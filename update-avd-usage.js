#!/usr/bin/env node
/*
 * update-avd-usage.js
 *
 * Reads each AVD VM's daily powered-on time from Azure Monitor and writes it into
 * the "Daily/Monthly Usage" workbook, flagging machines that ran past the scheduled
 * shutdown ("Extended").
 *
 * Uptime source: Azure VMs emit platform metrics ONLY while running. We count the
 * non-null per-minute (PT1M) buckets of a metric over the UTC day window; that count
 * is the number of minutes the VM was powered on (and it naturally sums multiple
 * on/off cycles and counts up to "now" if the VM is still running).
 *
 * Usage:
 *   node update-avd-usage.js                 # dry run: compute + print, do NOT touch the workbook
 *   node update-avd-usage.js --write         # write into config.workbook in place (a .bak is made first)
 *   node update-avd-usage.js --out copy.xlsx # write to a separate file, leave the original alone
 *   node update-avd-usage.js --date 2026-06-29   # target a specific UTC day (default: today)
 *   node update-avd-usage.js --workbook path.xlsx --sheet "Jun 26"
 *
 * Prereqs: `az login` to the right tenant, `az account set` will be applied from config.subscription,
 * Node + `npm install` in this folder (exceljs).
 */

const fs = require('fs');
const path = require('path');
const { execFile } = require('child_process');
const ExcelJS = require('exceljs');

// If output is piped to a reader that closes early (e.g. `| head`), further writes to
// stdout raise EPIPE — exit quietly instead of crashing.
process.stdout.on('error', (e) => { if (e.code === 'EPIPE') process.exit(0); });

// ---------- args ----------
const argv = process.argv.slice(2);
function flag(name) { return argv.includes(name); }
function opt(name, def) { const i = argv.indexOf(name); return i >= 0 && i + 1 < argv.length ? argv[i + 1] : def; }

const DO_WRITE = flag('--write') || argv.includes('--out');
const OUT_PATH = opt('--out', null);

const cfg = JSON.parse(fs.readFileSync(path.join(__dirname, 'config.json'), 'utf8'));
const WORKBOOK = opt('--workbook', cfg.workbook);
const SHEET_OVERRIDE = opt('--sheet', cfg.sheet);

// ---------- target day (UTC) ----------
// --yesterday : target the UTC day that just ended. Use this for the scheduled run
//   (e.g. ~00:45 UTC) so the day is fully closed and the FULL 24h window is counted —
//   this captures late extensions / anyone who worked past the prior evening, and the
//   detected stop time below is the real shutdown, not a mid-day snapshot.
const wantYesterday = flag('--yesterday');
const dateArg = opt('--date', null);
const now = new Date();
let target;
if (wantYesterday) target = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate()) - 86400000);
else if (dateArg) target = new Date(dateArg + 'T00:00:00Z');
else target = now;
const dayStart = new Date(Date.UTC(target.getUTCFullYear(), target.getUTCMonth(), target.getUTCDate(), 0, 0, 0));
// window end: now if targeting today, else end of that UTC day
const isToday = dayStart.getTime() === Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate());
const dayEnd = isToday ? now : new Date(dayStart.getTime() + 24 * 3600 * 1000);
const startIso = dayStart.toISOString().replace(/\.\d{3}Z$/, 'Z');
const endIso = dayEnd.toISOString().replace(/\.\d{3}Z$/, 'Z');

// Excel serial for the target day (Excel epoch = 1899-12-30)
const EXCEL_EPOCH = Date.UTC(1899, 11, 30);
const targetSerial = Math.round((dayStart.getTime() - EXCEL_EPOCH) / 86400000);

const MONTH_NAMES = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

// ---------- helpers ----------
function az(args) {
  return new Promise((resolve, reject) => {
    execFile('az', args, { shell: true, maxBuffer: 64 * 1024 * 1024 }, (err, stdout, stderr) => {
      if (err) return reject(new Error((stderr || err.message).trim()));
      resolve(stdout);
    });
  });
}

async function azJson(args) {
  const out = await az(args.concat(['-o', 'json']));
  return JSON.parse(out);
}

function fmtMinutes(mins) {
  if (!mins || mins <= 0) return 'No usage today';
  const h = Math.floor(mins / 60);
  const m = mins % 60;
  return `${h} hrs ${m} ${m === 1 ? 'min' : 'mins'}`;
}

function thresholdFor(vm, rg, region) {
  const o = cfg.thresholdOverrides || {};
  if (o.byVm && o.byVm[vm] != null) return o.byVm[vm];
  if (o.byResourceGroup && o.byResourceGroup[rg] != null) return o.byResourceGroup[rg];
  if (o.byRegion && o.byRegion[region] != null) return o.byRegion[region];
  return cfg.extendedThresholdMinutes;
}

// simple promise pool
async function pool(items, size, worker) {
  const results = new Array(items.length);
  let i = 0;
  const runners = Array.from({ length: Math.min(size, items.length) }, async () => {
    while (i < items.length) {
      const idx = i++;
      results[idx] = await worker(items[idx], idx);
    }
  });
  await Promise.all(runners);
  return results;
}

// Count non-null PT1M buckets within [startIso, endIso].
// Each bucket with data = 1 minute powered on. First/last bucket ≈ boot/stop time.
// Returns { minutes, firstStart, lastStop, ranToWindowEnd }.
async function uptimeDetail(resourceId) {
  const j = await azJson([
    'monitor', 'metrics', 'list',
    '--resource', `"${resourceId}"`,
    '--metric', `"${cfg.metric}"`,
    '--interval', 'PT1M',
    '--start-time', startIso,
    '--end-time', endIso,
    '--aggregation', 'average',
  ]);
  const ts = j.value && j.value[0] && j.value[0].timeseries && j.value[0].timeseries[0];
  if (!ts || !ts.data) return { minutes: 0, firstStart: null, lastStop: null, ranToWindowEnd: false };
  const live = ts.data.filter((d) => d.average != null);
  if (!live.length) return { minutes: 0, firstStart: null, lastStop: null, ranToWindowEnd: false };
  const lastStop = new Date(live[live.length - 1].timeStamp);
  // if the last running minute is within 2 min of the window end, the VM hadn't stopped
  // inside this window (it crossed the boundary / is still running) — flag it.
  const ranToWindowEnd = (dayEnd.getTime() - lastStop.getTime()) <= 2 * 60 * 1000;
  return { minutes: live.length, firstStart: new Date(live[0].timeStamp), lastStop, ranToWindowEnd };
}

const hhmm = (d) => d ? d.toISOString().slice(11, 16) + 'Z' : '--:--';

// Reconstruct uptime from the Activity Log — used when a sheet row's VM is no longer in
// the live inventory (deprovisioned). The platform-metrics API returns ResourceNotFound
// for a deleted VM, but the Activity Log keeps its control-plane events for ~90 days.
//
// A VM may be started several times in a day; we sum EVERY powered-on interval. The
// end-of-day shutdown is an `az vm deallocate`, so we also check that a deallocate
// actually closed the last run — if the VM was still running at the window end it was
// NOT deallocated (shutdown pipeline missed it / it crossed midnight), which we flag.
async function activityLogUptime(vmName, candidateSubIds, candidateRgs) {
  const seedFrom = new Date(dayStart.getTime() - 24 * 3600 * 1000).toISOString().replace(/\.\d{3}Z$/, 'Z');
  const pairs = [];
  for (const subId of candidateSubIds) for (const rg of candidateRgs) pairs.push({ subId, rg });
  for (const { subId, rg } of pairs) {
    const id = `/subscriptions/${subId}/resourceGroups/${rg}/providers/Microsoft.Compute/virtualMachines/${vmName}`;
    let ev;
    try {
      ev = await azJson(['monitor', 'activity-log', 'list', '--resource-id', `"${id}"`,
        '--start-time', seedFrom, '--end-time', endIso,
        '--query', `"[?status.value=='Succeeded'].{t:eventTimestamp, op:operationName.value}"`]);
    } catch (e) { continue; }
    if (!ev || !ev.length) continue;

    // start = power on; deallocate / powerOff / delete = power off
    const tr = ev.map((e) => ({
      t: new Date(e.t),
      on: /virtualMachines\/start\/action$/i.test(e.op),
      off: /(deallocate\/action|powerOff\/action|\/delete)$/i.test(e.op),
    })).filter((e) => e.on || e.off).sort((a, b) => a.t - b.t);
    if (!tr.length) continue;

    // seed running-state at dayStart from the last transition before it
    let running = false;
    for (const e of tr) { if (e.t < dayStart) running = e.on ? true : (e.off ? false : running); }

    let segStart = running ? dayStart : null;
    let minutes = 0, firstStart = running ? dayStart : null, lastStop = null, starts = 0, deallocs = 0;
    for (const e of tr) {
      if (e.t < dayStart || e.t >= dayEnd) continue;
      if (e.on && !running) { running = true; segStart = e.t; starts++; if (!firstStart) firstStart = e.t; }
      else if (e.off && running) { running = false; minutes += (e.t - segStart) / 60000; lastStop = e.t; deallocs++; segStart = null; }
    }
    if (running) { minutes += (dayEnd - segStart) / 60000; lastStop = dayEnd; } // never closed within the day

    return {
      minutes: Math.round(minutes),
      firstStart, lastStop,
      ranToWindowEnd: running,      // true = still running at window end
      deallocatedByEnd: !running,   // false = end-of-day deallocate did NOT happen
      starts, deallocs,
      source: 'activity-log', rg, hadEvents: true,
    };
  }
  return { minutes: 0, firstStart: null, lastStop: null, ranToWindowEnd: false, deallocatedByEnd: null, source: 'activity-log', hadEvents: false };
}

function cellDateSerial(cell) {
  const v = cell && cell.value;
  if (v == null) return null;
  if (v instanceof Date) return Math.round((Date.UTC(v.getUTCFullYear(), v.getUTCMonth(), v.getUTCDate()) - EXCEL_EPOCH) / 86400000);
  if (typeof v === 'number' && v > 40000 && v < 60000) return Math.round(v);
  return null;
}

function colToLetter(n) {
  let s = '';
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

(async () => {
  console.log(`AVD usage tracker`);
  console.log(`  target day (UTC): ${startIso} .. ${endIso}  (serial ${targetSerial})`);
  console.log(`  mode: ${DO_WRITE ? (OUT_PATH ? 'write -> ' + OUT_PATH : 'write in place') : 'DRY RUN (no file changes)'}`);

  // 1) VM inventory across ALL configured subscriptions (the fleets live in different subs).
  //    name -> {id, rg, region, powerState}. The full id carries its own subscription, so the
  //    metric/activity-log queries work regardless of which sub is "current".
  console.log('Loading VM inventory...');
  const subs = (cfg.subscriptions && cfg.subscriptions.length) ? cfg.subscriptions
    : (cfg.subscription ? [cfg.subscription] : [null]);
  let vms = [];
  for (const sub of subs) {
    try {
      const list = await azJson(['vm', 'list', '-d'].concat(sub ? ['--subscription', `"${sub}"`] : [])
        .concat(['--query', '"[].{name:name, id:id, rg:resourceGroup, region:location, power:powerState}"']));
      vms = vms.concat(list);
      console.log(`  ${list.length} VMs in ${sub || 'current subscription'}`);
    } catch (e) {
      console.error(`! vm list failed for ${sub}: ${e.message}`);
    }
  }
  const byName = new Map(vms.map((v) => [v.name.toLowerCase(), v]));

  // candidate subscription IDs + RGs for reconstructing deleted VMs from the Activity Log
  const candidateSubIds = [...new Set(vms.map((v) => v.id.split('/')[2]))];
  const candidateRgs = (cfg.deletedVmLookup && cfg.deletedVmLookup.resourceGroups && cfg.deletedVmLookup.resourceGroups.length)
    ? cfg.deletedVmLookup.resourceGroups
    : [...new Set(vms.map((v) => v.rg))];

  // 2) Open workbook
  if (!fs.existsSync(WORKBOOK)) {
    console.error(`! workbook not found: ${WORKBOOK}\n  Set config.workbook or pass --workbook <path>.`);
    process.exit(2);
  }
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(WORKBOOK);

  // pick sheet
  let sheet;
  if (SHEET_OVERRIDE) sheet = wb.getWorksheet(SHEET_OVERRIDE);
  if (!sheet) {
    // Match loosely: sheet naming is inconsistent across the workbook's history
    // (abbreviated vs. full month names, typos like "Novemeber", double spaces),
    // so match on the month's first 3 letters + the 2-digit year appearing anywhere.
    const monthAbbr = MONTH_NAMES[target.getUTCMonth()].toLowerCase();
    const yr = String(target.getUTCFullYear()).slice(2);
    sheet = wb.worksheets.find((s) => {
      const n = s.name.trim().toLowerCase();
      return n.startsWith(monthAbbr) && n.endsWith(yr);
    });
  }
  if (!sheet) { console.error('! could not resolve target sheet'); process.exit(2); }
  console.log(`  sheet: "${sheet.name}"`);

  // 3) Locate the target date column (match serial in row 1); else append to the right.
  const headerRow = sheet.getRow(1);
  let targetCol = null;
  let lastDateCol = 0;
  // Columns 1-3 are always Name/User/Region, even on a brand-new sheet whose header
  // row hasn't been touched yet (cellCount would otherwise read as low as 1, which
  // pushed the first appended date column into the User/Region columns themselves).
  let lastUsedCol = Math.max(headerRow.cellCount, 3);
  headerRow.eachCell({ includeEmpty: false }, (cell, col) => {
    const serial = cellDateSerial(cell);
    if (serial != null) { lastDateCol = Math.max(lastDateCol, col); if (serial === targetSerial) targetCol = col; }
  });
  let appended = false;
  if (!targetCol) {
    targetCol = Math.max(lastUsedCol, lastDateCol) + 1;
    appended = true;
  }
  console.log(`  ${appended ? 'appending NEW' : 'updating existing'} column ${colToLetter(targetCol)} (col ${targetCol})`);

  // 4) Walk data rows (name in column A). Header rows are 1-2.
  const jobs = [];
  const skipped = [];
  sheet.eachRow({ includeEmpty: false }, (row, rowNum) => {
    if (rowNum <= 2) return;
    const name = (row.getCell(1).value || '').toString().trim();
    if (!name) return;
    if ((cfg.skipPrefixes || []).some((p) => name.toLowerCase().startsWith(p.toLowerCase()))) {
      skipped.push({ rowNum, name, reason: 'skip-prefix' });
      return;
    }
    const vm = byName.get(name.toLowerCase());
    // VM gone from inventory (deprovisioned) -> reconstruct from Activity Log instead of skipping
    jobs.push({ rowNum, name, vm: vm || null, deleted: !vm });
  });
  const deletedCount = jobs.filter((j) => j.deleted).length;
  console.log(`  ${jobs.length} VM rows to query (${deletedCount} via Activity Log fallback), ${skipped.length} skipped`);

  // 5) Query uptime concurrently
  let done = 0;
  const results = await pool(jobs, cfg.concurrency || 8, async (job) => {
    let detail = { minutes: 0, lastStop: null, firstStart: null, ranToWindowEnd: false }, err = null;
    try {
      detail = job.vm
        ? await uptimeDetail(job.vm.id)
        : await activityLogUptime(job.name, candidateSubIds, candidateRgs);
    } catch (e) { err = e.message; }
    done++;
    process.stdout.write(`\r  queried ${done}/${jobs.length}   `);
    // deleted VM with no recoverable Activity Log data (deleted >retention, or never ran)
    const noData = !!job.deleted && !err && detail && detail.hadEvents === false;
    const rg = job.vm ? job.vm.rg : detail && detail.rg;
    const region = job.vm ? job.vm.region : null;
    const mins = detail ? detail.minutes : 0;
    const thr = thresholdFor(job.name, rg, region);
    const extended = !noData && mins > thr;
    let text = fmtMinutes(mins);
    if (extended) text += ' (Extended)';
    return { ...job, mins, extended, text, err, noData, power: job.vm ? job.vm.power : 'deleted', ...detail };
  });
  process.stdout.write('\n');

  // 6) Header for an appended column: always set the date (+ sub-header), even on a
  // brand-new sheet with no prior date column to copy style from -- otherwise the date
  // never gets recorded and every subsequent run can't tell this column is already used,
  // so it keeps re-resolving to the same column instead of appending a new one.
  if (appended) {
    const hc = headerRow.getCell(targetCol);
    hc.value = dayStart;            // exceljs stores as date; numFmt from copied style renders it
    const sub = sheet.getRow(2);
    const subc = sub.getCell(targetCol);
    subc.value = 'Daily/Avg time';
    if (lastDateCol) {
      const prev = headerRow.getCell(lastDateCol);
      if (prev.style) hc.style = JSON.parse(JSON.stringify(prev.style));
      const prevSub = sub.getCell(lastDateCol);
      subc.value = prevSub.value || 'Daily/Avg time';
      if (prevSub.style) subc.style = JSON.parse(JSON.stringify(prevSub.style));
    }
    headerRow.commit();
    sub.commit();
  }

  // "not deallocated by end of day" — VM was still running at the window end. For a past
  // day that means the end-of-day deallocate never happened (or it crossed midnight); for
  // a live "today" run it's just still running now, which is normal, so don't alarm on it.
  const notDeallocated = results.filter((r) => !r.err && !r.noData && r.ranToWindowEnd && !isToday);

  // 7) Write values (or just print in dry run)
  console.log('\nResults:');
  for (const r of results) {
    let tag;
    if (r.err) {
      tag = `ERROR: ${r.err}`;
    } else if (r.noData) {
      tag = 'no Activity Log data (deleted > retention, or never ran) — left untouched';
    } else if (r.mins > 0) {
      const span = `${hhmm(r.firstStart)}-${hhmm(r.lastStop)}`;
      const dealloc = r.ranToWindowEnd ? (isToday ? ' still running' : ' NOT deallocated') : '';
      const src = r.source === 'activity-log' ? ' via activity-log' : '';
      tag = `${r.text.padEnd(26)} [${span}${dealloc}${src}]`;
    } else {
      tag = r.text + (r.source === 'activity-log' ? '  [via activity-log]' : '');
    }
    console.log(`  ${r.name.padEnd(16)} ${tag}`);

    // don't write rows that errored or have no recoverable data — leave the cell untouched
    if (!DO_WRITE || r.err || r.noData) continue;
    const cell = sheet.getRow(r.rowNum).getCell(targetCol);
    cell.value = r.text;
    // preserve a sensible base style by copying the row's previous date cell, then add highlight
    if (appended && lastDateCol) {
      const srcCell = sheet.getRow(r.rowNum).getCell(lastDateCol);
      if (srcCell.style) cell.style = JSON.parse(JSON.stringify(srcCell.style));
    }
    // set highlight on extended, and explicitly clear any inherited fill otherwise
    if (r.extended) {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: cfg.highlightColorARGB } };
    } else {
      cell.fill = { type: 'pattern', pattern: 'none' };
    }
    // document reconstructed / anomalous cells with an Excel note (non-intrusive red triangle)
    const notes = [];
    if (r.source === 'activity-log') notes.push(`Reconstructed from Activity Log (VM deleted; RG ${r.rg}).`);
    if (!isToday && r.ranToWindowEnd) notes.push('Still running at end of UTC day — end-of-day deallocate not seen.');
    if (notes.length) cell.note = notes.join(' ');
  }

  if (notDeallocated.length) {
    console.log('\n⚠ Not deallocated by end of day (still running at UTC midnight):');
    for (const r of notDeallocated) console.log(`  row ${r.rowNum}: ${r.name}  last seen ${hhmm(r.lastStop)}`);
  }

  if (skipped.length) {
    console.log('\nSkipped (left untouched — handle manually):');
    for (const s of skipped) console.log(`  row ${s.rowNum}: ${s.name}  (${s.reason})`);
  }

  const extCount = results.filter((r) => r.extended && !r.err && !r.noData).length;
  const errCount = results.filter((r) => r.err).length;
  const noDataCount = results.filter((r) => r.noData).length;
  const recovered = results.filter((r) => r.source === 'activity-log' && !r.noData && !r.err).length;
  console.log(`\nSummary: ${results.length - errCount - noDataCount} written-eligible, ${extCount} extended, ${recovered} recovered via activity-log, ${notDeallocated.length} not-deallocated, ${errCount} errors, ${skipped.length + noDataCount} skipped/no-data.`);

  // 8) Save
  if (DO_WRITE) {
    const dest = OUT_PATH || WORKBOOK;
    if (!OUT_PATH && fs.existsSync(WORKBOOK)) {
      const bak = WORKBOOK.replace(/\.xlsx$/i, `.${startIso.slice(0, 10)}.bak.xlsx`);
      fs.copyFileSync(WORKBOOK, bak);
      console.log(`Backup: ${bak}`);
    }
    // Write via an in-memory buffer + atomic rename instead of exceljs's streaming writeFile.
    // The streaming path pipes a zip stream to an fs stream and can throw EPIPE on some
    // container/overlay filesystems; a single buffered write is robust and leaves the
    // original intact if the process is interrupted mid-write.
    const buf = await wb.xlsx.writeBuffer();
    const tmp = `${dest}.tmp-${process.pid}`;
    fs.writeFileSync(tmp, Buffer.from(buf));
    fs.renameSync(tmp, dest);
    console.log(`Saved: ${dest}`);
  } else {
    console.log('\n(dry run — re-run with --write to update the workbook, or --out <file> to write a copy)');
  }
})().catch((e) => { console.error('\nFATAL:', e.message); process.exit(1); });
