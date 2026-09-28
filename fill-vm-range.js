#!/usr/bin/env node
/*
 * fill-vm-range.js
 *
 * Fills daily uptime cells for ONE VM across a date range, without re-querying
 * every other VM already tracked. Useful when a VM's row was added to the sheet
 * after other rows had already been backfilled for that range.
 *
 * Usage:
 *   node fill-vm-range.js <vmName> <startDate YYYY-MM-DD> <endDate YYYY-MM-DD>          # dry run
 *   node fill-vm-range.js <vmName> <startDate YYYY-MM-DD> <endDate YYYY-MM-DD> --write  # write in place (.bak made first)
 *
 * Column for each day must already exist in the target sheet (created by a prior
 * run of update-avd-usage.js for that day) -- this script does not append columns.
 */
const fs = require('fs');
const path = require('path');
const { execFile } = require('child_process');
const ExcelJS = require('exceljs');

const argv = process.argv.slice(2);
const DO_WRITE = argv.includes('--write');
const positional = argv.filter((a) => !a.startsWith('--'));
const [VM_NAME, START, END] = positional;
if (!VM_NAME || !START || !END) {
  console.error('Usage: node fill-vm-range.js <vmName> <start YYYY-MM-DD> <end YYYY-MM-DD> [--write]');
  process.exit(1);
}

const cfg = JSON.parse(fs.readFileSync(path.join(__dirname, 'config.json'), 'utf8'));
const WORKBOOK = cfg.workbook;
const MONTH_NAMES = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const EXCEL_EPOCH = Date.UTC(1899, 11, 30);

function az(args) {
  return new Promise((resolve, reject) => {
    execFile('az', args, { shell: true, maxBuffer: 64 * 1024 * 1024 }, (err, stdout, stderr) => {
      if (err) return reject(new Error((stderr || err.message).trim()));
      resolve(stdout);
    });
  });
}
async function azJson(args) { return JSON.parse(await az(args.concat(['-o', 'json']))); }

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
async function uptimeMinutes(resourceId, startIso, endIso) {
  const j = await azJson(['monitor', 'metrics', 'list', '--resource', `"${resourceId}"`,
    '--metric', `"${cfg.metric}"`, '--interval', 'PT1M',
    '--start-time', startIso, '--end-time', endIso, '--aggregation', 'average']);
  const ts = j.value && j.value[0] && j.value[0].timeseries && j.value[0].timeseries[0];
  if (!ts || !ts.data) return 0;
  return ts.data.filter((d) => d.average != null).length;
}
function cellDateSerial(cell) {
  const v = cell && cell.value;
  if (v == null) return null;
  if (v instanceof Date) return Math.round((Date.UTC(v.getUTCFullYear(), v.getUTCMonth(), v.getUTCDate()) - EXCEL_EPOCH) / 86400000);
  if (typeof v === 'number' && v > 40000 && v < 60000) return Math.round(v);
  return null;
}

(async () => {
  console.log(`Locating VM "${VM_NAME}" across configured subscriptions...`);
  const subs = (cfg.subscriptions && cfg.subscriptions.length) ? cfg.subscriptions : [null];
  let vm = null;
  for (const sub of subs) {
    try {
      const list = await azJson(['vm', 'list', '-d'].concat(sub ? ['--subscription', `"${sub}"`] : [])
        .concat(['--query', `"[?name=='${VM_NAME}'].{name:name, id:id, rg:resourceGroup, region:location}"`]));
      if (list.length) { vm = list[0]; break; }
    } catch (e) { console.error(`! lookup failed in ${sub || 'current subscription'}: ${e.message}`); }
  }
  if (!vm) { console.error(`! VM "${VM_NAME}" not found in any configured subscription. Check the exact name in Azure and re-run.`); process.exit(2); }
  console.log(`  found: ${vm.id}`);

  if (!fs.existsSync(WORKBOOK)) { console.error(`! workbook not found: ${WORKBOOK}`); process.exit(2); }
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(WORKBOOK);

  const start = new Date(START + 'T00:00:00Z');
  const end = new Date(END + 'T00:00:00Z');
  const results = [];
  for (let day = start; day <= end; day = new Date(day.getTime() + 86400000)) {
    const dayStart = new Date(Date.UTC(day.getUTCFullYear(), day.getUTCMonth(), day.getUTCDate()));
    const dayEnd = new Date(dayStart.getTime() + 24 * 3600 * 1000);
    const startIso = dayStart.toISOString().replace(/\.\d{3}Z$/, 'Z');
    const endIso = dayEnd.toISOString().replace(/\.\d{3}Z$/, 'Z');
    const targetSerial = Math.round((dayStart.getTime() - EXCEL_EPOCH) / 86400000);
    const dateLabel = startIso.slice(0, 10);

    const monthAbbr = MONTH_NAMES[day.getUTCMonth()].toLowerCase();
    const yr = String(day.getUTCFullYear()).slice(2);
    const sheet = wb.worksheets.find((s) => { const n = s.name.trim().toLowerCase(); return n.startsWith(monthAbbr) && n.endsWith(yr); });
    if (!sheet) { console.error(`  ${dateLabel}  ! no matching sheet, skipped`); continue; }

    let targetCol = null;
    sheet.getRow(1).eachCell({ includeEmpty: false }, (cell, col) => { if (cellDateSerial(cell) === targetSerial) targetCol = col; });
    if (!targetCol) { console.error(`  ${dateLabel}  ! no existing column in "${sheet.name}" (run update-avd-usage.js for this day first), skipped`); continue; }

    let rowNum = null;
    sheet.eachRow({ includeEmpty: false }, (row, rn) => {
      if (rn > 2 && (row.getCell(1).value || '').toString().trim().toLowerCase() === VM_NAME.toLowerCase()) rowNum = rn;
    });
    if (!rowNum) { console.error(`  ${dateLabel}  ! VM row not found in "${sheet.name}", skipped`); continue; }

    const mins = await uptimeMinutes(vm.id, startIso, endIso);
    const thr = thresholdFor(vm.name, vm.rg, vm.region);
    const extended = mins > thr;
    const text = fmtMinutes(mins) + (extended ? ' (Extended)' : '');
    console.log(`  ${dateLabel}  ${text}`);
    results.push({ sheet, rowNum, targetCol, text, extended });
  }

  if (!DO_WRITE) {
    console.log('\n(dry run -- re-run with --write to save these values)');
    return;
  }
  const bak = WORKBOOK.replace(/\.xlsx$/i, `.fillvm-${VM_NAME}-${Date.now()}.bak.xlsx`);
  fs.copyFileSync(WORKBOOK, bak);
  console.log(`Backup: ${bak}`);
  for (const r of results) {
    const cell = r.sheet.getRow(r.rowNum).getCell(r.targetCol);
    cell.value = r.text;
    cell.fill = r.extended
      ? { type: 'pattern', pattern: 'solid', fgColor: { argb: cfg.highlightColorARGB } }
      : { type: 'pattern', pattern: 'none' };
  }
  const buf = await wb.xlsx.writeBuffer();
  const tmp = `${WORKBOOK}.tmp-${process.pid}`;
  fs.writeFileSync(tmp, Buffer.from(buf));
  fs.renameSync(tmp, WORKBOOK);
  console.log(`Saved: ${WORKBOOK}`);
})().catch((e) => { console.error('\nFATAL:', e.message); process.exit(1); });
