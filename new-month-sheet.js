#!/usr/bin/env node
/*
 * new-month-sheet.js
 *
 * Clones the VM roster (Name/User/Region columns + their styles/column widths)
 * from an existing month sheet into a brand-new sheet for the new month. Date
 * columns are NOT copied -- update-avd-usage.js appends those as it runs each day.
 *
 * Usage:
 *   node new-month-sheet.js "Aug 26" "July 26"   # <newSheetName> <sourceSheetName>
 *   node new-month-sheet.js "Aug 26"             # source defaults to the workbook's first sheet
 */
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const [NEW_NAME, SRC_NAME] = process.argv.slice(2);
if (!NEW_NAME) {
  console.error('Usage: node new-month-sheet.js "<newSheetName>" ["<sourceSheetName>"]');
  process.exit(1);
}

const cfg = JSON.parse(fs.readFileSync(path.join(__dirname, 'config.json'), 'utf8'));
const WORKBOOK = cfg.workbook;

(async () => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(WORKBOOK);

  const existing = wb.worksheets.find((s) => s.name.trim().toLowerCase() === NEW_NAME.trim().toLowerCase());
  if (existing) { console.error(`! a sheet named "${NEW_NAME}" already exists.`); process.exit(2); }

  const src = SRC_NAME
    ? wb.worksheets.find((s) => s.name.trim().toLowerCase() === SRC_NAME.trim().toLowerCase())
    : wb.worksheets[0];
  if (!src) { console.error(`! source sheet not found: ${SRC_NAME}`); process.exit(2); }
  console.log(`Cloning roster from "${src.name}" into new sheet "${NEW_NAME}"...`);

  const dst = wb.addWorksheet(NEW_NAME, { views: src.views });

  // column widths for the roster columns (A/B/C) + a sensible default for date columns
  for (let c = 1; c <= 3; c++) {
    const col = src.getColumn(c);
    dst.getColumn(c).width = col.width;
  }

  let copiedRows = 0;
  src.eachRow({ includeEmpty: false }, (srcRow, rowNum) => {
    const dstRow = dst.getRow(rowNum);
    dstRow.height = srcRow.height;
    for (let c = 1; c <= 3; c++) {
      const srcCell = srcRow.getCell(c);
      const dstCell = dstRow.getCell(c);
      dstCell.value = srcCell.value;
      if (srcCell.style) dstCell.style = JSON.parse(JSON.stringify(srcCell.style));
    }
    dstRow.commit();
    if (rowNum > 2 && (srcRow.getCell(1).value || '').toString().trim()) copiedRows++;
  });
  console.log(`  copied ${copiedRows} VM rows (+ header rows 1-2).`);

  const bak = WORKBOOK.replace(/\.xlsx$/i, `.newsheet-${Date.now()}.bak.xlsx`);
  fs.copyFileSync(WORKBOOK, bak);
  console.log(`Backup: ${bak}`);

  const buf = await wb.xlsx.writeBuffer();
  const tmp = `${WORKBOOK}.tmp-${process.pid}`;
  fs.writeFileSync(tmp, Buffer.from(buf));
  fs.renameSync(tmp, WORKBOOK);
  console.log(`Saved: ${WORKBOOK}`);
  console.log(`\nNext: run "node update-avd-usage.js --write --date <first day of month>" to populate day 1, then continue with --yesterday daily as usual.`);
})().catch((e) => { console.error('FATAL:', e.message); process.exit(1); });
