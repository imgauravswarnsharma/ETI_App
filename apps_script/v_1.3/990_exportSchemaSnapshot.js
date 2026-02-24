/**
 * Script Name: STEP3_exportSchemaSnapshot
 * Status: STABLE — ETI v1.3
 *
 * Purpose:
 * - Capture column order and header values for all sheets
 * - Freeze schema exactly as Google Sheets exposes it
 *
 * Explicit Non-Goals:
 * - Does NOT format headers (owned by sheet-level formatter)
 * - Does NOT mutate source sheets
 * - Does NOT interact with AppSheet or transactional data
 *
 * Input Dependencies:
 * - Active spreadsheet (all visible sheets)
 * 
 * Algorithm (Optimized – Logic Unchanged):
 * 1. Clear output sheet
 * 2. Build full output array in memory
 * 3. Write once using setValues()
 * 4. Apply separator formatting in batch
 *
 * Output Contract:
 * - Sheet: Schema_Snapshot (fully regenerated each run)
 * - Layout:
 *   Row 1  → Header (values only)
 *   Row 2–3 → Reserved empty rows
 *   Row 4+ → Data
 *   Two empty separator rows inserted when Sheet_Name changes
 *
 * Cosmetic Rules (Intentional & Limited):
 * - Separator rows are visually marked in columns A–D only
 * - Color: #FFFF00 (bright yellow)
 * - Purpose: improve human scanability between sheet blocks
 *
 * Idempotency:
 * - Safe to re-run; output is cleared and rebuilt deterministically
 */

function STEP3_exportSchemaSnapshot() {

  const dataSS = SpreadsheetApp.getActiveSpreadsheet();
  const metaSS = getMetadataSpreadsheet_();
  const sheets = dataSS.getSheets();

  const OUTPUT_SHEET_NAME = 'Schema_Snapshot';
  let out = metaSS.getSheetByName(OUTPUT_SHEET_NAME);
  if (!out) out = metaSS.insertSheet(OUTPUT_SHEET_NAME);

  out.clear();

  const headers = [
    'Sheet_Name',
    'Column_Index',
    'Column_Letter',
    'Header_Value'
  ];

  out.getRange(1, 1, 1, headers.length).setValues([headers]);

  let output = [];
  let separatorRows = [];
  let lastSheetName = null;

  sheets.forEach(sh => {

    const sheetName = sh.getName();
    const lastCol = sh.getLastColumn();

    if (lastCol === 0) {
      lastSheetName = sheetName;
      return;
    }

    // Insert separator (logical only for now)
    if (lastSheetName !== null && sheetName !== lastSheetName) {
      separatorRows.push(output.length + 4);
      separatorRows.push(output.length + 5);
      output.push(new Array(headers.length).fill(''));
      output.push(new Array(headers.length).fill(''));
    }

    const headerValues =
      sh.getRange(1, 1, 1, lastCol).getValues()[0];

    for (let idx = 0; idx < headerValues.length; idx++) {
      output.push([
        sheetName,
        idx + 1,
        columnToLetter(idx + 1),
        headerValues[idx]
      ]);
    }

    lastSheetName = sheetName;
  });

  // Single batch write
  if (output.length > 0) {
    out.getRange(4, 1, output.length, headers.length)
       .setValues(output);
  }

  // Apply separator formatting in batch
  separatorRows.forEach(r => {
    out.getRange(r, 1, 1, headers.length)
       .setBackground('#FFFF00');
  });
}

/**
 * Utility: Convert column number to letter (1 → A, 27 → AA)
 */
function columnToLetter(column) {
  let temp = '';
  let letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}