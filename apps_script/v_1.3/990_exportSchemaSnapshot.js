/**
 * STEP 3 — Schema Snapshot
 *
 * Purpose:
 * - Capture column order and headers for all sheets
 * - Freeze schema as Google Sheets sees it
 *
 * Output Sheet:
 * - Schema_Snapshot
 */

function STEP3_exportSchemaSnapshot() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  const OUTPUT_SHEET_NAME = 'Schema_Snapshot';
  let out = ss.getSheetByName(OUTPUT_SHEET_NAME);

  if (!out) {
    out = ss.insertSheet(OUTPUT_SHEET_NAME);
  }

  out.clearContents();
  out.appendRow([
    'Sheet_Name',
    'Column_Index',
    'Column_Letter',
    'Header_Value'
  ]);

  sheets.forEach(sh => {

    const headerRange = sh.getRange(1, 1, 1, sh.getLastColumn());
    const headers = headerRange.getValues()[0];

    headers.forEach((header, idx) => {
      out.appendRow([
        sh.getName(),
        idx + 1,
        columnToLetter(idx + 1),
        header
      ]);
    });

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
