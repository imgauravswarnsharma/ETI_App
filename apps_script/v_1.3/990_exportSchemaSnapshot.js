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

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  const OUTPUT_SHEET_NAME = 'Schema_Snapshot';
  let out = ss.getSheetByName(OUTPUT_SHEET_NAME);
  if (!out) out = ss.insertSheet(OUTPUT_SHEET_NAME);

  out.clear();

  const headers = [
    'Sheet_Name',
    'Column_Index',
    'Column_Letter',
    'Header_Value'
  ];

  // Header values only — formatting handled elsewhere
  out.getRange(1, 1, 1, headers.length).setValues([headers]);

  let writeRow = 4;
  let lastSheetName = null;

  sheets.forEach(sh => {

    const sheetName = sh.getName();

    // ---- Insert visual separator between sheet blocks ----
    if (lastSheetName !== null && sheetName !== lastSheetName) {

      // Mark the two gap rows (A–D only)
      out.getRange(writeRow, 1, 2, headers.length)
         .setBackground('#FFFF00');

      writeRow += 2;
    }

    const headerRange = sh.getRange(1, 1, 1, sh.getLastColumn());
    const headerValues = headerRange.getValues()[0];

    headerValues.forEach((header, idx) => {
      out.getRange(writeRow, 1, 1, headers.length).setValues([[
        sheetName,
        idx + 1,
        columnToLetter(idx + 1),
        header
      ]]);
      writeRow++;
    });

    lastSheetName = sheetName;
  });
}

/**
 * Utility: Convert column number to letter (1 → A, 27 → AA)
 * Pure utility — no side effects
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
