/**
 * Script Name: exportFormulaInventory_v2_manifest
 * Status: STABLE — ETI v1.3
 *
 * Purpose:
 * - Capture spreadsheet formulas as inert text (manifest mode)
 * - Provide a deterministic, text-only formula inventory
 *
 * Explicit Non-Goals:
 * - Does NOT evaluate formulas
 * - Does NOT format headers (owned by sheet-level formatter)
 * - Does NOT mutate source sheets
 * - Does NOT interact with AppSheet or transactional logic
 *
 * Rules (Locked):
 * - Read formulas from ROW 2 only
 * - Strip leading "="
 * - Store formulas as inert text
 * - Clear + rewrite output on every run
 *
 * Input Dependencies:
 * - Active spreadsheet (all sheets)
 *
 * Output Contract:
 * - Sheet: Formula_Inventory
 * - Layout:
 *   Row 1  → Header (values only)
 *   Row 2–3 → Reserved empty rows
 *   Row 4+ → Data
 *   Two empty separator rows inserted when Sheet_Name changes
 *
 * Cosmetic Rules (Intentional & Limited):
 * - Separator rows are visually marked in columns A–F only
 * - Color: #FFFF00 (bright yellow)
 * - Purpose: improve human scanability between sheet blocks
 *
 * Idempotency:
 * - Safe to re-run; output is fully cleared and rebuilt
 */

function exportFormulaInventory_v2_manifest() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  const OUTPUT_SHEET = 'Formula_Inventory';
  let out = ss.getSheetByName(OUTPUT_SHEET);
  if (!out) out = ss.insertSheet(OUTPUT_SHEET);

  // Full reset (values + formatting)
  out.clear();

  const headers = [
    'Sheet_Name',
    'Column_Index',
    'Column_Letter',
    'Column_Name',
    'Formula_A1_Text',
    'Formula_R1C1_Text'
  ];

  // Header values only — formatting handled elsewhere
  out.getRange(1, 1, 1, headers.length).setValues([headers]);

  let writeRow = 4;
  let lastSheetName = null;

  sheets.forEach(sh => {

    const sheetName = sh.getName();
    const lastCol = sh.getLastColumn();

    // Skip sheets with no columns
    if (lastCol === 0) {
      lastSheetName = sheetName;
      return;
    }

    // ---- Insert visual separator between sheet blocks ----
    if (lastSheetName !== null && sheetName !== lastSheetName) {

      // Mark the two gap rows (A–F only)
      out.getRange(writeRow, 1, 2, headers.length)
         .setBackground('#FFFF00');

      writeRow += 2;
    }

    const headersRow = sh.getRange(1, 1, 1, lastCol).getValues()[0];

    for (let col = 1; col <= lastCol; col++) {

      const cell = sh.getRange(2, col);

      let formulaA1 = cell.getFormula() || '';
      let formulaR1C1 = cell.getFormulaR1C1() || '';

      // Strip leading "=" to enforce manifest mode
      if (formulaA1.startsWith('=')) {
        formulaA1 = formulaA1.slice(1);
      }
      if (formulaR1C1.startsWith('=')) {
        formulaR1C1 = formulaR1C1.slice(1);
      }

      out.getRange(writeRow, 1, 1, headers.length).setValues([[
        sheetName,
        col,
        columnToLetter(col),
        headersRow[col - 1] || '',
        formulaA1,
        formulaR1C1
      ]]);

      writeRow++;
    }

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
