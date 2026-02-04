/**
 * Formula_Inventory v2 — MANIFEST MODE
 *
 * Rules:
 * - Read formulas from ROW 2 only
 * - Strip leading "="
 * - Store as inert text
 * - Clear + rewrite on every run
 */

function exportFormulaInventory_v2_manifest() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  const OUTPUT_SHEET = 'Formula_Inventory';

  let out = ss.getSheetByName(OUTPUT_SHEET);
  if (!out) out = ss.insertSheet(OUTPUT_SHEET);

  // Full reset (values + formatting)
  out.clear();

  out.appendRow([
    'Sheet_Name',
    'Column_Index',
    'Column_Letter',
    'Column_Name',
    'Formula_A1_Text',
    'Formula_R1C1_Text'
  ]);

  sheets.forEach(sh => {

    const sheetName = sh.getName();
    const lastCol = sh.getLastColumn();
    if (lastCol === 0) return;

    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];

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

      out.appendRow([
        sheetName,
        col,
        columnToLetter(col),
        headers[col - 1] || '',
        formulaA1,
        formulaR1C1
      ]);
    }
  });
}

/**
 * Utility: column number → letter
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
