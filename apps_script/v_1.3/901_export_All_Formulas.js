/**
 * Script Name: export_All_Formulas
 * Purpose:
 * - Extract all formulas from all sheets
 * - Preserve exact cell positions
 * - Output audit-safe inventory
 *
 * Output Sheet: _Formula_Inventory
 */

function export_All_Formulas() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  const OUTPUT_SHEET = '_Formula_Inventory';
  let out = ss.getSheetByName(OUTPUT_SHEET);
  if (!out) out = ss.insertSheet(OUTPUT_SHEET);
  out.clearContents();

  out.appendRow([
    'Sheet_Name',
    'Cell',
    'Formula_A1',
    'Formula_R1C1'
  ]);

  sheets.forEach(sh => {

    const range = sh.getDataRange();
    const formulasA1 = range.getFormulas();
    const formulasR1C1 = range.getFormulasR1C1();

    for (let r = 0; r < formulasA1.length; r++) {
      for (let c = 0; c < formulasA1[r].length; c++) {

        if (formulasA1[r][c]) {
          out.appendRow([
            sh.getName(),
            sh.getRange(r + 1, c + 1).getA1Notation(),
            formulasA1[r][c],
            formulasR1C1[r][c]
          ]);
        }
      }
    }
  });

}
