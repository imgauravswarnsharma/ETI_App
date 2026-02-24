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

  const dataSS = SpreadsheetApp.getActiveSpreadsheet();
  const metaSS = getMetadataSpreadsheet_();

  const sheets = dataSS.getSheets();


  const OUTPUT_SHEET = 'Formula_Inventory';
  let out = metaSS.getSheetByName(OUTPUT_SHEET);
  if (!out) out = metaSS.insertSheet(OUTPUT_SHEET);
  
  out.clear();

  const headers = [
    'Sheet_Name',
    'Column_Index',
    'Column_Letter',
    'Column_Name',
    'Formula_A1_Text',
    'Formula_R1C1_Text'
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

    if (lastSheetName !== null && sheetName !== lastSheetName) {
      separatorRows.push(output.length + 4); // +4 because metadata starts row 4
      separatorRows.push(output.length + 5);
      output.push(new Array(headers.length).fill(''));
      output.push(new Array(headers.length).fill(''));
    }

    const headersRow = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    const formulaA1 = sh.getRange(2, 1, 1, lastCol).getFormulas()[0];
    const formulaR1C1 = sh.getRange(2, 1, 1, lastCol).getFormulasR1C1()[0];

    for (let col = 1; col <= lastCol; col++) {

      let fA1 = formulaA1[col - 1] || '';
      let fR1C1 = formulaR1C1[col - 1] || '';

      if (fA1.startsWith('=')) fA1 = fA1.slice(1);
      if (fR1C1.startsWith('=')) fR1C1 = fR1C1.slice(1);

      output.push([
        sheetName,
        col,
        columnToLetter(col),
        headersRow[col - 1] || '',
        fA1,
        fR1C1
      ]);
    }

    lastSheetName = sheetName;
  });

  if (output.length > 0) {
    out.getRange(4, 1, output.length, headers.length)
       .setValues(output);
  }

  // Apply separator background in batch
  separatorRows.forEach(r => {
    out.getRange(r, 1, 1, headers.length)
       .setBackground('#FFFF00');
  });
}