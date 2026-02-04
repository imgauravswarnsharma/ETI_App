/**
 * SAFE regeneration into a shadow sheet
 * Never overwrites original
 */

function regenerateSheetPreview(sheetName) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheetName = `__REGEN__${sheetName}`;

  const SCHEMA = ss.getSheetByName('Schema_Snapshot').getDataRange().getValues();
  const FORMULAS = ss.getSheetByName('Formula_Inventory').getDataRange().getValues();

  // ---- Skip header row safely ----
  const cols = SCHEMA.slice(1).filter(r => r[0] === sheetName);
  if (cols.length === 0) {
    throw new Error(`Sheet not found in Schema_Snapshot: ${sheetName}`);
  }

  // ---- Create / reset preview sheet ----
  let sh = ss.getSheetByName(targetSheetName);
  if (!sh) sh = ss.insertSheet(targetSheetName);
  sh.clear();

  // ---- Write headers ----
  const headers = cols.map(r => r[3]);
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ---- Build formula lookup ----
  const formulaMap = {};
  for (let i = 1; i < FORMULAS.length; i++) {
    const [s, colIdx,,,, a1] = FORMULAS[i];
    if (s === sheetName && a1) {
      formulaMap[colIdx] = a1;
    }
  }

  // ---- Inject formulas into row 2 ----
  cols.forEach(col => {
    const colIdx = col[1];
    const formulaText = formulaMap[colIdx];
    if (formulaText) {
      sh.getRange(2, colIdx).setFormula('=' + formulaText);
    }
  });
}
