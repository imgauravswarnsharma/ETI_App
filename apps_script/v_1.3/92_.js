/**
 * Phase 1 Implementation
 * Step: Generate Derived_Column_Logic for DERIVED_LOCAL columns
 * Scope: Transaction_Resolution only
 */

function PHASE1_generateDerivedLocalLogic() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const SCHEMA_SHEET = 'Schema_Snapshot';
  const CLASS_SHEET  = 'Column_Classification';
  const FORMULA_SHEET = 'Formula_Inventory';
  const OUTPUT_SHEET = 'Derived_Column_Logic';

  const TARGET_SHEET_NAME = 'Transaction_Resolution';

  const schema = ss.getSheetByName(SCHEMA_SHEET).getDataRange().getValues();
  const classes = ss.getSheetByName(CLASS_SHEET).getDataRange().getValues();
  const formulas = ss.getSheetByName(FORMULA_SHEET).getDataRange().getValues();

  let out = ss.getSheetByName(OUTPUT_SHEET);
  if (!out) out = ss.insertSheet(OUTPUT_SHEET);
  out.clearContents();

  out.appendRow([
    'Sheet_Name',
    'Column_Index',
    'Column_Letter',
    'Column_Name',
    'Column_Class',
    'Formula_A1',
    'Formula_R1C1',
    'Ref_1_Sheet',
    'Ref_1_Column_Index',
    'Ref_1_Column_Name'
  ]);

  // Build schema lookup: Sheet + Column_Index → Column_Name
  const schemaMap = {};
  for (let i = 1; i < schema.length; i++) {
    const [sheet, colIdx,, colName] = schema[i];
    schemaMap[`${sheet}|${colIdx}`] = colName;
  }

  // Iterate classified columns
  for (let i = 1; i < classes.length; i++) {

    const [sheet, colIdx, colLetter, colName, colClass] = classes[i];

    if (
      sheet !== TARGET_SHEET_NAME ||
      colClass !== 'DERIVED_LOCAL'
    ) continue;

    // Find first formula cell in this column
    const formulaRow = formulas.find(r =>
      r[0] === sheet &&
      r[1].startsWith(colLetter)
    );

    if (!formulaRow) continue;

    const formulaA1 = formulaRow[2];
    const formulaR1C1 = formulaRow[3];

    // Extract RC[-n]
    const match = formulaR1C1.match(/RC\[(-\d+)\]/);
    if (!match) continue;

    const offset = parseInt(match[1], 10);
    const refColIdx = colIdx + offset;
    const refColName = schemaMap[`${sheet}|${refColIdx}`];

    out.appendRow([
      sheet,
      colIdx,
      colLetter,
      colName,
      colClass,
      formulaA1,
      formulaR1C1,
      sheet,
      refColIdx,
      refColName
    ]);
  }
}
