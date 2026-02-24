/**
 * Script Name: classifyColumns_fromManifest
 * Status: STABLE — ETI v1.3
 *
 * Purpose:
 * - Generate Column_Classification as a READ-ONLY analytical artifact
 * - Classify columns using Schema_Snapshot + Formula_Inventory (text-only)
 *
 * Explicit Non-Goals:
 * - Does NOT mutate schemas
 * - Does NOT write formulas
 * - Does NOT manage header formatting (owned by sheet-level formatter)
 * - Does NOT interact with AppSheet or transactional data
 *
 * Input Dependencies (Required):
 * - Schema_Snapshot
 * - Formula_Inventory (Formula_A1_Text)
 *
 * Algorithm (Optimized – Logic Unchanged):
 * 1. Read Schema + Formula inventory
 * 2. Build formula lookup map
 * 3. Build full classification output in memory
 * 4. Single batch write
 * 5. Apply separator formatting in batch
 * 
 * Output Contract:
 * - Sheet: Column_Classification (fully regenerated each run)
 * - Layout:
 *   Row 1  → Header (values only; formatting handled elsewhere)
 *   Row 2–3 → Reserved empty rows
 *   Row 4+ → Data
 *   Two empty separator rows inserted when Sheet_Name changes
 *
 * Cosmetic Rules (Intentional & Limited):
 * - Separator rows are visually marked in columns A–E only
 * - Purpose: improve human scanability between table blocks
 * - No data rows or logic rows are ever styled
 *
 * Idempotency:
 * - Safe to re-run; output is fully cleared and rebuilt
 */


function classifyColumns_fromManifest() {

  const dataSS = SpreadsheetApp.getActiveSpreadsheet();
  const metaSS = getMetadataSpreadsheet_();

  const SCHEMA_SHEET   = metaSS.getSheetByName('Schema_Snapshot');
  const FORMULA_SHEET  = metaSS.getSheetByName('Formula_Inventory');

  if (!SCHEMA_SHEET || !FORMULA_SHEET) {
    throw new Error('Required dependency sheet missing');
  }

  const SCHEMA   = SCHEMA_SHEET.getDataRange().getValues();
  const FORMULAS = FORMULA_SHEET.getDataRange().getValues();

  let out = metaSS.getSheetByName('Column_Classification');
  if (!out) out = metaSS.insertSheet('Column_Classification');

  out.clear();

  const headers = [
    'Sheet_Name',
    'Column_Index',
    'Column_Letter',
    'Column_Name',
    'Semantic_Class'
  ];

  out.getRange(1, 1, 1, headers.length).setValues([headers]);

  /* =========================================================
     BUILD FORMULA LOOKUP MAP
  ========================================================== */

  const formulaMap = {};
  for (let i = 1; i < FORMULAS.length; i++) {
    const [sheet, colIdx,,,, formulaA1] = FORMULAS[i];
    formulaMap[`${sheet}|${colIdx}`] = (formulaA1 || '');
  }

  /* =========================================================
     CLASSIFICATION BUILD (IN MEMORY)
  ========================================================== */

  const PASS_THROUGH_REGEX =
    /^IF\s*\(\s*[^,]+,\s*(?:[A-Z0-9_]+!)?\$?[A-Z]+\$?\d+\s*,\s*""\s*\)$/i;

  let output = [];
  let separatorRows = [];
  let lastSheet = null;

  for (let i = 1; i < SCHEMA.length; i++) {

    const [sheet, colIdx, colLetter, colName] = SCHEMA[i];

    if (lastSheet !== null && sheet !== lastSheet) {
      separatorRows.push(output.length + 4);
      separatorRows.push(output.length + 5);
      output.push(new Array(headers.length).fill(''));
      output.push(new Array(headers.length).fill(''));
    }

    const rawFormula = formulaMap[`${sheet}|${colIdx}`];

    let cls = 'EMPTY';

    if (rawFormula) {

      const f = rawFormula.replace(/\s+/g, ' ').trim();

      if (/(XLOOKUP|VLOOKUP|INDEX|MATCH)\s*\(/i.test(f)) {
        cls = 'DERIVED_LOOKUP';

      } else if (PASS_THROUGH_REGEX.test(f)) {
        cls = 'PASS_THROUGH';

      } else {
        cls = 'DERIVED_LOCAL';
      }
    }

    output.push([
      sheet,
      colIdx,
      colLetter,
      colName,
      cls
    ]);

    lastSheet = sheet;
  }

  /* =========================================================
     SINGLE BATCH WRITE
  ========================================================== */

  if (output.length > 0) {
    out.getRange(4, 1, output.length, headers.length)
       .setValues(output);
  }

  /* =========================================================
     BATCH SEPARATOR FORMATTING
  ========================================================== */

  separatorRows.forEach(r => {
    out.getRange(r, 1, 1, headers.length)
       .setBackground('#FFFF00');
  });
}