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

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const SCHEMA   = ss.getSheetByName('Schema_Snapshot').getDataRange().getValues();
  const FORMULAS = ss.getSheetByName('Formula_Inventory').getDataRange().getValues();

  let out = ss.getSheetByName('Column_Classification');
  if (!out) out = ss.insertSheet('Column_Classification');

  out.clear();

  const headers = [
    'Sheet_Name',
    'Column_Index',
    'Column_Letter',
    'Column_Name',
    'Semantic_Class'
  ];

  // Header values only — NO formatting here
  out.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Build lookup: Sheet|ColIndex → Formula_A1_Text
  const formulaMap = {};
  for (let i = 1; i < FORMULAS.length; i++) {
    const [sheet, colIdx,,,, formulaA1] = FORMULAS[i];
    formulaMap[`${sheet}|${colIdx}`] = (formulaA1 || '');
  }

  let writeRow  = 4;
  let lastSheet = null;

  // Robust PASS_THROUGH detector (after normalization)
  const PASS_THROUGH_REGEX =
    /^IF\s*\(\s*[^,]+,\s*(?:[A-Z0-9_]+!)?\$?[A-Z]+\$?\d+\s*,\s*""\s*\)$/i;

  for (let i = 1; i < SCHEMA.length; i++) {

    const [sheet, colIdx, colLetter, colName] = SCHEMA[i];

    // ---- Sheet boundary: insert 2 empty rows + visual marker ----
    if (lastSheet !== null && sheet !== lastSheet) {

      // Mark the two separator rows (A:E only)
      out.getRange(writeRow, 1, 2, 5)
         .setBackground('#FFFF00'); // yellow color

      writeRow += 2;
    }

    const rawFormula = formulaMap[`${sheet}|${colIdx}`];

    let cls = 'EMPTY';

    if (rawFormula) {

      // normalize whitespace
      const f = rawFormula.replace(/\s+/g, ' ').trim();

      if (/(XLOOKUP|VLOOKUP|INDEX|MATCH)\s*\(/i.test(f)) {
        cls = 'DERIVED_LOOKUP';

      } else if (PASS_THROUGH_REGEX.test(f)) {
        cls = 'PASS_THROUGH';

      } else {
        cls = 'DERIVED_LOCAL';
      }
    }

    out.getRange(writeRow, 1, 1, headers.length).setValues([[
      sheet,
      colIdx,
      colLetter,
      colName,
      cls
    ]]);

    lastSheet = sheet;
    writeRow++;
  }
}
