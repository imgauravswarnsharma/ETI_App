/**
 * Column_Classification — Manifest-based
 * Uses Formula_Inventory v2 (text only)
 *
 * Layout:
 * Row 1  → Header (bold, written once)
 * Row 2  → empty
 * Row 3  → empty
 * Row 4+ → data
 * 2 empty rows inserted when Sheet_Name changes
 */

function classifyColumns_fromManifest() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const SCHEMA = ss.getSheetByName('Schema_Snapshot').getDataRange().getValues();
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

  out.getRange(1, 1, 1, headers.length)
     .setValues([headers])
     .setFontWeight('bold');

  // Build lookup: Sheet|ColIndex → Formula_A1_Text
  const formulaMap = {};
  for (let i = 1; i < FORMULAS.length; i++) {
    const [sheet, colIdx,,,, formulaA1] = FORMULAS[i];
    formulaMap[`${sheet}|${colIdx}`] = (formulaA1 || '');
  }

  let writeRow = 4;
  let lastSheet = null;

  // Robust PASS_THROUGH detector (after normalization)
  const PASS_THROUGH_REGEX =
    /^IF\s*\(\s*[^,]+,\s*(?:[A-Z0-9_]+!)?\$?[A-Z]+\$?\d+\s*,\s*""\s*\)$/i;

  for (let i = 1; i < SCHEMA.length; i++) {

    const [sheet, colIdx, colLetter, colName] = SCHEMA[i];

    if (lastSheet !== null && sheet !== lastSheet) {
      writeRow += 2;
    }

    const rawFormula = formulaMap[`${sheet}|${colIdx}`];

    let cls = 'EMPTY';

    if (rawFormula) {

      // --- normalize whitespace (CRITICAL FIX) ---
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
