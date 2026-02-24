/**
 * Script Name: generateDerivedColumnLogic
 * Status: STABLE — ETI v1.3
 *
 * Purpose:
 * - Resolve derived column dependencies using inert formulas
 * - Expand same-sheet (R1C1) and cross-sheet (A1) references
 *
 * Explicit Non-Goals:
 * - Does NOT evaluate formulas
 * - Does NOT mutate schemas or formulas
 * - Does NOT format headers (owned by sheet-level formatter)
 * - Does NOT interact with AppSheet or transactional data
 *
 * Input Dependencies (Required):
 * - Schema_Snapshot
 * - Formula_Inventory
 * - Column_Classification
 *
 * Algorithm (Optimized – Logic Unchanged):
 * 1. Build schema and formula lookup maps
 * 2. Resolve references in memory
 * 3. Single batch write
 * 4. Batch separator formatting
 *
 * Output Contract:
 * - Sheet: Derived_Column_Logic (fully regenerated each run)
 * - Layout:
 *   Row 1  → Header (values only)
 *   Row 2–3 → Reserved empty rows
 *   Row 4+ → Data
 *   Two empty separator rows inserted when Sheet_Name changes
 *
 * Cosmetic Rules (Intentional & Limited):
 * - Separator rows are visually marked in columns A–H only
 * - Color: #FFFF00 (bright yellow)
 * - Purpose: improve human scanability between derived blocks
 *
 * Idempotency:
 * - Safe to re-run; output is cleared and rebuilt deterministically
 */

function generateDerivedColumnLogic() {

  const dataSS = SpreadsheetApp.getActiveSpreadsheet();
  const metaSS = getMetadataSpreadsheet_();

  const SCHEMA_SHEET   = metaSS.getSheetByName('Schema_Snapshot');
  const FORMULA_SHEET  = metaSS.getSheetByName('Formula_Inventory');
  const CLASS_SHEET    = metaSS.getSheetByName('Column_Classification');

  if (!SCHEMA_SHEET || !FORMULA_SHEET || !CLASS_SHEET) {
    throw new Error('Required dependency sheet missing');
  }

  const SCHEMA   = SCHEMA_SHEET.getDataRange().getValues();
  const FORMULAS = FORMULA_SHEET.getDataRange().getValues();
  const CLASS    = CLASS_SHEET.getDataRange().getValues();

  let out = metaSS.getSheetByName('Derived_Column_Logic');
  if (!out) out = metaSS.insertSheet('Derived_Column_Logic');
  out.clear();

  const headers = [
    'Sheet_Name',
    'Column_Index',
    'Column_Letter',
    'Column_Name',
    'Semantic_Class',
    'Formula_A1_Text',
    'Formula_R1C1_Text',
    'Resolved_References'
  ];

  out.getRange(1, 1, 1, headers.length).setValues([headers]);

  /* ===================== BUILD LOOKUPS ===================== */

  const schemaMap = {};
  const schemaLetterMap = {};

  for (let i = 1; i < SCHEMA.length; i++) {
    const [sheet, idx, letter, name] = SCHEMA[i];
    schemaMap[`${sheet}|${idx}`] = { letter, name };
    schemaLetterMap[`${sheet}|${letter}`] = { idx, name };
  }

  const formulaMap = {};
  for (let i = 1; i < FORMULAS.length; i++) {
    const [sheet, colIdx, , , a1, r1c1] = FORMULAS[i];
    formulaMap[`${sheet}|${colIdx}`] = { a1, r1c1 };
  }

  /* ===================== MAIN LOOP ===================== */

  let output = [];
  let separatorRows = [];
  let lastSheet = null;

  for (let i = 1; i < CLASS.length; i++) {

    const [sheet, colIdx, colLetter, colName, semantic] = CLASS[i];

    if (!['DERIVED_LOCAL', 'DERIVED_LOOKUP'].includes(semantic)) continue;

    if (lastSheet !== null && sheet !== lastSheet) {
      separatorRows.push(output.length + 4);
      separatorRows.push(output.length + 5);
      output.push(new Array(headers.length).fill(''));
      output.push(new Array(headers.length).fill(''));
    }

    const f = formulaMap[`${sheet}|${colIdx}`];
    if (!f) continue;

    const refs = new Set();

    /* ---------- R1C1 same-sheet references ---------- */

    const r1c1 = f.r1c1 || '';
    const rcMatches = r1c1.match(/RC\[[+-]?\d+\]/g) || [];

    rcMatches.forEach(token => {
      const offset = parseInt(token.match(/[+-]?\d+/)[0], 10);
      const targetIdx = colIdx + offset;
      const meta = schemaMap[`${sheet}|${targetIdx}`];
      if (meta) {
        refs.add(`${sheet}.${meta.letter} (${meta.name})`);
      }
    });

    /* ---------- A1 cross-sheet references ---------- */

    const a1 = f.a1 || '';
    const a1Matches = a1.match(/([A-Z0-9_]+)!\$?[A-Z]+/gi) || [];

    a1Matches.forEach(ref => {
      const [refSheet, colPart] = ref.split('!');
      const colLetterRef = colPart.replace(/[^A-Z]/gi, '');
      const meta = schemaLetterMap[`${refSheet}|${colLetterRef}`];
      if (meta) {
        refs.add(`${refSheet}.${colLetterRef} (${meta.name})`);
      }
    });

    output.push([
      sheet,
      colIdx,
      colLetter,
      colName,
      semantic,
      f.a1,
      f.r1c1,
      Array.from(refs).join('\n')
    ]);

    lastSheet = sheet;
  }

  /* ===================== SINGLE BATCH WRITE ===================== */

  if (output.length > 0) {
    out.getRange(4, 1, output.length, headers.length)
       .setValues(output);
  }

  /* ===================== BATCH SEPARATOR FORMATTING ===================== */

  separatorRows.forEach(r => {
    out.getRange(r, 1, 1, headers.length)
       .setBackground('#FFFF00');
  });
}