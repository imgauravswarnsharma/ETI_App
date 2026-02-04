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

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const SCHEMA   = ss.getSheetByName('Schema_Snapshot').getDataRange().getValues();
  const FORMULAS = ss.getSheetByName('Formula_Inventory').getDataRange().getValues();
  const CLASS    = ss.getSheetByName('Column_Classification').getDataRange().getValues();

  let out = ss.getSheetByName('Derived_Column_Logic');
  if (!out) out = ss.insertSheet('Derived_Column_Logic');
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

  // Header values only — formatting handled elsewhere
  out.getRange(1, 1, 1, headers.length).setValues([headers]);

  /* ===================== BUILD LOOKUPS ===================== */

  // key: Sheet|ColIndex → { letter, name }
  const schemaMap = {};
  for (let i = 1; i < SCHEMA.length; i++) {
    const [sheet, idx, letter, name] = SCHEMA[i];
    schemaMap[`${sheet}|${idx}`] = { letter, name };
  }

  // key: Sheet|ColIndex → { a1, r1c1 }
  const formulaMap = {};
  for (let i = 1; i < FORMULAS.length; i++) {
    const [sheet, colIdx, , , a1, r1c1] = FORMULAS[i];
    formulaMap[`${sheet}|${colIdx}`] = { a1, r1c1 };
  }

  let writeRow = 4;
  let lastSheet = null;

  /* ===================== MAIN LOOP ===================== */

  for (let i = 1; i < CLASS.length; i++) {

    const [sheet, colIdx, colLetter, colName, semantic] = CLASS[i];

    if (!['DERIVED_LOCAL', 'DERIVED_LOOKUP'].includes(semantic)) continue;

    // ---- Insert visual separator between sheet blocks ----
    if (lastSheet !== null && sheet !== lastSheet) {

      // Mark the two gap rows (A–H only)
      out.getRange(writeRow, 1, 2, headers.length)
         .setBackground('#FFFF00');

      writeRow += 2;
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
      const meta = SCHEMA.find(
        r => r[0] === refSheet && r[2] === colLetterRef
      );
      if (meta) {
        refs.add(`${refSheet}.${meta[2]} (${meta[3]})`);
      }
    });

    out.getRange(writeRow, 1, 1, headers.length).setValues([[
      sheet,
      colIdx,
      colLetter,
      colName,
      semantic,
      f.a1,
      f.r1c1,
      Array.from(refs).join('\n')
    ]]);

    lastSheet = sheet;
    writeRow++;
  }
}
