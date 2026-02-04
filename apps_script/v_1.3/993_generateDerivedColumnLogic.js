/**
 * Derived_Column_Logic — Reference Resolution v1
 * Reads inert formulas and expands dependencies
 */

function generateDerivedColumnLogic() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const SCHEMA = ss.getSheetByName('Schema_Snapshot').getDataRange().getValues();
  const FORMULAS = ss.getSheetByName('Formula_Inventory').getDataRange().getValues();
  const CLASS = ss.getSheetByName('Column_Classification').getDataRange().getValues();

  let out = ss.getSheetByName('Derived_Column_Logic');
  if (!out) out = ss.insertSheet('Derived_Column_Logic');
  out.clear();

  // ---- Header ----
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

  out.getRange(1, 1, 1, headers.length)
     .setValues([headers])
     .setFontWeight('bold');

  // ---- Build schema lookup ----
  // key: Sheet|ColIndex → { letter, name }
  const schemaMap = {};
  for (let i = 1; i < SCHEMA.length; i++) {
    const [sheet, idx, letter, name] = SCHEMA[i];
    schemaMap[`${sheet}|${idx}`] = { letter, name };
  }

  // ---- Build formula lookup ----
  const formulaMap = {};
  for (let i = 1; i < FORMULAS.length; i++) {
    const [sheet, colIdx,,,, a1, r1c1] = FORMULAS[i];
    formulaMap[`${sheet}|${colIdx}`] = { a1, r1c1 };
  }

  let writeRow = 4;
  let lastSheet = null;

  for (let i = 1; i < CLASS.length; i++) {

    const [sheet, colIdx, colLetter, colName, semantic] = CLASS[i];

    if (!['DERIVED_LOCAL', 'DERIVED_LOOKUP'].includes(semantic)) continue;

    if (lastSheet !== null && sheet !== lastSheet) {
      writeRow += 2;
    }

    const f = formulaMap[`${sheet}|${colIdx}`];
    if (!f) continue;

    const refs = new Set();

    // ---- R1C1 references (same sheet) ----
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

    // ---- A1 cross-sheet references ----
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
