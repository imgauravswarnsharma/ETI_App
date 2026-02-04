/**
 * Config_Data_Validation — Manifest Extract v1
 * Read-only, safe
 */

function extractDataValidationConfig() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const SCHEMA = ss.getSheetByName('Schema_Snapshot')
                   .getDataRange()
                   .getValues();

  let out = ss.getSheetByName('Config_Data_Validation');
  if (!out) out = ss.insertSheet('Config_Data_Validation');
  out.clear();

  const headers = [
    'Sheet_Name',
    'Column_Letter',
    'Column_Name',
    'Rule_Type',
    'Criteria',
    'Allowed_Values',
    'Help_Text',
    'Strict_Flag'
  ];

  out.getRange(1, 1, 1, headers.length)
     .setValues([headers])
     .setFontWeight('bold');

  let writeRow = 4;
  let lastSheet = null;

  // Build schema lookup: Sheet|ColIndex → { letter, name }
  const schemaMap = {};
  for (let i = 1; i < SCHEMA.length; i++) {
    const [sheet, idx, letter, name] = SCHEMA[i];
    schemaMap[`${sheet}|${idx}`] = { letter, name };
  }

  ss.getSheets().forEach(sh => {

    const sheetName = sh.getName();
    const maxCols = sh.getMaxColumns();

    for (let col = 1; col <= maxCols; col++) {

      const rule = sh.getRange(2, col).getDataValidation();
      if (!rule) continue;

      const meta = schemaMap[`${sheetName}|${col}`];
      if (!meta) continue;

      if (lastSheet !== null && sheetName !== lastSheet) {
        writeRow += 2;
      }

      const type = rule.getCriteriaType();
      const args = rule.getCriteriaValues();

      let criteria = type;
      let allowed = '';

      if (args && args.length) {
        allowed = args.map(v => {
          if (Array.isArray(v)) return v.flat().join(', ');
          return String(v);
        }).join(' | ');
      }

      out.getRange(writeRow, 1, 1, headers.length).setValues([[
        sheetName,
        meta.letter,
        meta.name,
        criteria,
        allowed,
        allowed,
        rule.getHelpText() || '',
        rule.getAllowInvalid() ? 'WARN' : 'STRICT'
      ]]);

      lastSheet = sheetName;
      writeRow++;
    }
  });
}
