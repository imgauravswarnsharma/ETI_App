/**
 * Config_Named_Ranges — Manifest Extract v1
 * Read-only, safe
 */

function extractNamedRangesConfig() {

  const dataSS = SpreadsheetApp.getActiveSpreadsheet();
  const metaSS = getMetadataSpreadsheet_();

  let out = metaSS.getSheetByName('Config_Named_Ranges');
  if (!out) out = metaSS.insertSheet('Config_Named_Ranges');
  out.clear();

  const headers = [
    'Range_Name',
    'Sheet_Name',
    'Range_A1'
  ];

  out.getRange(1, 1, 1, headers.length)
     .setValues([headers])
     .setFontWeight('bold');

  let writeRow = 4;

  const namedRanges = dataSS.getNamedRanges();

  namedRanges.forEach(nr => {
    const range = nr.getRange();
    const sheet = range.getSheet();

    out.getRange(writeRow, 1, 1, headers.length).setValues([[
      nr.getName(),
      sheet.getName(),
      range.getA1Notation()
    ]]);

    writeRow++;
  });
}