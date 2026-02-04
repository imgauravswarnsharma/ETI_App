function applyBaseFormatting_ToAllExistingSheets() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  for (const sheet of sheets) {
    applyBaseFormatting_(sheet);
  }

  console.log(
    `Base formatting applied to ${sheets.length} existing sheets`
  );
}