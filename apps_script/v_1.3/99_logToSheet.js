function logToSheet_(message, level) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shName = 'Script_Logs';
  let sh = ss.getSheetByName(shName);

  if (!sh) {
    sh = ss.insertSheet(shName);
    sh.appendRow(['Timestamp', 'Level', 'Message']);
  }

  sh.appendRow([
    new Date(),
    level || 'INFO',
    message
  ]);
}
