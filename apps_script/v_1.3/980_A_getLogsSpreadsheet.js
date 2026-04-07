function getLogsSpreadsheet_() {

  const FILE_NAME = 'ETI_App_v1.3_Logs';

  const files = DriveApp.getFilesByName(FILE_NAME);

  if (!files.hasNext()) {
    throw new Error(`Logs file "${FILE_NAME}" not found`);
  }

  return SpreadsheetApp.open(files.next());
}