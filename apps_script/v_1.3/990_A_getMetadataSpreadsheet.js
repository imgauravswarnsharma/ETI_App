function getMetadataSpreadsheet_() {

  const FILE_NAME = 'ETI_App_v1.3_Metadata';

  const files = DriveApp.getFilesByName(FILE_NAME);

  if (!files.hasNext()) {
    throw new Error(`Metadata file "${FILE_NAME}" not found`);
  }

  return SpreadsheetApp.open(files.next());
}