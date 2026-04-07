function reconcile_access_metadata_() {

  const SCRIPT_NAME = "reconcile_access_metadata_";

  const SCHEMA_SHEET = "Schema_Snapshot";
  const ACCESS_SHEET = "Column_Access_Control";

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const schemaSheet = ss.getSheetByName(SCHEMA_SHEET);
  if (!schemaSheet) throw new Error("Schema_Snapshot sheet missing");

  let accessSheet = ss.getSheetByName(ACCESS_SHEET);

  /* =========================
     CREATE ACCESS TABLE IF MISSING
     ========================= */

  if (!accessSheet) {

    accessSheet = ss.insertSheet(ACCESS_SHEET);

    const headers = [[
      "Column_Key",
      "Sheet_Name",
      "Column_Name",
      "Access_Level",
      "Admin_Action",
      "Source",
      "Created_At",
      "Updated_At",
      "Notes"
    ]];

    accessSheet.getRange(1,1,1,headers[0].length).setValues(headers);

    console.log(`[${SCRIPT_NAME}] Column_Access_Control created`);
  }

  /* =========================
     LOAD SCHEMA
     ========================= */

  const schemaData = schemaSheet.getDataRange().getValues();

  if (schemaData.length < 2) {
    console.log(`[${SCRIPT_NAME}] Schema empty`);
    return;
  }

  const schemaHeader = schemaData[0];

  const sc = name => schemaHeader.indexOf(name);

  const IDX_SCHEMA = {

    sheet: sc("Sheet_Name"),
    colName: sc("Column_Name")

  };

  if (IDX_SCHEMA.sheet === -1 || IDX_SCHEMA.colName === -1) {
    throw new Error("Schema_Snapshot missing required columns");
  }

  /* =========================
     LOAD ACCESS TABLE
     ========================= */

  const accessData = accessSheet.getDataRange().getValues();

  const existingKeys = new Set();

  for (let i = 1; i < accessData.length; i++) {

    const key = accessData[i][0];

    if (key) existingKeys.add(key);

  }

  /* =========================
     DISCOVERY LOOP
     ========================= */

  const rowsToInsert = [];

  for (let i = 1; i < schemaData.length; i++) {

    const sheetName = schemaData[i][IDX_SCHEMA.sheet];
    const columnName = schemaData[i][IDX_SCHEMA.colName];

    if (!sheetName || !columnName) continue;

    const columnKey = `${sheetName}::${columnName}`;

    if (existingKeys.has(columnKey)) continue;

    rowsToInsert.push([
      columnKey,
      sheetName,
      columnName,
      "DEV_L2",
      "REVIEW",
      "SYSTEM",
      new Date(),
      "",
      ""
    ]);

    existingKeys.add(columnKey);

  }

  /* =========================
     BATCH INSERT
     ========================= */

  if (rowsToInsert.length > 0) {

    accessSheet.getRange(
      accessSheet.getLastRow() + 1,
      1,
      rowsToInsert.length,
      rowsToInsert[0].length
    ).setValues(rowsToInsert);

  }

  console.log(`[${SCRIPT_NAME}] new columns detected: ${rowsToInsert.length}`);

}
