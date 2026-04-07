/*
================================================================================================
THIS FILE CONTAINS TESTED, DEPRECATED, REJECTED, TEMPRORY, DEBUGGING, SANITY, TESTING AND POSSIBLE FUTUTRE FEATURES SCRIPTS OR WORK IN PROGESS.

THIS SCRIPT DOES NOT CONTAIN ANY FUNCTION WHICH CONTAINS ACTIVE OR USED FUNCTIONS AS OF NOW AND WORKS AS VAULT FOR REFERENCES IF NEEDED BY DEVELOPER
================================================================================================
*/

function applyBaseFormatting_(sheet) {

  /* ===============================
     1. Freeze first two rows
     =============================== */
  sheet.setFrozenRows(2);

  /* ===============================
     2. Bold header row (Row 1)
     =============================== */
  const headerRange = sheet.getRange(
    1,
    1,
    1,
    sheet.getMaxColumns()
  );
  headerRange.setFontWeight('bold');

  /* ===============================
     3. Global alignment
     =============================== */
  const fullRange = sheet.getRange(
    1,
    1,
    sheet.getMaxRows(),
    sheet.getMaxColumns()
  );

  fullRange
    .setHorizontalAlignment('left')
    .setVerticalAlignment('top');
}



/*
========================================================================

++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

========================================================================
*/

// MANUAL TRIGGER
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



/*
========================================================================

++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

========================================================================
*/

//
/**
 * SAFE regeneration into a shadow sheet
 * Never overwrites original
 */

function regenerateSheetPreview(sheetName) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheetName = `${sheetName}`;

  const SCHEMA = ss.getSheetByName('Schema_Snapshot').getDataRange().getValues();
  const FORMULAS = ss.getSheetByName('Formula_Inventory').getDataRange().getValues();

  // ---- Skip header row safely ----
  const cols = SCHEMA.slice(1).filter(r => r[0] === sheetName);
  if (cols.length === 0) {
    throw new Error(`Sheet not found in Schema_Snapshot: ${sheetName}`);
  }

  // ---- Create / reset preview sheet ----
  let sh = ss.getSheetByName(targetSheetName);
  if (!sh) sh = ss.insertSheet(targetSheetName);
  sh.clear();

  // ---- Write headers ----
  const headers = cols.map(r => r[3]);
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ---- Build formula lookup ----
  const formulaMap = {};
  for (let i = 1; i < FORMULAS.length; i++) {
    const [s, colIdx,,,, a1] = FORMULAS[i];
    if (s === sheetName && a1) {
      formulaMap[colIdx] = a1;
    }
  }

  // ---- Inject formulas into row 2 ----
  cols.forEach(col => {
    const colIdx = col[1];
    const formulaText = formulaMap[colIdx];
    if (formulaText) {
      sh.getRange(2, colIdx).setFormula('=' + formulaText);
    }
  });
}



/*
========================================================================

++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

========================================================================
*/

//GENERATE COUNTER FOR HUMAN ITEM ID
function generateItemID_H() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const itemsSh = ss.getSheetByName('Lookup_Items');
  const ctrlSh  = ss.getSheetByName('Counter_Control');
  if (!itemsSh || !ctrlSh) throw new Error('Required sheet not found');

  // Read control table
  const ctrlData = ctrlSh.getDataRange().getValues();
  const ctrlHeader = ctrlData[0];

  const colCtrl = n => ctrlHeader.indexOf(n);

  const IDX_CTRL = {
    entity: colCtrl('Entity_Key'),
    total: colCtrl('Total_Counter')
  };

  for (const [k, v] of Object.entries(IDX_CTRL)) {
    if (v === -1) throw new Error(`Missing control column: ${k}`);
  }

  // Locate ITEM row
  let ctrlRow = -1;
  for (let i = 1; i < ctrlData.length; i++) {
    if (ctrlData[i][IDX_CTRL.entity] === 'ITEM') {
      ctrlRow = i;
      break;
    }
  }
  if (ctrlRow === -1) throw new Error('ITEM row not found in Counter_Control');

  let counter = Number(ctrlData[ctrlRow][IDX_CTRL.total]) || 0;

  // Read items
  const data = itemsSh.getDataRange().getValues();
  if (data.length < 2) return;

  const header = data[0];
  const col = n => header.indexOf(n);

  const IDX_ITEMS = {
    itemIdM: col('Item_ID_M'),
    itemIdH: col('Item_ID_H')
  };

  for (const [k, v] of Object.entries(IDX_ITEMS)) {
    if (v === -1) throw new Error(`Missing item column: ${k}`);
  }

  let wrote = false;

  for (let i = 1; i < data.length; i++) {
    const r = data[i];

    if (r[IDX_ITEMS.itemIdM] && !r[IDX_ITEMS.itemIdH]) {
      counter += 1;
      const humanId = 'ITEM-' + String(counter).padStart(6, '0');
      itemsSh.getRange(i + 1, IDX_ITEMS.itemIdH + 1).setValue(humanId);
      wrote = true;
    }
  }

  // Persist counter only if changes were made
  if (wrote) {
    ctrlSh.getRange(ctrlRow + 1, IDX_CTRL.total + 1).setValue(counter);
  }
}

/*
========================================================================

++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

========================================================================
*/

// 

function generateContextDropdowns() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const mapSheet = ss.getSheetByName("Mapping_Item_Brand_Product");
  const lookupBrands = ss.getSheetByName("Lookup_Brands");
  const lookupProducts = ss.getSheetByName("Lookup_Products");
  const stagingBrands = ss.getSheetByName("Staging_Lookup_Brands");
  const stagingProducts = ss.getSheetByName("Staging_Lookup_Products");
  const ctxSheet = ss.getSheetByName("Context_Dropdowns");

  const now = new Date();

  // Reset table
  ctxSheet.clearContents();

  const header = [
    "Context_Type",
    "Item_ID_Machine",
    "Brand_ID_Machine",
    "Product_ID_Machine",
    "Display_Value",
    "Priority",
    "Source",
    "Created_At"
  ];

  ctxSheet.getRange(1,1,1,header.length).setValues([header]);

  const rows = [];
  const seen = new Set();

  // -------------------------
  // Mapping: Item-Brand
  // -------------------------

  const mapData = mapSheet.getDataRange().getValues();
  const mapHeader = mapData[0];

  const idxItem = mapHeader.indexOf("Item_ID_Machine");
  const idxBrand = mapHeader.indexOf("Brand_ID_Machine");
  const idxProduct = mapHeader.indexOf("Product_ID_Machine");
  const idxBrandName = mapHeader.indexOf("Brand_Name_Canonical");
  const idxProductName = mapHeader.indexOf("Product_Name_Canonical");
  const idxActive = mapHeader.indexOf("Is_Mapping_Active");

  for (let i = 1; i < mapData.length; i++) {

    const r = mapData[i];

    if (idxActive < 0 || r[idxActive] !== true) continue;

    const item = r[idxItem];
    const brand = r[idxBrand];
    const product = r[idxProduct];

    const brandName = r[idxBrandName];
    const productName = r[idxProductName];

    const keyBrand = "IB_" + item + "_" + brand;

    if (!seen.has(keyBrand)) {

      rows.push([
        "ITEM_BRAND",
        item,
        brand,
        "",
        brandName,
        1,
        "Mapping_Item_Brand_Product",
        now
      ]);

      seen.add(keyBrand);

    }

    if (product && productName) {

      const keyProduct = "IBP_" + item + "_" + brand + "_" + product;

      if (!seen.has(keyProduct)) {

        rows.push([
          "ITEM_BRAND_PRODUCT",
          item,
          brand,
          product,
          productName,
          1,
          "Mapping_Item_Brand_Product",
          now
        ]);

        seen.add(keyProduct);

      }

    }

  }

  // -------------------------
  // Lookup Brands
  // -------------------------

  const lbData = lookupBrands.getDataRange().getValues();
  const lbHeader = lbData[0];

  const lbBrandID = lbHeader.indexOf("Brand_ID_Machine");
  const lbBrandName = lbHeader.indexOf("Brand_Name");
  const lbActive = lbHeader.indexOf("Is_Active");

  for (let i = 1; i < lbData.length; i++) {

    const r = lbData[i];

    if (lbActive < 0 || r[lbActive] !== true) continue;

    rows.push([
      "ITEM_BRAND",
      "",
      r[lbBrandID],
      "",
      r[lbBrandName],
      2,
      "Lookup_Brands",
      now
    ]);

  }

  // -------------------------
  // Lookup Products
  // -------------------------

  const lpData = lookupProducts.getDataRange().getValues();
  const lpHeader = lpData[0];

  const lpProductID = lpHeader.indexOf("Product_ID_Machine");
  const lpProductName = lpHeader.indexOf("Product_Name");
  const lpActive = lpHeader.indexOf("Is_Active");

  for (let i = 1; i < lpData.length; i++) {

    const r = lpData[i];

    if (lpActive < 0 || r[lpActive] !== true) continue;

    rows.push([
      "ITEM_BRAND_PRODUCT",
      "",
      "",
      r[lpProductID],
      r[lpProductName],
      2,
      "Lookup_Products",
      now
    ]);

  }

  // -------------------------
  // Staging Brands
  // -------------------------

  const sbData = stagingBrands.getDataRange().getValues();
  const sbHeader = sbData[0];

  const sbBrandID = sbHeader.indexOf("Staging_Brand_ID_Machine");
  const sbBrandName = sbHeader.indexOf("Brand_Name_Entered");
  const sbActive = sbHeader.indexOf("Is_Active");
  const sbPromoted = sbHeader.indexOf("Is_Lookup_Promoted");

  for (let i = 1; i < sbData.length; i++) {

    const r = sbData[i];

    if (sbActive < 0 || r[sbActive] !== true) continue;
    if (sbPromoted >= 0 && r[sbPromoted] === true) continue;

    rows.push([
      "ITEM_BRAND",
      "",
      r[sbBrandID],
      "",
      r[sbBrandName],
      3,
      "Staging_Brands",
      now
    ]);

  }

  // -------------------------
  // Staging Products
  // -------------------------

  const spData = stagingProducts.getDataRange().getValues();
  const spHeader = spData[0];

  const spProductID = spHeader.indexOf("Staging_Product_ID_Machine");
  const spProductName = spHeader.indexOf("Product_Name_Entered");
  const spActive = spHeader.indexOf("Is_Active");
  const spPromoted = spHeader.indexOf("Is_Lookup_Promoted");

  for (let i = 1; i < spData.length; i++) {

    const r = spData[i];

    if (spActive < 0 || r[spActive] !== true) continue;
    if (spPromoted >= 0 && r[spPromoted] === true) continue;

    rows.push([
      "ITEM_BRAND_PRODUCT",
      "",
      "",
      r[spProductID],
      r[spProductName],
      3,
      "Staging_Products",
      now
    ]);

  }

  // -------------------------
  // SORT ROWS (Priority → Display_Value)
  // -------------------------

  rows.sort((a, b) => {

    const priorityDiff = a[5] - b[5];

    if (priorityDiff !== 0) return priorityDiff;

    return String(a[4]).localeCompare(String(b[4]));

  });

  if (rows.length > 0) {

    ctxSheet.getRange(2,1,rows.length,rows[0].length).setValues(rows);

  }

}


/*
========================================================================

++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

========================================================================
*/

// 
/*
=========================================================
ACCESS GOVERNANCE MODULE
=========================================================

Pipeline Order:

metadata_pipeline_()
reconcile_access_control_metadata_()
apply_access_governance_()

=========================================================
*/


/*
---------------------------------------------------------
Reconcile Access Metadata
---------------------------------------------------------
Ensures every sheet column exists in
Column_Access_Control table.

Missing columns are inserted with:

Access_Level = DEV_L2
Source = SYSTEM
Admin_Action = REVIEW
---------------------------------------------------------
*/

function reconcile_access_control_metadata_(){

  const dataSS = SpreadsheetApp.getActiveSpreadsheet();   // READ (DATA)
  const supportSS = getSupportSpreadsheet_();              // WRITE (SUPPORT)

  const TABLE_NAME = "Column_Access_Control";
  const table = supportSS.getSheetByName(TABLE_NAME);

  if(!table) throw new Error(`${TABLE_NAME} not found in support file`);

  const now = new Date();

  const lastRow = table.getLastRow();

  let existingData = [];
  if(lastRow > 1){
    existingData = table.getRange(2,1,lastRow-1,10).getValues();
  }

  /* =========================
     BUILD EXISTING MAP
  ========================= */

  const existingMap = {};
  const seenMap = {};

  existingData.forEach((r, idx) => {

    const sheet = r[0];
    const columnName = r[3];

    if(!sheet || !columnName) return;

    const key = sheet + "::" + columnName;

    existingMap[key] = { rowIndex: idx + 2 };
    seenMap[key] = false;

  });

  /* =========================
     SCAN CURRENT SCHEMA
  ========================= */

  const rowsToInsert = [];

  dataSS.getSheets().forEach(sheet => {

    const sheetName = sheet.getName();
    const lastCol = sheet.getLastColumn();

    if(lastCol === 0) return;

    const headers = sheet.getRange(1,1,1,lastCol).getValues()[0];

    headers.forEach((header,i) => {

      if(!header) return;

      const key = sheetName + "::" + header;

      if(!existingMap[key]){

        rowsToInsert.push([
          sheetName,
          i+1,
          columnToLetter_(i+1),
          header,
          "DEV_L2",
          "SYSTEM",
          "REVIEW",
          now,
          "",
          ""
        ]);

      } else {
        seenMap[key] = true;
      }

    });

  });

  /* =========================
     INSERT NEW
  ========================= */

  if(rowsToInsert.length){

    table.getRange(
      table.getLastRow()+1,
      1,
      rowsToInsert.length,
      rowsToInsert[0].length
    ).setValues(rowsToInsert);

  }

  /* =========================
     SOFT DELETE
  ========================= */

  Object.keys(existingMap).forEach(key => {

    if(!seenMap[key]){

      const rowIndex = existingMap[key].rowIndex;

      table.getRange(rowIndex, 9).setValue(now);  // Updated_At
      table.getRange(rowIndex, 10).setValue(
        "AUTO_DELETED: Column no longer exists"
      );

    }

  });

}



/*
---------------------------------------------------------
Apply Access Governance
---------------------------------------------------------
Locks or unlocks columns based on:

Access_Mode switch
Column_Access_Control metadata
---------------------------------------------------------
*/

function apply_access_governance_(){

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const switches = getAutomationSwitchMap_();
  const mode = switches["Access_Mode"] || "BACKEND";

  console.log("Access Mode:", mode);


  /*
  -----------------------------------------------------
  GOD MODE
  -----------------------------------------------------
  */

  if(mode === "GOD"){

    removeAllProtections_(ss);
    return;

  }


  /*
  -----------------------------------------------------
  Access Level Map
  -----------------------------------------------------
  */

  const levelMap = {
    USER:1,
    DEV_L1:2,
    DEV_L2:3
  };

  const modeLevel = levelMap[mode] || 0;


  /*
  -----------------------------------------------------
  Read Access Table
  -----------------------------------------------------
  */

  const table = ss.getSheetByName("Column_Access_Control");

  const lastRow = table.getLastRow();

  if(lastRow < 2) return;

  const data = table.getRange(2,1,lastRow-1,7).getValues();


  const accessMap = {};

  data.forEach(r => {

    const sheet = r[0];
    const column = r[3];
    const level = r[4];
    const action = r[6];

    if(action !== "REVIEWED") return;

    if(!accessMap[sheet]) accessMap[sheet] = {};

    accessMap[sheet][column] = level;

  });


  /*
  -----------------------------------------------------
  Apply Protections
  -----------------------------------------------------
  */

  ss.getSheets().forEach(sheet => {

    const sheetName = sheet.getName();

    const lastCol = sheet.getLastColumn();
    const lastRow = sheet.getLastRow();

    if(lastCol === 0) return;


    /*
    Remove existing protections
    */

    sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)
    .forEach(p => p.remove());

    sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)
    .forEach(p => p.remove());


    /*
    Protect entire sheet
    */

    const sheetProtection = sheet.protect();
    sheetProtection.removeEditors(sheetProtection.getEditors());


    /*
    Read headers
    */

    const headers = sheet.getRange(1,1,1,lastCol).getValues()[0];


    headers.forEach((header,i) => {

      const level = accessMap[sheetName]?.[header];

      if(!level) return;

      const colLevel = levelMap[level];

      if(colLevel <= modeLevel){

        const range = sheet.getRange(3,i+1,lastRow-2,1);

        const protection = range.protect();

        protection.removeEditors(protection.getEditors());
        protection.setWarningOnly(false);

      }

    });


    /*
    Always protect header rows
    */

    const headerRange = sheet.getRange(1,1,2,lastCol);

    const headerProtection = headerRange.protect();

    headerProtection.removeEditors(headerProtection.getEditors());


    /*
    Special Rule:
    Protect Automation_Control column A
    for modes below DEV_L2
    */

    if(sheetName === "Automation_Control" && modeLevel < 3){

      const protectRange = sheet.getRange(3,1,lastRow-2,1);

      const p = protectRange.protect();

      p.removeEditors(p.getEditors());

    }

  });

}




/*
---------------------------------------------------------
Remove All Protections (GOD Mode)
---------------------------------------------------------
*/

function removeAllProtections_(ss){

  ss.getSheets().forEach(sheet => {

    sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)
    .forEach(p => p.remove());

    sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)
    .forEach(p => p.remove());

  });

}




/*
---------------------------------------------------------
Helper — Column Index to Letter
---------------------------------------------------------
*/

function columnToLetter_(column){

  let temp;
  let letter = "";

  while(column > 0){

    temp = (column - 1) % 26;

    letter = String.fromCharCode(temp + 65) + letter;

    column = (column - temp - 1) / 26;

  }

  return letter;

}



/*
========================================================================

++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

========================================================================
*/

// WORKING CLEAN LOGGER (ALREADY TESTED)


/**
 * Logger Name: ETI_Structured_Logger
 * Version: v1.3
 * Status: FIXED (Controller-driven, modular, safe)
 */


/*
-------------------------------------
ACTION LOGGER (DETAILED)
-------------------------------------
*/
function ETI_log_(payload) {

  // ---- Guard: Action Logging ----
  if (!isActionLogEnabled_()) return;

  // ---- Debug Console ----
  if (isDebugModeEnabled_()) {
    console.log(
      `[${payload.scriptName || ''}]`,
      payload.level || 'INFO',
      payload.action || '',
      payload.rowNumber ? `ROW ${payload.rowNumber}` : '',
      payload.details || ''
    );
  }

  const logSS = getLogsSpreadsheet_();
  const sheetName = 'Action_Logs';

  let sh = logSS.getSheetByName(sheetName);

  if (!sh) {
    sh = logSS.insertSheet(sheetName);

    sh.appendRow([
      'Timestamp',

      // ---- EXECUTION TRACE ----
      'Pipeline_Name',
      'Function_Name',
      'Sheet_Name',
      'Switch_Name',

      // ---- CONTEXT ----
      'Execution_ID',
      'Trigger_Type',
      'Run_Context',

      // ---- ACTION ----
      'Level',
      'Action',
      'Step_Name',
      'Row_Number',
      'Details',

      // ---- ERROR ----
      'Error_Message'
    ]);
  }

  sh.appendRow([
    new Date(),

    // ---- EXECUTION TRACE ----
    payload.pipelineName || '',
    payload.functionName || '',
    payload.sheetName || '',
    payload.switchName || '',

    // ---- CONTEXT ----
    payload.executionId || '',
    payload.triggerType || 'MANUAL',
    payload.runContext || '',

    // ---- ACTION ----
    payload.level || 'INFO',
    payload.action || '',
    payload.stepName || '',
    payload.rowNumber || '',
    payload.details || '',

    // ---- ERROR ----
    payload.errorMessage || ''
  ]);
}


/*
-------------------------------------
EXECUTION LOGGER (HIGH LEVEL)
-------------------------------------
*/
function ETI_logExecution_(payload) {

  // ---- Guard: Execution Logging ----
  if (!isExecutionLogEnabled_()) return;

  const logSS = getLogsSpreadsheet_();
  const sheetName = 'Execution_Logs';

  let sh = logSS.getSheetByName(sheetName);

  if (!sh) {
    sh = logSS.insertSheet(sheetName);

    sh.appendRow([
      'Timestamp_Start',
      'Timestamp_End',

      // ---- EXECUTION TRACE ----
      'Pipeline_Name',
      'Function_Name',
      'Switch_Name',

      // ---- CONTEXT ----
      'Execution_ID',
      'Trigger_Source',
      'Run_Context',

      // ---- RESULT ----
      'Status',
      'Duration_ms',
      'Error_Message'
    ]);
  }

  sh.appendRow([
    payload.startTime || '',
    payload.endTime || '',

    // ---- EXECUTION TRACE ----
    payload.pipelineName || '',
    payload.functionName || '',
    payload.switchName || '',

    // ---- CONTEXT ----
    payload.executionId || '',
    payload.triggerSource || '',
    payload.runContext || '',

    // ---- RESULT ----
    payload.status || '',
    payload.durationMs || '',
    payload.errorMessage || ''
  ]);
}
