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

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const table = ss.getSheetByName("Column_Access_Control");

  const now = new Date();

  const lastRow = table.getLastRow();

  let existing = [];

  if(lastRow > 1){
    existing = table.getRange(2,1,lastRow-1,4).getValues();
  }

  const map = {};

  existing.forEach(r => {

    const sheet = r[0];
    const columnName = r[3];

    if(!sheet || !columnName) return;

    const key = sheet + "::" + columnName;

    map[key] = true;

  });


  const rowsToInsert = [];

  ss.getSheets().forEach(sheet => {

    const sheetName = sheet.getName();
    const lastCol = sheet.getLastColumn();

    if(lastCol === 0) return;

    const headers = sheet.getRange(1,1,1,lastCol).getValues()[0];

    headers.forEach((header,i) => {

      if(!header) return;

      const key = sheetName + "::" + header;

      if(map[key]) return;

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

    });

  });


  if(rowsToInsert.length){

    table.getRange(
      table.getLastRow()+1,
      1,
      rowsToInsert.length,
      rowsToInsert[0].length
    ).setValues(rowsToInsert);

  }

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
