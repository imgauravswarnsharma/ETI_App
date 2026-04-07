function run_generateScriptCallMap_INTERNAL(){
  generateScriptCallMap_INTERNAL_();
}


function generateScriptCallMap_INTERNAL_(){

  const metadataSS = getMetadataSpreadsheet_();

  const FUNC_SHEET = "Script_Function_Inventory";
  const OUT_SHEET = "Script_Call_Map_INTERNAL";

  const funcSheet = metadataSS.getSheetByName(FUNC_SHEET);

  if(!funcSheet){
    throw new Error("Script_Function_Inventory not found");
  }

  const funcData = funcSheet.getDataRange().getValues();

  const validFunctions = new Set();

  for(let i=1;i<funcData.length;i++){

    const fn = funcData[i][0];

    if(fn) validFunctions.add(fn);

  }

  let sheet = metadataSS.getSheetByName(OUT_SHEET);
  if(!sheet) sheet = metadataSS.insertSheet(OUT_SHEET);

  sheet.clear();

  sheet.appendRow([
    "Caller_Function",
    "Called_Function",
    "File_Name",
    "Line_Number",
    "Detected_At"
  ]);

  const scriptId = ScriptApp.getScriptId();

  const response = UrlFetchApp.fetch(
    "https://script.googleapis.com/v1/projects/" + scriptId + "/content",
    {
      headers:{
        Authorization: "Bearer " + ScriptApp.getOAuthToken()
      }
    }
  );

  const project = JSON.parse(response.getContentText());

  const files = project.files || [];

  const rows = [];

  const now = new Date();

  const callRegex = /([A-Za-z0-9_]+)\s*\(/g;

  files.forEach(file => {

    if(file.type !== "SERVER_JS") return;

    const source = file.source || "";
    const fileName = file.name + ".gs";

    const funcRegex = /function\s+([A-Za-z0-9_]+)\s*\(/g;

    let funcMatch;

    while((funcMatch = funcRegex.exec(source)) !== null){

      const caller = funcMatch[1];

      const bodyStart = funcMatch.index;

      const bodyEnd = source.indexOf("}", bodyStart);

      const body = source.substring(bodyStart, bodyEnd);

      let callMatch;

      while((callMatch = callRegex.exec(body)) !== null){

        const called = callMatch[1];

        if(called === caller) continue;

        if(!validFunctions.has(called)) continue;

        const lineNumber = source
          .substring(0, callMatch.index)
          .split("\n").length;

        rows.push([
          caller,
          called,
          fileName,
          lineNumber,
          now
        ]);

      }

    }

  });

  if(rows.length){

    sheet.getRange(
      2,
      1,
      rows.length,
      rows[0].length
    ).setValues(rows);

  }

}