function run_generateScriptCallMap_RAW(){
  generateScriptCallMap_RAW_();
}


function generateScriptCallMap_RAW_(){

  const metadataSS = getMetadataSpreadsheet_();
  const SHEET_NAME = "Script_Call_Map_RAW";

  let sheet = metadataSS.getSheetByName(SHEET_NAME);
  if(!sheet) sheet = metadataSS.insertSheet(SHEET_NAME);

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

  const ignore = new Set([
    "if",
    "for",
    "while",
    "switch",
    "catch",
    "function",
    "return",
    "Logger",
    "console"
  ]);

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

        if(ignore.has(called)) continue;

        if(called === caller) continue;

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