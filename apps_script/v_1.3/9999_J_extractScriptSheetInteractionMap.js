function run_extractScriptSheetInteractionMap(){
  extractScriptSheetInteractionMap_();
}


function extractScriptSheetInteractionMap_(){

  const metadataSS = getMetadataSpreadsheet_();
  const SHEET_NAME = "Script_Sheet_Interaction_Map";

  let sheet = metadataSS.getSheetByName(SHEET_NAME);
  if(!sheet) sheet = metadataSS.insertSheet(SHEET_NAME);

  sheet.clear();

  sheet.appendRow([
    "Function_Name",
    "Sheet_Name",
    "Operation",
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

  const funcRegex = /function\s+([A-Za-z0-9_]+)\s*\(/g;

  files.forEach(file => {

    if(file.type !== "SERVER_JS") return;

    const source = file.source || "";
    const fileName = file.name + ".gs";

    let funcMatch;

    while((funcMatch = funcRegex.exec(source)) !== null){

      const fnName = funcMatch[1];

      const bodyStart = funcMatch.index;
      const bodyEnd = source.indexOf("}", bodyStart);

      const body = source.substring(bodyStart, bodyEnd);
      const lines = body.split("\n");

      const sheetVars = {};

      lines.forEach((line,i)=>{

        const trimmed = line.trim();

        const assignMatch = trimmed.match(/(const|let|var)\s+([A-Za-z0-9_]+)\s*=\s*.*getSheetByName\(["'](.+?)["']\)/);

        if(assignMatch){

          const variable = assignMatch[2];
          const sheetName = assignMatch[3];

          sheetVars[variable] = sheetName;
        }

        let operation = "";

        if(trimmed.includes("setValue(")) operation = "WRITE_CELL";
        if(trimmed.includes("setValues(")) operation = "WRITE_RANGE";
        if(trimmed.includes("appendRow(")) operation = "APPEND_ROW";
        if(trimmed.includes("clear(")) operation = "CLEAR";
        if(trimmed.includes("clearContents(")) operation = "CLEAR_CONTENT";
        if(trimmed.includes("getValues(")) operation = "READ_RANGE";
        if(trimmed.includes("getValue(")) operation = "READ_CELL";

        if(operation){

          let detectedSheet = "UNKNOWN";

          Object.keys(sheetVars).forEach(v=>{
            if(trimmed.includes(v + ".")){
              detectedSheet = sheetVars[v];
            }
          });

          const lineNumber =
            source.substring(0, funcMatch.index)
            .split("\n").length + i;

          rows.push([
            fnName,
            detectedSheet,
            operation,
            fileName,
            lineNumber,
            now
          ]);
        }

      });

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