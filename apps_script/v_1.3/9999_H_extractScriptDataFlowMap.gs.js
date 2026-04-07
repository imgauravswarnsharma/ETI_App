function run_extractScriptDataFlowMap(){
  extractScriptDataFlowMap_();
}


function extractScriptDataFlowMap_(){

  const metadataSS = getMetadataSpreadsheet_();

  const OUT_SHEET = "Script_DataFlow_Map";

  let sheet = metadataSS.getSheetByName(OUT_SHEET);

  if(!sheet) sheet = metadataSS.insertSheet(OUT_SHEET);

  sheet.clear();

  sheet.appendRow([
    "Function_Name",
    "File_Name",
    "Operation_Type",
    "Sheet_Reference",
    "Detected_Line",
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

      lines.forEach((line,i)=>{

        let operation = "";

        if(line.includes("getSheetByName")) operation = "SHEET_ACCESS";
        if(line.includes("getRange")) operation = "RANGE_ACCESS";
        if(line.includes("getValues")) operation = "READ_VALUES";
        if(line.includes("getValue")) operation = "READ_VALUE";
        if(line.includes("setValues")) operation = "WRITE_VALUES";
        if(line.includes("setValue")) operation = "WRITE_VALUE";
        if(line.includes("appendRow")) operation = "APPEND_ROW";
        if(line.includes("clear")) operation = "CLEAR_RANGE";

        if(operation){

          const lineNumber =
            source.substring(0, funcMatch.index)
            .split("\n").length + i;

          rows.push([
            fnName,
            fileName,
            operation,
            line.trim(),
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