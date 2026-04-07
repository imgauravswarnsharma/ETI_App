function run_extractScriptPerformanceMap(){
  extractScriptPerformanceMap_();
}


function extractScriptPerformanceMap_(){

  const metadataSS = getMetadataSpreadsheet_();
  const OUT_SHEET = "Script_Performance_Map";

  let sheet = metadataSS.getSheetByName(OUT_SHEET);
  if(!sheet) sheet = metadataSS.insertSheet(OUT_SHEET);

  sheet.clear();

  sheet.appendRow([
    "Function_Name",
    "File_Name",
    "Performance_Flag",
    "Detected_Line",
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

      let insideLoop = false;

      lines.forEach((line,i)=>{

        const trimmed = line.trim();

        if(
          trimmed.startsWith("for(") ||
          trimmed.startsWith("for (") ||
          trimmed.startsWith("while(") ||
          trimmed.startsWith("while (")
        ){
          insideLoop = true;
        }

        let flag = "";

        if(trimmed.includes("getValue(") && insideLoop){
          flag = "READ_IN_LOOP";
        }

        if(trimmed.includes("setValue(") && insideLoop){
          flag = "WRITE_IN_LOOP";
        }

        if(trimmed.includes("appendRow(")){
          flag = "APPEND_ROW_USAGE";
        }

        if(trimmed.includes("getRange(") && insideLoop){
          flag = "RANGE_ACCESS_IN_LOOP";
        }

        if(flag){

          const lineNumber =
            source.substring(0, funcMatch.index)
            .split("\n").length + i;

          rows.push([
            fnName,
            fileName,
            flag,
            trimmed,
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