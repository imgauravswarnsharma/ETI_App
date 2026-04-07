function run_extractScriptFunctionCodeSummary(){
  extractScriptFunctionCodeSummary_();
}


function extractScriptFunctionCodeSummary_(){

  const metadataSS = getMetadataSpreadsheet_();

  const OUT_SHEET = "Script_Function_Code_Summary";

  let sheet = metadataSS.getSheetByName(OUT_SHEET);

  if(!sheet) sheet = metadataSS.insertSheet(OUT_SHEET);

  sheet.clear();

  sheet.appendRow([
    "Function_Name",
    "File_Name",
    "Line_Number",
    "Parameter_Count",
    "Has_Return",
    "Approx_Line_Count",
    "First_Comment",
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

  const funcRegex = /function\s+([A-Za-z0-9_]+)\s*\(([^)]*)\)/g;

  files.forEach(file => {

    if(file.type !== "SERVER_JS") return;

    const source = file.source || "";
    const fileName = file.name + ".gs";

    let match;

    while((match = funcRegex.exec(source)) !== null){

      const fnName = match[1];
      const params = match[2];

      const paramCount = params.trim() === "" ? 0 : params.split(",").length;

      const startIndex = match.index;

      const bodyEnd = source.indexOf("}", startIndex);

      const body = source.substring(startIndex, bodyEnd);

      const hasReturn = body.includes("return");

      const approxLines = body.split("\n").length;

      const lineNumber = source
        .substring(0, startIndex)
        .split("\n").length;

      let comment = "";

      const commentRegex = /\/\*\*([\s\S]*?)\*\//;

      const commentMatch = commentRegex.exec(body);

      if(commentMatch){

        comment = commentMatch[1]
          .replace(/\n/g," ")
          .replace(/\*/g,"")
          .trim()
          .substring(0,200);

      }

      rows.push([
        fnName,
        fileName,
        lineNumber,
        paramCount,
        hasReturn,
        approxLines,
        comment,
        now
      ]);

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