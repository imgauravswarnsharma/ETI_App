function run_extractScriptFunctionInventory(){
  extractScriptFunctionInventory_();
}


function extractScriptFunctionInventory_(){

  const metadataSS = getMetadataSpreadsheet_();

  const FILE_TABLE = "Script_File_Inventory";
  const FUNC_TABLE = "Script_Function_Inventory";

  let fileSheet = metadataSS.getSheetByName(FILE_TABLE);
  if(!fileSheet) fileSheet = metadataSS.insertSheet(FILE_TABLE);

  let funcSheet = metadataSS.getSheetByName(FUNC_TABLE);
  if(!funcSheet) funcSheet = metadataSS.insertSheet(FUNC_TABLE);

  fileSheet.clear();
  funcSheet.clear();

  fileSheet.appendRow([
    "File_Name",
    "File_Type",
    "Line_Count",
    "Function_Count",
    "Detected_At"
  ]);

  funcSheet.appendRow([
    "Function_Name",
    "File_Name",
    "Line_Number",
    "Function_Type",
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

  const fileRows = [];
  const funcRows = [];

  const now = new Date();

  files.forEach(file => {

    if(file.type !== "SERVER_JS") return;

    const fileName = file.name + ".gs";
    const source = file.source || "";

    const lines = source.split("\n");

    const lineCount = lines.length;

    const functionRegex = /function\s+([A-Za-z0-9_]+)\s*\(/g;

    let match;
    let functionCount = 0;

    while((match = functionRegex.exec(source)) !== null){

      const functionName = match[1];

      const lineNumber = source
        .substring(0,match.index)
        .split("\n").length;

      const type = classifyFunctionType_(functionName);

      funcRows.push([
        functionName,
        fileName,
        lineNumber,
        type,
        now
      ]);

      functionCount++;

    }

    fileRows.push([
      fileName,
      "SERVER_JS",
      lineCount,
      functionCount,
      now
    ]);

  });


  if(fileRows.length){

    fileSheet.getRange(
      2,
      1,
      fileRows.length,
      fileRows[0].length
    ).setValues(fileRows);

  }


  if(funcRows.length){

    funcSheet.getRange(
      2,
      1,
      funcRows.length,
      funcRows[0].length
    ).setValues(funcRows);

  }

}


function classifyFunctionType_(name){

  if(name.startsWith("pipeline_")) return "PIPELINE";

  if(name.includes("Controller")) return "CONTROLLER";

  if(name.includes("process")) return "PROCESSOR";

  if(name.includes("populate")) return "PROCESSOR";

  if(name.includes("promote")) return "PROCESSOR";

  if(name.includes("cleanup")) return "PROCESSOR";

  if(name.includes("metadata")) return "METADATA";

  if(name.includes("access")) return "ACCESS";

  if(name.includes("log")) return "LOGGER";

  return "UTILITY";

}



