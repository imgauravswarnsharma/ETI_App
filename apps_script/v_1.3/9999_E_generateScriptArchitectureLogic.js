function run_generateScriptArchitectureLogic(){
  generateScriptArchitectureLogic_();
}


function generateScriptArchitectureLogic_(){

  const metadataSS = getMetadataSpreadsheet_();

  const FUNC_SHEET = "Script_Function_Inventory";
  const CALL_SHEET = "Script_Call_Map_INTERNAL";
  const PIPE_SHEET = "Script_Pipeline_Map";
  const OUT_SHEET  = "Script_Architecture_Logic";

  const funcSheet = metadataSS.getSheetByName(FUNC_SHEET);
  const callSheet = metadataSS.getSheetByName(CALL_SHEET);
  const pipeSheet = metadataSS.getSheetByName(PIPE_SHEET);

  if(!funcSheet) throw new Error("Script_Function_Inventory missing");
  if(!callSheet) throw new Error("Script_Call_Map_INTERNAL missing");
  if(!pipeSheet) throw new Error("Script_Pipeline_Map missing");

  const funcData = funcSheet.getDataRange().getValues();
  const callData = callSheet.getDataRange().getValues();
  const pipeData = pipeSheet.getDataRange().getValues();

  let sheet = metadataSS.getSheetByName(OUT_SHEET);

  if(!sheet) sheet = metadataSS.insertSheet(OUT_SHEET);

  sheet.clear();

  sheet.appendRow([
    "Architecture_Level",
    "Parent_Node",
    "Node_Name",
    "Node_Type",
    "Source_File",
    "Detected_At"
  ]);

  const rows = [];

  const now = new Date();

  /* --------------------------------------------------
     Build Function → Type Map
  -------------------------------------------------- */

  const funcTypeMap = {};
  const funcFileMap = {};

  for(let i=1;i<funcData.length;i++){

    const fn   = funcData[i][0];
    const file = funcData[i][1];
    const type = funcData[i][3];

    funcTypeMap[fn] = type;
    funcFileMap[fn] = file;

  }

  /* --------------------------------------------------
     Level 1 — Pipelines
  -------------------------------------------------- */

  const pipelines = Object.keys(funcTypeMap)
    .filter(fn => fn.startsWith("pipeline_") || fn.includes("pipeline"));

  pipelines.forEach(pipeline => {

    rows.push([
      1,
      "SYSTEM",
      pipeline,
      "PIPELINE",
      funcFileMap[pipeline] || "",
      now
    ]);

  });

  /* --------------------------------------------------
     Level 2 — Pipeline Steps
  -------------------------------------------------- */

  pipeData.forEach((row,i)=>{

    if(i===0) return;

    const pipeline = row[0];
    const called   = row[3];

    rows.push([
      2,
      pipeline,
      called,
      funcTypeMap[called] || "UNKNOWN",
      funcFileMap[called] || "",
      now
    ]);

  });

  /* --------------------------------------------------
     Level 3 — Utility Calls
  -------------------------------------------------- */

  callData.forEach((row,i)=>{

    if(i===0) return;

    const caller = row[0];
    const called = row[1];

    const type = funcTypeMap[called];

    if(!type) return;

    if(type === "UTILITY" || type === "LOGGER"){

      rows.push([
        3,
        caller,
        called,
        type,
        funcFileMap[called] || "",
        now
      ]);

    }

  });

  /* --------------------------------------------------
     Write Results
  -------------------------------------------------- */

  if(rows.length){

    sheet.getRange(
      2,
      1,
      rows.length,
      rows[0].length
    ).setValues(rows);

  }

}