function run_generateScriptPipelineMap(){
  generateScriptPipelineMap_();
}


function generateScriptPipelineMap_(){

  const metadataSS = getMetadataSpreadsheet_();

  const FUNC_SHEET = "Script_Function_Inventory";
  const CALL_SHEET = "Script_Call_Map_INTERNAL";
  const OUT_SHEET  = "Script_Pipeline_Map";

  const funcSheet = metadataSS.getSheetByName(FUNC_SHEET);
  const callSheet = metadataSS.getSheetByName(CALL_SHEET);

  if(!funcSheet) throw new Error("Script_Function_Inventory missing");
  if(!callSheet) throw new Error("Script_Call_Map_INTERNAL missing");

  const funcData = funcSheet.getDataRange().getValues();
  const callData = callSheet.getDataRange().getValues();

  const funcTypeMap = {};

  for(let i=1;i<funcData.length;i++){

    const fn = funcData[i][0];
    const type = funcData[i][3];

    funcTypeMap[fn] = type;

  }

  let sheet = metadataSS.getSheetByName(OUT_SHEET);

  if(!sheet) sheet = metadataSS.insertSheet(OUT_SHEET);

  sheet.clear();

  sheet.appendRow([
    "Pipeline",
    "Step_Order",
    "Caller_Function",
    "Called_Function",
    "Called_Type",
    "File_Name",
    "Detected_At"
  ]);

  const rows = [];

  const now = new Date();

  const pipelineFunctions = Object.keys(funcTypeMap)
    .filter(fn => fn.startsWith("pipeline_") || fn.includes("pipeline"));

  pipelineFunctions.forEach(pipeline => {

    let step = 1;

    callData.forEach((row,i)=>{

      if(i===0) return;

      const caller = row[0];
      const called = row[1];
      const file = row[2];

      if(caller !== pipeline) return;

      const type = funcTypeMap[called] || "UNKNOWN";

      rows.push([
        pipeline,
        step,
        caller,
        called,
        type,
        file,
        now
      ]);

      step++;

    });

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