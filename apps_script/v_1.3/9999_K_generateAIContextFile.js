function run_generateAIContext(){
  generateAIContext_();
}

function generateAIContext_(){

  const metadataSS = getMetadataSpreadsheet_();

  const OUT_SHEET = "AI_CONTEXT_EXPORT";

  let sheet = metadataSS.getSheetByName(OUT_SHEET);

  if(!sheet) sheet = metadataSS.insertSheet(OUT_SHEET);

  sheet.clear();

  let output = [];

  output.push("# ETI SYSTEM CONTEXT");
  output.push("");

  output.push("Architecture:");
  output.push("Google Sheets = Data Layer");
  output.push("Apps Script = Processing Pipelines");
  output.push("AppSheet = UI Layer");
  output.push("");


  /* =========================
     TABLE ARCHITECTURE
  ========================= */

  const schema = metadataSS.getSheetByName("Schema_Snapshot");

  if(schema){

    const data = schema.getDataRange().getValues();

    output.push("## TABLE ARCHITECTURE");

    const tables = new Set();

    data.slice(1).forEach(r=>{
      if(r[0]) tables.add(r[0]);
    });

    tables.forEach(t=>{
      output.push("- " + t);
    });

    output.push("");
  }



  /* =========================
     DERIVED COLUMN LOGIC
  ========================= */

  const derived = metadataSS.getSheetByName("Derived_Column_Logic");

  if(derived){

    const data = derived.getDataRange().getValues();

    output.push("## DERIVED COLUMN LOGIC");

    data.slice(1).forEach(r=>{

      const table = r[0];
      const column = r[1];
      const purpose = r[3];

      if(table && column){

        let line = "- " + table + "." + column;

        if(purpose) line += " : " + purpose;

        output.push(line);

      }

    });

    output.push("");
  }



  /* =========================
     SCRIPT FUNCTIONS
  ========================= */

  const functions = metadataSS.getSheetByName("Script_Function_Inventory");

  if(functions){

    const data = functions.getDataRange().getValues();

    output.push("## SCRIPT FUNCTIONS");

    data.slice(1).forEach(r=>{
      output.push("- " + r[0] + " (" + r[1] + ")");
    });

    output.push("");
  }



  /* =========================
     SCRIPT DEPENDENCIES
  ========================= */

  const callMap = metadataSS.getSheetByName("Script_Call_Map_INTERNAL");

  if(callMap){

    const data = callMap.getDataRange().getValues();

    output.push("## SCRIPT DEPENDENCIES");

    data.slice(1).forEach(r=>{
      output.push("- " + r[0] + " → " + r[1]);
    });

    output.push("");
  }



  /* =========================
     SCRIPT DATA FLOW
  ========================= */

  const flow = metadataSS.getSheetByName("Script_DataFlow_Map");

  if(flow){

    const data = flow.getDataRange().getValues();

    output.push("## SCRIPT DATA FLOW");

    data.slice(1).forEach(r=>{
      output.push("- " + r[0] + " : " + r[2] + " → " + r[3]);
    });

    output.push("");
  }



  /* ====================================================
     WRITE NORMAL READABLE CONTEXT (COLUMN A)
  ==================================================== */

  sheet.getRange(1,1,output.length,1)
       .setValues(output.map(x=>[x]));



  /* ====================================================
     CREATE MARKDOWN BLOCK
  ==================================================== */

  const md = output.join("\n");



  /* ====================================================
     SPLIT INTO CHUNKS (avoid 50k limit)
  ==================================================== */

  const CHUNK_SIZE = 40000;

  let chunks = [];

  for(let i=0;i<md.length;i+=CHUNK_SIZE){

    chunks.push(md.substring(i,i+CHUNK_SIZE));

  }



  /* ====================================================
     WRITE CHUNKS STARTING COLUMN C
  ==================================================== */

  sheet.getRange(1,3).setValue("COPY_PASTE_CONTEXT");

  sheet.getRange(2,3,1,chunks.length)
       .setValues([chunks]);

}