function run_exportAIContextMarkdown(){
  exportAIContextMarkdown_();
}

function exportAIContextMarkdown_(){

  const metaSS = getMetadataSpreadsheet_();

  let md = "";

  md += "# ETI SYSTEM CONTEXT\n\n";

  md += "Architecture:\n";
  md += "- Google Sheets = Data Layer\n";
  md += "- Apps Script = Processing Pipelines\n";
  md += "- AppSheet = UI Layer\n\n";


  /* ========================
     TABLE ARCHITECTURE
  ======================== */

  const schema = metaSS.getSheetByName("Schema_Snapshot");

  if(schema){

    md += "## TABLE ARCHITECTURE\n";

    const data = schema.getDataRange().getValues();

    const tables = new Set();

    data.slice(1).forEach(r=>{
      if(r[0]) tables.add(r[0]);
    });

    tables.forEach(t=>{
      md += "- " + t + "\n";
    });

    md += "\n";
  }


  /* ========================
     DERIVED COLUMN LOGIC
  ======================== */

  const derived = metaSS.getSheetByName("Derived_Column_Logic");

  if(derived){

    md += "## DERIVED COLUMN LOGIC\n";

    const data = derived.getDataRange().getValues();

    data.slice(1).forEach(r=>{

      const table = r[0];
      const column = r[1];
      const desc = r[3];

      if(table && column){

        md += "- " + table + "." + column;

        if(desc) md += " : " + desc;

        md += "\n";

      }

    });

    md += "\n";
  }


  /* ========================
     SCRIPT FUNCTIONS
  ======================== */

  const funcs = metaSS.getSheetByName("Script_Function_Inventory");

  if(funcs){

    md += "## SCRIPT FUNCTIONS\n";

    const data = funcs.getDataRange().getValues();

    data.slice(1).forEach(r=>{
      md += "- " + r[0] + " (" + r[1] + ")\n";
    });

    md += "\n";
  }


  /* ========================
     SCRIPT DEPENDENCIES
  ======================== */

  const deps = metaSS.getSheetByName("Script_Call_Map_INTERNAL");

  if(deps){

    md += "## SCRIPT DEPENDENCIES\n";

    const data = deps.getDataRange().getValues();

    data.slice(1).forEach(r=>{
      md += "- " + r[0] + " → " + r[1] + "\n";
    });

    md += "\n";
  }


  /* ========================
     SCRIPT DATA FLOW
  ======================== */

  const flow = metaSS.getSheetByName("Script_DataFlow_Map");

  if(flow){

    md += "## SCRIPT DATA FLOW\n";

    const data = flow.getDataRange().getValues();

    data.slice(1).forEach(r=>{
      md += "- " + r[0] + " : " + r[2] + " → " + r[3] + "\n";
    });

    md += "\n";
  }
/* ========================
   WRITE FILE TO DRIVE
======================== */

const fileName = "ETI_SYSTEM_CONTEXT.md";

const files = DriveApp.getFilesByName(fileName);

if(files.hasNext()){

  const file = files.next();
  file.setContent(md);

  Logger.log("Markdown updated: " + file.getUrl());

}else{

  const file = DriveApp.createFile(
    fileName,
    md,
    MimeType.PLAIN_TEXT
  );

  Logger.log("Markdown created: " + file.getUrl());

}
}