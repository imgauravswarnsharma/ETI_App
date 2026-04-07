function run_generateScriptArchitectureDiagram(){
  generateScriptArchitectureDiagram_();
}


function generateScriptArchitectureDiagram_(){

  const metadataSS = getMetadataSpreadsheet_();

  const ARCH_SHEET = "Script_Architecture_Logic";
  const OUT_SHEET  = "Script_Architecture_Diagram";

  const archSheet = metadataSS.getSheetByName(ARCH_SHEET);

  if(!archSheet){
    throw new Error("Script_Architecture_Logic sheet missing");
  }

  const data = archSheet.getDataRange().getValues();

  let sheet = metadataSS.getSheetByName(OUT_SHEET);

  if(!sheet) sheet = metadataSS.insertSheet(OUT_SHEET);

  sheet.clear();

  sheet.appendRow([
    "Diagram_Type",
    "Diagram_Definition",
    "Generated_At"
  ]);

  const edges = [];

  for(let i=1;i<data.length;i++){

    const parent = data[i][1];
    const node   = data[i][2];

    if(!parent || !node) continue;

    if(parent === "SYSTEM"){

      edges.push(`SYSTEM --> ${node}`);

    } else {

      edges.push(`${parent} --> ${node}`);

    }

  }

  const mermaid = [
    "graph TD",
    ...edges
  ].join("\n");

  const rows = [[
    "MERMAID",
    mermaid,
    new Date()
  ]];

  sheet.getRange(2,1,1,3).setValues(rows);

}