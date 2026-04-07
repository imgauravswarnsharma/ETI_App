/*
-------------------------------------
Utility Function — Read Switch Table
(Column-based controller)
-------------------------------------
*/
function getAutomationSwitchMap_(){

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Automation_Control");

  if(!sheet) return {};

  const lastCol = sheet.getLastColumn();

  if(lastCol === 0) return {};

  const header = sheet.getRange(1,1,1,lastCol).getValues()[0];
  const values = sheet.getRange(2,1,1,lastCol).getValues()[0];

  const map = {};

  for(let i=0;i<header.length;i++){

    const name = header[i];
    let value = values[i];

    if(name === "" || name === null) continue;

    if(value === "TRUE") value = true;
    if(value === "FALSE") value = false;

    map[name] = value;

  }

  return map;

}



/*
-------------------------------------
Utility Function — Reset Switch
(Column-based controller)
-------------------------------------
*/
function resetSwitch_(switchName){

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Automation_Control");

  if(!sheet) return;

  const lastCol = sheet.getLastColumn();

  const header = sheet.getRange(1,1,1,lastCol).getValues()[0];

  const index = header.indexOf(switchName);

  if(index === -1) return;

  sheet.getRange(2,index+1).setValue(false);

}




/*
-------------------------------------
Controller
-------------------------------------
*/
function automationController_onChange(e){

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(3000)) {
    Logger.log("Automation controller already running. Skipping.");
    return;
  }

  const startTime = new Date();

  try {

    const switches = getAutomationSwitchMap_();

    const executionLogging = switches["Enable_Execution_Logging"] === true;
    const actionLogging = switches["Enable_Script_Action_Logging"] === true;

    /* ACCESS MODE SWITCH */
    const accessMode = switches["Access_Mode"] || "BACKEND";

    if(executionLogging){
      console.log("=== Automation Controller Started ===");
    }


    const runItemsPipeline = switches["Run_Item_Pipeline"] || false;
    const runBrandsPipeline = switches["Run_Brand_Pipeline"] || false;
    const runProductsPipeline = switches["Run_Product_Pipeline"] || false;

    const runItemBrandMapping = switches["Run_Item_Brand_Mapping_Pipeline"] || false;
    const runItemBrandProductMapping = switches["Run_Item_Brand_Product_Mapping_Pipeline"] || false;

    const runIntegrityCheck =
      switches["Enable_Staging_Integrity_Check"] !== false;

    const populateItems = switches["Populate_Items_Staging"] || false;
    const populateBrands = switches["Populate_Brands_Staging"] || false;
    const populateProducts = switches["Populate_Products_Staging"] || false;

    const promoteItems = switches["Promote_Items_To_Lookup"] || false;
    const promoteBrands = switches["Promote_Brands_To_Lookup"] || false;
    const promoteProducts = switches["Promote_Products_To_Lookup"] || false;

    const generateMetadata = switches["Run_Metadata_Pipeline"] || false;

    const runScriptsMetadataPipeline = switches["Run_Scripts_Metadata_Pipeline"] || false;

    const runAccessModePipeline = switches["Run_Access_Mode_Pipeline"] || false;

    /* PERSISTENT SWITCH */
    const exportAIContext =
      switches["Enable_Export_AI_Context_md"] === true;



    /* =========================
       Item Pipeline
    ========================= */

    if(runItemsPipeline){

      if(actionLogging) console.log("Running Item Pipeline");

      pipeline_items_();
      resetSwitch_("Run_Item_Pipeline");

    }



    /* =========================
       Brand Pipeline
    ========================= */

    if(runBrandsPipeline){

      if(actionLogging) console.log("Running Brand Pipeline");

      pipeline_brands_();
      resetSwitch_("Run_Brand_Pipeline");

    }



    /* =========================
       Product Pipeline
    ========================= */

    if(runProductsPipeline){

      if(actionLogging) console.log("Running Product Pipeline");

      pipeline_products_();
      resetSwitch_("Run_Product_Pipeline");

    }



    /* =========================
       Item ↔ Brand Mapping
    ========================= */

    if(runItemBrandMapping){

      if(actionLogging) console.log("Running Item-Brand Mapping Pipeline");

      pipeline_item_brand_mapping_();
      resetSwitch_("Run_Item_Brand_Mapping_Pipeline");

    }



    /* =========================
       Item ↔ Brand ↔ Product Mapping
    ========================= */

    if(runItemBrandProductMapping){

      if(actionLogging) console.log("Running Item-Brand-Product Mapping Pipeline");

      pipeline_item_brand_product_mapping_();
      resetSwitch_("Run_Item_Brand_Product_Mapping_Pipeline");

    }



    /* =========================
       Staging Integrity Check
    ========================= */

    if(runIntegrityCheck){

      if(actionLogging) console.log("Running Staging Integrity Checks");

      processStagingItems_StateMachine();
      processStagingBrands_StateMachine();
      processStagingProducts_StateMachine();

    }



    /* =========================
       Populate Staging
    ========================= */

    if(populateItems){

      if(actionLogging) console.log("Populate Items Staging");

      populateStagingLookupItems_FromTransactionResolution();
      resetSwitch_("Populate_Items_Staging");

    }

    if(populateBrands){

      if(actionLogging) console.log("Populate Brands Staging");

      populateStagingLookupBrands_FromTransactionResolution();
      resetSwitch_("Populate_Brands_Staging");

    }

    if(populateProducts){

      if(actionLogging) console.log("Populate Products Staging");

      populateStagingLookupProducts_FromTransactionResolution();
      resetSwitch_("Populate_Products_Staging");

    }



    /* =========================
       Promotion
    ========================= */

    if(promoteItems){

      if(actionLogging) console.log("Promote Items To Lookup");

      promoteApprovedItems_FromStaging_ToLookup();
      resetSwitch_("Promote_Items_To_Lookup");

    }

    if(promoteBrands){

      if(actionLogging) console.log("Promote Brands To Lookup");

      promoteApprovedBrands_FromStaging_ToLookup();
      resetSwitch_("Promote_Brands_To_Lookup");

    }

    if(promoteProducts){

      if(actionLogging) console.log("Promote Products To Lookup");

      promoteApprovedProducts_FromStaging_ToLookup();
      resetSwitch_("Promote_Products_To_Lookup");

    }



    /* =========================
       Metadata Pipeline
    ========================= */

    if(generateMetadata){

      if(actionLogging) console.log("Running Metadata Pipeline");

      metadata_pipeline_();
      resetSwitch_("Run_Metadata_Pipeline");

    }



    /* =========================
       Script Metadata Pipeline
    ========================= */

    if(runScriptsMetadataPipeline){

      if(actionLogging) console.log("Running Script Metadata Pipeline");

      regenerateScriptArchitecture_();
      resetSwitch_("Run_Scripts_Metadata_Pipeline");

    }



    /* =========================
       AI Context Export (Persistent)
    ========================= */

    if(exportAIContext){

      if(actionLogging) console.log("Exporting AI Context Markdown");

      run_generateAIContext();
      run_exportAIContextMarkdown();

    }



    /* =========================
       Access Control Pipeline
    ========================= */

    if(runAccessModePipeline){
    
      pipeline_access_mode_();
      
      resetSwitch_("Run_Access_Mode_Pipeline");
    }



/*
---------------------------------------------------------------------
---------------------------------------------------------------------
*/



    if(executionLogging){

      const endTime = new Date();
      const duration = (endTime - startTime) / 1000;

      console.log("=== Automation Controller Completed ===");
      console.log("Execution Time (seconds): " + duration);

    }

  }
  catch(error){

    console.error("Automation Controller Error:", error);

    throw error;

  }
  finally{

    lock.releaseLock();

  }

}