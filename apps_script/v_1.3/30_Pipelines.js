/* =========================
   Item Pipeline
   ========================= */
function pipeline_items_(){

  // -------------------------------------
  // EXECUTION CONTEXT
  // -------------------------------------
  // Rule:
  // - If context already exists (Controller triggered), DO NOT reinitialize
  // - Only enhance it (set pipeline_name + run_context)
  // - If no context (manual run), initialize fresh

  const SCRIPT_NAME = 'Pipeline';
  const FUNCTION_NAME = 'pipeline_items_';

  let ctx = getExecutionContext_();

  if (ctx) {
    // Existing context → enhance only
    ctx.pipeline_name = FUNCTION_NAME;
    ctx.run_context = "PIPELINE";
  } else {
    // No context → manual execution
    initExecutionContext_({
      pipeline_name: FUNCTION_NAME,
      run_context: "PIPELINE"
    });
  }

  const t0 = new Date();

  try {

    ETI_log_({
      scriptName: SCRIPT_NAME,
      functionName: FUNCTION_NAME,
      level: 'INFO',
      action: 'PIPELINE START',
      details: 'Item pipeline execution started'
    });

    // -------------------------------------
    // ACTUAL PIPELINE FUNCTIONS 
    // -------------------------------------
    populateStagingLookupItems_FromTransactionResolution();
    processStagingItems_StateMachine();
    promoteApprovedItems_FromStaging_ToLookup();
    backfill_ItemIDs_Machine_LookupItems();
    cleanupOrphan_ItemIDs_Machine_LookupItems();

    const durationMs = new Date().getTime() - t0.getTime();

    // -------------------------------------
    // LOGGING
    // -------------------------------------
    ETI_log_({
      scriptName: SCRIPT_NAME,
      functionName: FUNCTION_NAME,
      level: 'INFO',
      action: 'PIPELINE END',
      details: `Pipeline completed successfully | DurationMs=${durationMs}`
    });

  } catch (err) {

    // -------------------------------------
    // ERROR LOGGING
    // -------------------------------------
    ETI_logError_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      err,
      'PIPELINE'
    );

    throw err;

  } finally {

    // -------------------------------------
    // CRITICAL: Flush buffered logs once
    // -------------------------------------
    flushLogs_();

  }
}

/* =========================
   Brand Pipeline
   ========================= */
function pipeline_brands_(){

  initExecutionContext_({
    pipeline_name: "pipeline_brands_",
    run_context: "PIPELINE"
  });

  populateStagingLookupBrands_FromTransactionResolution();
  processStagingBrands_StateMachine();
  promoteApprovedBrands_FromStaging_ToLookup();
  backfill_BrandIDs_Machine_LookupBrands();
  cleanupOrphan_BrandIDs_Machine_LookupBrands();
}


/* =========================
   Product Pipeline
   ========================= */
function pipeline_products_(){

  initExecutionContext_({
    pipeline_name: "pipeline_products_",
    run_context: "PIPELINE"
  });

  populateStagingLookupProducts_FromTransactionResolution();
  processStagingProducts_StateMachine();
  promoteApprovedProducts_FromStaging_ToLookup();
  backfill_ProductIDs_Machine_LookupProducts();
  cleanupOrphan_ProductIDs_Machine_LookupProducts();
}


/* =========================
   Item-Brand Mapping Pipeline
   ========================= */
function pipeline_item_brand_mapping_(){

  initExecutionContext_({
    pipeline_name: "pipeline_item_brand_mapping_",
    run_context: "PIPELINE"
  });

  populateMapping_Item_Brand_FromTransactionResolution();
  processMapping_Item_Brand_StateMachine();
  cleanupMapping_Item_Brand_InvalidRows();
}


/* =========================
   Item-Brand-Product Mapping Pipeline
   ========================= */
function pipeline_item_brand_product_mapping_(){

  initExecutionContext_({
    pipeline_name: "pipeline_item_brand_product_mapping_",
    run_context: "PIPELINE"
  });

  populateMapping_Item_Brand_Product_FromTransactionResolution();
  processMapping_Item_Brand_Product_StateMachine();
  cleanupMapping_Item_Brand_Product_InvalidRows();
}


/* =========================
   Sheets Metadata Pipeline
   ========================= */
function sheets_metadata_pipeline_(){

  initExecutionContext_({
    pipeline_name: "sheets_metadata_pipeline_",
    run_context: "PIPELINE"
  });

  STEP3_exportSchemaSnapshot();
  exportFormulaInventory_v2_manifest();
  classifyColumns_fromManifest();
  generateDerivedColumnLogic();
  reconcile_access_control_metadata_();
}


/* =========================
   Scripts Metadata Pipeline
   ========================= */
function scripts_metadata_pipeline_(){

  initExecutionContext_({
    pipeline_name: "scripts_metadata_pipeline_",
    run_context: "PIPELINE"
  });

  extractScriptFunctionInventory_();
  generateScriptCallMap_RAW_();
  generateScriptCallMap_INTERNAL_();
  generateScriptPipelineMap_();
  generateScriptArchitectureLogic_();
  generateScriptArchitectureDiagram_();
  extractScriptFunctionCodeSummary_();
  extractScriptDataFlowMap_();
  extractScriptPerformanceMap_();
  extractScriptSheetInteractionMap_();
  generateAIContext_();
  exportAIContextMarkdown_();
}


/* =========================
   Full Metadata Pipeline
   ========================= */
function full_metadata_pipeline_(){

  initExecutionContext_({
    pipeline_name: "full_metadata_pipeline_",
    run_context: "PIPELINE"
  });

  sheets_metadata_pipeline_();
  scripts_metadata_pipeline_();
}


/* =========================
   Access Governance Pipeline
   ========================= */
function pipeline_access_mode_(){

  initExecutionContext_({
    pipeline_name: "pipeline_access_mode_",
    run_context: "PIPELINE"
  });

  reconcile_access_control_metadata_();
  apply_access_governance_();
}