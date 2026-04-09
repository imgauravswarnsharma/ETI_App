/*
-------------------------------------
EXECUTION CONTEXT
-------------------------------------

Rule:

If context already exists (Controller triggered), DO NOT reinitialize.

Only enhance it (set pipeline_name + run_context)

If no context (manual run), initialize fresh
*/





/* =========================
   Item Pipeline
   ========================= */
function pipeline_items_(){

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

/*-------------------------------------
  ACTUAL PIPELINE FUNCTIONS 
-------------------------------------*/
    populateStagingLookupItems_FromTransactionResolution();
    processStagingItems_StateMachine();
    promoteApprovedItems_FromStaging_ToLookup();
    backfill_ItemIDs_Machine_LookupItems();
    cleanupOrphan_ItemIDs_Machine_LookupItems();

    const durationMs = new Date().getTime() - t0.getTime();

/*-------------------------------------
  LOGGING
-------------------------------------*/
    ETI_log_({
      scriptName: SCRIPT_NAME,
      functionName: FUNCTION_NAME,
      level: 'INFO',
      action: 'PIPELINE END',
      details: `Pipeline completed successfully | DurationMs=${durationMs}`
    });

  } catch (err) {

/*-------------------------------------
  ERROR LOGGING
-------------------------------------*/
    ETI_logError_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      err,
      'PIPELINE'
    );

    throw err;

  } finally {

    flushLogs_(); // CRITICAL: Flush buffered logs once
  
  }
}

/* =========================
   Brand Pipeline
   ========================= */
function pipeline_brands_(){

  // -------------------------------------
  // EXECUTION CONTEXT
  // -------------------------------------

  const SCRIPT_NAME = 'Pipeline';
  const FUNCTION_NAME = 'pipeline_brands_';

  let ctx = getExecutionContext_();

  if (ctx) {
    ctx.pipeline_name = FUNCTION_NAME;
    ctx.run_context = "PIPELINE";
  } else {
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
      details: 'Brand pipeline execution started'
    });

    // -------------------------------------
    // ACTUAL PIPELINE FUNCTIONS
    // -------------------------------------
    populateStagingLookupBrands_FromTransactionResolution();
    processStagingBrands_StateMachine();
    promoteApprovedBrands_FromStaging_ToLookup();
    backfill_BrandIDs_Machine_LookupBrands();
    cleanupOrphan_BrandIDs_Machine_LookupBrands();

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
   Product Pipeline
   ========================= */
function pipeline_products_(){

  /* -------------------------------------
     EXECUTION CONTEXT
  ------------------------------------- */

  const SCRIPT_NAME = 'Pipeline';
  const FUNCTION_NAME = 'pipeline_products_';

  let ctx = getExecutionContext_();

  if (ctx) {
    /* Existing context → enhance only */
    ctx.pipeline_name = FUNCTION_NAME;
    ctx.run_context = "PIPELINE";
  } else {
    /* No context → manual execution */
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
      details: 'Product pipeline execution started'
    });

    /* -------------------------------------
       ACTUAL PIPELINE FUNCTIONS
    ------------------------------------- */

    populateStagingLookupProducts_FromTransactionResolution();
    processStagingProducts_StateMachine();
    promoteApprovedProducts_FromStaging_ToLookup();
    backfill_ProductIDs_Machine_LookupProducts();
    cleanupOrphan_ProductIDs_Machine_LookupProducts();

    const durationMs = new Date().getTime() - t0.getTime();

    /* -------------------------------------
       LOGGING
    ------------------------------------- */

    ETI_log_({
      scriptName: SCRIPT_NAME,
      functionName: FUNCTION_NAME,
      level: 'INFO',
      action: 'PIPELINE END',
      details: `Pipeline completed successfully | DurationMs=${durationMs}`
    });

  } catch (err) {

    /* -------------------------------------
       ERROR LOGGING
    ------------------------------------- */

    ETI_logError_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      err,
      'PIPELINE'
    );

    throw err;

  } finally {

    /* -------------------------------------
       CRITICAL: Flush buffered logs once
    ------------------------------------- */

    flushLogs_();

  }
}


/* =========================
   Item-Brand Mapping Pipeline
   ========================= */
function pipeline_item_brand_mapping_(){

  /* -------------------------------------
     EXECUTION CONTEXT
  ------------------------------------- */

  const SCRIPT_NAME = 'Pipeline';
  const FUNCTION_NAME = 'pipeline_item_brand_mapping_';

  let ctx = getExecutionContext_();

  if (ctx) {
    /* Existing context → enhance only */
    ctx.pipeline_name = FUNCTION_NAME;
    ctx.run_context = "PIPELINE";
  } else {
    /* No context → manual execution */
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
      details: 'Item-Brand mapping pipeline execution started'
    });

    /* -------------------------------------
       ACTUAL PIPELINE FUNCTIONS
    ------------------------------------- */

    populateMapping_Item_Brand_FromTransactionResolution();
    processMapping_Item_Brand_StateMachine();
    cleanupMapping_Item_Brand_InvalidRows();

    const durationMs = new Date().getTime() - t0.getTime();

    /* -------------------------------------
       LOGGING
    ------------------------------------- */

    ETI_log_({
      scriptName: SCRIPT_NAME,
      functionName: FUNCTION_NAME,
      level: 'INFO',
      action: 'PIPELINE END',
      details: `Pipeline completed successfully | DurationMs=${durationMs}`
    });

  } catch (err) {

    /* -------------------------------------
       ERROR LOGGING
    ------------------------------------- */

    ETI_logError_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      err,
      'PIPELINE'
    );

    throw err;

  } finally {

    /* -------------------------------------
       CRITICAL: Flush buffered logs once
    ------------------------------------- */

    flushLogs_();

  }
}


/* =========================
   Item-Brand-Product Mapping Pipeline
   ========================= */
function pipeline_item_brand_product_mapping_(){

  /* -------------------------------------
     EXECUTION CONTEXT
  ------------------------------------- */

  const SCRIPT_NAME = 'Pipeline';
  const FUNCTION_NAME = 'pipeline_item_brand_product_mapping_';

  let ctx = getExecutionContext_();

  if (ctx) {
    /* Existing context → enhance only */
    ctx.pipeline_name = FUNCTION_NAME;
    ctx.run_context = "PIPELINE";
  } else {
    /* No context → manual execution */
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
      details: 'Item-Brand-Product mapping pipeline execution started'
    });

    /* -------------------------------------
       ACTUAL PIPELINE FUNCTIONS
    ------------------------------------- */

    populateMapping_Item_Brand_Product_FromTransactionResolution();
    processMapping_Item_Brand_Product_StateMachine();
    cleanupMapping_Item_Brand_Product_InvalidRows();

    const durationMs = new Date().getTime() - t0.getTime();

    /* -------------------------------------
       LOGGING
    ------------------------------------- */

    ETI_log_({
      scriptName: SCRIPT_NAME,
      functionName: FUNCTION_NAME,
      level: 'INFO',
      action: 'PIPELINE END',
      details: `Pipeline completed successfully | DurationMs=${durationMs}`
    });

  } catch (err) {

    /* -------------------------------------
       ERROR LOGGING
    ------------------------------------- */

    ETI_logError_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      err,
      'PIPELINE'
    );

    throw err;

  } finally {

    /* -------------------------------------
       CRITICAL: Flush buffered logs once
    ------------------------------------- */

    flushLogs_();

  }
}

/* =========================
   Sheets Metadata Pipeline
   ========================= */
function sheets_metadata_pipeline_(){

  /* -------------------------------------
     EXECUTION CONTEXT
  ------------------------------------- */

  const SCRIPT_NAME = 'Pipeline';
  const FUNCTION_NAME = 'sheets_metadata_pipeline_';

  let ctx = getExecutionContext_();

  if (ctx) {
    /* Existing context → enhance only */
    ctx.pipeline_name = FUNCTION_NAME;
    ctx.run_context = "PIPELINE";
  } else {
    /* No context → manual execution */
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
      details: 'Sheets metadata pipeline execution started'
    });

    /* -------------------------------------
       ACTUAL PIPELINE FUNCTIONS
    ------------------------------------- */

    STEP3_exportSchemaSnapshot();
    exportFormulaInventory_v2_manifest();
    classifyColumns_fromManifest();
    generateDerivedColumnLogic();
    //reconcile_access_control_metadata_();

    const durationMs = new Date().getTime() - t0.getTime();

    /* -------------------------------------
       LOGGING
    ------------------------------------- */

    ETI_log_({
      scriptName: SCRIPT_NAME,
      functionName: FUNCTION_NAME,
      level: 'INFO',
      action: 'PIPELINE END',
      details: `Pipeline completed successfully | DurationMs=${durationMs}`
    });

  } catch (err) {

    /* -------------------------------------
       ERROR LOGGING
    ------------------------------------- */

    ETI_logError_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      err,
      'PIPELINE'
    );

    throw err;

  } finally {

    /* -------------------------------------
       CRITICAL: Flush buffered logs once
    ------------------------------------- */

    flushLogs_();

  }
}


/* =========================
   Scripts Metadata Pipeline
   ========================= */
function scripts_metadata_pipeline_(){

  /* -------------------------------------
     EXECUTION CONTEXT
  ------------------------------------- */

  const SCRIPT_NAME = 'Pipeline';
  const FUNCTION_NAME = 'scripts_metadata_pipeline_';

  let ctx = getExecutionContext_();

  if (ctx) {
    /* Existing context → enhance only */
    ctx.pipeline_name = FUNCTION_NAME;
    ctx.run_context = "PIPELINE";
  } else {
    /* No context → manual execution */
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
      details: 'Scripts metadata pipeline execution started'
    });

    /* -------------------------------------
       ACTUAL PIPELINE FUNCTIONS
    ------------------------------------- */

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

    const durationMs = new Date().getTime() - t0.getTime();

    /* -------------------------------------
       LOGGING
    ------------------------------------- */

    ETI_log_({
      scriptName: SCRIPT_NAME,
      functionName: FUNCTION_NAME,
      level: 'INFO',
      action: 'PIPELINE END',
      details: `Pipeline completed successfully | DurationMs=${durationMs}`
    });

  } catch (err) {

    /* -------------------------------------
       ERROR LOGGING
    ------------------------------------- */

    ETI_logError_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      err,
      'PIPELINE'
    );

    throw err;

  } finally {

    /* -------------------------------------
       CRITICAL: Flush buffered logs once
    ------------------------------------- */

    flushLogs_();

  }
}


/* =========================
   Full Metadata Pipeline
   ========================= */
function full_metadata_pipeline_(){

  /* -------------------------------------
     EXECUTION CONTEXT
  ------------------------------------- */

  const SCRIPT_NAME = 'Pipeline';
  const FUNCTION_NAME = 'full_metadata_pipeline_';

  let ctx = getExecutionContext_();

  if (ctx) {
    /* Existing context → enhance only */
    ctx.pipeline_name = FUNCTION_NAME;
    ctx.run_context = "PIPELINE";
  } else {
    /* No context → manual execution */
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
      details: 'Full metadata pipeline execution started'
    });

    /* -------------------------------------
       ACTUAL PIPELINE FUNCTIONS
    ------------------------------------- */

    sheets_metadata_pipeline_();
    scripts_metadata_pipeline_();

    const durationMs = new Date().getTime() - t0.getTime();

    /* -------------------------------------
       LOGGING
    ------------------------------------- */

    ETI_log_({
      scriptName: SCRIPT_NAME,
      functionName: FUNCTION_NAME,
      level: 'INFO',
      action: 'PIPELINE END',
      details: `Pipeline completed successfully | DurationMs=${durationMs}`
    });

  } catch (err) {

    /* -------------------------------------
       ERROR LOGGING
    ------------------------------------- */

    ETI_logError_(
      SCRIPT_NAME,
      FUNCTION_NAME,
      err,
      'PIPELINE'
    );

    throw err;

  } finally {

    /* -------------------------------------
       CRITICAL: Flush buffered logs once
    ------------------------------------- */

    flushLogs_();

  }
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