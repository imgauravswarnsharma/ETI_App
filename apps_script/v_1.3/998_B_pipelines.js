/* =========================
   Item Pipeline
   ========================= */
function pipeline_items_(){

  populateStagingLookupItems_FromTransactionResolution();

  processStagingItems_StateMachine();

  promoteApprovedItems_FromStaging_ToLookup();

  backfill_ItemIDs_Machine_LookupItems();

  cleanupOrphan_ItemIDs_Machine_LookupItems();

}




/* =========================
   Brand Pipeline
   =====================+=== */
function pipeline_brands_(){

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

  populateStagingLookupProducts_FromTransactionResolution();

  processStagingProducts_StateMachine()

  promoteApprovedProducts_FromStaging_ToLookup();

  backfill_ProductIDs_Machine_LookupProducts();

  cleanupOrphan_ProductIDs_Machine_LookupProducts();

}




/* =========================
   Item-Brand Mapping Pipeline
   ========================= */

function pipeline_item_brand_mapping_(){

  populateMapping_Item_Brand_FromTransactionResolution();

  processMapping_Item_Brand_StateMachine();
  
  cleanupMapping_Item_Brand_InvalidRows();
}




/* =========================
   Item-Brand-Product Mapping Pipeline
   ========================= */
function pipeline_item_brand_product_mapping_(){

  populateMapping_Item_Brand_Product_FromTransactionResolution();

  processMapping_Item_Brand_Product_StateMachine();

  cleanupMapping_Item_Brand_Product_InvalidRows();
}




/* =========================
   Metadata Pipeline
   ========================= */
function metadata_pipeline_(){

  STEP3_exportSchemaSnapshot();

  exportFormulaInventory_v2_manifest();
  
  classifyColumns_fromManifest();

  generateDerivedColumnLogic();

  reconcile_access_metadata_();

}

/* =========================
   Scripts Metadata Pipeline
   ========================= */
function regenerateScriptArchitecture_(){

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
   Access Governance Pipeline
   ========================= */
function pipeline_access_mode_(){

  reconcile_access_control_metadata_();

  apply_access_governance_();

}


