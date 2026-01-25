/**
 * Script Name: cleanupMapping_Item_Brand_Product_InvalidRows
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Permanently DELETE invalid rows from Mapping_Item_Brand_Product
 * - A row is invalid if ANY of the following is missing:
 *   - Item_ID_Machine
 *   - Brand_ID_Machine
 *   - Product_ID_Machine
 *
 * Preconditions:
 * - Sheet must exist: Mapping_Item_Brand_Product
 * - Header row present in row 1
 * - Required columns (header-based):
 *   - Item_ID_Machine
 *   - Brand_ID_Machine
 *   - Product_ID_Machine
 *
 * Algorithm (Step-by-Step):
 * 1. Generate Execution_ID for traceability.
 * 2. Read Mapping_Item_Brand_Product into memory.
 * 3. Resolve column indexes from header row.
 * 4. Identify rows where any of the three IDs is blank.
 * 5. Delete those rows (bottom-up to preserve row indexing).
 * 6. Emit execution summary and completion logs.
 *
 * Failure Modes:
 * - Mapping_Item_Brand_Product sheet not found
 * - Required columns missing
 *
 * Reason for Deprecation (if applicable):
 * - N/A
 */

function cleanupMapping_Item_Brand_Product_InvalidRows() {

  const EXECUTION_ID = Utilities.getUuid();
  const SCRIPT_NAME  = 'cleanupMapping_Item_Brand_Product_InvalidRows';
  const MAP_SHEET    = 'Mapping_Item_Brand_Product';

  const t0 = new Date();
  console.log(`[${SCRIPT_NAME}] START`);

  ETI_log_({
    executionId: EXECUTION_ID,
    scriptName: SCRIPT_NAME,
    sheetName: MAP_SHEET,
    level: 'INFO',
    action: 'START',
    details: 'Execution started'
  });

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const mapSh = ss.getSheetByName(MAP_SHEET);

  if (!mapSh) {
    throw new Error(`Sheet ${MAP_SHEET} not found`);
  }

  const dataRange = mapSh.getDataRange();
  const data      = dataRange.getValues();

  if (data.length < 2) {
    ETI_log_({
      executionId: EXECUTION_ID,
      scriptName: SCRIPT_NAME,
      sheetName: MAP_SHEET,
      level: 'INFO',
      action: 'EXIT',
      details: 'No data rows found'
    });
    return;
  }

  const header = data[0];
  const col = name => header.indexOf(name);

  const IDX = {
    itemId: col('Item_ID_Machine'),
    brandId: col('Brand_ID_Machine'),
    productId: col('Product_ID_Machine')
  };

  for (const [k, v] of Object.entries(IDX)) {
    if (v === -1) {
      throw new Error(`Missing required column: ${k}`);
    }
  }

  /* ======================================================
     IDENTIFY INVALID ROWS
     ====================================================== */

  const rowsToDelete = [];

  for (let i = 1; i < data.length; i++) {
    const rowNum = i + 1;
    const itemId    = data[i][IDX.itemId];
    const brandId   = data[i][IDX.brandId];
    const productId = data[i][IDX.productId];

    if (!itemId || !brandId || !productId) {
      rowsToDelete.push(rowNum);
    }
  }

  /* ======================================================
     DELETE ROWS (BOTTOM-UP)
     ====================================================== */

  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    mapSh.deleteRow(rowsToDelete[i]);
  }

  const durationMs = new Date().getTime() - t0.getTime();

  console.log(
    `[${SCRIPT_NAME}] Deleted=${rowsToDelete.length}, DurationMs=${durationMs}`
  );

  ETI_log_({
    executionId: EXECUTION_ID,
    scriptName: SCRIPT_NAME,
    sheetName: MAP_SHEET,
    level: 'INFO',
    action: 'SUMMARY',
    details: `Deleted=${rowsToDelete.length}, DurationMs=${durationMs}`
  });

  ETI_log_({
    executionId: EXECUTION_ID,
    scriptName: SCRIPT_NAME,
    sheetName: MAP_SHEET,
    level: 'INFO',
    action: 'END',
    details: 'Execution completed successfully'
  });
}
