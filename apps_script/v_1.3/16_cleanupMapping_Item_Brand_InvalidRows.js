/**
 * Script Name: cleanupMapping_Item_Brand_InvalidRows
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Permanently DELETE invalid rows from Mapping_Item_Brand
 * - A row is invalid if:
 *   - Item_ID_Machine is missing OR
 *   - Brand_ID_Machine is missing
 * - Maintain a clean, authoritative mapping table
 *
 * Preconditions:
 * - Sheet must exist: Mapping_Item_Brand
 * - Header row present in row 1
 * - Required columns (header-based):
 *   - Item_ID_Machine
 *   - Brand_ID_Machine
 *
 * Algorithm (Step-by-Step):
 *
 * 1. Generate Execution_ID for traceability.
 * 2. Load Mapping_Item_Brand sheet.
 * 3. Resolve column indexes using header names.
 * 4. Scan all rows and identify invalid records:
 *      Item_ID_Machine missing
 *      OR
 *      Brand_ID_Machine missing
 * 5. Collect row numbers to delete.
 * 6. Delete rows in reverse order (bottom-up).
 * 7. Emit execution summary and completion logs.
 *
 * Failure Modes:
 * - Mapping_Item_Brand sheet not found
 * - Required column missing
 *
 * Reason for Deprecation:
 * - N/A
 */

function cleanupMapping_Item_Brand_InvalidRows() {

  const EXECUTION_ID = Utilities.getUuid();
  const SCRIPT_NAME  = 'cleanupMapping_Item_Brand_InvalidRows';
  const MAP_SHEET    = 'Mapping_Item_Brand';

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

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mapSh = ss.getSheetByName(MAP_SHEET);

  if (!mapSh) {
    throw new Error(`Sheet not found: ${MAP_SHEET}`);
  }

  const data = mapSh.getDataRange().getValues();

  if (data.length < 2) {

    ETI_log_({
      executionId: EXECUTION_ID,
      scriptName: SCRIPT_NAME,
      sheetName: MAP_SHEET,
      level: 'INFO',
      action: 'EXIT',
      details: 'No data rows present'
    });

    return;
  }

  const header = data[0];
  const col = n => header.indexOf(n);

  const IDX = {
    itemId: col('Item_ID_Machine'),
    brandId: col('Brand_ID_Machine')
  };

  for (const [key, idx] of Object.entries(IDX)) {

    if (idx === -1) {
      throw new Error(
        `Missing required column: ${
          key === 'itemId'
            ? 'Item_ID_Machine'
            : 'Brand_ID_Machine'
        }`
      );
    }
  }

  /* =========================
     IDENTIFY INVALID ROWS
     ========================= */

  const rowsToDelete = [];

  for (let i = 1; i < data.length; i++) {

    const rowNum = i + 1;

    const itemId  = data[i][IDX.itemId];
    const brandId = data[i][IDX.brandId];

    if (!itemId || !brandId) {
      rowsToDelete.push(rowNum);
    }
  }

  /* =========================
     DELETE ROWS (BOTTOM-UP)
     ========================= */

  for (let i = rowsToDelete.length - 1; i >= 0; i--) {

    mapSh.deleteRow(rowsToDelete[i]);

    ETI_log_({
      executionId: EXECUTION_ID,
      scriptName: SCRIPT_NAME,
      sheetName: MAP_SHEET,
      level: 'WARN',
      rowNumber: rowsToDelete[i],
      action: 'DELETE_INVALID_ROW',
      details: 'Missing Item_ID_Machine or Brand_ID_Machine'
    });
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