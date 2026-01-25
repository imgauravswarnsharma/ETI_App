/**
 * Script Name: cleanupOrphan_ItemIDs_Machine_LookupItems
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Clear orphan Item_ID_Machine where Item_Name is missing
 *
 * Preconditions:
 * - Spreadsheet contains a sheet named: Lookup_Items
 * - Header row present in row 1
 * - Required columns:
 *   - Item_Name
 *   - Item_ID_Machine
 *
 * Algorithm:
 * 1. Generate Execution_ID
 * 2. Load Lookup_Items
 * 3. Resolve column indexes
 * 4. For each row:
 *    a. If Item_Name is blank AND Item_ID_Machine exists → clear ID
 * 5. Write changes in one batch
 *
 * Failure Modes:
 * - Sheet missing
 * - Required column missing
 *
 * Reason for Deprecation:
 * - N/A
 */

function cleanupOrphan_ItemIDs_Machine_LookupItems() {

  const EXECUTION_ID = Utilities.getUuid();
  const SCRIPT_NAME  = 'cleanupOrphan_ItemIDs_Machine_LookupItems';
  const SHEET_NAME   = 'Lookup_Items';

  ETI_log_({
    executionId: EXECUTION_ID,
    scriptName: SCRIPT_NAME,
    sheetName: SHEET_NAME,
    level: 'INFO',
    action: 'START',
    details: 'Execution started'
  });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`Sheet ${SHEET_NAME} not found`);

  const range = sh.getDataRange();
  const data  = range.getValues();

  if (data.length < 2) {
    ETI_log_({
      executionId: EXECUTION_ID,
      scriptName: SCRIPT_NAME,
      sheetName: SHEET_NAME,
      level: 'WARN',
      action: 'EXIT',
      details: 'No data rows found'
    });
    return;
  }

  const header = data[0];
  const col = n => header.indexOf(n);

  const IDX = {
    itemName: col('Item_Name'),
    itemIdM: col('Item_ID_Machine')
  };

  for (const [k, v] of Object.entries(IDX)) {
    if (v === -1) throw new Error(`Missing required column: ${k}`);
  }

  let clearedCount = 0;
  const output = data.map(r => r.slice());

  for (let i = 1; i < output.length; i++) {
    const rowNum = i + 1;
    const name   = output[i][IDX.itemName];
    const itemId = output[i][IDX.itemIdM];

    if (!name && itemId) {
      output[i][IDX.itemIdM] = '';
      clearedCount++;

      ETI_log_({
        executionId: EXECUTION_ID,
        scriptName: SCRIPT_NAME,
        sheetName: SHEET_NAME,
        level: 'WARN',
        rowNumber: rowNum,
        action: 'CLEAR_ORPHAN_ID',
        details: `Item_Name missing; Cleared Item_ID_Machine: ${itemId}`
      });
    }
  }

  range.setValues(output);

  ETI_log_({
    executionId: EXECUTION_ID,
    scriptName: SCRIPT_NAME,
    sheetName: SHEET_NAME,
    level: 'INFO',
    action: 'SUMMARY',
    details: `Cleared=${clearedCount}`
  });

  ETI_log_({
    executionId: EXECUTION_ID,
    scriptName: SCRIPT_NAME,
    sheetName: SHEET_NAME,
    level: 'INFO',
    action: 'END',
    details: 'Execution completed successfully'
  });
}
