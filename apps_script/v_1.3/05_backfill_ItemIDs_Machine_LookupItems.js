/**
 * Script Name: backfill_ItemIDs_Machine_LookupItems
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Backfill Item_ID_Machine where Item_Name exists and ID is missing
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
 *    a. If Item_Name exists AND Item_ID_Machine is blank → generate UUID
 * 5. Write changes in one batch
 *
 * Failure Modes:
 * - Sheet missing
 * - Required column missing
 *
 * Reason for Deprecation:
 * - N/A
 */

function backfill_ItemIDs_Machine_LookupItems() {

  const EXECUTION_ID = Utilities.getUuid();
  const SCRIPT_NAME  = 'backfill_ItemIDs_Machine_LookupItems';
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

  let generatedCount = 0;
  const output = data.map(r => r.slice());

  for (let i = 1; i < output.length; i++) {
    const rowNum  = i + 1;
    const name    = output[i][IDX.itemName];
    const itemId  = output[i][IDX.itemIdM];

    if (name && !itemId) {
      const newId = Utilities.getUuid();
      output[i][IDX.itemIdM] = newId;
      generatedCount++;

      ETI_log_({
        executionId: EXECUTION_ID,
        scriptName: SCRIPT_NAME,
        sheetName: SHEET_NAME,
        level: 'INFO',
        rowNumber: rowNum,
        action: 'GENERATE_ID',
        details: `Generated Item_ID_Machine: ${newId}`
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
    details: `Generated=${generatedCount}`
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
