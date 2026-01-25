/**
 * Script Name: backfill_ProductIDs_Machine_LookupProducts
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Backfill Product_ID_Machine where Product_Name exists and ID is missing
 *
 * Preconditions:
 * - Spreadsheet contains a sheet named: Lookup_Products
 * - Header row present in row 1
 * - Required columns:
 *   - Product_Name
 *   - Product_ID_Machine
 *
 * Algorithm:
 * 1. Generate Execution_ID
 * 2. Load Lookup_Products
 * 3. Resolve column indexes
 * 4. For each row:
 *    a. If Product_Name exists AND Product_ID_Machine is blank → generate UUID
 * 5. Write changes in one batch
 *
 * Failure Modes:
 * - Sheet missing
 * - Required column missing
 *
 * Reason for Deprecation:
 * - N/A
 */

function backfill_ProductIDs_Machine_LookupProducts() {

  const EXECUTION_ID = Utilities.getUuid();
  const SCRIPT_NAME  = 'backfill_ProductIDs_Machine_LookupProducts';
  const SHEET_NAME   = 'Lookup_Products';

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
    productName: col('Product_Name'),
    productIdM:  col('Product_ID_Machine')
  };

  for (const [k, v] of Object.entries(IDX)) {
    if (v === -1) throw new Error(`Missing required column: ${k}`);
  }

  let generatedCount = 0;
  const output = data.map(r => r.slice());

  for (let i = 1; i < output.length; i++) {
    const rowNum = i + 1;
    const name   = output[i][IDX.productName];
    const prodId = output[i][IDX.productIdM];

    if (name && !prodId) {
      const newId = Utilities.getUuid();
      output[i][IDX.productIdM] = newId;
      generatedCount++;

      ETI_log_({
        executionId: EXECUTION_ID,
        scriptName: SCRIPT_NAME,
        sheetName: SHEET_NAME,
        level: 'INFO',
        rowNumber: rowNum,
        action: 'GENERATE_ID',
        details: `Generated Product_ID_Machine=${newId}`
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
