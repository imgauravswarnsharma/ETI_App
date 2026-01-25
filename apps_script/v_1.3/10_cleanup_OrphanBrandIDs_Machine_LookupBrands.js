/**
 * Script Name: cleanupOrphan_BrandIDs_Machine_LookupBrands
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Clear orphan Brand_ID_Machine where Brand_Name is missing
 *
 * Preconditions:
 * - Spreadsheet contains a sheet named: Lookup_Brands
 * - Header row present in row 1
 * - Required columns:
 *   - Brand_Name
 *   - Brand_ID_Machine
 *
 * Algorithm:
 * 1. Generate Execution_ID
 * 2. Load Lookup_Brands
 * 3. Resolve column indexes
 * 4. For each row:
 *    a. If Brand_Name is blank AND Brand_ID_Machine exists → clear ID
 * 5. Write changes in one batch
 *
 * Failure Modes:
 * - Sheet missing
 * - Required column missing
 *
 * Reason for Deprecation:
 * - N/A
 */

function cleanupOrphan_BrandIDs_Machine_LookupBrands() {

  const EXECUTION_ID = Utilities.getUuid();
  const SCRIPT_NAME  = 'cleanupOrphan_BrandIDs_Machine_LookupBrands';
  const SHEET_NAME   = 'Lookup_Brands';

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
    brandName: col('Brand_Name'),
    brandIdM:  col('Brand_ID_Machine')
  };

  for (const [k, v] of Object.entries(IDX)) {
    if (v === -1) throw new Error(`Missing required column: ${k}`);
  }

  let clearedCount = 0;
  const output = data.map(r => r.slice());

  for (let i = 1; i < output.length; i++) {
    const rowNum  = i + 1;
    const name    = output[i][IDX.brandName];
    const brandId = output[i][IDX.brandIdM];

    if (!name && brandId) {
      output[i][IDX.brandIdM] = '';
      clearedCount++;

      ETI_log_({
        executionId: EXECUTION_ID,
        scriptName: SCRIPT_NAME,
        sheetName: SHEET_NAME,
        level: 'WARN',
        rowNumber: rowNum,
        action: 'CLEAR_ORPHAN_ID',
        details: `Brand_Name missing; Cleared Brand_ID_Machine: ${brandId}`
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
