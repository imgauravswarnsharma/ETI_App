/**
 * Script Name: populateStagingLookupBrands_FromTransactionResolution
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Populate Staging_Lookup_Brands with unresolved brand entries
 * - One staging row per canonical brand (forward-only intake)
 * - Acts as review + promotion intake for Brand master data
 * - Mirrors Item staging behavior exactly (ETI v1.3 consistency)
 *
 * Preconditions:
 * - Sheets must exist:
 *   - Transaction_Resolution
 *   - Staging_Lookup_Brands
 * - Header row must exist in row 1 for both sheets
 * - Required columns must exist (header-based, order-independent)
 * - Is_Lookup_Promoted is script-owned and must not be manually edited
 *
 * Algorithm (Step-by-Step):
 * 1. Generate Execution_ID for traceability.
 * 2. Load Staging_Lookup_Brands and build an in-memory set of
 *    existing Brand_Name_Canonical values.
 * 3. Load Transaction_Resolution and resolve required column indexes.
 * 4. Iterate each transaction row:
 *    a. Skip if Txn_ID_Machine is missing.
 *    b. Skip if Brand_ID_Machine already exists.
 *    c. Skip if Brand_Name_Canonical is missing.
 *    d. Skip if canonical already exists in staging.
 *    e. Otherwise, append a new staging row with a full UUID.
 * 5. Batch-write all new staging rows.
 * 6. Emit execution summary logs.
 *
 * Failure Modes:
 * - Missing required sheet
 * - Missing required column
 *
 * Reason for Deprecation (if applicable):
 * - N/A
 * - This script remains ACTIVE for ETI v1.3.
 * - Superseded only if brand ingestion moves to ID-first
 *   or reconciliation-based flow in v1.4.
 */

function populateStagingLookupBrands_FromTransactionResolution() {

  const EXECUTION_ID = Utilities.getUuid();
  const SCRIPT_NAME  = 'populateStagingLookupBrands_FromTransactionResolution';
  const TXN_SHEET    = 'Transaction_Resolution';
  const STG_SHEET    = 'Staging_Lookup_Brands';

  const t0 = new Date();
  console.log(`[${SCRIPT_NAME}] START`);

  ETI_log_({
    executionId: EXECUTION_ID,
    scriptName: SCRIPT_NAME,
    sheetName: STG_SHEET,
    level: 'INFO',
    action: 'START',
    details: 'Execution started'
  });

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const tsSh  = ss.getSheetByName(TXN_SHEET);
  const stgSh = ss.getSheetByName(STG_SHEET);

  if (!tsSh || !stgSh) {
    throw new Error('Required sheet not found');
  }

  /* ======================================================
     READ EXISTING STAGING CANONICALS
     ====================================================== */

  const stgData = stgSh.getDataRange().getValues();
  const stgHdr  = stgData[0];
  const stgCol  = n => stgHdr.indexOf(n);

  const IDX_STG_CANON = stgCol('Brand_Name_Canonical');
  if (IDX_STG_CANON === -1) {
    throw new Error('Staging_Lookup_Brands missing Brand_Name_Canonical');
  }

  const stagingCanonSet = new Set();
  for (let i = 1; i < stgData.length; i++) {
    const v = stgData[i][IDX_STG_CANON];
    if (v) stagingCanonSet.add(String(v));
  }

  /* ======================================================
     READ TRANSACTION RESOLUTION
     ====================================================== */

  const tsData = tsSh.getDataRange().getValues();
  const tsHdr  = tsData[0];
  const tsCol  = n => tsHdr.indexOf(n);

  const IDX = {
    txnId: tsCol('Txn_ID_Machine'),
    brandId: tsCol('Brand_ID_Machine'),
    brandEntered: tsCol('Brand_Name_Entered'),
    brandCanon: tsCol('Brand_Name_Canonical')
  };

  for (const [k, v] of Object.entries(IDX)) {
    if (v === -1) {
      throw new Error(`Transaction_Resolution missing column: ${k}`);
    }
  }

  /* ======================================================
     COUNTERS (LOGGING ONLY)
     ====================================================== */

  let scanned = 0;
  let skipNoTxn = 0;
  let skipHasBrand = 0;
  let skipNoCanon = 0;
  let skipDuplicateCanon = 0;

  const rowsToAppend = [];

  /* ======================================================
     STAGING POPULATION LOOP
     ====================================================== */

  for (let i = 1; i < tsData.length; i++) {
    scanned++;
    const r = tsData[i];

    if (!r[IDX.txnId]) {
      skipNoTxn++;
      continue;
    }

    if (r[IDX.brandId]) {
      skipHasBrand++;
      continue;
    }

    const canon = r[IDX.brandCanon];
    if (!canon) {
      skipNoCanon++;
      continue;
    }

    if (stagingCanonSet.has(canon)) {
      skipDuplicateCanon++;
      continue;
    }

    const newRow = new Array(stgHdr.length).fill('');

    newRow[stgCol('Source_Txn_ID_Machine')]    = r[IDX.txnId];
    newRow[stgCol('Staging_Brand_ID_Machine')] = Utilities.getUuid(); // full UUID
    newRow[stgCol('Mapped_Brand_ID_Machine')]  = '';
    newRow[stgCol('Brand_Name_Entered')]       = r[IDX.brandEntered];
    newRow[stgCol('Brand_Name_Canonical')]     = canon;
    newRow[stgCol('Brand_Name_Approved')]      = '';
    newRow[stgCol('Is_Approved')]              = false;
    newRow[stgCol('Is_Active')]                = true;
    newRow[stgCol('Is_Archived')]              = false;
    newRow[stgCol('Is_Lookup_Promoted')]       = false;
    newRow[stgCol('Review_Status')]            = 'Pending';
    newRow[stgCol('Notes')]                    = '';

    rowsToAppend.push(newRow);
    stagingCanonSet.add(canon);
  }

  /* ======================================================
     WRITE STAGING
     ====================================================== */

  if (rowsToAppend.length > 0) {
    stgSh.getRange(
      stgSh.getLastRow() + 1,
      1,
      rowsToAppend.length,
      rowsToAppend[0].length
    ).setValues(rowsToAppend);
  }

  const durationMs = new Date().getTime() - t0.getTime();

  console.log(
    `[${SCRIPT_NAME}] Scanned=${scanned}, ` +
    `SkipNoTxn=${skipNoTxn}, ` +
    `SkipHasBrand=${skipHasBrand}, ` +
    `SkipNoCanon=${skipNoCanon}, ` +
    `SkipDuplicateCanon=${skipDuplicateCanon}, ` +
    `Appended=${rowsToAppend.length}, ` +
    `DurationMs=${durationMs}`
  );

  ETI_log_({
    executionId: EXECUTION_ID,
    scriptName: SCRIPT_NAME,
    sheetName: STG_SHEET,
    level: 'INFO',
    action: 'SUMMARY',
    details:
      `Scanned=${scanned}, ` +
      `SkipNoTxn=${skipNoTxn}, ` +
      `SkipHasBrand=${skipHasBrand}, ` +
      `SkipNoCanon=${skipNoCanon}, ` +
      `SkipDuplicateCanon=${skipDuplicateCanon}, ` +
      `Appended=${rowsToAppend.length}, ` +
      `DurationMs=${durationMs}`
  });

  ETI_log_({
    executionId: EXECUTION_ID,
    scriptName: SCRIPT_NAME,
    sheetName: STG_SHEET,
    level: 'INFO',
    action: 'END',
    details: 'Execution completed successfully'
  });
}
