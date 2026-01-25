/**
 * Script Name: populateMapping_Item_Brand_FromTransactionStaging
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Populate Mapping_Item_Brand with unique Item–Brand relationships
 *   discovered from Transaction_Staging
 * - One row per (Item_ID_Machine, Brand_ID_Machine)
 * - Transaction acts strictly as discovery context (first-seen evidence)
 *
 * Preconditions:
 * - Sheets must exist:
 *   - Transaction_Staging
 *   - Mapping_Item_Brand
 * - Header row must exist in row 1
 * - Required columns must exist (header-based, order-independent)
 *
 * Algorithm (Step-by-Step):
 * 1. Generate Execution_ID for traceability.
 * 2. Read Mapping_Item_Brand and build an in-memory set of existing
 *    (Item_ID_Machine, Brand_ID_Machine) pairs.
 * 3. Read Transaction_Staging and iterate rows in order:
 *    a. Skip rows without Txn_ID_Machine.
 *    b. Skip rows without Item_ID_Machine or Brand_ID_Machine.
 *    c. Capture the first-seen Item–Brand pair only once.
 * 4. Append new mapping rows for unseen pairs with:
 *    - First_Seen_Txn_ID
 *    - First_Seen_Txn_Date
 *    - Canonical snapshots (descriptive only)
 * 5. Batch-write all new rows.
 * 6. Emit execution summary and completion logs.
 *
 * Failure Modes:
 * - Required sheet missing
 * - Required column missing
 *
 * Reason for Deprecation (if applicable):
 * - N/A
 * - Script remains ACTIVE for ETI v1.3.
 * - Superseded only if mapping discovery becomes event-driven in v1.4.
 */

function populateMapping_Item_Brand_FromTransactionStaging() {

  const EXECUTION_ID = Utilities.getUuid();
  const SCRIPT_NAME  = 'populateMapping_Item_Brand_FromTransactionStaging';
  const TXN_SHEET    = 'Transaction_Staging';
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

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const tsSh  = ss.getSheetByName(TXN_SHEET);
  const mapSh = ss.getSheetByName(MAP_SHEET);

  if (!tsSh || !mapSh) {
    throw new Error('Required sheet not found');
  }

  /* ======================================================
     READ EXISTING MAPPINGS (DEDUP SET)
     ====================================================== */

  const mapData = mapSh.getDataRange().getValues();
  const mapHdr  = mapData[0];
  const mapCol  = n => mapHdr.indexOf(n);

  const IDX_MAP = {
    itemId: mapCol('Item_ID_Machine'),
    brandId: mapCol('Brand_ID_Machine')
  };

  for (const [k, v] of Object.entries(IDX_MAP)) {
    if (v === -1) {
      throw new Error(`Mapping_Item_Brand missing column: ${k}`);
    }
  }

  const existingSet = new Set();
  for (let i = 1; i < mapData.length; i++) {
    const itemId  = mapData[i][IDX_MAP.itemId];
    const brandId = mapData[i][IDX_MAP.brandId];
    if (itemId && brandId) {
      existingSet.add(itemId + '|' + brandId);
    }
  }

  /* ======================================================
     READ TRANSACTION STAGING
     ====================================================== */

  const tsData = tsSh.getDataRange().getValues();
  const tsHdr  = tsData[0];
  const tsCol  = n => tsHdr.indexOf(n);

  const IDX_TS = {
    txnId: tsCol('Txn_ID_Machine'),
    txnDate: tsCol('Txn_Date_Entered'),
    itemId: tsCol('Item_ID_Machine'),
    brandId: tsCol('Brand_ID_Machine'),
    itemCanon: tsCol('Item_Name_Canonical'),
    brandCanon: tsCol('Brand_Name_Canonical')
  };


  for (const [k, v] of Object.entries(IDX_TS)) {
    if (v === -1) {
      throw new Error(`Transaction_Staging missing column: ${k}`);
    }
  }

  /* ======================================================
     DISCOVERY LOOP (FIRST-SEEN PER PAIR)
     ====================================================== */

  const firstSeenMap = new Map();
  let scanned = 0;

  for (let i = 1; i < tsData.length; i++) {
    scanned++;
    const r = tsData[i];

    if (!r[IDX_TS.txnId]) continue;
    if (!r[IDX_TS.itemId] || !r[IDX_TS.brandId]) continue;

    const key = r[IDX_TS.itemId] + '|' + r[IDX_TS.brandId];
    if (firstSeenMap.has(key)) continue;

    firstSeenMap.set(key, {
      itemId: r[IDX_TS.itemId],
      brandId: r[IDX_TS.brandId],
      txnId: r[IDX_TS.txnId],
      txnDate: r[IDX_TS.txnDate] || '',
      itemCanon: r[IDX_TS.itemCanon] || '',
      brandCanon: r[IDX_TS.brandCanon] || ''
    });
  }

  /* ======================================================
     BUILD ROWS TO APPEND
     ====================================================== */

  const rowsToAppend = [];
  let appended = 0;

  for (const [key, v] of firstSeenMap.entries()) {
    if (existingSet.has(key)) continue;

    const newRow = new Array(mapHdr.length).fill('');

    newRow[mapCol('Item_ID_Machine')]        = v.itemId;
    newRow[mapCol('Brand_ID_Machine')]       = v.brandId;
    newRow[mapCol('First_Seen_Txn_ID')]      = v.txnId;
    newRow[mapCol('First_Seen_Txn_Date')]    = v.txnDate;
    newRow[mapCol('Item_Name_Canonical')]    = v.itemCanon;
    newRow[mapCol('Brand_Name_Canonical')]   = v.brandCanon;
    newRow[mapCol('Is_Mapping_Active')]      = true;
    newRow[mapCol('Is_Archived')]            = false;
    newRow[mapCol('Created_At')]             = new Date();
    newRow[mapCol('Notes')]                  = 'Discovered from transaction';

    rowsToAppend.push(newRow);
    existingSet.add(key);
    appended++;
  }

  /* ======================================================
     WRITE MAPPINGS
     ====================================================== */

  if (rowsToAppend.length > 0) {
    mapSh.getRange(
      mapSh.getLastRow() + 1,
      1,
      rowsToAppend.length,
      rowsToAppend[0].length
    ).setValues(rowsToAppend);
  }

  const durationMs = new Date().getTime() - t0.getTime();

  console.log(
    `[${SCRIPT_NAME}] Scanned=${scanned}, Added=${appended}, DurationMs=${durationMs}`
  );

  ETI_log_({
    executionId: EXECUTION_ID,
    scriptName: SCRIPT_NAME,
    sheetName: MAP_SHEET,
    level: 'INFO',
    action: 'SUMMARY',
    details: `Scanned=${scanned}, Added=${appended}, DurationMs=${durationMs}`
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
