/**
 * Script Name: populateMapping_Item_Brand_Product_FromTransactionResolution
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Populate Mapping_Item_Brand_Product with unique
 *   (Item_ID_Machine, Brand_ID_Machine, Product_ID_Machine) relationships
 *   discovered from Transaction_Resolution.
 * - Each mapping is recorded exactly once (first-seen semantics).
 * - Transaction rows serve strictly as discovery evidence.
 *
 * Preconditions:
 * - Sheets must exist:
 *     Transaction_Resolution
 *     Mapping_Item_Brand_Product
 * - Header row present in row 1
 * - Required columns must exist (header-based)
 *
 * Algorithm (Step-by-Step):
 *
 * 1. Generate Execution_ID for traceability.
 * 2. Load Mapping_Item_Brand_Product and build an in-memory set of existing
 *    identity keys:
 *
 *        Item_ID | Brand_ID | Product_ID
 *
 * 3. Load Transaction_Resolution.
 *
 * 4. Iterate rows sequentially:
 *
 *      Skip if:
 *        - Txn_ID_Machine missing
 *        - Item_ID_Machine missing
 *        - Brand_ID_Machine missing
 *        - Product_ID_Machine missing
 *
 *      Build identity key.
 *
 *      Capture only the FIRST occurrence of each identity key
 *      within the current execution batch.
 *
 * 5. Construct mapping rows using:
 *
 *      Canonical snapshots
 *      Evidence metadata
 *      Default governance flags
 *
 * 6. Append rows in batch write.
 *
 * 7. Emit execution summary logs.
 *
 * Failure Modes:
 * - Required sheet missing
 * - Required column missing
 *
 * Notes:
 * - Discovery script NEVER modifies existing rows.
 * - Governance reconciliation handled by processing script.
 */

function populateMapping_Item_Brand_Product_FromTransactionResolution() {

  const EXECUTION_ID = Utilities.getUuid();
  const SCRIPT_NAME  = 'populateMapping_Item_Brand_Product_FromTransactionResolution';

  const TXN_SHEET = 'Transaction_Resolution';
  const MAP_SHEET = 'Mapping_Item_Brand_Product';

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
  const tsSh  = ss.getSheetByName(TXN_SHEET);
  const mapSh = ss.getSheetByName(MAP_SHEET);

  if (!tsSh || !mapSh) {
    throw new Error('Required sheet not found');
  }

  /* =====================================================
     READ EXISTING MAPPINGS
     ===================================================== */

  const mapData = mapSh.getDataRange().getValues();
  const mapHdr  = mapData[0];

  const mapCol = n => mapHdr.indexOf(n);

  const IDX_MAP = {
    itemId: mapCol('Item_ID_Machine'),
    brandId: mapCol('Brand_ID_Machine'),
    productId: mapCol('Product_ID_Machine')
  };

  for (const [k,v] of Object.entries(IDX_MAP)) {
    if (v === -1) {
      throw new Error(`Mapping_Item_Brand_Product missing column: ${k}`);
    }
  }

  const existingSet = new Set();

  for (let i = 1; i < mapData.length; i++) {

    const iId = mapData[i][IDX_MAP.itemId];
    const bId = mapData[i][IDX_MAP.brandId];
    const pId = mapData[i][IDX_MAP.productId];

    if (iId && bId && pId) {
      existingSet.add(iId + '|' + bId + '|' + pId);
    }
  }

  /* =====================================================
     READ TRANSACTION_RESOLUTION
     ===================================================== */

  const tsData = tsSh.getDataRange().getValues();
  const tsHdr  = tsData[0];

  const txCol = n => tsHdr.indexOf(n);

  const IDX_TX = {

    txnId: txCol('Txn_ID_Machine'),
    txnDateEntered: txCol('Txn_Date_Entered'),
    createdAt: txCol('Created_At'),

    itemId: txCol('Item_ID_Machine'),
    brandId: txCol('Brand_ID_Machine'),
    productId: txCol('Product_ID_Machine'),

    itemCanon: txCol('Item_Name_Canonical'),
    brandCanon: txCol('Brand_Name_Canonical'),
    productCanon: txCol('Product_Name_Canonical')
  };

  for (const [k,v] of Object.entries(IDX_TX)) {
    if (v === -1) {
      throw new Error(`Transaction_Resolution missing column: ${k}`);
    }
  }

  /* =====================================================
     DISCOVERY LOOP
     ===================================================== */

  const firstSeenMap = new Map();

  let scanned = 0;

  for (let i = 1; i < tsData.length; i++) {

    scanned++;

    const r = tsData[i];

    if (!r[IDX_TX.txnId]) continue;

    if (!r[IDX_TX.itemId] ||
        !r[IDX_TX.brandId] ||
        !r[IDX_TX.productId]) continue;

    const key =
      r[IDX_TX.itemId] + '|' +
      r[IDX_TX.brandId] + '|' +
      r[IDX_TX.productId];

    if (firstSeenMap.has(key)) continue;

    /* resolve first seen date safely */

    let firstSeenDate = r[IDX_TX.txnDateEntered];

    if (!firstSeenDate) {
      firstSeenDate = r[IDX_TX.createdAt];
    }

    if (!firstSeenDate) {
      firstSeenDate = new Date();
    }

    firstSeenMap.set(key, {

      itemId: r[IDX_TX.itemId],
      brandId: r[IDX_TX.brandId],
      productId: r[IDX_TX.productId],

      txnId: r[IDX_TX.txnId],
      txnDate: firstSeenDate,

      itemCanon: r[IDX_TX.itemCanon] || '',
      brandCanon: r[IDX_TX.brandCanon] || '',
      productCanon: r[IDX_TX.productCanon] || ''
    });
  }

  /* =====================================================
     BUILD ROWS
     ===================================================== */

  const rowsToAppend = [];

  let appended = 0;

  for (const [key,v] of firstSeenMap.entries()) {

    if (existingSet.has(key)) continue;

    const row = new Array(mapHdr.length).fill('');

    row[mapCol('Item_Name_Canonical')]    = v.itemCanon;
    row[mapCol('Brand_Name_Canonical')]   = v.brandCanon;
    row[mapCol('Product_Name_Canonical')] = v.productCanon;

    row[mapCol('Is_Mapping_Active')]      = true;
    row[mapCol('Is_Analytics_Enabled')]   = true;
    row[mapCol('Is_Archived')]            = false;

    row[mapCol('Created_At')]             = new Date();
    row[mapCol('Notes')]                  = 'Discovered from Transaction_Resolution';

    row[mapCol('First_Seen_Txn_Date')]    = v.txnDate;
    row[mapCol('First_Seen_Txn_ID')]      = v.txnId;

    row[mapCol('Item_ID_Machine')]        = v.itemId;
    row[mapCol('Brand_ID_Machine')]       = v.brandId;
    row[mapCol('Product_ID_Machine')]     = v.productId;

    rowsToAppend.push(row);

    existingSet.add(key);
    appended++;
  }

  /* =====================================================
     WRITE
     ===================================================== */

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