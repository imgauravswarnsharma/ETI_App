/**
 * Script Name: populateMapping_Item_Brand_FromTransactionResolution
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Discover Item ↔ Brand relationships observed in Transaction_Resolution
 * - Insert mapping rows into Mapping_Item_Brand
 * - Preserve first-seen transaction evidence
 * - Capture canonical snapshots for audit visibility
 *
 * Mapping Identity:
 * (Item_ID_Machine, Brand_ID_Machine)
 *
 * Preconditions:
 * - Sheet must exist: Transaction_Resolution
 * - Sheet must exist: Mapping_Item_Brand
 * - Header row must exist in row 1
 *
 * Required columns in Transaction_Resolution:
 * - Txn_ID_Machine
 * - Created_At
 * - Item_ID_Machine
 * - Brand_ID_Machine
 * - Item_Name_Canonical
 * - Brand_Name_Canonical
 *
 * Required columns in Mapping_Item_Brand:
 * - Item_Name_Canonical
 * - Brand_Name_Canonical
 * - Item_Status_Snapshot
 * - Brand_Status_Snapshot
 * - Is_Mapping_Active
 * - Is_Analytics_Enabled
 * - Is_Archived
 * - Created_At
 * - Notes
 * - First_Seen_Txn_Date
 * - First_Seen_Txn_ID
 * - Item_ID_Machine
 * - Brand_ID_Machine
 *
 * Algorithm (Step-by-Step):
 *
 * 1. Load Mapping_Item_Brand and resolve column indexes.
 * 2. Build in-memory Set of existing mapping identities:
 *      Item_ID_Machine + Brand_ID_Machine
 *
 * 3. Load Transaction_Resolution rows.
 *
 * 4. Iterate each transaction:
 *
 *    Skip if:
 *      - Txn_ID_Machine missing
 *      - Item_ID_Machine missing
 *      - Brand_ID_Machine missing
 *
 * 5. Construct mapping identity key.
 *
 *    If identity already exists:
 *       → skip row
 *
 * 6. Create new mapping row:
 *
 *      Item_Name_Canonical
 *      Brand_Name_Canonical
 *
 *      Item_Status_Snapshot  = ""
 *      Brand_Status_Snapshot = ""
 *
 *      Is_Mapping_Active     = TRUE
 *      Is_Analytics_Enabled  = TRUE
 *      Is_Archived           = FALSE
 *
 *      Created_At = NOW()
 *      Notes      = "Discovered from Transaction_Resolution"
 *
 *      First_Seen_Txn_Date
 *      First_Seen_Txn_ID
 *
 *      Item_ID_Machine
 *      Brand_ID_Machine
 *
 * 7. Append rows in batch.
 *
 * 8. Emit execution summary and completion logs.
 *
 * Failure Modes:
 * - Required sheet missing
 * - Required column missing
 *
 * Reason for Deprecation:
 * - N/A
 */

function populateMapping_Item_Brand_FromTransactionResolution() {

  const EXECUTION_ID = Utilities.getUuid();
  const SCRIPT_NAME  = 'populateMapping_Item_Brand_FromTransactionResolution';

  const TXN_SHEET = 'Transaction_Resolution';
  const MAP_SHEET = 'Mapping_Item_Brand';

  const t0 = new Date();

  console.log(`[${SCRIPT_NAME}] START`);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const txSh = ss.getSheetByName(TXN_SHEET);
  const mpSh = ss.getSheetByName(MAP_SHEET);

  if (!txSh || !mpSh) {
    throw new Error('Required sheet not found');
  }

  /* =========================
     READ MAPPING TABLE
     ========================= */

  const mpData = mpSh.getDataRange().getValues();
  const mpHdr  = mpData[0];
  const mpCol  = n => mpHdr.indexOf(n);

  const IDX_MAP = {

    itemCanon: mpCol('Item_Name_Canonical'),
    brandCanon: mpCol('Brand_Name_Canonical'),

    itemStatus: mpCol('Item_Status_Snapshot'),
    brandStatus: mpCol('Brand_Status_Snapshot'),

    mapActive: mpCol('Is_Mapping_Active'),
    analytics: mpCol('Is_Analytics_Enabled'),
    archived: mpCol('Is_Archived'),

    createdAt: mpCol('Created_At'),
    notes: mpCol('Notes'),

    firstSeenDate: mpCol('First_Seen_Txn_Date'),
    firstSeenTxn: mpCol('First_Seen_Txn_ID'),

    itemId: mpCol('Item_ID_Machine'),
    brandId: mpCol('Brand_ID_Machine')
  };

  for (const [k,v] of Object.entries(IDX_MAP)) {
    if (v === -1) throw new Error(`Mapping_Item_Brand missing column: ${k}`);
  }

  /* =========================
     BUILD EXISTING IDENTITY SET
     ========================= */

  const existingSet = new Set();

  for (let i = 1; i < mpData.length; i++) {

    const itemId = mpData[i][IDX_MAP.itemId];
    const brandId = mpData[i][IDX_MAP.brandId];

    if (!itemId || !brandId) continue;

    const key = `${itemId}||${brandId}`;

    existingSet.add(key);
  }

  /* =========================
     READ TRANSACTION RESOLUTION
     ========================= */

  const txData = txSh.getDataRange().getValues();
  const txHdr  = txData[0];
  const txCol  = n => txHdr.indexOf(n);

  const IDX_TX = {

    txnId: txCol('Txn_ID_Machine'),

    txnDateEntered: txCol('Txn_Date_Entered'),
    createdAt: txCol('Created_At'),

    itemId: txCol('Item_ID_Machine'),
    brandId: txCol('Brand_ID_Machine'),

    itemCanon: txCol('Item_Name_Canonical'),
    brandCanon: txCol('Brand_Name_Canonical')
  };

  for (const [k,v] of Object.entries(IDX_TX)) {
    if (v === -1) {
      throw new Error(`Transaction_Resolution missing column: ${k}`);
    }
  }

  /* =========================
     COUNTERS
     ========================= */

  let scanned = 0;
  let skipNoTxn = 0;
  let skipNoItem = 0;
  let skipNoBrand = 0;
  let skipDuplicate = 0;

  const rowsToAppend = [];

  /* =========================
     DISCOVERY LOOP
     ========================= */

  for (let i = 1; i < txData.length; i++) {

    scanned++;

    const r = txData[i];

    const txnId = r[IDX_TX.txnId];
    const itemId = r[IDX_TX.itemId];
    const brandId = r[IDX_TX.brandId];

    if (!txnId) {
      skipNoTxn++;
      continue;
    }

    if (!itemId) {
      skipNoItem++;
      continue;
    }

    if (!brandId) {
      skipNoBrand++;
      continue;
    }

    const key = `${itemId}||${brandId}`;

    if (existingSet.has(key)) {
      skipDuplicate++;
      continue;
    }

    const row = new Array(mpHdr.length).fill('');

    row[IDX_MAP.itemCanon] = r[IDX_TX.itemCanon];
    row[IDX_MAP.brandCanon] = r[IDX_TX.brandCanon];

    row[IDX_MAP.itemStatus] = '';
    row[IDX_MAP.brandStatus] = '';

    row[IDX_MAP.mapActive] = true;
    row[IDX_MAP.analytics] = true;
    row[IDX_MAP.archived] = false;

    row[IDX_MAP.createdAt] = new Date();
    row[IDX_MAP.notes] = 'Discovered from Transaction_Resolution';

    let firstSeenDate = r[IDX_TX.txnDateEntered];

    if (!firstSeenDate) {
      firstSeenDate = r[IDX_TX.createdAt];
    }

    if (!firstSeenDate) {
      firstSeenDate = new Date();
    }

row[IDX_MAP.firstSeenDate] = firstSeenDate;
    row[IDX_MAP.firstSeenTxn] = txnId;

    row[IDX_MAP.itemId] = itemId;
    row[IDX_MAP.brandId] = brandId;

    rowsToAppend.push(row);

    existingSet.add(key);
  }

  /* =========================
     BATCH APPEND
     ========================= */

  if (rowsToAppend.length > 0) {

    mpSh.getRange(
      mpSh.getLastRow() + 1,
      1,
      rowsToAppend.length,
      mpHdr.length
    ).setValues(rowsToAppend);
  }

  const durationMs = new Date().getTime() - t0.getTime();

  console.log(`[${SCRIPT_NAME}] Rows scanned: ${scanned}`);
  console.log(`[${SCRIPT_NAME}] Skipped no Txn_ID: ${skipNoTxn}`);
  console.log(`[${SCRIPT_NAME}] Skipped no Item_ID: ${skipNoItem}`);
  console.log(`[${SCRIPT_NAME}] Skipped no Brand_ID: ${skipNoBrand}`);
  console.log(`[${SCRIPT_NAME}] Skipped duplicate: ${skipDuplicate}`);
  console.log(`[${SCRIPT_NAME}] Rows appended: ${rowsToAppend.length}`);
  console.log(`[${SCRIPT_NAME}] END – Duration(ms): ${durationMs}`);

}