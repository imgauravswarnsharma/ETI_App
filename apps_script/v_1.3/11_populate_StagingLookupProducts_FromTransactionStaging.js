/**
 * Script Name: populateStagingLookupProducts_FromTransactionResolution
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Populate Staging_Lookup_Products with unresolved product canonicals
 *   detected in Transaction_Resolution.
 * - Insert one staging row per unique canonical product requiring governance review.
 * - Initialize staging rows with the default governance state ("Review").
 * - Include item and brand context for governance visibility.
 *
 * Preconditions:
 * - Sheet must exist: Transaction_Resolution
 * - Sheet must exist: Staging_Lookup_Products
 * - Header row present in row 1
 *
 * Required columns in Transaction_Resolution:
 *   - Txn_ID_Machine
 *   - Item_ID_Machine
 *   - Product_ID_Machine
 *   - Item_Name_Entered
 *   - Brand_Name_Entered
 *   - Product_Name_Entered
 *   - Product_Name_Canonical
 *
 * Required columns in Staging_Lookup_Products:
 *   - Source_Item_Name
 *   - Source_Brand_Name
 *   - Product_Name_Entered
 *   - Product_Name_Canonical
 *   - Product_Name_Approved
 *   - Admin_Action
 *   - Is_Approved
 *   - Is_Active
 *   - Is_Archived
 *   - Is_Lookup_Promoted
 *   - Populated_At
 *   - Notes
 *   - Staging_Product_ID_Machine
 *   - Mapped_Product_ID_Machine
 *   - Source_Txn_ID_Machine
 *   - Source_Item_ID_Machine
 *   - Source_Product_ID_Machine
 *
 * Algorithm (Step-by-Step):
 *
 * 1. Load Staging_Lookup_Products and read header row.
 * 2. Build in-memory Set of existing staged product canonicals.
 * 3. Load Transaction_Resolution rows.
 * 4. Iterate rows:
 *
 *    Skip if:
 *      - Txn_ID_Machine missing
 *      - Product_ID_Machine exists
 *      - Product_Name_Canonical missing
 *      - canonical already staged
 *
 * 5. Create staging row:
 *
 *    Source_Item_Name
 *    Source_Brand_Name
 *
 *    Product_Name_Entered
 *    Product_Name_Canonical
 *    Product_Name_Approved
 *
 *    Admin_Action = Review
 *
 *    Is_Approved = FALSE
 *    Is_Active = FALSE
 *    Is_Archived = FALSE
 *    Is_Lookup_Promoted = FALSE
 *
 *    Source_Txn_ID_Machine
 *    Source_Item_ID_Machine
 *    Source_Product_ID_Machine
 *
 *    Staging_Product_ID_Machine (UUID)
 *
 *    Populated_At
 *    Notes
 *
 * 6. Batch append rows.
 * 7. Emit execution summary.
 *
 * Failure Modes:
 * - Required sheet missing
 * - Required column missing
 *
 * Reason for Deprecation:
 * - N/A
 */

function populateStagingLookupProducts_FromTransactionResolution() {

  const EXECUTION_ID = Utilities.getUuid();
  const SCRIPT_NAME = 'populateStagingLookupProducts_FromTransactionResolution';

  const TXN_SHEET = 'Transaction_Resolution';
  const STG_SHEET = 'Staging_Lookup_Products';

  const t0 = new Date();

  console.log(`[${SCRIPT_NAME}] START`);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tsSh = ss.getSheetByName(TXN_SHEET);
  const stgSh = ss.getSheetByName(STG_SHEET);

  if (!tsSh || !stgSh) {
    throw new Error('Required sheet not found');
  }

  /* ---------- Read staging header ---------- */

  const stgData = stgSh.getDataRange().getValues();
  const stgHdr = stgData[0];
  const stgCol = n => stgHdr.indexOf(n);

  const IDX_STG = {

    sourceItemName: stgCol('Source_Item_Name'),
    sourceBrandName: stgCol('Source_Brand_Name'),

    entered: stgCol('Product_Name_Entered'),
    canon: stgCol('Product_Name_Canonical'),
    approvedName: stgCol('Product_Name_Approved'),

    adminAction: stgCol('Admin_Action'),

    isApproved: stgCol('Is_Approved'),
    isActive: stgCol('Is_Active'),
    isArchived: stgCol('Is_Archived'),
    isPromoted: stgCol('Is_Lookup_Promoted'),

    populatedAt: stgCol('Populated_At'),
    notes: stgCol('Notes'),

    stagingId: stgCol('Staging_Product_ID_Machine'),
    mappedId: stgCol('Mapped_Product_ID_Machine'),

    sourceTxn: stgCol('Source_Txn_ID_Machine'),
    sourceItem: stgCol('Source_Item_ID_Machine'),
    sourceProduct: stgCol('Source_Product_ID_Machine')
  };

  for (const [k,v] of Object.entries(IDX_STG)) {
    if (v === -1) throw new Error(`Staging_Lookup_Products missing column: ${k}`);
  }

  /* ---------- Existing staging canonicals ---------- */

  const stagingCanonSet = new Set();

  for (let i = 1; i < stgData.length; i++) {
    const v = stgData[i][IDX_STG.canon];
    if (v) stagingCanonSet.add(String(v));
  }

  /* ---------- Read Transaction Resolution ---------- */

  const tsData = tsSh.getDataRange().getValues();
  const tsHdr = tsData[0];
  const tsCol = n => tsHdr.indexOf(n);

  const IDX = {

    txnId: tsCol('Txn_ID_Machine'),
    itemId: tsCol('Item_ID_Machine'),
    productId: tsCol('Product_ID_Machine'),

    itemName: tsCol('Item_Name_Entered'),
    brandName: tsCol('Brand_Name_Entered'),

    productEntered: tsCol('Product_Name_Entered'),
    productCanon: tsCol('Product_Name_Canonical')
  };

  for (const [k,v] of Object.entries(IDX)) {
    if (v === -1) {
      throw new Error(`Transaction_Resolution missing column: ${k}`);
    }
  }

  /* ---------- Counters ---------- */

  let scanned = 0;
  let skipNoTxn = 0;
  let skipHasProduct = 0;
  let skipNoCanon = 0;
  let skipDuplicateCanon = 0;

  const rowsToAppend = [];

  /* ---------- Processing Loop ---------- */

  for (let i = 1; i < tsData.length; i++) {

    scanned++;

    const r = tsData[i];

    if (!r[IDX.txnId]) {
      skipNoTxn++;
      continue;
    }

    if (r[IDX.productId]) {
      skipHasProduct++;
      continue;
    }

    const canon = r[IDX.productCanon];

    if (!canon) {
      skipNoCanon++;
      continue;
    }

    if (stagingCanonSet.has(canon)) {
      skipDuplicateCanon++;
      continue;
    }

    const row = new Array(stgHdr.length).fill('');

    row[IDX_STG.sourceItemName] = r[IDX.itemName];
    row[IDX_STG.sourceBrandName] = r[IDX.brandName];

    row[IDX_STG.entered] = r[IDX.productEntered];
    row[IDX_STG.canon] = canon;
    row[IDX_STG.approvedName] = '';

    row[IDX_STG.adminAction] = 'Review';

    row[IDX_STG.isApproved] = false;
    row[IDX_STG.isActive] = false;
    row[IDX_STG.isArchived] = false;
    row[IDX_STG.isPromoted] = false;

    row[IDX_STG.sourceTxn] = r[IDX.txnId];
    row[IDX_STG.sourceItem] = r[IDX.itemId];
    row[IDX_STG.sourceProduct] = r[IDX.productId];

    row[IDX_STG.stagingId] = Utilities.getUuid();
    row[IDX_STG.mappedId] = '';

    row[IDX_STG.populatedAt] = new Date();

    row[IDX_STG.notes] = 'Staged from Transaction_Resolution';

    rowsToAppend.push(row);

    stagingCanonSet.add(canon);
  }

  /* ---------- Batch Write ---------- */

  if (rowsToAppend.length > 0) {

    stgSh.getRange(
      stgSh.getLastRow() + 1,
      1,
      rowsToAppend.length,
      stgHdr.length
    ).setValues(rowsToAppend);
  }

  const durationMs = new Date().getTime() - t0.getTime();

  console.log(`[${SCRIPT_NAME}] Rows scanned: ${scanned}`);
  console.log(`[${SCRIPT_NAME}] Skipped no Txn_ID: ${skipNoTxn}`);
  console.log(`[${SCRIPT_NAME}] Skipped has Product_ID: ${skipHasProduct}`);
  console.log(`[${SCRIPT_NAME}] Skipped no canonical: ${skipNoCanon}`);
  console.log(`[${SCRIPT_NAME}] Skipped duplicate canonical: ${skipDuplicateCanon}`);
  console.log(`[${SCRIPT_NAME}] Rows appended: ${rowsToAppend.length}`);
  console.log(`[${SCRIPT_NAME}] END – Duration(ms): ${durationMs}`);

}