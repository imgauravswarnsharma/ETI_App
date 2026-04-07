/**
 * Script Name: populateStagingLookupItems_FromTransactionResolution
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Populate Staging_Lookup_Items with unresolved item canonicals detected in Transaction_Resolution.
 * - Insert one staging row per unique canonical value requiring governance review.
 * - Initialize staging rows with the default governance state ("Review").
 * - Serve as the intake bridge between transaction logging and the staging governance workflow.
 *
 * Preconditions:
 * - Sheet must exist: Transaction_Resolution
 * - Sheet must exist: Staging_Lookup_Items
 * - Header row present in row 1 for both sheets
 *
 * Required columns in Transaction_Resolution:
 *   - Txn_ID_Machine
 *   - Item_ID_Machine
 *   - Item_Name_Entered
 *   - Item_Name_Canonical
 *
 * Required columns in Staging_Lookup_Items:
 *   - Source_Txn_ID_Machine
 *   - Staging_Item_ID_Machine
 *   - Mapped_Item_ID_Machine
 *   - Item_Name_Entered
 *   - Item_Name_Canonical
 *   - Item_Name_Approved
 *   - Admin_Action
 *   - Is_Approved
 *   - Is_Active
 *   - Is_Archived
 *   - Is_Lookup_Promoted
 *   - Populated_At
 *   - Notes
 *
 * Algorithm (Step-by-Step):
 *
 * 1. Load Staging_Lookup_Items and read the header row.
 *    - Resolve column indexes dynamically using header names.
 *
 * 2. Build an in-memory Set of existing staged canonicals.
 *    - Ensures only one staging row exists per canonical value.
 *
 * 3. Load Transaction_Resolution rows.
 *    - Resolve required column indexes dynamically.
 *
 * 4. Iterate through Transaction_Resolution rows:
 *
 *    a. Skip rows where Txn_ID_Machine is missing.
 *
 *    b. Skip rows where Item_ID_Machine already exists
 *       (item already resolved in lookup).
 *
 *    c. Skip rows where Item_Name_Canonical is missing.
 *
 *    d. Skip rows where the canonical value already exists
 *       in the staging canonical Set.
 *
 *    e. For each remaining row:
 *       - Construct a new staging row array sized to the
 *         Staging_Lookup_Items header length.
 *
 *       Populate the following fields:
 *
 *       Source_Txn_ID_Machine   ← transaction reference
 *       Staging_Item_ID_Machine ← generated UUID
 *       Mapped_Item_ID_Machine  ← blank
 *       Item_Name_Entered       ← original entered value
 *       Item_Name_Canonical     ← canonical value
 *       Item_Name_Approved      ← blank
 *
 *       Admin_Action            ← "Review"
 *       Is_Approved             ← FALSE
 *       Is_Active               ← FALSE
 *       Is_Archived             ← FALSE
 *       Is_Lookup_Promoted      ← FALSE
 *
 *       Populated_At            ← script timestamp
 *       Notes                   ← "Staged from Transaction_Resolution"
 *
 *       All other staging columns remain blank so that
 *       sheet formulas populate derived values.
 *
 *    f. Add the canonical value to the in-memory Set to
 *       prevent duplicate staging rows within the same run.
 *
 * 5. Batch append all new rows to Staging_Lookup_Items.
 *
 * 6. Emit execution summary metrics to console logs
 *    and ETI_log_ for auditability.
 *
 * Failure Modes:
 * - Required sheet missing
 * - Required column missing
 *
 * Reason for Deprecation:
 * - N/A
 */

function populateStagingLookupItems_FromTransactionResolution() {

  const EXECUTION_ID = Utilities.getUuid();
  const SCRIPT_NAME = 'populateStagingLookupItems_FromTransactionResolution';
  const TXN_SHEET = 'Transaction_Resolution';
  const STG_SHEET = 'Staging_Lookup_Items';

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
    sourceTxn: stgCol('Source_Txn_ID_Machine'),
    stagingId: stgCol('Staging_Item_ID_Machine'),
    mappedId: stgCol('Mapped_Item_ID_Machine'),
    entered: stgCol('Item_Name_Entered'),
    canon: stgCol('Item_Name_Canonical'),
    approvedName: stgCol('Item_Name_Approved'),
    adminAction: stgCol('Admin_Action'),
    isApproved: stgCol('Is_Approved'),
    isActive: stgCol('Is_Active'),
    isArchived: stgCol('Is_Archived'),
    isPromoted: stgCol('Is_Lookup_Promoted'),
    populatedAt: stgCol('Populated_At'),
    notes: stgCol('Notes')
  };

  for (const [k,v] of Object.entries(IDX_STG)) {
    if (v === -1) throw new Error(`Staging_Lookup_Items missing column: ${k}`);
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
    itemEntered: tsCol('Item_Name_Entered'),
    itemCanon: tsCol('Item_Name_Canonical')
  };

  for (const [k,v] of Object.entries(IDX)) {
    if (v === -1) {
      throw new Error(`Transaction_Resolution missing column: ${k}`);
    }
  }

  /* ---------- Counters ---------- */

  let scanned = 0;
  let skipNoTxn = 0;
  let skipHasItem = 0;
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

    if (r[IDX.itemId]) {
      skipHasItem++;
      continue;
    }

    const canon = r[IDX.itemCanon];

    if (!canon) {
      skipNoCanon++;
      continue;
    }

    if (stagingCanonSet.has(canon)) {
      skipDuplicateCanon++;
      continue;
    }

    const row = new Array(stgHdr.length).fill('');

    row[IDX_STG.sourceTxn] = r[IDX.txnId];
    row[IDX_STG.stagingId] = Utilities.getUuid();
    row[IDX_STG.mappedId] = '';
    row[IDX_STG.entered] = r[IDX.itemEntered];
    row[IDX_STG.canon] = canon;
    row[IDX_STG.approvedName] = '';

    row[IDX_STG.adminAction] = 'Review';
    row[IDX_STG.isApproved] = false;
    row[IDX_STG.isActive] = false;
    row[IDX_STG.isArchived] = false;
    row[IDX_STG.isPromoted] = false;

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
  console.log(`[${SCRIPT_NAME}] Skipped has Item_ID: ${skipHasItem}`);
  console.log(`[${SCRIPT_NAME}] Skipped no canonical: ${skipNoCanon}`);
  console.log(`[${SCRIPT_NAME}] Skipped duplicate canonical: ${skipDuplicateCanon}`);
  console.log(`[${SCRIPT_NAME}] Rows appended: ${rowsToAppend.length}`);
  console.log(`[${SCRIPT_NAME}] END – Duration(ms): ${durationMs}`);

  ETI_log_({
    executionId: EXECUTION_ID,
    scriptName: SCRIPT_NAME,
    sheetName: STG_SHEET,
    level: 'INFO',
    action: 'SUMMARY',
    details:
      `Scanned=${scanned}, ` +
      `SkipNoTxn=${skipNoTxn}, ` +
      `SkipHasItem=${skipHasItem}, ` +
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