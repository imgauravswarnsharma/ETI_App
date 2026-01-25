/**
 * Script Name: populateStagingLookupItems_FromTransactionResolution
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Populate Staging_Lookup_Items with unresolved item canonicals
 * - One row per canonical value
 * - Acts as intake + review bridge for item approval workflow
 * - Forward-only, non-reconciling by design
 *
 * Preconditions:
 * - Sheet must exist: Transaction_Resolution
 * - Sheet must exist: Staging_Lookup_Items
 * - Header row present in row 1 for both sheets
 * - Required columns in Transaction_Resolution:
 *   - Txn_ID_Machine
 *   - Item_ID_Machine
 *   - Item_Name_Entered
 *   - Item_Name_Canonical
 * - Required columns in Staging_Lookup_Items:
 *   - Item_Name_Canonical
 *
 * Algorithm (Step-by-Step):
 * 1. Read existing canonicals from Staging_Lookup_Items
 * 2. Build in-memory canonical de-duplication set
 * 3. Read Transaction_Resolution rows
 * 4. For each row:
 *    a. Skip if Txn_ID missing
 *    b. Skip if Item_ID already present
 *    c. Skip if canonical missing
 *    d. Skip if canonical already staged
 *    e. Otherwise append new staging row
 * 5. Batch write appended rows
 * 6. Emit execution summary
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
  const tsSh  = ss.getSheetByName(TXN_SHEET);
  const stgSh = ss.getSheetByName(STG_SHEET);

  if (!tsSh || !stgSh) {
    throw new Error('Required sheet not found');
  }

  /* ---------- Read existing staging canonicals ---------- */
  const stgData = stgSh.getDataRange().getValues();
  const stgHdr  = stgData[0];
  const stgCol  = n => stgHdr.indexOf(n);

  const IDX_STG_CANON = stgCol('Item_Name_Canonical');
  if (IDX_STG_CANON === -1) {
    throw new Error('Staging_Lookup_Items.Item_Name_Canonical missing');
  }

  const stagingCanonSet = new Set();
  for (let i = 1; i < stgData.length; i++) {
    const v = stgData[i][IDX_STG_CANON];
    if (v) stagingCanonSet.add(String(v));
  }

  console.log(`[${SCRIPT_NAME}] Existing staging canonicals: ${stagingCanonSet.size}`);

  /* ---------- Read transaction Resolution ---------- */
  const tsData = tsSh.getDataRange().getValues();
  const tsHdr  = tsData[0];
  const tsCol  = n => tsHdr.indexOf(n);

  const IDX = {
    txnId: tsCol('Txn_ID_Machine'),
    itemId: tsCol('Item_ID_Machine'),
    itemEntered: tsCol('Item_Name_Entered'),
    itemCanon: tsCol('Item_Name_Canonical')
  };

  for (const [k, v] of Object.entries(IDX)) {
    if (v === -1) {
      throw new Error(`Transaction_Resolution missing column: ${k}`);
    }
  }

  /* ---------- Counters (logging only) ---------- */
  let scanned = 0;
  let skipNoTxn = 0;
  let skipHasItem = 0;
  let skipNoCanon = 0;
  let skipDuplicateCanon = 0;

  const rowsToAppend = [];

  for (let i = 1; i < tsData.length; i++) {
    scanned++;
    const r = tsData[i];

    // --- ORIGINAL GATES (UNCHANGED) ---
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

    /* ---------- WRITE ROW (COLUMN ORDER FIXED) ---------- */
    rowsToAppend.push([
      r[IDX.txnId],                                   // Source_Txn_ID_Machine
      Utilities.getUuid(), // Staging_Item_ID_Machine
      '',                                             // Mapped_Item_ID_Machine
      r[IDX.itemEntered],                             // Item_Name_Entered
      canon,                                          // Item_Name_Canonical
      '',                                             // Item_Name_Approved
      false,                                          // Is_Approved
      true,                                           // Is_Active
      false,                                          // Is_Archived
      'Pending',                                      // Review_Status
      ''                                              // Notes
    ]);

    stagingCanonSet.add(canon); // enforce one row per canonical (unchanged)
  }

  if (rowsToAppend.length > 0) {
    stgSh.getRange(
      stgSh.getLastRow() + 1,
      1,
      rowsToAppend.length,
      rowsToAppend[0].length
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
