/**
 * Script Name: promoteApprovedProducts_FromStaging_ToLookup
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Promote approved staging product rows into Lookup_Products exactly once
 * - Promotion authority = staging row (not canonical, not name)
 * - Record lineage via Is_Staging_Promoted (Lookup_Products)
 * - Record completion via Is_Lookup_Promoted (Staging_Lookup_Products)
 * - Preserve auditability and forward-only identity guarantees (ETI v1.3)
 *
 * Preconditions:
 * - Sheets must exist:
 *   - Staging_Lookup_Products
 *   - Lookup_Products
 * - Header row must exist in row 1
 * - Required columns must exist (header-based, order-independent)
 * - Flags Is_Lookup_Promoted and Is_Staging_Promoted are script-owned
 *
 * Algorithm (Step-by-Step):
 * 1. Generate Execution_ID for traceability.
 * 2. Load Lookup_Products and resolve column indexes via headers.
 * 3. Load Staging_Lookup_Products and resolve column indexes via headers.
 * 4. Iterate each staging row:
 *    a. Skip if Is_Approved ≠ TRUE.
 *    b. Skip if Is_Lookup_Promoted = TRUE (hard stop).
 *    c. Resolve final product name (Approved → fallback Entered).
 *    d. Generate full UUID for Product_ID_Machine.
 *    e. Construct a new Lookup_Products row using header-indexed placement.
 *    f. Mark Is_Staging_Promoted = TRUE in Lookup_Products.
 *    g. Write back Mapped_Product_ID_Machine, Is_Lookup_Promoted = TRUE,
 *       and contextual Notes into Staging_Lookup_Products.
 * 5. Batch-append all new Lookup_Products rows.
 * 6. Write back all staging updates.
 * 7. Emit execution summary and completion logs.
 *
 * Failure Modes:
 * - Missing required sheet
 * - Missing required column
 *
 * Reason for Deprecation (if applicable):
 * - N/A
 * - This script remains ACTIVE for ETI v1.3.
 * - Superseded only in v1.4 if product reconciliation
 *   or merge workflows are introduced.
 */

function promoteApprovedProducts_FromStaging_ToLookup() {

  const EXECUTION_ID = Utilities.getUuid();
  const SCRIPT_NAME  = 'promoteApprovedProducts_FromStaging_ToLookup';
  const STG_SHEET    = 'Staging_Lookup_Products';
  const LKP_SHEET    = 'Lookup_Products';

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
  const stgSh = ss.getSheetByName(STG_SHEET);
  const lkSh  = ss.getSheetByName(LKP_SHEET);

  if (!stgSh || !lkSh) {
    throw new Error('Required sheet not found');
  }

  /* ===================== READ LOOKUP ===================== */

  const lkData = lkSh.getDataRange().getValues();
  const lkHdr  = lkData[0];
  const lkCol  = n => lkHdr.indexOf(n);

  const IDX_LK = {
    productId: lkCol('Product_ID_Machine'),
    productHuman: lkCol('Product_ID_Human'),
    productName: lkCol('Product_Name'),
    productCanon: lkCol('Product_Name_Canonical'),
    isActive: lkCol('Is_Active'),
    isArchived: lkCol('Is_Archived'),
    isStgPromoted: lkCol('Is_Staging_Promoted'),
    notes: lkCol('Notes')
  };

  for (const [k, v] of Object.entries(IDX_LK)) {
    if (v === -1) {
      throw new Error(`Lookup_Products missing column: ${k}`);
    }
  }

  /* ===================== READ STAGING ===================== */

  const stgData = stgSh.getDataRange().getValues();
  const stgHdr  = stgData[0];
  const stgCol  = n => stgHdr.indexOf(n);

  const IDX_STG = {
    entered: stgCol('Prodcut_Name_Entered'),
    approved: stgCol('Prodcut_Name_Approved'),
    canon: stgCol('Prodcut_Name_Canonical'),
    isApproved: stgCol('Is_Approved'),
    isPromoted: stgCol('Is_Lookup_Promoted'),
    mappedId: stgCol('Mapped_Prodcut_ID_Machine'),
    notes: stgCol('Notes')
  };

  for (const [k, v] of Object.entries(IDX_STG)) {
    if (v === -1) {
      throw new Error(`Staging_Lookup_Products missing column: ${k}`);
    }
  }

  /* ===================== PROMOTION LOOP ===================== */

  const lookupAppendRows = [];
  const stagingUpdates  = [];

  let scanned = 0;
  let promoted = 0;
  let skipped = 0;

  for (let i = 1; i < stgData.length; i++) {
    scanned++;
    const rowNum = i + 1;
    const r = stgData[i];

    if (r[IDX_STG.isApproved] !== true) { skipped++; continue; }
    if (r[IDX_STG.isPromoted] === true) { skipped++; continue; }

    const finalName =
      r[IDX_STG.approved] ||
      r[IDX_STG.entered];

    if (!finalName) { skipped++; continue; }

    const canon = r[IDX_STG.canon] || '';

    const productIdMachine = Utilities.getUuid(); // FULL UUID

    const newLookupRow = new Array(lkHdr.length).fill('');
    newLookupRow[IDX_LK.productId]        = productIdMachine;
    newLookupRow[IDX_LK.productHuman]     = '';
    newLookupRow[IDX_LK.productName]      = finalName;
    newLookupRow[IDX_LK.productCanon]     = canon;
    newLookupRow[IDX_LK.isActive]         = true;
    newLookupRow[IDX_LK.isArchived]       = false;
    newLookupRow[IDX_LK.isStgPromoted]    = true;
    newLookupRow[IDX_LK.notes]            = 'Promoted from staging';

    lookupAppendRows.push(newLookupRow);

    const existingNote = r[IDX_STG.notes] || '';
    const newNote =
      (existingNote ? existingNote + ' | ' : '') +
      `Promoted to Lookup_Products → Product_ID_Machine=${productIdMachine}`;

    stagingUpdates.push({
      row: rowNum,
      mappedId: productIdMachine,
      note: newNote
    });

    promoted++;

    console.log(
      `[${SCRIPT_NAME}] PROMOTED row ${rowNum} → Product_ID_Machine=${productIdMachine}`
    );

    ETI_log_({
      executionId: EXECUTION_ID,
      scriptName: SCRIPT_NAME,
      sheetName: STG_SHEET,
      level: 'INFO',
      rowNumber: rowNum,
      action: 'PROMOTE_PRODUCT',
      details: `Promoted staging product → Product_ID_Machine=${productIdMachine}`
    });
  }

  /* ===================== WRITE LOOKUP ===================== */

  if (lookupAppendRows.length > 0) {
    lkSh.getRange(
      lkSh.getLastRow() + 1,
      1,
      lookupAppendRows.length,
      lookupAppendRows[0].length
    ).setValues(lookupAppendRows);
  }

  /* ===================== WRITE BACK TO STAGING ===================== */

  for (const u of stagingUpdates) {
    stgSh.getRange(u.row, IDX_STG.mappedId + 1).setValue(u.mappedId);
    stgSh.getRange(u.row, IDX_STG.isPromoted + 1).setValue(true);
    stgSh.getRange(u.row, IDX_STG.notes + 1).setValue(u.note);
  }

  const durationMs = new Date().getTime() - t0.getTime();

  console.log(
    `[${SCRIPT_NAME}] Scanned=${scanned}, Promoted=${promoted}, Skipped=${skipped}, DurationMs=${durationMs}`
  );

  ETI_log_({
    executionId: EXECUTION_ID,
    scriptName: SCRIPT_NAME,
    sheetName: STG_SHEET,
    level: 'INFO',
    action: 'SUMMARY',
    details: `Scanned=${scanned}, Promoted=${promoted}, Skipped=${skipped}, DurationMs=${durationMs}`
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
