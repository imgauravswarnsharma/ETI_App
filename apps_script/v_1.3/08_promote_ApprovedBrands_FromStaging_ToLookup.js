/**
 * Script Name: promoteApprovedBrands_FromStaging_ToLookup
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Promote approved staging brand rows into Lookup_Brands exactly once
 * - Promotion authority = staging row (not canonical, not name)
 * - Record lineage via Is_Staging_Promoted (Lookup_Brands)
 * - Record completion via Is_Lookup_Promoted (Staging_Lookup_Brands)
 * - Preserve auditability and forward-only identity guarantees (ETI v1.3)
 *
 * Preconditions:
 * - Sheets must exist:
 *   - Staging_Lookup_Brands
 *   - Lookup_Brands
 * - Header row must exist in row 1
 * - Required columns must exist (header-based, order-independent)
 * - Flags Is_Lookup_Promoted and Is_Staging_Promoted are script-owned
 *
 * Algorithm (Step-by-Step):
 * 1. Generate a unique Execution_ID for traceability.
 * 2. Load Lookup_Brands and resolve column indexes via headers.
 * 3. Load Staging_Lookup_Brands and resolve column indexes via headers.
 * 4. Iterate each staging row:
 *    a. Skip if Is_Approved ≠ TRUE.
 *    b. Skip if Is_Lookup_Promoted = TRUE (hard stop).
 *    c. Resolve final brand name (Brand_Name_Approved → fallback Brand_Name_Entered).
 *    d. Generate full UUID for Brand_ID_Machine.
 *    e. Construct a new Lookup_Brands row using header-indexed placement.
 *    f. Mark Is_Staging_Promoted = TRUE in Lookup_Brands.
 *    g. Write back Mapped_Brand_ID_Machine, Is_Lookup_Promoted = TRUE,
 *       and contextual Notes into Staging_Lookup_Brands.
 * 5. Batch-append all new Lookup_Brands rows.
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
 * - Superseded only in v1.4 if identity reconciliation
 *   or brand-merge workflows are introduced.
 */

function promoteApprovedBrands_FromStaging_ToLookup() {

  const EXECUTION_ID = Utilities.getUuid();
  const SCRIPT_NAME  = 'promoteApprovedBrands_FromStaging_ToLookup';
  const STG_SHEET    = 'Staging_Lookup_Brands';
  const LKP_SHEET    = 'Lookup_Brands';

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
    brandId: lkCol('Brand_ID_Machine'),
    brandHuman: lkCol('Brand_ID_Human'),
    brandName: lkCol('Brand_Name'),
    brandCanon: lkCol('Brand_Name_Canonical'),
    isActive: lkCol('Is_Active'),
    isArchived: lkCol('Is_Archived'),
    isStgPromoted: lkCol('Is_Staging_Promoted'),
    notes: lkCol('Notes')
  };

  for (const [k, v] of Object.entries(IDX_LK)) {
    if (v === -1) {
      throw new Error(`Lookup_Brands missing column: ${k}`);
    }
  }

  /* ===================== READ STAGING ===================== */

  const stgData = stgSh.getDataRange().getValues();
  const stgHdr  = stgData[0];
  const stgCol  = n => stgHdr.indexOf(n);

  const IDX_STG = {
    entered: stgCol('Brand_Name_Entered'),
    approved: stgCol('Brand_Name_Approved'),
    canon: stgCol('Brand_Name_Canonical'),
    isApproved: stgCol('Is_Approved'),
    isPromoted: stgCol('Is_Lookup_Promoted'),
    mappedId: stgCol('Mapped_Brand_ID_Machine'),
    notes: stgCol('Notes')
  };

  for (const [k, v] of Object.entries(IDX_STG)) {
    if (v === -1) {
      throw new Error(`Staging_Lookup_Brands missing column: ${k}`);
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

    const brandIdMachine = Utilities.getUuid(); // FULL UUID

    const newLookupRow = new Array(lkHdr.length).fill('');
    newLookupRow[IDX_LK.brandId]        = brandIdMachine;
    newLookupRow[IDX_LK.brandHuman]     = '';
    newLookupRow[IDX_LK.brandName]      = finalName;
    newLookupRow[IDX_LK.brandCanon]     = canon;
    newLookupRow[IDX_LK.isActive]       = true;
    newLookupRow[IDX_LK.isArchived]     = false;
    newLookupRow[IDX_LK.isStgPromoted]  = true;
    newLookupRow[IDX_LK.notes]          = 'Promoted from staging';

    lookupAppendRows.push(newLookupRow);

    const existingNote = r[IDX_STG.notes] || '';
    const newNote =
      (existingNote ? existingNote + ' | ' : '') +
      `Promoted to Lookup_Brands → Brand_ID_Machine=${brandIdMachine}`;

    stagingUpdates.push({
      row: rowNum,
      mappedId: brandIdMachine,
      note: newNote
    });

    promoted++;

    console.log(
      `[${SCRIPT_NAME}] PROMOTED row ${rowNum} → Brand_ID_Machine=${brandIdMachine}`
    );

    ETI_log_({
      executionId: EXECUTION_ID,
      scriptName: SCRIPT_NAME,
      sheetName: STG_SHEET,
      level: 'INFO',
      rowNumber: rowNum,
      action: 'PROMOTE_BRAND',
      details: `Promoted staging brand → Brand_ID_Machine=${brandIdMachine}`
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
