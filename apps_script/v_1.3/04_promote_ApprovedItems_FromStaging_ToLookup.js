/**
 * Script Name: promoteApprovedItems_FromStaging_ToLookup
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Version: v1.3.1
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Promote staging items into Lookup_Items once governance conditions are satisfied
 * - Maintain deterministic promotion authority
 * - Preserve lineage via Staging_Item_ID_Machine
 * - Provide full action-level logging for debugging
 *
 * Promotion Gate (ALL must be true):
 *
 *   Action_Review_Status = "Pending (Promotion)"
 *   Is_Pipeline_ready    = TRUE
 *   Is_Lookup_Promoted   = FALSE
 *
 * Promotion Results:
 *
 * Lookup_Items row inserted
 * Staging row updated with:
 *
 *   Mapped_Item_ID_Machine
 *   Is_Lookup_Promoted
 *   Action_Review_Status
 *   Entity_Owner
 *   Promotion_Label
 *   Promoted_At
 *   Item_Status
 *   Notes
 */

function promoteApprovedItems_FromStaging_ToLookup() {

  const EXECUTION_ID = Utilities.getUuid();
  const SCRIPT_NAME  = 'promoteApprovedItems_FromStaging_ToLookup';
  const STG_SHEET    = 'Staging_Lookup_Items';
  const LKP_SHEET    = 'Lookup_Items';

  const t0 = new Date();

  try {

    console.log(`[${SCRIPT_NAME}] START`);

    ETI_log_({
      executionId: EXECUTION_ID,
      scriptName: SCRIPT_NAME,
      sheetName: STG_SHEET,
      level: 'INFO',
      action: 'START',
      details: 'Promotion execution started'
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
      itemName: lkCol('Item_Name'),
      itemCanon: lkCol('Item_Name_Canonical'),
      isApproved: lkCol('Is_Approved'),
      isActive: lkCol('Is_Active'),
      isArchived: lkCol('Is_Archived'),
      isStgPromoted: lkCol('Is_Staging_Promoted'),
      sourceType: lkCol('Source_Type'),
      createdAt: lkCol('Created_At'),
      notes: lkCol('Notes'),
      itemIdMachine: lkCol('Item_ID_Machine'),
      stagingId: lkCol('Staging_ID_Machine')
    };

    for (const [k,v] of Object.entries(IDX_LK)) {
      if (v === -1) throw new Error(`Lookup_Items missing column: ${k}`);
    }

    /* ===================== READ STAGING ===================== */

    const stgData = stgSh.getDataRange().getValues();
    const stgHdr  = stgData[0];
    const stgCol  = n => stgHdr.indexOf(n);

    const IDX_STG = {
      entered: stgCol('Item_Name_Entered'),
      approved: stgCol('Item_Name_Approved'),
      canon: stgCol('Item_Name_Canonical'),
      reviewStatus: stgCol('Action_Review_Status'),
      pipelineReady: stgCol('Is_Pipeline_ready'),
      isPromoted: stgCol('Is_Lookup_Promoted'),
      stagingId: stgCol('Staging_Item_ID_Machine'),
      mappedId: stgCol('Mapped_Item_ID_Machine'),
      notes: stgCol('Notes'),
      isApproved: stgCol('Is_Approved'),
      isActive: stgCol('Is_Active'),
      isArchived: stgCol('Is_Archived'),
      entityOwner: stgCol('Entity_Owner'),
      promotionLabel: stgCol('Promotion_Label'),
      promotedAt: stgCol('Promoted_At'),
      itemStatus: stgCol('Item_Status')
    };

    for (const [k,v] of Object.entries(IDX_STG)) {
      if (v === -1) throw new Error(`Staging_Lookup_Items missing column: ${k}`);
    }

    /* ===================== PROMOTION LOOP ===================== */

    const lookupAppendRows = [];
    const stagingUpdates   = [];

    let scanned = 0;
    let promoted = 0;
    let skipped = 0;

    for (let i = 1; i < stgData.length; i++) {

      scanned++;

      const rowNum = i + 1;
      const r = stgData[i];

      if (r[IDX_STG.reviewStatus] !== 'Pending (Promotion)') { skipped++; continue; }
      if (r[IDX_STG.pipelineReady] !== true) { skipped++; continue; }
      if (r[IDX_STG.isPromoted] === true) { skipped++; continue; }

      const finalName =
        r[IDX_STG.approved] ||
        r[IDX_STG.entered];

      if (!finalName) { skipped++; continue; }

      const canon = r[IDX_STG.canon] || '';
      const stagingId = r[IDX_STG.stagingId];

      const itemIdMachine = Utilities.getUuid();

      const newLookupRow = new Array(lkHdr.length).fill('');

      newLookupRow[IDX_LK.itemName] = finalName;
      newLookupRow[IDX_LK.itemCanon] = canon;
      newLookupRow[IDX_LK.isApproved] = r[IDX_STG.isApproved];
      newLookupRow[IDX_LK.isActive]   = r[IDX_STG.isActive];
      newLookupRow[IDX_LK.isArchived] = r[IDX_STG.isArchived];

      newLookupRow[IDX_LK.isStgPromoted] = true;
      newLookupRow[IDX_LK.sourceType] = 'STAGING_PROMOTION';
      newLookupRow[IDX_LK.createdAt] = new Date();

      newLookupRow[IDX_LK.itemIdMachine] = itemIdMachine;
      newLookupRow[IDX_LK.stagingId] = stagingId;

      newLookupRow[IDX_LK.notes] =
        `Promoted from staging → Staging_ID=${stagingId}`;

      lookupAppendRows.push(newLookupRow);

      /* derive promoted Item_Status */

      let promotedStatus = '';

      if (r[IDX_STG.isApproved] && r[IDX_STG.isArchived])
        promotedStatus = 'Promoted (Archived)';

      else if (r[IDX_STG.isApproved] && r[IDX_STG.isActive])
        promotedStatus = 'Promoted (Live)';

      else if (r[IDX_STG.isApproved])
        promotedStatus = 'Promoted (Hidden Dropdown)';

      const existingNote = r[IDX_STG.notes] || '';

      const newNote =
        (existingNote ? existingNote + ' | ' : '') +
        `Promoted to Lookup_Items → Item_ID_Machine=${itemIdMachine}`;

      stagingUpdates.push({
        row: rowNum,
        mappedId: itemIdMachine,
        note: newNote,
        status: promotedStatus
      });

      /* Action Log */

      ETI_log_({
        executionId: EXECUTION_ID,
        scriptName: SCRIPT_NAME,
        sheetName: STG_SHEET,
        level: 'INFO',
        action: 'PROMOTION',
        details:
          `Row=${rowNum}, Staging_ID=${stagingId}, Item_ID=${itemIdMachine}, Item_Name=${finalName}`
      });

      promoted++;
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

    /* ===================== WRITE BACK STAGING ===================== */

    for (const u of stagingUpdates) {

      stgSh.getRange(u.row, IDX_STG.mappedId + 1).setValue(u.mappedId);
      stgSh.getRange(u.row, IDX_STG.isPromoted + 1).setValue(true);
      stgSh.getRange(u.row, IDX_STG.reviewStatus + 1).setValue('Promoted');
      stgSh.getRange(u.row, IDX_STG.entityOwner + 1).setValue('Lookup');
      stgSh.getRange(u.row, IDX_STG.promotionLabel + 1).setValue('Promoted');
      stgSh.getRange(u.row, IDX_STG.promotedAt + 1).setValue(new Date());
      stgSh.getRange(u.row, IDX_STG.itemStatus + 1).setValue(u.status);
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
      details: 'Promotion execution completed'
    });

  } catch (err) {

    ETI_log_({
      executionId: EXECUTION_ID,
      scriptName: SCRIPT_NAME,
      sheetName: STG_SHEET,
      level: 'ERROR',
      action: 'FAILED',
      details: err.message
    });

    throw err;
  }
}