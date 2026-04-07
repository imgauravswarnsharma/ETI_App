/**
 * Script Name: processStagingItems_StateMachine
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Deterministically process governance state for Staging_Lookup_Items.
 * - Repair any drift between Admin_Action and binary state flags.
 * - Derive all dependent columns directly in script (no formula reliance).
 * - Preserve audit trace of drift in Notes and ETI_log_.
 *
 * Governance State Machine (Authoritative)
 *
 * Admin_Action → Binary Flags → Derived State
 *
 * Admin_Action                Is_Approved   Is_Active   Is_Archived
 * ------------------------------------------------------------------
 * Review                      FALSE         FALSE       FALSE
 * Activate                    FALSE         TRUE        FALSE
 * Approve (UI Hidden)         TRUE          FALSE       FALSE
 * Approve & Activate          TRUE          TRUE        FALSE
 * Approve but Deprecate       TRUE          FALSE       TRUE
 * Reject                      FALSE         FALSE       TRUE
 *
 *
 * Post Promotion (handled by promotion script)
 *
 * Is_Approved   Is_Active   Is_Archived   Is_Lookup_Promoted
 * -----------------------------------------------------------
 * TRUE          FALSE       TRUE          TRUE   → Promoted (Archived)
 * TRUE          FALSE       FALSE         TRUE   → Promoted (Hidden Dropdown)
 * TRUE          TRUE        FALSE         TRUE   → Promoted (Live)
 *
 *
 * Derived Columns
 *
 * Valid_State =
 * NOT(
 *   (Is_Active = TRUE AND Is_Archived = TRUE)
 *   OR
 *   (Is_Lookup_Promoted = TRUE AND Is_Approved = FALSE)
 * )
 *
 *
 * Is_Pipeline_ready =
 * Is_Approved
 * AND NOT Is_Lookup_Promoted
 * AND Valid_State
 *
 *
 * Action_Review_Status
 *
 * if Is_Lookup_Promoted → Promoted
 * else if Is_Approved → Pending (Promotion)
 * else if Is_Archived → Rejected
 * else → Pending (Approval)
 *
 *
 * Item_Status
 *
 * If promoted:
 *   Approved + Active → Promoted (Live)
 *   Approved + Archived → Promoted (Archived)
 *   Approved only → Promoted (Hidden Dropdown)
 *
 * If not promoted:
 *   none → To be Reviewed
 *   active only → Active (Temporary)
 *   approved only → Approved (Hidden Dropdown)
 *   approved + active → Approved & Activated (Temporary)
 *   approved + archived → Approved (Archived)
 *   archived only → Rejected
 *
 *
 * Entity_Owner
 *
 * If Is_Lookup_Promoted → Lookup
 * Else → Staging
 *
 *
 * Integrity_Status
 *
 * VALID
 * REPAIRED
 * INVALID_STATE
 *
 *
 * Preconditions
 * - Sheet must exist: Staging_Lookup_Items
 * - Sheet must exist: Automation_Control
 * - Integrity control switch: Automation_Control!K2
 *
 * Failure Modes
 * - Required sheet missing
 * - Required column missing
 */

function processStagingItems_StateMachine() {

  const EXECUTION_ID = Utilities.getUuid();
  const SCRIPT_NAME = 'processStagingItems_StateMachine';
  const STG_SHEET = 'Staging_Lookup_Items';
  const CTRL_SHEET = 'Automation_Control';

  const t0 = new Date();

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  /* =========================
     Integrity Switch Check
     ========================= */

  const ctrl = ss.getSheetByName(CTRL_SHEET);
  if (!ctrl) throw new Error('Automation_Control sheet missing');

  const runIntegrity = ctrl.getRange("K2").getValue();

  if (runIntegrity !== true) {
    console.log(`[${SCRIPT_NAME}] skipped via control switch`);
    return;
  }

  console.log(`[${SCRIPT_NAME}] START`);

  const stgSh = ss.getSheetByName(STG_SHEET);
  if (!stgSh) throw new Error('Staging_Lookup_Items sheet missing');

  const data = stgSh.getDataRange().getValues();
  const hdr = data[0];

  const col = n => hdr.indexOf(n);

  const IDX = {
    adminAction: col('Admin_Action'),
    isApproved: col('Is_Approved'),
    isActive: col('Is_Active'),
    isArchived: col('Is_Archived'),
    isPromoted: col('Is_Lookup_Promoted'),
    pipelineReady: col('Is_Pipeline_ready'),
    validState: col('Valid_State'),
    actionStatus: col('Action_Review_Status'),
    itemStatus: col('Item_Status'),
    entityOwner: col('Entity_Owner'),
    integrity: col('Integrity_Status'),
    notes: col('Notes'),
    stagingId: col('Staging_Item_ID_Machine')
  };

  for (const [k,v] of Object.entries(IDX)) {
    if (v === -1) throw new Error(`Missing column: ${k}`);
  }

  let repaired = 0;
  let valid = 0;
  let invalid = 0;

  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "EEEE, MMMM d, yyyy 'at' HH:mm:ss"
  );

  /* =========================
     Row Processing Loop
     ========================= */

  for (let i = 1; i < data.length; i++) {

    const row = data[i];
    const admin = row[IDX.adminAction];
    const stagingId = row[IDX.stagingId];

    if (!admin) continue;

    let expected = {
      approved:false,
      active:false,
      archived:false
    };

    switch(admin) {

      case 'Review':
        break;

      case 'Activate':
        expected.active = true;
        break;

      case 'Approve (UI Hidden)':
        expected.approved = true;
        break;

      case 'Approve & Activate':
        expected.approved = true;
        expected.active = true;
        break;

      case 'Approve but Deprecate':
        expected.approved = true;
        expected.archived = true;
        break;

      case 'Reject':
        expected.archived = true;
        break;

      default:
        invalid++;
        row[IDX.integrity] = 'INVALID_ADMIN_ACTION';
        continue;
    }

    let drift = [];

    function repair(idx, expectedVal, name) {

      const actual = row[idx];

      if (actual !== expectedVal) {
        drift.push(`${name} expected=${expectedVal} found=${actual}`);
        row[idx] = expectedVal;
      }
    }

    repair(IDX.isApproved, expected.approved, 'Is_Approved');
    repair(IDX.isActive, expected.active, 'Is_Active');
    repair(IDX.isArchived, expected.archived, 'Is_Archived');

    const promoted = row[IDX.isPromoted];

    const validState =
      !(row[IDX.isActive] && row[IDX.isArchived]) &&
      !(promoted && !row[IDX.isApproved]);

    row[IDX.validState] = validState;

    if (!validState) {
      row[IDX.integrity] = 'INVALID_STATE';
      invalid++;
      continue;
    }

    const pipelineReady =
      row[IDX.isApproved] &&
      !promoted &&
      validState;

    row[IDX.pipelineReady] = pipelineReady;

    let reviewStatus = 'Pending (Approval)';

    if (promoted) reviewStatus = 'Promoted';
    else if (row[IDX.isApproved]) reviewStatus = 'Pending (Promotion)';
    else if (row[IDX.isArchived]) reviewStatus = 'Rejected';

    row[IDX.actionStatus] = reviewStatus;

    let itemStatus = 'To be Reviewed';

    if (promoted) {

      if (row[IDX.isActive])
        itemStatus = 'Promoted (Live)';
      else if (row[IDX.isArchived])
        itemStatus = 'Promoted (Archived)';
      else
        itemStatus = 'Promoted (Hidden Dropdown)';

    } else {

      if (row[IDX.isArchived] && !row[IDX.isApproved])
        itemStatus = 'Rejected';

      else if (row[IDX.isActive] && !row[IDX.isApproved])
        itemStatus = 'Active (Temporary)';

      else if (row[IDX.isApproved] && !row[IDX.isActive])
        itemStatus = 'Approved (Hidden Dropdown)';

      else if (row[IDX.isApproved] && row[IDX.isActive])
        itemStatus = 'Approved & Activated (Temporary)';

      else if (row[IDX.isApproved] && row[IDX.isArchived])
        itemStatus = 'Approved (Archived)';
    }

    row[IDX.itemStatus] = itemStatus;

    row[IDX.entityOwner] =
      promoted ? 'Lookup' : 'Staging';

    if (drift.length > 0) {

      repaired++;

      const msg =
        `Integrity drift repaired: ${drift.join(' | ')} — ${timestamp}`;

      row[IDX.notes] = msg;
      row[IDX.integrity] = 'REPAIRED';

      ETI_log_({
        executionId: EXECUTION_ID,
        scriptName: SCRIPT_NAME,
        sheetName: STG_SHEET,
        level: 'WARN',
        action: 'DRIFT_REPAIR',
        details: `Row=${i+1}, Staging_ID=${stagingId}, ${drift.join(' | ')}`
      });

    } else {

      valid++;

      row[IDX.integrity] = 'VALID';
      row[IDX.notes] =
        `Integrity check passed — ${timestamp}`;
    }
  }

  /* =========================
     Batch Write Back
     ========================= */

  stgSh
    .getRange(2,1,data.length-1,hdr.length)
    .setValues(data.slice(1));

  const durationMs = new Date().getTime() - t0.getTime();

  console.log(
    `[${SCRIPT_NAME}] VALID=${valid} REPAIRED=${repaired} INVALID=${invalid} Duration=${durationMs}`
  );

  ETI_log_({
    executionId: EXECUTION_ID,
    scriptName: SCRIPT_NAME,
    sheetName: STG_SHEET,
    level: 'INFO',
    action: 'SUMMARY',
    details: `Valid=${valid}, Repaired=${repaired}, Invalid=${invalid}, DurationMs=${durationMs}`
  });

  ETI_log_({
    executionId: EXECUTION_ID,
    scriptName: SCRIPT_NAME,
    sheetName: STG_SHEET,
    level: 'INFO',
    action: 'END',
    details: 'State machine processing completed'
  });

}