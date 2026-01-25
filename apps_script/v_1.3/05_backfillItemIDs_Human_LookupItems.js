/**
 * Script Name: backfillItemIDs_Human_LookupItems
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Generate sequential, human-readable Item_ID_Human values
 * - Format: ITEM-000001
 * - Fill only missing Item_ID_Human rows
 * - Preserve existing IDs
 * - Dual logging enabled (console + Script_Logs)
 *
 * Preconditions:
 * - Sheet must exist: Lookup_Items
 * - Header row present in row 1
 * - Required columns:
 *   - Item_Name
 *   - Item_ID_Human
 *
 * Algorithm (Step-by-Step):
 * 1. Read Lookup_Items into memory
 * 2. Extract max numeric suffix from existing Item_ID_Human
 * 3. Iterate rows:
 *    a. If Item_Name exists and Item_ID_Human is blank → assign next number
 *    b. Else → skip
 * 4. Write all updates in one batch
 * 5. Log execution summary
 *
 * Failure Modes:
 * - Lookup_Items sheet missing
 * - Required column missing
 *
 * Reason for Deprecation:
 * - N/A
 */

function backfillItemIDs_Human_LookupItems() {

  const EXECUTION_ID = Utilities.getUuid();
  const SCRIPT_NAME = 'backfillItemIDs_Human_LookupItems';
  const SHEET_NAME = 'Lookup_Items';
  const ID_PREFIX = 'ITEM-';
  const PAD_LENGTH = 6;

  console.log(`[${SCRIPT_NAME}] START`);
  ETI_log_({
    executionId: EXECUTION_ID,
    scriptName: SCRIPT_NAME,
    sheetName: SHEET_NAME,
    level: 'INFO',
    action: 'START',
    details: 'Execution started'
  });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`Sheet ${SHEET_NAME} not found`);

  const dataRange = sh.getDataRange();
  const data = dataRange.getValues();
  if (data.length < 2) {
    ETI_log_({
      executionId: EXECUTION_ID,
      scriptName: SCRIPT_NAME,
      sheetName: SHEET_NAME,
      level: 'WARN',
      action: 'EXIT',
      details: 'No data rows found'
    });
    return;
  }

  const header = data[0];
  const col = name => header.indexOf(name);

  const IDX = {
    itemName: col('Item_Name'),
    itemIdHuman: col('Item_ID_Human')
  };

  for (const [k, v] of Object.entries(IDX)) {
    if (v === -1) throw new Error(`Lookup_Items missing column: ${k}`);
  }

  // ---- Find current max Item_ID_Human ----
  let maxId = 0;

  for (let i = 1; i < data.length; i++) {
    const val = data[i][IDX.itemIdHuman];
    if (typeof val === 'string' && val.startsWith(ID_PREFIX)) {
      const num = parseInt(val.replace(ID_PREFIX, ''), 10);
      if (!isNaN(num)) maxId = Math.max(maxId, num);
    }
  }

  let generated = 0;
  const output = data.map(r => r.slice());

  for (let i = 1; i < output.length; i++) {
    const rowNum = i + 1;
    const itemName = output[i][IDX.itemName];
    const itemIdHuman = output[i][IDX.itemIdHuman];

    if (itemName && !itemIdHuman) {
      maxId++;
      const newHumanId =
        ID_PREFIX + String(maxId).padStart(PAD_LENGTH, '0');

      output[i][IDX.itemIdHuman] = newHumanId;
      generated++;

      console.log(
        `[${SCRIPT_NAME}] ROW ${rowNum} → GENERATED Item_ID_Human: ${newHumanId}`
      );

      ETI_log_({
        executionId: EXECUTION_ID,
        scriptName: SCRIPT_NAME,
        sheetName: SHEET_NAME,
        level: 'INFO',
        rowNumber: rowNum,
        action: 'GENERATE_ITEM_ID_HUMAN',
        details: `Generated Item_ID_Human: ${newHumanId}`
      });
    }
  }

  dataRange.setValues(output);

  console.log(`[${SCRIPT_NAME}] Item_ID_Human generated: ${generated}`);
  console.log(`[${SCRIPT_NAME}] END`);

  ETI_log_({
    executionId: EXECUTION_ID,
    scriptName: SCRIPT_NAME,
    sheetName: SHEET_NAME,
    level: 'INFO',
    action: 'SUMMARY',
    details: `Generated=${generated}, Last_Item_ID_Human=${ID_PREFIX}${String(maxId).padStart(PAD_LENGTH, '0')}`
  });

  ETI_log_({
    executionId: EXECUTION_ID,
    scriptName: SCRIPT_NAME,
    sheetName: SHEET_NAME,
    level: 'INFO',
    action: 'END',
    details: 'Execution completed successfully'
  });
}
