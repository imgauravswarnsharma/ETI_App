/**
 * Script Name: backfillTxnIDs_TransactionRaw
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Backfill Txn_ID_Machine for valid transaction rows in Transaction_Raw
 * - Preserve idempotency and non-blocking behavior
 * - Emit both console logs (debug) and structured sheet logs (audit)
 *
 * Preconditions:
 * - Sheet must exist: Transaction_Raw
 * - Header row present in row 1
 * - Required columns:
 *   - Trx_Date_Entered
 *   - Item_Name_Entered
 *   - Qty_Value_Entered
 *   - Qty_Unit_Entered
 *   - Price_Entered
 *   - Txn_ID_Machine
 *
 * Algorithm (Step-by-Step):
 * 1. Generate Execution_ID
 * 2. Load Transaction_Raw into memory
 * 3. Resolve column indexes from header
 * 4. Iterate rows:
 *    a. Skip invalid transactions
 *    b. Skip rows with existing Txn_ID_Machine
 *    c. Generate and write Txn_ID_Machine for valid rows
 * 5. Emit execution summary
 *
 * Failure Modes:
 * - Transaction_Raw sheet missing
 * - Required column missing
 *
 * Reason for Deprecation:
 * - N/A
 */

function backfillTxnIDs_TransactionRaw() {

  const EXECUTION_ID = Utilities.getUuid();
  const SCRIPT_NAME = 'backfillTxnIDs_TransactionRaw';
  const SHEET_NAME = 'Transaction_Raw';

  const t0 = new Date();

  // ---- START LOG ----
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

  const data = sh.getDataRange().getValues();
  if (data.length < 2) {
    console.log(`[${SCRIPT_NAME}] No data rows found`);
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
    trxDate: col('Trx_Date_Entered'),
    item: col('Item_Name_Entered'),
    qtyVal: col('Qty_Value_Entered'),
    qtyUnit: col('Qty_Unit_Entered'),
    price: col('Price_Entered'),
    txnId: col('Txn_ID_Machine')
  };

  for (const [k, v] of Object.entries(IDX)) {
    if (v === -1) {
      throw new Error(`Transaction_Raw missing column: ${k}`);
    }
  }

  let scanned = 0;
  let skipInvalid = 0;
  let skipHasTxnId = 0;
  let generated = 0;

  for (let i = 1; i < data.length; i++) {
    scanned++;
    const rowNum = i + 1;
    const r = data[i];

    const isValidTxn =
      r[IDX.trxDate] &&
      r[IDX.item] &&
      r[IDX.qtyVal] &&
      r[IDX.qtyUnit] &&
      r[IDX.price];

    if (!isValidTxn) {
      skipInvalid++;
      continue;
    }

    if (r[IDX.txnId]) {
      skipHasTxnId++;
      continue;
    }

    const machineId =
      Utilities.getUuid();

    sh.getRange(rowNum, IDX.txnId + 1).setValue(machineId);
    generated++;

    console.log(
      `[${SCRIPT_NAME}] ROW ${rowNum} → GENERATED Txn_ID_Machine: ${machineId}`
    );

    ETI_log_({
      executionId: EXECUTION_ID,
      scriptName: SCRIPT_NAME,
      sheetName: SHEET_NAME,
      level: 'INFO',
      rowNumber: rowNum,
      action: 'GENERATE_TXN_ID',
      details: `Generated Txn_ID_Machine: ${machineId}`
    });
  }

  const durationMs = new Date().getTime() - t0.getTime();

  // ---- SUMMARY LOG ----
  console.log(`[${SCRIPT_NAME}] Rows scanned: ${scanned}`);
  console.log(`[${SCRIPT_NAME}] Skipped invalid: ${skipInvalid}`);
  console.log(`[${SCRIPT_NAME}] Skipped existing Txn_ID: ${skipHasTxnId}`);
  console.log(`[${SCRIPT_NAME}] Generated Txn_ID: ${generated}`);
  console.log(`[${SCRIPT_NAME}] END – Duration(ms): ${durationMs}`);

  ETI_log_({
    executionId: EXECUTION_ID,
    scriptName: SCRIPT_NAME,
    sheetName: SHEET_NAME,
    level: 'INFO',
    action: 'SUMMARY',
    details: `Scanned=${scanned}, Invalid=${skipInvalid}, Existing=${skipHasTxnId}, Generated=${generated}, DurationMs=${durationMs}`
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
