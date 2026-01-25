/**
 * Script Name: cleanupInvalidTransactions_TransactionRaw
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Clean up invalid transaction rows after Txn_ID backfill
 * - Enforce prerequisite completeness invariant
 * - Clear Txn_ID_Machine where prerequisites are partially missing
 * - Preserve raw user-entered data (non-destructive)
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
 *    a. Check prerequisite completeness
 *    b. If partially filled AND Txn_ID_Machine exists → clear it
 *    c. If fully empty → ignore
 *    d. If fully valid → keep untouched
 * 5. Write all mutations back in a single batch
 * 6. Emit execution summary
 *
 * Failure Modes:
 * - Transaction_Raw sheet missing
 * - Required column missing
 *
 * Reason for Deprecation:
 * - N/A
 * - Remains ACTIVE for ETI v1.3
 */

function cleanupInvalidTransactions_TransactionRaw() {

  const EXECUTION_ID = Utilities.getUuid();
  const SCRIPT_NAME  = 'cleanupInvalidTransactions_TransactionRaw';
  const SHEET_NAME   = 'Transaction_Raw';

  const t0 = new Date();
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

  const range = sh.getDataRange();
  const data  = range.getValues();

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
  let cleared = 0;

  const output = data.map(r => r.slice());

  for (let i = 1; i < output.length; i++) {
    scanned++;
    const rowNum = i + 1;
    const r = output[i];

    const prereqValues = [
      r[IDX.trxDate],
      r[IDX.item],
      r[IDX.qtyVal],
      r[IDX.qtyUnit],
      r[IDX.price]
    ];

    const filledCount = prereqValues.filter(v => v !== '' && v !== null).length;

    const hasTxnId = r[IDX.txnId];

    // Partial prerequisite state + Txn_ID present → INVALID
    if (filledCount > 0 && filledCount < prereqValues.length && hasTxnId) {

      const oldId = r[IDX.txnId];
      output[i][IDX.txnId] = '';
      cleared++;

      console.log(
        `[${SCRIPT_NAME}] ROW ${rowNum} → CLEARED orphan Txn_ID_Machine: ${oldId}`
      );

      ETI_log_({
        executionId: EXECUTION_ID,
        scriptName: SCRIPT_NAME,
        sheetName: SHEET_NAME,
        level: 'WARN',
        rowNumber: rowNum,
        action: 'CLEAR_INVALID_TXN',
        details: `Cleared Txn_ID_Machine due to partial prerequisites: ${oldId}`
      });
    }
  }

  range.setValues(output);

  const durationMs = new Date().getTime() - t0.getTime();

  console.log(
    `[${SCRIPT_NAME}] Scanned=${scanned}, Cleared=${cleared}, DurationMs=${durationMs}`
  );

  ETI_log_({
    executionId: EXECUTION_ID,
    scriptName: SCRIPT_NAME,
    sheetName: SHEET_NAME,
    level: 'INFO',
    action: 'SUMMARY',
    details: `Scanned=${scanned}, Cleared=${cleared}, DurationMs=${durationMs}`
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
