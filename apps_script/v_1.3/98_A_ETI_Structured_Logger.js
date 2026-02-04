/**
 * Logger Name: ETI_Structured_Logger
 * Version: v1.3.2
 * Status: ACTIVE
 */

const ETI_DEBUG_MODE = true; // ← turn OFF later if needed

function ETI_log_(payload) {

  // ---- Console output for active debugging ----
  if (ETI_DEBUG_MODE) {
    console.log(
      `[${payload.scriptName}]`,
      payload.level || 'INFO',
      payload.action || '',
      payload.rowNumber ? `ROW ${payload.rowNumber}` : '',
      payload.details || ''
    );
  }

  // ---- Persistent sheet logging (audit trail) ----
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'Script_Logs';

  let sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.appendRow([
      'Timestamp',
      'Execution_ID',
      'Script_Name',
      'Sheet_Name',
      'Trigger_Type',
      'Level',
      'Row_Number',
      'Action',
      'Details'
    ]);
  }

  sh.appendRow([
    new Date(),
    payload.executionId,
    payload.scriptName,
    payload.sheetName || '',
    payload.triggerType || 'MANUAL',
    payload.level || 'INFO',
    payload.rowNumber || '',
    payload.action || '',
    payload.details || ''
  ]);
}
