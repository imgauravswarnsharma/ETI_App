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
  const logSS = getLogsSpreadsheet_();
  const sheetName = 'Script_Logs';

  let sh = logSS.getSheetByName(sheetName);
  if (!sh) {
    sh = logSS.insertSheet(sheetName);
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

function logToSheet_(message, level) {

  const logSS = getLogsSpreadsheet_();
  const shName = 'Script_Logs';

  let sh = logSS.getSheetByName(shName);

  if (!sh) {
    sh = logSS.insertSheet(shName);
    sh.appendRow(['Timestamp', 'Level', 'Message']);
  }

  sh.appendRow([
    new Date(),
    level || 'INFO',
    message
  ]);
}
