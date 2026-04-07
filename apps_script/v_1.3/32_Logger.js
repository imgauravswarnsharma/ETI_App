/**
 * Logger Name: ETI_Structured_Logger
 * Version: v1.3
 * Status: CLEAN (Singleton Context + Action Log Only)
 */


/*
-------------------------------------
GLOBAL EXECUTION CONTEXT (SINGLETON)
-------------------------------------
*/
var EXECUTION_CONTEXT = null;

function initExecutionContext_(options = {}) {

  EXECUTION_CONTEXT = {
    execution_id: Utilities.getUuid(),
    pipeline_name: options.pipeline_name || "",
    run_context: options.run_context || "STANDALONE",
    started_at: new Date()
  };

  return EXECUTION_CONTEXT;
}

function getExecutionContext_(){
  return EXECUTION_CONTEXT;
}


/*
-------------------------------------
ACTION LOGGER (DETAILED)
-------------------------------------
*/
function ETI_log_(payload) {

  // ---- Guard: Action Logging ----
  if (!isActionLogEnabled_()) return;

  const ctx = getExecutionContext_();

  // ---- Auto Inject Execution Context ----
  const executionId = ctx?.execution_id || '';
  const pipelineName = ctx?.pipeline_name || '';
  const runContext = ctx?.run_context || '';

  // ---- Debug Console ----
  if (isDebugModeEnabled_()) {
    console.log(
      `[${payload.scriptName || ''}]`,
      payload.level || 'INFO',
      payload.action || '',
      payload.rowNumber ? `ROW ${payload.rowNumber}` : '',
      payload.details || ''
    );
  }

  const logSS = getLogsSpreadsheet_();
  const sheetName = 'Action_Logs';

  let sh = logSS.getSheetByName(sheetName);

  if (!sh) {
    sh = logSS.insertSheet(sheetName);

    sh.appendRow([
      'Timestamp',

      // ---- EXECUTION TRACE ----
      'Pipeline_Name',
      'Function_Name',
      'Sheet_Name',
      'Switch_Name',

      // ---- CONTEXT ----
      'Execution_ID',
      'Trigger_Type',
      'Run_Context',

      // ---- ACTION ----
      'Level',
      'Action',
      'Step_Name',
      'Row_Number',
      'Details',

      // ---- ERROR ----
      'Error_Message'
    ]);
  }

  sh.appendRow([
    new Date(),

    // ---- EXECUTION TRACE ----
    pipelineName,
    payload.functionName || '',
    payload.sheetName || '',
    payload.switchName || '',

    // ---- CONTEXT ----
    executionId,
    payload.triggerType || 'MANUAL',
    runContext,

    // ---- ACTION ----
    payload.level || 'INFO',
    payload.action || '',
    payload.stepName || '',
    payload.rowNumber || '',
    payload.details || '',

    // ---- ERROR ----
    payload.errorMessage || ''
  ]);
}