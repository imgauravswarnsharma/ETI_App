/**
 * Logger Name: ETI_Structured_Logger
 * Version: v1.3
 * Status: Buffered Write + Debug Mirror
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
    trigger_type: options.trigger_type || "MANUAL",
    started_at: new Date()
  };

  // Reset buffer on new execution
  ETI_LOG_BUFFER = [];

  return EXECUTION_CONTEXT;
}

function getExecutionContext_(){
  return EXECUTION_CONTEXT;
}


/*
-------------------------------------
GLOBAL LOG BUFFER (BATCH WRITE)
-------------------------------------
*/
var ETI_LOG_BUFFER = [];


/*
-------------------------------------
SCHEMA (SINGLE SOURCE OF TRUTH)
-------------------------------------
*/
const ACTION_LOG_SCHEMA = [
  'Timestamp',
  'Level',

  'Trigger_Type',
  'Run_Context',

  'Action',
  'Debug_Message',
  'Error_Message',

  'Execution_ID',

  'Pipeline_Name',
  'Script_Name',
  'Function_Name',
  'Sheet_Name',
  'Switch_Name',

  'Step_Name',
  'Row_Number',

  'Details'
];


/*
-------------------------------------
LOG FORMATTER (CENTRALIZED)
-------------------------------------
*/

/* -------- Semantic Formatting -------- */
function formatAction_(action){
  return (action || '').replace(/_/g, ' ');
}

function formatStep_(step){
  return (step || '')
    .replace(/_/g, ' ')
    .toUpperCase();
}

/* -------- Safety Formatting -------- */
function sanitizeLogText_(text){
  if (!text) return '';
  return String(text)
    .replace(/\n/g, ' ')
    .replace(/\r/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/* -------- Orchestrator -------- */
function buildLogComponents_(payload){

  const actionDisplay = formatAction_(payload.action);

  const cleanDetails = sanitizeLogText_(payload.details);
  const cleanError = sanitizeLogText_(payload.errorMessage);

  const header = `[${payload.scriptName || ''} - ${payload.functionName || ''}${payload.sheetName ? ' - ' + payload.sheetName : ''}]`;

  const logLine = [
    header,
    actionDisplay,
    '⇒',
    payload.rowNumber ? `ROW ${payload.rowNumber}` : '',
    cleanDetails
  ].join(' ').replace(/\s+/g, ' ').trim();

  return {
    logLine,
    cleanDetails,
    cleanError
  };
}


/*
-------------------------------------
CORE LOGGER
-------------------------------------
*/
function ETI_log_(payload) {
  if (!payload) return;

  if (!payload.functionName && payload.scriptName) {
    payload.functionName = payload.scriptName;
  }

  const { logLine, cleanDetails, cleanError } = buildLogComponents_(payload);

  /*
  -------------------------------------
  DEBUG MODE (REAL-TIME)
  -------------------------------------
  */
  if (isDebugModeEnabled_()) {
    if (payload.level === 'ERROR') console.error(logLine);
    else console.log(logLine);
  }

  /*
  -------------------------------------
  ACTION LOG (BUFFERED)
  -------------------------------------
  */
  if (!isActionLogEnabled_()) return;

  const ctx = getExecutionContext_();

  ETI_LOG_BUFFER.push([
    new Date(),

    payload.level || 'INFO',

    payload.triggerType || ctx?.trigger_type || 'MANUAL',
    ctx?.run_context || 'STANDALONE',

    payload.action || '',
    logLine,
    cleanError,

    ctx?.execution_id || '',

    ctx?.pipeline_name || '',
    payload.scriptName || '',
    payload.functionName || '',
    payload.sheetName || '',
    payload.switchName || '',

    payload.stepName ? formatStep_(payload.stepName) : '',
    payload.rowNumber || '',

    cleanDetails
  ]);
}


/*
-------------------------------------
FLUSH LOGS (BATCH WRITE)
-------------------------------------
*/
function flushLogs_(){

  if (!isActionLogEnabled_()) return;
  if (!ETI_LOG_BUFFER || ETI_LOG_BUFFER.length === 0) return;

  const logSS = getLogsSpreadsheet_();
  const sheetName = 'Action_Logs';

  let sh = logSS.getSheetByName(sheetName);

  // Ensure schema
  if (!sh) {
    sh = logSS.insertSheet(sheetName);
    sh.appendRow(ACTION_LOG_SCHEMA);
  } else {
    const existingHeader = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];

    const isMismatch =
      existingHeader.length !== ACTION_LOG_SCHEMA.length ||
      existingHeader.some((h,i) => h !== ACTION_LOG_SCHEMA[i]);

    if (isMismatch) {
      sh.clear();
      sh.appendRow(ACTION_LOG_SCHEMA);

      ETI_debugLogInternal_('Logger', 'SchemaReset', 'WARN',
        'Action_Logs schema mismatch detected and reset');
    }
  }

  // Bulk write
  const startRow = sh.getLastRow() + 1;
  sh.getRange(startRow, 1, ETI_LOG_BUFFER.length, ACTION_LOG_SCHEMA.length)
    .setValues(ETI_LOG_BUFFER);

  // Clear buffer
  ETI_LOG_BUFFER = [];
}


/*
-------------------------------------
INTERNAL DEBUG (SAFE)
-------------------------------------
*/
function ETI_debugLogInternal_(script, fn, level, msg){

  if (!isDebugModeEnabled_()) return;

  const line = `[${script}::${fn}] ${level} ${msg}`;

  if (level === 'ERROR') console.error(line);
  else console.log(line);
}


/*
-------------------------------------
UTILITY: ERROR LOGGER
-------------------------------------
*/
function ETI_logError_(scriptName, functionName, error, stepName=''){

  ETI_log_({
    scriptName,
    functionName,
    level: 'ERROR',
    action: 'ERROR',
    stepName,
    details: error?.message || '',
    errorMessage: error?.stack || ''
  });
}


/*
-------------------------------------
UTILITY: STEP LOGGER
-------------------------------------
*/
function ETI_logStepStart_(scriptName, functionName, stepName){
  ETI_log_({
    scriptName,
    functionName,
    level: 'INFO',
    action: 'PROCESS',
    stepName,
    details: `${formatStep_(stepName)} started`
  });
}

function ETI_logStepEnd_(scriptName, functionName, stepName, summary=''){
  ETI_log_({
    scriptName,
    functionName,
    level: 'INFO',
    action: 'PROCESS',
    stepName,
    details: summary || `${formatStep_(stepName)} completed`
  });
}