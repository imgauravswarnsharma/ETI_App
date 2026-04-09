/**
 * =========================================================
 * SYSTEM CONTEXT: ETI STRUCTURED LOGGER
 * =========================================================
 *
 * POSITION IN ARCHITECTURE:
 * ---------------------------------------------------------
 * Logger = Core Infrastructure Layer
 *
 * Used by:
 * - Controller (entry point)
 * - Pipelines (execution grouping)
 * - All scripts (business logic)
 *
 * It is the SINGLE SOURCE OF TRUTH for:
 * - Debug logging
 * - Action logging (persistent audit)
 * - Execution context propagation
 *
 *
 * =========================================================
 * 1. CORE COMPONENTS
 * =========================================================
 *
 * 1. EXECUTION CONTEXT (GLOBAL STATE)
 * ---------------------------------------------------------
 * Defined via:
 *   initExecutionContext_()
 *   getExecutionContext_()
 *
 * Structure:
 * {
 *   execution_id   → unique UUID per run
 *   pipeline_name  → set at pipeline level
 *   run_context    → STANDALONE / PIPELINE
 *   trigger_type   → MANUAL / CONTROLLER
 *   started_at     → timestamp
 * }
 *
 * Behavior:
 * - Initialized ONLY at execution entry points
 *   (Controller OR manual pipeline/script)
 *
 * - Pipeline layer may MODIFY context (not recreate)
 *
 * - Logger READS context (never mutates it)
 *
 *
 * =========================================================
 * 2. LOGGING FLOW (END-TO-END)
 * =========================================================
 *
 * ENTRY POINTS:
 * ---------------------------------------------------------
 * A. Controller Execution
 *    → initExecutionContext_({ trigger_type: CONTROLLER })
 *
 * B. Manual Execution (Apps Script UI)
 *    → initExecutionContext_() OR implicit defaults
 *
 * C. Pipeline Execution
 *    → Enhances context:
 *       pipeline_name
 *       run_context = PIPELINE
 *
 *
 * LOGGING EXECUTION:
 * ---------------------------------------------------------
 * Step 1: Script calls ETI_log_(payload)
 *
 * Step 2: Logger builds:
 *   - Debug line (console)
 *   - Structured row (buffer)
 *
 * Step 3: Buffer accumulates logs
 *
 * Step 4: flushLogs_() writes to sheet (batch)
 *
 *
 * =========================================================
 * 3. LOG TYPES
 * =========================================================
 *
 * A. DEBUG LOG (REAL-TIME)
 * ---------------------------------------------------------
 * Output:
 *   Apps Script console
 *
 * Format:
 *   [Script - Function - Sheet] ACTION ⇒ Details
 *
 * Example:
 *   [Items - populate... - Staging_Lookup_Items] SUMMARY ⇒ Scanned=1999 | ...
 *
 * Purpose:
 * - Fast visual debugging
 * - Execution trace
 *
 *
 * B. ACTION LOG (PERSISTENT)
 * ---------------------------------------------------------
 * Output:
 *   Action_Logs sheet
 *
 * Stored as structured rows
 *
 * Includes:
 * - Execution metadata
 * - Debug message snapshot
 * - Error details
 *
 *
 * =========================================================
 * 4. ACTION LOG SCHEMA
 * =========================================================
 *
 * Columns:
 *
 * Timestamp
 * Level
 *
 * Trigger_Type
 * Run_Context
 *
 * Action
 * Debug_Message
 * Error_Message
 *
 * Execution_ID
 *
 * Pipeline_Name
 * Script_Name
 * Function_Name
 * Sheet_Name
 * Switch_Name
 *
 * Step_Name
 * Row_Number
 *
 * Details
 *
 *
 * Design Intent:
 * ---------------------------------------------------------
 * - Debug_Message = primary scan column
 * - Other columns = structured filtering / audit
 *
 *
 * =========================================================
 * 5. LOG FORMATTING SYSTEM
 * =========================================================
 *
 * Centralized via:
 *   buildLogComponents_()
 *
 * Sub-components:
 *
 * 1. formatAction_
 *    → Converts ACTION_NAME → "ACTION NAME"
 *
 * 2. formatStep_
 *    → Converts STEP_NAME → "STEP NAME"
 *
 * 3. sanitizeLogText_
 *    → Removes:
 *       - new lines
 *       - extra spaces
 *       - unsafe characters
 *
 * 4. Header Builder:
 *    → [Script - Function - Sheet]
 *
 *
 * RULE:
 * ---------------------------------------------------------
 * Logger controls:
 *   ✔ Structure
 *   ✔ Safety
 *
 * Script controls:
 *   ✔ Meaning
 *   ✔ Metrics
 *
 *
 * =========================================================
 * 6. BUFFERED WRITE SYSTEM
 * =========================================================
 *
 * Mechanism:
 * ---------------------------------------------------------
 * - Logs are NOT written immediately
 * - Stored in ETI_LOG_BUFFER
 * - Written in batch via flushLogs_()
 *
 * Benefits:
 * ---------------------------------------------------------
 * ✔ Performance optimized
 * ✔ Reduces API calls
 * ✔ Prevents partial writes
 *
 *
 * CRITICAL RULE:
 * ---------------------------------------------------------
 * flushLogs_() MUST be called:
 * - At pipeline end
 * - At controller wrapper end
 * - In finally blocks
 *
 *
 * =========================================================
 * 7. EXECUTION CONTEXT PROPAGATION
 * =========================================================
 *
 * FLOW:
 *
 * Controller
 *   ↓
 * initExecutionContext_
 *   ↓
 * Pipeline (enhances context)
 *   ↓
 * Script functions (read-only)
 *   ↓
 * Logger uses context
 *
 *
 * RULES:
 * ---------------------------------------------------------
 * ✔ Only ENTRY POINT initializes context
 * ✔ Lower layers MUST NOT reinitialize
 * ✔ Context can only be ENRICHED downstream
 *
 *
 * =========================================================
 * 8. ERROR HANDLING
 * =========================================================
 *
 * Utility:
 *   ETI_logError_()
 *
 * Captures:
 * - error.message → Details
 * - error.stack   → Error_Message
 *
 *
 * =========================================================
 * 9. STEP LOGGING (STANDARDIZED)
 * =========================================================
 *
 * Utilities:
 * - ETI_logStepStart_
 * - ETI_logStepEnd_
 *
 * Output:
 *   PROCESS ⇒ LOAD STAGING started
 *   PROCESS ⇒ LOAD STAGING completed
 *
 *
 * =========================================================
 * 10. DESIGN PRINCIPLES
 * =========================================================
 *
 * ✔ Single Source of Truth (Logger only)
 * ✔ No logging logic in business scripts
 * ✔ Context-driven logging
 * ✔ Plug-and-play across system
 * ✔ Performance-first (batch writes)
 * ✔ Readable debug-first design
 *
 *
 * =========================================================
 * 11. CONTROLLER + PIPELINE + LOGGER INTERPLAY
 * =========================================================
 *
 * Controller:
 *   → Defines trigger_type
 *   → Starts execution
 *
 * Pipeline:
 *   → Defines run_context
 *   → Defines pipeline_name
 *
 * Script:
 *   → Emits logs (no context logic)
 *
 * Logger:
 *   → Formats + persists logs
 *
 *
 * =========================================================
 * 12. FINAL ARCHITECTURE SUMMARY
 * =========================================================
 *
 * Controller  → WHEN to run
 * Pipeline    → WHAT group to run
 * Script      → HOW logic runs
 * Logger      → WHAT happened (audit + debug)
 *
 * =========================================================
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

  /*
  -------------------------------------
  EXECUTION CONTEXT FALLBACK (ADDED)
  -------------------------------------
  */
  let ctx = getExecutionContext_();
  if (!ctx) {
    initExecutionContext_({
      run_context: 'STANDALONE',
      trigger_type: 'MANUAL'
    });
    ctx = getExecutionContext_();
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

/*
-------------------------------------
BLANK ROW INJECTION (FIXED)
-------------------------------------
*/
let startRow = sh.getLastRow() + 1;

if (sh.getLastRow() > 1) {
  sh.insertRowBefore(startRow);

  // APPLY COLOR TO BLANK ROW
  sh.getRange(startRow, 1, 1, ACTION_LOG_SCHEMA.length)
    .setBackground('#fbbc04');

  startRow++; // CRITICAL FIX
}

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