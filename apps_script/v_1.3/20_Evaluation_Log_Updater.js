/* =========================
   GLOBAL DEBUG CONFIG
========================= */
const DEBUG_MODE = true; 
// Set to true → evaluator row will NOT be cleared
// Set to false → normal atomic clearing


/* =========================
   ENTRY TRIGGER
========================= */
function onChangeInstallable(e) {

  console.log("=== TRIGGER FIRED ===");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const evalSheet = ss.getSheetByName("Item_Buy_Evaluate");

  if (!evalSheet) {
    console.log("ERROR: Item_Buy_Evaluate sheet not found");
    return;
  }

  const data = evalSheet.getDataRange().getValues();
  if (data.length < 2) {
    console.log("No data rows found");
    return;
  }

  const headerMap = getHeaderMap_(evalSheet);
  const row = data[1];

  const evaluationId = getCell_(row, headerMap, "Evaluation_ID");
  const inputReady = getCell_(row, headerMap, "Input_Ready_For_Comparison");

  console.log("Row state:",
    "InputReady =", inputReady,
    "Evaluation_ID =", evaluationId
  );

  if (inputReady === true && evaluationId) {
    console.log("Gating PASSED → entering processor");
    processEvaluationRow_(evalSheet, 2);
  } else {
    console.log("Gating FAILED");
  }

}


/* =========================
   CORE PROCESSOR
========================= */
function processEvaluationRow_(evalSheet, rowIndex) {

  console.log("=== PROCESS START ===");

  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // Prevent parallel execution

  try {

    const headerMap = getHeaderMap_(evalSheet);

    // 🔹 Force recalculation BEFORE reading row
    console.log("Forcing sheet recalculation...");
    SpreadsheetApp.flush();
    Utilities.sleep(250);
    SpreadsheetApp.flush();

    const row = evalSheet
      .getRange(rowIndex, 1, 1, evalSheet.getLastColumn())
      .getValues()[0];

    const get = (col) => getCell_(row, headerMap, col);

    console.log("Validation snapshot:",
      "InputReady:", get("Input_Ready_For_Comparison"),
      "Compare_ID:", get("Compare_ID"),
      "Evaluation_ID:", get("Evaluation_ID")
    );


    // ====== GATING LOGIC ======
    if (get("Input_Ready_For_Comparison") !== true) return;
    if (Number(get("Compare_ID")) !== 1) return;
    if (!get("Evaluation_ID")) return;

    // NEW: ensure full evaluation chain complete
    if (!get("Summary_UI") || get("Summary_UI").toString().trim() === "") {
      console.log("Exit: Summary_UI not ready");
      return;
    }
    // ===========================

    const logSheet = SpreadsheetApp.getActive().getSheetByName("Item_Evaluation_Log");
    if (!logSheet) {
      console.log("ERROR: Item_Evaluation_Log sheet not found");
      return;
    }

    const logHeaderMap = getHeaderMap_(logSheet);

    const lastRow = logSheet.getLastRow();
    const logData =
      lastRow > 1
        ? logSheet.getRange(2, 1, lastRow - 1, logSheet.getLastColumn()).getValues()
        : [];

    const snapshot = buildSnapshot_(headerMap, row);

    const matchIndex = findMatchRow_(snapshot, logHeaderMap, logData);

    if (matchIndex !== -1) {

      console.log("UPDATING existing log row");

      logSheet.getRange(matchIndex + 2, 1, 1, logSheet.getLastColumn())
              .setValues([buildLogRow_(snapshot, logHeaderMap)]);

    } else {

      console.log("INSERTING new log row");

      logSheet.appendRow(buildLogRow_(snapshot, logHeaderMap));
    }

    if (!DEBUG_MODE) {
      console.log("Resetting evaluator row (DEBUG_MODE = false)");
      resetEvaluatorRow_(evalSheet, rowIndex, headerMap);
    } else {

  console.log("DEBUG_MODE enabled → Skipping evaluator reset");

}

    console.log("=== PROCESS END ===");

  } finally {
    lock.releaseLock();
  }
}


/* =========================
   SNAPSHOT BUILDER
========================= */
function buildSnapshot_(headerMap, row) {

  const snapshot = {};

  for (const key in headerMap) {
    snapshot[key] = row[headerMap[key]];
  }

  // Explicit field mapping
  snapshot["Evaluated_Item"] = snapshot["Planned_Item"];
  snapshot["Evaluated_Brand"] = snapshot["Planned_Brand"];
  snapshot["Evaluated_Product"] = snapshot["Planned_Product"];
  snapshot["Evaluated_Platform"] = snapshot["Current_Platform"];
  snapshot["Evaluated_Qty"] = snapshot["Planned_Qty"];
  snapshot["Evaluated_Qty_Unit"] = snapshot["Planned_Qty_Unit"];
  snapshot["Recorded_Normalised_Qty"] = snapshot["Planned_Normalised_Qty"];
  snapshot["Evaluated_Price"] = snapshot["Current_Price"];

  snapshot["Logged_At"] = snapshot["Evaluated_At"];
  snapshot["Eval_Ready_For_Logging"] = true;

  return snapshot;
}


/* =========================
   FIND MATCH ROW
========================= */
function findMatchRow_(snapshot, logHeaderMap, logData) {

  const norm = (v) => (v || "").toString().trim().toLowerCase();
  const num = (v) => Number(v);

  for (let i = 0; i < logData.length; i++) {

    const row = logData[i];

    const match =
      norm(getCell_(row, logHeaderMap, "Evaluated_Item")) === norm(snapshot["Evaluated_Item"]) &&
      norm(getCell_(row, logHeaderMap, "Evaluated_Brand")) === norm(snapshot["Evaluated_Brand"]) &&
      norm(getCell_(row, logHeaderMap, "Evaluated_Product")) === norm(snapshot["Evaluated_Product"]) &&
      norm(getCell_(row, logHeaderMap, "Evaluated_Platform")) === norm(snapshot["Evaluated_Platform"]) &&
      num(getCell_(row, logHeaderMap, "Recorded_Normalised_Qty")) === num(snapshot["Recorded_Normalised_Qty"]) &&
      num(getCell_(row, logHeaderMap, "Evaluated_Price")) === num(snapshot["Evaluated_Price"]);

    if (match) return i;
  }

  return -1;
}


/* =========================
   BUILD LOG ROW
========================= */
function buildLogRow_(snapshot, logHeaderMap) {

  const row = new Array(Object.keys(logHeaderMap).length).fill("");

  for (const logCol in logHeaderMap) {

    const snapshotKey = Object.keys(snapshot).find(
      k => k.toLowerCase().trim() === logCol.toLowerCase().trim()
    );

    if (snapshotKey !== undefined) {
      row[logHeaderMap[logCol]] = snapshot[snapshotKey];
    }

  }

  return row;
}


/* =========================
   RESET EVALUATOR
========================= */
function resetEvaluatorRow_(sheet, rowIndex, headerMap) {

  const clearCols = [
    "Planned_Item",
    "Planned_Brand",
    "Planned_Product",
    "Current_Platform",
    "Planned_Qty",
    "Planned_Qty_Unit",
    "Current_Price",
    "Evaluation_Date",
    "Evaluated_At"
  ];

  clearCols.forEach(col => {
    const idx = headerMap[col];
    if (idx !== undefined) {
      sheet.getRange(rowIndex, idx + 1).clearContent();
    }
  });

  const evalIdx = headerMap["Evaluation_ID"];
  if (evalIdx !== undefined) {
    sheet.getRange(rowIndex, evalIdx + 1).clearContent();
  }
}


/* =========================
   SAFE HEADER MAP (TRIM + CASE SAFE)
========================= */
function getHeaderMap_(sheet) {

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};

  headers.forEach((h, i) => {
    if (!h) return;
    const clean = h.toString().trim();
    map[clean] = i;
  });

  return map;
}


/* =========================
   SAFE CELL ACCESSOR
========================= */
function getCell_(row, headerMap, colName) {

  const key = Object.keys(headerMap).find(
    k => k.toLowerCase() === colName.toLowerCase()
  );

  if (!key) return undefined;

  return row[headerMap[key]];
}