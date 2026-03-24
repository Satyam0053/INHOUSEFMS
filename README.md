const CONFIG = {
  START_ROW: 7,
  COLS: {
    TIMESTAMP_A: 1, 
    SKU_B: 2,       
    DESIGN_C: 3,    
    JC_NO_D: 4,      
    QTY_C: 5,
    ARTWORK_F: 6,
    START_TS_I: 8,
    FABRIC_ACTUAL_J: 9,
    STATUS_K: 10,
    DELAY_COL: 12,
    CHECKBOX_M: 13,
    LOOKUP_O: 14,
    FABRICATOR_STATUS: 15,
    BUNDLING_PLAN_Q: 16,
    BUNDLING_ACTUAL: 17,
    BUNDLING_STATUS: 18,
    BUNDLING_DELAY: 19,
    // ✅ NEW BLANK COLUMN = 20 (T) — inserted here
    LEAD_TIME_H: 21,
    STITCH_PLAN_P: 22,
    FINAL_PACKAGING_STATUS: 57
  },
  FLOW: [
    { name: "Fabric Cutting", planned: 8,  actual: 9,  status: 10, doer: 11, delay: 12 },
    { name: "Bundling",       planned: 16, actual: 17, status: 18, doer: 11, delay: 19 },
    { name: "Stitching",      planned: 22, actual: 23, status: 24, doer: 25, delay: 26 },
    { name: "Quality Check",  planned: 27, actual: 28, status: 29, doer: 30, delay: 31 },
    { name: "Kaaj & Button",  planned: 32, actual: 33, status: 34, doer: 35, delay: 36 },
    { name: "Thread Cutting", planned: 37, actual: 38, status: 39, doer: 40, delay: 41 },
    { name: "Thread Sucking", planned: 42, actual: 43, status: 44, doer: 45, delay: 46 },
    { name: "Iron",           planned: 47, actual: 48, status: 49, doer: 50, delay: 51 },
    { name: "Packaging",      planned: 52, actual: 53, status: 54, doer: 55, delay: 56 }
  ]
};

const SHIFT_SETTINGS = {
  shiftOpen:  { h: 9,  m: 30 },
  shiftClose: { h: 20, m: 0  },
  offPattern: "1000000" // Sunday Off
};

/* ============================
    FABRICATOR EXPORT — BACKFILL EXISTING TIMESTAMPS
============================ */
function backfillFabricatorTimestamps() {
  const ss  = SpreadsheetApp.getActive();
  const sh  = ss.getSheetByName("Fabricator_Export");
  if (!sh) { ss.toast("Fabricator_Export sheet not found!", "Error"); return; }

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const data  = sh.getRange(2, 1, lastRow - 1, 6).getValues();
  const now   = new Date();
  let   count = 0;

  const output = data.map(row => {
    const colA     = row[0];
    const existing = row[5];
    if (colA !== "" && colA !== null && colA !== undefined && !existing) {
      count++;
      return [now];
    }
    return [existing || ""];
  });

  sh.getRange(2, 6, output.length, 1).setValues(output);
  ss.toast(`✅ Timestamps filled for ${count} rows`, "Backfill Complete");
}

/* ============================
    FABRICATOR EXPORT — AUTO TIMESTAMP ON CHANGE
============================ */
function onChangeFabricatorExport(e) {
  const ss  = SpreadsheetApp.getActive();
  const sh  = ss.getSheetByName("Fabricator_Export");
  if (!sh) return;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const data = sh.getRange(2, 1, lastRow - 1, 6).getValues();
  const now  = new Date();
  const colFOutput = [];
  let changed = false;

  data.forEach(row => {
    const colA     = row[0];
    const existing = row[5];
    if (colA !== "" && colA !== null && colA !== undefined && !existing) {
      colFOutput.push([now]);
      changed = true;
    } else {
      colFOutput.push([existing || ""]);
    }
  });

  if (changed) {
    sh.getRange(2, 6, colFOutput.length, 1).setValues(colFOutput);
  }
}

/* ============================
    HELPER: Valid date check
============================ */
function isValidDate_(d) {
  return d instanceof Date && !isNaN(d) && d.getFullYear() >= 2000;
}
/* ============================
    1. RUN ALL ROWS — BATCH
============================ */
function runAllRows() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("FMS");
  if (!sh) return;

  const lastRow = sh.getLastRow();
  if (lastRow < CONFIG.START_ROW) return;

  const doerCols1Based = new Set(CONFIG.FLOW.map(s => s.doer));

  const dataRange = sh.getRange(CONFIG.START_ROW, 1, lastRow - CONFIG.START_ROW + 1, 56);
  const data      = dataRange.getValues();
  const now       = new Date();

  const updates = data.map((row) => {
    const qty      = parseFloat(row[CONFIG.COLS.QTY_C - 1]) || 0;
    const colA     = row[CONFIG.COLS.TIMESTAMP_A - 1];
    const tatHours = Math.ceil(qty / 100);

    if (!qty || !colA) return row;

    if (String(row[CONFIG.COLS.FABRICATOR_STATUS - 1]).toLowerCase().trim() === "done") {
      return row;
    }

    if (!row[CONFIG.COLS.START_TS_I - 1]) {
      const startDate = colA instanceof Date ? colA : new Date(colA);
      row[CONFIG.COLS.START_TS_I - 1] = calculatePlanned_(startDate, tatHours);
    }

    CONFIG.FLOW.forEach((step, idx) => {
      const status     = String(row[step.status  - 1]).toLowerCase().trim();
      const rawActual  = row[step.actual  - 1]; // ✅ read raw before clearing
      let   plannedVal = row[step.planned - 1];
      let   actualVal  = isValidDate_(rawActual) ? rawActual : null;

      // ✅ Clear invalid planned (except protected cols)
      if (!isValidDate_(plannedVal)
          && step.planned !== CONFIG.COLS.BUNDLING_PLAN_Q
          && step.planned !== CONFIG.COLS.STITCH_PLAN_P) {
        plannedVal = null;
        row[step.planned - 1] = "";
      }

      // ✅ Only stamp actual if status is done AND actual is genuinely empty
      if (status === "done" && !actualVal) {
        actualVal = now;
        row[step.actual - 1] = actualVal;
      }

      // ✅ Clear actual if invalid (but not if we just set it)
      if (!isValidDate_(row[step.actual - 1])) {
        row[step.actual - 1] = "";
      }

      // Chain planned from previous step
      if (idx > 0
          && step.planned !== CONFIG.COLS.BUNDLING_PLAN_Q
          && step.planned !== CONFIG.COLS.STITCH_PLAN_P) {
        const prevStep   = CONFIG.FLOW[idx - 1];
        const prevActual = isValidDate_(row[prevStep.actual - 1]) ? row[prevStep.actual - 1] : null;
        const prevStatus = String(row[prevStep.status - 1]).toLowerCase().trim();
        if (prevStatus === "done" && prevActual && !plannedVal) {
          plannedVal = calculatePlanned_(prevActual, tatHours);
          row[step.planned - 1] = plannedVal;
        }
      }

      // Delay
      if (isValidDate_(plannedVal) && isValidDate_(actualVal) && actualVal > plannedVal) {
        row[step.delay - 1] = formatDuration_(actualVal - plannedVal);
      }
    });

    return row;
  });
}


/* ============================
    2. PROCESS SINGLE ROW — on edit
============================ */
function processRow_(sh, row) {
  const rowData  = sh.getRange(row, 1, 1, 56).getValues()[0];
  const qty      = parseFloat(rowData[CONFIG.COLS.QTY_C - 1]) || 0;
  const start    = rowData[CONFIG.COLS.TIMESTAMP_A - 1];
  const tatHours = Math.ceil(qty / 100);

  if (!qty || !start) return;

  if (String(rowData[CONFIG.COLS.FABRICATOR_STATUS - 1]).toLowerCase().trim() === "done") return;

  if (!rowData[CONFIG.COLS.START_TS_I - 1]) {
    sh.getRange(row, CONFIG.COLS.START_TS_I)
      .setValue(calculatePlanned_(new Date(start), tatHours));
  }

  CONFIG.FLOW.forEach((step, idx) => {
    const plannedCell = sh.getRange(row, step.planned);
    const actualCell  = sh.getRange(row, step.actual);
    const statusCell  = sh.getRange(row, step.status);
    const delayCell   = sh.getRange(row, step.delay);

    const status   = String(statusCell.getValue()).toLowerCase().trim();
    const rawActual = actualCell.getValue(); // ✅ read raw first
    let planned    = plannedCell.getValue();

    if (!isValidDate_(planned)) planned = null;

    // ✅ Only stamp actual if status is done AND cell is genuinely empty
    if (status === "done" && !isValidDate_(rawActual)) {
      actualCell.setValue(new Date());
    }

    // Chain planned from previous step
    if (idx > 0
        && step.planned !== CONFIG.COLS.BUNDLING_PLAN_Q
        && step.planned !== CONFIG.COLS.STITCH_PLAN_P
        && !planned) {
      const prev       = CONFIG.FLOW[idx - 1];
      const prevActual = sh.getRange(row, prev.actual).getValue();
      const prevStatus = String(sh.getRange(row, prev.status).getValue()).toLowerCase().trim();

      if (prevStatus === "done" && isValidDate_(prevActual)) {
        planned = calculatePlanned_(prevActual, tatHours);
        plannedCell.setValue(planned);
      }
    }

    // Delay
    const finalActual = actualCell.getValue();
    if (isValidDate_(planned) && isValidDate_(finalActual) && finalActual > planned) {
      delayCell.setValue(formatDuration_(finalActual - planned));
    }
  });

  updateFabricatorStatusRow_(sh, row);
}


/* ============================
    3. FABRICATOR STATUS — ALL ROWS
    ✅ Artwork  → checks col F = "artwork"
    ✅ Fabricator → checks col T (col 20) = "fabricator"
============================ */
function updateFabricatorStatus() {
  const ss  = SpreadsheetApp.getActive();
  const sh  = ss.getSheetByName("FMS");
  const fab = ss.getSheetByName("For_Looker_Febricator");

  if (!sh)  { ss.toast("FMS sheet not found!",             "Error"); return; }
  if (!fab) { ss.toast("For_Looker_Febricator not found!", "Error"); return; }

  const lastRow = sh.getLastRow();
  if (lastRow < CONFIG.START_ROW) return;

  // ✅ Read up to col 20 to include col T
  const fmsData = sh.getRange(
    CONFIG.START_ROW, 1,
    lastRow - CONFIG.START_ROW + 1, 20
  ).getValues();

  const fabLastRow = fab.getLastRow();
  const fabMap     = {};

  if (fabLastRow >= 2) {
    const fabData = fab.getRange(2, 1, fabLastRow - 1, 5).getValues();
    fabData.forEach(r => {
      const jc     = String(r[0]).trim();
      const status = String(r[4]).toLowerCase().trim();
      if (jc) fabMap[jc] = status;
    });
  }

  const output = fmsData.map(row => {
    const colF = String(row[CONFIG.COLS.ARTWORK_F - 1]).trim().toLowerCase(); // col 6
    const colT = String(row[19]).trim().toLowerCase();                         // col 20 (T)
    const jcNo = String(row[CONFIG.COLS.JC_NO_D  - 1]).trim();

    // ✅ Artwork — gate on col F = "artwork"
    if (colF === "artwork") {
      if (jcNo && fabMap[jcNo] === "done") return ["Done"];
      return [""];
    }

    // ✅ Fabricator — gate on col T = "fabricator"
    if (colT === "fabricator") {
      if (jcNo && fabMap[jcNo] === "done") return ["Done"];
      return [""];
    }

    // Neither condition met — blank
    return [""];
  });

  sh.getRange(
    CONFIG.START_ROW, CONFIG.COLS.FABRICATOR_STATUS,
    output.length, 1
  ).setValues(output);

  ss.toast("Fabricator Status updated!", "✅ Done");
}

/* ============================
    4. FABRICATOR STATUS — SINGLE ROW
    ✅ Artwork  → checks col F = "artwork"
    ✅ Fabricator → checks col T (col 20) = "fabricator"
============================ */
function updateFabricatorStatusRow_(sheet, row) {
  const ss  = SpreadsheetApp.getActive();
  const fab = ss.getSheetByName("For_Looker_Febricator");

  const colF = String(sheet.getRange(row, CONFIG.COLS.ARTWORK_F).getValue()).trim().toLowerCase(); // col 6
  const colT = String(sheet.getRange(row, 20).getValue()).trim().toLowerCase();                     // col T = 20

  // ✅ Neither Artwork (col F) nor Fabricator (col T) — blank and exit
  if (colF !== "artwork" && colT !== "fabricator") {
    sheet.getRange(row, CONFIG.COLS.FABRICATOR_STATUS).setValue("");
    return;
  }

  const jcNo = String(sheet.getRange(row, CONFIG.COLS.JC_NO_D).getValue()).trim();

  if (!jcNo || !fab) {
    sheet.getRange(row, CONFIG.COLS.FABRICATOR_STATUS).setValue("");
    return;
  }

  const fabLastRow = fab.getLastRow();
  if (fabLastRow < 2) {
    sheet.getRange(row, CONFIG.COLS.FABRICATOR_STATUS).setValue("");
    return;
  }

  const fabData = fab.getRange(2, 1, fabLastRow - 1, 5).getValues();

  let result = "";
  for (let i = 0; i < fabData.length; i++) {
    const fabJc     = String(fabData[i][0]).trim();
    const fabStatus = String(fabData[i][4]).toLowerCase().trim();
    if (fabJc === jcNo && fabStatus === "done") {
      result = "Done";
      break;
    }
  }

  sheet.getRange(row, CONFIG.COLS.FABRICATOR_STATUS).setValue(result);
}

/* ============================
    5. BUNDLING PLAN — COL 16
============================ */
function updateBundlingPlan() {
  const sh   = SpreadsheetApp.getActive().getSheetByName("FMS");
  const data = sh.getRange(
    CONFIG.START_ROW, 1,
    sh.getLastRow() - CONFIG.START_ROW + 1, 16
  ).getValues();

  const output = data.map(r => {
    const qty           = parseFloat(r[CONFIG.COLS.QTY_C - 1]) || 0;
    const tat           = Math.ceil(qty / 100);
    const cuttingStatus = String(r[CONFIG.COLS.STATUS_K - 1]).toLowerCase().trim(); // col 10
    const fabricActual  = r[CONFIG.COLS.FABRIC_ACTUAL_J - 1];                       // col 9

    if (cuttingStatus !== "done" || !isValidDate_(fabricActual)) return [""];

    return [calculatePlanned_(fabricActual, tat)];
  });

  sh.getRange(CONFIG.START_ROW, CONFIG.COLS.BUNDLING_PLAN_Q, output.length, 1).setValues(output);
  SpreadsheetApp.getActive().toast("Bundling Plan updated.", "✅ Done");
}

/* ============================
    6. STITCHING PLAN — COL 22
============================ */
function updateStitchPlan() {
  const sh   = SpreadsheetApp.getActive().getSheetByName("FMS");
  const data = sh.getRange(
    CONFIG.START_ROW, 1,
    sh.getLastRow() - CONFIG.START_ROW + 1, 22
  ).getValues();

  const output = data.map(r => {
    const qty        = parseFloat(r[CONFIG.COLS.QTY_C      - 1]) || 0;
    const leadDays   = parseInt(r[CONFIG.COLS.LEAD_TIME_H  - 1]);  // col 21 (U)
    const tat        = Math.ceil(qty / 100);
    const base       = r[CONFIG.COLS.BUNDLING_ACTUAL       - 1];   // col 17
    const bundlingSt = String(r[CONFIG.COLS.BUNDLING_STATUS - 1]).toLowerCase().trim(); // col 18

    // ✅ Col T (col 20) and Col U (col 21) gate
    const colT      = String(r[19]).trim().toLowerCase(); // col T = index 19
    const colU      = r[20];                              // col U = index 20
    const colUEmpty = colU === "" || colU === null || colU === undefined;

    // ✅ Only proceed if T = "inhouse" AND U is not blank
    if (colT !== "inhouse" || colUEmpty) return [""];

    // ✅ If bundling is done but lead days missing — wait
    if (bundlingSt === "done") {
      if (!leadDays || isNaN(leadDays) || leadDays <= 0) return [""];
    }

    if (isValidDate_(base)) {
      const leadHours = (isNaN(leadDays) || !leadDays) ? 0 : leadDays * 10.5;
      return [calculatePlanned_(base, leadHours + tat)];
    }

    return [""];
  });

  sh.getRange(CONFIG.START_ROW, CONFIG.COLS.STITCH_PLAN_P, output.length, 1).setValues(output);
  SpreadsheetApp.getActive().toast("Stitching Plan updated.", "✅ Done");
}

/* ============================
    7. SYNC HELPERS
============================ */
function syncAndUpdateBundling() {
  syncArtworkOnly();
  updateFabricatorStatus();
  updateBundlingPlan();
}

function syncAndUpdateStitch() {
  syncArtworkOnly();
  updateFabricatorStatus();
  updateStitchPlan();
}

/* ============================
    8. GENERATE COL H PLANNED
============================ */
function generateFabricCuttingPlanned() {
  const ss  = SpreadsheetApp.getActive();
  const sh  = ss.getSheetByName("FMS");
  if (!sh) return;

  const lastRow = sh.getLastRow();
  if (lastRow < CONFIG.START_ROW) return;

  const data   = sh.getRange(CONFIG.START_ROW, 1, lastRow - CONFIG.START_ROW + 1, 8).getValues();
  const output = [];
  let count    = 0;

  data.forEach((row) => {
    const colA     = row[CONFIG.COLS.TIMESTAMP_A - 1];
    const qty      = parseFloat(row[CONFIG.COLS.QTY_C - 1]) || 0;
    const existing = row[CONFIG.COLS.START_TS_I  - 1];
    const tatHours = Math.ceil(qty / 100);

    if (!qty || !colA) {
      output.push([existing || ""]);
      return;
    }

    const startDate = colA instanceof Date ? colA : new Date(colA);
    const planned   = calculatePlanned_(startDate, tatHours);
    output.push([planned]);
    count++;
  });

  sh.getRange(CONFIG.START_ROW, CONFIG.COLS.START_TS_I, output.length, 1).setValues(output);
  ss.toast(`Col H updated for ${count} rows`, "✅ Fabric Cutting Planned");
}

/* ============================
    9. UI MENU
============================ */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("⚙️FMS TOOLS")
    .addItem("▶️ Run Full Update",                    "runAllRows")
    .addItem("🔄 Sync Artwork (Col O)",                "syncArtworkOnly")
    .addItem("📦 Update Bundling Plan (Col 16)",       "syncAndUpdateBundling")
    .addItem("🧵 Update Stitching Plan (Col 22)",      "syncAndUpdateStitch")
    .addItem("🏭 Update Fabricator Status",            "updateFabricatorStatus")
    .addItem("📅 Generate Fabric Cutting Planned (H)", "generateFabricCuttingPlanned")
    .addItem("🕐 Backfill Fabricator Timestamps",      "backfillFabricatorTimestamps")
    .addSeparator()
    .addItem("✅ Sync All Checkboxes",                 "syncAllCheckboxes")
    .addItem("⏰ Refresh Overdue Coloring",            "refreshOverdueColoring")
    .addSeparator()
    .addItem("🖼 Refresh Borders",                     "applyBordersContiguous")
    .addToUi();

  ui.createMenu('Tasks Formulas')
    .addItem('⚙️ Initial Setup (Run First Time)', 'initialSetup')
    .addSeparator()
    .addSubMenu(ui.createMenu('📊 FMS to Looker Conversion')
      .addItem('1️⃣ Setup Looker Studio Sheets',     'createSetup')
      .addItem('2️⃣ Create Form with Prefilled URL', 'createFormWithPrefilledURL')
      .addItem('3️⃣ Convert FMS to Looker Format',   'convertFMStoDBWithDoers'))
    .addSeparator()
    .addItem('📥 Assisted ImportRange', 'importRangeFormula')
    .addSeparator()
    .addSubMenu(ui.createMenu('📐 FMS Formulas')
      .addItem('TAT with Working Hours',               'plannedwwh')
      .addItem('TAT in Days',                          'plannedindays')
      .addItem('T-x Formula',                          'plannedlead')
      .addItem('Specific Time',                        'specificTime')
      .addItem('Show Planned Only When Status is NO',  'tatifno')
      .addItem('Show Planned Only When Status is YES', 'tatifyes')
      .addItem('Set Actual Time',                      'actualTime')
      .addItem('Time Delay Formula',                   'timeDelay')
      .addItem('Vlookup Wizard',                       'vlookupFormula'))
    .addToUi();
}

/* ============================
    SETUP TRIGGER
============================ */
function setupFabricatorTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "onChangeFabricatorExport") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("onChangeFabricatorExport")
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onChange()
    .create();

  SpreadsheetApp.getActive().toast("✅ Fabricator trigger set up!", "Done");
}


/* ============================
    10. BRANDING
============================ */
function applyBranding(sheet, startRow, numRows) {
  if (!sheet || sheet.getName() !== "Form responses 2") return;
  var range = sheet.getRange(startRow, 1, numRows, 10);
  range.setFontFamily("RocknRoll One")
       .setFontSize(15)
       .setFontColor('#000000')
       .setVerticalAlignment("middle");
}

function onFormSubmit(e) {
  if (!e || !e.range) return;
  applyBranding(e.range.getSheet(), e.range.getRow(), 1);
}

function formatAllExistingData() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Form responses 2");
  if (sheet) {
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) applyBranding(sheet, 2, lastRow - 1);
  } else {
    Logger.log("Sheet 'Form responses 2' not found!");
  }
}

/* ============================
    11. UTILITIES
============================ */
function calculatePlanned_(start, hrs) {
  let cur = new Date(start);
  let rem = hrs;

  while (rem > 0) {
    if (SHIFT_SETTINGS.offPattern[cur.getDay()] === "1") {
      cur.setDate(cur.getDate() + 1);
      cur.setHours(SHIFT_SETTINGS.shiftOpen.h, SHIFT_SETTINGS.shiftOpen.m, 0, 0);
      continue;
    }

    let open  = new Date(cur);
    open.setHours(SHIFT_SETTINGS.shiftOpen.h,  SHIFT_SETTINGS.shiftOpen.m,  0, 0);
    let close = new Date(cur);
    close.setHours(SHIFT_SETTINGS.shiftClose.h, SHIFT_SETTINGS.shiftClose.m, 0, 0);

    if (cur < open) cur = open;

    if (cur >= close) {
      cur.setDate(cur.getDate() + 1);
      cur.setHours(SHIFT_SETTINGS.shiftOpen.h, SHIFT_SETTINGS.shiftOpen.m, 0, 0);
      continue;
    }

    let avail = (close - cur) / 3600000;

    if (rem <= avail) {
      cur.setTime(cur.getTime() + rem * 3600000);
      rem = 0;
    } else {
      rem -= avail;
      cur.setDate(cur.getDate() + 1);
      cur.setHours(SHIFT_SETTINGS.shiftOpen.h, SHIFT_SETTINGS.shiftOpen.m, 0, 0);
    }
  }
  return cur;
}

function formatDuration_(ms) {
  const h = Math.floor(ms / 3600000);
  const m = Math.floor((ms % 3600000) / 60000);
  return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
}

function syncAllCheckboxes() {
  const sh      = SpreadsheetApp.getActive().getSheetByName("FMS");
  const lastRow = sh.getLastRow();
  const data    = sh.getRange(
    CONFIG.START_ROW, CONFIG.COLS.ARTWORK_F,
    lastRow - CONFIG.START_ROW + 1
  ).getValues();

  data.forEach((row, i) => {
    const cellM = sh.getRange(CONFIG.START_ROW + i, CONFIG.COLS.CHECKBOX_M);
    if (String(row[0]).trim() === "Artwork") cellM.insertCheckboxes();
    else cellM.removeCheckboxes().clearContent();
  });
}
function applyBordersContiguous() {
  const sh      = SpreadsheetApp.getActive().getSheetByName("FMS");
  const lastRow = sh.getLastRow();
  if (lastRow < CONFIG.START_ROW) return;

  // ✅ Find true last row by checking col A (TIMESTAMP_A) for actual data
  const colAData = sh.getRange(CONFIG.START_ROW, CONFIG.COLS.TIMESTAMP_A, lastRow - CONFIG.START_ROW + 1, 1).getValues();
  
  let trueLastRow = CONFIG.START_ROW - 1;
  colAData.forEach((row, i) => {
    if (row[0] !== "" && row[0] !== null && row[0] !== undefined) {
      trueLastRow = CONFIG.START_ROW + i;
    }
  });

  // ✅ No data found at all — exit
  if (trueLastRow < CONFIG.START_ROW) return;

  // ✅ First clear ALL borders from start row to sheet's last row (removes stale borders on blank rows)
  sh.getRange(CONFIG.START_ROW, 1, lastRow - CONFIG.START_ROW + 1, 56)
    .setBorder(false, false, false, false, false, false);

  // ✅ Then apply borders only on rows that have data
  sh.getRange(CONFIG.START_ROW, 1, trueLastRow - CONFIG.START_ROW + 1, 56)
    .setBorder(true, true, true, true, true, true, "#6d696e", SpreadsheetApp.BorderStyle.DOUBLE);
}
function refreshOverdueColoring() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("FMS");
  if (!sh) return;

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < CONFIG.START_ROW) return;

  const data         = sh.getRange(CONFIG.START_ROW, 1, lastRow - CONFIG.START_ROW + 1, lastCol).getValues();
  const now          = new Date().getTime();
  const TWO_HOURS_MS = 2 * 60 * 60 * 1000;
  const backgrounds  = [];
  const fontColors   = [];

  data.forEach((row) => {
    const rowBg   = new Array(lastCol).fill(null);
    const rowFont = new Array(lastCol).fill(null);

    if (String(row[CONFIG.COLS.FABRICATOR_STATUS - 1]).toLowerCase().trim() === "done") {
      backgrounds.push(rowBg);
      fontColors.push(rowFont);
      return;
    }

    CONFIG.FLOW.forEach((step) => {
      const colIdx = step.planned - 1;
      if (colIdx >= lastCol) return;

      const status      = String(row[step.status - 1] || "").toLowerCase().trim();
      const plannedDate = row[colIdx];

      if (status !== "done" && isValidDate_(plannedDate)) {
        const diff = plannedDate.getTime() - now;
        if (diff <= 0) {
          rowBg[colIdx]   = "#ffff00";
          rowFont[colIdx] = "#000000";
        } else if (diff <= TWO_HOURS_MS) {
          rowBg[colIdx]   = "#ffffff";
          rowFont[colIdx] = "#ff0000";
        } else {
          rowBg[colIdx]   = "#ffffff";
          rowFont[colIdx] = "#000000";
        }
      }
    });

    backgrounds.push(rowBg);
    fontColors.push(rowFont);
  });

  const targetRange = sh.getRange(CONFIG.START_ROW, 1, backgrounds.length, lastCol);
  targetRange.setBackgrounds(backgrounds);
  targetRange.setFontColors(fontColors);
}

function lookupArtworkDateLocal_(jcNo, artData) {
  const searchId = String(jcNo).trim();
  if (!searchId) return null;
  for (let i = 1; i < artData.length; i++) {
    const rowJcNo = String(artData[i][3]).trim();
    const status  = String(artData[i][6]).toLowerCase().trim();
    if (rowJcNo === searchId && status === "done") {
      return artData[i][4] instanceof Date ? artData[i][4] : new Date(artData[i][4]);
    }
  }
  return null;
}

/* ============================
    12. ON EDIT
============================ */
/* ============================
    12. ON EDIT
============================ */
/* ============================
    2. PROCESS SINGLE ROW — on edit
    ✅ NO actual stamping — only chains planned + calculates delay
============================ */
function processRow_(sh, row) {
  const rowData  = sh.getRange(row, 1, 1, 56).getValues()[0];
  const qty      = parseFloat(rowData[CONFIG.COLS.QTY_C - 1]) || 0;
  const start    = rowData[CONFIG.COLS.TIMESTAMP_A - 1];
  const tatHours = Math.ceil(qty / 100);

  if (!qty || !start) return;
  if (String(rowData[CONFIG.COLS.FABRICATOR_STATUS - 1]).toLowerCase().trim() === "done") return;

  // ✅ Generate col H if empty
  if (!isValidDate_(rowData[CONFIG.COLS.START_TS_I - 1])) {
    sh.getRange(row, CONFIG.COLS.START_TS_I)
      .setValue(calculatePlanned_(new Date(start), tatHours));
  }

  CONFIG.FLOW.forEach((step, idx) => {
    const plannedCell = sh.getRange(row, step.planned);
    const actualCell  = sh.getRange(row, step.actual);
    const statusCell  = sh.getRange(row, step.status);
    const delayCell   = sh.getRange(row, step.delay);

    const status  = String(statusCell.getValue()).toLowerCase().trim();
    let   planned = plannedCell.getValue();
    if (!isValidDate_(planned)) planned = null;

    // ✅ Chain planned from previous step — skip protected cols
    if (idx > 0
        && step.planned !== CONFIG.COLS.BUNDLING_PLAN_Q
        && step.planned !== CONFIG.COLS.STITCH_PLAN_P
        && !planned) {
      const prev       = CONFIG.FLOW[idx - 1];
      const prevActual = sh.getRange(row, prev.actual).getValue();
      const prevStatus = String(sh.getRange(row, prev.status).getValue()).toLowerCase().trim();

      if (prevStatus === "done" && isValidDate_(prevActual)) {
        planned = calculatePlanned_(prevActual, tatHours);
        plannedCell.setValue(planned);
      }
    }

    // ✅ Delay — only calculate if both planned and actual are valid dates
    const actualVal = actualCell.getValue();
    if (isValidDate_(planned) && isValidDate_(actualVal) && actualVal > planned) {
      delayCell.setValue(formatDuration_(actualVal - planned));
    }
  });

  updateFabricatorStatusRow_(sh, row);
}

/* ============================
    12. ON EDIT
============================ */
function onEdit(e) {
  if (!e || !e.range) return;
  const sh  = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  // ✅ Fabricator_Export sheet — auto timestamp col F when col A filled
  if (sh.getName() === "Fabricator_Export") {
    if (row >= 2 && col === 1) stampFabricatorExportTimestamp_(sh, row);
    return;
  }

  if (sh.getName() !== "FMS") return;
  if (row < CONFIG.START_ROW) return;

  const val = String(e.value || "").toLowerCase().trim();

  // ✅ Checkbox logic for col F
  if (col === CONFIG.COLS.ARTWORK_F) {
    const cell = sh.getRange(row, CONFIG.COLS.CHECKBOX_M);
    if (val === "artwork") cell.insertCheckboxes();
    else cell.removeCheckboxes().clearContent();
  }

  // ✅ Auto generate col H when col A or col C edited and H is empty
  if (col === CONFIG.COLS.TIMESTAMP_A || col === CONFIG.COLS.QTY_C) {
    const rowData  = sh.getRange(row, 1, 1, 8).getValues()[0];
    const colA     = rowData[CONFIG.COLS.TIMESTAMP_A - 1];
    const qty      = parseFloat(rowData[CONFIG.COLS.QTY_C - 1]) || 0;
    const existing = rowData[CONFIG.COLS.START_TS_I  - 1];
    if (colA && qty && !isValidDate_(existing)) {
      const tatHours = Math.ceil(qty / 100);
      sh.getRange(row, CONFIG.COLS.START_TS_I)
        .setValue(calculatePlanned_(new Date(colA), tatHours));
    }
  }

  // ✅ Actual timestamp — ONLY when a STATUS column is edited to "done"
  // and the actual cell is genuinely empty — never overwrites existing
  const statusCols = new Set(CONFIG.FLOW.map(s => s.status));
  if (statusCols.has(col) && val === "done") {
    const step = CONFIG.FLOW.find(s => s.status === col);
    if (step) {
      const actualCell = sh.getRange(row, step.actual);
      if (!isValidDate_(actualCell.getValue())) {
        actualCell.setValue(new Date());
      }
    }
  }

  // ✅ Chain planned + delay only — no actual stamping inside
  processRow_(sh, row);

  // ✅ Stitch plan triggers
  if (col === CONFIG.COLS.BUNDLING_STATUS && val === "done") updateStitchPlan();
  if (col === CONFIG.COLS.LEAD_TIME_H) updateStitchPlan();
}
/* ============================
    13. SYNC ARTWORK
============================ */
function syncArtworkOnly() {
  const ss       = SpreadsheetApp.getActive();
  const sh       = ss.getSheetByName("FMS");
  const artSheet = ss.getSheetByName("Done Artwork Status");
  if (!sh || !artSheet) return;

  const lastRow = sh.getLastRow();
  const data    = sh.getRange(CONFIG.START_ROW, 1, lastRow - CONFIG.START_ROW + 1, 14).getValues();
  const artData = artSheet.getDataRange().getValues();
  const outputN = [];
  let count     = 0;

  data.forEach((row) => {
    const jcNo        = row[CONFIG.COLS.JC_NO_D  - 1];
    const statusValue = String(row[CONFIG.COLS.ARTWORK_F - 1]).trim().toLowerCase();
    let existingN     = row[CONFIG.COLS.LOOKUP_O - 1];

    const isValidStatus = statusValue === "artwork"    ||
                          statusValue === "fabricator" ||
                          statusValue === "in-house";

    if (isValidStatus) {
      const foundDate = lookupArtworkDateLocal_(jcNo, artData);
      if (foundDate) {
        existingN = foundDate;
        count++;
      }
    }
    outputN.push([existingN]);
  });

  sh.getRange(CONFIG.START_ROW, CONFIG.COLS.LOOKUP_O, outputN.length, 1).setValues(outputN);
  ss.toast(`Synced ${count} dates into Column O`, "✅ Sync Complete");
}


