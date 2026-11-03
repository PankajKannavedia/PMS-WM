// ================= MENU =================
function onOpen() {
  SpreadsheetApp.getUi().createMenu("Billing FMS")
    .addItem("Billing FMS", "importBillingFMS")
    .addItem("WBS Sheet", "importWBSSheet")
    .addItem("INVENTORY", "importInventory")
    .addItem("Manpower Cost", "importManpowerCost")
    .addItem("PROJECT BOQ", "importProjectBOQ")
    .addItem("PROJECT PLANNING", "importProjectPlanning")
    .addItem("Auditor Sheet Manpower", "importAuditorManpower")
    .addItem("Auditor Sheet Item", "importAuditorItem")
    .addItem("Auditor Sheet Drawing", "importAuditorDrawing")
    .addSeparator()
    .addItem("Fetch ALL", "fetchAllBillingFMSData")
    .addSeparator()
    .addItem("⚙ Setup Auto Refresh (every 10 min)", "installAutoRefresh")
    .addItem("⏹ Stop Auto Refresh", "removeAutoRefresh")
    .addItem("▶ Run Auto Refresh now", "runScheduler")
    .addItem("↻ Rebuild Log Summary", "rebuildLogSummary")
    .addToUi();
}

// ============ MANUAL / MENU FUNCTIONS ============
function importBillingFMS(e)       { runImport_(["Billing FMS"],            "importBillingFMS", e); }
function importWBSSheet(e)         { runImport_(["WBS Sheet"],              "importWBSSheet", e); }
function importInventory(e)        { runImport_(["INVENTORY"],              "importInventory", e); }
function importManpowerCost(e)     { runImport_(["Manpower Cost"],          "importManpowerCost", e); }
function importProjectBOQ(e)       { runImport_(["PROJECT BOQ"],            "importProjectBOQ", e); }
function importProjectPlanning(e)  { runImport_(["PROJECT PLANNING"],       "importProjectPlanning", e); }
function importAuditorManpower(e)  { runImport_(["Auditor Sheet Manpower"], "importAuditorManpower", e); }
function importAuditorItem(e)      { runImport_(["Auditor Sheet Item"],     "importAuditorItem", e); }
function importAuditorDrawing(e)   { runImport_(["Auditor Sheet Drawing"],  "importAuditorDrawing", e); }
function fetchAllBillingFMSData(e) { runImport_(null,                       "fetchAllBillingFMSData", e); }


// ============ SETTINGS ============
// Destinations shown in the log summary (must match Index column E / tab names)
const DESTINATIONS = [
  "Billing FMS", "WBS Sheet", "INVENTORY", "Manpower Cost", "PROJECT BOQ",
  "PROJECT PLANNING", "Auditor Sheet Manpower", "Auditor Sheet Item", "Auditor Sheet Drawing"
];

// Key column per destination (source sheet column letter):
// a row is copied only if this column is NOT blank and NOT 0. Use lowercase names.
const KEY_COLUMN = {
  "project boq": "A",
  "project planning": "H",
  "inventory": "W"
};

const CONFIG = {
  INDEX_SHEET: "Index",
  LOG_SHEET: "TriggerLogs",
  LOG_MAX_ROWS: 5000,
  START_ROW: 2,
  STATUS_COL: 7,                 // Index column G = live status
  READ_PARALLEL: 3,
  WRITE_PARALLEL: 4,
  WRITE_CHUNK_CELLS: 150000,
  MAX_RETRIES: 3,
  RUN_BUDGET_MS: 5 * 60 * 1000,  // stay under Apps Script 6-min limit
  AUTO_REFRESH_MINUTES: 10       // allowed: 1, 5, 10, 15, 30
};

const LOG_HEADERS = [
  "Timestamp", "Run By", "Destination", "Run Type", "Status", "Sources OK",
  "Sources Failed", "Rows Loaded", "Read (s)", "Write (s)", "Total (s)", "Error / Note"
];


// ============ AUTO REFRESH (single trigger) ============
function installAutoRefresh() {
  removeAutoRefresh(true);
  ScriptApp.newTrigger("runScheduler").timeBased().everyMinutes(CONFIG.AUTO_REFRESH_MINUTES).create();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Auto Refresh set: every " + CONFIG.AUTO_REFRESH_MINUTES + " min. Old triggers removed.", "Auto Refresh", 10);
}

function removeAutoRefresh(silent) {
  const names = ["runScheduler", "importBillingFMS", "importWBSSheet", "importInventory",
    "importManpowerCost", "importProjectBOQ", "importProjectPlanning", "importAuditorManpower",
    "importAuditorItem", "importAuditorDrawing", "fetchAllBillingFMSData"];
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (names.indexOf(t.getHandlerFunction()) > -1) { ScriptApp.deleteTrigger(t); removed++; }
  });
  if (silent !== true)
    SpreadsheetApp.getActiveSpreadsheet().toast(removed + " trigger(s) removed.", "Auto Refresh", 8);
}

// Refreshes the most out-of-date destinations first, as many as fit in one run
function runScheduler(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return; // a manual run is active - next cycle will continue

  const t0 = Date.now();
  const results = [];
  try {
    const ctx = prepare_(ss);
    const order = ctx.destOrder.slice().sort((a, b) => (getLast_(a).lastRun || 0) - (getLast_(b).lastRun || 0));

    for (let d = 0; d < order.length; d++) {
      const name = order[d];
      const est = (getLast_(name).dur || 60) * 1000;
      const elapsed = Date.now() - t0;
      if (results.length && elapsed + est > CONFIG.RUN_BUDGET_MS) break; // rest next cycle
      const tag = "Auto: " + name;
      results.push(processDestination_(ctx, name,
        (f, text) => showProgress_(ss, tag, f * 100, text)));
    }
  } catch (err) {
    results.push({ destination: "(setup)", status: "FAILED", errors: [{ msg: err.message }] });
  } finally {
    lock.releaseLock();
    writeLogs_(ss, "AutoRefresh", (e && e.triggerUid) ? "Time trigger" : "Manual", results);
  }
}


// ============ MANUAL RUN ============
function runImport_(filter, fnName, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const runType = (e && e.triggerUid) ? "Time trigger" : "Manual";
  const label = filter ? filter.join(", ") : "All";
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(runType === "Manual" ? 30000 : 10000)) {
    writeLogs_(ss, fnName, runType, [{ destination: label, status: "SKIPPED",
      errors: [{ msg: "Another import was running - try again in a minute" }] }]);
    ss.toast("Another import is running. Please try again in a minute.", label, 10);
    return;
  }

  const t0 = Date.now();
  const results = [];
  try {
    const ctx = prepare_(ss);
    const names = filter ? filter.map(f => ctx.lookup[f.trim().toLowerCase()] || f) : ctx.destOrder.slice();

    names.forEach((name, d) => {
      if (!ctx.groups[name]) {
        results.push({ destination: name, status: "FAILED",
          errors: [{ msg: "No rows in Index for this destination (check column E)" }] });
        return;
      }
      if (results.length && Date.now() - t0 > CONFIG.RUN_BUDGET_MS) {
        results.push({ destination: name, status: "SKIPPED",
          errors: [{ msg: "Time limit reached - run it separately" }] });
        return;
      }
      const tag = name + " (" + (d + 1) + "/" + names.length + ")";
      results.push(processDestination_(ctx, name,
        (f, text) => showProgress_(ss, tag, (d + f) / names.length * 100, text)));
    });
  } catch (err) {
    console.error(err.stack);
    results.push({ destination: label, status: "FAILED", errors: [{ msg: err.message }] });
  } finally {
    lock.releaseLock();
    writeLogs_(ss, fnName, runType, results);
  }

  const rows = results.reduce((a, r) => a + (r.rows || 0), 0);
  const bad = results.filter(r => r.status !== "SUCCESS").map(r => r.destination + ": " + r.status);
  ss.toast("■■■■■■■■■■ 100% | " + rows.toLocaleString() + " rows. " +
           (bad.length ? "Check: " + bad.join(", ") : "All OK"), "Completed - " + label, 15);
}


// ============ PREPARE (read Index once) ============
function prepare_(ss) {
  const indexSheet = ss.getSheetByName(CONFIG.INDEX_SHEET);
  if (!indexSheet) throw new Error("Index sheet not found!");
  const lastRow = indexSheet.getLastRow();
  if (lastRow < CONFIG.START_ROW) throw new Error("No rows in Index sheet");

  const numRows = lastRow - CONFIG.START_ROW + 1;
  const indexData = indexSheet.getRange(CONFIG.START_ROW, 1, numRows, 6).getValues();
  const linkData  = indexSheet.getRange(CONFIG.START_ROW, 2, numRows, 1).getRichTextValues();
  const token = ScriptApp.getOAuthToken();

  indexSheet.getRange(1, CONFIG.STATUS_COL).setValue("Status");
  const statusRange = indexSheet.getRange(CONFIG.START_ROW, CONFIG.STATUS_COL, numRows, 1);
  const statusVals = statusRange.getValues();

  const groups = {}, destOrder = [], lookup = {};
  let lastPasteCell = "";

  indexData.forEach((row, i) => {
    const rowNumber = CONFIG.START_ROW + i;
    const pasteCellRaw = String(row[5] || "").trim();  // merged F -> carry down
    if (pasteCellRaw) lastPasteCell = pasteCellRaw;

    const destinationName = String(row[4] || "").trim();
    if (!row[1] && !row[2] && !row[3] && !destinationName) return;

    const key = destinationName || "(No destination)";
    if (!groups[key]) { groups[key] = []; destOrder.push(key); lookup[key.toLowerCase()] = key; }

    try {
      const spreadsheetId = getSpreadsheetId_(row[1], linkData[i][0]);
      const sheetName = String(row[2] || "").trim();
      const range = String(row[3] || "").trim();
      if (!spreadsheetId) throw new Error("Missing Sheet ID / link (col B)");
      if (!sheetName) throw new Error("Missing Sheet Name (col C)");
      if (!range) throw new Error("Missing Range (col D)");
      if (!destinationName) throw new Error("Missing Destination Sheet (col E)");

      const rangeInfo = parseSourceRange_(range);
      const keyLetter = KEY_COLUMN[destinationName.toLowerCase()];
      let keyCol = null;
      if (keyLetter) {
        keyCol = columnToNumber_(keyLetter) - rangeInfo.startColumn + 1;
        if (keyCol < 1 || keyCol > rangeInfo.numberOfColumns)
          throw new Error("Key column " + keyLetter + " is outside range " + range);
      }

      const fullRange = "'" + sheetName.replace(/'/g, "''") + "'!" + range;
      groups[key].push({
        rowNumber,
        pasteCell: parseCell_(lastPasteCell || "A2"),
        numberOfColumns: rangeInfo.numberOfColumns,
        keyCol,
        request: {
          url: "https://sheets.googleapis.com/v4/spreadsheets/" + spreadsheetId +
               "/values/" + encodeURIComponent(fullRange) +
               "?valueRenderOption=UNFORMATTED_VALUE&dateTimeRenderOption=SERIAL_NUMBER&majorDimension=ROWS",
          headers: { Authorization: "Bearer " + token },
          muteHttpExceptions: true
        }
      });
    } catch (e) {
      groups[key].push({ rowNumber, error: e.message });
    }
  });

  const meta = Sheets.Spreadsheets.get(ss.getId(), { fields: "sheets.properties(sheetId,title,gridProperties)" });
  const props = {};
  meta.sheets.forEach(s => props[s.properties.title] = s.properties);

  return {
    ss, ssId: ss.getId(), token, groups, destOrder, lookup, props,
    setStatus: (r, t) => { statusVals[r - CONFIG.START_ROW][0] = t; },
    getStatus: (r) => String(statusVals[r - CONFIG.START_ROW][0]),
    saveStatus: () => { statusRange.setValues(statusVals); SpreadsheetApp.flush(); }
  };
}


// ============ PROCESS ONE DESTINATION ============
function processDestination_(ctx, name, prog) {
  const t0 = Date.now();
  const res = { destination: name, status: "SUCCESS", ok: 0, failed: 0, rows: 0,
                readS: 0, writeS: 0, totalS: 0, errors: [] };
  const srcs = ctx.groups[name] || [];
  const valid = [];

  srcs.forEach(s => {
    if (s.error) { res.errors.push({ row: s.rowNumber, msg: s.error }); ctx.setStatus(s.rowNumber, "FAILED: " + s.error); }
    else { valid.push(s); ctx.setStatus(s.rowNumber, "Waiting..."); }
  });
  ctx.saveStatus();

  try {
    if (!valid.length) throw new Error("No valid Index rows");
    const p = ctx.props[name];
    if (!p) throw new Error("Destination tab '" + name + "' not found in this file");

    // ---- READ ----
    const tr = Date.now();
    prog(0.01, "Reading " + valid.length + " source sheet(s)...");
    const responses = fetchInBatches_(valid.map(s => s.request), CONFIG.READ_PARALLEL, CONFIG.MAX_RETRIES,
      (done, total) => prog(0.40 * done / total, "Reading " + done + "/" + total + " sources"));
    res.readS = (Date.now() - tr) / 1000;

    const n = valid[0].numberOfColumns;
    const pasteCell = valid[0].pasteCell;
    let data = [];
    let readOk = 0;

    prog(0.42, "Cleaning rows...");
    responses.forEach((r, i) => {
      const s = valid[i];
      try {
        const code = r.getResponseCode();
        if (code !== 200) throw new Error(friendlyError_(code, r.getContentText()));
        if (s.numberOfColumns !== n) throw new Error("Range width differs from other rows (" + s.numberOfColumns + " vs " + n + " cols)");
        const values = JSON.parse(r.getContentText()).values || [];
        s._added = cleanRows_(values, n, s.keyCol, data);
        readOk++;
        ctx.setStatus(s.rowNumber, "Read " + s._added.toLocaleString() + " rows");
      } catch (e) {
        res.errors.push({ row: s.rowNumber, msg: e.message });
        ctx.setStatus(s.rowNumber, "FAILED: " + e.message);
      }
    });

    if (!readOk) throw new Error("All sources failed - old data kept");

    // ---- CLEAR + EXPAND + WRITE ----
    const tw = Date.now();
    const startRow = pasteCell.row, startCol = pasteCell.col;
    const endCol = startCol + n - 1;
    const q = "'" + name.replace(/'/g, "''") + "'!";

    prog(0.45, "Clearing old data...");
    Sheets.Spreadsheets.Values.batchClear(
      { ranges: [q + columnToLetter_(startCol) + startRow + ":" + columnToLetter_(endCol)] }, ctx.ssId);

    const expand = [];
    const neededRows = startRow + data.length - 1;
    if (neededRows > p.gridProperties.rowCount) {
      expand.push({ appendDimension: { sheetId: p.sheetId, dimension: "ROWS", length: neededRows - p.gridProperties.rowCount } });
      p.gridProperties.rowCount = neededRows;
    }
    if (endCol > p.gridProperties.columnCount) {
      expand.push({ appendDimension: { sheetId: p.sheetId, dimension: "COLUMNS", length: endCol - p.gridProperties.columnCount } });
      p.gridProperties.columnCount = endCol;
    }
    if (expand.length) Sheets.Spreadsheets.batchUpdate({ requests: expand }, ctx.ssId);

    res.rows = data.length;
    if (data.length) {
      prog(0.50, "Writing " + data.length.toLocaleString() + " rows...");
      const rowsPerChunk = Math.max(1, Math.floor(CONFIG.WRITE_CHUNK_CELLS / n));
      const writeRequests = [];
      for (let i = 0; i < data.length; i += rowsPerChunk) {
        const chunk = data.slice(i, i + rowsPerChunk);
        const r1 = startRow + i;
        writeRequests.push({
          url: "https://sheets.googleapis.com/v4/spreadsheets/" + ctx.ssId + "/values:batchUpdate",
          method: "post",
          contentType: "application/json",
          headers: { Authorization: "Bearer " + ctx.token },
          muteHttpExceptions: true,
          payload: JSON.stringify({
            valueInputOption: "RAW",
            data: [{ range: q + columnToLetter_(startCol) + r1 + ":" + columnToLetter_(endCol) + (r1 + chunk.length - 1),
                     majorDimension: "ROWS", values: chunk }]
          })
        });
      }
      data = null;

      const wr = fetchInBatches_(writeRequests, CONFIG.WRITE_PARALLEL, CONFIG.MAX_RETRIES,
        (done, total) => prog(0.50 + 0.50 * done / total, "Writing " + done + "/" + total + " chunks"));
      const bad = wr.filter(r => r.getResponseCode() !== 200);
      if (bad.length) throw new Error(bad.length + " write chunk(s) failed: " +
                                      friendlyError_(bad[0].getResponseCode(), bad[0].getContentText()));
    }
    res.writeS = (Date.now() - tw) / 1000;

    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MMM HH:mm");
    valid.forEach(s => {
      if (ctx.getStatus(s.rowNumber).indexOf("Read ") === 0)
        ctx.setStatus(s.rowNumber, "OK - " + s._added.toLocaleString() + " rows | " + stamp);
    });
    prog(1, "Done");
  } catch (e) {
    res.errors.push({ row: null, msg: e.message });
    valid.forEach(s => {
      if (ctx.getStatus(s.rowNumber).indexOf("FAILED") !== 0)
        ctx.setStatus(s.rowNumber, "FAILED: " + e.message.substring(0, 150));
    });
    console.error(name + ": " + e.message);
  }

  srcs.forEach(s => { if (ctx.getStatus(s.rowNumber).indexOf("OK") === 0) res.ok++; else res.failed++; });
  res.status = res.ok === 0 ? "FAILED" : (res.failed > 0 ? "PARTIAL" : "SUCCESS");
  if (res.ok === 0) res.rows = 0;
  res.totalS = (Date.now() - t0) / 1000;

  const last = getLast_(name);
  last.lastRun = Date.now();
  last.dur = Math.round(res.totalS);
  if (res.status !== "FAILED") last.lastOk = Date.now();
  setLast_(name, last);

  ctx.saveStatus();
  return res;
}


// ============ FRIENDLY ERRORS ============
function friendlyError_(code, text) {
  text = String(text || "");
  if (code === 400 && text.indexOf("Unable to parse range") > -1) return "Tab not found - check Sheet Name (col C) / Range (col D)";
  if (code === 400) return "Bad request - check Range (col D)";
  if (code === 401) return "Login expired - run again";
  if (code === 403) return "No access to source file - share it with the script owner";
  if (code === 404) return "Source file not found - check link (col B)";
  if (code === 429) return "Google quota limit (too many requests) - will retry next run";
  if (code === 500 || code === 503) return "Google busy / source file too heavy (timeout)";
  if (code === 599) return "Timeout / network error - source file too heavy";
  let m = text;
  try { m = JSON.parse(text).error.message; } catch (e) {}
  return "API " + code + ": " + String(m).substring(0, 100);
}

// Groups same messages: "Rows 56, 57, 58: Tab not found ..."
function summarizeErrors_(errors) {
  if (!errors || !errors.length) return "";
  const map = {}, order = [];
  errors.forEach(e => {
    const msg = String(e.msg || "").replace(/\s+/g, " ").trim();
    if (!map[msg]) { map[msg] = []; order.push(msg); }
    if (e.row) map[msg].push(e.row);
  });
  return order.map(msg => (map[msg].length ? "Row" + (map[msg].length > 1 ? "s " : " ") + map[msg].join(", ") + ": " : "") + msg)
              .join(" | ").substring(0, 450);
}


// ============ LOGGING ============
function getLogSheet_(ss) {
  let sh = ss.getSheetByName(CONFIG.LOG_SHEET);
  if (sh) {
    const head = sh.getRange(1, 1, 1, LOG_HEADERS.length).getValues()[0];
    if (head.join("|") !== LOG_HEADERS.join("|")) {
      if (sh.getLastRow() > 0) {
        sh.setName(CONFIG.LOG_SHEET + "_old_" +
          Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "ddMMM_HHmm"));
        sh = null;
      } else {
        sh.clear();
      }
    }
  }
  if (!sh) sh = ss.getSheetByName(CONFIG.LOG_SHEET) || ss.insertSheet(CONFIG.LOG_SHEET);

  if (sh.getRange("A1").getValue() !== LOG_HEADERS[0]) {
    sh.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS])
      .setFontWeight("bold").setBackground("#434343").setFontColor("#ffffff");
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 150); sh.setColumnWidth(2, 160); sh.setColumnWidth(3, 170);
    sh.setColumnWidth(12, 450);
    buildLogSummary_(sh);
  }
  return sh;
}

function writeLogs_(ss, runBy, runType, results) {
  try {
    if (!results || !results.length) return;
    const sh = getLogSheet_(ss);
    const now = new Date();
    const rows = results.map(r => [
      now, runBy, r.destination || "", runType, r.status || "",
      r.ok || 0, r.failed || 0, r.rows || 0,
      round1_(r.readS), round1_(r.writeS), round1_(r.totalS),
      summarizeErrors_(r.errors)
    ]);

    // Newest on top
    sh.insertRowsBefore(2, rows.length);
    const rng = sh.getRange(2, 1, rows.length, LOG_HEADERS.length);
    rng.setValues(rows)
       .setWrap(false).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
       .setVerticalAlignment("middle").setFontWeight("normal")
       .setBackground(null).setFontColor("#000000");
    sh.getRange(2, 1, rows.length, 1).setNumberFormat("dd-MMM-yyyy HH:mm:ss");
    sh.getRange(2, 9, rows.length, 3).setNumberFormat("0.0");
    sh.getRange(2, 8, rows.length, 1).setNumberFormat("#,##0");
    sh.setRowHeights(2, rows.length, 21);

    const colours = { SUCCESS: "#d9ead3", PARTIAL: "#fff2cc", FAILED: "#f4cccc", SKIPPED: "#e8e8e8" };
    sh.getRange(2, 5, rows.length, 1).setBackgrounds(rows.map(r => [colours[r[4]] || null]));

    // Keep only latest rows
    const extra = sh.getLastRow() - 1 - CONFIG.LOG_MAX_ROWS;
    if (extra > 0) sh.deleteRows(CONFIG.LOG_MAX_ROWS + 2, extra);
  } catch (e) {
    console.error("Log write failed: " + e.message);
  }
}

function round1_(v) { return Number(Number(v || 0).toFixed(1)); }

// Summary table at N:U, by destination
function buildLogSummary_(sh) {
  const head = ["Destination", "Success", "Partial", "Failed", "Skipped", "Success %", "Last Success", "Avg Time (s)"];
  sh.getRange(1, 14, 1, head.length).setValues([head])
    .setFontWeight("bold").setBackground("#1f4e78").setFontColor("#ffffff");

  const rows = DESTINATIONS.map((name, i) => {
    const r = i + 2;
    return [
      name,
      '=COUNTIFS($C:$C,N' + r + ',$E:$E,"SUCCESS")',
      '=COUNTIFS($C:$C,N' + r + ',$E:$E,"PARTIAL")',
      '=COUNTIFS($C:$C,N' + r + ',$E:$E,"FAILED")',
      '=COUNTIFS($C:$C,N' + r + ',$E:$E,"SKIPPED")',
      '=IFERROR(O' + r + '/(O' + r + '+P' + r + '+Q' + r + '),"")',
      '=IFERROR(1/(1/MAXIFS($A:$A,$C:$C,N' + r + ',$E:$E,"SUCCESS")),"never")',
      '=IFERROR(AVERAGEIFS($K:$K,$C:$C,N' + r + ',$E:$E,"<>SKIPPED"),"")'
    ];
  });
  sh.getRange(2, 14, rows.length, head.length).setValues(rows);
  sh.getRange(2, 19, rows.length, 1).setNumberFormat("0%");
  sh.getRange(2, 20, rows.length, 1).setNumberFormat("dd-MMM HH:mm");
  sh.getRange(2, 21, rows.length, 1).setNumberFormat("0.0");

  const tr = rows.length + 2;
  sh.getRange(tr, 14, 1, 5).setValues([[
    "TOTAL", "=SUM(O2:O" + (tr - 1) + ")", "=SUM(P2:P" + (tr - 1) + ")",
    "=SUM(Q2:Q" + (tr - 1) + ")", "=SUM(R2:R" + (tr - 1) + ")"
  ]]).setFontWeight("bold").setBackground("#d9e1f2");
  sh.setColumnWidth(14, 180);
}

function rebuildLogSummary() {
  const sh = getLogSheet_(SpreadsheetApp.getActiveSpreadsheet());
  sh.getRange(1, 14, 40, 8).clear();
  buildLogSummary_(sh);
}


// ============ LAST RUN MEMORY ============
function getLast_(name) {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty("last_" + name) || "{}"); }
  catch (e) { return {}; }
}
function setLast_(name, obj) {
  PropertiesService.getScriptProperties().setProperty("last_" + name, JSON.stringify(obj));
}


// ============ PROGRESS BAR (toast) ============
function showProgress_(ss, title, pct, text) {
  pct = Math.max(1, Math.min(100, Math.round(pct)));
  const filled = Math.round(pct / 10);
  ss.toast("■".repeat(filled) + "□".repeat(10 - filled) + " " + pct + "% | " + text, title, -1);
}


// ============ ROW CLEANING ============
function cleanRows_(values, n, keyCol, out) {
  let added = 0;
  for (let r = 0; r < values.length; r++) {
    const row = values[r];
    if (keyCol) {
      const k = row[keyCol - 1];
      if (k === undefined || k === null || k === 0 || k === "" ||
          (typeof k === "string" && (k.trim() === "" || k.trim() === "0"))) continue;
    } else {
      let hasData = false;
      for (let c = 0; c < row.length; c++) {
        const v = row[c];
        if (v !== null && v !== "" && !(typeof v === "string" && v.trim() === "")) { hasData = true; break; }
      }
      if (!hasData) continue;
    }
    const o = new Array(n);
    for (let c = 0; c < n; c++) o[c] = (c < row.length && row[c] !== null) ? row[c] : "";
    out.push(o);
    added++;
  }
  return added;
}


// ============ PARALLEL FETCH WITH RETRY ============
function fetchInBatches_(requests, parallel, maxRetries, onProgress) {
  const responses = new Array(requests.length);
  for (let i = 0; i < requests.length; i += parallel) {
    const idx = [];
    for (let j = i; j < Math.min(i + parallel, requests.length); j++) idx.push(j);
    const res = fetchAllWithRetry_(idx.map(j => requests[j]), maxRetries);
    res.forEach((r, k) => responses[idx[k]] = r);
    if (onProgress) onProgress(Math.min(i + parallel, requests.length), requests.length);
  }
  return responses;
}

function fetchAllWithRetry_(requests, maxRetries) {
  const responses = safeFetchAll_(requests);
  let pending = [];
  responses.forEach((r, i) => { const c = r.getResponseCode(); if (c === 429 || c >= 500) pending.push(i); });

  for (let attempt = 1; attempt <= maxRetries && pending.length; attempt++) {
    Utilities.sleep(Math.pow(2, attempt) * 1000 + Math.floor(Math.random() * 1000));
    const retry = safeFetchAll_(pending.map(i => requests[i]));
    const still = [];
    retry.forEach((r, k) => {
      const idx = pending[k];
      responses[idx] = r;
      const c = r.getResponseCode();
      if (c === 429 || c >= 500) still.push(idx);
    });
    pending = still;
  }
  return responses;
}

function safeFetchAll_(requests) {
  try {
    return UrlFetchApp.fetchAll(requests);
  } catch (e) {
    return requests.map(req => {
      try { return UrlFetchApp.fetch(req.url, req); }
      catch (err) { const m = String(err.message); return { getResponseCode: () => 599, getContentText: () => m }; }
    });
  }
}


// ============ HELPERS ============
function parseCell_(cell) {
  const m = String(cell).trim().match(/^\$?([A-Z]+)\$?(\d+)$/i);
  if (!m) throw new Error("Invalid paste cell (col F): " + cell);
  return { col: columnToNumber_(m[1]), row: Number(m[2]) };
}

function getSpreadsheetId_(cellValue, richTextValue) {
  let value = String(cellValue || "").trim();
  if (richTextValue) {
    let link = richTextValue.getLinkUrl();
    if (!link) {
      const run = richTextValue.getRuns().find(r => r.getLinkUrl());
      if (run) link = run.getLinkUrl();
    }
    if (link) value = link;
  }
  const m = value.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : value;
}

function parseSourceRange_(range) {
  const m = range.match(/^\$?([A-Z]+)\$?\d+:\$?([A-Z]+)(?:\$?\d+)?$/i);
  if (!m) throw new Error("Invalid Range (col D): " + range);
  const s = columnToNumber_(m[1]), e = columnToNumber_(m[2]);
  if (e < s) throw new Error("Invalid column order in Range (col D): " + range);
  return { numberOfColumns: e - s + 1, startColumn: s };
}

function columnToNumber_(col) {
  let n = 0; col = col.toUpperCase();
  for (let i = 0; i < col.length; i++) n = n * 26 + col.charCodeAt(i) - 64;
  return n;
}

function columnToLetter_(n) {
  let s = "";
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}