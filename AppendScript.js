// ================= MENU =================
function onOpen() {
  SpreadsheetApp.getUi().createMenu("Billing FMS")
    .addItem("Billing FMS", "importBillingFMS")
    .addItem("WBS Sheet", "importWBSSheet")
    .addItem("INVENTORY", "importInventory")
    .addItem("Manpower Cost", "importManpowerCost")
    .addItem("PROJECT BOQ", "importProjectBOQ")
    .addItem("PROJECT PLANNING", "importProjectPlanning")
    .addSeparator()
    .addItem("Fetch ALL", "fetchAllBillingFMSData")
    .addToUi();
}

// ============ SEPARATE SCRIPTS ============
function importBillingFMS()       { runImport_(["Billing FMS"]); }
function importWBSSheet()         { runImport_(["WBS Sheet"]); }
function importInventory()        { runImport_(["INVENTORY"]); }
function importManpowerCost()     { runImport_(["Manpower Cost"]); }
function importProjectBOQ()       { runImport_(["PROJECT BOQ"]); }
function importProjectPlanning()  { runImport_(["PROJECT PLANNING"]); }
function fetchAllBillingFMSData() { runImport_(null); } // null = all destinations


// ============ SETTINGS ============
// Key column per destination: a row is copied only if this column is NOT blank and NOT 0.
// Number = position inside the source range (1 = first column of the range).
// Destinations not listed here use the normal "skip fully blank rows" rule.
// Use lowercase names here.
const KEY_COLUMN = {
  "project boq": "A",        // ISO NO.
  "project planning": "H",   // H column must not be blank / 0
  "inventory": "W"           // W column must not be blank / 0
};


// ============ CORE ENGINE ============
function runImport_(destinationFilter) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssId = ss.getId();
  const indexSheet = ss.getSheetByName("Index");
  const START_ROW = 2;
  const WRITE_CHUNK_CELLS = 400000;
  const label = destinationFilter ? destinationFilter.join(", ") : "All";

  if (!indexSheet) throw new Error("Index sheet not found!");
  const lastRow = indexSheet.getLastRow();
  if (lastRow < START_ROW) { ss.toast("No source records found!", label, 5); return; }

  const filterSet = destinationFilter
    ? new Set(destinationFilter.map(n => n.trim().toLowerCase()))
    : null;

  ss.toast("Importing...", label, 3);

  const numRows = lastRow - START_ROW + 1;
  const indexData = indexSheet.getRange(START_ROW, 1, numRows, 6).getValues();   // A:F
  const linkData  = indexSheet.getRange(START_ROW, 2, numRows, 1).getRichTextValues();
  const token = ScriptApp.getOAuthToken();

  let successCount = 0, failedCount = 0, totalRecords = 0;
  const sources = [], requests = [];
  let lastPasteCell = "";

  // STEP 1: Read Index rows
  indexData.forEach((row, i) => {
    // Column F is merged -> carry value down (runs for every row, before filtering)
    const pasteCellRaw = String(row[5] || "").trim();
    if (pasteCellRaw) lastPasteCell = pasteCellRaw;

    const destinationName = String(row[4] || "").trim();
    if (filterSet && !filterSet.has(destinationName.toLowerCase())) return;

    try {
      const spreadsheetId = getSpreadsheetId_(row[1], linkData[i][0]);
      const sheetName = String(row[2] || "").trim();
      const range = String(row[3] || "").trim();

      if (!spreadsheetId || !sheetName || !range || !destinationName) {
        if (!row[1] && !row[2] && !row[3] && !row[4]) return; // fully blank Index row
        throw new Error("Missing Index sheet details.");
      }

      const fullRange = "'" + sheetName.replace(/'/g, "''") + "'!" + range;
      requests.push({
        url: "https://sheets.googleapis.com/v4/spreadsheets/" + spreadsheetId +
             "/values/" + encodeURIComponent(fullRange) +
             "?valueRenderOption=UNFORMATTED_VALUE&dateTimeRenderOption=SERIAL_NUMBER&majorDimension=ROWS",
        headers: { Authorization: "Bearer " + token },
        muteHttpExceptions: true
      });
      sources.push({
        rowNumber: START_ROW + i,
        destinationName,
        pasteCell: parseCell_(lastPasteCell || "A2"),
        numberOfColumns: parseSourceRange_(range).numberOfColumns,
        startColumn: parseSourceRange_(range).startColumn
      });
    } catch (e) {
      failedCount++;
      console.error("Index row " + (START_ROW + i) + ": " + e.message);
    }
  });

  if (!requests.length) { ss.toast("No matching Index rows found!", label, 5); return; }

  // STEP 2: Fetch all in parallel
  const responses = fetchAllWithRetry_(requests, 4);

  // STEP 3: Clean + group by destination (Index order kept)
  const destinationData = {}, destOrder = [];

  responses.forEach((response, i) => {
    const src = sources[i];
    try {
      if (response.getResponseCode() !== 200)
        throw new Error("API error " + response.getResponseCode() + ": " +
                        response.getContentText().substring(0, 300));

      const values = JSON.parse(response.getContentText()).values || [];
      const n = src.numberOfColumns;

      let dest = destinationData[src.destinationName];
      if (!dest) {
        dest = destinationData[src.destinationName] =
          { data: [], numberOfColumns: n, pasteCell: src.pasteCell };
        destOrder.push(src.destinationName);
      }
      if (dest.numberOfColumns !== n)
        throw new Error("Column count " + n + " differs from " + dest.numberOfColumns +
                        " for " + src.destinationName);

      const keyLetter = KEY_COLUMN[src.destinationName.toLowerCase()];
      let keyCol = null;
      if (keyLetter) {
        keyCol = columnToNumber_(keyLetter) - src.startColumn + 1; // position inside the range
        if (keyCol < 1 || keyCol > n)
          throw new Error("Key column " + keyLetter + " is outside the source range for " + src.destinationName);
      }
      let added = 0;

      for (const row of values) {
        if (keyCol) {
          // Copy only if the key column has a real value (not blank, not 0)
          const k = row[keyCol - 1];
          if (k === undefined || k === null || k === 0 ||
              (typeof k === "string" && (k.trim() === "" || k.trim() === "0"))) continue;
        } else {
          // Skip fully blank rows
          let hasData = false;
          for (const v of row) {
            if (v !== null && v !== "" && !(typeof v === "string" && v.trim() === "")) { hasData = true; break; }
          }
          if (!hasData) continue;
        }

        // Exact width: pad short rows, cut long ones
        const out = new Array(n);
        for (let c = 0; c < n; c++) out[c] = (c < row.length && row[c] !== null) ? row[c] : "";
        dest.data.push(out);
        added++;
      }

      totalRecords += added;
      successCount++;
      console.log("Index row " + src.rowNumber + " -> " + src.destinationName + ": " + added + " rows");
    } catch (e) {
      failedCount++;
      console.error("Index row " + src.rowNumber + ": " + e.message);
    }
  });

  // STEP 4: Clear old data, then paste from start cell
  const meta = Sheets.Spreadsheets.get(ssId, { fields: "sheets.properties(sheetId,title,gridProperties)" });
  const props = {};
  meta.sheets.forEach(s => props[s.properties.title] = s.properties);

  const clearRanges = [], expandRequests = [], writes = [];

  destOrder.forEach(name => {
    const p = props[name];
    if (!p) { failedCount++; console.error("Destination sheet not found: " + name); return; }
    const d = destinationData[name];
    const startRow = d.pasteCell.row, startCol = d.pasteCell.col;
    const endCol = startCol + d.numberOfColumns - 1;
    const q = "'" + name.replace(/'/g, "''") + "'!";

    // Clear target columns from start row to bottom (removes stale rows)
    clearRanges.push(q + columnToLetter_(startCol) + startRow + ":" + columnToLetter_(endCol));

    const neededRows = startRow + d.data.length - 1;
    if (neededRows > p.gridProperties.rowCount)
      expandRequests.push({ appendDimension: { sheetId: p.sheetId, dimension: "ROWS",
        length: neededRows - p.gridProperties.rowCount } });
    if (endCol > p.gridProperties.columnCount)
      expandRequests.push({ appendDimension: { sheetId: p.sheetId, dimension: "COLUMNS",
        length: endCol - p.gridProperties.columnCount } });

    if (d.data.length) writes.push({ q, name, startRow, startCol, endCol, data: d.data });
  });

  if (clearRanges.length) Sheets.Spreadsheets.Values.batchClear({ ranges: clearRanges }, ssId);
  if (expandRequests.length) Sheets.Spreadsheets.batchUpdate({ requests: expandRequests }, ssId);

  let batch = [], batchCells = 0;
  const flush = () => {
    if (!batch.length) return;
    Sheets.Spreadsheets.Values.batchUpdate({ valueInputOption: "RAW", data: batch }, ssId);
    batch = []; batchCells = 0;
  };

  writes.forEach(w => {
    const cols = w.endCol - w.startCol + 1;
    const rowsPerChunk = Math.max(1, Math.floor(WRITE_CHUNK_CELLS / cols));
    for (let i = 0; i < w.data.length; i += rowsPerChunk) {
      const chunk = w.data.slice(i, i + rowsPerChunk);
      const r1 = w.startRow + i;
      batch.push({
        range: w.q + columnToLetter_(w.startCol) + r1 + ":" + columnToLetter_(w.endCol) + (r1 + chunk.length - 1),
        majorDimension: "ROWS",
        values: chunk
      });
      batchCells += chunk.length * cols;
      if (batchCells >= WRITE_CHUNK_CELLS) flush();
    }
    console.log(w.name + ": " + w.data.length + " rows pasted from " + columnToLetter_(w.startCol) + w.startRow);
  });
  flush();

  // STEP 5: Summary
  console.log(label + " | OK: " + successCount + " | Failed: " + failedCount + " | Rows: " + totalRecords);
  ss.toast("Completed! " + totalRecords + " rows pasted. Failed: " + failedCount, label, 10);
}



// Fetch all in parallel; retry only failed ones (429 / 5xx) with increasing wait
function fetchAllWithRetry_(requests, maxRetries) {
  const responses = UrlFetchApp.fetchAll(requests);
  let pending = [];
  responses.forEach((r, i) => {
    const code = r.getResponseCode();
    if (code === 429 || code >= 500) pending.push(i);
  });

  for (let attempt = 1; attempt <= maxRetries && pending.length; attempt++) {
    Utilities.sleep(Math.pow(2, attempt) * 1000 + Math.floor(Math.random() * 1000)); // ~2s, 4s, 8s, 16s
    console.log("Retry " + attempt + " for " + pending.length + " request(s)");
    const retryResponses = UrlFetchApp.fetchAll(pending.map(i => requests[i]));
    const stillFailing = [];
    retryResponses.forEach((r, k) => {
      const idx = pending[k];
      responses[idx] = r;
      const code = r.getResponseCode();
      if (code === 429 || code >= 500) stillFailing.push(idx);
    });
    pending = stillFailing;
  }
  return responses;
}



// ============ HELPERS ============
// "A2" -> {col:1,row:2}
function parseCell_(cell) {
  const m = String(cell).trim().match(/^\$?([A-Z]+)\$?(\d+)$/i);
  if (!m) throw new Error("Invalid paste cell: " + cell);
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
  if (!m) throw new Error("Invalid source range: " + range);
  const s = columnToNumber_(m[1]), e = columnToNumber_(m[2]);
  if (e < s) throw new Error("Invalid column range: " + range);
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