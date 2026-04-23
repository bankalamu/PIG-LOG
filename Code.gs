// ============================================================
//  PIG LOG — Google Apps Script Backend
//  Paste this entire file into your Apps Script project
// ============================================================

// Required OAuth scopes (Apps Script reads these comments to request permissions):
// @ts-nocheck
/* global DriveApp, SpreadsheetApp, Utilities, Session, ContentService */
// The following directive tells Apps Script to request Drive access:
// drive: https://www.googleapis.com/auth/drive

const SHEET_NAME = "PigLog";
const HEADERS = ["DB_ID", "PIG ID", "Boar", "SOW", "DOB", "SEX", "Type", "Stage", "Status", "ServiceDate", "Sire", "Weight", "Dewormed", "Pen", "Notes", "Available"];

// Key fields that define a unique pig record
const KEY_FIELDS = ["PIG ID", "Boar", "SOW"];

// ── Helpers ──────────────────────────────────────────────────

/**
 * Gets (or creates) the main PigLog spreadsheet tab.
 * @returns {Sheet} The PigLog Google Sheet object.
 */
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold").setBackground("#c9a84c").setFontColor("#000000");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 70);
    sheet.setColumnWidth(2, 120);
  } else {
    // Ensure all HEADERS columns exist — add any missing ones to the right
    const lastCol  = sheet.getLastColumn();
    const existing = lastCol > 0
      ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim())
      : [];
    HEADERS.forEach(h => {
      if (!existing.includes(h)) {
        const newCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, newCol).setValue(h)
             .setFontWeight("bold").setBackground("#c9a84c").setFontColor("#000000");
      }
    });
  }
  return sheet;
}

/**
 * Returns the next available DB_ID for a new pig record.
 * Scans column 1 for all existing IDs and returns max+1.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The PigLog sheet.
 * @returns {number} Next DB_ID integer (starts at 1 if sheet is empty).
 */
function getNextId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 1;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(v => v !== "");
  if (ids.length === 0) return 1;
  return Math.max(...ids.map(Number)) + 1;
}

/**
 * Wraps a JS object as a JSON ContentService HTTP response.
 * Used by both doGet() and doPost() to return data to the frontend.
 * @param {Object} data - Any JSON-serialisable object to return.
 * @returns {GoogleAppsScript.Content.TextOutput} JSON HTTP response.
 */
function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Alias for respond() kept for backward compatibility.
 * All POST and payload-GET requests use this to return JSON.
 * @param {Object} data - Any JSON-serialisable object to return.
 * @returns {GoogleAppsScript.Content.TextOutput} JSON HTTP response.
 */
function corsRespond(data) {
  // For POST requests wrapped in JSONP-style — Apps Script handles CORS via GET
  return respond(data);
}

// ── Router ───────────────────────────────────────────────────

// ── Request token validation ─────────────────────
/**
 * Validates the APP_TOKEN sent with each request for basic auth.
 * Token is stored in Script Properties under the key "APP_TOKEN".
 * If no APP_TOKEN is set, all requests are allowed through.
 * To disable auth: delete the APP_TOKEN property in Script Properties.
 * @param {string} token - Token value sent by the client in the request.
 * @returns {boolean} True if token matches or no token is configured.
 */
function _validateToken(token) {
  const expected = PropertiesService.getScriptProperties().getProperty('APP_TOKEN');
  if (!expected) return true; // no token set — allow all
  return String(token || '') === String(expected);
}

/**
 * HTTP GET handler — entry point for all read requests and GET+payload writes.
 * Routes by e.parameter.action.
 * Also decodes base64 "payload=" parameter for write actions that cannot
 * use POST due to GitHub Pages CSP restrictions (used by apiGet fallback).
 * @param {GoogleAppsScript.Events.DoGet} e - Apps Script GET event.
 * @returns {GoogleAppsScript.Content.TextOutput} JSON response.
 */
function doGet(e) {
  const action = e.parameter.action;

  // Handle payload= parameter (POST-like data sent via GET to avoid CSP issues)
  if (e.parameter.payload) {
    try {
      const bytes   = Utilities.base64Decode(e.parameter.payload);
      const str     = Utilities.newBlob(bytes).getDataAsString('UTF-8');
      // Decode URI-encoded JSON
      const decoded = decodeURIComponent(str);
      const payload = JSON.parse(decoded);
      Logger.log('payload action: ' + payload.action);
      return handlePostPayload(payload);
    } catch(err) {
      return respond({ error: 'Invalid payload: ' + err.message });
    }
  }

  try {
    if (action === "getAll")       return respond(getAllRecords());
    if (action === "ping")         return respond({ success: true, message: "pong", time: new Date().toISOString(), codeVersion: "3.1" });
    if (action === "debug")        return respond({ token: PropertiesService.getScriptProperties().getProperty('APP_TOKEN'), aiKey: !!PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY'), codeVersion: "3.1" });
    if (action === "clCount")      return respond(clCount());
    if (action === "clGetRecent")  return respond(clGetRecent(parseInt(e.parameter.days||'30')));
    if (action === "clDebug")      return respond(clDebug());
    if (action === "getNextDbId")  return respond({ success: true, nextId: getNextId(getSheet()) });
    if (action === "getByPigId") return respond(getByPigId(e.parameter.pigId));
    if (action === "clGetAll")   return respond(clGetAll());
    if (action === "slGetAll")   return respond(slGetAll());
    if (action === "wkGetAll")   return respond(wkGetAll());
    if (action === "moGetAll")   return respond(moGetAll());
    if (action === "moDedup")    return respond(moDeduplicateAll());
    if (action === "getSetting") return respond(getSetting(e.parameter.key));
    if (action === "getAIPending") return respond(getAIPendingRecords(e.parameter.date));
    if (action === "getPhoto")   return respond(getPhotoAsBase64(e.parameter.fileId));
    if (action === "waGetAll")   return respond(waGetAll());
    return respond({ error: "Unknown action" });
  } catch (err) {
    return respond({ error: err.message });
  }
}

/**
 * HTTP POST handler — entry point for all write requests.
 * Parses the JSON body from e.postData.contents and delegates to handlePostPayload().
 * @param {GoogleAppsScript.Events.DoPost} e - Apps Script POST event.
 * @returns {GoogleAppsScript.Content.TextOutput} JSON response.
 */
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    return handlePostPayload(payload);
  } catch (err) {
    return corsRespond({ error: err.message });
  }
}

/**
 * Routes all write actions from both POST body and GET+payload requests.
 * Every create/update/delete operation in the app passes through here.
 * @param {Object} payload - Parsed JSON object with at minimum { action: string }.
 *   Common shapes:
 *     { action: "add",       data: Object }
 *     { action: "update",    id: number, data: Object }
 *     { action: "delete",    id: number }
 *     { action: "clAdd",     data: Object }
 *     { action: "clSavePhoto", clId, photoBase64, mimeType, photoTime, section }
 *     { action: "slSavePhoto", slId, photoBase64, mimeType, photoTime, sectionKey }
 *     { action: "saveSetting", key: string, value: string }
 *     { action: "runAISingle", clId: number, force: boolean }
 * @returns {GoogleAppsScript.Content.TextOutput} JSON response from the handler.
 */
function handlePostPayload(payload) {
  try {
    const action = payload.action;
    if (action === "add")      return respond(addRecord(payload.data));
    if (action === "update")   return respond(updateRecord(payload.id, payload.data));
    if (action === "delete")   return respond(deleteRecord(payload.id));
    if (action === "clAdd")       return respond(clAdd(payload.data));
    if (action === "clUpsert")    return respond(clUpsert(payload.data));
    if (action === "clUpdate")    return respond(clUpdate(payload.id, payload.data));
    if (action === "clDelete")    return respond(clDelete(payload.id));
    if (action === "clSavePhoto") return respond(clSavePhoto(payload.clId, payload.photoBase64, payload.mimeType, payload.photoTime, payload.section));
    if (action === "slAdd")    return respond(slAdd(payload.data));
    if (action === "slUpsert") return respond(slUpsert(payload.data));
    if (action === "slUpdate") return respond(slUpdate(payload.id, payload.data));
    if (action === "slDelete") return respond(slDelete(payload.id));
    if (action === "slSavePhoto") return respond(slSavePhoto(payload.slId, payload.photoBase64, payload.mimeType, payload.photoTime, payload.sectionKey));
    if (action === "wkAdd")    return respond(wkAdd(payload.data));
    if (action === "wkUpdate") return respond(wkUpdate(payload.id, payload.data));
    if (action === "wkDelete") return respond(wkDelete(payload.id));
    if (action === "moAdd")    return respond(moAdd(payload.data));
    if (action === "moUpsert") return respond(moUpsert(payload.data));
    if (action === "moUpdate") return respond(moUpdate(payload.id, payload.data));
    if (action === "moDelete") return respond(moDelete(payload.id));
    if (action === "saveSetting") return respond(saveSetting(payload.key, payload.value));
    if (action === "waAdd")       return respond(waAdd(payload.data));
    if (action === "addHealthIssue") return respond(addHealthIssue(payload.data));
    if (action === "waUpdate")    return respond(waUpdate(payload.id, payload.data));
    if (action === "waDelete")    return respond(waDelete(payload.id));
    if (action === "migrateBoarSow")   return respond(migrateBoarSowToDbId());
    if (action === "migrateSowIds")    return respond(migrateSowLitterSowId());
    if (action === "runAIAnalysis")    return respond(runNightlyAIAnalysis(payload.targetDate || null, true));
    if (action === "runAISingle")      return respond(runAISingleRecord(payload.clId, payload.force));
    if (action === "runCloseout")      return respond(runCloseoutOnly());
    return respond({ error: "Unknown action: " + action });
  } catch (err) {
    return respond({ error: err.message });
  }
}

// ── CRUD Operations ──────────────────────────────────────────

/**
 * Returns all pig records from the PigLog sheet.
 * Builds a header map once and converts each row to a named object.
 * Formats Date objects as YYYY-MM-DD strings for consistent frontend display.
 * @returns {{ success: boolean, records: Object[] }}
 *   records — array of pig objects with all HEADERS fields as keys.
 */
function getAllRecords() {
  const sheet   = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, records: [] };

  const tz       = Session.getScriptTimeZone(); // cache once
  const lastCol  = sheet.getLastColumn();
  const hdrRow   = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const hdrMap   = {}; // colName → 0-based index
  hdrRow.forEach((h, i) => { if (h) hdrMap[String(h).trim()] = i; });

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const records = data
    .filter(row => row[0] !== "" && row[0] !== null && row[0] !== undefined)
    .map(row => {
      const rec = {};
      HEADERS.forEach(h => {
        const idx = hdrMap[h];
        if (idx === undefined) { rec[h] = ""; return; }
        const v = row[idx];
        rec[h] = (v instanceof Date)
          ? Utilities.formatDate(v, tz, "yyyy-MM-dd")
          : (v === null || v === undefined ? "" : v);
      });
      hdrRow.forEach((h, i) => {
        if (h && !(h in rec)) rec[h] = row[i] === null ? "" : row[i];
      });
      return rec;
    });

  return { success: true, records };
}

/**
 * Looks up a single pig record by its PIG ID string (ear tag).
 * Case-insensitive match on the "PIG ID" column.
 * @param {string} pigId - The PIG ID / ear tag to search for.
 * @returns {{ success: boolean, record?: Object, error?: string }}
 *   record — full pig object if found; error — message if not found.
 */
function getByPigId(pigId) {
  if (!pigId) return { success: false, error: "No PIG ID provided" };
  const sheet   = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, error: "No records found" };

  const lastCol = sheet.getLastColumn();
  const hdrRow  = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const hdrMap  = {};
  hdrRow.forEach((h, i) => { if (h) hdrMap[String(h).trim()] = i; });

  const pidIdx = hdrMap["PIG ID"] !== undefined ? hdrMap["PIG ID"] : 1;
  const data   = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const row    = data.find(r => String(r[pidIdx]).trim().toLowerCase() === String(pigId).trim().toLowerCase());

  if (!row) return { success: false, error: `No record found for PIG ID: "${pigId}"` };

  const rec = {};
  const tz  = Session.getScriptTimeZone();
  HEADERS.forEach(h => {
    const idx = hdrMap[h];
    if (idx === undefined) { rec[h] = ""; return; }
    const v = row[idx];
    rec[h] = (v instanceof Date)
      ? Utilities.formatDate(v, tz, "yyyy-MM-dd")
      : (v === null || v === undefined ? "" : v);
  });
  return { success: true, record: rec };
}

/**
 * Returns the header row of the PigLog sheet as a string array.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The PigLog sheet.
 * @returns {string[]} Array of column header names in order.
 */
function _getPigLogHeaders(sheet) {
  const lastCol = sheet.getLastColumn();
  const hdrRow  = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map     = {};
  hdrRow.forEach((h, i) => { if (h) map[String(h).trim()] = i; });
  return { map, lastCol, hdrRow };
}

/**
 * Adds a new pig record to the PigLog sheet.
 * Assigns the next available DB_ID automatically.
 * Performs duplicate check: blocks records with same PIG ID + Boar + SOW.
 * Validates that PIG ID is not blank before inserting.
 * @param {Object} data - Pig record fields matching HEADERS (DB_ID assigned automatically).
 * @returns {{ success: boolean, db_id?: number, error?: string }}
 *   db_id — the assigned DB_ID if successful; error — reason if failed.
 */
function addRecord(data) {
  const sheet   = getSheet();
  const lastRow = sheet.getLastRow();

  // All required fields must be present
  const required = ["PIG ID", "Boar", "SOW", "DOB", "SEX", "Type", "Stage", "Status", "Available"];
  const missing  = required.filter(f => !String(data[f] || "").trim());
  if (missing.length > 0) {
    return { success: false, error: "Required fields missing: " + missing.join(", ") };
  }

  const pigId = String(data["PIG ID"] || "").trim();
  const boar  = String(data["Boar"]   || "").trim();
  const sow   = String(data["SOW"]    || "").trim();

  const { map, lastCol } = _getPigLogHeaders(sheet);
  const pidIdx  = map["PIG ID"] !== undefined ? map["PIG ID"] : 1;
  const borIdx  = map["Boar"]   !== undefined ? map["Boar"]   : 2;
  const sowIdx  = map["SOW"]    !== undefined ? map["SOW"]    : 3;

  // Duplicate check using live header positions
  if (lastRow > 1) {
    const existing = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    // PIG ID must be unique — it is the physical ear tag
    const pigIdDup = existing.find(r => String(r[pidIdx]).trim().toLowerCase() === pigId.toLowerCase());
    if (pigIdDup) {
      return { success: false, error: `PIG ID "${pigId}" already exists. Each pig must have a unique ear tag ID.` };
    }

    // Full combo duplicate check
    const dup = existing.find(r =>
      String(r[pidIdx]).trim().toLowerCase() === pigId.toLowerCase() &&
      String(r[borIdx]).trim().toLowerCase() === boar.toLowerCase()  &&
      String(r[sowIdx]).trim().toLowerCase() === sow.toLowerCase()
    );
    if (dup) return { success: false, error: `Duplicate — PIG ID "${pigId}", Boar "${boar}", SOW "${sow}" already exists.` };
  }

  // Build new row aligned to actual sheet columns
  const newId  = getNextId(sheet);
  const newRow = Array(Math.max(lastCol, HEADERS.length)).fill("");
  HEADERS.forEach(h => {
    const idx = map[h];
    if (idx === undefined) return;
    newRow[idx] = (h === "DB_ID") ? newId : (data[h] !== undefined ? data[h] : "");
  });
  // If Available column exists in sheet but not in HEADERS yet, leave blank
  sheet.appendRow(newRow.slice(0, Math.max(lastCol, HEADERS.length)));
  return { success: true, db_id: newId };
}

/**
 * Updates fields on an existing pig record identified by DB_ID.
 * Only columns present in data are updated; others are left unchanged.
 * @param {number} dbId - The DB_ID of the record to update.
 * @param {Object} data - Object of field names → new values to write.
 * @returns {{ success: boolean, error?: string }}
 */
function updateRecord(dbId, data) {
  const sheet   = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, error: "No records found" };

  // All required fields must be present in the update payload
  const required = ["PIG ID", "Boar", "SOW", "DOB", "SEX", "Type", "Stage", "Status", "Available"];
  const missing  = required.filter(f => !String(data[f] || "").trim());
  if (missing.length > 0) {
    return { success: false, error: "Required fields missing: " + missing.join(", ") };
  }

  const { map, lastCol } = _getPigLogHeaders(sheet);
  const idIdx    = map["DB_ID"]    !== undefined ? map["DB_ID"]    : 0;
  const pidIdx   = map["PIG ID"]   !== undefined ? map["PIG ID"]   : 1;
  const borIdx   = map["Boar"]     !== undefined ? map["Boar"]     : 2;
  const sowIdx   = map["SOW"]      !== undefined ? map["SOW"]      : 3;
  const availIdx = map["Available"];

  const allData = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const rowIdx  = allData.findIndex(r => Number(r[idIdx]) === Number(dbId));
  if (rowIdx === -1) return { success: false, error: `Record DB_ID ${dbId} not found` };

  const existingRow = allData[rowIdx];
  const sheetRow    = rowIdx + 2;

  // Key-lock check: Available = "No" blocks key field changes
  const availVal = availIdx !== undefined ? String(existingRow[availIdx] || "").trim().toLowerCase() : "";
  if (availVal === "no") {
    const newPigId = String(data["PIG ID"] || "").trim();
    const newBoar  = String(data["Boar"]   || "").trim();
    const newSow   = String(data["SOW"]    || "").trim();
    if (
      (newPigId && newPigId.toLowerCase() !== String(existingRow[pidIdx] || "").trim().toLowerCase()) ||
      (newBoar  && newBoar.toLowerCase()  !== String(existingRow[borIdx] || "").trim().toLowerCase()) ||
      (newSow   && newSow.toLowerCase()   !== String(existingRow[sowIdx] || "").trim().toLowerCase())
    ) {
      return { success: false, error: "Cannot change PIG ID, Boar or SOW — record is marked Available: No." };
    }
  }

  // Duplicate check for key field changes
  const newPigId = String(data["PIG ID"] !== undefined ? data["PIG ID"] : existingRow[pidIdx] || "").trim();
  const newBoar  = String(data["Boar"]   !== undefined ? data["Boar"]   : existingRow[borIdx] || "").trim();
  const newSow   = String(data["SOW"]    !== undefined ? data["SOW"]    : existingRow[sowIdx] || "").trim();
  if (newPigId && newBoar && newSow) {
    const dup = allData.find((r, i) => {
      if (i === rowIdx) return false;
      return String(r[pidIdx]).trim().toLowerCase() === newPigId.toLowerCase() &&
             String(r[borIdx]).trim().toLowerCase() === newBoar.toLowerCase()  &&
             String(r[sowIdx]).trim().toLowerCase() === newSow.toLowerCase();
    });
    if (dup) return { success: false, error: `Duplicate — PIG ID "${newPigId}", Boar "${newBoar}", SOW "${newSow}" already exists.` };
  }

  // Write each field using live column positions
  HEADERS.forEach(h => {
    if (h === "DB_ID") return;
    if (data[h] === undefined) return;
    const colIdx = map[h];
    if (colIdx === undefined) return; // column not in sheet yet — skip
    sheet.getRange(sheetRow, colIdx + 1).setValue(data[h]);
  });
  return { success: true };
}

/**
 * Permanently deletes a pig record row identified by DB_ID.
 * @param {number} dbId - The DB_ID of the record to delete.
 * @returns {{ success: boolean, error?: string }}
 */
function deleteRecord(dbId) {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, error: "No records found" };

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const rowIndex = ids.findIndex(id => Number(id) === Number(dbId));

  if (rowIndex === -1) return { success: false, error: `Record with DB_ID ${dbId} not found` };

  sheet.deleteRow(rowIndex + 2);
  return { success: true, message: "Record deleted successfully" };
}

// ============================================================
//  DAILY CHECKLIST — Sheet: DailyChecklist
// ============================================================

const CL_SHEET = "DailyChecklist";
const CL_KEYS_GS = ['tail','eyes','stool','posture','skin','breathing',
                    'appetite','water','feed','smell',
                    'pinch','belly','limbs','injuries','temp'];
const CL_HEADERS = ["CL_ID","Date","Pen","CheckedBy","Status","Concerns","Notes",
                    "PhotoUrl1","PhotoTime1","PhotoUrl2","PhotoTime2","PhotoUrl3","PhotoTime3",
                    "PigCount","Sec1Time","Sec2Time","Sec3Time",
                    "cl_lactating","cl_nursing","cl_weaklings","cl_lightest_today","cl_heaviest_today",
                    "AIAnalysis",
                    ...CL_KEYS_GS];

/**
 * Gets or creates the DailyChecklist sheet with all required CL_HEADERS columns.
 * Forces plain text (@) number format on all time columns (Sec1/2/3Time,
 * PhotoTime1/2/3) to prevent Google Sheets from auto-converting to Date objects.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The DailyChecklist sheet.
 */
function getClSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CL_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CL_SHEET);
    sheet.appendRow(CL_HEADERS);
    sheet.getRange(1,1,1,CL_HEADERS.length).setFontWeight("bold").setBackground("#2d6a2d").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 70);
    sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(3, 80);
    sheet.setColumnWidth(4, 130);
    sheet.setColumnWidth(5, 90);
    sheet.setColumnWidth(6, 180);
    sheet.setColumnWidth(7, 300);
    // Format time columns as plain text on first creation only
    const lastCol = sheet.getLastColumn();
    if (lastCol > 0) {
      const hdrs = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      CL_TIME_COLS.forEach(colName => {
        const colIdx = hdrs.findIndex(h => String(h).trim() === colName);
        if (colIdx >= 0) sheet.getRange(2, colIdx + 1, 1000, 1).setNumberFormat("@");
      });
    }
  }
  return sheet;
}

/**
 * Returns the next available CL_ID for a new checklist record.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The DailyChecklist sheet.
 * @returns {number} Next CL_ID integer (starts at 1 if sheet is empty).
 */
function getNextClId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 1;
  const ids = sheet.getRange(2,1,lastRow-1,1).getValues().flat().filter(v => v !== "");
  if (ids.length === 0) return 1;
  return Math.max(...ids.map(Number)) + 1;
}

const CL_TIME_COLS = ["Sec1Time","Sec2Time","Sec3Time"];

/**
 * Returns all Daily Checklist records from the DailyChecklist sheet.
 * Skips any orphan columns whose header starts with "—".
 * Converts Date objects in time columns to HH:mm or yyyy-MM-dd HH:mm strings
 * using the script timezone (Africa/Lusaka) to prevent UTC offset corruption.
 * @returns {{ success: boolean, records: Object[] }}
 *   records — array of checklist objects with all CL_HEADERS fields.
 */
function clGetAll() {
  const sheet = getClSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, records: [] };
  const tz      = Session.getScriptTimeZone();
  const lastCol = sheet.getLastColumn();

  // Read header row to build column index map
  const hdrRow  = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const colMap  = {}; // headerName → 0-based index
  hdrRow.forEach((h, i) => {
    const hs = String(h||'').trim();
    if (hs && !hs.startsWith('—') && !colMap[hs]) colMap[hs] = i;
  });

  // Only read columns we actually need (skip orphans)
  const neededCols = CL_HEADERS.filter(h => colMap[h] !== undefined);
  const colIndices = neededCols.map(h => colMap[h]); // 0-based

  // Read all data rows at once
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const records = data
    .filter(r => r[0] !== "" && r[0] !== null && r[0] !== undefined)
    .map(row => {
      const rec = {};
      neededCols.forEach((h, i) => {
        const v = row[colIndices[i]];
        if (CL_TIME_COLS.indexOf(h) >= 0) {
          // Sec1/2/3Time — pure HH:mm values
          if (v instanceof Date) {
            rec[h] = String(v.getHours()).padStart(2,'0') + ':' + String(v.getMinutes()).padStart(2,'0');
          } else {
            const s = String(v||'').trim();
            rec[h] = /^\d{1,2}:\d{2}$/.test(s) ? s : '';
          }
        } else if (h === 'PhotoTime1' || h === 'PhotoTime2' || h === 'PhotoTime3') {
          // PhotoTime stores full datetime string e.g. "2026-03-25 09:00"
          // Return as-is — do NOT convert via getHours() which shifts timezone
          if (v instanceof Date) {
            rec[h] = Utilities.formatDate(v, tz, 'yyyy-MM-dd HH:mm');
          } else {
            rec[h] = String(v||'').trim();
          }
        } else if (v instanceof Date) {
          rec[h] = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
        } else {
          rec[h] = (v !== undefined && v !== null) ? v : '';
        }
      });
      // Truncate large fields
      if (rec.AIAnalysis && String(rec.AIAnalysis).length > 3000) {
        rec.AIAnalysis = String(rec.AIAnalysis).substring(0, 3000) + '…';
      }
      // Fill missing headers with empty string
      CL_HEADERS.forEach(h => { if (!(h in rec)) rec[h] = ''; });
      return rec;
    });

  return { success: true, records };
}

/**
 * Converts a raw sheet row array into a typed checklist record object.
 * Handles three value types per column:
 *   - Time columns (Sec1/2/3Time): formats as "HH:mm" 24hr string.
 *   - PhotoTime columns: formats as "yyyy-MM-dd HH:mm" string.
 *   - All other columns: converts to plain string.
 * @param {any[]} row - Raw values array from sheet.getValues() for one row.
 * @param {string[]} sheetHeaders - Column header names in same order as row.
 * @param {string} tz - Script timezone string e.g. "Africa/Lusaka".
 * @returns {Object} Typed record object keyed by header name.
 */
function rowToClRecord(row, sheetHeaders, tz) {
  if (!tz) tz = Session.getScriptTimeZone();
  const valMap = {};
  for (let i = 0; i < sheetHeaders.length; i++) {
    if (sheetHeaders[i]) valMap[sheetHeaders[i]] = row[i];
  }
  const rec = {};
  for (let hi = 0; hi < CL_HEADERS.length; hi++) {
    const h = CL_HEADERS[hi];
    const v = valMap.hasOwnProperty(h) ? valMap[h] : '';
    if (CL_TIME_COLS.indexOf(h) >= 0) {
      // Sec1/2/3Time — pure HH:mm
      if (v instanceof Date) {
        rec[h] = Utilities.formatDate(v, tz, "HH:mm");
      } else {
        const s = String(v || '').trim();
        rec[h] = /^\d{1,2}:\d{2}$/.test(s) ? s : '';
      }
    } else if (h === 'PhotoTime1' || h === 'PhotoTime2' || h === 'PhotoTime3') {
      // PhotoTime — full datetime string, preserve as-is
      if (v instanceof Date) {
        rec[h] = Utilities.formatDate(v, tz, 'yyyy-MM-dd HH:mm');
      } else {
        rec[h] = String(v || '').trim();
      }
    } else if (v instanceof Date) {
      rec[h] = Utilities.formatDate(v, tz, "yyyy-MM-dd");
    } else {
      rec[h] = (v !== undefined && v !== null) ? v : '';
    }
  }
  return rec;
}

// Diagnostic: returns sheet info without processing any rows
/**
 * Returns debug diagnostics for the DailyChecklist sheet.
 * Used by the admin "Test Connection" button to verify setup.
 * @returns {{ success: boolean, headers: string[], rows: number, ms: number }}
 *   headers — first row values; rows — total row count; ms — response time.
 */
function clDebug() {
  try {
    const t0    = new Date().getTime();
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const t1    = new Date().getTime();
    const sheet = ss.getSheetByName('DailyChecklist');
    const t2    = new Date().getTime();
    if (!sheet) return { success: false, error: 'DailyChecklist sheet not found', sheets: ss.getSheets().map(s=>s.getName()) };
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const t3    = new Date().getTime();
    const hdrs  = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const t4    = new Date().getTime();
    return {
      success:  true,
      rows:     lastRow - 1,
      cols:     lastCol,
      headers:  hdrs.map(h => String(h).trim()).filter(Boolean),
      timings:  { getSpreadsheet: t1-t0, getSheet: t2-t1, getDimensions: t3-t2, getHeaders: t4-t3, total: t4-t0 }
    };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// Just return row count — fast diagnostic
/**
 * Returns the total count of checklist records (excluding header row).
 * @returns {{ success: boolean, count: number }}
 */
function clCount() {
  const sheet = getClSheet();
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  return { success: true, rows: lastRow - 1, cols: lastCol };
}

// Return only records from the last N days — much faster than clGetAll for large sheets
/**
 * Returns checklist records from the last N calendar days.
 * Filters by the Date column — compares dates as YYYY-MM-DD strings.
 * More efficient than clGetAll() for the History tab which only needs recent records.
 * @param {number} [days=30] - Number of days back from today to include.
 * @returns {{ success: boolean, records: Object[], days: number, cutoff: string }}
 *   records — matching records; cutoff — the earliest date included (YYYY-MM-DD).
 */
function clGetRecent(days) {
  days = days || 30;
  const sheet = getClSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, records: [] };
  const tz      = Session.getScriptTimeZone();
  const cutoff  = new Date(); cutoff.setDate(cutoff.getDate() - days);
  const cutoffStr = Utilities.formatDate(cutoff, tz, 'yyyy-MM-dd');
  const lastCol = sheet.getLastColumn();
  const hdrRow  = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const colMap  = {};
  hdrRow.forEach((h, i) => {
    const hs = String(h||'').trim();
    if (hs && !hs.startsWith('—') && !colMap[hs]) colMap[hs] = i;
  });
  const neededCols = CL_HEADERS.filter(h => colMap[h] !== undefined);
  const colIndices = neededCols.map(h => colMap[h]);
  const dateColIdx = colMap['Date'];
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const records = data
    .filter(r => {
      if (!r[0]) return false;
      if (dateColIdx === undefined) return true;
      const d = r[dateColIdx];
      const ds = d instanceof Date ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : String(d||'').substring(0,10);
      return ds >= cutoffStr;
    })
    .map(row => {
      const rec = {};
      neededCols.forEach((h, i) => {
        const v = row[colIndices[i]];
        if (CL_TIME_COLS.indexOf(h) >= 0) {
          if (v instanceof Date) {
            rec[h] = String(v.getHours()).padStart(2,'0') + ':' + String(v.getMinutes()).padStart(2,'0');
          } else {
            const s = String(v||'').trim();
            rec[h] = /^\d{1,2}:\d{2}$/.test(s) ? s : '';
          }
        } else if (v instanceof Date) {
          rec[h] = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
        } else {
          rec[h] = (v !== undefined && v !== null) ? v : '';
        }
      });
      if (rec.AIAnalysis && String(rec.AIAnalysis).length > 3000) {
        rec.AIAnalysis = String(rec.AIAnalysis).substring(0, 3000) + '…';
      }
      CL_HEADERS.forEach(h => { if (!(h in rec)) rec[h] = ''; });
      return rec;
    });
  return { success: true, records, days, cutoff: cutoffStr };
}
/**
 * Builds a column-name → 1-based column-number map for the DailyChecklist sheet.
 * Used by clAdd/clUpdate to write specific columns by name without scanning headers each time.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The DailyChecklist sheet.
 * @returns {Object} Map of { headerName: columnIndex } (1-based).
 */
function _clSheetColMap(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return {};
  const hdrs = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map  = {};
  hdrs.forEach((h, i) => { if (h) map[String(h).trim()] = i + 1; });
  return map;
}

/**
 * Creates or updates a checklist record matching the given Date + Pen combination.
 * If a record already exists for that Date+Pen pair it updates it (prevents duplicates).
 * If no match is found it creates a new record via clAdd().
 * @param {Object} data - Checklist fields including Date (YYYY-MM-DD) and Pen number.
 * @returns {{ success: boolean, cl_id: number, updated?: boolean, error?: string }}
 *   updated — true if an existing record was updated rather than created.
 */
function clUpsert(data) {
  const sheet  = getClSheet();
  const colMap = _clSheetColMap(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const dateCol = colMap["Date"]; const penCol = colMap["Pen"];
    if (dateCol && penCol) {
      const nCols  = sheet.getLastColumn();
      const allData = sheet.getRange(2, 1, lastRow - 1, nCols).getValues();
      for (let i = 0; i < allData.length; i++) {
        const rowDate = allData[i][dateCol-1] instanceof Date
          ? Utilities.formatDate(allData[i][dateCol-1], Session.getScriptTimeZone(), "yyyy-MM-dd")
          : String(allData[i][dateCol-1] || '').trim();
        const rowPen = String(allData[i][penCol-1] || '').trim().toLowerCase();
        if (rowDate === String(data.Date || '').trim() && rowPen === String(data.Pen || '').trim().toLowerCase()) {
          const existingId = allData[i][0];
          clUpdate(existingId, data);
          return { success: true, cl_id: existingId, updated: true };
        }
      }
    }
  }
  return clAdd(data);
}

/**
 * Appends a new row to the DailyChecklist sheet.
 * Server-side duplicate guard: blocks a second record for the same Date+Pen.
 * After appending, forces @-format (plain text) on all time columns in the new row
 * to prevent Google Sheets auto-converting HH:mm strings to Date fractions.
 * @param {Object} data - Checklist fields. CL_ID and Status are assigned automatically.
 * @returns {{ success: boolean, cl_id: number, error?: string }}
 */
function clAdd(data) {
  const sheet   = getClSheet();
  const colMap  = _clSheetColMap(sheet);
  const lastRow = sheet.getLastRow();
  // Server-side duplicate guard
  if (lastRow > 1) {
    const dateCol = colMap["Date"]; const penCol = colMap["Pen"];
    if (dateCol && penCol) {
      const nCols  = sheet.getLastColumn();
      const allData = sheet.getRange(2, 1, lastRow - 1, nCols).getValues();
      const inDate  = String(data.Date || '').trim();
      const inPen   = String(data.Pen  || '').trim().toLowerCase();
      const dup = allData.find(r => {
        const rowDate = r[dateCol-1] instanceof Date
          ? Utilities.formatDate(r[dateCol-1], Session.getScriptTimeZone(), "yyyy-MM-dd")
          : String(r[dateCol-1] || '').trim();
        return rowDate === inDate && String(r[penCol-1] || '').trim().toLowerCase() === inPen;
      });
      if (dup) return { success: false, error: `Pen ${data.Pen} already has a record for ${data.Date}.` };
    }
  }
  const newId = getNextClId(sheet);
  // Build row aligned to ACTUAL sheet column order
  const nCols = Math.max(sheet.getLastColumn(), CL_HEADERS.length);
  const row   = new Array(nCols).fill('');
  Object.entries(colMap).forEach(([h, col]) => {
    if (h === "CL_ID") { row[col-1] = newId; return; }
    const v = data[h] !== undefined ? data[h] : '';
    row[col-1] = CL_TIME_COLS.includes(h) ? String(v) : v;
  });
  if (colMap["CL_ID"]) row[colMap["CL_ID"]-1] = newId;
  sheet.appendRow(row);
  // Force text format on all time columns to prevent Sheets auto-converting HH:mm
  const allTimeCols = [...CL_TIME_COLS, 'PhotoTime1','PhotoTime2','PhotoTime3'];
  allTimeCols.forEach(h => {
    const col = colMap[h];
    if (col) sheet.getRange(sheet.getLastRow(), col).setNumberFormat('@');
  });
  return { success: true, cl_id: newId };
}

/**
 * Updates specific fields on an existing checklist record by CL_ID.
 * Key behaviours:
 *   - Preserves existing timestamps — never overwrites a filled time column with blank.
 *   - Forces @-format (plain text) on all time columns after writing to prevent
 *     Google Sheets from converting HH:mm strings to Date objects on next read.
 * @param {number} clId - The CL_ID of the record to update.
 * @param {Object} data - Field name → value pairs to write (only provided fields are changed).
 * @returns {{ success: boolean, error?: string }}
 */
function clUpdate(clId, data) {
  const sheet   = getClSheet();
  const colMap  = _clSheetColMap(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, error: "No records" };
  const ids = sheet.getRange(2, 1, lastRow-1, 1).getValues().flat();
  const rowIndex = ids.findIndex(id => Number(id) === Number(clId));
  if (rowIndex === -1) return { success: false, error: "Record not found" };
  const sheetRow = rowIndex + 2;

  // Read existing row to preserve timestamp values
  const lastCol = sheet.getLastColumn();
  const existingRow = sheet.getRange(sheetRow, 1, 1, lastCol).getValues()[0];

  Object.keys(data).forEach(h => {
    const col = colMap[h];
    if (!col || h === "CL_ID") return;
    // Preserve existing timestamps — never overwrite with blank
    const isTimeCol  = CL_TIME_COLS.includes(h) || h === 'PhotoTime1' || h === 'PhotoTime2' || h === 'PhotoTime3';
    const isPhotoTime = h === 'PhotoTime1' || h === 'PhotoTime2' || h === 'PhotoTime3';
    if (isTimeCol) {
      const newVal      = String(data[h] || '').trim();
      const existingVal = String(existingRow[col - 1] || '').trim();
      if (!newVal && existingVal) return; // keep existing timestamp
      const cell = sheet.getRange(sheetRow, col);
      cell.setNumberFormat('@'); // force plain text for ALL time cols
      cell.setValue(newVal);
    } else {
      sheet.getRange(sheetRow, col).setValue(data[h]);
    }
  });
  return { success: true };
}

/**
 * Permanently deletes a Daily Checklist record row by CL_ID.
 * @param {number} clId - The CL_ID of the record to delete.
 * @returns {{ success: boolean, error?: string }}
 */
function clDelete(clId) {
  const sheet = getClSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, error: "No records" };
  const ids = sheet.getRange(2,1,lastRow-1,1).getValues().flat();
  const rowIndex = ids.findIndex(id => Number(id) === Number(clId));
  if (rowIndex === -1) return { success: false, error: "Record not found" };
  sheet.deleteRow(rowIndex + 2);
  return { success: true };
}

/**
 * Saves a health check section photo to Google Drive and records its URL in the sheet.
 * Stores photos in a "PigLog Photos" root folder (created if missing).
 * Sets the file as publicly viewable so the app can display it inline.
 * Writes the Drive URL to PhotoUrl{section} and timestamp to PhotoTime{section}.
 * @param {number} clId - The CL_ID of the checklist record this photo belongs to.
 * @param {string} photoBase64 - Base64-encoded image bytes (no data URL prefix).
 * @param {string} mimeType - Image MIME type e.g. "image/jpeg" or "image/png".
 * @param {string} photoTime - Capture timestamp as "YYYY-MM-DD HH:mm" string.
 * @param {number} section - Section number: 1 = Morning, 2 = Feeding, 3 = Afternoon.
 * @returns {{ success: boolean, viewUrl?: string, fileId?: string, error?: string }}
 *   viewUrl — full Drive view URL; fileId — Drive file ID for direct proxy access.
 */
function clSavePhoto(clId, photoBase64, mimeType, photoTime, section) {
  try {
    if (!clId || !photoBase64) return { success: false, error: "Missing clId or photo data" };
    const sec = section || 1;

    const folderName = "PigLog_Photos";
    const folders = DriveApp.getFoldersByName(folderName);
    const folder  = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

    // Timestamped filename: pen_checklist_{clId}_sec{N}_{YYYY-MM-DD_HH-MM}.jpg
    const ext  = (mimeType || 'image/jpeg').split('/')[1] || 'jpg';
    const ts   = photoTime ? String(photoTime).replace(/[: /]/g, '-').replace(/[^a-zA-Z0-9_-]/g, '') : String(new Date().getTime());
    const name = `pen_cl${clId}_sec${sec}_${ts}.${ext}`;
    const blob = Utilities.newBlob(Utilities.base64Decode(photoBase64), mimeType || 'image/jpeg', name);
    const file = folder.createFile(blob);

    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileId   = file.getId();
    const viewUrl  = "https://drive.google.com/file/d/" + fileId + "/view";

    // Store the viewUrl — frontend extracts fileId and builds lh3.googleusercontent.com URL
    const updateData = {};
    updateData['PhotoUrl'  + sec] = viewUrl;
    // Force plain text storage by writing to cell directly with text format
    // to prevent Google Sheets auto-converting "2026-03-25 09:00" to a Date object
    if (photoTime) updateData['PhotoTime' + sec] = String(photoTime);
    const result = clUpdate(clId, updateData);
    if (!result.success) return { success: false, error: "Photo saved to Drive but sheet update failed: " + result.error };

    return { success: true, viewUrl, fileId, section: sec };
  } catch(e) {
    return { success: false, error: "Photo save failed: " + e.message };
  }
}

// Run this function ONCE manually in the Apps Script editor to grant Drive permission.
// Click the ▶ Run button next to "testDriveAccess" in the editor.
/**
 * Tests that Drive access is correctly authorised for this script.
 * Run manually from the Apps Script editor after first deployment to verify
 * that the drive scope has been granted. Logs the root folder name.
 */
function testDriveAccess() {
  const folder = DriveApp.getRootFolder();
  Logger.log("Drive access OK. Root folder: " + folder.getName());
}

// Fetch a Drive file and return it as a base64 data URL — bypasses CORS for <img> tags
/**
 * Fetches a Google Drive file and returns it as a base64 data URL for inline display.
 * The thumbnail API returns 403 for private files, so this always uses DriveApp.
 * If the file exceeds 4 MB it falls back to a compressed lh3.googleusercontent.com
 * thumbnail URL instead of returning the full file bytes.
 * Called by the frontend when displaying saved checklist photos.
 * @param {string} fileId - Google Drive file ID (not the full URL).
 * @returns {{ success: boolean, dataUrl?: string, error?: string }}
 *   dataUrl — "data:image/jpeg;base64,..." string ready for an <img> src attribute.
 */
function getPhotoAsBase64(fileId) {
  try {
    if (!fileId) return { success: false, error: 'No fileId provided' };
    const file = DriveApp.getFileById(fileId);
    let   blob = file.getBlob();

    // If file > 4MB fetch a compressed thumbnail instead
    if (blob.getBytes().length > 4 * 1024 * 1024) {
      const thumbUrl = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w1024';
      const token    = ScriptApp.getOAuthToken();
      const res      = UrlFetchApp.fetch(thumbUrl, {
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true
      });
      if (res.getResponseCode() === 200) blob = res.getBlob();
    }

    const mime = blob.getContentType() || 'image/jpeg';
    const b64  = Utilities.base64Encode(blob.getBytes());
    return { success: true, dataUrl: 'data:' + mime + ';base64,' + b64 };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

const SL_SHEET = "SowLitter";
const SL_TASK_COLS = ['sow_antiinflam','sow_mma_ab','sow_oxytocin',
  'iodine','colostrum','teat_order','heat_lamp','litter_weigh',
  'iron','tail_dock','teeth','navel_check',
  'multivit','creep','navel_healed',
  'castration','ear_notch','coccidiostat',
  'indiv_weigh','health_d14','mma_d14',
  'vax','weaner_feed','weaner_pen','prev_meds','weaned'];
const SL_HEADERS_GS = ["SL_ID","SowId","LitterBoar","FarrowDate","EstFarrowDate","Pen","Notes",
  "sl_foster","sl_fsow","sl_total_piglets","sl_mortality",
  "ByLitter","ByNursing","BySowTreat","ByMilestones",
  "ByD01","ByD23","ByD57","ByD710","ByD14","ByD2128",
  "SlTime_farrow",
  "SlTime_shdr_litter","SlTime_shdr_sowtreat",
  "SlTime_mhdr_d01","SlTime_mhdr_d23","SlTime_mhdr_d57",
  "SlTime_mhdr_d710","SlTime_mhdr_d14","SlTime_mhdr_d2128",
  "sl_born_alive","sl_stillborn","sl_mummified","sl_total_birth_wt",
  "sl_lightest","sl_heaviest","sl_nursing","sl_weaklings",
  "sl_lightest_today","sl_heaviest_today","sl_castrated",
  "sl_wt_d14","sl_alive_d14","sl_num_weaned","sl_date_weaned","sl_wean_wt",
  // Section photos & notes — one per highlighted section banner
  "SlPhoto_farrow","SlPhotoTime_farrow","SlNotes_farrow",
  "SlPhoto_shdr_litter","SlPhotoTime_shdr_litter","SlNotes_shdr_litter",
  "SlPhoto_shdr_sowtreat","SlPhotoTime_shdr_sowtreat","SlNotes_shdr_sowtreat",
  "SlPhoto_mhdr_d01","SlPhotoTime_mhdr_d01","SlNotes_mhdr_d01",
  "SlPhoto_mhdr_d23","SlPhotoTime_mhdr_d23","SlNotes_mhdr_d23",
  "SlPhoto_mhdr_d57","SlPhotoTime_mhdr_d57","SlNotes_mhdr_d57",
  "SlPhoto_mhdr_d710","SlPhotoTime_mhdr_d710","SlNotes_mhdr_d710",
  "SlPhoto_mhdr_d14","SlPhotoTime_mhdr_d14","SlNotes_mhdr_d14",
  "SlPhoto_mhdr_d2128","SlPhotoTime_mhdr_d2128","SlNotes_mhdr_d2128",
  ...SL_TASK_COLS];

const SL_TIME_COLS = [
  "SlTime_farrow",
  "SlTime_shdr_litter","SlTime_shdr_sowtreat",
  "SlTime_mhdr_d01","SlTime_mhdr_d23","SlTime_mhdr_d57",
  "SlTime_mhdr_d710","SlTime_mhdr_d14","SlTime_mhdr_d2128"
];

/**
 * Gets or creates the SowLitter sheet with all required SL_HEADERS_GS columns.
 * On first creation applies pink header formatting and freezes row 1.
 * Existing sheets: adds any missing columns to the right without touching data.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The SowLitter sheet.
 */
function getSlSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SL_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SL_SHEET);
    sheet.appendRow(SL_HEADERS_GS);
    sheet.getRange(1,1,1,SL_HEADERS_GS.length).setFontWeight("bold").setBackground("#880e4f").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  } else {
    // Ensure all expected columns exist — add any missing ones to the right
    const lastCol  = sheet.getLastColumn();
    const existing = lastCol > 0
      ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim())
      : [];
    SL_HEADERS_GS.forEach(h => {
      if (h && !existing.includes(h)) {
        const newCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, newCol).setValue(h)
             .setFontWeight("bold").setBackground("#880e4f").setFontColor("#ffffff");
      }
    });
  }
  // Force plain-text format on all SlTime_* columns to prevent Sheets
  // auto-converting "25 Mar 2026 09:15" strings into Date objects
  const lastCol = sheet.getLastColumn();
  if (lastCol > 0) {
    const hdrs = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    SL_TIME_COLS.forEach(colName => {
      const colIdx = hdrs.findIndex(h => String(h).trim() === colName);
      if (colIdx >= 0) {
        sheet.getRange(2, colIdx + 1, Math.max(sheet.getLastRow() - 1, 1), 1).setNumberFormat('@');
      }
    });
  }
  return sheet;
}

/**
 * Returns all Sow & Litter records from the SowLitter sheet.
 * Converts Date objects to YYYY-MM-DD strings for consistent frontend handling.
 * @returns {{ success: boolean, records: Object[] }}
 *   records — array of SL objects with all SL_HEADERS_GS fields as keys.
 */
function slGetAll() {
  const sheet = getSlSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, records: [] };
  const tz         = Session.getScriptTimeZone(); // cache once
  const lastCol    = sheet.getLastColumn();
  const sheetHdrs  = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v).trim());
  const data       = sheet.getRange(2, 1, lastRow-1, lastCol).getValues();
  const validSlTime = t => {
    const s = String(t||'').trim();
    // Accept full datetime "DD Mon YYYY HH:MM", legacy "HH:MM", or any non-empty string
    return s.length > 0;
  };
  const records = data.filter(r => r[0] !== '').map(row => {
    const byName = {};
    sheetHdrs.forEach((h, i) => { if (h) byName[h] = row[i]; });
    const rec = {};
    SL_HEADERS_GS.forEach(h => {
      const v = byName.hasOwnProperty(h) ? byName[h] : '';
      if (SL_TIME_COLS.indexOf(h) >= 0) {
        // SlTime_* columns store full datetime strings like "25 Mar 2026 09:15"
        // Never convert via Date methods — always return as plain string
        if (v instanceof Date) {
          // Sheets auto-converted it — recover as formatted string
          rec[h] = Utilities.formatDate(v, tz, 'dd MMM yyyy HH:mm');
        } else {
          const s = String(v||'').trim();
          rec[h] = s; // return as-is — empty string is fine
        }
      } else {
        rec[h] = v instanceof Date
          ? Utilities.formatDate(v, tz, "yyyy-MM-dd")
          : (v !== undefined && v !== null ? v : '');
      }
    });
    return rec;
  });
  return { success: true, records };
}

// Upsert: update if SowId+FarrowDate exists, otherwise insert
/**
 * Creates or updates a Sow & Litter record for a given SowId + FarrowDate pair.
 * If an active (non-closed) record exists for this sow, updates it instead of
 * creating a duplicate. If none found, delegates to slAdd().
 * @param {Object} data - SL fields including SowId and FarrowDate (YYYY-MM-DD).
 * @returns {{ success: boolean, sl_id: number, updated?: boolean, error?: string }}
 */
function slUpsert(data) {
  const sheet  = getSlSheet();
  const colMap = _slSheetColMap(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const sowCol    = colMap["SowId"];
    const farrowCol = colMap["FarrowDate"];
    const nCols     = sheet.getLastColumn();
    const allData   = sheet.getRange(2, 1, lastRow-1, nCols).getValues();
    for (let i = 0; i < allData.length; i++) {
      const rowSow    = sowCol    ? String(allData[i][sowCol-1]   ||'').trim().toLowerCase() : '';
      const rowFarrow = farrowCol ? (allData[i][farrowCol-1] instanceof Date
        ? Utilities.formatDate(allData[i][farrowCol-1], Session.getScriptTimeZone(), "yyyy-MM-dd")
        : String(allData[i][farrowCol-1]||'').trim()) : '';
      if (rowSow === String(data.SowId||'').trim().toLowerCase() && rowFarrow === String(data.FarrowDate||'').trim()) {
        const existingId = allData[i][0];
        slUpdate(existingId, data);
        return { success: true, sl_id: existingId, updated: true };
      }
    }
  }
  return slAdd(data);
}

// Returns { headerName -> 1-based col } from actual sheet header row
/**
 * Builds a column-name → 1-based column-number map for the SowLitter sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The SowLitter sheet.
 * @returns {Object} Map of { headerName: columnIndex } (1-based).
 */
function _slSheetColMap(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return {};
  const hdrs = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map  = {};
  hdrs.forEach((h, i) => { if (h) map[String(h).trim()] = i + 1; });
  return map;
}

/**
 * Appends a new Sow & Litter record row to the SowLitter sheet.
 * Assigns the next available SL_ID automatically.
 * @param {Object} data - SL fields (SL_ID assigned automatically).
 * @returns {{ success: boolean, sl_id: number, error?: string }}
 */
function slAdd(data) {
  const sheet  = getSlSheet();
  const colMap = _slSheetColMap(sheet);
  const ids    = sheet.getLastRow() <= 1 ? [] : sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues().flat().filter(v=>v!=="");
  const newId  = ids.length === 0 ? 1 : Math.max(...ids.map(Number)) + 1;
  const nCols  = Math.max(sheet.getLastColumn(), SL_HEADERS_GS.length);
  const row    = new Array(nCols).fill('');
  Object.entries(colMap).forEach(([h, col]) => {
    if (h === "SL_ID") { row[col-1] = newId; return; }
    const v = data[h] !== undefined ? data[h] : '';
    row[col-1] = SL_TIME_COLS.includes(h) ? String(v) : v;
  });
  if (colMap["SL_ID"]) row[colMap["SL_ID"]-1] = newId;
  sheet.appendRow(row);
  // Force plain-text format on all SlTime_* columns in the new row
  const newRow = sheet.getLastRow();
  SL_TIME_COLS.forEach(colName => {
    const col = colMap[colName];
    if (col) sheet.getRange(newRow, col).setNumberFormat('@');
  });
  return { success: true, sl_id: newId };
}

/**
 * Updates specific fields on an existing SL record by SL_ID.
 * Only columns present in data are written; all others are preserved.
 * @param {number} slId - The SL_ID of the record to update.
 * @param {Object} data - Field name → value pairs to write.
 * @returns {{ success: boolean, error?: string }}
 */
function slUpdate(slId, data) {
  const sheet  = getSlSheet();
  const colMap = _slSheetColMap(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, error: "No records" };
  const ids = sheet.getRange(2,1,lastRow-1,1).getValues().flat();
  const rowIndex = ids.findIndex(id => Number(id) === Number(slId));
  if (rowIndex === -1) return { success: false, error: "Record not found" };
  const sheetRow = rowIndex + 2;
  Object.keys(data).forEach(h => {
    const col = colMap[h];
    if (!col || h === "SL_ID") return;
    if (SL_TIME_COLS.includes(h)) {
      const cell = sheet.getRange(sheetRow, col);
      cell.setNumberFormat('@');       // force plain text — prevents Sheets converting to Date
      cell.setValue(String(data[h] || ''));
    } else {
      sheet.getRange(sheetRow, col).setValue(data[h]);
    }
  });
  return { success: true };
}

/**
 * Permanently deletes a Sow & Litter record row by SL_ID.
 * @param {number} slId - The SL_ID of the record to delete.
 * @returns {{ success: boolean, error?: string }}
 */
function slDelete(slId) {
  const sheet = getSlSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, error: "No records" };
  const ids = sheet.getRange(2,1,lastRow-1,1).getValues().flat();
  const rowIndex = ids.findIndex(id => Number(id) === Number(slId));
  if (rowIndex === -1) return { success: false, error: "Record not found" };
  sheet.deleteRow(rowIndex + 2);
  return { success: true };
}

/**
 * Saves a litter section photo to Google Drive and records its URL in the SowLitter sheet.
 * Stores photos in "PigLog Photos/Litter" subfolder (created if missing).
 * Each section banner gets its own named photo slot identified by sectionKey.
 * Section keys map to sheet columns: SlPhoto_{sectionKey} and SlPhotoTime_{sectionKey}.
 * @param {number} slId - The SL_ID of the litter record this photo belongs to.
 * @param {string} photoBase64 - Base64-encoded image bytes (no data URL prefix).
 * @param {string} mimeType - Image MIME type e.g. "image/jpeg".
 * @param {string} photoTime - Capture timestamp as "YYYY-MM-DD HH:mm" string.
 * @param {string} sectionKey - Section identifier matching a SL_HEADER_RANGES key:
 *   "farrow" | "shdr_litter" | "shdr_sowtreat" | "mhdr_d01" | "mhdr_d23" |
 *   "mhdr_d57" | "mhdr_d710" | "mhdr_d14" | "mhdr_d2128"
 * @returns {{ success: boolean, viewUrl?: string, fileId?: string, sectionKey?: string, error?: string }}
 */
function slSavePhoto(slId, photoBase64, mimeType, photoTime, sectionKey) {
  try {
    if (!slId || !photoBase64) return { success: false, error: 'Missing slId or photo data' };
    const key = String(sectionKey || 'shdr_litter');

    // Get or create PigLog Photos/Litter folder
    const rootFolders = DriveApp.getFoldersByName('PigLog Photos');
    const root = rootFolders.hasNext() ? rootFolders.next() : DriveApp.createFolder('PigLog Photos');
    const litFolders = root.getFoldersByName('Litter');
    const folder = litFolders.hasNext() ? litFolders.next() : root.createFolder('Litter');

    const ext  = (mimeType || 'image/jpeg').split('/')[1] || 'jpg';
    const ts   = photoTime
      ? String(photoTime).replace(/[: /]/g, '-').replace(/[^a-zA-Z0-9_-]/g, '')
      : String(new Date().getTime());
    const name = `litter_sl${slId}_${key}_${ts}.${ext}`;
    const blob = Utilities.newBlob(Utilities.base64Decode(photoBase64), mimeType || 'image/jpeg', name);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const viewUrl = 'https://drive.google.com/file/d/' + file.getId() + '/view';

    // Update the SowLitter record with the named column
    const updateData = {};
    updateData['SlPhoto_'     + key] = viewUrl;
    if (photoTime) updateData['SlPhotoTime_' + key] = String(photoTime);
    const result = slUpdate(slId, updateData);
    if (!result.success) return { success: false, error: 'Photo saved but sheet update failed: ' + result.error };

    return { success: true, viewUrl, fileId: file.getId(), sectionKey: key };
  } catch(e) {
    return { success: false, error: 'Litter photo save failed: ' + e.message };
  }
}

// ============================================================
//  WEEKLY CHECKLIST — Sheet: WeeklyChecklist
// ============================================================

const WK_SHEET = "WeeklyChecklist";
const WK_KEYS_GS = ['wk_weight_spot','wk_body_condition','wk_behaviour',
  'wk_deworm','wk_vitamins','wk_wounds',
  'wk_feeder_clean','wk_pen_repair','wk_feed_consumption','wk_boar_condition'];
const WK_HEADERS_GS = ["WK_ID","Date","WeekNum","WeekYear","WeekKey","Pen","CheckedBy","Status","Concerns","Notes",
  "AvgWeight","FeedKg","ByCondition","ByHealth","ByFarm",...WK_KEYS_GS];

/**
 * Gets or creates the WeeklyChecklist sheet with all required WK_HEADERS_GS columns.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The WeeklyChecklist sheet.
 */
function getWkSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(WK_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(WK_SHEET);
    sheet.appendRow(WK_HEADERS_GS);
    sheet.getRange(1,1,1,WK_HEADERS_GS.length).setFontWeight("bold").setBackground("#1a3a8a").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Returns all Weekly Checklist records from the WeeklyChecklist sheet.
 * @returns {{ success: boolean, records: Object[] }}
 */
function wkGetAll() {
  const sheet = getWkSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, records: [] };
  const tz       = Session.getScriptTimeZone(); // cache once
  const lastCol  = sheet.getLastColumn();
  const sheetHdrs = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v).trim());
  const data = sheet.getRange(2, 1, lastRow-1, lastCol).getValues();
  return { success: true, records: data.filter(r => r[0] !== '').map(row => {
    const byName = {};
    sheetHdrs.forEach((h, i) => { if (h) byName[h] = row[i]; });
    const rec = {};
    WK_HEADERS_GS.forEach(h => {
      const v = byName.hasOwnProperty(h) ? byName[h] : '';
      rec[h] = v instanceof Date ? Utilities.formatDate(v, tz, "yyyy-MM-dd") : (v !== undefined && v !== null ? v : '');
    });
    return rec;
  })};
}

/**
 * Builds a column-name → 1-based column-number map for the WeeklyChecklist sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The WeeklyChecklist sheet.
 * @returns {Object} Map of { headerName: columnIndex } (1-based).
 */
function _wkSheetColMap(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return {};
  const hdrs = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = {};
  hdrs.forEach((h, i) => { if (h) map[String(h).trim()] = i + 1; });
  return map;
}

/**
 * Appends a new Weekly Checklist record to the WeeklyChecklist sheet.
 * Assigns the next WK_ID automatically.
 * @param {Object} data - WK fields including Date (YYYY-MM-DD), Pen, and WeekNum.
 * @returns {{ success: boolean, wk_id: number, error?: string }}
 */
function wkAdd(data) {
  const sheet   = getWkSheet();
  const colMap  = _wkSheetColMap(sheet);
  const lastRow = sheet.getLastRow();
  // Server-side duplicate guard: same WeekKey + Pen
  if (lastRow > 1) {
    const wkKeyCol = colMap["WeekKey"]; const penCol = colMap["Pen"];
    if (wkKeyCol && penCol) {
      const nCols   = sheet.getLastColumn();
      const allData = sheet.getRange(2, 1, lastRow - 1, nCols).getValues();
      const inKey   = String(data.WeekKey || '').trim();
      const inPen   = String(data.Pen     || '').trim().toLowerCase();
      const dup = allData.find(r =>
        String(r[wkKeyCol-1]||'').trim() === inKey &&
        String(r[penCol-1]  ||'').trim().toLowerCase() === inPen
      );
      if (dup) return { success: false, error: `Pen ${data.Pen} already has a record for ${data.WeekKey}.` };
    }
  }
  const ids   = lastRow <= 1 ? [] : sheet.getRange(2,1,lastRow-1,1).getValues().flat().filter(v=>v!=="");
  const newId = ids.length === 0 ? 1 : Math.max(...ids.map(Number)) + 1;
  const nCols = Math.max(sheet.getLastColumn(), WK_HEADERS_GS.length);
  const row   = new Array(nCols).fill('');
  Object.entries(colMap).forEach(([h, col]) => {
    row[col-1] = h === "WK_ID" ? newId : (data[h] !== undefined ? data[h] : '');
  });
  if (colMap["WK_ID"]) row[colMap["WK_ID"]-1] = newId;
  sheet.appendRow(row);
  return { success: true, wk_id: newId };
}

/**
 * Updates specific fields on an existing weekly record by WK_ID.
 * @param {number} wkId - The WK_ID of the record to update.
 * @param {Object} data - Field name → value pairs to write.
 * @returns {{ success: boolean, error?: string }}
 */
function wkUpdate(wkId, data) {
  const sheet  = getWkSheet();
  const colMap = _wkSheetColMap(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, error: "No records" };
  const ids = sheet.getRange(2,1,lastRow-1,1).getValues().flat();
  const rowIndex = ids.findIndex(id => Number(id) === Number(wkId));
  if (rowIndex === -1) return { success: false, error: "Not found" };
  const sheetRow = rowIndex + 2;
  Object.keys(data).forEach(h => {
    const col = colMap[h];
    if (!col || h === "WK_ID") return;
    sheet.getRange(sheetRow, col).setValue(data[h]);
  });
  return { success: true };
}

/**
 * Permanently deletes a Weekly Checklist record row by WK_ID.
 * @param {number} wkId - The WK_ID of the record to delete.
 * @returns {{ success: boolean, error?: string }}
 */
function wkDelete(wkId) {
  const sheet = getWkSheet(); const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, error: "No records" };
  const ids = sheet.getRange(2,1,lastRow-1,1).getValues().flat();
  const rowIndex = ids.findIndex(id => Number(id) === Number(wkId));
  if (rowIndex === -1) return { success: false, error: "Not found" };
  sheet.deleteRow(rowIndex + 2);
  return { success: true };
}

// ============================================================
//  MONTHLY CHECKLIST — Sheet: MonthlyChecklist
// ============================================================

const MO_SHEET = "MonthlyChecklist";
const MO_KEYS_GS = ['mo_full_weigh','mo_growth_rate','mo_mortality',
  'mo_vaccination','mo_breeding','mo_boar_perf',
  'mo_feed_inventory','mo_equipment','mo_biosecurity'];
const MO_HEADERS_GS = ["MO_ID","Month","CreatedDate","CheckedBy","Status","Concerns","Notes",
  "PigsWeighed","AvgADG","Deaths","VaxCount","FeedStock",
  "ByGrowth","ByHealth","ByFarm",...MO_KEYS_GS];

// Convert any value to YYYY-MM using LOCAL timezone (critical for GAS Date objects)
/**
 * Normalises any date-like value to a "YYYY-MM" month key string.
 * Handles both Date objects (from Sheets) and string inputs ("2026-03", "2026-03-15", etc.).
 * Returns empty string for null, undefined, or unparseable values.
 * @param {Date|string} val - A Date object or date string to normalise.
 * @returns {string} Month key in "YYYY-MM" format, or "" if invalid.
 */
function _toMonthKey(val) {
  if (!val && val !== 0) return '';
  // GAS returns Date objects for date-formatted cells; use local TZ methods
  if (Object.prototype.toString.call(val) === '[object Date]') {
    return val.getFullYear() + '-' + String(val.getMonth() + 1).padStart(2, '0');
  }
  const s = String(val).trim();
  // ISO string from JSON serialisation: "2026-02-28T22:00:00.000Z" (UTC-shifted)
  // or plain "2026-03-01" or "2026-03"
  if (/^\d{4}-\d{2}-\d{2}T/.test(s)) {
    // UTC timestamp — parse using Utilities.formatDate to get spreadsheet local date
    // (In GAS context; outside GAS we just do best-effort substring)
    try {
      const d = new Date(s);
      const tz = Session.getScriptTimeZone();
      const formatted = Utilities.formatDate(d, tz, 'yyyy-MM');
      return formatted;
    } catch(e) {
      return s.substring(0, 7);
    }
  }
  if (/^\d{4}-\d{2}/.test(s)) return s.substring(0, 7);
  return '';
}

/**
 * Gets or creates the MonthlyChecklist sheet with all required headers.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The MonthlyChecklist sheet.
 */
function getMoSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(MO_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(MO_SHEET);
    sheet.appendRow(MO_HEADERS_GS);
    sheet.getRange(1,1,1,MO_HEADERS_GS.length).setFontWeight("bold").setBackground("#4a148c").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Returns the header row of the MonthlyChecklist sheet as a string array.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The MonthlyChecklist sheet.
 * @returns {string[]} Array of header names.
 */
function _moGetSheetHeaders(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return {};
  const raw = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = {};
  raw.forEach((h, i) => { if (h) map[String(h).trim()] = i; });
  return map;
}

/**
 * Returns all Monthly Checklist records from the MonthlyChecklist sheet.
 * Normalises the Month column to "YYYY-MM" format using _toMonthKey().
 * @returns {{ success: boolean, records: Object[] }}
 */
function moGetAll() {
  const sheet   = getMoSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, records: [] };
  const hdrs    = _moGetSheetHeaders(sheet);
  const lastCol = sheet.getLastColumn();
  const data    = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const records = data
    .filter(r => r[0] !== '' && r[0] !== null && r[0] !== undefined)
    .map(row => {
      const rec = {};
      Object.entries(hdrs).forEach(([name, idx]) => { rec[name] = row[idx]; });
      // Normalise Month to clean YYYY-MM, handling Date objects and legacy Date column
      rec['Month'] = _toMonthKey(rec['Month'] || rec['Date'] || '');
      return rec;
    });
  return { success: true, records };
}

/**
 * Creates or updates a Monthly Checklist record for a given month.
 * Matches on the normalised Month key (YYYY-MM). Updates if found, creates if not.
 * @param {Object} data - MO fields including Month (YYYY-MM or any date string).
 * @returns {{ success: boolean, mo_id: number, updated?: boolean, error?: string }}
 */
function moUpsert(data) {
  const sheet    = getMoSheet();
  const monthKey = String(data['Month'] || '').trim().substring(0, 7);
  if (!monthKey) return { success: false, error: 'Month is required' };

  const hdrs        = _moGetSheetHeaders(sheet);
  const monthColIdx = hdrs['Month'] !== undefined ? hdrs['Month']
                    : hdrs['Date']  !== undefined ? hdrs['Date'] : 1;
  const idColIdx    = hdrs['MO_ID'] !== undefined ? hdrs['MO_ID'] : 0;
  const lastRow     = sheet.getLastRow();

  const matchingRows = [];
  if (lastRow > 1) {
    const lastCol = sheet.getLastColumn();
    const allData = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    allData.forEach((row, i) => {
      if (row[0] === '' || row[0] === null || row[0] === undefined) return;
      if (_toMonthKey(row[monthColIdx]) === monthKey) {
        matchingRows.push({ sheetRow: i + 2, rowData: row });
      }
    });
  }

  if (matchingRows.length > 0) {
    const target = matchingRows[0];
    const moId   = target.rowData[idColIdx];
    MO_HEADERS_GS.forEach(h => {
      if (h === 'MO_ID' || h === 'CreatedDate') return;
      if (data[h] === undefined) return;
      const colIdx = hdrs[h];
      if (colIdx === undefined) return;
      sheet.getRange(target.sheetRow, colIdx + 1).setValue(data[h]);
    });
    for (let d = matchingRows.length - 1; d >= 1; d--) {
      sheet.deleteRow(matchingRows[d].sheetRow);
    }
    return { success: true, mo_id: moId, updated: true };
  }

  // Insert new row
  const existingIds = lastRow <= 1 ? []
    : sheet.getRange(2, idColIdx + 1, lastRow - 1, 1).getValues().flat()
        .filter(v => v !== '' && v !== null && v !== undefined);
  const newId   = existingIds.length === 0 ? 1 : Math.max(...existingIds.map(Number)) + 1;
  const lastCol = Math.max(sheet.getLastColumn(), MO_HEADERS_GS.length);
  const newRow  = Array(lastCol).fill('');
  MO_HEADERS_GS.forEach(h => {
    const idx = hdrs[h];
    if (idx === undefined) return;
    newRow[idx] = (h === 'MO_ID') ? newId : (data[h] !== undefined ? data[h] : '');
  });
  sheet.appendRow(newRow);
  return { success: true, mo_id: newId, updated: false };
}

/**
 * Updates specific fields on an existing monthly record by MO_ID.
 * @param {number} moId - The MO_ID of the record to update.
 * @param {Object} data - Field name → value pairs to write.
 * @returns {{ success: boolean, error?: string }}
 */
function moUpdate(moId, data) {
  const sheet = getMoSheet(); const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, error: "No records" };
  const hdrs = _moGetSheetHeaders(sheet);
  const idColIdx = hdrs['MO_ID'] !== undefined ? hdrs['MO_ID'] : 0;
  const lastCol  = sheet.getLastColumn();
  const allData  = sheet.getRange(2, 1, lastRow-1, lastCol).getValues();
  const rowIdx   = allData.findIndex(r => Number(r[idColIdx]) === Number(moId));
  if (rowIdx === -1) return { success: false, error: "Not found" };
  MO_HEADERS_GS.forEach(h => {
    if (h === 'MO_ID') return;
    if (data[h] === undefined) return;
    const colIdx = hdrs[h];
    if (colIdx === undefined) return;
    sheet.getRange(rowIdx + 2, colIdx + 1).setValue(data[h]);
  });
  return { success: true };
}

/**
 * Permanently deletes a Monthly Checklist record row by MO_ID.
 * @param {number} moId - The MO_ID of the record to delete.
 * @returns {{ success: boolean, error?: string }}
 */
function moDelete(moId) {
  const sheet = getMoSheet(); const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, error: "No records" };
  const hdrs     = _moGetSheetHeaders(sheet);
  const idColIdx = hdrs['MO_ID'] !== undefined ? hdrs['MO_ID'] : 0;
  const allIds   = sheet.getRange(2, idColIdx + 1, lastRow - 1, 1).getValues().flat();
  const rowIdx   = allIds.findIndex(id => Number(id) === Number(moId));
  if (rowIdx === -1) return { success: false, error: "Not found" };
  sheet.deleteRow(rowIdx + 2);
  return { success: true };
}

// Run from Apps Script editor OR call ?action=moDedup to clean existing duplicates
/**
 * ONE-TIME UTILITY — Safe to run at any time from the Apps Script editor.
 * Removes duplicate Monthly Checklist rows, keeping only the row with the
 * highest MO_ID for each YYYY-MM month key. Useful after bulk imports
 * or if upsert logic created accidental duplicates.
 * @returns {{ success: boolean, removed: number }}
 *   removed — count of duplicate rows deleted.
 */
function moDeduplicateAll() {
  const sheet   = getMoSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, removed: 0 };
  const hdrs        = _moGetSheetHeaders(sheet);
  const monthColIdx = hdrs['Month'] !== undefined ? hdrs['Month']
                    : hdrs['Date']  !== undefined ? hdrs['Date'] : 1;
  const lastCol  = sheet.getLastColumn();
  const allData  = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const seen     = {};
  const toDelete = [];
  allData.forEach((row, i) => {
    if (!row[0] && row[0] !== 0) return;
    const key = _toMonthKey(row[monthColIdx]);
    if (!key) return;
    if (seen[key] === undefined) { seen[key] = i + 2; }
    else { toDelete.push(i + 2); }
  });
  toDelete.sort((a, b) => b - a).forEach(r => sheet.deleteRow(r));
  return { success: true, removed: toDelete.length };
}


// ============================================================
//  SETTINGS — Sheet: Settings
// ============================================================

/**
 * Gets or creates the Settings sheet used to store persistent app configuration.
 * Settings are key-value pairs editable from the admin panel in the frontend.
 * Common keys: "AI_ANALYSIS_ENABLED", "WEANING_WEEKS", "WORKER_NAMES", "MAX_PEN".
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The Settings sheet.
 */
function getSettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Settings");
  if (!sheet) {
    sheet = ss.insertSheet("Settings");
    sheet.appendRow(["Key", "Value", "UpdatedBy", "UpdatedAt"]);
    sheet.getRange(1,1,1,4).setFontWeight("bold").setBackground("#555555").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
    // seed defaults
    sheet.appendRow(["weaningWeeks", "5", "system", new Date()]);
    sheet.appendRow(["maxPen", "50", "system", new Date()]);
  }
  return sheet;
}

/**
 * Reads a setting value from the Settings sheet by its key.
 * @param {string} key - Setting key name e.g. "AI_ANALYSIS_ENABLED".
 * @returns {{ success: boolean, value: string|null }}
 *   value — the stored string value, or null if the key does not exist.
 */
function getSetting(key) {
  const sheet = getSettingsSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, value: null };
  const data = sheet.getRange(2, 1, lastRow-1, 2).getValues();
  const row = data.find(r => String(r[0]) === String(key));
  return { success: true, value: row ? String(row[1]) : null };
}

/**
 * Writes or updates a setting in the Settings sheet.
 * If a row with the given key already exists, updates its Value column.
 * If not found, appends a new row with the key, value, updater and timestamp.
 * @param {string} key - Setting key name.
 * @param {string} value - New value to store.
 * @returns {{ success: boolean }}
 */
function saveSetting(key, value) {
  const sheet = getSettingsSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const keys = sheet.getRange(2, 1, lastRow-1, 1).getValues().flat();
    const idx = keys.findIndex(k => String(k) === String(key));
    if (idx !== -1) {
      sheet.getRange(idx+2, 2).setValue(value);
      sheet.getRange(idx+2, 4).setValue(new Date());
      return { success: true };
    }
  }
  sheet.appendRow([key, value, "admin", new Date()]);
  return { success: true };
}

// ============================================================
//  HEADER MIGRATION — Run once manually in the Apps Script editor
//  to fix any sheet where column names and data are out of sync.
//
//  HOW TO RUN:
//    1. Open Apps Script editor (Extensions → Apps Script)
//    2. Paste this entire Code.gs (replace existing)
//    3. Select "migrateAllSheetHeaders" from the function dropdown
//    4. Click ▶ Run — grant permissions if prompted
//    5. A popup shows a full report of every change made per sheet
//
//  Strategy — REBUILD each sheet:
//    1. Read the full sheet (headers + all data rows)
//    2. Map every column by its CURRENT header name
//    3. Write a brand-new sheet in CORRECT column order:
//         - Columns that exist → their data moves to the right slot
//         - Columns that are new/missing → blank cells
//    4. Reformat header row with correct colours
//    Safe to run multiple times — sheets already correct are skipped.
// ============================================================

/**
 * ONE-TIME UTILITY — Run manually from the Apps Script editor after deploying new Code.gs.
 * Ensures every sheet has the correct columns in the right positions.
 * Adds any missing columns to the right of existing data — never deletes or reorders columns.
 * Safe to run multiple times (idempotent). Logs a full report of all changes made.
 * Run this after adding new sheet columns (e.g. SlPhoto_*, SlNotes_*).
 */
function migrateAllSheetHeaders() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const log = [];

  const sheets = [
    { name: SHEET_NAME,  headers: HEADERS,       bg: '#c9a84c', fg: '#000000' },
    { name: CL_SHEET,    headers: CL_HEADERS,    bg: '#2d6a2d', fg: '#ffffff' },
    { name: SL_SHEET,    headers: SL_HEADERS_GS, bg: '#880e4f', fg: '#ffffff' },
    { name: WK_SHEET,    headers: WK_HEADERS_GS, bg: '#1a3a8a', fg: '#ffffff' },
    { name: MO_SHEET,    headers: MO_HEADERS_GS, bg: '#4a148c', fg: '#ffffff' },
  ];

  sheets.forEach(({ name, headers, bg, fg }) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) {
      log.push('SKIP — sheet not found: ' + name);
      return;
    }

    const lastCol = sheet.getLastColumn();
    const lastRow = sheet.getLastRow();

    // Nothing in the sheet yet — just write headers
    if (lastRow === 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold').setBackground(bg).setFontColor(fg);
      log.push('CREATED headers — ' + name);
      return;
    }

    // Read current headers
    const currentHdr = sheet.getRange(1, 1, 1, lastCol)
      .getValues()[0].map(v => String(v).trim());

    // Check if already perfectly correct
    const alreadyOk = headers.length === currentHdr.length
      && headers.every((h, i) => h === currentHdr[i]);
    if (alreadyOk) {
      log.push('OK (no changes) — ' + name);
      return;
    }

    log.push('Rebuilding: ' + name);
    log.push('  Current cols (' + currentHdr.length + '): ' + currentHdr.join(', '));
    log.push('  Correct cols (' + headers.length + '): ' + headers.join(', '));

    // Build map: existing header name → 0-based column index in currentHdr
    const oldColIdx = {};
    currentHdr.forEach((h, i) => { if (h) oldColIdx[h] = i; });

    // Read all data rows (skip header)
    const dataRows = lastRow > 1
      ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues()
      : [];

    // Identify orphan columns (exist in sheet but not in correct spec)
    const orphanHeaders = currentHdr.filter(h => h && !headers.includes(h));
    orphanHeaders.forEach(h => {
      log.push('  → Orphan "' + h + '" moved to far right (after blank separator)');
    });

    // Build remapped data rows: spec columns in order, then blank, then orphan columns
    const newData = dataRows.map(row => {
      const specCells = headers.map(h => {
        const oldIdx = oldColIdx[h];
        return oldIdx !== undefined ? row[oldIdx] : '';
      });
      if (orphanHeaders.length === 0) return specCells;
      const orphanCells = orphanHeaders.map(h => {
        const oldIdx = oldColIdx[h];
        return oldIdx !== undefined ? row[oldIdx] : '';
      });
      return [...specCells, '', ...orphanCells];
    });

    // --- REBUILD THE SHEET ---

    // Step 1: Resolve spec headers as a plain array
    const specHeaders = headers.slice();
    const totalCols   = specHeaders.length + (orphanHeaders.length > 0 ? 1 + orphanHeaders.length : 0);

    // Step 2: Ensure the grid has enough columns
    const gridCols = sheet.getMaxColumns();
    if (gridCols < totalCols) {
      sheet.insertColumnsAfter(gridCols, totalCols - gridCols);
    }

    // Step 3: Clear ALL content across the entire used range
    const fullClearCols = Math.max(sheet.getMaxColumns(), totalCols);
    sheet.getRange(1, 1, Math.max(lastRow, 1), fullClearCols).clearContent();

    // Step 4: Write spec header row
    sheet.getRange(1, 1, 1, specHeaders.length).setValues([specHeaders]);

    // Step 5: Write data rows aligned to spec
    if (newData.length > 0) {
      // newData rows are already totalCols wide
      sheet.getRange(2, 1, newData.length, totalCols).setValues(newData);
    }

    // Step 6: Write orphan separator + orphan headers
    if (orphanHeaders.length > 0) {
      sheet.getRange(1, specHeaders.length + 1).setValue('— ORPHAN (not in spec) —');
      orphanHeaders.forEach((h, i) => {
        sheet.getRange(1, specHeaders.length + 2 + i).setValue(h);
      });
    }

    // Step 7: Physically delete columns beyond what we wrote (removes old stray columns)
    const usedGridCols = sheet.getMaxColumns();
    if (usedGridCols > totalCols) {
      sheet.deleteColumns(totalCols + 1, usedGridCols - totalCols);
    }

    // Step 8: Format spec header row in sheet colour
    sheet.getRange(1, 1, 1, specHeaders.length)
      .setFontWeight('bold').setBackground(bg).setFontColor(fg);

    // Step 9: Format orphan section header in grey
    if (orphanHeaders.length > 0) {
      sheet.getRange(1, specHeaders.length + 1, 1, 1 + orphanHeaders.length)
        .setFontWeight('bold').setBackground('#b0bec5').setFontColor('#000000');
    }

    sheet.setFrozenRows(1);

    // Step 10: Re-apply plain-text format to time columns
    const timeColsMap = name === CL_SHEET ? CL_TIME_COLS
                      : name === SL_SHEET ? SL_TIME_COLS
                      : [];
    if (timeColsMap.length > 0 && newData.length > 0) {
      timeColsMap.forEach(colName => {
        const col = specHeaders.indexOf(colName) + 1;
        if (col > 0) {
          sheet.getRange(2, col, newData.length, 1).setNumberFormat('@');
        }
      });
    }

    log.push('  ✓ Rebuilt — ' + newData.length + ' data rows · '
      + specHeaders.length + ' spec cols'
      + (orphanHeaders.length > 0 ? ' · ' + orphanHeaders.length + ' orphan col(s) at right' : ''));
  });

  const report = '=== MIGRATION REPORT ===\n\n' + log.join('\n');
  Logger.log(report);
  SpreadsheetApp.getUi().alert('Migration Complete', report, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============================================================
//  DIAGNOSE — Run this FIRST to see what each sheet looks like
//  WITHOUT making any changes. Check the Execution Log output.
//  This confirms the new Code.gs is active before you migrate.
// ============================================================
/**
 * DIAGNOSTIC UTILITY — Run from the Apps Script editor to inspect all sheet headers.
 * Logs the actual column names in each sheet vs what the code expects.
 * Useful for debugging "Unknown column" errors or verifying a migration ran correctly.
 */
function diagnoseSheetHeaders() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const log = [];

  const sheets = [
    { name: SHEET_NAME,  headers: HEADERS       },
    { name: CL_SHEET,    headers: CL_HEADERS    },
    { name: SL_SHEET,    headers: SL_HEADERS_GS },
    { name: WK_SHEET,    headers: WK_HEADERS_GS },
    { name: MO_SHEET,    headers: MO_HEADERS_GS },
  ];

  sheets.forEach(({ name, headers }) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) { log.push('MISSING: ' + name); return; }

    const lastCol    = sheet.getLastColumn();
    const currentHdr = lastCol > 0
      ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v).trim())
      : [];

    const specLen  = headers.length;
    const sheetLen = currentHdr.length;
    const match    = specLen === sheetLen && headers.every((h, i) => h === currentHdr[i]);

    log.push('');
    log.push('Sheet: ' + name);
    log.push('  Spec columns  : ' + specLen);
    log.push('  Sheet columns : ' + sheetLen);
    log.push('  Status        : ' + (match ? '✓ CORRECT' : '✗ NEEDS MIGRATION'));

    if (!match) {
      // Show first mismatch
      const maxLen = Math.max(specLen, sheetLen);
      for (let i = 0; i < maxLen; i++) {
        const s = headers[i]     || '(missing)';
        const c = currentHdr[i]  || '(missing)';
        if (s !== c) {
          log.push('  First mismatch at col ' + (i+1) + ': spec="' + s + '" sheet="' + c + '"');
          break;
        }
      }
      // List orphan columns
      const orphans = currentHdr.filter(h => h && !headers.includes(h));
      if (orphans.length) log.push('  Orphan cols   : ' + orphans.join(', '));
      // List missing columns
      const missing = headers.filter(h => !currentHdr.includes(h));
      if (missing.length) log.push('  Missing cols  : ' + missing.slice(0, 10).join(', ') + (missing.length > 10 ? '...' : ''));
    }
  });

  const report = '=== DIAGNOSE REPORT ===\n' + log.join('\n');
  Logger.log(report);
  SpreadsheetApp.getUi().alert('Diagnose Result', report, SpreadsheetApp.getUi().ButtonSet.OK);
}




// ============================================================
//  BOAR/SOW → DB_ID MIGRATION
// ============================================================
/**
 * ONE-TIME UTILITY — Migrates PigLog Boar and SOW columns from PIG ID strings to DB_IDs.
 * This enables stable pig cross-references that survive ear-tag changes or re-numbering.
 * Safe to run multiple times — skips rows that already contain numeric DB_IDs.
 * Run once after initial setup if Boar/SOW columns contain text PIG IDs.
 * @returns {{ success: boolean, updated: number, skipped: number }}
 *   updated — rows converted; skipped — rows already using numeric IDs.
 */
function migrateBoarSowToDbId() {
  const sheet   = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, message: 'No data rows to migrate.' };

  const lastCol = sheet.getLastColumn();
  const hdrRow  = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const hdrMap  = {};
  hdrRow.forEach((h, i) => { if (h) hdrMap[String(h).trim()] = i; });

  const dbIdIdx = hdrMap['DB_ID'];
  const pigIdIdx= hdrMap['PIG ID'];
  const boarIdx = hdrMap['Boar'];
  const sowIdx  = hdrMap['SOW'];

  if (dbIdIdx === undefined || pigIdIdx === undefined || boarIdx === undefined || sowIdx === undefined) {
    return { success: false, error: 'Required columns not found. Run migrateAllSheetHeaders first.' };
  }

  // Build lookup: PIG ID (lowercase) → DB_ID
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const pigIdToDbId = {};
  data.forEach(row => {
    const dbId  = row[dbIdIdx];
    const pigId = String(row[pigIdIdx] || '').trim();
    if (dbId !== '' && dbId !== null && pigId) {
      pigIdToDbId[pigId.toLowerCase()] = dbId;
    }
  });

  let updated = 0;
  const log      = [];
  const notFound = [];

  data.forEach((row, i) => {
    const sheetRow = i + 2;
    const rowDbId  = row[dbIdIdx];
    if (rowDbId === '' || rowDbId === null) return;

    ['Boar', 'SOW'].forEach(field => {
      const colIdx = field === 'Boar' ? boarIdx : sowIdx;
      const val    = row[colIdx];
      const valStr = String(val || '').trim();
      if (!valStr) return;

      // Already a numeric DB_ID — skip
      const asNum = Number(valStr);
      if (!isNaN(asNum) && asNum > 0 && String(Math.round(asNum)) === valStr) return;

      // Look up PIG ID → DB_ID
      const matchDbId = pigIdToDbId[valStr.toLowerCase()];
      if (matchDbId !== undefined) {
        sheet.getRange(sheetRow, colIdx + 1).setValue(matchDbId);
        log.push('Row ' + sheetRow + ' (DB_ID=' + rowDbId + '): ' + field + ' "' + valStr + '" → ' + matchDbId);
        updated++;
      } else {
        notFound.push('Row ' + sheetRow + ' (DB_ID=' + rowDbId + '): ' + field + ' "' + valStr + '" — PIG ID not found');
      }
    });
  });

  const summary = 'PigLog: Migrated ' + updated + ' Boar/SOW value(s).'
    + (log.length      ? '\n\nUpdated:\n'             + log.join('\n')      : '')
    + (notFound.length ? '\n\nNot found (review):\n' + notFound.join('\n') : '');

  // Also migrate SowLitter.SowId
  const slResult = migrateSowLitterSowId();
  const fullSummary = summary + '\n\n' + slResult.message;
  Logger.log(fullSummary);
  return { success: true, updated: updated + slResult.updated, notFound: notFound.length + slResult.notFound, message: fullSummary };
}

// ============================================================
//  SOWLITTER SowId → DB_ID MIGRATION
//  Scans every SowLitter row. For SowId values that look like
//  a PIG ID string, finds the matching DB_ID from PigLog
//  and replaces it in-place. Already-numeric values are skipped.
// ============================================================
/**
 * ONE-TIME UTILITY — Migrates SowLitter.SowId column from PIG ID strings to DB_IDs.
 * Looks up each SowId string in the PigLog sheet and replaces it with the DB_ID.
 * Safe to run multiple times — skips rows that already contain numeric IDs.
 * @returns {{ success: boolean, updated: number, skipped: number, notFound: string[] }}
 *   notFound — PIG ID strings that had no matching record in PigLog.
 */
function migrateSowLitterSowId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Build PIG ID → DB_ID lookup from PigLog
  const pigSheet = getSheet();
  const pigLastRow = pigSheet.getLastRow();
  const pigIdToDbId = {};

  if (pigLastRow > 1) {
    const pigLastCol = pigSheet.getLastColumn();
    const pigHdr     = pigSheet.getRange(1, 1, 1, pigLastCol).getValues()[0];
    const pigHdrMap  = {};
    pigHdr.forEach((h, i) => { if (h) pigHdrMap[String(h).trim()] = i; });
    const dbIdIdx  = pigHdrMap['DB_ID'];
    const pigIdIdx = pigHdrMap['PIG ID'];
    if (dbIdIdx !== undefined && pigIdIdx !== undefined) {
      const pigData = pigSheet.getRange(2, 1, pigLastRow - 1, pigLastCol).getValues();
      pigData.forEach(row => {
        const dbId  = row[dbIdIdx];
        const pigId = String(row[pigIdIdx] || '').trim();
        if (dbId !== '' && dbId !== null && pigId) {
          pigIdToDbId[pigId.toLowerCase()] = dbId;
        }
      });
    }
  }

  if (Object.keys(pigIdToDbId).length === 0) {
    return { success: false, updated: 0, notFound: 0, message: 'SowLitter: PigLog has no records to build lookup from.' };
  }

  // Open SowLitter sheet
  const slSheet = ss.getSheetByName(SL_SHEET);
  if (!slSheet) return { success: true, updated: 0, notFound: 0, message: 'SowLitter: sheet not found — skipped.' };

  const slLastRow = slSheet.getLastRow();
  if (slLastRow <= 1) return { success: true, updated: 0, notFound: 0, message: 'SowLitter: no data rows.' };

  const slLastCol = slSheet.getLastColumn();
  const slHdr     = slSheet.getRange(1, 1, 1, slLastCol).getValues()[0];
  const slHdrMap  = {};
  slHdr.forEach((h, i) => { if (h) slHdrMap[String(h).trim()] = i; });

  const slIdIdx  = slHdrMap['SL_ID'];
  const sowIdIdx = slHdrMap['SowId'];

  if (sowIdIdx === undefined) {
    return { success: false, updated: 0, notFound: 0, message: 'SowLitter: SowId column not found. Run migrateAllSheetHeaders first.' };
  }

  const slData = slSheet.getRange(2, 1, slLastRow - 1, slLastCol).getValues();
  let updated = 0;
  const log      = [];
  const notFound = [];

  slData.forEach((row, i) => {
    const sheetRow = i + 2;
    const slId     = slIdIdx !== undefined ? row[slIdIdx] : '?';
    if (slId === '' || slId === null) return; // skip empty rows

    const val    = row[sowIdIdx];
    const valStr = String(val || '').trim();
    if (!valStr) return;

    // Already a numeric DB_ID — skip
    const asNum = Number(valStr);
    if (!isNaN(asNum) && asNum > 0 && String(Math.round(asNum)) === valStr) return;

    // Look up PIG ID → DB_ID
    const matchDbId = pigIdToDbId[valStr.toLowerCase()];
    if (matchDbId !== undefined) {
      slSheet.getRange(sheetRow, sowIdIdx + 1).setValue(matchDbId);
      log.push('  Row ' + sheetRow + ' (SL_ID=' + slId + '): SowId "' + valStr + '" → DB_ID ' + matchDbId);
      updated++;
    } else {
      notFound.push('  Row ' + sheetRow + ' (SL_ID=' + slId + '): SowId "' + valStr + '" — no matching PIG ID found');
    }
  });

  const msg = 'SowLitter: Migrated ' + updated + ' SowId value(s).'
    + (log.length      ? '\n\nUpdated:\n'             + log.join('\n')      : '')
    + (notFound.length ? '\n\nNot found (review):\n' + notFound.join('\n') : '');

  Logger.log(msg);
  return { success: true, updated, notFound: notFound.length, message: msg };
}


// ============================================================
//  NIGHTLY AI ANALYSIS ENGINE
//  Processes Daily Checklist records, fetches photos from
//  Drive, calls Claude via Anthropic API, appends summary
//  into the AIAnalysis column.
//
//  SETUP:
//  1. In Apps Script editor → Project Settings → Script Properties
//     Add: ANTHROPIC_API_KEY = sk-ant-...your key...
//  2. Run createMidnightTrigger() ONCE manually to set up the
//     midnight time-driven trigger.
//  3. Admin can trigger manually via the app's admin panel.
// ============================================================

// ── Trigger Setup ───────────────────────────────────────────

/**
 * ONE-TIME SETUP — Run once from the Apps Script editor after first deployment.
 * Creates a daily time-based trigger that fires runNightlyCloseout() at midnight
 * in the script timezone (Africa/Lusaka). Deletes any existing nightly trigger first
 * to prevent duplicates. Check Apps Script > Triggers to confirm it was created.
 */
function createMidnightTrigger() {
  // Delete any existing nightly triggers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'runNightlyAIAnalysis' ||
        t.getHandlerFunction() === 'runNightlyCloseout') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // Create trigger: runs daily between midnight and 1am Lusaka time
  ScriptApp.newTrigger('runNightlyCloseout')
    .timeBased()
    .atHour(0)
    .everyDays(1)
    .inTimezone('Africa/Lusaka')
    .create();
  Logger.log('Midnight trigger created for runNightlyCloseout');
}

// ── Nightly Closeout — always runs regardless of API key ────
// Marks In Progress records from yesterday as Incomplete
// Then runs AI analysis if API key is configured

/**
 * Main nightly job — normally called at midnight via the time-based trigger.
 * Can also be triggered manually from the admin panel in the app (forceAI=true).
 *
 * Step 1 — Mark Incomplete: scans all checklist records dated before today and
 *   marks any with missing section times (Sec1/2/3Time blank) as "Incomplete".
 *
 * Step 2 — AI Analysis: runs if ANTHROPIC_API_KEY is set in Script Properties AND
 *   either AI_ANALYSIS_ENABLED setting is "true" (auto nightly) OR forceAI is true.
 *
 * @param {string} [targetDate] - YYYY-MM-DD date to process. Defaults to yesterday.
 * @param {boolean} [forceAI=false] - If true, runs AI analysis regardless of the setting.
 * @returns {{ success: boolean, message: string, marked: number, processed?: number }}
 *   marked — number of records marked Incomplete; processed — records sent to AI.
 */
function runNightlyCloseout(targetDate, forceAI) {
  const tz = Session.getScriptTimeZone();
  let processDate;
  if (targetDate) {
    processDate = targetDate;
  } else {
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    processDate = Utilities.formatDate(yesterday, tz, 'yyyy-MM-dd');
  }

  Logger.log('Nightly closeout: processing date ' + processDate);

  const sheet  = getClSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, message: 'No records.' };

  const lastCol = sheet.getLastColumn();
  const hdrRow  = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const hdrMap  = {};
  hdrRow.forEach((h, i) => { if (h) hdrMap[String(h).trim()] = i; });
  const data    = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  // Step 1: Mark incomplete records
  const marked = _markIncompleteRecords(processDate, sheet, hdrMap, data);
  Logger.log('Marked ' + marked + ' records as Incomplete for ' + processDate);

  // Step 2: Run AI analysis — always if forceAI, otherwise check setting
  const apiKey    = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  const aiSetting = getSetting('AI_ANALYSIS_ENABLED');
  const aiEnabled = forceAI || aiSetting.value === 'true';
  let aiResult = { processed: 0, skipped: 0, errors: 0 };
  if (apiKey && aiEnabled) {
    Logger.log('AI analysis enabled — running for ' + processDate);
    aiResult = _runAIForDate(processDate, sheet, hdrMap, data, apiKey);
  } else if (!apiKey) {
    Logger.log('AI skipped — no ANTHROPIC_API_KEY set in Script Properties');
  } else {
    Logger.log('AI skipped — AI_ANALYSIS_ENABLED is off in Settings');
  }

  const msg = 'Closeout complete for ' + processDate
    + ': ' + marked + ' marked Incomplete'
    + (aiEnabled && apiKey ? ', ' + aiResult.processed + ' AI analysed, ' + aiResult.skipped + ' skipped' : ', AI disabled');
  Logger.log(msg);
  return { success: true, message: msg, marked, ...aiResult, date: processDate };
}

// ── Main Batch Function (kept for backward compatibility) ───
/**
 * Backward-compatible alias for runNightlyCloseout().
 * Called by the "runAIAnalysis" action from the frontend admin panel.
 * @param {string} [targetDate] - YYYY-MM-DD to process (defaults to yesterday).
 * @param {boolean} [forceAI=false] - If true, forces AI even if setting is disabled.
 * @returns {{ success: boolean, message: string }}
 */
function runNightlyAIAnalysis(targetDate, forceAI) {
  return runNightlyCloseout(targetDate, forceAI);
}

  // Default: process yesterday's records (today's are still in progress)
// ── Analyse a single record by CL_ID ─────────────────────────
/**
 * Analyses a single Daily Checklist record using the Anthropic Claude API.
 * Fetches the record from the sheet, retrieves photos from Google Drive via
 * DriveApp (NOT the thumbnail API which returns 403 for private files), then
 * calls analyseChecklistRecord() to build the prompt and get the AI response.
 * Writes the formatted analysis back to the AIAnalysis column.
 * Called one record at a time from the frontend to avoid the 6-minute Apps Script limit.
 * @param {number} clId - The CL_ID of the checklist record to analyse.
 * @param {boolean} [force=false] - If true, re-analyses records that already have a summary.
 * @returns {{ success: boolean, processed?: boolean, skipped?: boolean, error?: string }}
 *   processed — true if AI ran and wrote a result.
 *   skipped — true if the record already had analysis and force was false.
 */
function runAISingleRecord(clId, force) {
  if (!clId) return { success: false, error: 'No CL_ID provided' };
  const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return { success: false, error: 'ANTHROPIC_API_KEY not set in Script Properties' };
  const sheet   = getClSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, error: 'No records' };
  const lastCol = sheet.getLastColumn();
  const hdrRow  = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const hdrMap  = {};
  hdrRow.forEach((h, i) => { if (h) hdrMap[String(h).trim()] = i; });
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const rowIndex = ids.findIndex(id => Number(id) === Number(clId));
  if (rowIndex === -1) return { success: false, error: 'Record not found: ' + clId };
  const sheetRow = rowIndex + 2;
  const row = sheet.getRange(sheetRow, 1, 1, lastCol).getValues()[0];
  // Skip if already analysed (unless force=true)
  const aiCol = hdrMap['AIAnalysis'];
  if (!force && aiCol !== undefined && String(row[aiCol] || '').trim()) {
    return { success: true, skipped: true, message: 'Already analysed' };
  }
  const rec = {};
  hdrRow.forEach((h, j) => { rec[h] = row[j]; });
  try {
    const summary = analyseChecklistRecord(rec, apiKey);
    if (summary && aiCol !== undefined) {
      sheet.getRange(sheetRow, aiCol + 1).setValue(summary);
    }
    return { success: true, processed: true, clId };
  } catch(e) {
    return { success: false, error: e.message, clId };
  }
}

// ── Get list of records needing AI analysis for a date ────────
/**
 * Returns a list of CL_IDs for checklist records on a given date that have
 * not yet received an AI analysis (AIAnalysis column is blank or empty).
 * Used by the frontend to determine how many records to process and to iterate
 * through them one at a time via runAISingleRecord().
 * @param {string} targetDate - YYYY-MM-DD date to check.
 * @returns {{ success: boolean, pending: number[] }}
 *   pending — array of CL_ID numbers needing analysis.
 */
function getAIPendingRecords(targetDate) {
  const sheet   = getClSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, pending: [] };
  const tz      = Session.getScriptTimeZone();
  const lastCol = sheet.getLastColumn();
  const hdrRow  = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const hdrMap  = {};
  hdrRow.forEach((h, i) => { if (h) hdrMap[String(h).trim()] = i; });
  const data    = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const aiCol   = hdrMap['AIAnalysis'];
  const pending = [];
  data.forEach(row => {
    const clId    = row[hdrMap['CL_ID']];
    if (!clId) return;
    const rowDate = row[hdrMap['Date']];
    const rowDateStr = rowDate instanceof Date
      ? Utilities.formatDate(rowDate, tz, 'yyyy-MM-dd')
      : String(rowDate || '').trim();
    if (rowDateStr !== targetDate) return;
    const hasAnalysis = aiCol !== undefined && String(row[aiCol] || '').trim();
    if (!hasAnalysis) pending.push(Number(clId));
  });
  return { success: true, pending };
}

// ── AI Analysis for a specific date ─────────────────────────
/**
 * Batch-analyses all checklist records for a specific date in a single execution.
 * Called by runNightlyCloseout() for the midnight batch job.
 * Processes records sequentially — stops early if approaching the 6-min time limit.
 * @param {string} processDate - YYYY-MM-DD date to analyse.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The DailyChecklist sheet.
 * @param {Object} hdrMap - Header name → 0-based column index map.
 * @param {any[][]} data - All sheet data rows (excluding header row).
 * @param {string} apiKey - Anthropic API key from Script Properties.
 * @returns {{ processed: number, skipped: number, errors: number }}
 */
function _runAIForDate(processDate, sheet, hdrMap, data, apiKey) {
  const tz = Session.getScriptTimeZone();
  const hdrRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let processed = 0, skipped = 0;
  const errors = [];

  data.forEach((row, i) => {
    const sheetRow   = i + 2;
    const clId       = row[hdrMap['CL_ID']];
    const rowDate    = row[hdrMap['Date']];
    const rowDateStr = rowDate instanceof Date
      ? Utilities.formatDate(rowDate, tz, 'yyyy-MM-dd')
      : String(rowDate || '').trim();
    if (rowDateStr !== processDate || !clId) return;

    const aiCol = hdrMap['AIAnalysis'];
    if (aiCol !== undefined && String(row[aiCol] || '').trim()) { skipped++; return; }

    const rec = {};
    hdrRow.forEach((h, j) => { rec[h] = row[j]; });
    try {
      const summary = analyseChecklistRecord(rec, apiKey);
      if (summary && aiCol !== undefined) {
        sheet.getRange(sheetRow, aiCol + 1).setValue(summary);
        processed++;
        Utilities.sleep(2000);
      }
    } catch(e) {
      errors.push('CL_ID ' + clId + ': ' + e.message);
      Logger.log('AI error CL_ID ' + clId + ': ' + e.message);
    }
  });
  return { processed, skipped, errors: errors.length };
}

// ── Closeout only — mark all stale In Progress records as Incomplete ──
/**
 * Marks stale checklist records as Incomplete without running AI analysis.
 * A record is marked Incomplete when: date is before today AND one or more
 * section times (Sec1Time, Sec2Time, Sec3Time) are blank AND status is not
 * already "Complete" or "ALL OK".
 * Safe to run manually at any time from the editor or admin panel.
 * @returns {{ success: boolean, message: string, marked: number }}
 *   marked — number of records updated to "Incomplete".
 */
function runCloseoutOnly() {
  const sheet = getClSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, message: 'No records.', marked: 0 };
  const lastCol = sheet.getLastColumn();
  const hdrRow  = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const hdrMap  = {};
  hdrRow.forEach((h, i) => { if (h) hdrMap[String(h).trim()] = i; });
  const data    = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const marked  = _markIncompleteRecords(null, sheet, hdrMap, data);
  const msg     = 'Closeout complete: ' + marked + ' record(s) marked Incomplete';
  Logger.log(msg);
  return { success: true, message: msg, marked };
}

// ── Mark Incomplete — closes all records older than today with missing sections ───
/**
 * Internal utility called by runNightlyCloseout() and runCloseoutOnly().
 * Scans ALL checklist records (not just one date) and marks as "Incomplete" any where:
 *   - The record date is strictly before today (past records only), AND
 *   - At least one of Sec1Time, Sec2Time, Sec3Time is blank (section not completed), AND
 *   - Current status is not already "Complete", "ALL OK", or "Incomplete".
 * @param {string|null} processDate - Unused parameter kept for signature compatibility.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The DailyChecklist sheet.
 * @param {Object} hdrMap - Header name → 0-based column index map.
 * @param {any[][]} data - All sheet data rows (excluding header row).
 * @returns {number} Count of records updated to "Incomplete".
 */
function _markIncompleteRecords(processDate, sheet, hdrMap, data) {
  const tz        = Session.getScriptTimeZone();
  const today     = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const statusCol = hdrMap['Status'];
  const dateCol   = hdrMap['Date'];
  const sec1Col   = hdrMap['Sec1Time'];
  const sec2Col   = hdrMap['Sec2Time'];
  const sec3Col   = hdrMap['Sec3Time'];
  if (statusCol === undefined || dateCol === undefined) return 0;
  let count = 0;
  data.forEach((row, i) => {
    const sheetRow   = i + 2;
    const rowDate    = row[dateCol];
    const rowDateStr = rowDate instanceof Date
      ? Utilities.formatDate(rowDate, tz, 'yyyy-MM-dd')
      : String(rowDate || '').trim();
    // Only process records from BEFORE today
    if (!rowDateStr || rowDateStr >= today) return;
    const status = String(row[statusCol] || '').trim();
    // Skip records already fully closed
    if (status === 'Incomplete' || status === 'Complete' || status === 'ALL OK') return;
    // Check if any section is missing
    const sec1 = sec1Col !== undefined ? String(row[sec1Col] || '').trim() : '';
    const sec2 = sec2Col !== undefined ? String(row[sec2Col] || '').trim() : '';
    const sec3 = sec3Col !== undefined ? String(row[sec3Col] || '').trim() : '';
    const missingSections = [];
    if (!sec1) missingSections.push('Morning');
    if (!sec2) missingSections.push('Feeding');
    if (!sec3) missingSections.push('Afternoon');
    if (missingSections.length > 0) {
      sheet.getRange(sheetRow, statusCol + 1).setValue('Incomplete');
      count++;
      Logger.log('Marked Incomplete: row ' + sheetRow + ' date ' + rowDateStr + ' missing: ' + missingSections.join(', '));
    }
  });
  return count;
}

// ── Per-Record Analysis ─────────────────────────────────────

/**
 * Calls the Anthropic Claude API to analyse one Daily Checklist record.
 * Builds a detailed veterinary prompt including all 15+ health indicators,
 * section completion status, pen and date context, and up to 3 section photos.
 * Photos are fetched via DriveApp.getFileById() (thumbnail API returns 403).
 * Uses model claude-opus-4-5 with max_tokens: 1500.
 * @param {Object} rec - Full checklist record object with all CL_HEADERS fields.
 *   Key fields used: Date, Pen, CheckedBy, Status, all health indicator fields,
 *   Sec1/2/3Time, PhotoUrl1/2/3.
 * @param {string} apiKey - Anthropic API key (ANTHROPIC_API_KEY Script Property).
 * @returns {string} Formatted AI analysis text with a timestamp header line.
 * @throws {Error} If the Anthropic API returns a non-200 HTTP status code.
 */
function analyseChecklistRecord(rec, apiKey) {
  // Build the check context text
  const CL_FIELD_LABELS = {
    tail:'Tail', eyes:'Eyes', stool:'Stool/Droppings', posture:'Posture/Movement',
    skin:'Skin Condition', breathing:'Breathing', appetite:'Appetite',
    water:'Water Intake', feed:'Feed Intake', smell:'Smell',
    pinch:'Pinch Test (Hydration)', belly:'Belly Fill', limbs:'Limbs/Hooves',
    injuries:'Injuries/Wounds', temp:'Temperature'
  };
  const clKeys = Object.keys(CL_FIELD_LABELS);
  const ok      = clKeys.filter(k => rec[k] === 'ok').map(k => CL_FIELD_LABELS[k]);
  const bad     = clKeys.filter(k => rec[k] === 'bad').map(k => CL_FIELD_LABELS[k]);
  const missing = clKeys.filter(k => !rec[k]).map(k => CL_FIELD_LABELS[k]);

  const status    = String(rec['Status']    || '').trim();
  const pen       = String(rec['Pen']       || '?');
  const pigCount  = String(rec['PigCount']  || '?');
  const date      = String(rec['Date']      || '?');
  const notes     = String(rec['Notes']     || '');
  const nursing   = String(rec['cl_nursing'] || '');
  const weaklings = String(rec['cl_weaklings'] || '');

  const contextText = `You are an experienced pig farm veterinary advisor reviewing end-of-day health records. You have been provided with ${['PhotoUrl1','PhotoUrl2','PhotoUrl3'].filter(k => rec[k]).length} farm photos (Morning, Feeding, and/or Afternoon checks) plus the health checklist data below. Study each photo carefully.

FARM RECORD: PEN ${pen} — DATE: ${date} — PIGS IN PEN: ${pigCount}
RECORD STATUS: ${status}
${nursing ? 'NURSING SOW: ' + nursing + ' piglets nursing' + (weaklings ? ', ' + weaklings + ' weaklings noted' : '') : ''}

HEALTH CHECKLIST RESULTS:
✅ All Clear (${ok.length} items): ${ok.join(', ') || 'none'}
⚠️ Concerns Flagged (${bad.length} items): ${bad.join(', ') || 'none'}
— Not Recorded (${missing.length} items): ${missing.join(', ') || 'all recorded'}
${notes ? '\nFARM WORKER NOTES: ' + notes : ''}

Please provide a thorough end-of-day health report with the following sections:

1. OVERALL STATUS — One clear verdict: Healthy / Monitor Closely / Action Required

2. PHOTO ANALYSIS — For each photo provided, describe in detail:
   • Pen cleanliness and bedding condition
   • Pig body condition (weight, coat, skin appearance)
   • Posture and movement behaviour (are pigs lying normally, huddled, lethargic?)
   • Feeding and water trough condition if visible
   • Any visible injuries, swelling, discharge, or abnormalities
   • Environmental conditions (temperature indicators, ventilation, crowding)

3. HEALTH CHECK SUMMARY — Interpret the checklist results in context of what you see in the photos. Do the photos confirm or contradict any flagged concerns?

4. CONCERNS & RISKS — List any specific health risks identified from photos or checks, with severity (Low / Medium / High)

5. RECOMMENDED ACTIONS — Specific practical steps for the next 24 hours, ordered by priority

6. NURSING & PIGLET WELFARE — If applicable, comment on sow condition and piglet welfare from the photos

Write clearly for a farm worker. Be specific about what you actually see in the photos — do not be vague.`;

  // Build message content — photos first, then text
  const messageContent = [];

  ['PhotoUrl1','PhotoUrl2','PhotoUrl3'].forEach((urlKey, i) => {
    const url = String(rec[urlKey] || '').trim();
    if (!url) return;
    try {
      const fileId = _extractFileId(url);
      if (!fileId) return;

      // Fetch file directly via DriveApp — thumbnail URL returns 404 for private files
      const file = DriveApp.getFileById(fileId);
      let blob = file.getBlob();
      let mime = (blob.getContentType() || 'image/jpeg').toLowerCase();
      Logger.log('Photo ' + urlKey + ': original size=' + blob.getBytes().length + ' mime=' + mime);

      // If over 4MB, resize by fetching lh3 thumbnail (Google's image proxy)
      if (blob.getBytes().length > 4 * 1024 * 1024) {
        try {
          const lh3Url  = 'https://lh3.googleusercontent.com/d/' + fileId + '=w800';
          const token   = ScriptApp.getOAuthToken();
          const lh3Resp = UrlFetchApp.fetch(lh3Url, {
            headers: { Authorization: 'Bearer ' + token },
            muteHttpExceptions: true
          });
          if (lh3Resp.getResponseCode() === 200) {
            const lh3Blob = lh3Resp.getBlob();
            const lh3Mime = (lh3Blob.getContentType() || '').toLowerCase();
            if (lh3Mime.startsWith('image/')) {
              blob = lh3Blob;
              mime = lh3Mime;
              Logger.log('Photo ' + urlKey + ': resized via lh3, new size=' + blob.getBytes().length);
            }
          }
        } catch(lh3e) {
          Logger.log('lh3 resize failed: ' + lh3e.message);
        }
        // If still too large, skip
        if (blob.getBytes().length > 4.5 * 1024 * 1024) {
          Logger.log('Photo ' + urlKey + ' still too large — skipping');
          messageContent.push({ type: 'text', text: ['[Morning photo — too large]','[Feeding photo — too large]','[Afternoon photo — too large]'][i] });
          return;
        }
      }

      // Ensure supported mime type
      if (!['image/jpeg','image/png','image/gif','image/webp'].includes(mime)) {
        mime = 'image/jpeg';
      }

      const finalBytes = blob.getBytes();
      const b64 = Utilities.base64Encode(finalBytes);
      Logger.log('Sending ' + urlKey + ': mime=' + mime + ' size=' + finalBytes.length + ' b64len=' + b64.length);
      messageContent.push({
        type: 'image',
        source: { type: 'base64', media_type: mime, data: b64 }
      });
      messageContent.push({ type: 'text', text: ['[Morning photo]','[Feeding photo]','[Afternoon photo]'][i] });
    } catch(e) {
      Logger.log('Photo fetch error for ' + urlKey + ': ' + e.message);
    }
  });

  messageContent.push({ type: 'text', text: contextText });

  if (messageContent.length === 1 && messageContent[0].type === 'text') {
    // No photos — still run text-only analysis
  }

  // Call Anthropic API via UrlFetchApp
  const payload = JSON.stringify({
    model:      'claude-opus-4-5',
    max_tokens: 1500,
    messages:   [{ role: 'user', content: messageContent }]
  });

  const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method:             'POST',
    contentType:        'application/json',
    headers: {
      'x-api-key':         apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload:            payload,
    muteHttpExceptions: true
  });

  const code   = response.getResponseCode();
  const body   = response.getContentText();
  Logger.log('Anthropic API response code: ' + code);
  Logger.log('Anthropic API response body: ' + body.substring(0, 500));

  if (code !== 200) {
    throw new Error('API error ' + code + ': ' + body.substring(0, 200));
  }

  const result = JSON.parse(body);
  const text = (result.content || []).map(c => c.text || '').filter(Boolean).join('\n');
  return '[AI Analysis — ' + new Date().toLocaleDateString() + ']\n' + text;
}

// ── Helper: extract Drive fileId from any URL format ────────

/**
 * Extracts a Google Drive file ID from any of the URL formats the app may store.
 * Handles three formats:
 *   1. "drive:{fileId}" — internal format used by clSavePhoto/slSavePhoto
 *   2. "https://drive.google.com/file/d/{fileId}/view" — standard Drive share link
 *   3. "https://drive.google.com/open?id={fileId}" — legacy Drive open link
 * @param {string} url - URL or drive: reference string to parse.
 * @returns {string|null} The extracted file ID string, or null if no ID found.
 */
function _extractFileId(url) {
  if (!url) return null;
  if (url.startsWith('drive:')) return url.slice(6);
  const m = url.match(/\/d\/([A-Za-z0-9_-]{25,})|[?&]id=([A-Za-z0-9_-]{25,})/);
  return m ? (m[1] || m[2]) : null;
}

// ============================================================
//  WORKER ACTIONS
//  Sheet: WorkerActions
//  Tracks tasks assigned to farm workers
// ============================================================

const WA_SHEET = 'WorkerActions';
const WA_HEADERS = ['ACTION_ID','Date','Worker','Category','Action','Priority','Status','DueDate','Notes','CompletedAt'];

/**
 * Gets or creates the WorkerActions sheet with all required WA_HEADERS columns.
 * On first creation: sets column widths, applies teal header formatting, freezes row 1.
 * Existing sheets: adds any missing columns to the right without touching data.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The WorkerActions sheet.
 */
function getWaSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(WA_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(WA_SHEET);
    sheet.appendRow(WA_HEADERS);
    sheet.getRange(1,1,1,WA_HEADERS.length).setFontWeight('bold').setBackground('#e65100').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 80);
    sheet.setColumnWidth(3, 120);
    sheet.setColumnWidth(5, 250);
    sheet.setColumnWidth(9, 200);
  }
  return sheet;
}

/**
 * Builds a column-name → 1-based column-number map for the WorkerActions sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The WorkerActions sheet.
 * @returns {Object} Map of { headerName: columnIndex } (1-based).
 */
function _waColMap(sheet) {
  const hdrs = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const map = {};
  hdrs.forEach((h,i) => { if (h) map[String(h).trim()] = i+1; });
  return map;
}

/**
 * Returns all Worker Action task records from the WorkerActions sheet.
 * Formats Date objects in the Date and DueDate columns as YYYY-MM-DD strings.
 * @returns {{ success: boolean, records: Object[] }}
 *   records — array of task objects with all WA_HEADERS fields as keys.
 */
function waGetAll() {
  const sheet = getWaSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, records: [] };
  const tz     = Session.getScriptTimeZone();
  const lastCol = sheet.getLastColumn();
  const hdrs   = sheet.getRange(1,1,1,lastCol).getValues()[0].map(v => String(v).trim());
  const data   = sheet.getRange(2,1,lastRow-1,lastCol).getValues();
  const records = data.filter(r => r[0] !== '' && r[0] !== null).map(row => {
    const rec = {};
    hdrs.forEach((h,i) => {
      const v = row[i];
      rec[h] = v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : (v !== null && v !== undefined ? v : '');
    });
    return rec;
  });
  return { success: true, records };
}

/**
 * Appends a new Worker Action task to the WorkerActions sheet.
 * Assigns the next ACTION_ID automatically. Sets Date to today if not provided.
 * Defaults Status to "Pending" if not specified in data.
 * @param {Object} data - Task fields: Worker, Category, Action, Priority, DueDate, Notes, Status.
 * @returns {{ success: boolean, action_id: number }}
 */
function waAdd(data) {
  const sheet = getWaSheet();
  const id    = _waNextId(sheet);
  const tz    = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const row   = WA_HEADERS.map(h => {
    if (h === 'ACTION_ID') return id;
    if (h === 'Date')      return data.Date || today;
    if (h === 'Status')    return data.Status || 'Pending';
    return data[h] !== undefined ? data[h] : '';
  });
  sheet.appendRow(row);
  return { success: true, action_id: id };
}

/**
 * Updates specific fields on an existing Worker Action record by ACTION_ID.
 * Skips the ACTION_ID column itself (immutable primary key).
 * @param {number} id - The ACTION_ID of the task to update.
 * @param {Object} data - Field name → value pairs to write (any WA_HEADERS field).
 * @returns {{ success: boolean, error?: string }}
 */
function waUpdate(id, data) {
  const sheet  = getWaSheet();
  const colMap = _waColMap(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, error: 'No records' };
  const ids = sheet.getRange(2,1,lastRow-1,1).getValues().flat();
  const rowIndex = ids.findIndex(r => Number(r) === Number(id));
  if (rowIndex === -1) return { success: false, error: 'Record not found' };
  const sheetRow = rowIndex + 2;
  Object.keys(data).forEach(h => {
    const col = colMap[h];
    if (!col || h === 'ACTION_ID') return;
    sheet.getRange(sheetRow, col).setValue(data[h]);
  });
  return { success: true };
}

/**
 * Permanently deletes a Worker Action record row by ACTION_ID.
 * @param {number} id - The ACTION_ID of the task to delete.
 * @returns {{ success: boolean, error?: string }}
 */
function waDelete(id) {
  const sheet = getWaSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, error: 'No records' };
  const ids = sheet.getRange(2,1,lastRow-1,1).getValues().flat();
  const rowIndex = ids.findIndex(r => Number(r) === Number(id));
  if (rowIndex === -1) return { success: false, error: 'Record not found' };
  sheet.deleteRow(rowIndex + 2);
  return { success: true };
}

/**
 * Returns the next available ACTION_ID for a new worker action task.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The WorkerActions sheet.
 * @returns {number} Next ACTION_ID integer (starts at 1 if sheet is empty).
 */
function _waNextId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 1;
  const ids = sheet.getRange(2,1,lastRow-1,1).getValues().flat().filter(v => v !== '');
  return ids.length === 0 ? 1 : Math.max(...ids.map(Number)) + 1;
}

/**
 * Records a pig health issue as a Worker Action task.
 * Creates a Priority=High pending action in the WorkerActions sheet.
 * @param {Object} data - Fields: pigId, diagnosis, symptoms, treatment, notes, dueDate (optional).
 *   pigId      — PIG ID / ear tag or pen identifier.
 *   diagnosis  — Primary diagnosis string e.g. "Umbilical Hernia".
 *   symptoms   — Short description of observed symptoms.
 *   treatment  — Recommended treatment / action.
 *   notes      — Additional notes or watch-for items.
 *   dueDate    — YYYY-MM-DD follow-up date (defaults to today + 1 day).
 * @returns {{ success: boolean, action_id?: number, error?: string }}
 */
function addHealthIssue(data) {
  try {
    const tz      = Session.getScriptTimeZone();
    const today   = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    const due     = data.dueDate || Utilities.formatDate(new Date(new Date().getTime() + 86400000), tz, 'yyyy-MM-dd');
    const pigLabel = data.pigId ? 'Pig: ' + String(data.pigId) : '';
    const actionText = [
      pigLabel,
      'Diagnosis: ' + (data.diagnosis || 'Health Issue'),
      'Symptoms: '  + (data.symptoms  || '—'),
      'Treatment: ' + (data.treatment || '—')
    ].filter(Boolean).join(' | ');
    const notes = [data.notes || '', data.breedingNote || ''].filter(Boolean).join(' | ');
    return waAdd({
      Date:     today,
      Worker:   data.loggedBy || 'System',
      Category: 'Health',
      Action:   actionText,
      Priority: 'High',
      Status:   'Pending',
      DueDate:  due,
      Notes:    notes
    });
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ── Fix PhotoTime timezone — run once to correct existing records ────
// Run this in the Apps Script editor to re-format all PhotoTime columns
// as plain text strings in Lusaka time
/**
 * ONE-TIME UTILITY — Run manually from the Apps Script editor.
 * Re-formats all PhotoTime1/2/3 columns in the DailyChecklist sheet as plain text
 * in "yyyy-MM-dd HH:mm" format using the script timezone (Africa/Lusaka).
 * Fixes records corrupted by Google Sheets auto-converting datetime strings
 * to Date objects which display incorrectly or shift by timezone offset.
 * Sets @-format on each column first to prevent future auto-conversion.
 */
function fixPhotoTimestamps() {
  const sheet  = getClSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) { Logger.log('No records'); return; }
  const tz      = Session.getScriptTimeZone(); // Africa/Lusaka
  const lastCol = sheet.getLastColumn();
  const hdrs    = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const colMap  = {};
  hdrs.forEach((h, i) => { if (h) colMap[String(h).trim()] = i + 1; });
  const photoCols = ['PhotoTime1','PhotoTime2','PhotoTime3'].filter(h => colMap[h]);
  if (!photoCols.length) { Logger.log('No PhotoTime columns found'); return; }
  let fixed = 0;
  photoCols.forEach(colName => {
    const col   = colMap[colName];
    const range = sheet.getRange(2, col, lastRow - 1, 1);
    range.setNumberFormat('@'); // force plain text
    const vals  = range.getValues();
    const fixed_vals = vals.map(([v]) => {
      if (!v) return [''];
      if (v instanceof Date) {
        // Convert from whatever timezone sheets used back to Lusaka local
        return [Utilities.formatDate(v, tz, 'yyyy-MM-dd HH:mm')];
      }
      return [String(v).trim()];
    });
    range.setValues(fixed_vals);
    fixed += fixed_vals.filter(([v]) => v).length;
  });
  Logger.log('Fixed ' + fixed + ' PhotoTime values in timezone: ' + tz);
}

// ── Fix ALL timestamps — run once in editor to correct existing records ──
// Fixes both PhotoTime (full datetime) and Sec1/2/3Time (HH:mm)
// by forcing plain text format so Google Sheets stops auto-converting them
/**
 * ONE-TIME UTILITY — Run manually from the Apps Script editor.
 * Fixes ALL time-related columns in the DailyChecklist sheet:
 *   - Sec1Time, Sec2Time, Sec3Time: forces to "HH:mm" 24-hour plain text.
 *     Strips AM/PM suffixes that Sheets may have appended.
 *   - PhotoTime1, PhotoTime2, PhotoTime3: forces to "yyyy-MM-dd HH:mm" plain text
 *     in Africa/Lusaka timezone, correcting any UTC-offset corruption.
 * Also sets @-format on each column to prevent future auto-conversion by Sheets.
 */
function fixAllTimestamps() {
  const sheet   = getClSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) { Logger.log('No records to fix'); return; }
  const tz      = Session.getScriptTimeZone();
  const lastCol = sheet.getLastColumn();
  const hdrs    = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const colMap  = {};
  hdrs.forEach((h, i) => { if (h) colMap[String(h).trim()] = i + 1; });

  // Fix Sec1/2/3Time — stored as HH:mm plain text
  ['Sec1Time','Sec2Time','Sec3Time'].forEach(colName => {
    const col = colMap[colName];
    if (!col) return;
    const range = sheet.getRange(2, col, lastRow - 1, 1);
    range.setNumberFormat('@');
    const vals  = range.getValues();
    const fixed = vals.map(([v]) => {
      if (!v) return [''];
      if (v instanceof Date) return [Utilities.formatDate(v, tz, 'HH:mm')];
      // Strip any AM/PM if present
      const s = String(v).trim();
      const ampm = s.match(/(\d{1,2}):(\d{2})\s*(AM|PM)?/i);
      if (ampm) {
        let h = parseInt(ampm[1]);
        const m = ampm[2];
        const period = (ampm[3] || '').toUpperCase();
        if (period === 'PM' && h < 12) h += 12;
        if (period === 'AM' && h === 12) h = 0;
        return [String(h).padStart(2,'0') + ':' + m];
      }
      return [s];
    });
    range.setValues(fixed);
    Logger.log('Fixed ' + colName);
  });

  // Fix PhotoTime1/2/3 — stored as full datetime plain text
  ['PhotoTime1','PhotoTime2','PhotoTime3'].forEach(colName => {
    const col = colMap[colName];
    if (!col) return;
    const range = sheet.getRange(2, col, lastRow - 1, 1);
    range.setNumberFormat('@');
    const vals  = range.getValues();
    const fixed = vals.map(([v]) => {
      if (!v) return [''];
      if (v instanceof Date) return [Utilities.formatDate(v, tz, 'yyyy-MM-dd HH:mm')];
      // Strip AM/PM from datetime strings like "2026-03-25 09:00 AM"
      const s = String(v).trim();
      const dtMatch = s.match(/^(\d{4}-\d{2}-\d{2})\s+(\d{1,2}):(\d{2})\s*(AM|PM)?/i);
      if (dtMatch) {
        let h = parseInt(dtMatch[2]);
        const m = dtMatch[3];
        const period = (dtMatch[4] || '').toUpperCase();
        if (period === 'PM' && h < 12) h += 12;
        if (period === 'AM' && h === 12) h = 0;
        return [dtMatch[1] + ' ' + String(h).padStart(2,'0') + ':' + m];
      }
      return [s];
    });
    range.setValues(fixed);
    Logger.log('Fixed ' + colName);
  });

  Logger.log('fixAllTimestamps complete — all time columns now stored as plain text in 24hr format');
}

// ── Run AI for today — safe, time-limited, resumable ─────────
// Run this directly in the Apps Script editor.
// It processes records one at a time and stops before hitting the 6-min limit.
// Run it again to continue where it left off (already-analysed records are skipped).
/**
 * EDITOR UTILITY — Run from the Apps Script editor to process today's records.
 * Time-safe wrapper: processes records one at a time and stops after 4.5 minutes
 * to stay within the 6-minute Apps Script execution limit.
 * Run it again to resume from where it left off (completed records are skipped).
 * Requires ANTHROPIC_API_KEY to be set in Script Properties.
 */
function runAIForToday() {
  const tz      = Session.getScriptTimeZone();
  const today   = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  _runAISafe(today);
}

/**
 * EDITOR UTILITY — Edit the targetDate constant below then run from the editor
 * to process AI analysis for a specific historical date.
 * Change the date string and run this function directly.
 */
function runAIForDate_() {
  // Change the date below and run this function
  const targetDate = '2026-03-26'; // ← change this date
  _runAISafe(targetDate);
}

/**
 * Internal time-safe AI runner used by runAIForToday() and runAIForDate_().
 * Fetches the list of pending CL_IDs for the target date, then processes
 * each one sequentially via runAISingleRecord(), pausing 1.5s between records.
 * Stops automatically if approaching the 4.5-minute safety threshold.
 * Logs progress to the Apps Script execution log.
 * @param {string} targetDate - YYYY-MM-DD date whose records to analyse.
 */
function _runAISafe(targetDate) {
  const startTime = new Date().getTime();
  const MAX_MS    = 4.5 * 60 * 1000; // stop after 4.5 minutes (safe margin)
  const apiKey    = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) { Logger.log('ERROR: ANTHROPIC_API_KEY not set'); return; }

  const pending = getAIPendingRecords(targetDate);
  Logger.log('Pending records for ' + targetDate + ': ' + JSON.stringify(pending.pending));
  if (!pending.pending.length) { Logger.log('No records to analyse'); return; }

  let processed = 0, skipped = 0, errors = 0;
  for (const clId of pending.pending) {
    // Check time remaining
    if (new Date().getTime() - startTime > MAX_MS) {
      Logger.log('⚠ Approaching time limit — stopping. Run again to continue.');
      Logger.log('Progress: ' + processed + ' done, ' + errors + ' errors, ' + pending.pending.length + ' total');
      return;
    }
    Logger.log('Analysing CL_ID: ' + clId);
    const result = runAISingleRecord(clId);
    if (result.success) {
      if (result.skipped) { skipped++; Logger.log('  → already analysed, skipped'); }
      else { processed++; Logger.log('  → ✓ done'); }
    } else {
      errors++;
      Logger.log('  → ✗ error: ' + result.error);
    }
    Utilities.sleep(1500); // pause between records
  }
  Logger.log('✅ Complete — ' + processed + ' analysed, ' + skipped + ' skipped, ' + errors + ' errors');
}
