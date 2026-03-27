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

function getNextId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 1;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(v => v !== "");
  if (ids.length === 0) return 1;
  return Math.max(...ids.map(Number)) + 1;
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function corsRespond(data) {
  // For POST requests wrapped in JSONP-style — Apps Script handles CORS via GET
  return respond(data);
}

// ── Router ───────────────────────────────────────────────────

// ── Request token validation ─────────────────────
// Set APP_TOKEN in Script Properties to enable validation.
// Leave unset to allow all requests (backward compatible).
function _validateToken(token) {
  const expected = PropertiesService.getScriptProperties().getProperty('APP_TOKEN');
  if (!expected) return true; // no token set — allow all (backward compatible)
  return String(token || '') === String(expected);
}

function doGet(e) {
  const action = e.parameter.action;
  if (action !== 'ping' && !_validateToken(e.parameter.token)) {
    return respond({ error: 'Unauthorized' });
  }
  try {
    if (action === "getAll")       return respond(getAllRecords());
    if (action === "ping")         return respond({ success: true, message: "pong", time: new Date().toISOString() });
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

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    if (!_validateToken(payload.token)) {
      return corsRespond({ error: 'Unauthorized' });
    }
    if (action === "add")      return corsRespond(addRecord(payload.data));
    if (action === "update")   return corsRespond(updateRecord(payload.id, payload.data));
    if (action === "delete")   return corsRespond(deleteRecord(payload.id));
    if (action === "clAdd")       return corsRespond(clAdd(payload.data));
    if (action === "clUpsert")    return corsRespond(clUpsert(payload.data));
    if (action === "clUpdate")    return corsRespond(clUpdate(payload.id, payload.data));
    if (action === "clDelete")    return corsRespond(clDelete(payload.id));
    if (action === "clSavePhoto") return corsRespond(clSavePhoto(payload.clId, payload.photoBase64, payload.mimeType, payload.photoTime, payload.section));
    if (action === "slAdd")    return corsRespond(slAdd(payload.data));
    if (action === "slUpsert") return corsRespond(slUpsert(payload.data));
    if (action === "slUpdate") return corsRespond(slUpdate(payload.id, payload.data));
    if (action === "slDelete") return corsRespond(slDelete(payload.id));
    if (action === "wkAdd")    return corsRespond(wkAdd(payload.data));
    if (action === "wkUpdate") return corsRespond(wkUpdate(payload.id, payload.data));
    if (action === "wkDelete") return corsRespond(wkDelete(payload.id));
    if (action === "moAdd")    return corsRespond(moAdd(payload.data));
    if (action === "moUpsert") return corsRespond(moUpsert(payload.data));
    if (action === "moUpdate") return corsRespond(moUpdate(payload.id, payload.data));
    if (action === "moDelete") return corsRespond(moDelete(payload.id));
    if (action === "saveSetting") return corsRespond(saveSetting(payload.key, payload.value));
    if (action === "waAdd")       return corsRespond(waAdd(payload.data));
    if (action === "waUpdate")    return corsRespond(waUpdate(payload.id, payload.data));
    if (action === "waDelete")    return corsRespond(waDelete(payload.id));
    if (action === "migrateBoarSow")   return corsRespond(migrateBoarSowToDbId());
    if (action === "migrateSowIds")    return corsRespond(migrateSowLitterSowId());
    if (action === "runAIAnalysis")    return corsRespond(runNightlyAIAnalysis(payload.targetDate || null, true));
    if (action === "runAISingle")      return corsRespond(runAISingleRecord(payload.clId));
    if (action === "runCloseout")      return corsRespond(runCloseoutOnly());
    return corsRespond({ error: "Unknown action" });
  } catch (err) {
    return corsRespond({ error: err.message });
  }
}

// ── CRUD Operations ──────────────────────────────────────────

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

function _getPigLogHeaders(sheet) {
  const lastCol = sheet.getLastColumn();
  const hdrRow  = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map     = {};
  hdrRow.forEach((h, i) => { if (h) map[String(h).trim()] = i; });
  return { map, lastCol, hdrRow };
}

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

function getNextClId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 1;
  const ids = sheet.getRange(2,1,lastRow-1,1).getValues().flat().filter(v => v !== "");
  if (ids.length === 0) return 1;
  return Math.max(...ids.map(Number)) + 1;
}

const CL_TIME_COLS = ["Sec1Time","Sec2Time","Sec3Time"];

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
function clCount() {
  const sheet = getClSheet();
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  return { success: true, rows: lastRow - 1, cols: lastCol };
}

// Return only records from the last N days — much faster than clGetAll for large sheets
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
function _clSheetColMap(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return {};
  const hdrs = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map  = {};
  hdrs.forEach((h, i) => { if (h) map[String(h).trim()] = i + 1; });
  return map;
}

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
function testDriveAccess() {
  const folder = DriveApp.getRootFolder();
  Logger.log("Drive access OK. Root folder: " + folder.getName());
}

// Fetch a Drive file and return it as a base64 data URL — bypasses CORS for <img> tags
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
  "sl_foster","sl_total_piglets","sl_mortality",
  "ByLitter","ByNursing","BySowTreat","ByMilestones",
  "ByD01","ByD23","ByD57","ByD710","ByD14","ByD2128",
  "SlTime_shdr_litter","SlTime_shdr_sowtreat",
  "SlTime_mhdr_d01","SlTime_mhdr_d23","SlTime_mhdr_d57",
  "SlTime_mhdr_d710","SlTime_mhdr_d14","SlTime_mhdr_d2128",
  "sl_born_alive","sl_stillborn","sl_mummified","sl_total_birth_wt",
  "sl_lightest","sl_heaviest","sl_nursing","sl_weaklings",
  "sl_lightest_today","sl_heaviest_today","sl_castrated",
  "sl_wt_d14","sl_alive_d14","sl_num_weaned","sl_date_weaned","sl_wean_wt",
  ...SL_TASK_COLS];

const SL_TIME_COLS = [
  "SlTime_shdr_litter","SlTime_shdr_sowtreat",
  "SlTime_mhdr_d01","SlTime_mhdr_d23","SlTime_mhdr_d57",
  "SlTime_mhdr_d710","SlTime_mhdr_d14","SlTime_mhdr_d2128"
];

function getSlSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SL_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SL_SHEET);
    sheet.appendRow(SL_HEADERS_GS);
    sheet.getRange(1,1,1,SL_HEADERS_GS.length).setFontWeight("bold").setBackground("#880e4f").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function slGetAll() {
  const sheet = getSlSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, records: [] };
  const tz         = Session.getScriptTimeZone(); // cache once
  const lastCol    = sheet.getLastColumn();
  const sheetHdrs  = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v).trim());
  const data       = sheet.getRange(2, 1, lastRow-1, lastCol).getValues();
  const validTime  = t => /^\d{1,2}:\d{2}$/.test(String(t).trim());
  const records = data.filter(r => r[0] !== '').map(row => {
    const byName = {};
    sheetHdrs.forEach((h, i) => { if (h) byName[h] = row[i]; });
    const rec = {};
    SL_HEADERS_GS.forEach(h => {
      const v = byName.hasOwnProperty(h) ? byName[h] : '';
      if (SL_TIME_COLS.indexOf(h) >= 0) {
        if (v instanceof Date) {
          rec[h] = String(v.getHours()).padStart(2,'0') + ':' + String(v.getMinutes()).padStart(2,'0');
        } else {
          rec[h] = validTime(v) ? String(v).trim() : '';
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
function _slSheetColMap(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return {};
  const hdrs = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map  = {};
  hdrs.forEach((h, i) => { if (h) map[String(h).trim()] = i + 1; });
  return map;
}

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
  return { success: true, sl_id: newId };
}

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
    const v = SL_TIME_COLS.includes(h) ? String(data[h]) : data[h];
    sheet.getRange(sheetRow, col).setValue(v);
  });
  return { success: true };
}

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

// ============================================================
//  WEEKLY CHECKLIST — Sheet: WeeklyChecklist
// ============================================================

const WK_SHEET = "WeeklyChecklist";
const WK_KEYS_GS = ['wk_weight_spot','wk_body_condition','wk_behaviour',
  'wk_deworm','wk_vitamins','wk_wounds',
  'wk_feeder_clean','wk_pen_repair','wk_feed_consumption','wk_boar_condition'];
const WK_HEADERS_GS = ["WK_ID","Date","WeekNum","WeekYear","WeekKey","Pen","CheckedBy","Status","Concerns","Notes",
  "AvgWeight","FeedKg","ByCondition","ByHealth","ByFarm",...WK_KEYS_GS];

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

function _wkSheetColMap(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return {};
  const hdrs = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = {};
  hdrs.forEach((h, i) => { if (h) map[String(h).trim()] = i + 1; });
  return map;
}

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

function _moGetSheetHeaders(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return {};
  const raw = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = {};
  raw.forEach((h, i) => { if (h) map[String(h).trim()] = i; });
  return map;
}

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

function getSetting(key) {
  const sheet = getSettingsSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, value: null };
  const data = sheet.getRange(2, 1, lastRow-1, 2).getValues();
  const row = data.find(r => String(r[0]) === String(key));
  return { success: true, value: row ? String(row[1]) : null };
}

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
function runNightlyAIAnalysis(targetDate, forceAI) {
  return runNightlyCloseout(targetDate, forceAI);
}

  // Default: process yesterday's records (today's are still in progress)
// ── Analyse a single record by CL_ID ─────────────────────────
function runAISingleRecord(clId) {
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
  // Skip if already analysed
  const aiCol = hdrMap['AIAnalysis'];
  if (aiCol !== undefined && String(row[aiCol] || '').trim()) {
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

// ── Mark Incomplete — closes all In Progress records older than today ───
function _markIncompleteRecords(processDate, sheet, hdrMap, data) {
  const tz        = Session.getScriptTimeZone();
  const today     = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const statusCol = hdrMap['Status'];
  const dateCol   = hdrMap['Date'];
  if (statusCol === undefined || dateCol === undefined) return 0;
  let count = 0;
  data.forEach((row, i) => {
    const sheetRow   = i + 2;
    const rowDate    = row[dateCol];
    const rowDateStr = rowDate instanceof Date
      ? Utilities.formatDate(rowDate, tz, 'yyyy-MM-dd')
      : String(rowDate || '').trim();
    // Mark any record from BEFORE today that is still open
    if (!rowDateStr || rowDateStr >= today) return;
    const status = String(row[statusCol] || '').trim();
    if (status === 'In Progress' || status === 'PARTIAL' || status === '') {
      sheet.getRange(sheetRow, statusCol + 1).setValue('Incomplete');
      count++;
      Logger.log('Marked Incomplete: row ' + sheetRow + ' date ' + rowDateStr);
    }
  });
  return count;
}

// ── Per-Record Analysis ─────────────────────────────────────

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

      // Always fetch a compressed thumbnail for AI — never the full file
      // sz=w800 gives ~200-400KB which is well under the 5MB API limit
      const thumbUrl = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w800';
      const token    = ScriptApp.getOAuthToken();
      const response = UrlFetchApp.fetch(thumbUrl, {
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true
      });

      let blob;
      if (response.getResponseCode() === 200) {
        blob = response.getBlob();
        Logger.log('Photo ' + urlKey + ': thumbnail fetched, size=' + blob.getBytes().length + ' bytes');
      } else {
        // Thumbnail failed — try the original but only if under 4MB
        Logger.log('Thumbnail fetch failed (' + response.getResponseCode() + ') for ' + urlKey + ' — trying original');
        const file = DriveApp.getFileById(fileId);
        blob = file.getBlob();
        const size = blob.getBytes().length;
        if (size > 4 * 1024 * 1024) {
          Logger.log('Original too large (' + size + ' bytes) — skipping ' + urlKey);
          messageContent.push({ type: 'text', text: ['[Morning photo — too large to send]','[Feeding photo — too large to send]','[Afternoon photo — too large to send]'][i] });
          return;
        }
      }

      // Final size check before sending
      const finalBytes = blob.getBytes();
      if (finalBytes.length > 4.5 * 1024 * 1024) {
        Logger.log('Photo still too large after thumbnail (' + finalBytes.length + ') — skipping');
        messageContent.push({ type: 'text', text: ['[Morning photo — could not compress]','[Feeding photo — could not compress]','[Afternoon photo — could not compress]'][i] });
        return;
      }

      const mime = blob.getContentType() || 'image/jpeg';
      const b64  = Utilities.base64Encode(finalBytes);
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
    model:      'claude-sonnet-4-5',
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

function _waColMap(sheet) {
  const hdrs = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const map = {};
  hdrs.forEach((h,i) => { if (h) map[String(h).trim()] = i+1; });
  return map;
}

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

function _waNextId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 1;
  const ids = sheet.getRange(2,1,lastRow-1,1).getValues().flat().filter(v => v !== '');
  return ids.length === 0 ? 1 : Math.max(...ids.map(Number)) + 1;
}

// ── Fix PhotoTime timezone — run once to correct existing records ────
// Run this in the Apps Script editor to re-format all PhotoTime columns
// as plain text strings in Lusaka time
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
function runAIForToday() {
  const tz      = Session.getScriptTimeZone();
  const today   = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  _runAISafe(today);
}

function runAIForDate_() {
  // Change the date below and run this function
  const targetDate = '2026-03-26'; // ← change this date
  _runAISafe(targetDate);
}

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
