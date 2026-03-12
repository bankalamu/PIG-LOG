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
const HEADERS = ["DB_ID", "PIG ID", "Boar", "SOW", "DOB", "SEX", "Type", "Stage", "Status", "Weight", "Dewormed", "Pen", "Notes", "Available"];

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

function doGet(e) {
  const action = e.parameter.action;
  try {
    if (action === "getAll")     return respond(getAllRecords());
    if (action === "getByPigId") return respond(getByPigId(e.parameter.pigId));
    if (action === "clGetAll")   return respond(clGetAll());
    if (action === "slGetAll")   return respond(slGetAll());
    if (action === "wkGetAll")   return respond(wkGetAll());
    if (action === "moGetAll")   return respond(moGetAll());
    if (action === "moDedup")    return respond(moDeduplicateAll());
    if (action === "getSetting") return respond(getSetting(e.parameter.key));
    return respond({ error: "Unknown action" });
  } catch (err) {
    return respond({ error: err.message });
  }
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    if (action === "add")      return corsRespond(addRecord(payload.data));
    if (action === "update")   return corsRespond(updateRecord(payload.id, payload.data));
    if (action === "delete")   return corsRespond(deleteRecord(payload.id));
    if (action === "clAdd")       return corsRespond(clAdd(payload.data));
    if (action === "clUpsert")    return corsRespond(clUpsert(payload.data));
    if (action === "clUpdate")    return corsRespond(clUpdate(payload.id, payload.data));
    if (action === "clDelete")    return corsRespond(clDelete(payload.id));
    if (action === "clSavePhoto") return corsRespond(clSavePhoto(payload.clId, payload.photoBase64, payload.mimeType, payload.photoTime));
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

  const lastCol  = sheet.getLastColumn();
  const hdrRow   = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const hdrMap   = {}; // colName → 0-based index
  hdrRow.forEach((h, i) => { if (h) hdrMap[String(h).trim()] = i; });

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const records = data
    .filter(row => row[0] !== "" && row[0] !== null && row[0] !== undefined)
    .map(row => {
      const rec = {};
      // Always include all known HEADERS (fill missing ones with "")
      HEADERS.forEach(h => {
        const idx = hdrMap[h];
        if (idx === undefined) { rec[h] = ""; return; }
        const v = row[idx];
        rec[h] = (v instanceof Date)
          ? Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd")
          : (v === null || v === undefined ? "" : v);
      });
      // Also expose any extra columns in the sheet not in HEADERS
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
  HEADERS.forEach(h => {
    const idx = hdrMap[h];
    if (idx === undefined) { rec[h] = ""; return; }
    const v = row[idx];
    rec[h] = (v instanceof Date)
      ? Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd")
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

  // All three key fields required
  const pigId = String(data["PIG ID"] || "").trim();
  const boar  = String(data["Boar"]   || "").trim();
  const sow   = String(data["SOW"]    || "").trim();
  if (!pigId || !boar || !sow) {
    return { success: false, error: "PIG ID, Boar and SOW are all required." };
  }

  const { map, lastCol } = _getPigLogHeaders(sheet);
  const pidIdx  = map["PIG ID"] !== undefined ? map["PIG ID"] : 1;
  const borIdx  = map["Boar"]   !== undefined ? map["Boar"]   : 2;
  const sowIdx  = map["SOW"]    !== undefined ? map["SOW"]    : 3;

  // Duplicate check using live header positions
  if (lastRow > 1) {
    const existing = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
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
const CL_HEADERS = ["CL_ID","Date","Pen","CheckedBy","Status","Concerns","Notes","PhotoUrl","PhotoTime","PigCount",
                    "Sec1Time","Sec2Time","Sec3Time",
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

function rowToClRecord(row) {
  const rec = {};
  CL_HEADERS.forEach((h,i) => {
    if (row[i] instanceof Date) {
      rec[h] = Utilities.formatDate(row[i], Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else {
      rec[h] = row[i];
    }
  });
  return rec;
}

// ── Checklist CRUD ──

function clGetAll() {
  const sheet = getClSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, records: [] };
  const lastCol = Math.max(sheet.getLastColumn(), CL_HEADERS.length);
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const records = data.filter(r => r[0] !== "" && r[0] !== null && r[0] !== undefined)
                      .map(row => rowToClRecord(row));
  return { success: true, records };
}

function clUpsert(data) {
  const sheet = getClSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const dateIdx = CL_HEADERS.indexOf("Date");
    const penIdx  = CL_HEADERS.indexOf("Pen");
    const allData = sheet.getRange(2, 1, lastRow - 1, CL_HEADERS.length).getValues();
    for (let i = 0; i < allData.length; i++) {
      const rowDate = allData[i][dateIdx] instanceof Date
        ? Utilities.formatDate(allData[i][dateIdx], Session.getScriptTimeZone(), "yyyy-MM-dd")
        : String(allData[i][dateIdx] || '').trim();
      const rowPen = String(allData[i][penIdx] || '').trim().toLowerCase();
      const inPen  = String(data.Pen || '').trim().toLowerCase();
      if (rowDate === String(data.Date || '').trim() && rowPen === inPen) {
        const existingId = allData[i][0];
        clUpdate(existingId, data);
        return { success: true, cl_id: existingId, updated: true };
      }
    }
  }
  return clAdd(data);
}

function clAdd(data) {
  const sheet = getClSheet();
  // Server-side duplicate guard: reject if same Date+Pen already exists
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const dateIdx = CL_HEADERS.indexOf("Date");
    const penIdx  = CL_HEADERS.indexOf("Pen");
    const allData = sheet.getRange(2, 1, lastRow - 1, CL_HEADERS.length).getValues();
    const inDate  = String(data.Date || '').trim();
    const inPen   = String(data.Pen  || '').trim().toLowerCase();
    const dup = allData.find(r => {
      const rowDate = r[dateIdx] instanceof Date
        ? Utilities.formatDate(r[dateIdx], Session.getScriptTimeZone(), "yyyy-MM-dd")
        : String(r[dateIdx] || '').trim();
      return rowDate === inDate && String(r[penIdx] || '').trim().toLowerCase() === inPen;
    });
    if (dup) return { success: false, error: `Pen ${data.Pen} already has a record for ${data.Date}. Use update instead.` };
  }
  const newId = getNextClId(sheet);
  const row = CL_HEADERS.map(h => {
    if (h === "CL_ID") return newId;
    return data[h] !== undefined ? data[h] : "";
  });
  sheet.appendRow(row);
  return { success: true, cl_id: newId };
}

function clUpdate(clId, data) {
  const sheet = getClSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, error: "No records" };
  const ids = sheet.getRange(2,1,lastRow-1,1).getValues().flat();
  const rowIndex = ids.findIndex(id => Number(id) === Number(clId));
  if (rowIndex === -1) return { success: false, error: "Record not found" };
  const sheetRow = rowIndex + 2;
  CL_HEADERS.forEach((h, colIndex) => {
    if (h === "CL_ID") return;
    if (data[h] !== undefined) sheet.getRange(sheetRow, colIndex+1).setValue(data[h]);
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

function clSavePhoto(clId, photoBase64, mimeType, photoTime) {
  try {
    if (!clId || !photoBase64) return { success: false, error: "Missing clId or photo data" };

    // Get or create PigLog Photos folder in Drive
    const folderName = "PigLog_Photos";
    const folders = DriveApp.getFoldersByName(folderName);
    const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

    // Decode base64 and save file
    const ext  = (mimeType || 'image/jpeg').split('/')[1] || 'jpg';
    const blob = Utilities.newBlob(Utilities.base64Decode(photoBase64), mimeType || 'image/jpeg',
                                   'pen_photo_cl' + clId + '_' + new Date().getTime() + '.' + ext);
    const file = folder.createFile(blob);

    // Make file publicly viewable (anyone with link)
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileId  = file.getId();
    const viewUrl = "https://drive.google.com/file/d/" + fileId + "/view";
    const thumbUrl = "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w400";

    // Update the PhotoUrl and PhotoTime columns in the sheet
    const updateData = { PhotoUrl: viewUrl };
    if (photoTime) updateData.PhotoTime = photoTime;
    const result = clUpdate(clId, updateData);
    if (!result.success) return { success: false, error: "Photo saved to Drive but sheet update failed: " + result.error };

    return { success: true, viewUrl, thumbUrl, fileId };
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

const SL_SHEET = "SowLitter";
const SL_TASK_COLS = ['sow_antiinflam','sow_mma_ab','sow_oxytocin',
  'iodine','colostrum','teat_order','heat_lamp','litter_weigh',
  'iron','tail_dock','teeth','navel_check',
  'multivit','creep','navel_healed',
  'castration','ear_notch','coccidiostat',
  'indiv_weigh','health_d14','mma_d14',
  'vax','weaner_feed','weaner_pen','prev_meds','weaned'];
const SL_HEADERS_GS = ["SL_ID","SowId","FarrowDate","Pen","Notes",
  "ByLitter","ByNursing","BySowTreat","ByMilestones",
  "ByD01","ByD23","ByD57","ByD710","ByD14","ByD2128",
  "sl_born_alive","sl_stillborn","sl_mummified","sl_total_birth_wt",
  "sl_lightest","sl_heaviest","sl_nursing","sl_weaklings",
  "sl_lightest_today","sl_heaviest_today","sl_castrated",
  "sl_wt_d14","sl_alive_d14","sl_num_weaned","sl_date_weaned","sl_wean_wt",
  ...SL_TASK_COLS];

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
  const data = sheet.getRange(2,1,lastRow-1,SL_HEADERS_GS.length).getValues();
  const records = data.filter(r => r[0] !== '').map(row => {
    const rec = {};
    SL_HEADERS_GS.forEach((h,i) => {
      rec[h] = row[i] instanceof Date
        ? Utilities.formatDate(row[i], Session.getScriptTimeZone(), "yyyy-MM-dd")
        : row[i];
    });
    return rec;
  });
  return { success: true, records };
}

// Upsert: update if SowId+FarrowDate exists, otherwise insert
function slUpsert(data) {
  const sheet = getSlSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const allData = sheet.getRange(2, 1, lastRow-1, SL_HEADERS_GS.length).getValues();
    const sowIdIdx  = SL_HEADERS_GS.indexOf("SowId");
    const farrowIdx = SL_HEADERS_GS.indexOf("FarrowDate");
    for (let i = 0; i < allData.length; i++) {
      const rowSow    = String(allData[i][sowIdIdx]||'').trim().toLowerCase();
      const rowFarrow = allData[i][farrowIdx] instanceof Date
        ? Utilities.formatDate(allData[i][farrowIdx], Session.getScriptTimeZone(), "yyyy-MM-dd")
        : String(allData[i][farrowIdx]||'').trim();
      if (rowSow === String(data.SowId||'').trim().toLowerCase() && rowFarrow === String(data.FarrowDate||'').trim()) {
        const existingId = allData[i][0];
        slUpdate(existingId, data);
        return { success: true, sl_id: existingId, updated: true };
      }
    }
  }
  return slAdd(data);
}

function slAdd(data) {
  const sheet = getSlSheet();
  const ids = sheet.getLastRow() <= 1 ? [] : sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues().flat().filter(v=>v!=="");
  const newId = ids.length === 0 ? 1 : Math.max(...ids.map(Number)) + 1;
  const row = SL_HEADERS_GS.map(h => h === "SL_ID" ? newId : (data[h] !== undefined ? data[h] : ""));
  sheet.appendRow(row);
  return { success: true, sl_id: newId };
}

function slUpdate(slId, data) {
  const sheet = getSlSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, error: "No records" };
  const ids = sheet.getRange(2,1,lastRow-1,1).getValues().flat();
  const rowIndex = ids.findIndex(id => Number(id) === Number(slId));
  if (rowIndex === -1) return { success: false, error: "Record not found" };
  const sheetRow = rowIndex + 2;
  SL_HEADERS_GS.forEach((h, colIndex) => {
    if (h === "SL_ID") return;
    if (data[h] !== undefined) sheet.getRange(sheetRow, colIndex+1).setValue(data[h]);
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
  const data = sheet.getRange(2,1,lastRow-1,WK_HEADERS_GS.length).getValues();
  return { success: true, records: data.filter(r=>r[0]!=='').map(row => {
    const rec = {};
    WK_HEADERS_GS.forEach((h,i) => { rec[h] = row[i] instanceof Date ? Utilities.formatDate(row[i], Session.getScriptTimeZone(), "yyyy-MM-dd") : row[i]; });
    return rec;
  })};
}

function wkAdd(data) {
  const sheet   = getWkSheet();
  const lastRow = sheet.getLastRow();
  // Server-side duplicate guard: same WeekKey + Pen
  if (lastRow > 1) {
    const wkKeyIdx = WK_HEADERS_GS.indexOf("WeekKey");
    const penIdx   = WK_HEADERS_GS.indexOf("Pen");
    const allData  = sheet.getRange(2, 1, lastRow - 1, WK_HEADERS_GS.length).getValues();
    const inKey    = String(data.WeekKey || '').trim();
    const inPen    = String(data.Pen     || '').trim().toLowerCase();
    const dup = allData.find(r =>
      String(r[wkKeyIdx]||'').trim() === inKey &&
      String(r[penIdx]  ||'').trim().toLowerCase() === inPen
    );
    if (dup) return { success: false, error: `Pen ${data.Pen} already has a record for ${data.WeekKey}. Use update instead.` };
  }
  const ids   = lastRow <= 1 ? [] : sheet.getRange(2,1,lastRow-1,1).getValues().flat().filter(v=>v!=="");
  const newId = ids.length === 0 ? 1 : Math.max(...ids.map(Number)) + 1;
  sheet.appendRow(WK_HEADERS_GS.map(h => h==="WK_ID" ? newId : (data[h]!==undefined ? data[h] : "")));
  return { success: true, wk_id: newId };
}

function wkUpdate(wkId, data) {
  const sheet = getWkSheet(); const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, error: "No records" };
  const ids = sheet.getRange(2,1,lastRow-1,1).getValues().flat();
  const rowIndex = ids.findIndex(id => Number(id) === Number(wkId));
  if (rowIndex === -1) return { success: false, error: "Not found" };
  WK_HEADERS_GS.forEach((h,i) => { if(h!=="WK_ID" && data[h]!==undefined) sheet.getRange(rowIndex+2,i+1).setValue(data[h]); });
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
