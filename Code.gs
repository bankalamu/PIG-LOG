// ============================================================
//  PIG LOG — Google Apps Script Backend
//  Paste this entire file into your Apps Script project
// ============================================================

const SHEET_NAME = "PigLog";
const HEADERS = ["DB_ID", "PIG ID", "Boar", "SOW", "DOB", "SEX", "Status", "Weight", "Dewormed", "Pen", "Notes"];

// ── Helpers ──────────────────────────────────────────────────

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold").setBackground("#c9a84c").setFontColor("#000000");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 70);   // DB_ID
    sheet.setColumnWidth(2, 120);  // PIG ID
    sheet.setColumnWidth(11, 250); // Notes
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
    if (action === "clAdd")    return corsRespond(clAdd(payload.data));
    if (action === "clUpdate") return corsRespond(clUpdate(payload.id, payload.data));
    if (action === "clDelete") return corsRespond(clDelete(payload.id));
    if (action === "slAdd")    return corsRespond(slAdd(payload.data));
    if (action === "slUpsert") return corsRespond(slUpsert(payload.data));
    if (action === "slUpdate") return corsRespond(slUpdate(payload.id, payload.data));
    if (action === "slDelete") return corsRespond(slDelete(payload.id));
    if (action === "wkAdd")    return corsRespond(wkAdd(payload.data));
    if (action === "wkUpdate") return corsRespond(wkUpdate(payload.id, payload.data));
    if (action === "wkDelete") return corsRespond(wkDelete(payload.id));
    if (action === "moAdd")    return corsRespond(moAdd(payload.data));
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
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, records: [] };

  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  const records = data
    .filter(row => row[0] !== "")
    .map(row => {
      const rec = {};
      HEADERS.forEach((h, i) => {
        // Format dates nicely
        if (row[i] instanceof Date) {
          rec[h] = Utilities.formatDate(row[i], Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else {
          rec[h] = row[i];
        }
      });
      return rec;
    });

  return { success: true, records };
}

function getByPigId(pigId) {
  if (!pigId) return { success: false, error: "No PIG ID provided" };
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, error: "No records found" };

  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  const row = data.find(r => String(r[1]).toLowerCase() === String(pigId).toLowerCase());

  if (!row) return { success: false, error: `No record found for PIG ID: "${pigId}"` };

  const rec = {};
  HEADERS.forEach((h, i) => {
    if (row[i] instanceof Date) {
      rec[h] = Utilities.formatDate(row[i], Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else {
      rec[h] = row[i];
    }
  });
  return { success: true, record: rec };
}

function addRecord(data) {
  const sheet = getSheet();
  const newId = getNextId(sheet);
  const row = HEADERS.map((h, i) => {
    if (h === "DB_ID") return newId;
    return data[h] !== undefined ? data[h] : "";
  });
  sheet.appendRow(row);
  return { success: true, db_id: newId, message: "Record added successfully" };
}

function updateRecord(dbId, data) {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, error: "No records found" };

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const rowIndex = ids.findIndex(id => Number(id) === Number(dbId));

  if (rowIndex === -1) return { success: false, error: `Record with DB_ID ${dbId} not found` };

  const sheetRow = rowIndex + 2; // +2 for header row and 0-index
  HEADERS.forEach((h, colIndex) => {
    if (h === "DB_ID") return; // Never overwrite DB_ID
    if (data[h] !== undefined) {
      sheet.getRange(sheetRow, colIndex + 1).setValue(data[h]);
    }
  });

  return { success: true, message: "Record updated successfully" };
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
  const data = sheet.getRange(2,1,lastRow-1,CL_HEADERS.length).getValues();
  const records = data.filter(r => r[0] !== "").map(rowToClRecord);
  return { success: true, records };
}

function clAdd(data) {
  const sheet = getClSheet();
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

// ============================================================
//  SOW & LITTER — Sheet: SowLitter
// ============================================================

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
const WK_HEADERS_GS = ["WK_ID","Date","Pen","CheckedBy","Status","Concerns","Notes",
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
  const sheet = getWkSheet();
  const ids = sheet.getLastRow() <= 1 ? [] : sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues().flat().filter(v=>v!=="");
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
const MO_HEADERS_GS = ["MO_ID","Month","CheckedBy","Status","Concerns","Notes",
  "PigsWeighed","AvgADG","Deaths","VaxCount","FeedStock",
  "ByGrowth","ByHealth","ByFarm",...MO_KEYS_GS];

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

function moGetAll() {
  const sheet = getMoSheet(); const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, records: [] };
  const data = sheet.getRange(2,1,lastRow-1,MO_HEADERS_GS.length).getValues();
  return { success: true, records: data.filter(r=>r[0]!=='').map(row => {
    const rec = {};
    MO_HEADERS_GS.forEach((h,i) => { rec[h] = row[i]; });
    return rec;
  })};
}

function moAdd(data) {
  const sheet = getMoSheet();
  const ids = sheet.getLastRow() <= 1 ? [] : sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues().flat().filter(v=>v!=="");
  const newId = ids.length === 0 ? 1 : Math.max(...ids.map(Number)) + 1;
  sheet.appendRow(MO_HEADERS_GS.map(h => h==="MO_ID" ? newId : (data[h]!==undefined ? data[h] : "")));
  return { success: true, mo_id: newId };
}

function moUpdate(moId, data) {
  const sheet = getMoSheet(); const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, error: "No records" };
  const ids = sheet.getRange(2,1,lastRow-1,1).getValues().flat();
  const rowIndex = ids.findIndex(id => Number(id) === Number(moId));
  if (rowIndex === -1) return { success: false, error: "Not found" };
  MO_HEADERS_GS.forEach((h,i) => { if(h!=="MO_ID" && data[h]!==undefined) sheet.getRange(rowIndex+2,i+1).setValue(data[h]); });
  return { success: true };
}

function moDelete(moId) {
  const sheet = getMoSheet(); const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, error: "No records" };
  const ids = sheet.getRange(2,1,lastRow-1,1).getValues().flat();
  const rowIndex = ids.findIndex(id => Number(id) === Number(moId));
  if (rowIndex === -1) return { success: false, error: "Not found" };
  sheet.deleteRow(rowIndex + 2);
  return { success: true };
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
    // seed default
    sheet.appendRow(["weaningWeeks", "5", "system", new Date()]);
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
