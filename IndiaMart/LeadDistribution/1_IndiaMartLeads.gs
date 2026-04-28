const DAILY_PIPELINE = {
  MAIN_SPREADSHEET_ID: '1OPq4GtU-wrQFUbN9Zvgpc_WO_O6fvhl0MUQKeTz1rGw',
  INDIAMART_CRM_KEY: 'mRy0FrFr4XrITves7nyK7liNqlHEnzBl',
  INDIAMART_URL: 'https://mapi.indiamart.com/wservce/crm/crmListing/v2/',

  SHEETS: {
    CONFIG: 'Config',
    FRESH: 'FreshLeads',
    RAW: 'RawData',
    WORK: 'Work',
    EMP_LEADCALLMSG: 'LeadCallMsg',
    EMP_FOLLOWUP: 'Followup'
  },

  COL: {
    CONFIG_SHEET_ID: 2, // B
    WORK_PHONE: 5, // E
    EMP_PHONE: 5, // E
    EMP_WORKED_FLAG: 8 // H
  }
  ,
  CLEAN_FRESH_AFTER_MOVE: true
};

function setupDaily8AMTrigger() {
  dailyDeleteTriggersByHandler_('runDailyIndiaMartPipeline');

  ScriptApp.newTrigger('runDailyIndiaMartPipeline')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();

  Logger.log('Created daily trigger for 8:00 AM.');
}

function runDailyIndiaMartPipeline() {
  const ss = SpreadsheetApp.openById(DAILY_PIPELINE.MAIN_SPREADSHEET_ID);
  const configSheet = dailyGetOrCreateSheet_(ss, DAILY_PIPELINE.SHEETS.CONFIG);
  const freshSheet = dailyGetOrCreateSheet_(ss, DAILY_PIPELINE.SHEETS.FRESH);
  const rawSheet = dailyGetOrCreateSheet_(ss, DAILY_PIPELINE.SHEETS.RAW);
  const workSheet = dailyGetOrCreateSheet_(ss, DAILY_PIPELINE.SHEETS.WORK);

  // Move any worked rows from employee LeadCallMsg -> Followup before fetching new leads
  const preMoveSummary = dailyMoveLeadCallMsgWorkedToFollowup_(configSheet);

  const leads = dailyFetchIndiaMartLeads_(configSheet);

  // 🚫 STOP ONLY FOR 429
  // if (leads === "RATE_LIMIT") {
  //   Logger.log("Stopped due to API rate limit");
  //   return;
  // }

  // ❌ Other API failure → continue
  if (leads === null) {
    Logger.log("API failed, but continuing with Fresh → Work");
  }

  // ✅ Success but 0 leads
  if (Array.isArray(leads) && leads.length === 0) {
    Logger.log("0 leads fetched");
  }

  // Write fetched API rows to both RawData and FreshLeads before Fresh -> Work move.
  if (Array.isArray(leads) && leads.length > 0) {
    dailyAppendRawLeads_(rawSheet, leads);
    dailyAppendRawLeads_(freshSheet, leads);
  }

  const moved = dailyCleanFreshLeadsToWork_(freshSheet, workSheet);
  dailyClearConfigColumnsCD_(configSheet);
  const syncSummary = dailyProcessEmployeeSheetsAndSyncToWork_(configSheet, workSheet);
  const clearedAssigned = dailyClearWorkAssignedWhereHEmpty_(workSheet);

  const fetchedCount = Array.isArray(leads) ? leads.length : 0;

  Logger.log(
    'Daily run complete. fetched=' + fetchedCount +
    ', moved_to_work=' + moved +
    ', moved_to_followup=' + ((preMoveSummary.movedToFollowup || 0) + (syncSummary.movedToFollowup || 0)) +
    ', cleared_leadcallmsg_rows=' + syncSummary.clearedLeadCallMsgRows +
    ', synced_work_rows=' + syncSummary.syncedWorkRows +
    ', deleted_followup_closed_rows=' + syncSummary.deletedClosedFromFollowup +
    ', cleared_work_N_when_H_empty=' + clearedAssigned
  );
}

function dailyProcessEmployeeSheetsAndSyncToWork_(configSheet, workSheet) {
  const summary = {
    movedToFollowup: 0,
    clearedLeadCallMsgRows: 0,
    syncedWorkRows: 0,
    deletedClosedFromFollowup: 0
  };

  const configLastRow = configSheet.getLastRow();
  if (configLastRow < 2) {
    return summary;
  }

  const sheetIdValues = configSheet
    .getRange(2, DAILY_PIPELINE.COL.CONFIG_SHEET_ID, configLastRow - 1, 1)
    .getValues();

  const workLastRow = workSheet.getLastRow();
  const workLastCol = Math.max(workSheet.getLastColumn(), DAILY_PIPELINE.COL.WORK_PHONE);
  if (workLastRow < 2) {
    return summary;
  }

  const workData = workSheet.getRange(2, 1, workLastRow - 1, workLastCol).getValues();
  const phoneToWorkIndexes = new Map();

  for (let i = 0; i < workData.length; i++) {
    const phone = dailyNormalizePhone_(workData[i][DAILY_PIPELINE.COL.WORK_PHONE - 1]);
    if (!phone) {
      continue;
    }
    if (!phoneToWorkIndexes.has(phone)) {
      phoneToWorkIndexes.set(phone, []);
    }
    phoneToWorkIndexes.get(phone).push(i);
  }

  for (let i = 0; i < sheetIdValues.length; i++) {
    const sheetId = dailyToCleanString_(sheetIdValues[i][0]);
    if (!sheetId) {
      continue;
    }

    let empSS;
    try {
      empSS = SpreadsheetApp.openById(sheetId);
    } catch (e) {
      Logger.log('Failed to open employee spreadsheet: ' + sheetId + ', error=' + e);
      continue;
    }

    const leadCallMsg = empSS.getSheetByName(DAILY_PIPELINE.SHEETS.EMP_LEADCALLMSG);
    const followup = dailyGetOrCreateSheet_(empSS, DAILY_PIPELINE.SHEETS.EMP_FOLLOWUP);

    if (leadCallMsg && leadCallMsg.getLastRow() >= 2) {
      const leadLastRow = leadCallMsg.getLastRow();
      const leadLastCol = Math.max(leadCallMsg.getLastColumn(), DAILY_PIPELINE.COL.EMP_PHONE, DAILY_PIPELINE.COL.EMP_WORKED_FLAG);
      const leadRows = leadCallMsg.getRange(2, 1, leadLastRow - 1, leadLastCol).getValues();
      const workedRows = [];

      for (let r = 0; r < leadRows.length; r++) {
        const row = leadRows[r];
        const workedVal = dailyToCleanString_(row[DAILY_PIPELINE.COL.EMP_WORKED_FLAG - 1]);
        if (workedVal) {
          workedRows.push(row);
        }
      }

      if (workedRows.length > 0) {
        const followupStartRow = Math.max(followup.getLastRow() + 1, 2);
        const normalizedWorked = dailyNormalizeRows_(workedRows, leadLastCol);
        followup.getRange(followupStartRow, 1, normalizedWorked.length, normalizedWorked[0].length).setValues(normalizedWorked);
        summary.movedToFollowup += normalizedWorked.length;
      }

      // Requirement: after processing, clear all data rows from LeadCallMsg.
      leadCallMsg.deleteRows(2, leadLastRow - 1);
      summary.clearedLeadCallMsgRows += (leadLastRow - 1);
    }

    const followLastRow = followup.getLastRow();
    if (followLastRow < 2) {
      continue;
    }

    const followLastCol = Math.max(followup.getLastColumn(), DAILY_PIPELINE.COL.EMP_PHONE, 10); // ensure column J is available
    const followData = followup.getRange(2, 1, followLastRow - 1, followLastCol).getValues();
    const followRowsToDelete = [];

    for (let f = 0; f < followData.length; f++) {
      const followRow = followData[f];

      const statusJ = dailyToCleanString_(followRow[9]).toLowerCase(); // column J
      if (statusJ === 'closed') {
        followRowsToDelete.push(f + 2);
      }

      const phone = dailyNormalizePhone_(followRow[DAILY_PIPELINE.COL.EMP_PHONE - 1]);
      if (!phone) {
        continue;
      }

      const matches = phoneToWorkIndexes.get(phone);
      if (!matches || matches.length === 0) {
        continue;
      }

      for (let m = 0; m < matches.length; m++) {
        const workIdx = matches[m];
        const target = workData[workIdx];
        const copyLen = Math.min(target.length, followRow.length);
        for (let c = 0; c < copyLen; c++) {
          target[c] = followRow[c];
        }
        summary.syncedWorkRows += 1;
      }
    }

    if (followRowsToDelete.length > 0) {
      followRowsToDelete.sort((a, b) => b - a);
      for (let d = 0; d < followRowsToDelete.length; d++) {
        followup.deleteRow(followRowsToDelete[d]);
        summary.deletedClosedFromFollowup += 1;
      }
    }
  }

  workSheet.getRange(2, 1, workData.length, workLastCol).setValues(workData);
  return summary;
}

function dailyFetchIndiaMartLeads_(configSheet) {
  if (!DAILY_PIPELINE.INDIAMART_CRM_KEY) {
    throw new Error('CRM key missing');
  }

  const now = new Date();
  let lastFetch = configSheet.getRange("G2").getValue();

  let from;

  if (lastFetch && !isNaN(new Date(lastFetch).getTime())) {
    let parsed = new Date(lastFetch);

    // If only date (00:00:00), fallback to 1 hour
    if (
      parsed.getHours() === 0 &&
      parsed.getMinutes() === 0 &&
      parsed.getSeconds() === 0
    ) {
      from = new Date(now.getTime() - 60 * 60 * 1000);
    } else {
      from = parsed;
    }
  } else {
    from = new Date(now.getTime() - 60 * 60 * 1000);
  }

  const startTime = dailyFormatIndiaMartDate_(from);
  const endTime = dailyFormatIndiaMartDate_(now);

  const url =
    DAILY_PIPELINE.INDIAMART_URL +
    '?glusr_crm_key=' + encodeURIComponent(DAILY_PIPELINE.INDIAMART_CRM_KEY) +
    '&start_time=' + encodeURIComponent(startTime) +
    '&end_time=' + encodeURIComponent(endTime);

  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const status = response.getResponseCode();
    const body = response.getContentText();

    // ❌ HANDLE 429 RATE LIMIT
    if (status === 429) {
      Logger.log("Try again after 5 minutes");
      SpreadsheetApp.getUi().alert("Try again after 5 minutes");
    
      return []; // continue pipeline
    }

    if (status < 200 || status >= 300) {
      Logger.log('API Error: ' + status);
      return [];
    }

    const parsed = JSON.parse(body);

    const leads =
      (parsed && Number(parsed.CODE) === 200 && Array.isArray(parsed.RESPONSE))
        ? parsed.RESPONSE
        : [];

    // ✅ ONLY UPDATE G2 ON SUCCESS
    const cell = configSheet.getRange("G2");
    cell.setValue(now);
    cell.setNumberFormat("dd-mmm-yyyy hh:mm:ss"); // ✅ force proper format

    return leads;

  } catch (e) {
    Logger.log("API failed: " + e.toString());
    return null;
  }
}

function dailyAppendRawLeads_(sheet, leads) {
  if (!Array.isArray(leads) || leads.length === 0) {
    return;
  }

  const existingHeaders = dailyGetHeaderRow_(sheet);
  const apiHeaders = dailyCollectHeadersFromObjects_(leads);
  const headers = dailyMergeHeaderArrays_(existingHeaders, apiHeaders);
  if (headers.length === 0) {
    return;
  }

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const rows = leads.map(obj => headers.map(h => (obj[h] === undefined || obj[h] === null ? '' : obj[h])));
  dailyAppendRows_(sheet, rows);
}

function dailyCleanFreshLeadsToWork_(freshSheet, workSheet) {
  const freshData = dailyGetSheetDataWithoutHeader_(freshSheet);
  if (freshData.length === 0) {
    return 0;
  }

  const existingPhones = new Set();
  const workData = dailyGetSheetDataWithoutHeader_(workSheet);
  for (let i = 0; i < workData.length; i++) {
    const phone = dailyNormalizePhone_(workData[i][4]);
    if (phone) {
      existingPhones.add(phone);
    }
  }

  const cleanRows = [];
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy');

  for (let i = 0; i < freshData.length; i++) {
    const row = freshData[i];
    const customerName = dailyToCleanString_(row[3]);
    const phone = dailyNormalizePhone_(row[4]);
    const leadDateRaw = dailyToCleanString_(row[2]);
    const leadDate = dailyFormatToDDMMYYYY_(leadDateRaw);
    const keyword = dailyToCleanString_(row[18]); // FreshLeads column S

    if (!phone || existingPhones.has(phone)) {
      continue;
    }

    existingPhones.add(phone);
    cleanRows.push(['', '', keyword, customerName, phone, leadDate, 'Indiamart']);
  }

  if (cleanRows.length > 0) {
    dailyAppendRows_(workSheet, cleanRows, 7);
    dailyRenumberCountColumn_(workSheet, 2);

    if (DAILY_PIPELINE.CLEAN_FRESH_AFTER_MOVE) {
      const lastRow = freshSheet.getLastRow();
      const lastCol = freshSheet.getLastColumn();
      if (lastRow >= 2 && lastCol >= 1) {
        freshSheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
      }
    }
  }

  return cleanRows.length;
}

function dailyClearConfigColumnsCD_(configSheet) {
  const last = configSheet.getLastRow();
  if (last < 2) {
    return;
  }
  const numRows = last - 1;
  configSheet.getRange(2, 3, numRows, 1).clearContent(); // C
  configSheet.getRange(2, 4, numRows, 1).clearContent(); // D
}

function dailyClearWorkAssignedWhereHEmpty_(workSheet) {
  const lastRow = workSheet.getLastRow();
  if (lastRow < 2) {
    return 0;
  }

  const rowCount = lastRow - 1;
  const hVals = workSheet.getRange(2, 8, rowCount, 1).getDisplayValues(); // H
  const nVals = workSheet.getRange(2, 14, rowCount, 1).getValues(); // N

  let cleared = 0;
  for (let i = 0; i < rowCount; i++) {
    const h = dailyToCleanString_(hVals[i] && hVals[i][0]);
    const n = dailyToCleanString_(nVals[i] && nVals[i][0]);
    if (h === '' && n !== '') {
      nVals[i][0] = '';
      cleared += 1;
    }
  }

  if (cleared > 0) {
    workSheet.getRange(2, 14, rowCount, 1).setValues(nVals);
  }

  return cleared;
}

function dailyGetOrCreateSheet_(ss, sheetName) {
  let sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
  }
  return sh;
}

function dailyGetHeaderRow_(sheet) {
  if (sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) {
    return [];
  }
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function dailyGetSheetDataWithoutHeader_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    return [];
  }
  return sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
}

function dailyAppendRows_(sheet, rows, minColumns) {
  if (!Array.isArray(rows) || rows.length === 0) {
    return;
  }
  const normalized = dailyNormalizeRows_(rows, minColumns || 1);
  const startRow = Math.max(sheet.getLastRow() + 1, 2);
  sheet.getRange(startRow, 1, normalized.length, normalized[0].length).setValues(normalized);
}

function dailyNormalizeRows_(rows, minCols) {
  let width = Math.max(1, minCols || 1);
  for (let i = 0; i < rows.length; i++) {
    width = Math.max(width, rows[i].length);
  }

  const out = [];
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i].slice();
    while (r.length < width) {
      r.push('');
    }
    if (r.length > width) {
      r.length = width;
    }
    out.push(r);
  }
  return out;
}

function dailyMergeHeaderArrays_(a, b) {
  const seen = new Set();
  const merged = [];

  function push(arr) {
    for (let i = 0; i < arr.length; i++) {
      const h = dailyToCleanString_(arr[i]);
      if (!h || seen.has(h)) {
        continue;
      }
      seen.add(h);
      merged.push(h);
    }
  }

  push(a || []);
  push(b || []);
  return merged;
}

function dailyCollectHeadersFromObjects_(objects) {
  const seen = new Set();
  const out = [];

  for (let i = 0; i < objects.length; i++) {
    const keys = Object.keys(objects[i] || {});
    for (let k = 0; k < keys.length; k++) {
      const key = dailyToCleanString_(keys[k]);
      if (!key || seen.has(key)) {
        continue;
      }
      seen.add(key);
      out.push(key);
    }
  }
  return out;
}

function dailyDeleteTriggersByHandler_(handlerName) {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === handlerName) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function dailyEnsureWorkHeader_(workSheet) {
  if (workSheet.getLastRow() >= 1) {
    return;
  }
  workSheet.getRange(1, 1, 1, 7).setValues([[
    'Date',
    'Count',
    'Keyword',
    'Customer Name',
    'Phone',
    'Lead Date',
    'Source'
  ]]);
}

function dailyRenumberCountColumn_(sheet, dataStartRow) {
  const lastRow = sheet.getLastRow();
  if (lastRow < dataStartRow) {
    return;
  }
  const n = lastRow - dataStartRow + 1;
  const values = [];
  for (let i = 0; i < n; i++) {
    values.push([i + 1]);
  }
  sheet.getRange(dataStartRow, 2, n, 1).setValues(values);
}

function dailyFormatIndiaMartDate_(dt) {
  const dd = ('0' + dt.getDate()).slice(-2);
  const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const mon = monthNames[dt.getMonth()];
  const yyyy = dt.getFullYear();
  const hh = ('0' + dt.getHours()).slice(-2);
  const mm = ('0' + dt.getMinutes()).slice(-2);
  const ss = ('0' + dt.getSeconds()).slice(-2);
  return dd + '-' + mon + '-' + yyyy + hh + ':' + mm + ':' + ss;
}

function dailyNormalizePhone_(val) {
  if (val === undefined || val === null) {
    return '';
  }
  let digits = String(val).replace(/\D/g, '');
  if (!digits) {
    return '';
  }
  if (digits.length > 10) {
    digits = digits.slice(-10);
  }
  if (digits.length < 10) {
    return '';
  }
  return digits;
}

function dailyToCleanString_(val) {
  if (val === undefined || val === null) {
    return '';
  }
  return String(val).trim();
}

function dailyFormatToDDMMYYYY_(val) {
  if (val === undefined || val === null) {
    return '';
  }
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd-MM-yyyy');
  }
  const s = String(val).trim();
  if (!s) return '';

  // Try to parse with Date
  const parsed = new Date(s);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'dd-MM-yyyy');
  }

  // Try common numeric date patterns: yyyy-mm-dd or dd-mm-yyyy or dd/mm/yyyy
  const m = s.match(/^(\d{1,4})[\/-](\d{1,2})[\/-](\d{1,4})/);
  if (m) {
    const a = m[1], b = m[2], c = m[3];
    if (a.length === 4) {
      // a = year
      const yyyy = a.padStart(4, '0');
      const mm = b.padStart(2, '0');
      const dd = c.padStart(2, '0');
      return dd + '-' + mm + '-' + yyyy;
    }
    if (c.length === 4) {
      // c = year, assume a = dd
      const dd = a.padStart(2, '0');
      const mm = b.padStart(2, '0');
      const yyyy = c.padStart(4, '0');
      return dd + '-' + mm + '-' + yyyy;
    }
  }

  // Fallback: return original string
  return s;
}

function dailyMoveLeadCallMsgWorkedToFollowup_(configSheet) {
  const summary = {
    movedToFollowup: 0,
    deletedFromLeadCallMsg: 0
  };

  const configLastRow = configSheet.getLastRow();
  if (configLastRow < 2) {
    return summary;
  }

  const sheetIdValues = configSheet
    .getRange(2, DAILY_PIPELINE.COL.CONFIG_SHEET_ID, configLastRow - 1, 1)
    .getValues();

  for (let i = 0; i < sheetIdValues.length; i++) {
    const sheetId = dailyToCleanString_(sheetIdValues[i][0]);
    if (!sheetId) {
      continue;
    }

    let empSS;
    try {
      empSS = SpreadsheetApp.openById(sheetId);
    } catch (e) {
      Logger.log('Failed to open employee spreadsheet: ' + sheetId + ', error=' + e);
      continue;
    }

    const leadCallMsg = empSS.getSheetByName(DAILY_PIPELINE.SHEETS.EMP_LEADCALLMSG);
    const followup = dailyGetOrCreateSheet_(empSS, DAILY_PIPELINE.SHEETS.EMP_FOLLOWUP);

    if (!leadCallMsg || leadCallMsg.getLastRow() < 2) {
      continue;
    }

    const leadLastRow = leadCallMsg.getLastRow();
    const leadLastCol = Math.max(leadCallMsg.getLastColumn(), DAILY_PIPELINE.COL.EMP_PHONE, DAILY_PIPELINE.COL.EMP_WORKED_FLAG);
    const leadValues = leadCallMsg.getRange(2, 1, leadLastRow - 1, leadLastCol).getValues();
    const leadValidations = leadCallMsg.getRange(2, 1, leadLastRow - 1, leadLastCol).getDataValidations();

    const rowsToMove = [];
    const valsToMove = [];
    const indicesToDelete = [];

    for (let r = 0; r < leadValues.length; r++) {
      const row = leadValues[r];
      const workedVal = dailyToCleanString_(row[DAILY_PIPELINE.COL.EMP_WORKED_FLAG - 1]);
      if (workedVal) {
        rowsToMove.push(row);
        valsToMove.push(leadValidations[r]);
        indicesToDelete.push(r + 2);
      }
    }

    if (rowsToMove.length === 0) {
      continue;
    }

    const startRow = Math.max(followup.getLastRow() + 1, 2);
    const normalized = dailyNormalizeRows_(rowsToMove, leadLastCol);
    followup.getRange(startRow, 1, normalized.length, normalized[0].length).setValues(normalized);
    summary.movedToFollowup += normalized.length;

    // Apply data validations row by row (preserve dropdowns)
    for (let j = 0; j < valsToMove.length; j++) {
      const rowValidations = valsToMove[j] || [];
      // ensure width matches
      while (rowValidations.length < normalized[j].length) {
        rowValidations.push(null);
      }
      followup.getRange(startRow + j, 1, 1, rowValidations.length).setDataValidations([rowValidations]);
    }

    // delete original rows from LeadCallMsg (delete from bottom to top)
    indicesToDelete.sort((a, b) => b - a);
    for (let k = 0; k < indicesToDelete.length; k++) {
      leadCallMsg.deleteRow(indicesToDelete[k]);
      summary.deletedFromLeadCallMsg += 1;
    }
  }

  return summary;
}
