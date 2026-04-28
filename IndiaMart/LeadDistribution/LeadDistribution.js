// V2 pipeline (user-facing functions only):
// 1) completeRun() -> run full flow once manually
// 2) toggleLeadPipelineV2Trigger() -> start/stop daily 8AM trigger

const LeadPipelineV2 = (function () {
  const CFG = Object.freeze({
    CRM_KEY: 'mRy0FrFr4XrITves7nyK7liNqlHEnzBl',
    BASE_URL: 'https://mapi.indiamart.com/wservce/crm/crmListing/v2/',
    FETCH_HOURS: 24,
    DAILY_RUN_HOUR: 8,
    DEFAULT_LEADS_PER_EMPLOYEE: 50,
    SHEETS: {
      CONFIG: 'Config',
      RAWDATA: 'Rawdata',
      BUYLEADS: 'FreshLeads',
      WORK: 'work',
      EXTERNAL: 'LeadCallMsg',
      FOLLOWUP: 'Followup'
    },
    CONFIG_COLS: {
      EMP_NAME: 1,
      EMP_ID: 2,
      SOURCE_SHEET: 4,
      SOURCE_LAST_COPIED_ROW: 5,
      SOURCE_PER_EMPLOYEE: 6,
      LEADS_PER_EMPLOYEE: 8
    },
    COLS: {
      PHONE: 5,
      STATUS_H: 8,
      LINK_O: 14
    },
    WHATSAPP_LINK_TEXT: 'Whatsapp',
    EMPLOYEE_DATA_START_ROW: 3,
    PROPS: {
      DAILY_TRIGGER_ENABLED: 'LEAD_PIPELINE_V2_DAILY_TRIGGER_ENABLED'
    }
  });

  function _toNonNegativeInt(val, fallback) {
    const n = parseInt(String(val === undefined || val === null ? '' : val).trim(), 10);
    if (isNaN(n) || n < 0) return fallback;
    return n;
  }

  function _formatIndiaMartDate(dt) {
    const dd = ('0' + dt.getDate()).slice(-2);
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const mon = months[dt.getMonth()];
    const yyyy = dt.getFullYear();
    const hh = ('0' + dt.getHours()).slice(-2);
    const mm = ('0' + dt.getMinutes()).slice(-2);
    const ss = ('0' + dt.getSeconds()).slice(-2);
    return dd + '-' + mon + '-' + yyyy + hh + ':' + mm + ':' + ss;
  }

  function _normalizePhone(val) {
    if (val === undefined || val === null) return '';
    let s = String(val).replace(/\D/g, '');
    if (!s) return '';
    if (s.length > 10) s = s.slice(-10);
    return s;
  }

  function _normalizeSourceName(val) {
    if (val === undefined || val === null) return '';
    return String(val).trim().toLowerCase();
  }

  function _getRowSourceLabel(row) {
    if (!row || row.length < 7) return 'unknown';
    const src = _normalizeSourceName(row[6]); // column G
    return src || 'unknown';
  }

  function _getOrCreateSheet(ss, name) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    return sh;
  }

  function _getHeaders(sheet) {
    if (!sheet || sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) return [];
    return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  }

  function _mergeHeaders(headersA, headersB, rows) {
    const merged = [];
    const seen = new Set();

    (headersA || []).forEach(h => {
      if (!seen.has(h)) {
        seen.add(h);
        merged.push(h);
      }
    });

    (headersB || []).forEach(h => {
      if (!seen.has(h)) {
        seen.add(h);
        merged.push(h);
      }
    });

    (rows || []).forEach(obj => {
      Object.keys(obj || {}).forEach(k => {
        if (!seen.has(k)) {
          seen.add(k);
          merged.push(k);
        }
      });
    });

    return merged;
  }

  function _rowsToMatrix(rows, headers) {
    return (rows || []).map(obj => headers.map(h => (obj[h] === undefined ? '' : obj[h])));
  }

  function _fetchIndiaMartRows() {
    const now = new Date();
    const start = new Date(now.getTime() - CFG.FETCH_HOURS * 60 * 60 * 1000);
    const params = {
      glusr_crm_key: CFG.CRM_KEY,
      start_time: _formatIndiaMartDate(start),
      end_time: _formatIndiaMartDate(now)
    };

    const qs = Object.keys(params)
      .map(k => encodeURIComponent(k) + '=' + encodeURIComponent(params[k]))
      .join('&');

    const url = CFG.BASE_URL + '?' + qs;
    try {
      const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      const body = resp.getContentText();
      const json = JSON.parse(body);
      if (json && Number(json.CODE) === 200) return json.RESPONSE || [];
      Logger.log('IndiaMART API error: %s', body);
      return [];
    } catch (e) {
      Logger.log('IndiaMART API request failed: %s', e.toString());
      return [];
    }
  }

  function _appendFetchedToRawAndBuy(rows) {
    if (!rows || !rows.length) return;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const raw = _getOrCreateSheet(ss, CFG.SHEETS.RAWDATA);
    const buy = _getOrCreateSheet(ss, CFG.SHEETS.BUYLEADS);

    const headers = _mergeHeaders(_getHeaders(raw), _getHeaders(buy), rows);
    if (!headers.length) return;

    raw.getRange(1, 1, 1, headers.length).setValues([headers]);
    buy.getRange(1, 1, 1, headers.length).setValues([headers]);

    const matrix = _rowsToMatrix(rows, headers);
    const rawStart = Math.max(raw.getLastRow() + 1, 2);
    const buyStart = Math.max(buy.getLastRow() + 1, 2);
    raw.getRange(rawStart, 1, matrix.length, headers.length).setValues(matrix);
    buy.getRange(buyStart, 1, matrix.length, headers.length).setValues(matrix);
  }

  function _syncEmployeesToWorkInternal() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const employees = _getEmployees();
    if (!employees.length) {
      Logger.log('No employees available for sync.');
      return;
    }

    let work = ss.getSheetByName(CFG.SHEETS.WORK);
    if (!work) work = ss.insertSheet(CFG.SHEETS.WORK);

    const workLast = work.getLastRow();
    if (workLast < 3) {
      Logger.log('No work rows available for employee sync.');
      return;
    }

    const workCols = Math.max(13, work.getLastColumn());
    const workRange = work.getRange(3, 1, workLast - 2, workCols);
    const workData = workRange.getValues();

    const workMap = new Map();
    for (let i = 0; i < workData.length; i++) {
      const p = _normalizePhone(workData[i][4]);
      if (!p) continue;
      if (!workMap.has(p)) workMap.set(p, []);
      workMap.get(p).push(i);
    }

    let updates = 0;
    for (let r = 0; r < employees.length; r++) {
      const id = employees[r].id ? String(employees[r].id).trim() : '';
      if (!id) continue;

      let ext;
      try {
        ext = SpreadsheetApp.openById(id);
      } catch (e) {
        Logger.log('Failed opening employee spreadsheet %s: %s', id, e.toString());
        continue;
      }

      // Requirement: sync from Followup sheet to master work.
      const src = ext.getSheetByName(CFG.SHEETS.FOLLOWUP);
      if (!src || src.getLastRow() < CFG.EMPLOYEE_DATA_START_ROW) continue;

      const srcCols = Math.max(13, src.getLastColumn());
      const srcData = src.getRange(
        CFG.EMPLOYEE_DATA_START_ROW,
        1,
        src.getLastRow() - (CFG.EMPLOYEE_DATA_START_ROW - 1),
        srcCols
      ).getValues();

      for (let i = 0; i < srcData.length; i++) {
        const sRow = srcData[i];
        const p = _normalizePhone(sRow[4]);
        if (!p) continue;

        const matchIdx = workMap.get(p);
        if (!matchIdx || !matchIdx.length) continue;

        for (let m = 0; m < matchIdx.length; m++) {
          const idx = matchIdx[m];
          const t = workData[idx];
          t[0] = sRow[0] !== undefined ? sRow[0] : '';
          t[6] = sRow[6] !== undefined ? sRow[6] : '';
          t[7] = sRow[7] !== undefined ? sRow[7] : '';
          t[8] = sRow[8] !== undefined ? sRow[8] : '';
          t[9] = sRow[9] !== undefined ? sRow[9] : '';
          t[10] = sRow[10] !== undefined ? sRow[10] : '';
          t[11] = sRow[11] !== undefined ? sRow[11] : '';
          t[12] = sRow[12] !== undefined ? sRow[12] : '';
          updates += 1;
        }
      }
    }

    if (updates > 0) workRange.setValues(workData);
    Logger.log('Internal employee sync updated rows=%s', updates);
  }

  function _moveWorkedRowsToFollowupAndClearInbox(extSpreadsheet, inboxSheet, employeeId) {
    const inboxLast = inboxSheet.getLastRow();
    if (inboxLast < CFG.EMPLOYEE_DATA_START_ROW) return;

    const inboxCols = Math.max(inboxSheet.getLastColumn(), CFG.COLS.LINK_O);
    const inboxRange = inboxSheet.getRange(
      CFG.EMPLOYEE_DATA_START_ROW,
      1,
      inboxLast - (CFG.EMPLOYEE_DATA_START_ROW - 1),
      inboxCols
    );
    const inboxData = inboxRange.getValues();
    const workedSourceRows = [];
    for (let i = 0; i < inboxData.length; i++) {
      const bVal = inboxData[i][1];
      const hVal = inboxData[i][CFG.COLS.STATUS_H - 1];
      const bFilled = bVal !== undefined && bVal !== null && String(bVal).trim() !== '';
      const hFilled = hVal !== undefined && hVal !== null && String(hVal).trim() !== '';
      if (bFilled && hFilled) {
        workedSourceRows.push(CFG.EMPLOYEE_DATA_START_ROW + i);
      }
    }

    if (workedSourceRows.length) {
      const followup = _getOrCreateSheet(extSpreadsheet, CFG.SHEETS.FOLLOWUP);
      if (followup.getLastRow() < 2) {
        const headerRows = Math.min(2, inboxSheet.getLastRow());
        if (headerRows > 0) {
          inboxSheet.getRange(1, 1, headerRows, inboxCols).copyTo(followup.getRange(1, 1, headerRows, inboxCols));
        }
      }
      const start = Math.max(followup.getLastRow() + 1, CFG.EMPLOYEE_DATA_START_ROW);
      for (let i = 0; i < workedSourceRows.length; i++) {
        const srcRow = workedSourceRows[i];
        const dstRow = start + i;
        inboxSheet.getRange(srcRow, 1, 1, inboxCols).copyTo(followup.getRange(dstRow, 1, 1, inboxCols));
      }
    }

    inboxSheet.deleteRows(CFG.EMPLOYEE_DATA_START_ROW, inboxLast - (CFG.EMPLOYEE_DATA_START_ROW - 1));
    Logger.log('Employee %s: moved %s worked rows to Followup and cleared LeadCallMsg.', employeeId, String(workedSourceRows.length));
  }

  function _toWhatsAppPhone(phoneRaw) {
    const p = _normalizePhone(phoneRaw);
    if (!p) return '';
    if (p.length === 10) return '91' + p;
    return p;
  }

  function _buildPrefilledWhatsAppMessage(name, keyword) {
    const who = (name !== undefined && name !== null && String(name).trim() !== '') ? String(name).trim() : 'there';
    const topic = (keyword !== undefined && keyword !== null && String(keyword).trim() !== '') ? String(keyword).trim() : 'your requirement';
    return 'hello ' + who + ', we got your requirement from indiamart regarding ' + topic + ',Please find out below refrence images.';
  }

  function _setWhatsAppLinksInColumnO(sheet, startRow, rowsData) {
    if (!rowsData || !rowsData.length) return;

    const out = [];
    for (let i = 0; i < rowsData.length; i++) {
      const row = rowsData[i] || [];
      const phone = _toWhatsAppPhone(row[4]);
      if (!phone) {
        out.push(['']);
        continue;
      }

      // Build a simple universal link with just the phone number
      const url = 'https://web.whatsapp.com/send?phone=' + phone.replace(/\D/g, ''); 

      // Push the hyperlink to your output array
      out.push(['=HYPERLINK("' + url + '","' + CFG.WHATSAPP_LINK_TEXT + '")']);

    }
    const range = sheet.getRange(startRow, CFG.COLS.LINK_O, rowsData.length, 1);
    range.clearContent();
    range.setValues(out);
  }

  function _renumberWorkCountColumn() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const work = ss.getSheetByName(CFG.SHEETS.WORK);
    if (!work) return;

    const last = work.getLastRow();
    if (last < 3) return;

    const n = last - 2;
    const seq = [];
    for (let i = 0; i < n; i++) seq.push([i + 1]);
    work.getRange(3, 2, n, 1).setValues(seq);
  }

  function _renumberSheetCountColumn(sheet) {
    if (!sheet) return;
    const last = sheet.getLastRow();
    if (last < CFG.EMPLOYEE_DATA_START_ROW) return;

    const n = last - (CFG.EMPLOYEE_DATA_START_ROW - 1);
    const seq = [];
    for (let i = 0; i < n; i++) seq.push([i + 1]);
    sheet.getRange(CFG.EMPLOYEE_DATA_START_ROW, 2, n, 1).setValues(seq);
  }

  function _getEmployees() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = ss.getSheetByName(CFG.SHEETS.CONFIG);
    const employees = [];

    if (config && config.getLastRow() >= 2) {
      const rows = config.getRange(2, 1, config.getLastRow() - 1, 4).getValues();
      for (let i = 0; i < rows.length; i++) {
        const name = rows[i][CFG.CONFIG_COLS.EMP_NAME - 1] ? String(rows[i][CFG.CONFIG_COLS.EMP_NAME - 1]).trim() : '';
        const id = rows[i][CFG.CONFIG_COLS.EMP_ID - 1] ? String(rows[i][CFG.CONFIG_COLS.EMP_ID - 1]).trim() : '';
        if (!id) continue;
        employees.push({ name: name || id, id: id });
      }
    }

    return employees;
  }

  function _getSourceSheetNamesFromConfig() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = ss.getSheetByName(CFG.SHEETS.CONFIG);
    if (!config || config.getLastRow() < 2) return [];

    const vals = config.getRange(2, CFG.CONFIG_COLS.SOURCE_SHEET, config.getLastRow() - 1, 1).getValues();
    const names = [];
    const seen = new Set();
    for (let i = 0; i < vals.length; i++) {
      const n = vals[i][0] ? String(vals[i][0]).trim() : '';
      if (!n) continue;
      if (seen.has(n.toLowerCase())) continue;
      seen.add(n.toLowerCase());
      names.push(n);
    }
    return names;
  }

  function _getSourceSheetConfigsFromConfig() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = ss.getSheetByName(CFG.SHEETS.CONFIG);
    if (!config || config.getLastRow() < 2) return [];

    const width = Math.max(CFG.CONFIG_COLS.SOURCE_LAST_COPIED_ROW, CFG.CONFIG_COLS.SOURCE_PER_EMPLOYEE);
    const rows = config.getRange(2, 1, config.getLastRow() - 1, width).getValues();
    const list = [];
    const seen = new Set();

    for (let i = 0; i < rows.length; i++) {
      const configRow = i + 2;
      const sheetName = rows[i][CFG.CONFIG_COLS.SOURCE_SHEET - 1] ? String(rows[i][CFG.CONFIG_COLS.SOURCE_SHEET - 1]).trim() : '';
      if (!sheetName) continue;

      const quotaRaw = rows[i][CFG.CONFIG_COLS.SOURCE_PER_EMPLOYEE - 1];
      const quotaText = quotaRaw === undefined || quotaRaw === null ? '' : String(quotaRaw).trim();
      if (quotaText === '') continue; // blank in F means skip source sheet

      const perEmployeeCount = _toNonNegativeInt(quotaText, -1);
      if (perEmployeeCount < 0) continue;

      const key = sheetName.toLowerCase();
      if (seen.has(key)) continue;
      seen.add(key);

      const rawLast = rows[i][CFG.CONFIG_COLS.SOURCE_LAST_COPIED_ROW - 1];
      const parsed = parseInt(String(rawLast || '0'), 10);
      const lastCopiedRow = isNaN(parsed) ? 0 : parsed;

      list.push({
        configRow: configRow,
        sheetName: sheetName,
        lastCopiedRow: lastCopiedRow,
        perEmployeeCount: perEmployeeCount,
        isFlexible: perEmployeeCount === 0
      });
    }

    return list;
  }

  function _getActiveSourceConfigsForDistribution() {
    return _getSourceSheetConfigsFromConfig().filter(function (cfg) {
      const name = String(cfg.sheetName).toLowerCase();
      return name !== String(CFG.SHEETS.WORK).toLowerCase()
        && name !== String(CFG.SHEETS.RAWDATA).toLowerCase()
        && name !== String(CFG.SHEETS.BUYLEADS).toLowerCase();
    });
  }

  function _getEffectiveLeadsPerEmployeeFromSourceRules(configuredLeadsPerEmployee) {
    const sourceConfigs = _getActiveSourceConfigsForDistribution();
    let fixedTotal = 0;
    let hasFlexible = false;

    for (let i = 0; i < sourceConfigs.length; i++) {
      const c = sourceConfigs[i];
      if (c.perEmployeeCount === 0) {
        hasFlexible = true;
      } else if (c.perEmployeeCount > 0) {
        fixedTotal += c.perEmployeeCount;
      }
    }

    if (fixedTotal > 0 && !hasFlexible) {
      return {
        leadsPerEmployee: fixedTotal,
        overrideByFixedRules: true,
        fixedTotal: fixedTotal
      };
    }

    return {
      leadsPerEmployee: configuredLeadsPerEmployee,
      overrideByFixedRules: false,
      fixedTotal: fixedTotal
    };
  }

  function _getLeadsPerEmployeeFromConfig(defaultValue) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = ss.getSheetByName(CFG.SHEETS.CONFIG);
    if (!config) return defaultValue;
    const val = config.getRange(2, CFG.CONFIG_COLS.LEADS_PER_EMPLOYEE).getValue(); // H2
    const parsed = _toNonNegativeInt(val, -1);
    if (parsed <= 0) return defaultValue;
    return parsed;
  }

  function _normalizeRowsToColumnCount(rows, minCols) {
    if (!rows || !rows.length) return { rows: [], cols: Math.max(0, minCols || 0) };

    let cols = Math.max(0, minCols || 0);
    for (let i = 0; i < rows.length; i++) {
      const len = rows[i] ? rows[i].length : 0;
      if (len > cols) cols = len;
    }

    const out = [];
    for (let i = 0; i < rows.length; i++) {
      const src = rows[i] ? rows[i].slice() : [];
      while (src.length < cols) src.push('');
      if (src.length > cols) src.length = cols;
      out.push(src);
    }

    return { rows: out, cols: cols };
  }

  function _prepareEmployeeSheetsForNextCycle() {
    const employees = _getEmployees();
    if (!employees.length) {
      Logger.log('No employee IDs available for prepare step.');
      return;
    }

    for (let i = 0; i < employees.length; i++) {
      const emp = employees[i];
      let ext;
      try {
        ext = SpreadsheetApp.openById(emp.id);
      } catch (e) {
        Logger.log('Failed opening employee sheet %s in prepare step: %s', emp.id, e.toString());
        continue;
      }

      const inbox = _getOrCreateSheet(ext, CFG.SHEETS.EXTERNAL);
      try {
        _moveWorkedRowsToFollowupAndClearInbox(ext, inbox, emp.id);
        const followup = ext.getSheetByName(CFG.SHEETS.FOLLOWUP);
        _renumberSheetCountColumn(followup);
      } catch (e) {
        Logger.log('Prepare step failed for %s: %s', emp.id, e.toString());
      }
    }
  }

  function _getHeaderIndexMap(headers) {
    const map = {};
    for (let i = 0; i < headers.length; i++) {
      map[String(headers[i])] = i;
    }
    return map;
  }

  function _extractByHeader(row, idxMap, names) {
    for (let i = 0; i < names.length; i++) {
      const idx = idxMap[names[i]];
      if (idx !== undefined && row[idx] !== undefined && row[idx] !== null && String(row[idx]).trim() !== '') {
        return String(row[idx]);
      }
    }
    return '';
  }

  function _moveBuyLeadsToWorkAndClearBuyLeads() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const buy = ss.getSheetByName(CFG.SHEETS.BUYLEADS);
    if (!buy || buy.getLastRow() < 2) {
      Logger.log('BuyLeads has no data.');
      return;
    }

    const headers = _getHeaders(buy);
    const idxMap = _getHeaderIndexMap(headers);
    const lastRow = buy.getLastRow();
    const lastCol = buy.getLastColumn();
    const buyData = buy.getRange(2, 1, lastRow - 1, lastCol).getValues();

    let work = ss.getSheetByName(CFG.SHEETS.WORK);
    if (!work) work = ss.insertSheet(CFG.SHEETS.WORK);
    const workHeaders = ['', 'Count', 'Keyword', 'Customer Name', 'Phone', 'Lead Date', 'Source'];
    work.getRange(1, 1, 1, workHeaders.length).setValues([workHeaders]);
    const outRows = [];
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

    for (let i = 0; i < buyData.length; i++) {
      const row = buyData[i];

      const name = _extractByHeader(row, idxMap, ['SENDER_NAME', 'NAME']) || (row[3] ? String(row[3]) : '');
      const phone = _extractByHeader(row, idxMap, ['SENDER_MOBILE', 'MOBILE', 'MOBILE_NO']) || (row[4] ? String(row[4]) : '');

      if (!phone || String(phone).trim() === '') continue;

      // Requirement: column C in work should come directly from FreshLeads column S.
      const keyword = (row[18] !== undefined && row[18] !== null) ? String(row[18]) : '';

      outRows.push(['', '', keyword, name, phone, today, 'Indiamart']);
    }

    if (outRows.length) {
      const workLast = work.getLastRow();
      const nextRow = workLast >= 3 ? workLast + 1 : 3;
      work.getRange(nextRow, 1, outRows.length, outRows[0].length).setValues(outRows);

      _renumberWorkCountColumn();
    }

    // Clear BuyLeads data completely except header row.
    const currentLast = buy.getLastRow();
    if (currentLast > 1) {
      buy.deleteRows(2, currentLast - 1);
    }
  }

  function _removeBlockedIndiaMartRowsFromWork() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const work = ss.getSheetByName(CFG.SHEETS.WORK);
    if (!work || work.getLastRow() < 3) return 0;

    const workCols = Math.max(7, work.getLastColumn());
    const data = work.getRange(3, 1, work.getLastRow() - 2, workCols).getValues();
    const rowsToDelete = [];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const source = _normalizeSourceName(row[6]);
      const enquiry = row[2] === undefined || row[2] === null ? '' : String(row[2]).toLowerCase();
      if (source === 'indiamart' && enquiry.indexOf('kapila pashu aahar') !== -1) {
        rowsToDelete.push(3 + i);
      }
    }

    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      work.deleteRow(rowsToDelete[i]);
    }

    if (rowsToDelete.length) {
      _renumberWorkCountColumn();
    }

    Logger.log('Removed blocked IndiaMART work rows: %s', rowsToDelete.length);
    return rowsToDelete.length;
  }

  function _getAssignableWorkIndexes(workData) {
    const indexes = [];
    let skippedInvalid = 0;

    for (let i = 0; i < workData.length; i++) {
      const bVal = workData[i][1];
      const phone = _normalizePhone(workData[i][4]);
      const hVal = workData[i][7];
      const hEmpty = hVal === undefined || hVal === null || String(hVal).trim() === '';
      const bFilled = bVal !== undefined && bVal !== null && String(bVal).trim() !== '';
      const hasPhone = phone !== '';

      if (hEmpty && bFilled && hasPhone) {
        indexes.push(i);
      } else if (hEmpty && !hasPhone) {
        skippedInvalid += 1;
      }
    }

    return { indexes: indexes, skippedInvalid: skippedInvalid };
  }

  function _topUpWorkFromConfigSheets(requiredTotal, assigneeCount) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let work = ss.getSheetByName(CFG.SHEETS.WORK);
    if (!work) work = ss.insertSheet(CFG.SHEETS.WORK);

    const workCols = Math.max(14, work.getLastColumn());
    const workLast = work.getLastRow();
    const workData = workLast >= 3 ? work.getRange(3, 1, workLast - 2, workCols).getValues() : [];
    const assignableNow = _getAssignableWorkIndexes(workData).indexes.length;
    const needed = Math.max(0, requiredTotal - assignableNow);
    Logger.log('Top-up check: total unworked in work=%s, required rows=%s, additional needed=%s', assignableNow, requiredTotal, needed);
    if (needed <= 0) return 0;

    const sourceConfigs = _getActiveSourceConfigsForDistribution();

    if (!sourceConfigs.length) {
      Logger.log('Top-up required (%s) but no source sheets found in Config column D.', needed);
      return 0;
    }

    const existingPhones = new Set();
    for (let i = 0; i < workData.length; i++) {
      const p = _normalizePhone(workData[i][4]);
      if (p) existingPhones.add(p);
    }

    const pools = [];
    for (let s = 0; s < sourceConfigs.length; s++) {
      const cfgRow = sourceConfigs[s];
      const shName = cfgRow.sheetName;
      const sourceStartRow = Math.max(CFG.EMPLOYEE_DATA_START_ROW, (cfgRow.lastCopiedRow || 0) + 1);
      const sh = ss.getSheetByName(shName);
      if (!sh || sh.getLastRow() < 3) {
        pools.push({
          name: shName,
          key: _normalizeSourceName(shName),
          configRow: cfgRow.configRow,
          perEmployeeCount: cfgRow.perEmployeeCount,
          isFlexible: cfgRow.isFlexible,
          rows: [],
          ptr: 0,
          picked: 0,
          skippedDuplicatePhone: 0,
          sourceStartRow: sourceStartRow,
          lastSelectedSourceRow: 0,
          currentLastCopiedRow: cfgRow.lastCopiedRow
        });
        continue;
      }

      const startRow = sourceStartRow;
      if (startRow > sh.getLastRow()) {
        pools.push({
          name: shName,
          key: _normalizeSourceName(shName),
          configRow: cfgRow.configRow,
          perEmployeeCount: cfgRow.perEmployeeCount,
          isFlexible: cfgRow.isFlexible,
          rows: [],
          ptr: 0,
          picked: 0,
          skippedDuplicatePhone: 0,
          sourceStartRow: sourceStartRow,
          lastSelectedSourceRow: 0,
          currentLastCopiedRow: cfgRow.lastCopiedRow
        });
        continue;
      }

      const srcCols = Math.max(workCols, sh.getLastColumn());
      const srcData = sh.getRange(startRow, 1, sh.getLastRow() - (startRow - 1), srcCols).getValues();
      const candidates = [];
      let skippedDuplicatePhone = 0;

      for (let i = 0; i < srcData.length; i++) {
        const r = srcData[i];
        const phone = _normalizePhone(r[4]);
        if (phone && existingPhones.has(phone)) {
          skippedDuplicatePhone += 1;
          continue;
        }
        if (phone) existingPhones.add(phone);

        // Copy rows exactly as they are from source sheets.
        const row = r.slice(0, workCols);
        while (row.length < workCols) row.push('');
        // Clear assignment columns before re-queueing.
        row[13] = '';
        row[14] = '';
        candidates.push({
          data: row,
          sourceRow: startRow + i
        });
      }

      pools.push({
        name: shName,
        key: _normalizeSourceName(shName),
        configRow: cfgRow.configRow,
        perEmployeeCount: cfgRow.perEmployeeCount,
        isFlexible: cfgRow.isFlexible,
        rows: candidates,
        ptr: 0,
        picked: 0,
        skippedDuplicatePhone: skippedDuplicatePhone,
        sourceStartRow: sourceStartRow,
        lastSelectedSourceRow: 0,
        currentLastCopiedRow: cfgRow.lastCopiedRow
      });
    }

    let left = needed;
    const selectedBySource = new Map();
    for (let i = 0; i < pools.length; i++) {
      selectedBySource.set(pools[i].key, []);
    }

    function takeFromPool(pool, count) {
      let taken = 0;
      while (count > 0 && pool.ptr < pool.rows.length && left > 0) {
        const picked = pool.rows[pool.ptr];
        selectedBySource.get(pool.key).push(picked.data);
        pool.lastSelectedSourceRow = picked.sourceRow;
        pool.ptr += 1;
        pool.picked += 1;
        count -= 1;
        left -= 1;
        taken += 1;
      }
      return taken;
    }

    const fixedPools = pools.filter(p => !p.isFlexible && p.perEmployeeCount > 0);
    const flexiblePools = pools.filter(p => p.isFlexible);

    if (!fixedPools.length && flexiblePools.length) {
      // If all source rules are 0, pull equally from all listed sources.
      while (left > 0) {
        let progressed = false;
        for (let i = 0; i < pools.length && left > 0; i++) {
          if (takeFromPool(pools[i], 1) > 0) progressed = true;
        }
        if (!progressed) break;
      }
    } else {
      // First satisfy fixed per-employee counts from Config column F.
      for (let i = 0; i < fixedPools.length && left > 0; i++) {
        const p = fixedPools[i];
        const mustTake = p.perEmployeeCount * assigneeCount;
        takeFromPool(p, mustTake);
      }

      // Fill remaining requirement using flexible (F=0) sources first.
      while (left > 0 && flexiblePools.length) {
        let progressed = false;
        for (let i = 0; i < flexiblePools.length && left > 0; i++) {
          if (takeFromPool(flexiblePools[i], 1) > 0) progressed = true;
        }
        if (!progressed) break;
      }

      // Final fallback: use any source with available rows.
      while (left > 0) {
        let progressed = false;
        for (let i = 0; i < pools.length && left > 0; i++) {
          if (takeFromPool(pools[i], 1) > 0) progressed = true;
        }
        if (!progressed) break;
      }
    }

    // Append in config order so work sheet stays source-grouped (not scrambled).
    const selected = [];
    for (let i = 0; i < pools.length; i++) {
      const rows = selectedBySource.get(pools[i].key) || [];
      for (let r = 0; r < rows.length; r++) selected.push(rows[r]);
    }

    if (!selected.length) {
      Logger.log('Top-up required (%s) but no rows available in source sheets from current E pointers.', needed);
      return 0;
    }

    const normalized = _normalizeRowsToColumnCount(selected, Math.max(workCols, work.getLastColumn(), CFG.COLS.LINK_O));
    const start = Math.max(work.getLastRow() + 1, 3);
    work.getRange(start, 1, normalized.rows.length, normalized.cols).setValues(normalized.rows);
    _renumberWorkCountColumn();

    // Persist source progress in Config column E (last copied source row).
    const configSheet = ss.getSheetByName(CFG.SHEETS.CONFIG);
    if (configSheet) {
      for (let i = 0; i < pools.length; i++) {
        const p = pools[i];
        let nextLastCopied = _toNonNegativeInt(p.currentLastCopiedRow, 0);
        if (p.picked > 0 && p.lastSelectedSourceRow > 0) {
          nextLastCopied = p.lastSelectedSourceRow;
        }
        configSheet.getRange(p.configRow, CFG.CONFIG_COLS.SOURCE_LAST_COPIED_ROW).setValue(nextLastCopied);
        Logger.log('Source %s progress: start row=%s, picked=%s, last copied row(E)=%s', p.name, p.sourceStartRow, p.picked, nextLastCopied);
      }
    }

    const details = pools.map(function (p) {
      const rule = p.isFlexible ? 'F=0' : ('F=' + p.perEmployeeCount);
      return p.name + ' {picked=' + p.picked + ', ' + rule + ', skippedDuplicatePhone=' + p.skippedDuplicatePhone + '}';
    }).join(', ');
    Logger.log('Top-up added %s rows (needed=%s, target=%s, employees=%s). Source-wise pickup: %s', selected.length, needed, requiredTotal, assigneeCount, details);
    return selected.length;
  }

  function _selectRowsForAssignment(unassignedAll, workData, targetCount) {
    const indiamart = [];
    const bySource = new Map();

    for (let i = 0; i < unassignedAll.length; i++) {
      const idx = unassignedAll[i];
      const src = _getRowSourceLabel(workData[idx]);
      if (src === 'indiamart') {
        indiamart.push(idx);
        continue;
      }
      if (!bySource.has(src)) bySource.set(src, []);
      bySource.get(src).push(idx);
    }

    const selected = [];
    for (let i = 0; i < indiamart.length && selected.length < targetCount; i++) {
      selected.push(indiamart[i]);
    }

    const sources = Array.from(bySource.keys());
    let ptr = 0;
    while (selected.length < targetCount) {
      let progressed = false;
      for (let i = 0; i < sources.length && selected.length < targetCount; i++) {
        const s = sources[(ptr + i) % sources.length];
        const arr = bySource.get(s);
        if (!arr || !arr.length) continue;
        selected.push(arr.shift());
        progressed = true;
      }
      if (!progressed) break;
      ptr = (ptr + 1) % Math.max(1, sources.length);
    }

    return selected;
  }

  function _buildEmployeeBucketsBySource(unassignedIndexes, workData, assigneeCount, perEmployeeTarget) {
    const buckets = [];
    for (let i = 0; i < assigneeCount; i++) buckets.push([]);

    const bySource = new Map();
    for (let i = 0; i < unassignedIndexes.length; i++) {
      const idx = unassignedIndexes[i];
      const src = _getRowSourceLabel(workData[idx]);
      if (!bySource.has(src)) bySource.set(src, []);
      bySource.get(src).push(idx);
    }

    const sources = Array.from(bySource.keys());
    let start = 0;
    for (let si = 0; si < sources.length; si++) {
      const arr = bySource.get(sources[si]);
      for (let i = 0; i < arr.length; i++) {
        let best = -1;
        let bestLoad = Number.MAX_SAFE_INTEGER;

        // Keep total leads balanced, while preserving source round-robin behavior.
        for (let k = 0; k < assigneeCount; k++) {
          const empIdx = (start + k) % assigneeCount;
          const load = buckets[empIdx].length;
          if (perEmployeeTarget > 0 && load >= perEmployeeTarget) continue;
          if (load < bestLoad) {
            bestLoad = load;
            best = empIdx;
          }
        }

        if (best < 0) {
          for (let k = 0; k < assigneeCount; k++) {
            const empIdx = (start + k) % assigneeCount;
            const load = buckets[empIdx].length;
            if (load < bestLoad) {
              bestLoad = load;
              best = empIdx;
            }
          }
        }

        buckets[best].push(arr[i]);
        start = (best + 1) % assigneeCount;
      }
    }

    return buckets;
  }

  function _distributeUnassignedWorkRows() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const employees = _getEmployees();
    if (!employees.length) {
      Logger.log('No employee IDs available.');
      return;
    }

    let work = ss.getSheetByName(CFG.SHEETS.WORK);
    if (!work) work = ss.insertSheet(CFG.SHEETS.WORK);

    const workCols = Math.max(14, work.getLastColumn());

    // Build valid assignment targets first, then distribute round-robin for strict fairness.
    const assignees = [];
    for (let i = 0; i < employees.length; i++) {
      const emp = employees[i];
      let ext;
      try {
        ext = SpreadsheetApp.openById(emp.id);
      } catch (e) {
        Logger.log('Failed opening employee sheet %s: %s', emp.id, e.toString());
        continue;
      }

      const target = _getOrCreateSheet(ext, CFG.SHEETS.EXTERNAL);
      assignees.push({
        name: emp.name,
        id: emp.id,
        ext: ext,
        target: target
      });
    }

    if (!assignees.length) {
      Logger.log('No valid employee sheets available for distribution.');
      return;
    }

    const configuredLeadsPerEmployee = _getLeadsPerEmployeeFromConfig(CFG.DEFAULT_LEADS_PER_EMPLOYEE);
    const effectiveRule = _getEffectiveLeadsPerEmployeeFromSourceRules(configuredLeadsPerEmployee);
    const leadsPerEmployee = effectiveRule.leadsPerEmployee;
    if (effectiveRule.overrideByFixedRules) {
      Logger.log('Using fixed source rules only: per-employee target=%s from Config column F (H2 ignored for this run).', effectiveRule.fixedTotal);
    }
    const requiredTotal = assignees.length * leadsPerEmployee;
    const beforeWorkLast = work.getLastRow();
    const beforeData = beforeWorkLast >= 3 ? work.getRange(3, 1, beforeWorkLast - 2, workCols).getValues() : [];
    const beforeUnworked = _getAssignableWorkIndexes(beforeData).indexes.length;
    Logger.log('Distribution check: employees=%s, leadsPerEmployee(H2 or default)=%s, required rows=%s, current unworked in work=%s', assignees.length, leadsPerEmployee, requiredTotal, beforeUnworked);

    _topUpWorkFromConfigSheets(requiredTotal, assignees.length);

    // Re-read work after top-up.
    const refreshedLast = work.getLastRow();
    if (refreshedLast < 3) {
      Logger.log('No rows available in work even after top-up from Config sources.');
      return;
    }
    const refreshedData = work.getRange(3, 1, refreshedLast - 2, workCols).getValues();
    const assignableResult = _getAssignableWorkIndexes(refreshedData);
    const unassignedAll = assignableResult.indexes;
    Logger.log('Total unworked rows in work after top-up: %s; required rows: %s', unassignedAll.length, requiredTotal);

    if (assignableResult.skippedInvalid > 0) {
      Logger.log('Skipped %s invalid unassigned rows (missing phone).', assignableResult.skippedInvalid);
    }

    if (!unassignedAll.length) {
      Logger.log('No unassigned rows in work (column H empty).');
      return;
    }

    const targetCount = Math.min(requiredTotal, unassignedAll.length);
    const unassigned = _selectRowsForAssignment(
      unassignedAll,
      refreshedData,
      targetCount
    );
    if (unassignedAll.length < requiredTotal) {
      Logger.log('Available assignable rows %s are below required %s (%s x %s).', unassignedAll.length, requiredTotal, leadsPerEmployee, assignees.length);
    }

    const buckets = _buildEmployeeBucketsBySource(unassigned, refreshedData, assignees.length, leadsPerEmployee);

    const distributionSizes = buckets.map(b => b.length).join(',');
    Logger.log('Distribution plan (rows per employee in order): %s', distributionSizes);

    const assigned = [];

    for (let ei = 0; ei < assignees.length; ei++) {
      const emp = assignees[ei];
      const target = emp.target;
      const bucket = buckets[ei];
      if (!bucket.length) continue;

      const rowsToAppend = [];
      const srcRows = [];
      for (let k = 0; k < bucket.length; k++) {
        const idx = bucket[k];
        const sheetRow = 3 + idx;
        const row = work.getRange(sheetRow, 1, 1, workCols).getValues()[0];
        row[13] = emp.name;
        rowsToAppend.push(row);
        srcRows.push(sheetRow);
      }
      if (!rowsToAppend.length) continue;

      const startRow = Math.max(target.getLastRow() + 1, CFG.EMPLOYEE_DATA_START_ROW);
      let baseNum = 0;
      let suffix = '';
      for (let rr = target.getLastRow(); rr >= CFG.EMPLOYEE_DATA_START_ROW; rr--) {
        const v = target.getRange(rr, 2).getValue();
        if (v !== undefined && v !== null && String(v).trim() !== '') {
          const txt = String(v);
          const m = txt.match(/^(\d+)\s*(.*)$/);
          if (m) {
            baseNum = parseInt(m[1], 10);
            suffix = m[2] ? ' ' + m[2] : '';
          } else {
            const asNum = parseInt(txt, 10);
            if (!isNaN(asNum)) baseNum = asNum;
          }
          break;
        }
      }

      const prepared = [];
      const cols = Math.max(14, target.getLastColumn(), workCols);
      for (let i = 0; i < rowsToAppend.length; i++) {
        const r = rowsToAppend[i].slice();
        r[1] = String(baseNum + i + 1) + (suffix || '');
        while (r.length <= 13) r.push('');
        r[13] = '';
        prepared.push(r);
      }

      const preparedNormalized = _normalizeRowsToColumnCount(prepared, cols);

      try {
        target.getRange(startRow, 1, preparedNormalized.rows.length, preparedNormalized.cols).setValues(preparedNormalized.rows);

        // Copy dropdowns from source work rows to employee sheet for appended rows.
        try {
          const dvPerCol = new Array(preparedNormalized.cols).fill(null);
          for (let c = 0; c < preparedNormalized.cols; c++) {
            let newDv = null;
            for (let s = 0; s < srcRows.length; s++) {
              const srcCell = work.getRange(srcRows[s], c + 1);
              const dv = srcCell.getDataValidation();
              if (!dv) continue;

              const crit = dv.getCriteriaType();
              const critVals = dv.getCriteriaValues();
              if (crit === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
                const list = critVals[0] || [];
                newDv = SpreadsheetApp.newDataValidation().requireValueInList(list, true).build();
                break;
              }
              if (crit === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
                const rng = critVals[0];
                if (!rng || !rng.getValues) continue;
                const list = rng
                  .getValues()
                  .reduce((acc, row) => acc.concat(row), [])
                  .map(String)
                  .filter(v => v !== '');
                if (list.length) {
                  newDv = SpreadsheetApp.newDataValidation().requireValueInList(list, true).build();
                  break;
                }
              }
            }
            if (newDv) dvPerCol[c] = newDv;
          }

          for (let c = 0; c < preparedNormalized.cols; c++) {
            if (!dvPerCol[c]) continue;
            target.getRange(startRow, c + 1, preparedNormalized.rows.length, 1).setDataValidation(dvPerCol[c]);
          }
        } catch (e) {
          Logger.log('Dropdown copy to employee sheet failed for %s: %s', emp.id, e.toString());
        }

        for (let i = 0; i < srcRows.length; i++) {
          assigned.push({ row: srcRows[i], name: emp.name });
        }

        try {
          _setWhatsAppLinksInColumnO(target, startRow, preparedNormalized.rows);
        } catch (e) {
          Logger.log('Failed setting WhatsApp links in column O for %s: %s', emp.id, e.toString());
        }
      } catch (e) {
        Logger.log('Failed appending to employee sheet %s: %s', emp.id, e.toString());
      }
    }

    assigned.forEach(rec => {
      try {
        work.getRange(rec.row, 14).setValue(rec.name);
      } catch (e) {
        Logger.log('Failed marking work row %s: %s', rec.row, e.toString());
      }
    });

    Logger.log('Assigned total rows to employees: %s', assigned.length);
  }

  function runOnce() {
    Logger.log('runLeadPipelineV2Once started');

    // Step 1: move worked rows (B/H filled) to Followup, then clear LeadCallMsg data rows.
    try {
      _prepareEmployeeSheetsForNextCycle();
    } catch (e) {
      Logger.log('Prepare employee sheets step failed: %s', e.toString());
    }

    // Step 2: sync updates from Followup sheet to master work.
    try {
      _syncEmployeesToWorkInternal();
    } catch (e) {
      Logger.log('Employee sync step failed: %s', e.toString());
    }

    // Step 2 and 3: fetch IndiaMART and append to RawData + BuyLeads
    let fetchedRows = [];
    try {
      fetchedRows = _fetchIndiaMartRows();
      if (fetchedRows.length) _appendFetchedToRawAndBuy(fetchedRows);
      Logger.log('Fetched IndiaMART rows: %s', fetchedRows.length);
    } catch (e) {
      Logger.log('Fetch/store step failed: %s', e.toString());
    }

    // Step 4: MasterSheet-style move from BuyLeads to work and clear BuyLeads
    try {
      _moveBuyLeadsToWorkAndClearBuyLeads();
    } catch (e) {
      Logger.log('Move BuyLeads -> work step failed: %s', e.toString());
    }

    try {
      _removeBlockedIndiaMartRowsFromWork();
    } catch (e) {
      Logger.log('Blocked IndiaMART cleanup step failed: %s', e.toString());
    }

    // Step 5: divide all work rows where H is empty and push to employees
    try {
      _distributeUnassignedWorkRows();
    } catch (e) {
      Logger.log('Distribution step failed: %s', e.toString());
    }

    Logger.log('runLeadPipelineV2Once completed');
  }

  function _getNextDailyRunDateAtHour(hour) {
    const now = new Date();
    const next = new Date(now);
    next.setHours(hour, 0, 0, 0);
    if (next.getTime() <= now.getTime()) {
      next.setDate(next.getDate() + 1);
    }
    return next;
  }

  function _clearDailyTriggers() {
    const handlers = new Set([
      'completeRun',
      'runLeadPipelineV2ScheduledDailyRun',
      'runLeadPipelineV2Once'
    ]);
    const existing = ScriptApp.getProjectTriggers().filter(t => handlers.has(t.getHandlerFunction()));
    existing.forEach(t => ScriptApp.deleteTrigger(t));
    return existing.length;
  }

  function _scheduleNextDailyRun() {
    _clearDailyTriggers();
    const nextRun = _getNextDailyRunDateAtHour(CFG.DAILY_RUN_HOUR);
    ScriptApp
      .newTrigger('completeRun')
      .timeBased()
      .at(nextRun)
      .create();

    Logger.log(
      'Lead pipeline next daily run scheduled for %s (script timezone: %s).',
      Utilities.formatDate(nextRun, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
      Session.getScriptTimeZone()
    );
  }

  function rescheduleIfDailyEnabled() {
    const props = PropertiesService.getScriptProperties();
    const enabled = props.getProperty(CFG.PROPS.DAILY_TRIGGER_ENABLED) === 'true';
    if (enabled) {
      _scheduleNextDailyRun();
    }
  }

  function toggleTrigger() {
    const props = PropertiesService.getScriptProperties();
    const enabled = props.getProperty(CFG.PROPS.DAILY_TRIGGER_ENABLED) === 'true';

    if (enabled) {
      const removed = _clearDailyTriggers();
      props.setProperty(CFG.PROPS.DAILY_TRIGGER_ENABLED, 'false');
      Logger.log('Lead pipeline daily trigger stopped.');
      Logger.log('Removed %s scheduled daily trigger(s).', removed);
      return;
    }

    props.setProperty(CFG.PROPS.DAILY_TRIGGER_ENABLED, 'true');
    _scheduleNextDailyRun();
    Logger.log('Lead pipeline daily trigger started. It will run every day at %s:00 in script timezone.', CFG.DAILY_RUN_HOUR);
  }

  
  return {
    runOnce: runOnce,
    toggleTrigger: toggleTrigger,
    rescheduleIfDailyEnabled: rescheduleIfDailyEnabled
  };
})();

function dailytrigger() {
  LeadPipelineV2.toggleTrigger();
}

function completeRun() {
  try {
    LeadPipelineV2.runOnce();
  } finally {
    LeadPipelineV2.rescheduleIfDailyEnabled();
  }
}