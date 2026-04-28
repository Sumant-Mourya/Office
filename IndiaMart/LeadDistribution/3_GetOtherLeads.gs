const EMP_REFRESH_PIPELINE = {
  MAIN_SPREADSHEET_ID: '1OPq4GtU-wrQFUbN9Zvgpc_WO_O6fvhl0MUQKeTz1rGw',

  SHEETS: {
    CONFIG: 'Config',
    WORK: 'Work',
    EMPLOYEE_TARGET: 'LeadCallMsg'
  },

  COL: {
    D: 4,
    J: 10,
    K: 11,
    L: 12
  },

  EMP_COL: {
    H: 8,
    I: 9,
    J: 10
  },

};

function runEmployeeRefreshPipeline() {
  const activeSS = SpreadsheetApp.getActiveSpreadsheet();
  const employeeSheet = empRefreshGetOrCreateSheet_(activeSS, EMP_REFRESH_PIPELINE.SHEETS.EMPLOYEE_TARGET);

  const currentSpreadsheetId = activeSS.getId();

  const mainSS = SpreadsheetApp.openById(EMP_REFRESH_PIPELINE.MAIN_SPREADSHEET_ID);
  const configSheet = mainSS.getSheetByName(EMP_REFRESH_PIPELINE.SHEETS.CONFIG);
  const workSheet = mainSS.getSheetByName(EMP_REFRESH_PIPELINE.SHEETS.WORK);

  const configData = configSheet.getDataRange().getValues();

  let employeeRowIndex = -1;

  // 🔍 Find employee row in config (Column B)
  for (let i = 1; i < configData.length; i++) {
    const sheetId = String(configData[i][1]).trim();
    if (sheetId === currentSpreadsheetId) {
      employeeRowIndex = i;
      break;
    }
  }

  if (employeeRowIndex === -1) {
    Logger.log("Employee not found in config");
    return;
  }

  const row = configData[employeeRowIndex];
  const employeeName = empRefreshToCleanString_(row[0]); // Config column A

  let currentCount = Number(row[3]) || 0; // Column D

  // Employee-specific limit: read from main Config sheet cell N2 (column 14, row 2).
  const limitCell = configSheet.getRange(2, 14);
  const limitRaw = limitCell.getValue();
  const limitDisplay = limitCell.getDisplayValue();
  let employeeLimit;

  if (typeof limitRaw === 'number' && !isNaN(limitRaw)) {
    employeeLimit = Number(limitRaw);
  } else {
    const parsed = parseInt(String(limitDisplay).replace(/[^0-9-]/g, ''), 10);
    if (!isNaN(parsed) && parsed > 0) {
      employeeLimit = parsed;
    } else {
      employeeLimit = 50;
    }
  }

  Logger.log('Refresh: activeSS=' + currentSpreadsheetId + ' configRow=' + (employeeRowIndex + 1) + ' currentCount=' + currentCount + ' limitSource=Config!N2 limitRaw=' + limitRaw + ' limitDisplay=' + limitDisplay + ' resolvedLimit=' + employeeLimit);

  if (currentCount >= employeeLimit) {
    Logger.log('Limit reached (' + employeeLimit + ')');
    return;
  }

  let remainingLimit = employeeLimit - currentCount;

    // Collect all source rows present in Config column J (process them sequentially)
    const sources = [];
    for (let i = 1; i < configData.length; i++) {
      const srcName = String(configData[i][9] || '').trim();
      if (!srcName) continue;
      const srcLast = Number(configData[i][10]) || 0; // K
      const srcMax = Number(configData[i][11]) || 0; // L
      sources.push({ sheetRow: i + 1, name: srcName, lastProcessed: srcLast, maxFetch: srcMax });
    }

    if (sources.length === 0) {
      Logger.log('No source sheets configured in Config column J');
      return;
    }

    let totalFetched = 0;
    let totalFromWork = 0;

    for (let s = 0; s < sources.length && remainingLimit > 0; s++) {
      const src = sources[s];
      if (!src.maxFetch || src.maxFetch <= 0) {
        continue;
      }

      const requiredForSource = Math.min(src.maxFetch, remainingLimit);
      let fulfilledFromWork = 0;

      // 1) Work-first: consume matching unworked rows from Work (G=source, H empty)
      const workMatches = empRefreshFindWorkRowsForSource_(workSheet, src.name, requiredForSource);
      if (workMatches.length > 0) {
        empRefreshAppendRowsFromRowObjectsWithValidations_(employeeSheet, workSheet, workMatches);
        empRefreshAssignWorkRowsToEmployee_(workSheet, workMatches, employeeName);
        fulfilledFromWork = workMatches.length;
        totalFromWork += fulfilledFromWork;
        remainingLimit -= fulfilledFromWork;
        Logger.log('EmployeeRefresh: source=' + src.name + ' fulfilled_from_work=' + fulfilledFromWork + ' required=' + requiredForSource + ' (skipping source sheet fetch)');
        continue;
      }

      const requiredFromSource = requiredForSource;

      const sourceSheet = mainSS.getSheetByName(src.name);
      if (!sourceSheet) {
        Logger.log('Source sheet not found: ' + src.name);
        continue;
      }

      const lastRow = sourceSheet.getLastRow();
      const startRow = Math.max(2, Number(src.lastProcessed) + 1); // always skip header

      Logger.log('EmployeeRefresh: processing source=' + src.name + ' sheetRow=' + src.sheetRow + ' startRow=' + startRow + ' lastRow=' + lastRow + ' maxFetch=' + src.maxFetch + ' remainingLimit=' + remainingLimit);

      if (startRow > lastRow) {
        Logger.log('No new data for ' + src.name);
        continue;
      }

      const available = lastRow - startRow + 1;
      const take = Math.min(requiredFromSource, remainingLimit, available);
      if (take <= 0) {
        Logger.log('Nothing to fetch from ' + src.name);
        continue;
      }

      const width = sourceSheet.getLastColumn();
      const rows = sourceSheet.getRange(startRow, 1, take, width).getValues();

      // Transform rows: clean phone (col E -> index 4) and format date (col F -> index 5), set source in col G
      for (let r = 0; r < rows.length; r++) {
        const rowVals = rows[r];
        // Phone: keep digits and take last 10 if longer
        try {
          const rawPhone = rowVals[4];
          let digits = '';
          if (rawPhone !== undefined && rawPhone !== null) digits = String(rawPhone).replace(/\D/g, '');
          if (digits.length > 10) digits = digits.slice(-10);
          rowVals[4] = digits;
        } catch (e) {
          // ignore
        }

        // Lead date: normalize to dd-MM-yyyy
        try {
          const rawDate = rowVals[5];
          const formatted = empRefreshFormatToDDMMYYYY_(rawDate);
          rowVals[5] = formatted;
        } catch (e) {
          // ignore
        }

        // Source: keep source value from the sheet (do not force case), so validations match
      }

      // Append transformed rows to Work and employee LeadCallMsg, preserving dropdowns
      // Work rows get assigned employee name in column N.
      empRefreshAppendRowsWithValidations_(workSheet, sourceSheet, startRow, rows, { assignedToName: employeeName });
      empRefreshAppendRowsWithValidations_(employeeSheet, sourceSheet, startRow, rows);

      // Update this source's last-processed (column K)
      const newLast = startRow + take - 1;
      configSheet.getRange(src.sheetRow, 11).setValue(newLast);

      totalFetched += take;
      remainingLimit -= take;
    }

    // Update employee distributed count (column D) for both source-fetched and work-assigned rows.
    const totalAssignedToEmployee = totalFetched + totalFromWork;
    if (totalAssignedToEmployee > 0) {
      const countCell = configSheet.getRange(employeeRowIndex + 1, 4);
      const latestCount = Number(countCell.getValue()) || 0;
      countCell.setValue(latestCount + totalAssignedToEmployee);
    }

    Logger.log('Fetched total ' + totalFetched + ' rows from source sheets and assigned ' + totalFromWork + ' rows from Work for employee ' + currentSpreadsheetId);
}

function empRefreshAllRowsWorked_(sheet) {
  const data = empRefreshGetSheetDataWithoutHeader_(sheet);
  if (data.length === 0) {
    return true;
  }

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const h = empRefreshToCleanString_(row[EMP_REFRESH_PIPELINE.EMP_COL.H - 1]);
    const ii = empRefreshToCleanString_(row[EMP_REFRESH_PIPELINE.EMP_COL.I - 1]);
    const j = empRefreshToCleanString_(row[EMP_REFRESH_PIPELINE.EMP_COL.J - 1]);

    const hasAnyLeadData = row.some(v => empRefreshToCleanString_(v) !== '');
    if (!hasAnyLeadData) {
      continue;
    }

    if (!h || !ii || !j) {
      return false;
    }
  }

  return true;
}

function empRefreshGetOrCreateSheet_(ss, sheetName) {
  let sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
  }
  return sh;
}

function empRefreshGetSheetDataWithoutHeader_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    return [];
  }
  return sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
}

function empRefreshAppendRows_(sheet, rows) {
  if (!Array.isArray(rows) || rows.length === 0) {
    return;
  }

  const normalized = empRefreshNormalizeRows_(rows, 1);
  const startRow = Math.max(sheet.getLastRow() + 1, 2);
  sheet.getRange(startRow, 1, normalized.length, normalized[0].length).setValues(normalized);
}

function empRefreshAppendRowsWithValidations_(targetSheet, sourceSheet, sourceStartRow, rows, options) {
  if (!Array.isArray(rows) || rows.length === 0) return;

  const normalized = empRefreshNormalizeRows_(rows, 1);
  const rowCount = normalized.length;
  let colCount = normalized[0].length;
  const targetStartRow = Math.max(targetSheet.getLastRow() + 1, 2);

  for (let r = 0; r < rowCount; r++) {
    // Column B should always be current row + 1.
    if (normalized[r].length >= 2) {
      normalized[r][1] = targetStartRow + r + 1;
    }
    // Optional: when writing into Work, fill column N with employee name.
    if (options && options.assignedToName) {
      while (normalized[r].length < 14) {
        normalized[r].push('');
      }
      normalized[r][13] = options.assignedToName;
    }
  }

  colCount = normalized[0].length;

  targetSheet.getRange(targetStartRow, 1, rowCount, colCount).setValues(normalized);
  // Keep column F consistently displayed in dd-MM-yyyy.
  if (colCount >= 6) {
    targetSheet.getRange(targetStartRow, 6, rowCount, 1).setNumberFormat('dd-MM-yyyy');
  }

  if (options && options.assignedToName) {
    targetSheet.getRange(targetStartRow, 14, rowCount, 1).setValues(Array(rowCount).fill([options.assignedToName]));
  }

  // Try to copy data validations safely. Convert any "range-based" validations into a list-of-items
  try {
    const sourceLastCol = sourceSheet.getLastColumn();
    const sourceRange = sourceSheet.getRange(sourceStartRow, 1, rowCount, sourceLastCol);
    const sourceValidations = sourceRange.getDataValidations();

    const targetValidations = [];
    for (let r = 0; r < rowCount; r++) {
      const srcRow = sourceValidations[r] || [];
      const tgtRow = [];
      for (let c = 0; c < colCount; c++) {
        const dv = srcRow[c] || null;
        if (!dv) {
          tgtRow.push(null);
          continue;
        }

        let newDv = null;
        try {
          const type = dv.getCriteriaType();
          const crit = dv.getCriteriaValues();
          if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
            const items = crit && crit[0] ? crit[0] : [];
            newDv = SpreadsheetApp.newDataValidation().requireValueInList(items, true).build();
          } else if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
            const srcRange = crit && crit[0] ? crit[0] : null;
            if (srcRange && typeof srcRange.getValues === 'function') {
              const vals2d = srcRange.getValues();
              const items = [].concat.apply([], vals2d).map(v => (v === null || v === undefined) ? '' : String(v)).filter(x => x !== '');
              if (items.length > 0) {
                newDv = SpreadsheetApp.newDataValidation().requireValueInList(items, true).build();
              } else {
                newDv = null;
              }
            }
          } else {
            // unsupported types: skip
            newDv = null;
          }
        } catch (e) {
          newDv = null;
        }

        tgtRow.push(newDv);
      }
      targetValidations.push(tgtRow);
    }

    // Apply validations (cells with null will have no validation)
    targetSheet.getRange(targetStartRow, 1, targetValidations.length, colCount).setDataValidations(targetValidations);
  } catch (e) {
    Logger.log('empRefreshAppendRowsWithValidations_ failed to copy validations: ' + e);
  }
}

function empRefreshFindWorkRowsForSource_(workSheet, sourceName, limit) {
  if (!workSheet || limit <= 0) return [];
  const lastRow = workSheet.getLastRow();
  const lastCol = Math.max(workSheet.getLastColumn(), 14);
  if (lastRow < 2 || lastCol < 1) return [];

  const data = workSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const displayData = workSheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
  const out = [];
  const wanted = empRefreshNormalizeSource_(sourceName);

  for (let i = 0; i < data.length && out.length < limit; i++) {
    const row = data[i];
    const displayRow = displayData[i] || [];
    const src = empRefreshNormalizeSource_(row[6]); // G
    const hRaw = row[7]; // H
    const hDisplay = empRefreshToCleanString_(displayRow[7]);
    const nDisplay = empRefreshToCleanString_(displayRow[13]); // N
    if (src !== wanted) continue;
    if (hDisplay !== '') continue;
    if (hRaw !== '' && hRaw !== null && hRaw !== false) continue;
    if (nDisplay !== '') continue;
    out.push({ sheetRow: i + 2, values: row });
  }

  return out;
}

function empRefreshAssignWorkRowsToEmployee_(workSheet, rowObjects, employeeName) {
  if (!Array.isArray(rowObjects) || rowObjects.length === 0) return;

  for (let i = 0; i < rowObjects.length; i++) {
    const rowNo = rowObjects[i].sheetRow;
    workSheet.getRange(rowNo, 14).setValue(employeeName); // N
    workSheet.getRange(rowNo, 2).setValue(rowNo + 1); // B = current row + 1
  }
}

function empRefreshAppendRowsFromRowObjectsWithValidations_(targetSheet, sourceSheet, rowObjects) {
  if (!Array.isArray(rowObjects) || rowObjects.length === 0) return;

  const rows = rowObjects.map(x => (x.values || []).slice());
  const normalized = empRefreshNormalizeRows_(rows, 1);
  const rowCount = normalized.length;
  let colCount = normalized[0].length;
  const targetStartRow = Math.max(targetSheet.getLastRow() + 1, 2);

  for (let i = 0; i < rowCount; i++) {
    if (normalized[i].length >= 2) {
      normalized[i][1] = targetStartRow + i + 1;
    }
  }

  colCount = normalized[0].length;
  targetSheet.getRange(targetStartRow, 1, rowCount, colCount).setValues(normalized);
  if (colCount >= 6) {
    targetSheet.getRange(targetStartRow, 6, rowCount, 1).setNumberFormat('dd-MM-yyyy');
  }

  // Copy validations row-by-row from Work
  const validations = [];
  for (let i = 0; i < rowObjects.length; i++) {
    const srcRowNum = rowObjects[i].sheetRow;
    const dvRow = sourceSheet.getRange(srcRowNum, 1, 1, colCount).getDataValidations()[0] || [];
    const outDvRow = [];
    for (let c = 0; c < colCount; c++) {
      const dv = dvRow[c] || null;
      if (!dv) {
        outDvRow.push(null);
        continue;
      }

      let newDv = null;
      try {
        const type = dv.getCriteriaType();
        const crit = dv.getCriteriaValues();
        if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
          const items = crit && crit[0] ? crit[0] : [];
          newDv = SpreadsheetApp.newDataValidation().requireValueInList(items, true).build();
        } else if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
          const srcRange = crit && crit[0] ? crit[0] : null;
          if (srcRange && typeof srcRange.getValues === 'function') {
            const vals2d = srcRange.getValues();
            const items = [].concat.apply([], vals2d).map(v => (v === null || v === undefined) ? '' : String(v)).filter(x => x !== '');
            if (items.length > 0) {
              newDv = SpreadsheetApp.newDataValidation().requireValueInList(items, true).build();
            }
          }
        }
      } catch (e) {
        newDv = null;
      }
      outDvRow.push(newDv);
    }
    validations.push(outDvRow);
  }

  if (validations.length > 0) {
    targetSheet.getRange(targetStartRow, 1, validations.length, colCount).setDataValidations(validations);
  }
}

function empRefreshNormalizeSource_(val) {
  return empRefreshToCleanString_(val).toLowerCase();
}

function empRefreshNormalizeRows_(rows, minCols) {
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

function empRefreshToCleanString_(val) {
  if (val === undefined || val === null) {
    return '';
  }
  return String(val).trim();
}

function empRefreshToNonNegativeNumber_(val, fallback) {
  const n = Number(val);
  if (isNaN(n) || n < 0) {
    return fallback;
  }
  return n;
}

function empRefreshFormatToDDMMYYYY_(val) {
  if (val === undefined || val === null) {
    return '';
  }

  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd-MM-yyyy');
  }

  const s = String(val).trim();
  if (!s) return '';

  // Excel/Sheets serial date support.
  const serial = Number(s);
  if (!isNaN(serial) && serial > 20000 && serial < 100000) {
    const ms = Math.round((serial - 25569) * 86400 * 1000);
    const dt = new Date(ms);
    if (!isNaN(dt.getTime())) {
      return Utilities.formatDate(dt, Session.getScriptTimeZone(), 'dd-MM-yyyy');
    }
  }

  const parsed = new Date(s);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'dd-MM-yyyy');
  }

  const m = s.match(/^(\d{1,4})[\/\-.](\d{1,2})[\/\-.](\d{1,4})/);
  if (m) {
    const a = m[1], b = m[2], c = m[3];
    if (a.length === 4) {
      return c.padStart(2, '0') + '-' + b.padStart(2, '0') + '-' + a.padStart(4, '0');
    }
    if (c.length === 4) {
      return a.padStart(2, '0') + '-' + b.padStart(2, '0') + '-' + c.padStart(4, '0');
    }
  }

  return s;
}
