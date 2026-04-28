const DISTRIBUTION_PIPELINE = {
  MAIN_SPREADSHEET_ID: '1OPq4GtU-wrQFUbN9Zvgpc_WO_O6fvhl0MUQKeTz1rGw',

  SHEETS: {
    CONFIG: 'Config',
    WORK: 'Work',
    EMPLOYEE_TARGET: 'LeadCallMsg'
  },

  COL: {
    EMPLOYEE_NAME: 1, // A
    EMPLOYEE_SHEET_ID: 2, // B
    IS_PRESENT: 3, // C
    DISTRIBUTED_COUNT: 4, // D
    WEIGHT: 5 // E
  },

  EMPLOYEE_SHEET_DATA_START_ROW: 2,
  ASSIGNED_TO_COL: 14, // N
  PER_EMPLOYEE_LIMIT: 50
};
function runLeadDistributionFromWork() {
  const ss = SpreadsheetApp.openById(DISTRIBUTION_PIPELINE.MAIN_SPREADSHEET_ID);
  const configSheet = distributionGetOrCreateSheet_(ss, DISTRIBUTION_PIPELINE.SHEETS.CONFIG);
  const workSheet = distributionGetOrCreateSheet_(ss, DISTRIBUTION_PIPELINE.SHEETS.WORK);
  const configLastRow = configSheet.getLastRow();
  if (configLastRow < 2) {
    Logger.log('Config has no employee rows.');
    return;
  }

  const configData = configSheet
    .getRange(
      2,
      DISTRIBUTION_PIPELINE.COL.EMPLOYEE_NAME,
      configLastRow - 1,
      DISTRIBUTION_PIPELINE.COL.WEIGHT - DISTRIBUTION_PIPELINE.COL.EMPLOYEE_NAME + 1
    )
    .getValues();

  const employees = [];

  for (let i = 0; i < configData.length; i++) {
    const row = configData[i];
    const employeeName = distributionToCleanString_(row[0]); // A
    const sheetId = distributionToCleanString_(row[1]); // B
    const presentRaw = distributionToCleanString_(row[2]); // C
    const currentCount = distributionToNonNegativeNumber_(row[3], 0); // D - already distributed
    const weight = distributionToNonNegativeNumber_(row[4], 0); // E

    if (!sheetId) {
      continue;
    }

    employees.push({
      configRow: i + 2,
      employeeName: employeeName,
      sheetId: sheetId,
      present: distributionIsPresentFlag_(presentRaw),
      weight: weight,
      current: currentCount,
      capacity: Math.max(0, DISTRIBUTION_PIPELINE.PER_EMPLOYEE_LIMIT - currentCount),
      assigned: 0
    });
  }

  const presentEmployees = employees.filter(e => e.present);
  if (presentEmployees.length === 0) {
    Logger.log('No present employees in Config column C. No distribution done.');
    distributionWriteDistributedCounts_(configSheet, employees);
    return;
  }

  const workLastRow = workSheet.getLastRow();
  const workLastCol = workSheet.getLastColumn();

  // fetch existing assigned-to column (N) values so we can update them in batch
  let workAssigned = [];
  if (workLastRow >= 2) {
    workAssigned = workSheet.getRange(2, DISTRIBUTION_PIPELINE.ASSIGNED_TO_COL, workLastRow - 1, 1).getValues();
  }

  const workData = distributionGetSheetDataWithoutHeader_(workSheet);
  const indiamartLeads = [];

  for (let i = 0; i < workData.length; i++) {
    const row = workData[i];
    const hasAnyData = row.some(v => distributionToCleanString_(v) !== '');
    if (!hasAnyData) {
      continue;
    }

    // Only consider leads whose column H is empty
    const colH = distributionToCleanString_(row[7]); // H is index 7
    if (colH !== '') {
      continue;
    }

    const src = distributionNormalizeSourceKey_(row[6]); // G is index 6
    if (!distributionIsIndiaMartSource_(src)) {
      continue;
    }

    const entry = { values: row, sheetRow: i + 2 };
    indiamartLeads.push(entry);
  }

  if (indiamartLeads.length === 0) {
    Logger.log('No unworked indiamart leads in Work sheet.');
    distributionWriteDistributedCounts_(configSheet, employees);
    return;
  }

  // Prepare per-employee buckets for assigned rows
  const assignedRowsPerEmp = [];
  for (let i = 0; i < presentEmployees.length; i++) {
    assignedRowsPerEmp.push([]);
  }

  // Distribute only indiamart leads by weight among present employees.
  const imCounts = distributionCalculateWeightedShares_(indiamartLeads.length, presentEmployees);
  for (let i = 0; i < presentEmployees.length; i++) {
    const cnt = imCounts[i] || 0;
    for (let k = 0; k < cnt; k++) {
      const rowObj = indiamartLeads.shift();
      if (!rowObj) break;
      assignedRowsPerEmp[i].push(rowObj);
      presentEmployees[i].assigned += 1;
    }
  }

  // 3) Commit assignments: append to employee sheets and mark Work column N
  let totalDistributed = 0;
  for (let i = 0; i < presentEmployees.length; i++) {
    const emp = presentEmployees[i];
    const rowsForEmployee = assignedRowsPerEmp[i] || [];
    if (rowsForEmployee.length === 0) {
      emp.assigned = 0;
      continue;
    }

    const empSS = SpreadsheetApp.openById(emp.sheetId);
    const empSheet = distributionGetOrCreateSheet_(empSS, DISTRIBUTION_PIPELINE.SHEETS.EMPLOYEE_TARGET);
    distributionAppendRowsFromWork_(workSheet, empSheet, rowsForEmployee, emp.employeeName);
    emp.assigned = rowsForEmployee.length;
    totalDistributed += emp.assigned;

    // mark assigned name in Work column N for these rows
    for (let r = 0; r < rowsForEmployee.length; r++) {
      const sheetRow = rowsForEmployee[r].sheetRow;
      const idx = sheetRow - 2; // index into workAssigned
      if (idx >= 0) {
        while (workAssigned.length <= idx) {
          workAssigned.push(['']);
        }
        workAssigned[idx][0] = emp.employeeName;
      }
    }
  }

  distributionWriteDistributedCounts_(configSheet, employees);
  Logger.log('Distribution complete. total_distributed=' + totalDistributed);

  // write back updated Assigned-To column (N) into Work sheet
  if (workLastRow >= 2) {
    const rowsToWrite = workLastRow - 1;
    const writeValues = [];
    for (let i = 0; i < rowsToWrite; i++) {
      writeValues.push(workAssigned[i] || ['']);
    }
    workSheet.getRange(2, DISTRIBUTION_PIPELINE.ASSIGNED_TO_COL, rowsToWrite, 1).setValues(writeValues);
  }
}

function distributionNormalizeSourceKey_(val) {
  const t = distributionToCleanString_(val).toLowerCase();
  return t || 'unknown';
}

function distributionIsIndiaMartSource_(sourceKey) {
  return distributionNormalizeSourceKey_(sourceKey) === 'indiamart';
}

function distributionCalculateWeightedShares_(totalLeads, employees) {
  if (employees.length === 0) {
    return [];
  }

  let totalWeight = 0;
  for (let i = 0; i < employees.length; i++) {
    totalWeight += distributionToNonNegativeNumber_(employees[i].weight, 0);
  }

  const useEqual = totalWeight <= 0;
  const raw = [];
  let assigned = 0;

  for (let i = 0; i < employees.length; i++) {
    const w = useEqual ? 1 : distributionToNonNegativeNumber_(employees[i].weight, 0);
    const share = useEqual ? totalLeads / employees.length : (totalLeads * w) / totalWeight;

    const floored = Math.floor(share);
    raw.push({ idx: i, base: floored, frac: share - floored });
    assigned += floored;
  }

  let remaining = totalLeads - assigned;
  raw.sort((a, b) => b.frac - a.frac);

  for (let i = 0; i < raw.length && remaining > 0; i++) {
    raw[i].base += 1;
    remaining -= 1;
  }

  raw.sort((a, b) => a.idx - b.idx);
  return raw.map(x => x.base);
}

function distributionWriteDistributedCounts_(configSheet, employees) {
  for (let i = 0; i < employees.length; i++) {
    const e = employees[i];
    const cell = configSheet.getRange(e.configRow, DISTRIBUTION_PIPELINE.COL.DISTRIBUTED_COUNT);
    const prev = Number(cell.getValue()) || 0;
    cell.setValue(prev + (e.assigned || 0));
  }
}

function distributionGetOrCreateSheet_(ss, sheetName) {
  let sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
  }
  return sh;
}

function distributionGetSheetDataWithoutHeader_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    return [];
  }
  return sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
}

function distributionAppendRows_(sheet, rows) {
  if (!Array.isArray(rows) || rows.length === 0) {
    return;
  }

  const normalized = distributionNormalizeRows_(rows, 1);
  const startRow = Math.max(sheet.getLastRow() + 1, 2);
  sheet.getRange(startRow, 1, normalized.length, normalized[0].length).setValues(normalized);
}

function distributionAppendRowsFromWork_(workSheet, employeeSheet, leadRows, employeeName) {
  if (!Array.isArray(leadRows) || leadRows.length === 0) {
    return;
  }

  const employeeLastRow = employeeSheet.getLastRow();
  const startRow = Math.max(employeeLastRow + 1, DISTRIBUTION_PIPELINE.EMPLOYEE_SHEET_DATA_START_ROW);
  const workLastCol = Math.max(workSheet.getLastColumn(), DISTRIBUTION_PIPELINE.ASSIGNED_TO_COL);
  const preparedRows = [];
  const validations = [];

  // determine starting count for Count column (col 2)
  let startCount = 0;
  if (employeeLastRow >= DISTRIBUTION_PIPELINE.EMPLOYEE_SHEET_DATA_START_ROW) {
    const lastCountVal = employeeSheet.getRange(employeeLastRow, 2).getValue();
    startCount = Number(lastCountVal) || (employeeLastRow - DISTRIBUTION_PIPELINE.EMPLOYEE_SHEET_DATA_START_ROW + 1);
  } else {
    startCount = 0;
  }

  for (let i = 0; i < leadRows.length; i++) {
    const source = leadRows[i];
    const output = source.values.slice();

    while (output.length < workLastCol) {
      output.push('');
    }

    // set Count as continuation from existing last count
    output[1] = startCount + i + 1;
    // Keep assignee name only in Work sheet; do not copy column N into employee sheet.
    output[DISTRIBUTION_PIPELINE.ASSIGNED_TO_COL - 1] = '';
    preparedRows.push(output);

    const dvRow = workSheet.getRange(source.sheetRow, 1, 1, workLastCol).getDataValidations()[0] || [];
    const convRow = [];
    for (let c = 0; c < workLastCol; c++) {
      const dvCell = dvRow[c] || null;
      if (!dvCell) {
        convRow.push(null);
        continue;
      }

      let newDv = null;
      try {
        const type = dvCell.getCriteriaType();
        const crit = dvCell.getCriteriaValues();
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
      convRow.push(newDv);
    }
    validations.push(convRow);
  }

  employeeSheet.getRange(startRow, 1, preparedRows.length, workLastCol).setValues(preparedRows);
  if (validations.length > 0) {
    employeeSheet.getRange(startRow, 1, validations.length, workLastCol).setDataValidations(validations);
  }
}

function distributionNormalizeRows_(rows, minCols) {
  let width = Math.max(1, minCols || 1);
  for (let i = 0; i < rows.length; i++) {
    width = Math.max(width, rows[i].length);
  }

  const output = [];
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i].slice();
    while (r.length < width) {
      r.push('');
    }
    if (r.length > width) {
      r.length = width;
    }
    output.push(r);
  }
  return output;
}

function distributionIsPresentFlag_(val) {
  const t = distributionToCleanString_(val).toLowerCase();
  return t === 'yes' || t === 'y' || t === '1' || t === 'true' || t === 'present';
}

function distributionToCleanString_(val) {
  if (val === undefined || val === null) {
    return '';
  }
  return String(val).trim();
}

function distributionToNonNegativeNumber_(val, fallback) {
  const n = Number(val);
  if (isNaN(n) || n < 0) {
    return fallback;
  }
  return n;
}
