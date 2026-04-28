const LEAD_FLOW = {
  LEAD_SHEET: 'LeadCallMsg',
  FOLLOWUP_SHEET: 'followup',
  HISTORY_SHEET: 'History',
  START_ROW: 3,
  PHONE_COL: 5,   // E
  WHATSAPP_COL: 8, // H
  CALLING_COL: 9,  // I
  STATUS_COL: 10   // J
};

function onEdit(e) { 
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row < LEAD_FLOW.START_ROW) return;

  const name = sheet.getName();
  const isLead = name === LEAD_FLOW.LEAD_SHEET;
  const isFollowup = name.toLowerCase() === LEAD_FLOW.FOLLOWUP_SHEET;
  if (!isLead && !isFollowup) return;

  const ss = sheet.getParent();
  const history = ss.getSheetByName(LEAD_FLOW.HISTORY_SHEET);
  if (!history) return;

  const oldValue = String(e.oldValue || '').trim();
  const newValue = String(e.value || '').trim();

  if (isLead) {
    handleLeadCallMsgEdit_(sheet, history, row, col, oldValue, newValue);
    return;
  }

  handleFollowupEdit_(sheet, history, row, col, oldValue, newValue);
}

function handleLeadCallMsgEdit_(sheet, historySheet, row, col, oldValue, newValue) {
  if (col !== LEAD_FLOW.WHATSAPP_COL && col !== LEAD_FLOW.CALLING_COL && col !== LEAD_FLOW.STATUS_COL) {
    return;
  }

  const rowPack = getRowPack_(sheet, row);
  if (!rowPack.phone) return;

  if (col === LEAD_FLOW.WHATSAPP_COL || col === LEAD_FLOW.CALLING_COL) {
    sheet.getRange(row, LEAD_FLOW.STATUS_COL).setValue('Follow up');
    removeRowsByPhone_(historySheet, rowPack.phone);
    return;
  }

  if (newValue === 'Follow up') {
    removeRowsByPhone_(historySheet, rowPack.phone);
    return;
  }

  if (newValue === 'Closed') {
    if (!isRowAtoIFilled_(rowPack.values)) {
      sheet.getRange(row, LEAD_FLOW.STATUS_COL).setValue(oldValue || '');
      return;
    }
    copyRowToHistoryIfMissing_(historySheet, rowPack);
    return;
  }

  if (newValue === 'Pending' || (oldValue === 'Closed' && newValue !== 'Closed')) {
    removeRowsByPhone_(historySheet, rowPack.phone);
  }
}

function handleFollowupEdit_(sheet, historySheet, row, col, oldValue, newValue) {
  if (col !== LEAD_FLOW.STATUS_COL) return;

  const rowPack = getRowPack_(sheet, row);
  if (!rowPack.phone) return;

  if (newValue === 'Closed') {
    if (!isRowAtoIFilled_(rowPack.values)) {
      sheet.getRange(row, LEAD_FLOW.STATUS_COL).setValue(oldValue || 'Follow up');
      return;
    }
    copyRowToHistoryIfMissing_(historySheet, rowPack);
    return;
  }

  if (newValue === 'Follow up' && oldValue === 'Closed') {
    removeRowsByPhone_(historySheet, rowPack.phone);
    return;
  }

  if (newValue !== 'Follow up') {
    sheet.getRange(row, LEAD_FLOW.STATUS_COL).setValue('Follow up');
  }
}

function getRowPack_(sheet, row) {
  const lastCol = sheet.getLastColumn();
  const range = sheet.getRange(row, 1, 1, lastCol);
  const values = range.getValues()[0];
  const formulas = range.getFormulas()[0];
  const phone = normalizePhone_(values[LEAD_FLOW.PHONE_COL - 1]);

  return {
    values: values,
    formulas: formulas,
    totalCols: lastCol,
    phone: phone
  };
}

function copyRowToHistoryIfMissing_(historySheet, rowPack) {
  const exists = hasPhoneInSheet_(historySheet, rowPack.phone);
  if (exists) return;

  const targetRow = Math.max(historySheet.getLastRow() + 1, LEAD_FLOW.START_ROW);
  historySheet.getRange(targetRow, 1, 1, rowPack.totalCols).setValues([rowPack.values]);

  for (let i = 0; i < rowPack.formulas.length; i++) {
    if (rowPack.formulas[i]) {
      historySheet.getRange(targetRow, i + 1).setFormula(rowPack.formulas[i]);
    }
  }

  historySheet.getRange(targetRow, 1, 1, rowPack.totalCols).clearDataValidations();
}

function hasPhoneInSheet_(sheet, phone) {
  const lastRow = sheet.getLastRow();
  if (lastRow < LEAD_FLOW.START_ROW) return false;

  const phones = sheet
    .getRange(LEAD_FLOW.START_ROW, LEAD_FLOW.PHONE_COL, lastRow - LEAD_FLOW.START_ROW + 1, 1)
    .getValues();

  for (let i = 0; i < phones.length; i++) {
    if (normalizePhone_(phones[i][0]) === phone) {
      return true;
    }
  }
  return false;
}

function removeRowsByPhone_(sheet, phone) {
  const lastRow = sheet.getLastRow();
  if (lastRow < LEAD_FLOW.START_ROW) return;

  const rows = sheet
    .getRange(LEAD_FLOW.START_ROW, LEAD_FLOW.PHONE_COL, lastRow - LEAD_FLOW.START_ROW + 1, 1)
    .getValues();

  for (let i = rows.length - 1; i >= 0; i--) {
    if (normalizePhone_(rows[i][0]) === phone) {
      sheet.deleteRow(i + LEAD_FLOW.START_ROW);
    }
  }
}

function isRowAtoIFilled_(rowValues) {
  for (let i = 0; i < 9; i++) {
    if (String(rowValues[i] === null || rowValues[i] === undefined ? '' : rowValues[i]).trim() === '') {
      return false;
    }
  }
  return true;
}

function normalizePhone_(value) {
  const digits = String(value === null || value === undefined ? '' : value).replace(/\D/g, '');
  if (!digits) return '';
  return digits.length > 10 ? digits.slice(-10) : digits;
}