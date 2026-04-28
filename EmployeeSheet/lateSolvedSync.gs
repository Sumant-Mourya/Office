/**
 * Google Apps Script: Pull delayed orders into target sheet every day at 8 AM.
 *
 * How to use:
 * 1) Create a standalone or bound Apps Script project.
 * 2) Paste this file.
 * 3) If target is in another spreadsheet, set TARGET_SPREADSHEET_ID.
 * 4) Run setupMorningTrigger() once (authorize when asked).
 * 5) Optional: run syncLateSolvedOrders() manually to test.
 */

const CONFIG = {
  TARGET_SHEET_NAME: '1.LateSolved',
  // Keep null when target sheet is in the active spreadsheet (bound script).
  TARGET_SPREADSHEET_ID: null,

  PARENT_SPREADSHEET_IDS: [
    '1YiY43sxWZu-WtCQlU9dqNhN7EVE6umpqZMu7mBJcX-U',
    '1utDGphOGzMxHmQq7YcaILQlbzix9JrsXcVll4p0nnlc'
  ],

  SOURCE_SHEET_NAME: 'All Orders',
  HEADER_ROW: 2,
  DATA_START_ROW: 3,

  // You mentioned ">50 days" in your rule, so this is set to 50.
  AGE_DAYS_THRESHOLD: 50,

  // Copy these headers from source to same header names in target.
  COPY_HEADERS: [
    'INTERNAL_ORDERNO',
    'DELIVERY_PICKUP_DATE',
    'PURCHASE_DATE',
    'FULLNAME',
    'ITEM_NAME',
    'PAYMENT_STATUS'
  ],

  PURCHASE_DATE_HEADER: 'PURCHASE_DATE'
};

/**
 * Main sync function.
 */
function syncLateSolvedOrders() {
  const now = new Date();
  const targetSheet = getTargetSheet_();

  const targetHeaderMap = getHeaderMap_(targetSheet, CONFIG.HEADER_ROW);
  validateRequiredHeaders_(targetHeaderMap, CONFIG.COPY_HEADERS, {
    scope: 'TARGET',
    spreadsheetId: targetSheet.getParent().getId(),
    spreadsheetName: targetSheet.getParent().getName(),
    sheetName: targetSheet.getName(),
    headerRow: CONFIG.HEADER_ROW
  });
  const removedOrderNos = removeGreenRowsFromTarget_(targetSheet, targetHeaderMap);

  const collectedRows = [];

  CONFIG.PARENT_SPREADSHEET_IDS.forEach((spreadsheetId) => {
    const sourceSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sourceSheet = sourceSpreadsheet.getSheetByName(CONFIG.SOURCE_SHEET_NAME);

    if (!sourceSheet) {
      Logger.log('Skipped %s: sheet "%s" not found.', spreadsheetId, CONFIG.SOURCE_SHEET_NAME);
      return;
    }

    const sourceHeaderMap = getHeaderMap_(sourceSheet, CONFIG.HEADER_ROW);
    validateRequiredHeaders_(sourceHeaderMap, CONFIG.COPY_HEADERS, {
      scope: 'SOURCE',
      spreadsheetId: sourceSpreadsheet.getId(),
      spreadsheetName: sourceSpreadsheet.getName(),
      sheetName: sourceSheet.getName(),
      headerRow: CONFIG.HEADER_ROW
    });
    validateRequiredHeaders_(sourceHeaderMap, [CONFIG.PURCHASE_DATE_HEADER], {
      scope: 'SOURCE',
      spreadsheetId: sourceSpreadsheet.getId(),
      spreadsheetName: sourceSpreadsheet.getName(),
      sheetName: sourceSheet.getName(),
      headerRow: CONFIG.HEADER_ROW
    });

    const lastRow = sourceSheet.getLastRow();
    const lastCol = sourceSheet.getLastColumn();

    if (lastRow < CONFIG.DATA_START_ROW) {
      return;
    }

    const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
    const dataRange = sourceSheet.getRange(CONFIG.DATA_START_ROW, 1, numRows, lastCol);
    const values = dataRange.getValues();
    const backgrounds = dataRange.getBackgrounds();

    const purchaseDateCol0 = sourceHeaderMap[CONFIG.PURCHASE_DATE_HEADER] - 1;
    const internalOrderNoCol0 = sourceHeaderMap.INTERNAL_ORDERNO - 1;

    values.forEach((row, rowIndex) => {
      const orderNoKey = normalizeKey_(row[internalOrderNoCol0]);
      if (orderNoKey && removedOrderNos.has(orderNoKey)) return;

      const purchaseDateValue = row[purchaseDateCol0];
      const purchaseDate = normalizeDate_(purchaseDateValue);
      if (!purchaseDate) return;

      const ageInDays = Math.floor((now.getTime() - purchaseDate.getTime()) / 86400000);
      if (ageInDays <= CONFIG.AGE_DAYS_THRESHOLD) return;

      // Skip if PURCHASE_DATE cell has any non-white fill color.
      const bgColor = backgrounds[rowIndex][purchaseDateCol0];
      if (!isWhiteColor_(bgColor)) return;

      const outputRow = new Array(CONFIG.COPY_HEADERS.length).fill('');
      CONFIG.COPY_HEADERS.forEach((header, i) => {
        const sourceCol0 = sourceHeaderMap[header] - 1;
        outputRow[i] = row[sourceCol0];
      });

      collectedRows.push(outputRow);
    });
  });

  writeToTarget_(targetSheet, targetHeaderMap, collectedRows);
  Logger.log('Sync complete. Rows written: %s', collectedRows.length);
}

function removeGreenRowsFromTarget_(targetSheet, targetHeaderMap) {
  const dataStart = CONFIG.DATA_START_ROW;
  const lastRow = targetSheet.getLastRow();
  const lastCol = targetSheet.getLastColumn();
  const removedOrderNos = new Set();

  if (lastRow < dataStart || lastCol === 0) {
    return removedOrderNos;
  }

  const numRows = lastRow - dataStart + 1;
  const range = targetSheet.getRange(dataStart, 1, numRows, lastCol);
  const values = range.getValues();
  const backgrounds = range.getBackgrounds();
  const rowsToDelete = [];

  const internalOrderNoCol0 = targetHeaderMap.INTERNAL_ORDERNO
    ? targetHeaderMap.INTERNAL_ORDERNO - 1
    : -1;

  values.forEach((row, rowIndex) => {
    const hasAnyValue = row.some((cell) => String(cell).trim() !== '');
    if (!hasAnyValue) return;

    const hasGreenCell = backgrounds[rowIndex].some((color) => isGreenColor_(color));
    if (!hasGreenCell) return;

    if (internalOrderNoCol0 >= 0) {
      const orderNoKey = normalizeKey_(row[internalOrderNoCol0]);
      if (orderNoKey) removedOrderNos.add(orderNoKey);
    }

    rowsToDelete.push(dataStart + rowIndex);
  });

  // Delete from bottom to top so row indexes stay valid.
  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    targetSheet.deleteRow(rowsToDelete[i]);
  }

  return removedOrderNos;
}

function getTargetSheet_() {
  const ss = CONFIG.TARGET_SPREADSHEET_ID
    ? SpreadsheetApp.openById(CONFIG.TARGET_SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  const sheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  if (!sheet) {
    throw new Error('Target sheet not found: ' + CONFIG.TARGET_SHEET_NAME);
  }
  return sheet;
}

function getHeaderMap_(sheet, headerRow) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    throw new Error('Header row is empty in sheet: ' + sheet.getName());
  }

  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  const map = {};

  headers.forEach((header, index) => {
    const key = normalizeHeaderName_(header);
    if (key) map[key] = index + 1;
  });

  return map;
}

function validateRequiredHeaders_(headerMap, requiredHeaders, context) {
  const missing = requiredHeaders.filter((h) => !headerMap[normalizeHeaderName_(h)]);
  if (missing.length > 0) {
    const foundHeaders = Object.keys(headerMap).sort().join(', ');
    const location = context
      ? [
          'scope=' + context.scope,
          'spreadsheetId=' + context.spreadsheetId,
          'spreadsheetName=' + context.spreadsheetName,
          'sheetName=' + context.sheetName,
          'headerRow=' + context.headerRow
        ].join(' | ')
      : 'scope=UNKNOWN';

    throw new Error(
      'Missing required headers: ' + missing.join(', ') +
      ' || Location: ' + location +
      ' || Found headers (normalized): ' + foundHeaders
    );
  }
}

function normalizeHeaderName_(value) {
  return String(value == null ? '' : value)
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}

function normalizeDate_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  const parsed = new Date(value);
  if (isNaN(parsed)) return null;

  return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
}

function isWhiteColor_(hex) {
  if (!hex) return true;
  const c = String(hex).trim().toLowerCase();
  return c === '#ffffff' || c === '#fff' || c === 'white';
}

function isGreenColor_(hex) {
  if (!hex) return false;

  const c = String(hex).trim().toLowerCase();
  if (c === 'green') return true;

  const commonGreens = new Set([
    '#00ff00',
    '#b6d7a8',
    '#93c47d',
    '#6aa84f',
    '#38761d',
    '#34a853'
  ]);
  if (commonGreens.has(c)) return true;

  const rgb = hexToRgb_(c);
  if (!rgb) return false;

  // Heuristic for green shades.
  return rgb.g > 80 && rgb.g >= rgb.r * 1.2 && rgb.g >= rgb.b * 1.2;
}

function hexToRgb_(hex) {
  const m = /^#?([0-9a-f]{6})$/i.exec(hex);
  if (!m) return null;

  const s = m[1];
  return {
    r: parseInt(s.substring(0, 2), 16),
    g: parseInt(s.substring(2, 4), 16),
    b: parseInt(s.substring(4, 6), 16)
  };
}

function normalizeKey_(value) {
  return String(value == null ? '' : value).trim();
}

function writeToTarget_(targetSheet, targetHeaderMap, collectedRows) {
  const dataStart = CONFIG.DATA_START_ROW;
  const lastRow = targetSheet.getLastRow();
  const lastCol = targetSheet.getLastColumn();

  // Clear previous generated data section (rows 3+), keep header rows 1 and 2.
  if (lastRow >= dataStart && lastCol > 0) {
    targetSheet.getRange(dataStart, 1, lastRow - dataStart + 1, lastCol).clearContent();
  }

  if (collectedRows.length === 0) {
    return;
  }

  // Build a full-width matrix so values land under matching target headers.
  const targetWidth = Math.max(lastCol, Math.max(...Object.values(targetHeaderMap)));
  const output = collectedRows.map((srcRow) => {
    const row = new Array(targetWidth).fill('');
    CONFIG.COPY_HEADERS.forEach((header, i) => {
      const col1 = targetHeaderMap[header];
      if (col1) row[col1 - 1] = srcRow[i];
    });
    return row;
  });

  targetSheet.getRange(dataStart, 1, output.length, targetWidth).setValues(output);
}
