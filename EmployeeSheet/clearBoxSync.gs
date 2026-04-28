/**
 * Sync returned/cancel orders into target sheet 2.ClearBox.
 * Header row is 2, data starts at row 3.
 */
function syncClearBoxOrders() {
  const targetSheet = getClearBoxTargetSheet_();
  const targetHeaderMap = getHeaderMap_(targetSheet, CONFIG.HEADER_ROW);

  validateRequiredHeaders_(targetHeaderMap, ['INTERNAL_ORDERNO'], {
    scope: 'TARGET_CLEARBOX',
    spreadsheetId: targetSheet.getParent().getId(),
    spreadsheetName: targetSheet.getParent().getName(),
    sheetName: targetSheet.getName(),
    headerRow: CONFIG.HEADER_ROW
  });

  const rowsToWrite = [];

  CONFIG.PARENT_SPREADSHEET_IDS.forEach((spreadsheetId) => {
    const sourceSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sourceSheet = sourceSpreadsheet.getSheetByName(CONFIG.SOURCE_SHEET_NAME);

    if (!sourceSheet) {
      Logger.log('ClearBox skipped %s: sheet "%s" not found.', spreadsheetId, CONFIG.SOURCE_SHEET_NAME);
      return;
    }

    const sourceHeaderMap = getHeaderMap_(sourceSheet, CONFIG.HEADER_ROW);
    validateRequiredHeaders_(sourceHeaderMap, ['INTERNAL_ORDERNO', 'STATUS'], {
      scope: 'SOURCE_CLEARBOX',
      spreadsheetId: sourceSpreadsheet.getId(),
      spreadsheetName: sourceSpreadsheet.getName(),
      sheetName: sourceSheet.getName(),
      headerRow: CONFIG.HEADER_ROW
    });

    const lastRow = sourceSheet.getLastRow();
    const lastCol = sourceSheet.getLastColumn();
    if (lastRow < CONFIG.DATA_START_ROW) return;

    const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
    const values = sourceSheet.getRange(CONFIG.DATA_START_ROW, 1, numRows, lastCol).getValues();

    const internalOrderNoCol0 = sourceHeaderMap.INTERNAL_ORDERNO - 1;
    const statusCol0 = sourceHeaderMap.STATUS - 1;

    values.forEach((row) => {
      const orderNo = row[internalOrderNoCol0];
      const status = normalizeKey_(row[statusCol0]).toUpperCase();
      if (!status) return;

      const mappedValue = mapStatusForClearBox_(status);
      if (!mappedValue) return;

      rowsToWrite.push({
        internalOrderNo: orderNo,
        columnCValue: mappedValue
      });
    });
  });

  writeClearBoxTarget_(targetSheet, targetHeaderMap, rowsToWrite);
  Logger.log('ClearBox sync complete. Rows written: %s', rowsToWrite.length);
}

function getClearBoxTargetSheet_() {
  const ss = CONFIG.TARGET_SPREADSHEET_ID
    ? SpreadsheetApp.openById(CONFIG.TARGET_SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  const sheet = ss.getSheetByName('2.ClearBox');
  if (!sheet) {
    throw new Error('Target sheet not found: 2.ClearBox');
  }

  return sheet;
}

function mapStatusForClearBox_(statusText) {
  if (statusText.indexOf('RETURN') !== -1) return 'RTO';
  if (statusText.indexOf('CANCEL') !== -1) return 'CANCEL';
  return '';
}

function writeClearBoxTarget_(targetSheet, targetHeaderMap, rowsToWrite) {
  const dataStart = CONFIG.DATA_START_ROW;
  const lastRow = targetSheet.getLastRow();
  const lastCol = Math.max(targetSheet.getLastColumn(), 3);

  if (lastRow >= dataStart && lastCol > 0) {
    targetSheet.getRange(dataStart, 1, lastRow - dataStart + 1, lastCol).clearContent();
  }

  if (rowsToWrite.length === 0) return;

  const internalOrderNoCol1 = targetHeaderMap.INTERNAL_ORDERNO;
  const output = rowsToWrite.map((item) => {
    const row = new Array(lastCol).fill('');
    row[internalOrderNoCol1 - 1] = item.internalOrderNo;
    row[2] = item.columnCValue;
    return row;
  });

  targetSheet.getRange(dataStart, 1, output.length, lastCol).setValues(output);
}
