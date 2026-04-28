const STATUS_MOVER_PARENT_FOLDER_ID = '1699pqPDxIXUNjJ9H7jf0tru6oLugKUPf';
const STATUS_MOVER_TRIGGER_HANDLER = 'runOrderStatusMover';
const STATUS_MOVER_CONFIG_SHEET = 'Configuration';
const STATUS_MOVER_MAIN_ID_CELL = 'B1';
const STATUS_MOVER_MAIN_TAB_CELL = 'B2';
const ORDER_DATE_COLUMN_INDEX = 8; // Column H (1-based)
const STATUS_COLUMN_INDEX = 28; // Column AB (1-based)
const STATUS_MOVER_LOG_SHEET_NAME = 'log';
const ENABLE_STATUS_MOVER_LOGGING = true;
const YEAR_FOLDER_CACHE_ = {};
const MONTH_SPREADSHEET_CACHE_ = {};

const STATUS_DATASET = [
  'canceled',
  'cancelled',
  'cancel',
  'return',
  'returned',
  'rto',
  'deliver',
  'delivered',
  'complete',
  'completed',
  'Alert'
];

const TARGET_HEADERS = [
  'PortalOrderID',
  'Delivery or Pickup date',
  'Delivery Status',
  'Payment Status',
  'Staff Notes',
  'Brand Name',
  'Internal OrderNo',
  'Purchase Date',
  'Sales Channel',
  'SKU',
  'Item Name',
  'Cateogry',
  'Image',
  'OrderNote',
  'Qty',
  'Currency',
  'Currency Price',
  'Fullname',
  'AddressLine1',
  'AddressLine2',
  'City',
  'State',
  'Pincode',
  'Country',
  'Phone',
  'Courier Name',
  'Tracking Code',
  'Status',
  'Shipping Charge',
  'Image url',
  'Listing Url',
  'New Tracking',
  'INR',
  'Price in INR',
  'Product Cost',
  'Ecommerce Expense+Taxes',
  'Estimated Profit',
  'Return Request Date',
  'Return Courier Name',
  'Return Tracking Code',
  'RTO Recieved Date',
  'Wrong Return Claims',
  'Shipping Charge'
];

function showTriggerControlDialog() {
  const html = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <style>
          body {
            font-family: Arial;
            text-align: center;
            padding: 20px;
          }
          button {
            padding: 10px 20px;
            font-size: 14px;
            margin-top: 15px;
            cursor: pointer;
          }
          .status {
            margin-top: 10px;
            font-weight: bold;
          }
        </style>
      </head>
      <body>

        <h3>Order Status Mover</h3>

        <div class="status" id="status">Checking...</div>

        <button onclick="toggle()">Start / Stop</button>

        <script>
          function loadStatus() {
            google.script.run.withSuccessHandler(function(res) {
              document.getElementById('status').innerText =
                res ? "🟢 Running" : "🔴 Stopped";
            }).isTriggerRunning();
          }

          function toggle() {
            google.script.run.withSuccessHandler(function(res) {
              if (res.action === 'started') {
                document.getElementById('status').innerText = "🟢 Running";
              } else {
                document.getElementById('status').innerText = "🔴 Stopped";
              }
            }).toggleOrderStatusMoverTrigger();
          }

          loadStatus();
        </script>

      </body>
    </html>
  `)
  .setWidth(300)
  .setHeight(200);

  SpreadsheetApp.getUi().showModalDialog(html, 'Trigger Control');
}

function isTriggerRunning() {
  const triggers = ScriptApp.getProjectTriggers();
  return triggers.some(t => t.getHandlerFunction() === STATUS_MOVER_TRIGGER_HANDLER);
}

function toggleOrderStatusMoverTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  const existing = triggers.filter(function(t) {
    return t.getHandlerFunction() === STATUS_MOVER_TRIGGER_HANDLER;
  });

  if (existing.length > 0) {
    existing.forEach(function(t) {
      ScriptApp.deleteTrigger(t);
    });
    clearRunLock_();
    logStatusMover_('INFO', 'Trigger stopped', { deletedTriggers: existing.length });
    return {
      action: 'stopped',
      deletedTriggers: existing.length
    };
  }

  createStatusMoverTrigger_();

  return {
    action: 'started',
    mode: 'every_minute_test'
  };
}

function createStatusMoverTrigger_() {
  ScriptApp.newTrigger(STATUS_MOVER_TRIGGER_HANDLER)
    .timeBased()
    .everyMinutes(1)
    // .everyDays(1)
    // .atHour(8)
    .create();

  logStatusMover_('INFO', 'Every 1 minute trigger created (test mode)');
}

function runOrderStatusMover() {
  const executionLock = LockService.getScriptLock();
  if (!executionLock.tryLock(1000)) {
    logStatusMover_('WARN', 'Skipping run: previous execution still in progress', {
      lockAge: getRunLockAgeMinutes_()
    });
    return { skipped: true, reason: 'concurrent_run_lock' };
  }

  setRunLock_();
  const startedAt = new Date();
  logStatusMover_('INFO', 'Run started');

  try {
    const config = getStatusMoverConfig_();
    logStatusMover_('INFO', 'Configuration loaded', {
      mainSpreadsheetId: config.mainSpreadsheetId,
      mainOrdersSheetName: config.mainOrdersSheetName
    });

    const mainSpreadsheet = SpreadsheetApp.openById(config.mainSpreadsheetId);
    const mainSheet = mainSpreadsheet.getSheetByName(config.mainOrdersSheetName);
    if (!mainSheet) {
      throw new Error('Main orders sheet not found: ' + config.mainOrdersSheetName);
    }

    // Queue logic: always check row 3 (after 2 header rows). Move row 3 only if AB has
    // a configured keyword. Stop immediately when row 3 does not match.
    let moved = 0;
    const movedByTarget = {};
    let stopReason = 'row3_status_not_match';
    let lastTargetSpreadsheetId = '';
    let lastTargetSpreadsheetName = '';

    const targetSheetCache = {};
    const targetDimensionsSynced = {};

    while (true) {
      const lastRow = runWithRetry_(function() {
        return mainSheet.getLastRow();
      }, 'row3Queue:getLastRow');

      if (lastRow < 3) {
        stopReason = 'no_data_row_after_headers';
        break;
      }

      const row3Range = mainSheet.getRange(3, 1, 1, TARGET_HEADERS.length);
      const row3Values = runWithRetry_(function() {
        return row3Range.getValues()[0];
      }, 'row3Queue:getRow3Values');

      const matchedStatusKeyword = findMatchedStatusKeyword_(row3Values[STATUS_COLUMN_INDEX - 1]);
      if (!matchedStatusKeyword) {
        stopReason = 'row3_status_not_match';
        break;
      }

      const orderDate = parseOrderDate_(row3Values[ORDER_DATE_COLUMN_INDEX - 1]);
      if (!orderDate) {
        logStatusMover_('WARN', 'Date format not matched', {
          row: 3,
          column: 'H',
          value: String(row3Values[ORDER_DATE_COLUMN_INDEX - 1] == null ? '' : row3Values[ORDER_DATE_COLUMN_INDEX - 1])
        });
        stopReason = 'row3_invalid_order_date';
        break;
      }

      const targetInfo = getTargetSpreadsheetByOrderDate_(orderDate);
      const cacheKey = targetInfo.spreadsheet.getId() + '::' + targetInfo.monthlyName;
      if (!targetSheetCache[cacheKey]) {
        targetSheetCache[cacheKey] = ensureTargetSheet_(targetInfo.spreadsheet, targetInfo.monthlyName);
      }
      if (!targetDimensionsSynced[cacheKey]) {
        syncColumnWidths_(mainSheet, targetSheetCache[cacheKey], TARGET_HEADERS.length);
        targetDimensionsSynced[cacheKey] = true;
      }

      appendCopiedRowFromSource_(mainSheet, 3, targetSheetCache[cacheKey], TARGET_HEADERS.length);

      runWithRetry_(function() {
        mainSheet.deleteRow(3);
        return true;
      }, 'row3Queue:deleteRow3');

      moved++;
      movedByTarget[targetInfo.monthlyName] = (movedByTarget[targetInfo.monthlyName] || 0) + 1;
      lastTargetSpreadsheetId = targetInfo.spreadsheet.getId();
      lastTargetSpreadsheetName = targetInfo.spreadsheet.getName();
    }

    const result = {
      moved: moved,
      movedByTarget: movedByTarget,
      stopReason: stopReason,
      lastTargetSpreadsheetId: lastTargetSpreadsheetId,
      lastTargetSpreadsheetName: lastTargetSpreadsheetName,
      durationMs: new Date().getTime() - startedAt.getTime()
    };

    logStatusMover_('INFO', 'Run completed', result);
    clearRunLock_();
    executionLock.releaseLock();

    return result;
  } catch (e) {
    logStatusMover_('ERROR', 'Run failed', {
      message: e && e.message ? e.message : String(e),
      stack: e && e.stack ? e.stack : ''
    });
    clearRunLock_();
    executionLock.releaseLock();
    throw e;
  }
}

// ─── Run lock helpers (stale after 15 minutes) ──────────────────────────────
const RUN_LOCK_KEY_ = 'STATUS_MOVER_RUNNING_SINCE';
const RUN_LOCK_STALE_MS_ = 15 * 60 * 1000;

function setRunLock_() {
  PropertiesService.getScriptProperties().setProperty(RUN_LOCK_KEY_, String(new Date().getTime()));
}

function clearRunLock_() {
  PropertiesService.getScriptProperties().deleteProperty(RUN_LOCK_KEY_);
}

function isRunLocked_() {
  return false;
}

function getRunLockAgeMinutes_() {
  const val = PropertiesService.getScriptProperties().getProperty(RUN_LOCK_KEY_);
  if (!val) return 0;
  const lockedAt = parseInt(val, 10);
  if (isNaN(lockedAt)) return 0;
  return Math.round((new Date().getTime() - lockedAt) / 60000);
}

function getStatusMoverConfig_() {
  const localSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = localSpreadsheet.getSheetByName(STATUS_MOVER_CONFIG_SHEET);
  if (!configSheet) {
    throw new Error('Configuration sheet not found in current spreadsheet.');
  }

  const mainSpreadsheetId = String(configSheet.getRange(STATUS_MOVER_MAIN_ID_CELL).getDisplayValue() || '').trim();
  const mainOrdersSheetName = String(configSheet.getRange(STATUS_MOVER_MAIN_TAB_CELL).getDisplayValue() || '').trim();

  if (!mainSpreadsheetId) {
    throw new Error('Configuration B1 is empty. Put main spreadsheet ID in B1.');
  }
  if (!mainOrdersSheetName) {
    throw new Error('Configuration B2 is empty. Put main orders sheet name in B2.');
  }

  return {
    mainSpreadsheetId: mainSpreadsheetId,
    mainOrdersSheetName: mainOrdersSheetName
  };
}

function getTargetSpreadsheetByOrderDate_(orderDate) {
  const yearName = String(orderDate.getFullYear());
  const monthName = MONTH_NAMES_[orderDate.getMonth()];
  const yearSuffix = String(orderDate.getFullYear()).slice(-2);
  const monthlyName = monthName + yearSuffix; // Example: March26
  const monthCacheKey = yearName + '::' + monthlyName;

  if (MONTH_SPREADSHEET_CACHE_[monthCacheKey]) {
    return MONTH_SPREADSHEET_CACHE_[monthCacheKey];
  }

  const parentFolder = runWithRetry_(function() {
    return DriveApp.getFolderById(STATUS_MOVER_PARENT_FOLDER_ID);
  }, 'getTargetSpreadsheetByOrderDate_:getParentFolder');
  const yearFolderInfo = getOrCreateFolderByName_(parentFolder, yearName);
  const yearFolder = yearFolderInfo.folder;
  const monthlyFiles = runWithRetry_(function() {
    return yearFolder.getFilesByName(monthlyName);
  }, 'getTargetSpreadsheetByOrderDate_:getFilesByName');

  while (monthlyFiles.hasNext()) {
    const file = monthlyFiles.next();
    if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
      const foundSpreadsheet = runWithRetry_(function() {
        return SpreadsheetApp.openById(file.getId());
      }, 'getTargetSpreadsheetByOrderDate_:openExistingById');

      const existingInfo = {
        spreadsheet: foundSpreadsheet,
        yearFolderCreated: yearFolderInfo.created,
        monthSpreadsheetCreated: false,
        monthlyName: monthlyName,
        yearName: yearName
      };
      MONTH_SPREADSHEET_CACHE_[monthCacheKey] = existingInfo;
      return existingInfo;
    }
  }

  const createdFile = runWithRetry_(function() {
    return SpreadsheetApp.create(monthlyName);
  }, 'getTargetSpreadsheetByOrderDate_:createSpreadsheet');
  const createdId = createdFile.getId();
  const createdDriveFile = runWithRetry_(function() {
    return DriveApp.getFileById(createdId);
  }, 'getTargetSpreadsheetByOrderDate_:getCreatedDriveFile');
  runWithRetry_(function() {
    yearFolder.addFile(createdDriveFile);
    return true;
  }, 'getTargetSpreadsheetByOrderDate_:addFileToYearFolder');

  // Remove from root My Drive to keep only inside year folder.
  const root = runWithRetry_(function() {
    return DriveApp.getRootFolder();
  }, 'getTargetSpreadsheetByOrderDate_:getRootFolder');
  runWithRetry_(function() {
    root.removeFile(createdDriveFile);
    return true;
  }, 'getTargetSpreadsheetByOrderDate_:removeFromRoot');

  logStatusMover_('INFO', 'Monthly spreadsheet created', {
    yearFolder: yearName,
    monthlySpreadsheetName: monthlyName,
    monthlySpreadsheetId: createdId
  });

  const createdInfo = {
    spreadsheet: createdFile,
    yearFolderCreated: yearFolderInfo.created,
    monthSpreadsheetCreated: true,
    monthlyName: monthlyName,
    yearName: yearName
  };
  MONTH_SPREADSHEET_CACHE_[monthCacheKey] = createdInfo;
  return createdInfo;
}

function getOrCreateFolderByName_(parentFolder, folderName) {
  const cacheKey = parentFolder.getId() + '::' + folderName;
  if (YEAR_FOLDER_CACHE_[cacheKey]) {
    return { folder: YEAR_FOLDER_CACHE_[cacheKey], created: false };
  }

  const folders = runWithRetry_(function() {
    return parentFolder.getFoldersByName(folderName);
  }, 'getOrCreateFolderByName_:getFoldersByName');
  if (folders.hasNext()) {
    const existing = folders.next();
    YEAR_FOLDER_CACHE_[cacheKey] = existing;
    return { folder: existing, created: false };
  }

  const created = runWithRetry_(function() {
    return parentFolder.createFolder(folderName);
  }, 'getOrCreateFolderByName_:createFolder');
  logStatusMover_('INFO', 'Year folder created', {
    folderName: folderName,
    folderId: created.getId()
  });
  YEAR_FOLDER_CACHE_[cacheKey] = created;
  return { folder: created, created: true };
}

function getStatusMoverLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(STATUS_MOVER_LOG_SHEET_NAME) || ss.getSheetByName('Log');
  if (!sheet) {
    sheet = ss.insertSheet(STATUS_MOVER_LOG_SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 3).setValues([['timestamp', 'level', 'Log']]);
  }

  return sheet;
}

function logStatusMover_(level, message, details) {
  if (!ENABLE_STATUS_MOVER_LOGGING) return;

  const lv = String(level || 'INFO').toUpperCase();
  let text = String(message || '');

  if (details !== undefined && details !== null) {
    if (typeof details === 'string') {
      text += ' | ' + details;
    } else {
      try {
        text += ' | ' + JSON.stringify(details);
      } catch (e) {
        text += ' | [unserializable details]';
      }
    }
  }

  try {
    const logSheet = getStatusMoverLogSheet_();
    logSheet.appendRow([new Date(), lv, text]);
  } catch (e) {
    Logger.log('STATUS_MOVER_LOG_FAILED: ' + e + ' | ' + lv + ': ' + text);
  }
}

function ensureTargetSheet_(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }

  deleteDefaultSheetIfNeeded_(spreadsheet, sheetName);

  const hasHeader = sheet.getLastRow() >= 1;
  if (!hasHeader) {
    sheet.getRange(1, 1, 1, TARGET_HEADERS.length).setValues([TARGET_HEADERS]);
  } else {
    const currentHeader = sheet.getRange(1, 1, 1, TARGET_HEADERS.length).getDisplayValues()[0];
    const headerMismatch = TARGET_HEADERS.some(function(h, i) {
      return String(currentHeader[i] || '').trim() !== h;
    });
    if (headerMismatch) {
      sheet.getRange(1, 1, 1, TARGET_HEADERS.length).setValues([TARGET_HEADERS]);
    }
  }

  // Apply bold + font size 12 to header row.
  const headerRange = sheet.getRange(1, 1, 1, TARGET_HEADERS.length);
  headerRange
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#ffff00');

  return sheet;
}

function appendCopiedRowFromSource_(sourceSheet, sourceRowIndex, targetSheet, width) {
  const sourceRange = sourceSheet.getRange(sourceRowIndex, 1, 1, width);

  // ✅ exact values (keeps date format)
  const rowDisplayValues = sourceRange.getDisplayValues();
  const rowBackgrounds = sourceRange.getBackgrounds();
  const rowFontColors = sourceRange.getFontColors();

  const sourceRowHeight = sourceSheet.getRowHeight(sourceRowIndex);

  // Fix column M
  if (rowDisplayValues[0].length >= 13) {
    rowDisplayValues[0][12] = '';
  }

  // ✅ IMPORTANT: use setValues instead of appendRow
  const targetRow = targetSheet.getLastRow() + 1;
  const targetRange = targetSheet.getRange(targetRow, 1, 1, width);
  
  // force everything as text (IMPORTANT for Column H)
  targetRange.setNumberFormat("@STRING@");
  
  // insert values (keeps dd-MM-yyyy exactly)
  targetRange.setValues(rowDisplayValues);
  
  // styles
  targetRange.setBackgrounds(rowBackgrounds);
  targetRange.setFontColors(rowFontColors);
  
  // align
  targetRange
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  
  // keep ONLY this if needed (safe)
  targetSheet.getRange(targetRow, 2).setNumberFormat("MMM dd, yyyy");
  
  // ❌ REMOVE THIS LINE (important fix)
  // targetSheet.getRange(targetRow, 8).setNumberFormat("dd-MM-yyyy");
  
  // image formula
  targetSheet.getRange(targetRow, 13).setFormula(`=IMAGE(AD${targetRow})`);
  
  // row height
  targetSheet.setRowHeight(targetRow, sourceRowHeight);
}

function deleteDefaultSheetIfNeeded_(spreadsheet, targetSheetName) {
    const defaultSheet = spreadsheet.getSheetByName('Sheet1');
    if (!defaultSheet) return;
    if (defaultSheet.getName() === targetSheetName) return;
    if (spreadsheet.getSheets().length <= 1) return;

    runWithRetry_(function() {
      spreadsheet.deleteSheet(defaultSheet);
      return true;
    }, 'deleteDefaultSheetIfNeeded_:deleteSheet1');
  }

function syncColumnWidths_(sourceSheet, targetSheet, width) {
  for (let c = 1; c <= width; c++) {
    const sourceWidth = sourceSheet.getColumnWidth(c);
    runWithRetry_(function() {
      targetSheet.setColumnWidth(c, sourceWidth);
      return true;
    }, 'syncColumnWidths_:setColumnWidth:' + c);
  }
}

function findMatchedStatusKeyword_(status) {
  const s = String(status == null ? '' : status).trim().toLowerCase();
  if (!s) return '';

  for (let i = 0; i < STATUS_DATASET.length; i++) {
    const kw = String(STATUS_DATASET[i] || '').trim().toLowerCase();
    if (!kw) continue;
    if (s.indexOf(kw) !== -1) {
      return kw;
    }
  }

  return '';
}

const MONTH_NAMES_ = [
  'January',
  'February',
  'March',
  'April',
  'May',
  'June',
  'July',
  'August',
  'September',
  'October',
  'November',
  'December'
];

function parseOrderDate_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return value;
  }

  const text = String(value == null ? '' : value).trim();
  if (!text) return null;

  const parts = text.split('-');
  if (parts.length !== 3) return null;

  const day = parseInt(parts[0], 10);
  const month = parseInt(parts[1], 10);
  const year = parseInt(parts[2], 10);

  if (isNaN(day) || isNaN(month) || isNaN(year)) return null;

  const parsed = new Date(year, month - 1, day);
  if (
    parsed.getFullYear() !== year ||
    parsed.getMonth() !== month - 1 ||
    parsed.getDate() !== day
  ) {
    return null;
  }

  return parsed;
}

function runWithRetry_(fn, label) {
  const maxAttempts = 4;
  let lastError;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      return fn();
    } catch (e) {
      lastError = e;
      const message = e && e.message ? e.message : String(e);
      const isRetryableServiceError =
        message.indexOf('Service error: Spreadsheets') !== -1 ||
        message.indexOf('Service error: Drive') !== -1 ||
        message.indexOf('Service invoked too many times') !== -1;

      logStatusMover_('WARN', 'Retryable operation failed', {
        label: label || 'operation',
        attempt: attempt,
        message: message,
        retrying: isRetryableServiceError && attempt < maxAttempts
      });

      if (!isRetryableServiceError || attempt === maxAttempts) {
        throw e;
      }

      Utilities.sleep(1000 * attempt);
    }
  }

  throw lastError;
}
