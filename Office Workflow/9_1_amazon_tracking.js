function Order_Status_Amazon() {

  const STATUS_FOLDER_ID = "1lYAVF9zHdl2ZdV6xSZfltwrA3YRYYlnd";
  const RETURN_FOLDER_ID = "1DsLkBqTrU7Z94NexVDdlFk5koJZxmUlC";
  const ui = SpreadsheetApp.getUi();
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  function showAutoCloseDialog(title, lines, autoCloseSeconds) {
    const safeLines = lines
      .map(function (line) {
        return String(line)
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;");
      })
      .join("<br>");

    const html = HtmlService.createHtmlOutput(
      '<div style="font-family:Arial,sans-serif;font-size:13px;line-height:1.6;padding:8px;">' +
      safeLines +
      '</div>' +
      '<script>setTimeout(function(){google.script.host.close();}, ' + (Math.max(2, autoCloseSeconds) * 1000) + ');</script>'
    )
      .setWidth(430)
      .setHeight(260);

    ui.showModelessDialog(html, title);
  }

  function listFolderFiles(folder) {
    const out = [];
    if (!folder) return out;

    const files = folder.getFiles();
    while (files.hasNext()) out.push(files.next());
    return out;
  }

  function parseSeparatedText(text) {
    const firstLine = String(text || "").split(/\r?\n/, 1)[0] || "";
    const delimiterCounts = {
      "\t": (firstLine.match(/\t/g) || []).length,
      ",": (firstLine.match(/,/g) || []).length,
      ";": (firstLine.match(/;/g) || []).length,
      "|": (firstLine.match(/\|/g) || []).length
    };
    const delimiter = Object.keys(delimiterCounts).sort(function (a, b) {
      return delimiterCounts[b] - delimiterCounts[a];
    })[0] || "\t";

    return Utilities.parseCsv(String(text || ""), delimiter);
  }

  function readRowsFromFile(file) {
    try {
      if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
        const ss = SpreadsheetApp.openById(file.getId());
        const sh = ss.getSheets()[0];
        if (!sh) return [];
        return sh.getDataRange().getDisplayValues();
      }

      return parseSeparatedText(file.getBlob().getDataAsString());
    } catch (e) {
      Logger.log("Warning: failed to read file " + file.getName() + " | " + e.message);
      return null;
    }
  }

  function normalizeHeaderName(value) {
    return String(value || "")
      .replace(/^\uFEFF/, "")
      .replace(/^"|"$/g, "")
      .trim()
      .toLowerCase()
      .replace(/[\s_]+/g, "-");
  }

  function normalizeStatus(rawStatus) {
    if (!rawStatus) return "";

    const s = String(rawStatus).trim();
    const STATUS_MAP = {
      "Pending - Waiting for Pick Up": "Waiting for Pick Up",
      "Shipped - Delivered to Buyer": "Delivered",
      "Shipped - Out for Delivery": "Out for Delivery",
      "Shipped - Picked Up": "Picked Up",
      "Shipped - Returned to Seller": "Returned to Seller",
      "Shipped - Returning to Seller": "Returning to Seller",
      "Cancelled": "Cancelled",
      "Shipped": "Shipped",
      "Pending": "Pending"
    };

    return STATUS_MAP[s] || s;
  }

  const configSheet = activeSpreadsheet.getSheetByName("Configuration");
  if (!configSheet) {
    showAutoCloseDialog("Amazon Tracking Update", ["Configuration sheet not found."], 8);
    return;
  }

  const configLastRow = configSheet.getLastRow();
  const sheetIds = configLastRow < 2
    ? []
    : configSheet
      .getRange("F2:F" + configLastRow)
      .getValues()
      .flat()
      .map(id => id.toString().trim())
      .filter(id => id);

  if (sheetIds.length === 0) {
    showAutoCloseDialog("Amazon Tracking Update", ["No sheet found in column F of Configuration."], 8);
    return;
  }

  let statusFolder;
  try {
    statusFolder = DriveApp.getFolderById(STATUS_FOLDER_ID);
  } catch (e) {
    Logger.log("Warning: Amazon Status Folder Not Found: " + STATUS_FOLDER_ID);
    statusFolder = null;
  }

  let returnFolder;
  try {
    returnFolder = DriveApp.getFolderById(RETURN_FOLDER_ID);
  } catch (e) {
    Logger.log("Warning: Amazon Return Folder Not Found: " + RETURN_FOLDER_ID);
    returnFolder = null;
  }

  const statusFilesAll = listFolderFiles(statusFolder);
  const returnFilesAll = listFolderFiles(returnFolder);

  const statusFileStats = {
    total: statusFilesAll.length,
    parsed: 0,
    headerMatched: 0,
    headerMissing: 0,
    errors: 0
  };

  const returnFileStats = {
    total: returnFilesAll.length,
    parsed: 0,
    errors: 0
  };

  const statusMap = {};
  const returnMap = {};
  const missingStatusHeaderFiles = {};

  for (let f = 0; f < statusFilesAll.length; f++) {
    const file = statusFilesAll[f];
    const rows = readRowsFromFile(file);
    if (rows === null) {
      statusFileStats.errors++;
      continue;
    }

    if (!rows || rows.length === 0) continue;
    statusFileStats.parsed++;

    const headerCols = (rows[0] || []).map(normalizeHeaderName);
    let shipmentStatusCol = headerCols.findIndex(function (h) {
      return h === "shipment-status" || h === "shipmentstatus" || h === "status" || h === "order-status";
    });

    let startRow = 1;
    if (shipmentStatusCol === -1) {
      if ((rows[0] || []).length >= 5) {
        const firstOrder = rows[0][0] ? String(rows[0][0]).trim() : "";
        const firstStatus = rows[0][4] ? String(rows[0][4]).trim() : "";
        const looksLikeHeader = /order|shipment|status/i.test(firstOrder) || /status/i.test(firstStatus);

        if (!looksLikeHeader && firstOrder && firstStatus) {
          shipmentStatusCol = 4;
          startRow = 0;
        }
      }
    }

    if (shipmentStatusCol === -1) {
      statusFileStats.headerMissing++;
      missingStatusHeaderFiles[file.getName()] = true;
      Logger.log("Warning: shipment-status column not found in file: " + file.getName());
      continue;
    }

    statusFileStats.headerMatched++;

    for (let i = startRow; i < rows.length; i++) {
      const cols = rows[i];
      if (!cols || cols.length <= shipmentStatusCol) continue;

      const orderId = cols[0] ? String(cols[0]).trim() : "";
      const rawStatus = cols[shipmentStatusCol] ? String(cols[shipmentStatusCol]).trim() : "";

      if (!orderId || !rawStatus) continue;
      statusMap[orderId] = normalizeStatus(rawStatus);
    }
  }

  for (let f = 0; f < returnFilesAll.length; f++) {
    const file = returnFilesAll[f];
    const rows = readRowsFromFile(file);
    if (rows === null) {
      returnFileStats.errors++;
      continue;
    }

    if (!rows || rows.length === 0) continue;
    returnFileStats.parsed++;

    const firstCell = rows[0][0] ? String(rows[0][0]).trim().toLowerCase() : "";
    const startRow = /order/.test(firstCell) ? 1 : 0;

    for (let i = startRow; i < rows.length; i++) {
      const cols = rows[i];
      if (!cols || cols.length < 24) continue;

      const orderId = cols[0] ? String(cols[0]).trim() : "";
      const deliveryDate = cols[23] ? String(cols[23]).trim() : "";

      if (!orderId) continue;
      returnMap[orderId] = deliveryDate
        ? "RETURNED RECEIVED"
        : "Returned But Not Received";
    }
  }

  showAutoCloseDialog(
    "Amazon Tracking Update",
    [
      "Sheet found in Configuration column F: " + sheetIds.length,
      !statusFolder
        ? "Status folder not found"
        : "Status folder file found: " + statusFileStats.total + " (parsed: " + statusFileStats.parsed + ")",
      !returnFolder
        ? "Return folder not found"
        : "Return folder file found: " + returnFileStats.total + " (parsed: " + returnFileStats.parsed + ")"
    ],
    8
  );

  let totalStatusUpdatedRows = 0;
  let totalReturnUpdatedRows = 0;
  let processedSheetCount = 0;
  let failedSheetCount = 0;
  for (const targetSheetId of sheetIds) {
    try {
      const sheet = SpreadsheetApp.openById(targetSheetId.toString().trim()).getActiveSheet();
      const START_ROW = 3;
      const lastRow = sheet.getLastRow();
      if (lastRow < START_ROW) continue;

      const numRows = lastRow - START_ROW + 1;

      const orderIds = sheet.getRange(START_ROW, 1, numRows, 1).getValues().flat();
      const currentStatuses = sheet.getRange(START_ROW, 28, numRows, 1).getValues().flat();
      const statusUpdatedRows = {};
      const returnUpdatedRows = {};

      const orderRowMap = {};
      orderIds.forEach((id, idx) => {
        const key = id ? id.toString().trim() : "";
        if (key) orderRowMap[key] = idx;
      });

      const statusOut = [...currentStatuses];
      const sourceOut = Array(numRows).fill("");

      activeSpreadsheet.toast("Updating delivery status", "Amazon Tracking Update", 3);
      Logger.log("Updating delivery status for sheet: " + targetSheetId);

      const orderIdKeys = Object.keys(orderRowMap);
      for (let i = 0; i < orderIdKeys.length; i++) {
        const orderId = orderIdKeys[i];
        const newStatus = statusMap[orderId];
        if (!newStatus) continue;

        const rowIdx = orderRowMap[orderId];
        statusOut[rowIdx] = newStatus;
        sourceOut[rowIdx] = "updated from amazon order status";
        statusUpdatedRows[rowIdx] = true;
      }

      activeSpreadsheet.toast("Updating return status", "Amazon Tracking Update", 3);
      Logger.log("Updating return status for sheet: " + targetSheetId);

      for (let i = 0; i < orderIdKeys.length; i++) {
        const orderId = orderIdKeys[i];
        const returnStatus = returnMap[orderId];
        if (!returnStatus) continue;

        const rowIdx = orderRowMap[orderId];
        statusOut[rowIdx] = returnStatus;
        sourceOut[rowIdx] = "updated from return files";
        returnUpdatedRows[rowIdx] = true;
      }

      sheet.getRange(START_ROW, 28, numRows, 1).setValues(statusOut.map(v => [v]));
      sheet.getRange(START_ROW, 29, numRows, 1).setValues(sourceOut.map(v => [v]));

      totalStatusUpdatedRows += Object.keys(statusUpdatedRows).length;
      totalReturnUpdatedRows += Object.keys(returnUpdatedRows).length;
      processedSheetCount++;

      Logger.log("Amazon sync completed for sheet: " + targetSheetId);

    } catch (e) {
      failedSheetCount++;
      Logger.log("Failed for sheet ID: " + targetSheetId + " | " + e.message);
    }
  }

  showAutoCloseDialog(
    "Amazon Tracking Update Summary",
    [
      "Sheet found in Configuration column F: " + sheetIds.length,
      !statusFolder
        ? "Status folder not found"
        : "Status folder file found: " + statusFileStats.total +
          " (parsed: " + statusFileStats.parsed +
          ", matched header: " + statusFileStats.headerMatched +
          ", missing header: " + statusFileStats.headerMissing +
          ", read error: " + statusFileStats.errors + ")",
      !returnFolder
        ? "Return folder not found"
        : "Return folder file found: " + returnFileStats.total +
          " (parsed: " + returnFileStats.parsed +
          ", read error: " + returnFileStats.errors + ")",
      "Updated from amazon order status: " + totalStatusUpdatedRows + " row",
      "Updated from return files: " + totalReturnUpdatedRows + " row",
      "Status files missing shipment-status header: " + Object.keys(missingStatusHeaderFiles).length,
      "Sheets processed: " + processedSheetCount,
      "Sheets failed: " + failedSheetCount
    ],
    15
  );

  Logger.log("Amazon status then return sync completed for all sheets");
}