function Order_Status_Flipkart() {

  const STATUS_FOLDER_ID = "1G3utXu1VrZ_ih-mgKLyx4bK8p7KFTzhW";
  const RETURN_FOLDER_ID = "1ufpDlGAW9BeP6TUAL9bd8Nlj9PvFL_Df";

  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration");
  const configLastRow = configSheet.getLastRow();
  if (configLastRow < 1) return;
  const sheetIds = configSheet.getRange("F1:F" + configLastRow).getValues().flat().filter(id => id.toString().trim());

  for (const targetSheetId of sheetIds) {
    try {
      const sheet = SpreadsheetApp.openById(targetSheetId.toString().trim()).getActiveSheet();
      const START_ROW = 3;
      const lastRow = sheet.getLastRow();
      if (lastRow < START_ROW) continue;

  const numRows = lastRow - START_ROW + 1;

  // ================= READ MAIN SHEET =================
  const orderIds = sheet.getRange(START_ROW, 1, numRows, 1).getValues().flat();
  const currentStatuses = sheet.getRange(START_ROW, 28, numRows, 1).getValues().flat();

  function normalizeId(id) {
    return id.toString().replace(/[^0-9]/g, "");
  }

  const orderRowMap = {};
  orderIds.forEach((id, idx) => {
    if (id) orderRowMap[normalizeId(id)] = idx;
  });

  const statusOut = [...currentStatuses];
  const sourceOut = Array(numRows).fill("");

  // ================= HELPER =================
  function openExcelAsSheet(file) {
    try {
      if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
        return SpreadsheetApp.openById(file.getId());
      }
      const converted = Drive.Files.copy(
        { mimeType: MimeType.GOOGLE_SHEETS, title: file.getName() },
        file.getId()
      );
      return SpreadsheetApp.openById(converted.id);
    } catch (e) {
      Logger.log("❌ Conversion failed: " + file.getName());
      return null;
    }
  }

  // =====================================================
  // 1️⃣ RETURN FILES (SAFE)
  // =====================================================
  let returnFolder;
  try {
    returnFolder = DriveApp.getFolderById(RETURN_FOLDER_ID);
  } catch (e) {
    Logger.log("⚠️ Flipkart Return Folder Not Found: " + RETURN_FOLDER_ID);
    returnFolder = null;
  }

  if (returnFolder) {
    const returnFiles = returnFolder.getFiles();

    while (returnFiles.hasNext()) {
      const file = returnFiles.next();
      if (!file.getName().toLowerCase().endsWith(".xlsx")) continue;

      const ss = openExcelAsSheet(file);
      if (!ss) continue;

      const returnsSheet = ss.getSheetByName("Returns");
      if (!returnsSheet) {
        DriveApp.getFileById(ss.getId()).setTrashed(true);
        continue;
      }

      const data = returnsSheet.getDataRange().getValues();

      for (let i = 1; i < data.length; i++) {
        const rawOrderId = data[i][1];    // Column B
        const returnStatus = data[i][16]; // Column Q
        if (!rawOrderId || !returnStatus) continue;

        const normalizedId = normalizeId(rawOrderId);
        if (orderRowMap[normalizedId] === undefined) continue;

        const idx = orderRowMap[normalizedId];
        statusOut[idx] = returnStatus.toString().trim();
        sourceOut[idx] = "updated from return excel";
      }

      DriveApp.getFileById(ss.getId()).setTrashed(true);
    }
  }

  // =====================================================
  // 2️⃣ STATUS FILES (SAFE)
  // =====================================================
  let statusFolder;
  try {
    statusFolder = DriveApp.getFolderById(STATUS_FOLDER_ID);
  } catch (e) {
    Logger.log("⚠️ Flipkart Status Folder Not Found: " + STATUS_FOLDER_ID);
    statusFolder = null;
  }

  if (statusFolder) {
    const statusFiles = statusFolder.getFiles();

    while (statusFiles.hasNext()) {
      const file = statusFiles.next();
      if (!file.getName().toLowerCase().endsWith(".xlsx")) continue;

      const ss = openExcelAsSheet(file);
      if (!ss) continue;

      const ordersSheet = ss.getSheetByName("Orders");
      if (!ordersSheet) {
        DriveApp.getFileById(ss.getId()).setTrashed(true);
        continue;
      }

      const data = ordersSheet.getDataRange().getValues();

      for (let i = 1; i < data.length; i++) {
        const orderId = data[i][1];   // Column B
        const rawStatus = data[i][6]; // Column G
        if (!orderId || !rawStatus) continue;

        const normalizedId = normalizeId(orderId);
        if (orderRowMap[normalizedId] === undefined) continue;

        const idx = orderRowMap[normalizedId];

        // skip if already updated by return logic
        if (sourceOut[idx]) continue;

        statusOut[idx] = rawStatus.toString().trim();
        sourceOut[idx] = "updated from status excel";
      }

      DriveApp.getFileById(ss.getId()).setTrashed(true);
    }
  }

  // ================= WRITE BACK =================
  sheet.getRange(START_ROW, 28, numRows, 1).setValues(statusOut.map(v => [v]));
  sheet.getRange(START_ROW, 29, numRows, 1).setValues(sourceOut.map(v => [v]));

  Logger.log("✅ Flipkart sync completed for sheet: " + targetSheetId);

    } catch (e) {
      Logger.log("❌ Failed for sheet ID: " + targetSheetId + " | " + e.message);
    }
  }

  Logger.log("✅ Flipkart status + return sync completed for all sheets");
}