function Order_Status_Meesho() {

  const STATUS_FOLDER_ID = "1iIFyKgQQ3dycrG1PxJxCk0TNkB7UAOng";
  const RETURN_FOLDER_ID = "1sutxzHRswErflE0VCUlf0GKPVj4VRG1a";

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

  // ================= MAIN SHEET =================
  const orderIds = sheet.getRange(START_ROW, 1, numRows, 1).getValues().flat();
  const currentStatuses = sheet.getRange(START_ROW, 28, numRows, 1).getValues().flat();

  // 🔥 ONE & ONLY normalization function
  function normalizeOrderId(id) {
    if (!id) return "";
    return id
      .toString()
      .split("_")[0]
      .replace(/[^0-9]/g, "")
      .trim();
  }

  // 🔁 Map baseOrderId → all matching rows
  const orderRowMap = {};
  orderIds.forEach((id, idx) => {
    const key = normalizeOrderId(id);
    if (!key) return;
    if (!orderRowMap[key]) orderRowMap[key] = [];
    orderRowMap[key].push(idx);
  });

  const statusOut = [...currentStatuses];
  const sourceOut = Array(numRows).fill("");

  // =====================================================
  // 1️⃣ RETURN FILES (SAFE)
  // =====================================================
  let returnFolder;
  try {
    returnFolder = DriveApp.getFolderById(RETURN_FOLDER_ID);
  } catch (e) {
    Logger.log("⚠️ Meesho Return Folder Not Found: " + RETURN_FOLDER_ID);
    returnFolder = null;
  }

  if (returnFolder) {
    const returnFiles = returnFolder.getFiles();

    while (returnFiles.hasNext()) {
      const file = returnFiles.next();
      try {
        const csvData = Utilities.parseCsv(
          file.getBlob().getDataAsString()
        );

        for (let i = 1; i < csvData.length; i++) {
          try {
            if (!csvData[i] || csvData[i].length < 9) continue;

            const rawOrderId = csvData[i][8]; // Column I
            if (!rawOrderId) continue;

            const normalizedId = normalizeOrderId(rawOrderId);
            const rows = orderRowMap[normalizedId];
            if (!rows) continue;

            rows.forEach(idx => {
              statusOut[idx] = "RETURNED RECEIVED";
              sourceOut[idx] = "updated from meesho return csv";
            });

          } catch (rowErr) {
            Logger.log(`⚠️ Return row skipped | ${file.getName()} | CSV Row ${i + 1}`);
          }
        }

      } catch (fileErr) {
        Logger.log(`❌ Return file failed | ${file.getName()}`);
      }
    }
  }

  // =====================================================
  // 2️⃣ STATUS FILES (SAFE)
  // =====================================================
  let statusFolder;
  try {
    statusFolder = DriveApp.getFolderById(STATUS_FOLDER_ID);
  } catch (e) {
    Logger.log("⚠️ Meesho Status Folder Not Found: " + STATUS_FOLDER_ID);
    statusFolder = null;
  }

  if (statusFolder) {
    const statusFiles = statusFolder.getFiles();

    while (statusFiles.hasNext()) {
      const file = statusFiles.next();
      try {
        const csvData = Utilities.parseCsv(
          file.getBlob().getDataAsString()
        );

        for (let i = 1; i < csvData.length; i++) {
          try {
            if (!csvData[i] || csvData[i].length < 2) continue;

            const rawStatus = csvData[i][0];  // Column A
            const rawOrderId = csvData[i][1]; // Column B
            if (!rawOrderId || !rawStatus) continue;

            const normalizedId = normalizeOrderId(rawOrderId);
            const rows = orderRowMap[normalizedId];
            if (!rows) continue;

            rows.forEach(idx => {
              // ❗ return has priority
              if (sourceOut[idx]) return;

              statusOut[idx] = rawStatus.toString().trim();
              sourceOut[idx] = "updated from meesho status csv";
            });

          } catch (rowErr) {
            Logger.log(`⚠️ Status row skipped | ${file.getName()} | CSV Row ${i + 1}`);
          }
        }

      } catch (fileErr) {
        Logger.log(`❌ Status file failed | ${file.getName()}`);
      }
    }
  }

  // ================= WRITE BACK =================
  sheet.getRange(START_ROW, 28, numRows, 1)
       .setValues(statusOut.map(v => [v]));

  sheet.getRange(START_ROW, 29, numRows, 1)
       .setValues(sourceOut.map(v => [v]));

  Logger.log("✅ Meesho sync completed for sheet: " + targetSheetId);

    } catch (e) {
      Logger.log("❌ Failed for sheet ID: " + targetSheetId + " | " + e.message);
    }
  }

  Logger.log("✅ Meesho status + return sync completed for all sheets");
}