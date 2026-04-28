function processFollowUpsAndSync() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Config"); // your config sheet
  const workSheet = ss.getSheetByName("Work");

  if (!configSheet || !workSheet) {
    Logger.log("Missing Config or Work sheet");
    return;
  }

  const configData = configSheet.getRange(2, 2, configSheet.getLastRow() - 1, 1).getValues();

  // Build Work Sheet Map (phone → row index)
  const workData = workSheet.getDataRange().getValues();
  const workMap = new Map();

  for (let i = 1; i < workData.length; i++) {
    const phone = normalizePhone(workData[i][4]);
    if (!phone) continue;

    if (!workMap.has(phone)) workMap.set(phone, []);
    workMap.get(phone).push(i);
  }

  // LOOP EMPLOYEES
  configData.forEach(row => {
    const sheetId = row[0];
    if (!sheetId) return;

    let empSS;
    try {
      empSS = SpreadsheetApp.openById(sheetId);
    } catch (e) {
      Logger.log("Cannot open: " + sheetId);
      return;
    }

    const leadSheet = empSS.getSheetByName("LeadCallMsg");
    if (!leadSheet) return;

    let followSheet = empSS.getSheetByName("FollowUp");
    if (!followSheet) followSheet = empSS.insertSheet("FollowUp");

    const leadData = leadSheet.getDataRange().getValues();
    let newFollowRows = [];
    let rowsToDelete = [];

    // STEP 1: MOVE DATA WHERE COLUMN H FILLED
    for (let i = 1; i < leadData.length; i++) {
      const rowData = leadData[i];

      if (rowData[7]) { // Column H (index 7)
        const phone = normalizePhone(rowData[4]);

        newFollowRows.push([
          rowData[0], // A
          rowData[6], // G
          rowData[7], // H
          rowData[8], // I
          rowData[9], // J
          rowData[10], // K
          rowData[11], // L
          rowData[12], // M
          phone
        ]);

        rowsToDelete.push(i + 1);
      }
    }

    // Append to FollowUp
    if (newFollowRows.length > 0) {
      followSheet
        .getRange(followSheet.getLastRow() + 1, 1, newFollowRows.length, newFollowRows[0].length)
        .setValues(newFollowRows);
    }

    // Delete rows safely (bottom to top)
    rowsToDelete.reverse().forEach(r => leadSheet.deleteRow(r));

    // STEP 2: SYNC FOLLOWUP → WORK SHEET
    const followData = followSheet.getDataRange().getValues();

    for (let i = 1; i < followData.length; i++) {
      const fRow = followData[i];
      const phone = normalizePhone(fRow[8]);

      if (!phone) continue;

      const matches = workMap.get(phone);
      if (!matches) continue;

      matches.forEach(idx => {
        // Update Work Sheet columns (A, G–M same as your logic)
        workData[idx][0] = fRow[0];  // A
        workData[idx][6] = fRow[1];  // G
        workData[idx][7] = fRow[2];  // H
        workData[idx][8] = fRow[3];  // I
        workData[idx][9] = fRow[4];  // J
        workData[idx][10] = fRow[5]; // K
        workData[idx][11] = fRow[6]; // L
        workData[idx][12] = fRow[7]; // M
      });
    }
  });

  // FINAL WRITE BACK
  workSheet.getRange(1, 1, workData.length, workData[0].length).setValues(workData);

  Logger.log("Follow-up sync completed");
}

/**
 * 📱 Normalize phone → last 10 digits
 */
function normalizePhone(val) {
  if (!val) return "";

  let num = val.toString().replace(/\D/g, "");
  if (num.length >= 10) return num.slice(-10);

  return "";
}