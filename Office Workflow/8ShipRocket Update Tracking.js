function syncshiprocket() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Get Master Sheet Configs
  const configSheet = ss.getSheetByName("Configuration");
  if (!configSheet) throw new Error("Sheet 'Configuration' not found!");
  
  const targetSheetId = configSheet.getRange("B1").getValue().toString().trim();
  const targetSheetName = configSheet.getRange("B2").getValue().toString().trim();
  
  // Hardcoded Folder ID as requested
  const folderId = "1Rt0KdF5ktafz8g6ZX6VU-7Sw9UdUAvQW"; 
  
  // 2. Access Master Sheet
  const masterSS = SpreadsheetApp.openById(targetSheetId);
  const masterSheet = masterSS.getSheetByName(targetSheetName);
  const masterLastRow = masterSheet.getLastRow();
  
  if (masterLastRow < 3) return; 

  // 3. Find the Latest CSV File
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFilesByType(MimeType.CSV); // Changed to CSV
  
  let latestFile = null;
  let lastUpdate = 0;

  while (files.hasNext()) {
    let file = files.next();
    if (file.getLastUpdated().getTime() > lastUpdate) {
      lastUpdate = file.getLastUpdated().getTime();
      latestFile = file;
    }
  }

  if (!latestFile) {
    console.log("No CSV files found in the folder.");
    return;
  }

  // 4. Read CSV Data directly (No temp file needed)
  const csvData = Utilities.parseCsv(latestFile.getBlob().getDataAsString());
  
  // 5. Map CSV Data: Match Col A (index 0) | Data Col AE (30) & AG (32)
  const csvMap = new Map();
  for (let i = 0; i < csvData.length; i++) {
    let orderIdFromCSV = csvData[i][0]?.toString().trim();
    if (orderIdFromCSV) {
      // Index 30 is Col AE, Index 32 is Col AG
      csvMap.set(orderIdFromCSV, [csvData[i][30], csvData[i][32]]);
    }
  }

  // 6. Process Master Sheet and Update
  const masterValues = masterSheet.getRange(1, 1, masterLastRow, 27).getValues();
  let updatesCount = 0;

  for (let i = 2; i < masterLastRow; i++) {
    let colAA = masterValues[i][26]; 
    let colG = masterValues[i][6].toString().trim();

    // If Tracking (AA) is empty and Order ID (G) exists
    if ((colAA === "" || colAA === null) && colG !== "") {
      if (csvMap.has(colG)) {
        let matchedData = csvMap.get(colG);
        let courierName = matchedData[0]?.toString() || "";
        let rawTracking = matchedData[1]?.toString() || "";

        // Remove quotes: ''LP953931802IN'' -> LP953931802IN
        let cleanTracking = rawTracking.replace(/['"]+/g, '');

        if (courierName.toLowerCase().includes("india post")) {
          // Write to Z (Courier) and AA (Tracking)
          masterSheet.getRange(i + 1, 26).setValue("India Post"); 
          masterSheet.getRange(i + 1, 27).setValue(cleanTracking); 
          updatesCount++;
        }
      }
    }
  }

  console.log("Sync Complete! Updated " + updatesCount + " India Post rows using: " + latestFile.getName());
}