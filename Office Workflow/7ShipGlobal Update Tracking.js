function syncshipgloabal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Get Configs
  const configSheet = ss.getSheetByName("Configuration");
  if (!configSheet) throw new Error("Sheet 'Configuration' not found!");
  
  const targetSheetId = configSheet.getRange("B1").getValue();
  const targetSheetName = configSheet.getRange("B2").getValue();
  
  // 2. Access Master Sheet
  const masterSS = SpreadsheetApp.openById(targetSheetId);
  const masterSheet = masterSS.getSheetByName(targetSheetName);
  const masterLastRow = masterSheet.getLastRow();
  
  if (masterLastRow < 3) return; 

  // 3. Get the LATEST Excel File (Any name)
  const folderId = "1rFY1PAeIyjFDWYwUj11ct6cqT77XPcno";
  const folder = DriveApp.getFolderById(folderId);
  
  // CHANGED: Now looking for ANY Excel file type instead of a specific name
  const files = folder.getFilesByType(MimeType.MICROSOFT_EXCEL);
  
  let latestFile = null;
  let lastUpdate = 0;

  while (files.hasNext()) {
    let file = files.next();
    // Compare timestamps to find the most recently uploaded/edited file
    if (file.getLastUpdated().getTime() > lastUpdate) {
      lastUpdate = file.getLastUpdated().getTime();
      latestFile = file;
    }
  }

  if (!latestFile) {
    console.log("No Excel files found in the folder.");
    return;
  }

  console.log("Found latest file: " + latestFile.getName());

  // 4. Convert Excel to Temp Google Sheet
  const blob = latestFile.getBlob();
  const resource = {
    title: "Temp_Sync_File_" + new Date().getTime(),
    mimeType: MimeType.GOOGLE_SHEETS
  };
  
  const tempFile = Drive.Files.insert(resource, blob, {convert: true});
  const tempSS = SpreadsheetApp.openById(tempFile.id);
  const excelSheet = tempSS.getSheets()[0];
  const excelData = excelSheet.getDataRange().getValues();
  
  // Map Excel Data: Col E (index 4) -> [Col T (index 19), Col U (index 20)]
  const excelMap = new Map();
  for (let i = 0; i < excelData.length; i++) {
    let orderId = excelData[i][4]?.toString().trim();
    if (orderId) {
      excelMap.set(orderId, [excelData[i][19], excelData[i][20]]);
    }
  }

  // 5. Process Master Sheet and Update
  const masterValues = masterSheet.getRange(1, 1, masterLastRow, 27).getValues();
  let updatesCount = 0;

  for (let i = 2; i < masterLastRow; i++) {
    let colAA = masterValues[i][26]; 
    let colG = masterValues[i][6].toString().trim();

    if ((colAA === "" || colAA === null) && colG !== "") {
      if (excelMap.has(colG)) {
        let matchedData = excelMap.get(colG);
        masterSheet.getRange(i + 1, 26).setValue(matchedData[0]); 
        masterSheet.getRange(i + 1, 27).setValue(matchedData[1]); 
        updatesCount++;
      }
    }
  }

  // 6. Cleanup: Delete temp file
  Drive.Files.remove(tempFile.id);

  console.log("Sync Complete! Updated " + updatesCount + " rows using: " + latestFile.getName());
}