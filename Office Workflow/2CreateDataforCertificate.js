/**
 * Scans the Workflow file, creates/updates a tab named "Certificate Data" 
 * in the current spreadsheet, and saves the current Spreadsheet ID to B5.
 */
function processWorkflowFile() {
  const workflowFolderId = "1j3EYxaX4l6umHu9PVOLKE0hPpDksgz9O";
  const targetFileName = "Workflow";
  const targetSheetName = "Certificate Data";

  const folder = DriveApp.getFolderById(workflowFolderId);
  const files = folder.getFilesByName(targetFileName);

  if (!files.hasNext()) {
    throw new Error("'Workflow' file not found in the specified folder.");
  }

  const workflowSS = SpreadsheetApp.open(files.next());
  const masterSS = SpreadsheetApp.getActiveSpreadsheet();
  
  const itemDetailsSheet = masterSS.getSheetByName("Item Details");
  if (!itemDetailsSheet) {
    throw new Error("'Item Details' sheet not found in Master spreadsheet.");
  }
  const itemRefData = itemDetailsSheet.getDataRange().getValues();
  const itemMap = buildItemMap(itemRefData);

  let outputData = [];
  const sheets = workflowSS.getSheets();

  sheets.forEach(sheet => {
    if (sheet.getName() === "Today Data") return;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const colK = data[i][10]; 
      const colL = String(data[i][11] || "").toLowerCase(); 

      if (colK === 1 || colK === "1") {
        for (let keyword in itemMap) {
          if (colL.includes(keyword.toLowerCase())) {
            const details = itemMap[keyword];
            
            // Updated mapping to include the "Item" column
            outputData.push([
              details["Name"] || "NA",
              details["Color"] || "NA",
              details["Weight"] || "NA",
              details["Shape & Cut"] || "NA",
              details["Measurement"] || "NA",
              details["Hindi Name"] || "NA",
              details["Item"] || "NA",           // NEW COLUMN ADDED HERE
              details["Specific Gravity"] || "NA",
              details["Optic Character"] || "NA",
              details["Refractive Index"] || "NA",
              "NA", // Remark
              "NA", // Issued To
              details["Images"] || "NA"
            ]);
            break; 
          }
        }
      }
    }
  });

  if (outputData.length === 0) {
    SpreadsheetApp.getUi().alert("No matching data found to process.");
    return;
  }

  // Handle Sheet Overwrite
  let certSheet = masterSS.getSheetByName(targetSheetName);
  if (certSheet) {
    masterSS.deleteSheet(certSheet);
  }
  certSheet = masterSS.insertSheet(targetSheetName);

  // Updated Headers to include "Item"
  const headers = [
    "Name", "Color", "Weight", "Shape & Cut", "Measurement", "Hindi Name", 
    "Item", "Specific Gravity", "Optic Character", "Refractive Index", 
    "Remark", "Issued To", "Images"
  ];

  certSheet.appendRow(headers);
  certSheet.getRange(2, 1, outputData.length, headers.length).setValues(outputData);
  certSheet.autoResizeColumns(1, headers.length);
  certSheet.getRange(1, 1, 1, headers.length).setBackground("#d9ead3").setFontWeight("bold");
  
  certSheet.activate();

  // Store ID in Configuration
  const masterId = masterSS.getId();
  const configSheet = masterSS.getSheetByName("Configuration");
  if (configSheet) {
    configSheet.getRange("B5").setValue(masterId);
  }

  // UI remains the same
  showSuccessUI(outputData.length);
}

/**
 * Success UI Helper (Extracted for cleanliness)
 */
function showSuccessUI(count) {
  const htmlContent = `
    <style>
      body { margin: 0; padding: 0; font-family: 'Segoe UI', Tahoma, sans-serif; overflow: hidden; }
      .container { display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100vh; text-align: center; }
      .icon { font-size: 40px; margin-bottom: 10px; }
      .title { color: #0b8043; font-weight: bold; font-size: 18px; margin-bottom: 5px; }
      .msg { color: #555; font-size: 14px; margin-bottom: 20px; }
      .btn { background-color: #1a73e8; color: white; padding: 10px 30px; border: none; border-radius: 4px; font-weight: bold; cursor: pointer; }
    </style>
    <div class="container">
      <div class="icon">✅</div>
      <div class="title">Tab Updated!</div>
      <div class="msg">The "Certificate Data" tab has been refreshed with <b>${count} records</b>.</div>
      <button class="btn" onclick="google.script.host.close()">Got it</button>
    </div>
  `;
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent).setWidth(400).setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, ' ');
}

/**
 * Helper to map Item Details.
 */
function buildItemMap(data) {
  const headers = data[0];
  const map = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const nameKey = String(row[0]).trim();
    if (!nameKey) continue;
    
    map[nameKey] = {};
    headers.forEach((header, index) => {
      let val = row[index];
      map[nameKey][header] = (val === "" || val === null || val === undefined) ? "NA" : val;
    });
  }
  return map;
}