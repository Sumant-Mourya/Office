/**
 * Main entry point: Opens the date picker dialog.
 */
function generateCategorySheet() {
  const html = HtmlService.createHtmlOutputFromFile('0DatePicker_CreateSheet')
      .setWidth(450)
      .setHeight(400)
      .setTitle('Report Generator');
  SpreadsheetApp.getUi().showModalDialog(html, 'Select Date Range');
}

/**
 * Process selection with Overwrite and Same-Folder Strategy.
 * Fetches data from an External Spreadsheet defined in the "Configuration" sheet.
 */
function processDateSelection_CreateSheet(dateData) {
  if (!dateData || !dateData.start) throw new Error("Invalid date data.");

  // 1. ACCESS CONFIGURATION
  const activeSS = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = activeSS.getSheetByName("Configuration");
  if (!configSheet) throw new Error("Could not find 'Configuration' sheet.");

  const targetSSId = configSheet.getRange("B1").getValue().toString().trim();
  const targetSheetName = configSheet.getRange("B2").getValue().toString().trim();

  if (!targetSSId || !targetSheetName) {
    throw new Error("Configuration Error: Cell B1 (ID) or B2 (Sheet Name) is empty.");
  }

  // 2. GET DESTINATION FOLDER (Folder where THIS Master Spreadsheet lives)
  const currentFile = DriveApp.getFileById(activeSS.getId());
  const destinationFolder = currentFile.getParents().next();

  // 3. OPEN EXTERNAL SOURCE
  let sourceSheet;
  try {
    sourceSheet = SpreadsheetApp.openById(targetSSId).getSheetByName(targetSheetName);
  } catch (e) {
    throw new Error("Could not access Source Sheet. Check ID and Permissions.");
  }
  
  if (!sourceSheet) throw new Error("Source sheet '" + targetSheetName + "' not found in external spreadsheet.");

  let tempFiles = []; 
  let newSpreadsheet, docObj;

  try {
    let startDate = new Date(dateData.start);
    let endDate = new Date(dateData.end);
    startDate.setHours(0, 0, 0, 0);
    endDate.setHours(0, 0, 0, 0);

    const data = sourceSheet.getDataRange().getValues();
    const headers = data[1]; // Header row (Row 2)
    // Insert new column heading before column K (index 10)
    // Column letters: A=0 ... J=9, K=10 -> insert at 10
    if (headers && headers.splice) {
      headers.splice(10, 0, 'Want Certificate');
    }
    let categoryData = {};
    let allFilteredData = [];

    // Filter Logic
    for (let i = 2; i < data.length; i++) {
      const row = data[i];
      const rowDate = getDateOnly(row[7]); // Column H (unchanged)
      const category = row[11];            // Column L in source (original index 11)
      if (rowDate && rowDate >= startDate && rowDate <= endDate && category) {
        // Create a copy of the row and insert an empty value at new column K (index 10)
        const newRow = row.slice();
        newRow.splice(10, 0, ''); // keep 'Want Certificate' empty for now

        if (!categoryData[category]) categoryData[category] = [];
        categoryData[category].push(newRow);
        allFilteredData.push(newRow);
      }
    }

    if (allFilteredData.length === 0) throw new Error("No data found for this range.");

    const startStr = formatDate(startDate);
    const endStr = formatDate(endDate);
    const ssName = "Workflow";

    // 4. OVERWRITE LOGIC: Delete existing Spreadsheet with same name in this folder
    const existingSS = destinationFolder.getFilesByName(ssName);
    while (existingSS.hasNext()) {
      existingSS.next().setTrashed(true);
    }

    // Create Spreadsheet
    newSpreadsheet = SpreadsheetApp.create(ssName);
    let ssFile = DriveApp.getFileById(newSpreadsheet.getId());
    
    // Move to folder immediately
    destinationFolder.addFile(ssFile);
    DriveApp.getRootFolder().removeFile(ssFile);
    tempFiles.push(newSpreadsheet.getId());

    for (let category in categoryData) {
      const sheet = newSpreadsheet.insertSheet(category);
      sheet.appendRow(headers);
      categoryData[category].forEach(row => sheet.appendRow(row));
      
      // If category is Astro, trigger Doc generation with Overwrite logic
      if (category.toLowerCase() === "astro") {
        const docName = "Astro_Report_" + startStr + "_to_" + endStr;
        const existingDoc = destinationFolder.getFilesByName(docName);
        while (existingDoc.hasNext()) {
          existingDoc.next().setTrashed(true);
        }
      }
    }

    // Cleanup generated Spreadsheet
    const combinedSheet = newSpreadsheet.insertSheet("Today Data");
    combinedSheet.appendRow(headers);
    allFilteredData.forEach(row => combinedSheet.appendRow(row));
    const defaultSheet = newSpreadsheet.getSheetByName("Sheet1");
    if (defaultSheet) newSpreadsheet.deleteSheet(defaultSheet);

    return {
      ssUrl: newSpreadsheet.getUrl(),
      ssId: newSpreadsheet.getId(),
      docUrl: docObj ? docObj.url : null,
      docId: docObj ? docObj.id : null
    };

  } catch (err) {
    // Trash partial files if error occurs
    tempFiles.forEach(id => {
      try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {}
    });
    throw new Error(err.message);
  }
}

/**
 * HELPER FUNCTIONS
 */

function getDateOnly(v) { 
  return v instanceof Date ? new Date(v.getFullYear(), v.getMonth(), v.getDate()) : null; 
}

function formatDate(d) { 
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd-MM-yyyy"); 
}
