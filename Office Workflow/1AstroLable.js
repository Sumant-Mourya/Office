/**
 * Main entry point for the Direct Astro Doc Generator.
 */
function generateAstroDocDirectly() {
  const html = HtmlService.createHtmlOutputFromFile('1Astro_DatePicker')
      .setWidth(450)
      .setHeight(400)
      .setTitle('Astro Doc Generator');
  SpreadsheetApp.getUi().showModalDialog(html, 'Select Date Range');
}

/**
 * Process selection: Skips spreadsheet creation and goes straight to Google Doc.
 */
function processDateSelection_Astro(dateData) {
  if (!dateData || !dateData.start) throw new Error("Invalid date data.");

  // 1. ACCESS CONFIGURATION
  const activeSS = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = activeSS.getSheetByName("Configuration");
  if (!configSheet) throw new Error("Could not find 'Configuration' sheet.");

  const targetSSId = configSheet.getRange("B1").getValue().toString().trim();
  const targetSheetName = configSheet.getRange("B2").getValue().toString().trim();

  // 2. FOLDER & SOURCE SETUP
  const destinationFolder = DriveApp.getFileById(activeSS.getId()).getParents().next();
  let sourceSheet;
  try {
    sourceSheet = SpreadsheetApp.openById(targetSSId).getSheetByName(targetSheetName);
  } catch (e) {
    throw new Error("Could not access Source Sheet. Check ID and Permissions.");
  }
  
  if (!sourceSheet) throw new Error("Source sheet '" + targetSheetName + "' not found.");

  try {
    let startDate = new Date(dateData.start);
    let endDate = new Date(dateData.end);
    startDate.setHours(0, 0, 0, 0);
    endDate.setHours(0, 0, 0, 0);

    const data = sourceSheet.getDataRange().getValues();
    const headers = data[1]; 
    
    // Insert 'Want Certificate' header for consistent indexing
    if (headers && headers.splice) {
      headers.splice(10, 0, 'Want Certificate');
    }

    let astroRows = [];

    // 3. FILTER LOGIC (Astro Only)
    for (let i = 2; i < data.length; i++) {
      const row = data[i];
      const rowDate = getDateOnly(row[7]); // Column H
      const category = row[11];            // Column L (Original Source)
      
      if (rowDate && rowDate >= startDate && rowDate <= endDate && category && category.toLowerCase() === "astro") {
        const newRow = row.slice();
        newRow.splice(10, 0, ''); // Add empty index for 'Want Certificate'
        astroRows.push(newRow);
      }
    }

    if (astroRows.length === 0) throw new Error("No Astro data found for this range.");

    const startStr = formatDate(startDate);
    const endStr = formatDate(endDate);
    const docName = "Astro_Report";

    // 4. OVERWRITE LOGIC
    const existingDoc = destinationFolder.getFilesByName(docName);
    while (existingDoc.hasNext()) {
      existingDoc.next().setTrashed(true);
    }

    // 5. CREATE DOC
    const docObj = createAstroDoc(astroRows, headers, startStr, endStr, destinationFolder);

    return {
      docUrl: docObj.url,
      docId: docObj.id,
      ssUrl: null // No spreadsheet created in this version
    };

  } catch (err) {
    throw new Error(err.message);
  }
}

/**
 * Creates the 3x3 grid Document.
 * (This is the same logic as your original to ensure layout consistency)
 */
function createAstroDoc(rows, headers, startStr, endStr, destinationFolder) {
  const docName = "Astro_Report";
  const doc = DocumentApp.create(docName);
  const docId = doc.getId();
  
  const docFile = DriveApp.getFileById(docId);
  destinationFolder.addFile(docFile);
  DriveApp.getRootFolder().removeFile(docFile);

  const body = doc.getBody();
  body.setPageWidth(595.27).setPageHeight(841.89);
  body.setMarginLeft(5).setMarginRight(5).setMarginTop(5).setMarginBottom(5);

  for (let i = 0; i < rows.length; i += 9) {
    const chunk = rows.slice(i, i + 9);
    const table = body.appendTable();
    table.setAttributes({ [DocumentApp.Attribute.WIDTH]: 585, [DocumentApp.Attribute.BORDER_WIDTH]: 1 });

    let currentRow;
    chunk.forEach((rowData, index) => {
      if (index % 3 === 0) {
        currentRow = table.appendTableRow();
        currentRow.setMinimumHeight(270); 
      }

      const cell = currentRow.appendTableCell();
      cell.setPaddingTop(2).setPaddingBottom(2).setPaddingLeft(5).setPaddingRight(5);
      cell.setVerticalAlignment(DocumentApp.VerticalAlignment.TOP);

      const colG = String(rowData[6] || ""); 
      const colH = rowData[7] instanceof Date ? Utilities.formatDate(rowData[7], Session.getScriptTimeZone(), "dd-MM-yyyy") : String(rowData[7] || "");
      const colK = String(rowData[11] || ""); 
      const colN = String(rowData[14] || ""); 
      const imgUrl = String(rowData[30] || ""); 

      const headerTable = cell.appendTable([[colG, colH]]);
      headerTable.setBorderWidth(0);
      const hRow = headerTable.getRow(0);
      
      [0, 1].forEach(idx => {
        const p = hRow.getCell(idx).getChild(0).asParagraph();
        p.setBold(true).setFontSize(9).setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1.0);
        if (idx === 1) p.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
      });

      const pName = cell.appendParagraph(colK);
      pName.setAlignment(DocumentApp.HorizontalAlignment.CENTER)
           .setItalic(true).setSpacingBefore(2).setSpacingAfter(2)
           .setLineSpacing(1.0).setFontSize(colK.length > 60 ? 8 : 15);

      if (imgUrl && imgUrl.toLowerCase().startsWith('http')) {
        try {
          const response = UrlFetchApp.fetch(imgUrl);
          const img = cell.appendImage(response.getBlob());
          const ratio = Math.min(160 / img.getWidth(), 115 / img.getHeight());
          img.setWidth(img.getWidth() * ratio).setHeight(img.getHeight() * ratio);
          img.getParent().asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        } catch (e) {
          cell.appendParagraph("[No Image]").setAlignment(DocumentApp.HorizontalAlignment.CENTER).setFontSize(8);
        }
      }

      const pFooter = cell.appendParagraph(colN);
      pFooter.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setItalic(true).setFontSize(10);

      if (cell.getChild(0).getType() === DocumentApp.ElementType.PARAGRAPH && cell.getChild(0).asParagraph().getText() === "") {
        cell.removeChild(cell.getChild(0));
      }
    });

    while (currentRow.getNumChildren() < 3) { 
      currentRow.appendTableCell(" ").setAttributes({ [DocumentApp.Attribute.BORDER_COLOR]: '#FFFFFF' });
    }

    if (i + 9 < rows.length) body.appendPageBreak();
  }

  doc.saveAndClose();
  return { url: doc.getUrl(), id: docId };
}

// Re-using same Helpers
function getDateOnly(v) { return v instanceof Date ? new Date(v.getFullYear(), v.getMonth(), v.getDate()) : null; }
function formatDate(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd-MM-yyyy"); }