function launchCertificateGenerator() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Sheet Names
  const targetSheetName = "Certificate Data";
  const configSheetName = "Configuration";
  
  const certSheet = ss.getSheetByName(targetSheetName);
  const configSheet = ss.getSheetByName(configSheetName);
  
  // 1. Validation
  if (!certSheet) {
    ui.alert("Error: Could not find the '" + targetSheetName + "' tab.");
    return;
  }
  if (!configSheet) {
    ui.alert("Error: Could not find the '" + configSheetName + "' tab.");
    return;
  }

  // 2. Get Numbers from Configuration (D1 = Report, D2 = Lot)
  // .getValue() handles numbers directly
  let startReportNo = parseInt(configSheet.getRange("D1").getValue());
  let startLotNo = parseInt(configSheet.getRange("D2").getValue());

  if (isNaN(startReportNo) || isNaN(startLotNo)) {
    ui.alert("Error: Configuration sheet D1 or D2 does not contain valid numbers.");
    return;
  }

  const data = certSheet.getDataRange().getValues();
  // Filter out empty rows to get actual count of certificates to be made
  const rowsToProcess = data.slice(1).filter(row => row[0] !== "");
  
  if (rowsToProcess.length === 0) {
    ui.alert("No data found in Certificate Data sheet.");
    return;
  }

  // 3. Map Data for HTML Template
  const certData = rowsToProcess.map((row, index) => {
    
    // Image URL Cleanup (Handles =IMAGE formulas)
    let rawImage = String(row[12] || ""); 
    let cleanUrl = rawImage;
    if (rawImage.includes('IMAGE("')) {
      try { cleanUrl = rawImage.split('("')[1].split('")')[0]; } catch(e) { cleanUrl = rawImage; }
    }

    return {
      reportNo: (startReportNo + index).toString(), 
      lotNo:    (startLotNo + index).toString(),    
      name:     row[0],
      color:    row[1],    
      weight:   row[2],    
      shape:    row[3],    
      measure:  row[4],    
      hindi:    row[5],    
      itemName: row[6],    
      sg:       row[7],    
      optic:    row[8],    
      ri:       row[9],    
      remark:   row[10],
      issue:    row[11] || "NA",
      imageUrl: cleanUrl,  
      imageTitle: row[0]   
    };
  });

  // 4. Update Configuration Sheet with NEW starting numbers for next time
  const nextReportNo = startReportNo + certData.length;
  const nextLotNo = startLotNo + certData.length;
  
  configSheet.getRange("D1").setValue(nextReportNo);
  configSheet.getRange("D2").setValue(nextLotNo);

  // 5. Generate UI
  const template = HtmlService.createTemplateFromFile('2Certificate Page'); 
  template.certs = certData;
  
  const htmlOutput = template.evaluate()
      .setWidth(1100)
      .setHeight(850)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setTitle('GAL Print System');
      
  ui.showModalDialog(htmlOutput, 'Generating ' + certData.length + ' Certificates...');
}