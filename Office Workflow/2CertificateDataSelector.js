function showFloatingUI() {
  const html = HtmlService.createHtmlOutputFromFile('2CertificateSearchUi')
    .setWidth(450).setHeight(350).setTitle('Lightning Entry');
  SpreadsheetApp.getUi().showModelessDialog(html, ' ');
}

// Loads all data once when UI opens
function getInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const itemSheet = ss.getSheetByName("Item Details");
  const data = itemSheet.getDataRange().getValues();
  return {
    headers: data[0],
    rows: data.slice(1)
  };
}

// Background Task: Just appends, no calculation
function silentAdd(rowData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const certSheet = ss.getSheetByName("Certificate Data");
  certSheet.appendRow(rowData);
}
