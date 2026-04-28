function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Assign leads')
    .addItem('Assign leads','runLeadDistributionFromWork')
    .addSeparator()
    .addItem('Indiamart Leads', 'runHourlyPipeline')
    .addToUi();

}