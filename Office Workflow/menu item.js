function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Office Works')
    .addItem('Seema Jain','launchJaipurTransfer')
    .addSeparator()
    .addItem('Update Order Status', 'Order_Status')
    .addSeparator()
    .addItem('Create Data Sheet','generateCategorySheet')
    .addSeparator()
    .addItem('Create Gst','Create_Gst')
    .addToUi();
  SpreadsheetApp.getUi()
    .createMenu('Certificate')
    .addItem('Astro Label', 'generateAstroDocDirectly')
    .addSeparator()
    .addItem('Open Item Search', 'showFloatingUI')
    .addSeparator()
    .addItem('Generate Certificate', 'launchCertificateGenerator')
    .addToUi();
  SpreadsheetApp.getUi()
    .createMenu("Label")
    .addItem("Generate Labels", "openLabelGeneratorUI")
    .addToUi();
  SpreadsheetApp.getUi()
    .createMenu('Late Sheet')
    .addItem('Late Orders','runOrdersReports')
    .addToUi();
  SpreadsheetApp.getUi()
    .createMenu('Update Tracking Id')
    .addItem('Ship Rocket','syncshiprocket')
    .addSeparator()
    .addItem('Ship Global','syncshipgloabal')
    .addToUi();
  SpreadsheetApp.getUi()
    .createMenu('Title Creation')
    .addItem('Title Creation','startAutomatedProcess')
    .addToUi();
}





function Order_Status() {
  Order_Status_Amazon();
  Order_Status_Flipkart();
  Order_Status_Meesho();
}