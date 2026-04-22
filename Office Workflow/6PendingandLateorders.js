function runOrdersReports(){

  Logger.log("===== START ORDER REPORT GENERATION =====");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ss.getSheetByName("Configuration");

  const sourceId = config.getRange("B1").getDisplayValue().trim();
  const sourceSheetName = config.getRange("B2").getDisplayValue().trim();

  const sourceSS = SpreadsheetApp.openById(sourceId);
  const sourceSheet = sourceSS.getSheetByName(sourceSheetName);

  const lastRow = sourceSheet.getLastRow();
  const totalColumns = sourceSheet.getLastColumn();

  Logger.log("Rows detected: " + lastRow);
  Logger.log("Columns detected: " + totalColumns);

  const values = sourceSheet.getRange(1,1,lastRow,totalColumns).getValues();
  const formulas = sourceSheet.getRange(1,1,lastRow,totalColumns).getFormulas();
  const displayDates = sourceSheet.getRange(1,8,lastRow,1).getDisplayValues();
  const backgrounds = sourceSheet.getRange(1,8,lastRow,1).getBackgrounds();

  const today = new Date();

  let pendingRows = [];
  let last15Rows = [];

  let totalWhiteRows = 0;

  // Header (row 2)
  pendingRows.push(values[1]);
  last15Rows.push(values[1]);

  for(let i=2;i<values.length;i++){

    const bgColor = backgrounds[i][0];

    // Skip non-white rows completely
    if(bgColor && bgColor !== "#ffffff" && bgColor !== "#fff") continue;

    totalWhiteRows++;

    const displayDate = displayDates[i][0];
    if(!displayDate) continue;

    const parts = displayDate.split(/[\/\-]/);
    const purchaseDate = new Date(parts[2],parts[1]-1,parts[0]);

    const diffDays=(today-purchaseDate)/(1000*60*60*24);

    let row=[];

    for(let c=0;c<totalColumns;c++){

      let formula=formulas[i][c];

      if(formula){

        if(/=IMAGE\(/i.test(formula)){

          const newRow=(diffDays>15?pendingRows.length:last15Rows.length)+1;

          formula=formula.replace(/([A-Z]+)\d+/,function(match,col){
            return col+newRow;
          });

        }

        row.push(formula);

      }else{

        row.push(values[i][c]);

      }

    }

    if(diffDays > 15){
      pendingRows.push(row);
    }else{
      last15Rows.push(row);
    }

  }

  Logger.log("Total white rows detected: " + totalWhiteRows);
  Logger.log("Pending rows (>15 days): " + (pendingRows.length-1));
  Logger.log("Last 15 days rows: " + (last15Rows.length-1));

  Logger.log("Validation check (pending + last15): " + ((pendingRows.length-1)+(last15Rows.length-1)));

  const fileName="Pending Orders Report";

  const currentFile=DriveApp.getFileById(ss.getId());
  const parentFolder=currentFile.getParents().next();

  const existingFiles=parentFolder.getFilesByName(fileName);
  while(existingFiles.hasNext()) existingFiles.next().setTrashed(true);

  const report=SpreadsheetApp.create(fileName);
  Utilities.sleep(2000);

  const reportFile=DriveApp.getFileById(report.getId());
  parentFolder.addFile(reportFile);
  DriveApp.getRootFolder().removeFile(reportFile);

  const pendingSheet=report.getSheets()[0];
  pendingSheet.setName("Pending Order");

  pendingSheet.getRange(1,1,pendingRows.length,totalColumns)
  .setValues(pendingRows);

  const last15Sheet=report.insertSheet("Last 15 Days Orders");

  last15Sheet.getRange(1,1,last15Rows.length,totalColumns)
  .setValues(last15Rows);

  const url="https://docs.google.com/spreadsheets/d/"+report.getId();

  const html=HtmlService.createHtmlOutput(`
    <div style="font-family:Arial;text-align:center;padding:20px">
      <h2>Reports Created Successfully</h2>
      <p>Pending Orders and Last 15 Days reports generated.</p>
      <a href="${url}" target="_blank">
        <button style="padding:12px 25px;font-size:16px;background:#1a73e8;color:white;border:none;border-radius:6px">
          Open Report Spreadsheet
        </button>
      </a>
    </div>
  `).setWidth(360).setHeight(250);

  SpreadsheetApp.getUi().showModalDialog(html,"Report Ready");

  Logger.log("===== REPORT GENERATION COMPLETE =====");

}