/**
 * Launch the Shiprocket UI
 */
function shiprocketdatamaker() {
  const html = HtmlService.createHtmlOutputFromFile('5ShipRocketDatePicker')
      .setWidth(400)
      .setHeight(350)
      .setTitle('Shiprocket Export');
  SpreadsheetApp.getUi().showModalDialog(html, 'Generate Shiprocket Data');
}

function sanitizeText(value) {
  if (!value) return "";

  let text = String(value);

  const charMap = {
    // German
    'ß':'ss','ẞ':'SS',
    'ä':'ae','Ä':'Ae',
    'ö':'oe','Ö':'Oe',
    'ü':'ue','Ü':'Ue',

    // Scandinavian
    'æ':'ae','Æ':'AE',
    'ø':'o','Ø':'O',
    'å':'a','Å':'A',

    // French
    'œ':'oe','Œ':'OE',
    'ç':'c','Ç':'C',

    // Spanish
    'ñ':'n','Ñ':'N',

    // Polish
    'ł':'l','Ł':'L',
    'ą':'a','Ą':'A',
    'ę':'e','Ę':'E',
    'ś':'s','Ś':'S',
    'ć':'c','Ć':'C',
    'ń':'n','Ń':'N',
    'ż':'z','Ż':'Z',
    'ź':'z','Ź':'Z',

    // Czech / Slovak
    'č':'c','Č':'C',
    'ď':'d','Ď':'D',
    'ě':'e','Ě':'E',
    'ň':'n','Ň':'N',
    'ř':'r','Ř':'R',
    'š':'s','Š':'S',
    'ť':'t','Ť':'T',
    'ž':'z','Ž':'Z',

    // Turkish
    'ğ':'g','Ğ':'G',
    'ş':'s','Ş':'S',
    'ı':'i','İ':'I',

    // Romanian
    'ș':'s','Ș':'S',
    'ț':'t','Ț':'T',
    'ă':'a','Ă':'A',
    'â':'a','Â':'A',
    'î':'i','Î':'I',

    // Icelandic
    'ð':'d','Ð':'D',
    'þ':'th','Þ':'TH',

    // Croatian / Serbian
    'đ':'d','Đ':'D',

    // Vietnamese
    'đ':'d','Đ':'D'
  };

  // Replace mapped characters
  text = text.replace(/[^\u0000-\u007E]/g, function(c) {
    return charMap[c] || c;
  });

  // Remove accents (é → e, á → a etc.)
  text = text.normalize("NFD").replace(/[\u0300-\u036f]/g, "");

  // Remove all special characters
  text = text.replace(/[^a-zA-Z0-9]/g, " ");

  // Remove extra spaces
  text = text.replace(/\s+/g, " ").trim();

  return text;
}

/**
 * Process logic called by the HTML
 */
/**
 * Core logic for Shiprocket: Creates a temp file and returns a download link.
 */
/**
 * Process logic for Shiprocket: Creates a temp file and returns a download link.
 */
function processAndGenerateShiprocket(dateRange) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Configuration");
  if (!configSheet) throw new Error("Configuration sheet not found.");

  const sourceId = configSheet.getRange("B1").getDisplayValue().trim();
  const sourceSheetName = configSheet.getRange("B2").getDisplayValue().trim();
  const sourceSheet = SpreadsheetApp.openById(sourceId).getSheetByName(sourceSheetName);
  const sourceData = sourceSheet.getDataRange().getValues();

  const headers = [
    "Order ID", "Channel", "Order Date", "Purpose", "Currency", "First Name", "Last Name", 
    "Email", "Mobile", "Address 1", "Address 2", "Country", "Postcode", "City", "State", 
    "Master SKU", "Product Name", "HSN Code", "Quantity", "Tax", "VAT", "Unit Price", 
    "Invoice Date", "Length", "Breadth", "Height", "Weight", "IOSS", "EORI", "Terms", 
    "Franchise", "Seller ID", "Courier ID"
  ];

  const results = [headers];
  const startDate = new Date(dateRange.start);
  const endDate = new Date(dateRange.end);
  endDate.setHours(23, 59, 59, 999);
  const todayDate = Utilities.formatDate(new Date(), "GMT+5:30", "dd-MM-yyyy");
  
  const excludeCountries = ["usa", "united states", "united kingdom", "france", "spain", "germany", "peru", "india"];

  for (let i = 1; i < sourceData.length; i++) {
    const row = sourceData[i];
    
    // 1. SKIP THE HEADER/EMPTY ROWS
    if (!row[6] || String(row[6]).toLowerCase().includes("order id")) continue;

    // 2. DATE FILTER
    if (!row[7]) continue;
    const rowDate = new Date(row[7]);
    if (isNaN(rowDate.getTime()) || rowDate < startDate || rowDate > endDate) continue;

    // 3. COUNTRY FILTER
    const rawCountry = String(row[23] || "").trim();
    if (excludeCountries.includes(rawCountry.toLowerCase())) continue;

    // Initialize row with nulls to ensure empty cells stay empty
    let newRow = new Array(33).fill(null); 
    
    newRow[0] = row[6] || null;           
    newRow[1] = "Custom";         
    newRow[2] = todayDate;        
    newRow[3] = "Sample";         
    newRow[4] = "USD";            

    // Name Logic
    const fullName = String(row[17] || "").trim();
    if (fullName) {
      const nameParts = fullName.split(/\s+/);
      newRow[5] = sanitizeText(nameParts[0]); 
      newRow[6] = nameParts.length > 1 ? sanitizeText(nameParts.slice(1).join(" ")) : ".";
    }

    newRow[7] = "uttarahomes@gmail.com"; 

    // --- MOBILE LOGIC (Cleaned & Padded) ---
    let rawPhone = String(row[24] || "").trim();
    if (rawPhone !== "") {
      let cleanPhone = rawPhone.toLowerCase().split("ext")[0].replace(/\D/g, "");
      // Ensure exactly 10 digits
      newRow[8] = cleanPhone.length > 10 ? cleanPhone.slice(-10) : cleanPhone.padStart(10, '0');
    } else {
      newRow[8] = null; // Keep empty if source is empty
    }

    newRow[9]  = sanitizeText(row[18]);   
    newRow[10] = sanitizeText(row[19]);   
    newRow[11] = sanitizeText(row[23]);   
    newRow[12] = row[22] || null;   
    newRow[13] = sanitizeText(row[20]);   
    newRow[14] = sanitizeText(row[21]);   
    newRow[15] = sanitizeText(row[9]);   
    newRow[16] = "Fashion Jewellry"; 
    newRow[17] = "71179010";         
    newRow[18] = "1";                
    newRow[21] = "12";               
    newRow[22] = todayDate;          
    newRow[23] = "10";               
    newRow[24] = "3";                
    newRow[25] = "2";                
    newRow[26] = "0.05";             
    newRow[29] = "CIF";              

    results.push(newRow);
  }

  // Create temporary Spreadsheet
  const tempSS = SpreadsheetApp.create("Shiprocket_Export_" + Utilities.formatDate(new Date(), "GMT+5:30", "yyyy-MM-dd_HHmm"));
  const tempSheet = tempSS.getSheets()[0];
  
  // Set the Mobile Column (Column I) to Plain Text BEFORE inserting data to preserve zeros
  tempSheet.getRange(1, 9, results.length, 1).setNumberFormat("@"); 
  
  // Write all data at once
  tempSheet.getRange(1, 1, results.length, headers.length).setValues(results);
  
  // Clean up: delete file after 10 minutes to save Drive space (optional)
  // DriveApp.getFileById(tempSS.getId()).setTrashed(true); 

  return {
    url: "https://docs.google.com/spreadsheets/d/" + tempSS.getId() + "/export?format=xlsx",
    count: results.length - 1
  };
}