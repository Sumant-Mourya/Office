function openLabelGeneratorUI() {
  const html = HtmlService.createHtmlOutputFromFile('4ShipGlobalANDrockethtml')
    .setWidth(420)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Generating Labels');
}

const REFUND_CHECK_FOLDER_ID = '1GEzKGmiSTGL4KOHWpXt4e4AwfhBfTo5O';

function getLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName('log') || ss.getSheetByName('Log');
}

function logEvent_(level, message, details) {
  const lv = String(level || 'INFO').toUpperCase();
  let text = String(message || '');

  if (details !== undefined && details !== null) {
    if (typeof details === 'string') {
      text += ' | ' + details;
    } else {
      try {
        text += ' | ' + JSON.stringify(details);
      } catch (e) {
        text += ' | [unserializable details]';
      }
    }
  }

  try {
    const logSheet = getLogSheet_();
    if (!logSheet) {
      Logger.log(lv + ': ' + text);
      return;
    }
    logSheet.appendRow([new Date(), lv, text]);
  } catch (e) {
    Logger.log('LOG_WRITE_FAILED: ' + e + ' | ' + lv + ': ' + text);
  }
}

function getEffectiveLastColumn_(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const header = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0] || [];
  for (let i = header.length - 1; i >= 0; i--) {
    if (String(header[i] || '').trim() !== '') return i + 1;
  }
  return lastCol;
}

function normalizeOrderId(value) {
  let id = String(value == null ? '' : value).trim();
  if (!id) return '';

  // Matching rule: keep all characters as-is, only remove spaces and hyphens.
  id = id.replace(/[\s-]+/g, '');
  return id;
}

function getExternalMainSheetFromConfig_() {
  const activeSs = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = activeSs.getSheetByName('Configuration');
  if (!configSheet) throw new Error('Configuration sheet not found.');

  const mainSpreadsheetId = String(configSheet.getRange('B1').getDisplayValue() || '').trim();
  const mainSheetName = String(configSheet.getRange('B2').getDisplayValue() || '').trim();
  if (!mainSpreadsheetId) throw new Error('Configuration!B1 (spreadsheet ID) is empty.');
  if (!mainSheetName) throw new Error('Configuration!B2 (sheet name) is empty.');

  const mainSs = SpreadsheetApp.openById(mainSpreadsheetId);
  const mainSheet = mainSs.getSheetByName(mainSheetName);
  if (!mainSheet) throw new Error('Sheet not found in external spreadsheet: ' + mainSheetName);

  return {
    spreadsheetId: mainSpreadsheetId,
    sheetName: mainSheetName,
    sheet: mainSheet
  };
}

function markExternalSheetRowsRedByOrderIds_(orderIdsSet) {
  const ids = Array.from(orderIdsSet || []);
  if (ids.length === 0) {
    return { externalRowsMarkedRed: 0, checkedRows: 0 };
  }

  const external = getExternalMainSheetFromConfig_();
  const sheet = external.sheet;
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  if (lastRow <= 1) {
    return { externalRowsMarkedRed: 0, checkedRows: 0 };
  }

  const idSet = new Set(ids);
  const idDisplayValues = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();

  let marked = 0;
  for (let i = 0; i < idDisplayValues.length; i++) {
    const normalized = normalizeOrderId(idDisplayValues[i][0]);
    if (!normalized) continue;
    if (!idSet.has(normalized)) continue;

    sheet.getRange(i + 2, 1, 1, lastCol).setBackground('#ff0000');
    marked++;
  }

  return {
    externalRowsMarkedRed: marked,
    checkedRows: idDisplayValues.length,
    externalSheetName: external.sheetName,
    externalSpreadsheetId: external.spreadsheetId
  };
}

function parseCsvSmart_(content) {
  const raw = String(content == null ? '' : content).replace(/^\uFEFF/, '');
  if (!raw.trim()) return [];

  const lines = raw.split(/\r?\n/).slice(0, 10);
  let commaScore = 0;
  let semicolonScore = 0;
  let tabScore = 0;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    commaScore += (line.match(/,/g) || []).length;
    semicolonScore += (line.match(/;/g) || []).length;
    tabScore += (line.match(/\t/g) || []).length;
  }

  let delimiter = ',';
  if (semicolonScore > commaScore && semicolonScore >= tabScore) delimiter = ';';
  if (tabScore > commaScore && tabScore > semicolonScore) delimiter = '\t';

  return Utilities.parseCsv(raw, delimiter);
}

function findColumnIndexByHeader_(headerRow, matcher) {
  for (let i = 0; i < headerRow.length; i++) {
    const normalized = String(headerRow[i] || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
    if (matcher(normalized)) return i;
  }
  return -1;
}

function extractRefundOrderIds_(rows) {
  const ids = new Set();
  if (!rows || rows.length === 0) return ids;

  const header = rows[0] || [];
  const statusByHeader = findColumnIndexByHeader_(header, function(h) {
    return h === 'status' || h.indexOf('delivery status') !== -1 || h.indexOf('order status') !== -1;
  });
  const orderIdByHeader = findColumnIndexByHeader_(header, function(h) {
    return h === 'order id' || h === 'portal order id' || h.indexOf('order id') !== -1 || h.indexOf('orderno') !== -1;
  });

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] || [];

    // Default mapping requested by user: C=Status, D=Order ID.
    let status = String(row[2] || '').toLowerCase().trim();
    let orderId = normalizeOrderId(row[3]);

    // Fallback to header-mapped columns if default columns are empty.
    if ((!status || !orderId) && i > 0) {
      if (statusByHeader !== -1) status = String(row[statusByHeader] || '').toLowerCase().trim();
      if (orderIdByHeader !== -1) orderId = normalizeOrderId(row[orderIdByHeader]);
    }

    if (status.indexOf('refund') !== -1 && orderId) {
      ids.add(orderId);
    }
  }

  return ids;
}

function precheckRefundOrders() {
  const startedAt = new Date();
  logEvent_('INFO', 'Refund precheck started', { folderId: REFUND_CHECK_FOLDER_ID });

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('ReadyToShip');
    if (!sheet) throw new Error('ReadyToShip sheet not found.');

    const folder = DriveApp.getFolderById(REFUND_CHECK_FOLDER_ID);
    const files = folder.getFiles();

    const refundOrderIds = new Set();
    let csvFileCount = 0;

    while (files.hasNext()) {
      const file = files.next();
      const name = String(file.getName() || '').toLowerCase();
      const mime = String(file.getMimeType() || '').toLowerCase();
      const isCsv = name.endsWith('.csv') || mime.indexOf('csv') !== -1 || mime.indexOf('comma-separated-values') !== -1;
      if (!isCsv) continue;

      csvFileCount++;
      let rows = [];
      try {
        const content = file.getBlob().getDataAsString();
        rows = parseCsvSmart_(content);
      } catch (e) {
        logEvent_('WARN', 'Skipping malformed CSV', { fileName: file.getName(), error: String(e) });
        continue;
      }

      const idsFromFile = extractRefundOrderIds_(rows);
      idsFromFile.forEach(function(id) {
        refundOrderIds.add(id);
      });

      logEvent_('INFO', 'Refund scan file summary', {
        fileName: file.getName(),
        rowsRead: rows.length,
        refundIdsFoundInFile: idsFromFile.size
      });
    }

    const lastRow = sheet.getLastRow();
    const lastCol = getEffectiveLastColumn_(sheet);

    let externalMarkingSummary = { externalRowsMarkedRed: 0, checkedRows: 0 };
    try {
      externalMarkingSummary = markExternalSheetRowsRedByOrderIds_(refundOrderIds);
      logEvent_('INFO', 'External sheet marking complete', externalMarkingSummary);
    } catch (markErr) {
      logEvent_('WARN', 'External sheet marking skipped/failed', {
        message: markErr && markErr.message ? markErr.message : String(markErr)
      });
    }

    logEvent_('INFO', 'Refund IDs extracted', {
      csvFilesChecked: csvFileCount,
      refundOrderIdsFound: refundOrderIds.size,
      readyToShipLastRow: lastRow,
      readyToShipLastCol: lastCol,
      externalRowsMarkedRed: externalMarkingSummary.externalRowsMarkedRed || 0
    });

    if (lastRow <= 1) {
      return {
        csvFilesChecked: csvFileCount,
        refundOrderIdsFound: refundOrderIds.size,
        rowsMoved: 0
      };
    }

    const bodyRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    const bodyValues = bodyRange.getValues();
    const orderIdDisplayValues = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();

    const matchedSheetRows = [];

    for (let i = 0; i < bodyValues.length; i++) {
      const orderId = normalizeOrderId(orderIdDisplayValues[i][0]); // ReadyToShip column A (Order ID)
      if (refundOrderIds.has(orderId)) {
        matchedSheetRows.push(i + 2); // absolute row number in sheet
      }
    }

    if (matchedSheetRows.length === 0) {
      logEvent_('INFO', 'Refund precheck complete with no sheet matches', {
        csvFilesChecked: csvFileCount,
        refundOrderIdsFound: refundOrderIds.size,
        rowsMarkedRed: 0,
        durationMs: new Date().getTime() - startedAt.getTime()
      });
      return {
        csvFilesChecked: csvFileCount,
        refundOrderIdsFound: refundOrderIds.size,
        rowsMoved: 0,
        rowsMarkedRed: 0
      };
    }

    // Do not move rows. Mark only the exact matched rows red.
    for (let i = 0; i < matchedSheetRows.length; i++) {
      sheet.getRange(matchedSheetRows[i], 1, 1, lastCol).setBackground('#ff0000');
    }

    const result = {
      csvFilesChecked: csvFileCount,
      refundOrderIdsFound: refundOrderIds.size,
      rowsMoved: 0,
      rowsMarkedRed: matchedSheetRows.length
    };

    logEvent_('INFO', 'Refund precheck complete', {
      csvFilesChecked: csvFileCount,
      refundOrderIdsFound: refundOrderIds.size,
      rowsMoved: 0,
      rowsMarkedRed: matchedSheetRows.length,
      durationMs: new Date().getTime() - startedAt.getTime()
    });
    return result;
  } catch (e) {
    logEvent_('ERROR', 'Refund precheck failed', {
      message: e && e.message ? e.message : String(e),
      stack: e && e.stack ? e.stack : ''
    });
    throw e;
  }
}

function deleteRedRowsFromReadyToShip() {
  logEvent_('INFO', 'Delete red rows started');
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ReadyToShip');
    if (!sheet) throw new Error('ReadyToShip sheet not found.');

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow <= 1 || lastCol === 0) return { deletedRows: 0 };

    // Check only column A background; moved rows are always painted full-row red.
    const backgrounds = sheet.getRange(2, 1, lastRow - 1, 1).getBackgrounds();

    const rowsToDelete = [];
    for (let i = 0; i < backgrounds.length; i++) {
      const c = String(backgrounds[i][0] || '').toLowerCase();
      const isRed = c === '#ff0000' || c === 'red';
      if (isRed) rowsToDelete.push(i + 2);
    }

    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      sheet.deleteRow(rowsToDelete[i]);
    }

    logEvent_('INFO', 'Delete red rows completed', { deletedRows: rowsToDelete.length });
    return { deletedRows: rowsToDelete.length };
  } catch (e) {
    logEvent_('ERROR', 'Delete red rows failed', {
      message: e && e.message ? e.message : String(e),
      stack: e && e.stack ? e.stack : ''
    });
    throw e;
  }
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


// 🔥 MAIN FUNCTION
function generateAllFiles() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ReadyToShip");
  const data = sheet.getDataRange().getValues();

  const today1 = Utilities.formatDate(new Date(), "GMT+5:30", "yyyy-MM-dd");
  const today2 = Utilities.formatDate(new Date(), "GMT+5:30", "yyyy-MM-dd");

  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration");
  if (!configSheet) throw new Error("Configuration sheet not found.");

  const configLastRow = Math.max(configSheet.getLastRow(), 2);
  const configValues = configSheet.getRange(2, 11, configLastRow - 1, 4).getValues(); // K:L:M:N

  const shipglobalCountries = new Set();
  const shipglobalCodeMap = {};
  const shiprocketCountries = new Set();
  const shiprocketCodeMap = {};

  for (let j = 0; j < configValues.length; j++) {
    const [shipglobalCountry, shipglobalCode, shiprocketCountry, shiprocketCode] = configValues[j];
    const normalizedGlobal = String(shipglobalCountry || "").toLowerCase().trim();
    const normalizedRocket = String(shiprocketCountry || "").toLowerCase().trim();

    if (normalizedGlobal) {
      shipglobalCountries.add(normalizedGlobal);
      if (shipglobalCode) shipglobalCodeMap[normalizedGlobal] = String(shipglobalCode).trim();
    }

    if (normalizedRocket) {
      shiprocketCountries.add(normalizedRocket);
      if (shiprocketCode) shiprocketCodeMap[normalizedRocket] = String(shiprocketCode).trim();
    }
  }

  let shiprocketData = [[
    "Order ID","Channel","Order Date","Purpose","Currency","First Name","Last Name",
    "Email","Mobile","Address 1","Address 2","Country","Postcode","City","State",
    "Master SKU","Product Name","HSN Code","Quantity","Tax","VAT","Unit Price",
    "Invoice Date","Length","Breadth","Height","Weight","IOSS","EORI","Terms",
    "Franchise","Seller ID","Courier ID"
  ]];

  let shipglobalData = [[
    "invoice_no","invoice_date","order_reference","service","package_weight",
    "package_length","package_breadth","package_height","currency_code","csb5_status",
    "customer_shipping_firstname","customer_shipping_lastname","customer_shipping_mobile",
    "customer_shipping_email","customer_shipping_company","customer_shipping_address",
    "customer_shipping_address_2","customer_shipping_address_3","customer_shipping_city",
    "customer_shipping_postcode","customer_shipping_country_code","customer_shipping_state",
    "vendor_order_item_name","vendor_order_item_sku","vendor_order_item_quantity",
    "vendor_order_item_unit_price","vendor_order_item_hsn","vendor_order_item_tax_rate",
    "ioss_number","csbv5_limit_comfirmation"
  ]];

  for (let i = 1; i < data.length; i++) {

    const row = data[i];
    if (!row[6]) continue;

    const country = String(row[23] || "").toLowerCase().trim();
    const isShiprocketCountry = shiprocketCountries.has(country);
    const isShipglobalCountry = shipglobalCountries.has(country);

    // Shiprocket
    if (isShiprocketCountry) {
      let r = new Array(33).fill("");

      r[0] = row[6];
      r[1] = "Custom";
      r[2] = "=\"" + today1 + "\"";
      r[3] = "Sample";
      r[4] = "USD";

      // Name Logic
      const fullName = String(row[17] || "").trim();
      if (fullName) {
        const nameParts = fullName.split(/\s+/);
        r[5] = sanitizeText(nameParts[0]); 
        r[6] = nameParts.length > 1 ? sanitizeText(nameParts.slice(1).join(" ")) : ".";
      }

      r[7] = "uttarahomes@gmail.com";

      // --- MOBILE LOGIC (Cleaned & Padded) ---
      let rawPhone = String(row[24] || "").trim();
      if (rawPhone !== "") {
        let cleanPhone = rawPhone.toLowerCase().split("ext")[0].replace(/\D/g, "");
        // Ensure exactly 10 digits
        r[8] = cleanPhone.length > 10 ? cleanPhone.slice(-10) : cleanPhone.padStart(10, '0');
      } else {
        r[8] = ""; // Keep empty if source is empty
      }

      r[9]  = sanitizeText(row[18]);   
      r[10] = sanitizeText(row[19]);   
      r[11] = sanitizeText(row[23]);   
      r[12] = row[22] || "";   
      r[13] = sanitizeText(row[20]);   
      r[14] = sanitizeText(row[21]);   
      r[15] = sanitizeText(row[9]);   
      r[16] = "Fashion Jewellry"; 
      r[17] = "71179010";         
      r[18] = "1";                
      r[21] = "12";               
      r[22] = "=\"" + today1 + "\"";          
      r[23] = "10";               
      r[24] = "3";                
      r[25] = "2";                
      r[26] = "0.05";             
      r[29] = "CIF";              

      shiprocketData.push(r);
    }

    // Shipglobal
    if (isShipglobalCountry) {
      let r = new Array(30).fill("");

      r[0] = row[6];
      r[1] = "=\"" + today2 + "\"";
      r[2] = row[6];
      r[3] = country === "usa" || country === "united states" ? "ShipGlobal Super Saver" : "ShipGlobal Direct";
      r[4] = "0.05";
      r[5] = "10";
      r[6] = "3";
      r[7] = "2";
      r[8] = "USD";

      // Name Logic
      const fullName = String(row[17] || "").trim();
      if (fullName) {
        const nameParts = fullName.split(/\s+/);
        r[10] = sanitizeText(nameParts[0]); 
        r[11] = nameParts.length > 1 ? sanitizeText(nameParts.slice(1).join(" ")) : ".";
      }

      let rawPhone = String(row[24] || "").trim();
      if (rawPhone !== "") {
        let cleanPhone = rawPhone.toLowerCase().split("ext")[0].replace(/\D/g, "");
        r[12] = '="' + (cleanPhone.length > 10 ? cleanPhone.slice(-10) : cleanPhone.padStart(10, '0')) + '"';
      } else {
        r[12] = "";
      }

      r[13] = "uttarahomes@gmail.com";
      r[15] = sanitizeText(row[18]);
      r[16] = sanitizeText(row[19]);
      r[18] = sanitizeText(row[20]);
      r[19] = row[22] || "";
      r[20] = shipglobalCodeMap[country] || "";
      r[21] = sanitizeText(row[21]);
      r[22] = "Jewellry Fashion";
      r[23] = row[9];
      r[24] = "1";
      r[25] = "6";
      r[26] = "71179010";

      shipglobalData.push(r);
    }
  }

  let shiprocketCsv = "";
  let shipglobalCsv = "";

  if (shiprocketData.length > 1) {
    const rocketCsv = shiprocketData.map(r => r.map(c => `"${String(c).replace(/"/g,'""')}"`).join(",")).join("\n");
    shiprocketCsv = "data:text/csv;charset=utf-8," + encodeURIComponent(rocketCsv);
  }

  if (shipglobalData.length > 1) {
    const globalCsv = shipglobalData.map(r => r.map(c => `"${String(c).replace(/"/g,'""')}"`).join(",")).join("\n");
    shipglobalCsv = "data:text/csv;charset=utf-8," + encodeURIComponent(globalCsv);
  }

  return {
    shiprocketCsv: shiprocketCsv,
    shipglobalCsv: shipglobalCsv
  };
}