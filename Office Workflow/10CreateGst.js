function Create_Gst() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get external spreadsheet ID from configuration sheet
  const configSheet = ss.getSheetByName("configuration");
  if (!configSheet) throw new Error("configuration sheet not found.");

  const externalSpreadsheetId = configSheet.getRange("B1").getValue();
  if (!externalSpreadsheetId) throw new Error("No spreadsheet ID found in configuration!B1");

  // Open external spreadsheet
  const externalSS = SpreadsheetApp.openById(externalSpreadsheetId);

  // Get All Orders from external spreadsheet
  const sourceSheet = externalSS.getSheetByName("All Orders");
  if (!sourceSheet) throw new Error("All Orders sheet not found in external spreadsheet.");

  // If DATASET or EXPORT DATA sheets are missing, prompt user to create them and rerun
  const missingSheets = [];
  if (!ss.getSheetByName('DATASET')) missingSheets.push('DATASET');
  if (!ss.getSheetByName('EXPORT DATA')) missingSheets.push('EXPORT DATA');
  if (missingSheets.length) {
    try {
      const ui = SpreadsheetApp.getUi();
      const msg = 'The spreadsheet is missing the following required sheet(s): ' + missingSheets.join(', ') + 
                  '.\nWould you like the script to create them now? After creation please rerun the tool.';
      const resp = ui.alert('Missing Sheets', msg, ui.ButtonSet.YES_NO);
      if (resp === ui.Button.YES) {
        if (missingSheets.indexOf('DATASET') !== -1) {
          const ds = ss.insertSheet('DATASET');
          try { ds.getRange(1,1,1,4).setValues([['ItemName','ItemType','HSN Code','GST%']]); } catch(e) { /* ignore */ }
        }
        if (missingSheets.indexOf('EXPORT DATA') !== -1) {
          const ex = ss.insertSheet('EXPORT DATA');
          try { ex.getRange(1,1,1,6).setValues([['Brand','Currency','Rate','Net','State','Date']]); } catch(e) { /* ignore */ }
        }
        ui.alert('Sheets created', 'Missing sheets created. Please rerun the "Create 3b GST" menu item to continue.');
      } else {
        ui.alert('Action required', 'Please add the missing sheet(s) (' + missingSheets.join(', ') + ') and rerun the tool.');
      }
    } catch (err) {
      Logger.log('Missing sheets and UI not available: ' + err);
      throw new Error('Required sheets missing: ' + missingSheets.join(', ') + '. Please add them and rerun.');
    }
    return;
  }

  const lastRow = sourceSheet.getLastRow();
  const lastCol = sourceSheet.getLastColumn();
  if (lastRow < 3) return;

  const headers = sourceSheet.getRange(2, 1, 1, lastCol).getValues()[0];
  const data = sourceSheet.getRange(3, 1, lastRow - 2, lastCol).getValues();

  const purchaseDateIndex = findColumn(headers, "PURCHASE_DATE");
  if (purchaseDateIndex === -1) throw new Error("PURCHASE_DATE column not found.");

  // collect available months from data to show in a dropdown
  const monthsSet = {};
  if (data && data.length) {
    data.forEach(row => {
      const purchaseDate = row[purchaseDateIndex];
      if (!purchaseDate) return;
      let month = null;
      if (purchaseDate instanceof Date && !isNaN(purchaseDate)) {
        month = purchaseDate.getMonth() + 1;
      } else {
        const parts = purchaseDate.toString().split(/[-\/\.]/);
        if (parts.length >= 2) {
          const first = parseInt(parts[0], 10);
          const second = parseInt(parts[1], 10);
          month = first > 12 ? second : first;
        }
      }
      if (month) monthsSet[month] = true;
    });
  }

  const availableMonths = Object.keys(monthsSet).map(m => parseInt(m, 10)).sort((a,b)=>a-b);
  if (availableMonths.length === 0) {
    SpreadsheetApp.getUi().alert('No months found in Purchase Date column.');
    return;
  }

  // Build a simple HTML dialog with a dropdown for months (show names, values are numbers)
  const monthNames = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  const optionsHtml = availableMonths.map(m => '<option value="' + m + '">' + monthNames[m-1] + ' (' + m + ')' + '</option>').join('');
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family: Arial; padding:18px; width:100%; box-sizing:border-box;">' +
    '<h2 style="margin:0 0 8px 0">Create a 3b GST for</h2>' +
    '<div>' +
      '<div style="margin-bottom:8px">Select month:</div>' +
      '<select id="monthSelect" style="font-size:14px; padding:6px; width:100%">' + optionsHtml + '</select>' +
      '<div style="margin-top:14px; text-align:right">' +
        '<button id="createBtn" onclick="createNow()" style="padding:6px 12px; margin-right:8px">Create</button>' +
        '<span id="spinner" style="display:none; vertical-align:middle; margin-right:8px">' +
          '<span style="display:inline-block;width:18px;height:18px;border:2px solid #ccc;border-top-color:#333;border-radius:50%;animation:spin 1s linear infinite"></span>' +
        '</span>' +
        '<button onclick="google.script.host.close()" style="padding:6px 12px">Cancel</button>' +
      '</div>' +
    '</div>' +
    '<style>@keyframes spin{to{transform:rotate(360deg)}}</style>' +
    '<script>' +
    'function createNow(){' +
    '  var btn = document.getElementById("createBtn");' +
    '  var sel = document.getElementById("monthSelect");' +
    '  var spinner = document.getElementById("spinner");' +
    '  btn.disabled = true; sel.disabled = true; spinner.style.display = "inline-block"; btn.textContent = "Creating...";' +
    '  google.script.run.withSuccessHandler(function(){ google.script.host.close(); }).withFailureHandler(function(err){ spinner.style.display = "none"; btn.disabled = false; sel.disabled = false; btn.textContent = "Create"; alert("Error: " + (err && err.message ? err.message : err)); }).processCopyForMonth(parseInt(sel.value,10));' +
    '}' +
    '</script>' +
    '</div>'
  ).setWidth(620).setHeight(240);

  try {
    SpreadsheetApp.getUi().showModalDialog(html, 'Create a 3b GST for');
  } catch (uiErr) {
    Logger.log('Cannot show UI dialog: ' + uiErr);
    throw new Error('Cannot open dialog. Run this function from the spreadsheet UI (open the spreadsheet and use the custom menu)');
  }
}


function processCopyForMonth(targetMonth, outputName) {
  const targetCountry = "India";
  const baseInvoiceNumber = "BE\\24-25\\48264";
  const outputFolderId = "1WBvQXpbAPTrQIPTEnsYPEd9ezeUJ5z9q"; // destination Drive folder

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get external spreadsheet ID from configuration sheet
  const configSheet = ss.getSheetByName("configuration");
  if (!configSheet) throw new Error("configuration sheet not found.");

  const externalSpreadsheetId = configSheet.getRange("B1").getValue();
  if (!externalSpreadsheetId) throw new Error("No spreadsheet ID found in configuration!B1");

  // Open external spreadsheet
  const externalSS = SpreadsheetApp.openById(externalSpreadsheetId);

  // Get All Orders from external spreadsheet
  const sourceSheet = externalSS.getSheetByName("All Orders");
  if (!sourceSheet) throw new Error("All Orders sheet not found in external spreadsheet.");

  const lastRow = sourceSheet.getLastRow();
  const lastCol = sourceSheet.getLastColumn();
  if (lastRow < 3) return;

  const headers = sourceSheet.getRange(2, 1, 1, lastCol).getValues()[0];
  const data = sourceSheet.getRange(3, 1, lastRow - 2, lastCol).getValues();

  // Read DATASET sheet (if present) — detect header row robustly (handles inserted rows)
  const datasetSheet = ss.getSheetByName("DATASET");
  let datasetHeaders = [];
  let datasetData = [];
  if (datasetSheet) {
    const dsLastRow = datasetSheet.getLastRow();
    const dsLastCol = datasetSheet.getLastColumn();
    if (dsLastRow >= 2 && dsLastCol >= 1) {
      // attempt to find the header row by scanning the first few rows for known header tokens
      let headerRowIndex = 1;
      const maxHeaderScan = Math.min(10, dsLastRow);
      for (let r = 1; r <= maxHeaderScan; r++) {
        const rowVals = datasetSheet.getRange(r, 1, 1, dsLastCol).getValues()[0];
        const norm = rowVals.map(c => (c || "").toString().trim().toLowerCase());
        if (norm.some(c => c.indexOf('itemname') !== -1 || c.indexOf('item name') !== -1 || c.indexOf('itemtype') !== -1 || c.indexOf('hsn') !== -1 || c.indexOf('gst') !== -1)) {
          headerRowIndex = r;
          break;
        }
      }
      datasetHeaders = datasetSheet.getRange(headerRowIndex, 1, 1, dsLastCol).getValues()[0];
      const numDataRows = dsLastRow - headerRowIndex;
      if (numDataRows > 0) {
        datasetData = datasetSheet.getRange(headerRowIndex + 1, 1, numDataRows, dsLastCol).getValues();
      }
      Logger.log('Dataset header row: ' + headerRowIndex + ', data rows: ' + (datasetData ? datasetData.length : 0));
    }
  }

  const purchaseDateIndex = findColumn(headers, "PURCHASE_DATE");
  const brandIndex = findColumn(headers, "BRAND_NAME");
  const salesChannelIndex = findColumn(headers, "SALES_CHANNEL");
  const stateIndex = findColumn(headers, "STATE");
  const pincodeIndex = findColumn(headers, "PINCODE");
  const priceIndex = findColumn(headers, "CURRENCY_PRICE");
  const countryIndex = findColumn(headers, "COUNTRY");
  const orderIdIndex = findColumn(headers, "PORTAL_ORDER_ID");
  let itemNameIndex = findColumn(headers, "ITEM_NAME");
  let hsnIndex = findColumn(headers, "HSN_CODE");
  // Note: findColumn is now robust and handles common variants; log resolved headers
  Logger.log('Resolved columns — Item Name index: ' + itemNameIndex + (itemNameIndex >= 0 ? ' ("' + headers[itemNameIndex] + '")' : '') + ', HSN index: ' + hsnIndex + (hsnIndex >= 0 ? ' ("' + headers[hsnIndex] + '")' : ''));

  if ([purchaseDateIndex, brandIndex, salesChannelIndex, stateIndex,
       pincodeIndex, priceIndex, countryIndex, orderIdIndex].includes(-1)) {
    throw new Error("Required columns missing.");
  }

  const invoiceMatch = baseInvoiceNumber.match(/^(.*\\)(\d+)$/);
  if (!invoiceMatch) throw new Error("Invalid invoice format.");
  const invoicePrefix = invoiceMatch[1];
  let invoiceCounter = parseInt(invoiceMatch[2], 10);

  const normalizedTargetCountry = normalize(targetCountry);

  const headersOut = [
    "Date","Order ID","Invoice No.","Party Name","Order Status","GSTIN/UIN",
    "Ledger Name","State Name","Pincode","Item Name","HSN Code",
    "Currency","conversion Rate","Invoice Value","Taxable amount",
    "IGST","CGST","SGST","Tcs Rate","Tcs Amount"
  ];

  // Determine dataset column indices (if dataset present)
  let datasetItemNameIndex = -1;
  let datasetItemTypeIndex = -1;
  let datasetHsnIndex = -1;
  let datasetGstIndex = -1;
  if (datasetHeaders && datasetHeaders.length) {
    datasetItemNameIndex = findColumn(datasetHeaders, 'ITEMNAME');
    if (datasetItemNameIndex === -1) datasetItemNameIndex = findColumn(datasetHeaders, 'Item Name');
    datasetItemTypeIndex = findColumn(datasetHeaders, 'ItemType');
    if (datasetItemTypeIndex === -1) datasetItemTypeIndex = findColumn(datasetHeaders, 'item type');
    datasetHsnIndex = findColumn(datasetHeaders, 'HSN Code');
    if (datasetHsnIndex === -1) datasetHsnIndex = findColumn(datasetHeaders, 'hsn');
    datasetGstIndex = findColumn(datasetHeaders, 'GST%');
    if (datasetGstIndex === -1) datasetGstIndex = findColumn(datasetHeaders, 'GST');
    if (datasetGstIndex === -1) datasetGstIndex = findColumn(datasetHeaders, 'gst');
    Logger.log('Dataset columns — ITEMNAME: ' + datasetItemNameIndex + (datasetItemNameIndex >= 0 ? ' ("' + datasetHeaders[datasetItemNameIndex] + '")' : '') + ', ITEMTYPE: ' + datasetItemTypeIndex + ', HSN: ' + datasetHsnIndex + ', GST: ' + datasetGstIndex);
  }

  // --- DETECT EXPORT DATA SHEET (used to populate per-brand EXPORT sheet) ---
  const exportSheet = ss.getSheetByName("EXPORT DATA");
  let exportHeaders = [];
  let exportData = [];
  let exportBrandIndex = -1, exportCurrencyIndex = -1, exportRateIndex = -1, exportNetIndex = -1, exportStateIndex = -1, exportDateIndex = -1;
  if (exportSheet) {
    const exLastRow = exportSheet.getLastRow();
    const exLastCol = exportSheet.getLastColumn();
    if (exLastRow >= 1 && exLastCol >= 1) {
      // try to find header row within first 10 rows
      let headerRowIndex = 1;
      const maxScan = Math.min(10, exLastRow);
      for (let r = 1; r <= maxScan; r++) {
        const rowVals = exportSheet.getRange(r, 1, 1, exLastCol).getValues()[0];
        const norm = rowVals.map(c => (c || "").toString().trim().toLowerCase());
        if (norm.some(c => c.indexOf('brand') !== -1 || c.indexOf('currency') !== -1 || c.indexOf('rate') !== -1 || c.indexOf('net') !== -1)) {
          headerRowIndex = r;
          break;
        }
      }
      exportHeaders = exportSheet.getRange(headerRowIndex, 1, 1, exLastCol).getValues()[0];
      const numExportRows = exLastRow - headerRowIndex;
      if (numExportRows > 0) exportData = exportSheet.getRange(headerRowIndex + 1, 1, numExportRows, exLastCol).getValues();

      exportBrandIndex = findColumn(exportHeaders, 'Brand');
      exportCurrencyIndex = findColumn(exportHeaders, 'Currency');
      exportRateIndex = findColumn(exportHeaders, 'Rate');
      exportNetIndex = findColumn(exportHeaders, 'Net');
      exportStateIndex = findColumn(exportHeaders, 'State');
      exportDateIndex = findColumn(exportHeaders, 'Date');

      Logger.log('Export data header row: ' + (exportSheet ? 'found' : 'none') + ', rows: ' + (exportData ? exportData.length : 0));
    }
  }

  // build brand map for selected month
  const brandMap = {};
  data.forEach(row => {
    const purchaseDate = row[purchaseDateIndex];
    const brandRaw = row[brandIndex];
    if (!purchaseDate || !brandRaw) return;

    let month = null;
    if (purchaseDate instanceof Date && !isNaN(purchaseDate)) month = purchaseDate.getMonth() + 1;
    else {
      const parts = purchaseDate.toString().split(/[-\/\.]/);
      if (parts.length >= 2) {
        const first = parseInt(parts[0], 10);
        const second = parseInt(parts[1], 10);
        month = first > 12 ? second : first;
      }
    }
    if (month !== targetMonth) return;
    if (targetCountry !== "") {
      const countryValue = row[countryIndex];
      if (!countryValue) return;
      if (normalize(countryValue) !== normalizedTargetCountry) return;
    }
    const brandValue = brandRaw.toString().trim();
    const key = normalize(brandValue);
    if (key) brandMap[key] = brandValue;
  });

  const brandsToProcess = Object.keys(brandMap);

  let totalRowsWritten = 0;
  const perBrandCounts = [];
  let totalBrandsWithData = 0;
  const createdFiles = []; // collect created file names and URLs

  // determine spreadsheet name template
  const monthNames = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  const mn = monthNames[targetMonth-1] || ("Month " + targetMonth);
  // if user provided an outputName template, use it; otherwise use default template with {brand}
  const template = (outputName && outputName.toString().trim() !== "") ? outputName.toString() : (mn + ' GSTR 3B: {brand} 25-26');

  // Create (or pick) a month-named subfolder inside the configured output folder.
  // If the folder exists, create a unique name like "January-1", "January-2" etc.
  const parentOutFolder = DriveApp.getFolderById(outputFolderId);
  // Determine appropriate year for the chosen month from available purchase dates
  let yearForFolder = (new Date()).getFullYear();
  for (let i = 0; i < data.length; i++) {
    const pd = data[i][purchaseDateIndex];
    if (!pd) continue;
    let m = null;
    let y = null;
    if (pd instanceof Date && !isNaN(pd)) {
      m = pd.getMonth() + 1;
      y = pd.getFullYear();
    } else {
      const parts = pd.toString().split(/[-\/.]/).map(s => s.trim()).filter(Boolean);
      if (parts.length >= 2) {
        const first = parseInt(parts[0], 10);
        const second = parseInt(parts[1], 10);
        if (!isNaN(first) && !isNaN(second)) {
          m = first > 12 ? second : first;
        }
      }
      // try to find a 4-digit year token
      for (let p = parts.length - 1; p >= 0; p--) {
        const n = parseInt(parts[p], 10);
        if (!isNaN(n) && n > 31) { y = n; break; }
      }
    }
    if (m === targetMonth && y) { yearForFolder = y; break; }
  }

  let subfolderName = mn + ' ' + yearForFolder;
  let destFolder = null;
  // If folder exists, append numeric suffix like "January 2026-1", "January 2026-2"
  if (parentOutFolder.getFoldersByName(subfolderName).hasNext()) {
    let idx = 1;
    let candidate = subfolderName + '-' + idx;
    while (parentOutFolder.getFoldersByName(candidate).hasNext()) {
      idx++;
      candidate = subfolderName + '-' + idx;
    }
    destFolder = parentOutFolder.createFolder(candidate);
    subfolderName = candidate;
  } else {
    destFolder = parentOutFolder.createFolder(subfolderName);
  }

  brandsToProcess.forEach(normalizedBrandKey => {
    const brandDisplay = brandMap[normalizedBrandKey];
    if (!brandDisplay) return;

    let sheetName = brandDisplay.replace(/[\\\/\?\*\[\]:]/g, "").substring(0, 80);
    sheetName = sheetName.replace(/\b\w/g, l => l.toUpperCase());

    // create a separate spreadsheet for this brand using template (replace {brand})
    let finalBrandFileName = template;
    if (/{brand}/i.test(finalBrandFileName)) {
      finalBrandFileName = finalBrandFileName.replace(/\{brand\}/ig, sheetName);
    } else if (/{Brand}/.test(finalBrandFileName)) {
      finalBrandFileName = finalBrandFileName.replace(/\{Brand\}/g, sheetName);
    } else {
      // fallback to default naming if template doesn't include brand placeholder
      finalBrandFileName = mn + ' GSTR 3B: ' + sheetName + ' 25-26';
    }
    const brandSs = SpreadsheetApp.create(finalBrandFileName);
    const brandFile = DriveApp.getFileById(brandSs.getId());
    // add the created brand spreadsheet into the month subfolder
    destFolder.addFile(brandFile);
    DriveApp.getRootFolder().removeFile(brandFile);
    // store created file info for confirmation
    try { createdFiles.push({ name: finalBrandFileName, url: brandFile.getUrl() }); } catch (e) { createdFiles.push({ name: finalBrandFileName, url: '' }); }

    const outputRows = [];

    data.forEach(row => {
      const brandRaw = row[brandIndex];
      if (!brandRaw) return;
      if (normalize(brandRaw.toString().trim()) !== normalizedBrandKey) return;

      const purchaseDate = row[purchaseDateIndex];
      if (!purchaseDate) return;

      let month = null;
      if (purchaseDate instanceof Date && !isNaN(purchaseDate)) {
        month = purchaseDate.getMonth() + 1;
      } else {
        const parts = purchaseDate.toString().split(/[-\/\.]/);
        if (parts.length >= 2) {
          const first = parseInt(parts[0], 10);
          const second = parseInt(parts[1], 10);
          month = first > 12 ? second : first;
        }
      }
      if (month !== targetMonth) return;

      if (targetCountry !== "") {
        const countryValue = row[countryIndex];
        if (!countryValue) return;
        if (normalize(countryValue) !== normalizedTargetCountry) return;
      }

      let partyName = "";
      const salesChannel = row[salesChannelIndex];
      if (salesChannel) partyName = salesChannel.toString().split(".")[0].trim();

      const stateRaw = (row[stateIndex] || "").toString().trim();
      const countryRaw = (row[countryIndex] || "").toString().trim();
      const countryNorm = normalize(countryRaw);
      const stateNorm = stateRaw.toLowerCase();
      const ledgerName = (countryNorm === 'india' && (stateNorm === 'uttarakhand' || stateNorm === 'uk'))
        ? "GST INTRASTATE SALES"
        : "GST INTERSTATE SALES";

      let invoiceValue = row[priceIndex] || "";
      if (typeof invoiceValue === "string") invoiceValue = invoiceValue.replace(/,/g, "");

      const outRow = [
        purchaseDate,
        "",
        // invoicePrefix + invoiceCounter++,
        "",
        partyName,
        "",
        "",
        ledgerName,
        stateRaw,
        row[pincodeIndex] || "",
        (itemNameIndex !== -1 ? (row[itemNameIndex] || "") : ""),
        (hsnIndex !== -1 ? (row[hsnIndex] || "") : ""),
        "INR",
        "1.00",
        invoiceValue,
        "",
        "",
        "",
        "",
        0.01,
        ""
      ];
      // Post-process item name using DATASET mapping rules (will be applied later per-brand)
      outputRows.push(outRow);
    });

    if (outputRows.length > 0) {
      // Apply DATASET mapping to outputRows if dataset is available
      if (datasetData && datasetData.length && datasetItemNameIndex !== -1) {
        for (let rr = 0; rr < outputRows.length; rr++) {
          const currentItem = (outputRows[rr][9] || "").toString().trim();
          if (!currentItem) continue;
          const curLower = currentItem.toLowerCase();

          // 1) Try exact matches first (case-insensitive)
          const exactMatches = [];
          for (let d = 0; d < datasetData.length; d++) {
            const dsItem = (datasetData[d][datasetItemNameIndex] || "").toString().trim();
            if (!dsItem) continue;
            if (dsItem.toLowerCase() === curLower) exactMatches.push(d);
          }

          let matchedIndex = -1;
          if (exactMatches.length === 1) matchedIndex = exactMatches[0];
          else if (exactMatches.length > 1) {
            // Multiple exact matches found — keep original item name
            continue;
          } else {
            // 2) Token-scoring match: prefer dataset entries whose tokens best match the item
            const curTokens = curLower.replace(/[^a-z0-9]+/g, ' ').trim().split(/\s+/).filter(Boolean).map(t => (t.length>3 && t.endsWith('s')) ? t.slice(0,-1) : t);
            let bestIdx = -1;
            let bestScore = 0;
            let bestTokenCount = 0;
            for (let d = 0; d < datasetData.length; d++) {
              const dsItemRaw = (datasetData[d][datasetItemNameIndex] || "").toString().trim().toLowerCase();
              if (!dsItemRaw) continue;
              const dsTokens = dsItemRaw.replace(/[^a-z0-9]+/g, ' ').trim().split(/\s+/).filter(Boolean).map(t => (t.length>3 && t.endsWith('s')) ? t.slice(0,-1) : t);
              if (dsTokens.length === 0) continue;
              // count matched tokens
              let matchCount = 0;
              for (let ti = 0; ti < dsTokens.length; ti++) {
                const tok = dsTokens[ti];
                if (curTokens.indexOf(tok) !== -1) matchCount++;
              }
              const score = matchCount / dsTokens.length; // fraction of dataset tokens present in item
              if (score > bestScore || (score === bestScore && dsTokens.length > bestTokenCount)) {
                bestScore = score;
                bestIdx = d;
                bestTokenCount = dsTokens.length;
              }
            }
            // accept only if a reasonable fraction of tokens matched
            if (bestIdx !== -1 && bestScore >= 0.5) {
              matchedIndex = bestIdx;
            } else {
              // no confident match — keep original item name
              continue;
            }
          }

          if (matchedIndex !== -1) {
            const dsRow = datasetData[matchedIndex];
            const itemTypeVal = (datasetItemTypeIndex !== -1) ? (dsRow[datasetItemTypeIndex] || "") : "";
            const hsnVal = (datasetHsnIndex !== -1) ? (dsRow[datasetHsnIndex] || "") : "";
            // Replace complete item name with GEMSTONE (item type) as requested
            if (itemTypeVal && itemTypeVal.toString().trim() !== "") outputRows[rr][9] = itemTypeVal.toString();
            // Fill HSN column if available
            if (hsnVal && hsnVal.toString().trim() !== "") outputRows[rr][10] = hsnVal.toString();
            
            // TAX CALCULATIONS
            const gstPercentage = (datasetGstIndex !== -1) ? parseFloat(dsRow[datasetGstIndex] || 0) : 0;
            const invoiceValue = parseFloat(outputRows[rr][13] || 0) || 0;
            const tcsRate = parseFloat(outputRows[rr][18] || 0) || 0;
            const ledgerName = (outputRows[rr][6] || "").toString().trim().toUpperCase();
            
            // Calculate CGST, SGST, IGST based on ledger type
            let igstVal = 0, cgstVal = 0, sgstVal = 0;
            if (ledgerName.includes("INTRASTATE")) {
              // Intrastate: CGST = SGST = (InvoiceValue × GST%/2) / 100; IGST = 0
              cgstVal = (invoiceValue * gstPercentage / 2) / 100;
              sgstVal = cgstVal;
              igstVal = 0;
              outputRows[rr][15] = 0;           // IGST = 0
              outputRows[rr][16] = cgstVal;     // CGST
              outputRows[rr][17] = sgstVal;     // SGST
            } else {
              // Interstate: IGST = (InvoiceValue × GST%) / 100; CGST = SGST = 0
              igstVal = (invoiceValue * gstPercentage) / 100;
              cgstVal = 0;
              sgstVal = 0;
              outputRows[rr][15] = igstVal;     // IGST
              outputRows[rr][16] = 0;           // CGST = 0
              outputRows[rr][17] = 0;           // SGST = 0
            }
            
            // Taxable Amount = (IGST + CGST + SGST) + Invoice Value
            const taxableAmount = invoiceValue-(igstVal + cgstVal + sgstVal);
            outputRows[rr][14] = taxableAmount;
            
            // Calculate TCS Amount = Tcs Rate × Taxable Amount (no division)
            outputRows[rr][19] = tcsRate * taxableAmount;
          }
        }
      }
      // write into the brand spreadsheet's first sheet
      let sheet = brandSs.getSheets()[0];
      try { sheet.setName(sheetName); } catch (e) { /* ignore rename errors */ }
      sheet.clear();

      sheet.getRange(1, 1, 1, headersOut.length)
        .setValues([headersOut])
        .setFontWeight("bold")
        .setBackground("#FFFF00")
        .setHorizontalAlignment("center");

      sheet.getRange(2, 1, outputRows.length, headersOut.length).setValues(outputRows);
      // Format: center all, add borders, and auto-resize columns
      const totalRows = outputRows.length + 1; // header + data
      const dataRange = sheet.getRange(1, 1, totalRows, headersOut.length);
      dataRange.setHorizontalAlignment("center");
      dataRange.setBorder(true, true, true, true, true, true);
      for (let c = 1; c <= headersOut.length; c++) {
        try { sheet.autoResizeColumn(c); } catch (e) { /* ignore if not supported */ }
      }

      totalRowsWritten += outputRows.length;
      totalBrandsWithData++;
      perBrandCounts.push({ name: finalBrandFileName, count: outputRows.length });
    }
      // --- Create and populate EXPORT sheet in the brand spreadsheet from EXPORT DATA ---
      try {
        if (exportData && exportData.length && exportBrandIndex !== -1 && exportCurrencyIndex !== -1 && exportRateIndex !== -1 && exportNetIndex !== -1) {
          // remove existing EXPORT sheet if present
          try {
            const existing = brandSs.getSheetByName('EXPORT');
            if (existing) brandSs.deleteSheet(existing);
          } catch (e) { /* ignore */ }

          const exportSheetOut = brandSs.insertSheet('EXPORT');
          const exportHeadersOut = ["Date","Invoice Number","PartyName","Ledger Name","State Name","Item Name","Currency","Price","conversion Rate","INR TOTAL Price","Taxable Amount","IGST","CGST","SGST"];
          exportSheetOut.getRange(1,1,1,exportHeadersOut.length).setValues([exportHeadersOut]).setFontWeight('bold');

          // compute last date of chosen month (use current year)
          const year = new Date().getFullYear();
          const lastDay = new Date(year, targetMonth, 0);

          // party name mapping by currency
          const partyMap = {
            'CAD': 'AMAZON_CA',
            'GBP': 'AMAZON_UK',
            'USD': 'AMAZON_US',
            'EUR': 'AMAZON_DE',
            'JPY': 'AMAZON_JP',
            'AUD': 'AMAZON_AU',
            'MXN': 'AMAZON_MX',
            'BRL': 'AMAZON_BR'
          };

          const exportRowsOut = [];
          for (let e = 0; e < exportData.length; e++) {
            const erow = exportData[e];
            const brandVal = (erow[exportBrandIndex] || '').toString().trim();
            if (!brandVal) continue;
            if (normalize(brandVal) !== normalizedBrandKey) continue;

            // if export has a date column, filter by chosen month
            if (exportDateIndex !== -1) {
              const v = erow[exportDateIndex];
              if (v) {
                let m = null;
                if (v instanceof Date && !isNaN(v)) m = v.getMonth() + 1;
                else {
                  const parts = v.toString().split(/[-\/\.]/);
                  if (parts.length >= 2) {
                    const first = parseInt(parts[0],10);
                    const second = parseInt(parts[1],10);
                    m = first > 12 ? second : first;
                  }
                }
                if (m !== targetMonth) continue;
              }
            }

            const currency = (erow[exportCurrencyIndex] || '').toString().trim();
            const rate = parseFloat((erow[exportRateIndex] || 0).toString().replace(/,/g, '')) || 1;
            const inrTotal = parseFloat((erow[exportNetIndex] || 0).toString().replace(/,/g, '')) || 0;
            const price = rate === 0 ? 0 : (inrTotal / rate);

            // determine item name by price
            let itemNameCalc = '';
            if (price < 50) itemNameCalc = 'JEWELLERY';
            else if (price > 50 && price < 200) itemNameCalc = 'ROUGH_STONE';
            else if (price > 200) itemNameCalc = 'PUJA ARTICALS';

            // IGST calculation based on item type
            let igstVal = 0;
            if (itemNameCalc === 'JEWELLERY') igstVal = inrTotal * 0.03;
            else if (itemNameCalc === 'ROUGH_STONE') igstVal = inrTotal * 0.0025;
            else igstVal = 0;

            const taxable = inrTotal - igstVal;
            const party = partyMap[currency] || '';
            const stateName = (exportStateIndex !== -1 ? (erow[exportStateIndex] || '') : '');

            const rowOut = [
              lastDay,
              // invoicePrefix + invoiceCounter++,
              "",
              party,
              'GST EXPORT',
              stateName,
              itemNameCalc,
              currency,
              parseFloat(price.toFixed(2)),
              parseFloat(rate.toFixed(4)),
              parseFloat(inrTotal.toFixed(2)),
              parseFloat(taxable.toFixed(2)),
              parseFloat(igstVal.toFixed(2)),
              0,
              0
            ];
            exportRowsOut.push(rowOut);
          }

          if (exportRowsOut.length) {
            exportSheetOut.getRange(2,1,exportRowsOut.length, exportHeadersOut.length).setValues(exportRowsOut);
            exportSheetOut.getRange(1,1,exportRowsOut.length+1, exportHeadersOut.length).setHorizontalAlignment('center');
            for (let c = 1; c <= exportHeadersOut.length; c++) {
              try { exportSheetOut.autoResizeColumn(c); } catch (e) { }
            }
          }
        }
      } catch (ee) { Logger.log('Error populating EXPORT sheet: ' + ee); }
  });

  Logger.log("Processed brands with data: " + totalBrandsWithData + ", total rows: " + totalRowsWritten);
  perBrandCounts.forEach(b => Logger.log("Brand: " + b.name + " — rows: " + b.count));

  // Show confirmation dialog with links to created files
  if (createdFiles.length > 0) {
    const ui = SpreadsheetApp.getUi();
    const listHtml = createdFiles.map(f => {
      const url = f.url ? f.url : '#';
      const safeName = f.name.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
      return '<div style="margin-bottom:6px"><a href="' + url + '" target="_blank">' + safeName + '</a></div>';
    }).join('');

    const html = HtmlService.createHtmlOutput('<div style="font-family:Arial;padding:12px;max-height:400px;overflow:auto">' +
      '<h3 style="margin-top:0">Created Spreadsheets</h3>' + listHtml +
      '<div style="margin-top:12px;text-align:right"><button onclick="google.script.host.close()">Close</button></div>' +
      '<script>function open(url){window.open(url, "_blank");}</script>' +
      '</div>').setWidth(520).setHeight(300);

    ui.showModalDialog(html, 'Files created');
  }
}


// 🔧 HELPERS
function findColumn(headers, name) {
  if (!headers || !headers.length) return -1;
  const target = name.toString().trim().toLowerCase();

  // 1) Exact header match (trimmed, case-insensitive)
  for (let i = 0; i < headers.length; i++) {
    if (headers[i].toString().trim().toLowerCase() === target) return i;
  }

  // 2) Normalized exact match (remove non-alphanumeric characters)
  const normTarget = target.replace(/[^a-z0-9]/g, '');
  for (let i = 0; i < headers.length; i++) {
    const h = headers[i].toString().trim().toLowerCase();
    const normH = h.replace(/[^a-z0-9]/g, '');
    if (normH === normTarget && normH !== '') return i;
  }

  // 3) Token-based match: require all tokens from the target to appear in header text
  const tokens = (normTarget.match(/[a-z0-9]+/g) || []);
  if (tokens.length) {
    for (let i = 0; i < headers.length; i++) {
      const h = headers[i].toString().trim().toLowerCase();
      let all = true;
      for (let t = 0; t < tokens.length; t++) {
        const tok = tokens[t];
        if (h.indexOf(tok) === -1 && h.replace(/[^a-z0-9]/g, '').indexOf(tok) === -1) { all = false; break; }
      }
      if (all) return i;
    }
  }

  // 4) No match found
  return -1;
}

function normalize(text) {
  return text.toString().trim().replace(/\s+/g, " ").toLowerCase();
}
