/**
 * Opens the Date Picker Dialog for Jaipur Data Transfer.
 */
function launchJaipurTransfer() {
  const html = HtmlService.createHtmlOutputFromFile('3JaipurPicker')
      .setWidth(400)
      .setHeight(300)
      .setTitle('Transfer Jaipur Data');
  SpreadsheetApp.getUi().showModalDialog(html, 'Select Date Range');
}

/**
 * Processes the transfer based on the selected date range.
 * Now points to INTERNAL sheets to avoid "Target Access" errors.
 */
/**
 * Processes the transfer using:
 * Source: B1 (ID), B2 (Name)
 * Target: B4 (ID), B5 (Name)
 */
function processJaipurTransfer(dateRange) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. VERIFY INTERNAL CONFIGURATION
  const configSheet = ss.getSheetByName("Configuration");
  const referenceSheet = ss.getSheetByName("Seema Jain Data");

  if (!configSheet) throw new Error("Sheet 'Configuration' not found.");
  if (!referenceSheet) throw new Error("Sheet 'Seema Jain Data' not found.");

  // 2. GET CONFIG VALUES USING DISPLAY VALUE (Prevents Date errors)
  const sourceId = configSheet.getRange("B1").getDisplayValue().trim();
  const sourceSheetName = configSheet.getRange("B2").getDisplayValue().trim();
  
  const targetId = configSheet.getRange("B4").getDisplayValue().trim();
  const targetSheetName = configSheet.getRange("B5").getDisplayValue().trim();

  if (!targetId || !targetSheetName) {
    throw new Error("Target ID (B4) or Target Name (B5) is empty in Configuration.");
  }

  // 3. LOAD REFERENCE KEYWORDS (Seema Jain Data)
  const refData = referenceSheet.getRange("A1:A" + referenceSheet.getLastRow()).getValues();
  const keywords = refData.flat().filter(String);

  // 4. OPEN SOURCE SPREADSHEET (External)
  let sourceSheet;
  try {
    sourceSheet = SpreadsheetApp.openById(sourceId).getSheetByName(sourceSheetName);
    if (!sourceSheet) throw new Error("Source sheet name not found.");
  } catch (e) { 
    throw new Error("Source Access Error (B1/B2): " + e.message); 
  }

  // 5. OPEN TARGET SPREADSHEET (External or Internal via B4 ID)
  let targetSheet;
  try {
    const targetSS = SpreadsheetApp.openById(targetId);
    targetSheet = targetSS.getSheetByName(targetSheetName);
    if (!targetSheet) throw new Error("Target sheet name '" + targetSheetName + "' not found in target file.");
  } catch (e) {
    throw new Error("Target Access Error (B4/B5): " + e.message);
  }

  // 6. DEFINE DATE RANGE
  const start = new Date(dateRange.start);
  const end = new Date(dateRange.end);
  start.setHours(0,0,0,0);
  end.setHours(23,59,59,999);

  const sourceRawData = sourceSheet.getDataRange().getValues();
  const matchedRows = [];

  // 7. FILTER MATCHED SOURCE ROWS (KEEP KEYWORD MATCHING)
  for (let i = 1; i < sourceRawData.length; i++) {
    const row = sourceRawData[i];
    if (!row[7] || !row[11]) continue;

    const rowDate = new Date(row[7]); // Source H
    const colLValue = String(row[11]).toLowerCase(); // Source L
    const colKValue = String(row[10] || ""); // Source K

    const cityMatch = colLValue.includes("jaipur") || colLValue.includes("astro");

    if (rowDate >= start && rowDate <= end && cityMatch) {
      // Preserved exactly as existing flow.
      let matchedKeyword = "NA";
      for (let kw of keywords) {
        if (colKValue.toLowerCase().includes(kw.toLowerCase())) {
          matchedKeyword = kw;
          break;
        }
      }

      matchedRows.push({ row, matchedKeyword });
    }
  }

  // 8. INSERT EACH ROW AT TOP (ROW 3), SO NEXT ROW AGAIN COMES ON TOP
  if (matchedRows.length === 0) {
    return "No Jaipur/Astro records found for the selected dates.";
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  for (let i = 0; i < matchedRows.length; i++) {
    const sourceRow = matchedRows[i].row;

    // Insert a new row under header and shift existing data down.
    targetSheet.insertRowsBefore(3, 1);

    // Increment Column A from the row that shifted down to row 4.
    const belowSerial = targetSheet.getRange(4, 1).getDisplayValue().trim();
    const newSerial = incrementCodeFromBelow(belowSerial, today);

    const targetRow = new Array(24).fill(""); // A:X
    targetRow[0] = newSerial;                        // A incremented code
    targetRow[1] = new Date(today);                  // B order date = today
    targetRow[2] = sourceRow[6] || "";               // C from G
    targetRow[4] = "Pending";                       // E always Pending
    targetRow[5] = sourceRow[9] || "";               // F from J
    targetRow[6] = sourceRow[10] || "";              // G from K
    targetRow[7] = sourceRow[14] || "";              // H from O
    // I, J, K, L remain empty
    targetRow[12] = convertIndToUsSize(sourceRow[13]); // M from N converted IND -> US
    targetRow[13] = sourceRow[13] || "";             // N from N
    // O remains empty
    targetRow[15] = sourceRow[29] || "";             // P from AD
    // Q, R remain empty
    targetRow[18] = "=IMAGE(P3)";                    // S formula from P
    // T, U, V, W, X remain empty

    targetSheet.getRange(3, 1, 1, 24).setValues([targetRow]);
  }

  return "✅ Success: " + matchedRows.length + " rows added at top of " + targetSheetName;
}

function incrementCodeFromBelow(belowCode, today) {
  const text = String(belowCode || "").trim();
  const match = text.match(/^(.*-)(\d+)$/);

  if (match) {
    const prefix = match[1];
    const nextNum = Number(match[2]) + 1;
    return prefix + nextNum;
  }

  const monthYear = Utilities.formatDate(today, Session.getScriptTimeZone(), "MMM-yy").toUpperCase();
  return monthYear + "-1";
}

function convertIndToUsSize(sourceNValue) {
  const cleaned = String(sourceNValue || "").replace(/\s+/g, "");
  const match = cleaned.match(/IND:([0-9]+(?:\.[0-9]+)?)/i);

  if (!match) return "";

  const indSize = parseFloat(match[1]);
  if (isNaN(indSize)) return "";

  const indToUsMap = {
  7:4.25,
  8:4.5,
  9:4.75,
  9.5:5,
  10:5.25,
  10.5:5.5,
  11:5.75,
  12:6,
  12.5:6.25,
  13:6.5,
  13.5:6.75,
  14:7,
  15:7.25,
  15.5:7.5,
  16:7.75,
  17:8,
  17.5:8.25,
  18:8.5,
  19:8.75,
  19.5:9,
  20:9.25,
  21:9.5,
  21.5:9.75,
  22:10,
  23:10.25,
  23.5:10.5,
  24:10.75,
  25:11,
  25.5:11.25,
  26:11.5,
  26.5:11.75,
  27:12,
  28:12.25,
  28.5:12.5,
  29:12.75,
  30:13,
  30.5:13.25,
  31:13.5
  };

  const usSize = indToUsMap[indSize];
  if (usSize === undefined) return "";

  return "US:" + usSize;
}

