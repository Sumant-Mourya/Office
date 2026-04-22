function handleEdit(e) {
  try {
    if (!e || !e.range) return;

    var sheet = e.range.getSheet();
    if (sheet.getName() !== "ReadyToShip") return;

    var startRow = e.range.getRow();
    var endRow = e.range.getLastRow();
    var startCol = e.range.getColumn();
    var endCol = e.range.getLastColumn();

    // Check if the edited range includes column 7 (G)
    if (startCol > 7 || endCol < 7) return;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var readySheet = ss.getSheetByName("ReadyToShip");

    // 🔧 Config
    var config = ss.getSheetByName("Configuration");
    var externalId = config.getRange("B1").getValue();
    var sheetName = config.getRange("B2").getValue();

    logMessage("INFO", "Using sheet: " + sheetName);

    var extSS = SpreadsheetApp.openById(externalId);
    var extSheet = extSS.getSheetByName(sheetName);

    if (!extSheet) {
      logMessage("ERROR", "Sheet not found");
      return;
    }

    var extRange = extSheet.getDataRange();
    var values = extRange.getDisplayValues();

    // Process each row in the edited range
    for (var currentRow = startRow; currentRow <= endRow; currentRow++) {
      if (currentRow === 1) continue; // Skip header

      var orderId = sheet.getRange(currentRow, 7).getValue();
      if (!orderId) {
        sheet.getRange(currentRow, 8).clearContent(); // clear status
        continue;
      }

      logMessage("INFO", "Start: " + orderId + " at row " + currentRow);

      // 🔴 Duplicate check (Column G)
      var lastRow = readySheet.getLastRow();
      if (lastRow > 1) {
        var existing = readySheet.getRange(2, 7, lastRow - 1).getValues();
        var count = 0;

        for (var i = 0; i < existing.length; i++) {
          if (existing[i][0] == orderId) count++;
        }

        if (count > 1) {
          readySheet.getRange(currentRow, 8).setValue("Duplicate ❌");
          logMessage("WARN", "Duplicate: " + orderId);
          continue;
        }
      }

      // 🔍 Search in Column G
      var found = false;
      for (var i = 0; i < values.length; i++) {
        if (values[i][6] == orderId) {

          var rowData = values[i];

          // ✅ Copy full row
          readySheet.getRange(currentRow, 1, 1, rowData.length)
            .setValues([rowData]);

          // ✅ Column M = IMAGE(AD[row])
          readySheet.getRange(currentRow, 13)
            .setFormula('=IMAGE(AD' + currentRow + ')');

          // ✅ Clear status (overwrite duplicate or old state)
          readySheet.getRange(currentRow, 8).clearContent();

          logMessage("SUCCESS", "Copied: " + orderId);
          found = true;
          break;
        }
      }

      if (!found) {
        // ❌ Not found
        readySheet.getRange(currentRow, 8).setValue("Not Found ❌");
        logMessage("WARN", "Not found: " + orderId);
      }
    }

  } catch (err) {
    logMessage("ERROR", err.toString());
  }
}


// ✅ LOG FUNCTION
function logMessage(level, message) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName("Log");

    if (!logSheet) return;

    if (logSheet.getLastRow() === 0) {
      logSheet.appendRow(["Timestamp", "Level", "Log"]);
    }

    logSheet.appendRow([new Date(), level, message]);

  } catch (e) {
    Logger.log(e);
  }
}