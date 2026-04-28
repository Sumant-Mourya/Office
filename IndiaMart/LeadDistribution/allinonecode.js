function runAllOperations() {
  Logger.log("🚀 STARTING ALL OPERATIONS");

  try {
    // 1. First sync all data from sheets
    Logger.log("📊 Step 1: Running Full Sync...");
    runFullSync();

    // 2. Then update tracking details
    Logger.log("📦 Step 2: Updating Tracking Details...");
    updateTrackingDetails();

    // 3. Finally calculate incentives
    Logger.log("💰 Step 3: Calculating Incentives...");
    UpdateIncentive();

    Logger.log("✅ ALL OPERATIONS COMPLETED SUCCESSFULLY!");
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "All operations completed successfully!",
    );
  } catch (error) {
    Logger.log("❌ ERROR: " + error.toString());
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Error occurred: " + error.message,
    );
  }
}

function updateTrackingDetails() {
  const sourceSS = SpreadsheetApp.openById(
    "1YiY43sxWZu-WtCQlU9dqNhN7EVE6umpqZMu7mBJcX-U",
  );
  const sourceSheet = sourceSS.getSheets()[0];
  const sourceData = sourceSheet.getDataRange().getValues();

  const targetSS = SpreadsheetApp.openById(
    "1nX5NnhT4B6I0oRuTDgPLmrVfmBS46Y3cXuMQkTBQZWs",
  );
  const targetSheet = targetSS.getSheetByName("Incentive_Report");
  const targetData = targetSheet.getDataRange().getValues();

  const courierSheet = targetSS.getSheetByName("CourierCodes");
  const courierData = courierSheet.getDataRange().getValues();

  // ===== COURIER MAP =====
  const courierMap = new Map();
  for (let i = 1; i < courierData.length; i++) {
    const name = String(courierData[i][0]).toLowerCase().trim();
    const code = courierData[i][1];
    if (name && code) courierMap.set(name, code);
  }

  // ===== MAIN MAP =====
  const map = new Map();
  for (let i = 1; i < sourceData.length; i++) {
    const trackingCode = sourceData[i][26];
    if (trackingCode) {
      map.set(String(trackingCode).trim(), sourceData[i]);
    }
  }

  // ===== PROCESS =====
  for (let i = 1; i < targetData.length; i++) {
    const code = targetData[i][9];
    if (!code) continue;

    const cleanCode = String(code).trim();
    const match = map.get(cleanCode);

    if (match) {
      const status = match[27] || "";
      const prevStatus = targetData[i][24]; // Column Y (Last Status)
      const courierNameRaw = match[25];
      const country = match[23] || "";

      let trackingLink = "https://t.17track.net/en#nums=" + cleanCode;

      // ===== COURIER CODE =====
      if (courierNameRaw) {
        const courierName = String(courierNameRaw).toLowerCase().trim();
        const fcCode = courierMap.get(courierName);
        if (fcCode) trackingLink += "&fc=" + fcCode;
      }

      targetData[i][10] = trackingLink;
      targetData[i][11] = status;

      // ===== DELIVERY TIME =====
      let deliveryText = "";

      if (!country || String(country).trim() === "") {
        deliveryText = "Delivery timeline will be shared soon";
      } else if (String(country).toLowerCase().includes("india")) {
        deliveryText = "Usual delivery time is 3–5 days";
      } else {
        deliveryText =
          "Usual delivery time is 15–20 days (international shipping)";
      }

      // ===== COMMON DATA =====
      const mobile = targetData[i][4];
      const name = targetData[i][2] || "Sir";
      const paymentMethod = targetData[i][3];

      if (mobile) {
        const cleanMobile = "91" + String(mobile).replace(/\D/g, "");

        // =========================
        // 🚚 TRANSIT MESSAGE
        // =========================
        const msg = `Hello ${name} 👋,

Your order is in transit 🚚  

📦 ${deliveryText}  

Tracking Code: ${cleanCode}  
Track here: ${trackingLink}  

Please pick up the courier call 📞 so delivery happens smoothly 😊  

Thanks 🙏`;

        const url =
          "https://web.whatsapp.com/send?phone=" +
          cleanMobile +
          "&text=" +
          encodeURIComponent(msg);

        targetData[i][13] = `=HYPERLINK("${url}","Send Message")`;

        // =========================
        // 💰 COD MESSAGE
        // =========================
        if (
          String(paymentMethod).toUpperCase() === "COD" &&
          String(status).toLowerCase().includes("deliver")
        ) {
          const codMsg = `Hello ${name} 😊,

Your order has been delivered 🎉  
Kindly confirm COD payment 💰  

Thank you for shopping with us 🙏`;

          const codUrl =
            "https://web.whatsapp.com/send?phone=" +
            cleanMobile +
            "&text=" +
            encodeURIComponent(codMsg);

          targetData[i][14] = `=HYPERLINK("${codUrl}","Send COD Msg")`;
        } else {
          targetData[i][14] = "";
        }

        // =========================
        // 🔁 RTO MESSAGE
        // =========================
        if (String(status).toLowerCase().includes("return")) {
          const rtoMsg = `Hello ${name},

Your return has been completed 🔁  
If you need any help, feel free to contact us 😊`;

          const rtoUrl =
            "https://web.whatsapp.com/send?phone=" +
            cleanMobile +
            "&text=" +
            encodeURIComponent(rtoMsg);

          targetData[i][15] = `=HYPERLINK("${rtoUrl}","Send RTO Msg")`;
        } else {
          targetData[i][15] = "";
        }

        // =========================
        // ⭐ FEEDBACK MESSAGE
        // =========================
        if (String(status).toLowerCase().includes("deliver")) {
          const feedbackMsg = `Hi ${name} 😊,

We hope you had a great experience with us ✨  

⭐ Google Feedback  
https://g.page/r/CYDC5_X5wVDMEAE/review  

🛍️ IndiaMART Feedback  
https://IndiaMART.in/j2lZzgpW  

📞 Contact: 08043878940  

Thanks for choosing us 🙏  
- Team 55Carat`;

          const feedbackUrl =
            "https://web.whatsapp.com/send?phone=" +
            cleanMobile +
            "&text=" +
            encodeURIComponent(feedbackMsg);

          targetData[i][17] = `=HYPERLINK("${feedbackUrl}","Send Feedback")`;
        } else {
          targetData[i][17] = "";
        }

        // =========================
        // 📲 DYNAMIC WHATSAPP ENGINE
        // =========================

        // Only update if status changed
        if (status !== prevStatus) {
          let finalMsg = "";

          // 🚚 TRANSIT
          if (String(status).toLowerCase().includes("transit")) {
            finalMsg = `Hello ${name} 👋,

Your order is in transit 🚚  

📦 ${deliveryText}  

Tracking Code: ${cleanCode}  
Track here: ${trackingLink}  

Please pick up the courier call 📞  

Thanks 🙏`;
          }

          // ✅ DELIVERED
          else if (String(status).toLowerCase().includes("deliver")) {
            finalMsg = `Hello ${name} 😊,

🎉 Your order has been delivered!

Tracking Code: ${cleanCode}  

We hope you loved it ❤️  

⭐ Please share your feedback:
https://g.page/r/CYDC5_X5wVDMEAE/review  

🛍️ IndiaMART:
https://IndiaMART.in/j2lZzgpW  

Thanks 🙏  
- Team 55Carat`;
          }

          // 🔁 RETURN
          else if (String(status).toLowerCase().includes("return")) {
            finalMsg = `Hello ${name},

Your order return has been completed 🔁  

If you need any help, feel free to contact us 😊`;
          }

          // 📦 DEFAULT
          else {
            finalMsg = `Hello ${name},

Your order status is: ${status}

Track here:
${trackingLink}`;
          }

          const finalUrl =
            "https://web.whatsapp.com/send?phone=" +
            cleanMobile +
            "&text=" +
            encodeURIComponent(finalMsg);

          // 👉 Column W (23)
          targetData[i][22] = `=HYPERLINK("${finalUrl}","Send WhatsApp")`;

          // 👉 Column X (24)
          targetData[i][23] = "Pending";

          // 👉 Save current status → Column Y (25)
          targetData[i][24] = status;
        }
      }
    }

    // =========================
    // 📞 CALL REQUEST DEFAULT
    // =========================
    if (!targetData[i][18] || targetData[i][18] === "") {
      targetData[i][18] = "Pending";
    }
  }
 // 🔥 FORCE FIXED COLUMN SIZE (IMPORTANT)
const REQUIRED_COLUMNS = 25; // up to Last Status column

// Ensure all rows have same column length
for (let i = 0; i < targetData.length; i++) {
  while (targetData[i].length < REQUIRED_COLUMNS) {
    targetData[i].push("");
  }
}

// Ensure sheet also has enough columns
const currentCols = targetSheet.getMaxColumns();
if (currentCols < REQUIRED_COLUMNS) {
  targetSheet.insertColumnsAfter(currentCols, REQUIRED_COLUMNS - currentCols);
}

  targetSheet
    .getRange(1, 1, targetData.length, targetData[0].length)
    .setValues(targetData);

  Logger.log("✅ FINAL SYSTEM WITH DELIVERY + CALL DEFAULT LIVE");
}
function updateTrackingForRow_(row) {
  Logger.log("🚀 updateTrackingForRow_ | Row: " + row);

  const sourceSS = SpreadsheetApp.openById("1ScYG5YKco_-aFPk2QqZvwcPkdRlhgJDSQ4A5KnkZbQw");
  const sourceSheet = sourceSS.getSheets()[0];
  const sourceData = sourceSheet.getDataRange().getValues();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getActiveSheet();
  const lastCol = targetSheet.getLastColumn();
  const rowData = targetSheet.getRange(row, 1, 1, lastCol).getValues()[0];

  const courierSheet = ss.getSheetByName("CourierCodes");
  if (!courierSheet) {
    Logger.log("⚠️ CourierCodes sheet not found");
    return;
  }
  const courierData = courierSheet.getDataRange().getValues();

  // ===== COURIER MAP =====
  const courierMap = new Map();
  for (let i = 1; i < courierData.length; i++) {
    const cName = String(courierData[i][0]).toLowerCase().trim();
    const cCode = courierData[i][1];
    if (cName && cCode) courierMap.set(cName, cCode);
  }

  // ===== SOURCE MAP (tracking code → row) =====
  const map = new Map();
  for (let i = 1; i < sourceData.length; i++) {
    const trackingCode = sourceData[i][26];
    if (trackingCode) {
      map.set(String(trackingCode).trim(), sourceData[i]);
    }
  }

  const cleanCode = String(rowData[9]).trim(); // Column J (10) = index 9
  if (!cleanCode) {
    Logger.log("⚠️ No tracking code in row " + row);
    return;
  }

  const match = map.get(cleanCode);
  if (!match) {
    // No match in source — still generate a basic tracking link
    const basicLink = "https://t.17track.net/en#nums=" + cleanCode;
    targetSheet.getRange(row, 11).setValue(basicLink); // K = TrackingLink
    Logger.log("ℹ️ No source match for: " + cleanCode + " — basic link set");
    ss.toast("🔗 Tracking link created (no status data yet)", "Tracking", 3);
    return;
  }

  // ===== EXTRACT SOURCE DATA =====
  const status = match[27] || "";
  const prevStatus = rowData[24] || ""; // Column Y (25) = index 24 = Last Status
  const courierNameRaw = match[25];
  const country = match[23] || "";

  // ===== BUILD TRACKING LINK =====
  let trackingLink = "https://t.17track.net/en#nums=" + cleanCode;

  if (courierNameRaw) {
    const courierName = String(courierNameRaw).toLowerCase().trim();
    const fcCode = courierMap.get(courierName);
    if (fcCode) trackingLink += "&fc=" + fcCode;
  }

  // ===== WRITE: TrackingLink (K=11) & Order Status (L=12) =====
  targetSheet.getRange(row, 11).setValue(trackingLink);
  targetSheet.getRange(row, 12).setValue(status);

  // ===== DELIVERY TEXT =====
  let deliveryText = "";
  if (!country || String(country).trim() === "") {
    deliveryText = "Delivery timeline will be shared soon";
  } else if (String(country).toLowerCase().includes("india")) {
    deliveryText = "Usual delivery time is 3–5 days";
  } else {
    deliveryText = "Usual delivery time is 15–20 days (international shipping)";
  }

  // ===== COMMON DATA =====
  const name = rowData[2] || "Sir";           // Col C (3)
  const mobile = rowData[4];                  // Col E (5)
  const paymentMethod = rowData[3];           // Col D (4)

  if (!mobile) {
    Logger.log("⚠️ No mobile number — skipping WhatsApp messages");
    ss.toast("🔗 Tracking link + status updated (no mobile for messages)", "Tracking", 3);
    return;
  }

  const cleanMobile = "91" + String(mobile).replace(/\D/g, "");

  // =========================
  // 🚚 TRANSIT MESSAGE (Col N = 14)
  // =========================
  const transitMsg = `Hello ${name} 👋,

Your order is in transit 🚚  

📦 ${deliveryText}  

Tracking Code: ${cleanCode}  
Track here: ${trackingLink}  

Please pick up the courier call 📞 so delivery happens smoothly 😊  

Thanks 🙏`;

  const transitUrl =
    "https://web.whatsapp.com/send?phone=" + cleanMobile +
    "&text=" + encodeURIComponent(transitMsg);

  targetSheet.getRange(row, 14).setFormula(
    `=HYPERLINK("${transitUrl}","Send Message")`
  );

  // =========================
  // 💰 COD MESSAGE (Col O = 15)
  // =========================
  if (
    String(paymentMethod).toUpperCase() === "COD" &&
    String(status).toLowerCase().includes("deliver")
  ) {
    const codMsg = `Hello ${name} 😊,

Your order has been delivered 🎉  
Kindly confirm COD payment 💰  

Thank you for shopping with us 🙏`;

    const codUrl =
      "https://web.whatsapp.com/send?phone=" + cleanMobile +
      "&text=" + encodeURIComponent(codMsg);

    targetSheet.getRange(row, 15).setFormula(
      `=HYPERLINK("${codUrl}","Send COD Msg")`
    );
  } else {
    targetSheet.getRange(row, 15).setValue("");
  }

  // =========================
  // 🔁 RTO MESSAGE (Col P = 16)
  // =========================
  if (String(status).toLowerCase().includes("return")) {
    const rtoMsg = `Hello ${name},

Your return has been completed 🔁  
If you need any help, feel free to contact us 😊`;

    const rtoUrl =
      "https://web.whatsapp.com/send?phone=" + cleanMobile +
      "&text=" + encodeURIComponent(rtoMsg);

    targetSheet.getRange(row, 16).setFormula(
      `=HYPERLINK("${rtoUrl}","Send RTO Msg")`
    );
  } else {
    targetSheet.getRange(row, 16).setValue("");
  }

  // =========================
  // ⭐ FEEDBACK MESSAGE (Col R = 18)
  // =========================
  if (String(status).toLowerCase().includes("deliver")) {
    const feedbackMsg = `Hi ${name} 😊,

We hope you had a great experience with us ✨  

⭐ Google Feedback  
https://g.page/r/CYDC5_X5wVDMEAE/review  

🛍️ IndiaMART Feedback  
https://IndiaMART.in/j2lZzgpW  

📞 Contact: 08043878940  

Thanks for choosing us 🙏  
- Team 55Carat`;

    const feedbackUrl =
      "https://web.whatsapp.com/send?phone=" + cleanMobile +
      "&text=" + encodeURIComponent(feedbackMsg);

    targetSheet.getRange(row, 18).setFormula(
      `=HYPERLINK("${feedbackUrl}","Send Feedback")`
    );
  } else {
    targetSheet.getRange(row, 18).setValue("");
  }

  // =========================
  // 📲 DYNAMIC WHATSAPP ENGINE (Col W=23, X=24, Y=25)
  // Only update if status changed from last known
  // =========================
  if (status !== prevStatus) {
    let finalMsg = "";

    // 🚚 TRANSIT
    if (String(status).toLowerCase().includes("transit")) {
      finalMsg = `Hello ${name} 👋,

Your order is in transit 🚚  

📦 ${deliveryText}  

Tracking Code: ${cleanCode}  
Track here: ${trackingLink}  

Please pick up the courier call 📞  

Thanks 🙏`;
    }

    // ✅ DELIVERED
    else if (String(status).toLowerCase().includes("deliver")) {
      finalMsg = `Hello ${name} 😊,

🎉 Your order has been delivered!

Tracking Code: ${cleanCode}  

We hope you loved it ❤️  

⭐ Please share your feedback:
https://g.page/r/CYDC5_X5wVDMEAE/review  

🛍️ IndiaMART:
https://IndiaMART.in/j2lZzgpW  

Thanks 🙏  
- Team 55Carat`;
    }

    // 🔁 RETURN
    else if (String(status).toLowerCase().includes("return")) {
      finalMsg = `Hello ${name},

Your order return has been completed 🔁  

If you need any help, feel free to contact us 😊`;
    }

    // 📦 DEFAULT
    else {
      finalMsg = `Hello ${name},

Your order status is: ${status}

Track here:
${trackingLink}`;
    }

    const finalUrl =
      "https://web.whatsapp.com/send?phone=" + cleanMobile +
      "&text=" + encodeURIComponent(finalMsg);

    // Col W (23) = WhatsApp Message link
    targetSheet.getRange(row, 23).setFormula(
      `=HYPERLINK("${finalUrl}","Send WhatsApp")`
    );

    // Col X (24) = Message Status → reset to Pending
    targetSheet.getRange(row, 24).setValue("Pending");

    // Col Y (25) = Save current status as Last Status
    targetSheet.getRange(row, 25).setValue(status);

    Logger.log("  📲 Dynamic WhatsApp updated (status changed)");
  } else {
    Logger.log("  ℹ️ Status unchanged — skipping dynamic WhatsApp");
  }

  // =========================
  // 📞 CALL REQUEST DEFAULT (Col S = 19)
  // =========================
  const currentCallStatus = rowData[18]; // index 18 = Col S
  if (!currentCallStatus || currentCallStatus === "") {
    targetSheet.getRange(row, 19).setValue("Pending");
  }

  Logger.log("✅ Row " + row + " fully updated");
}

// ═══════════════════════════════════════════════════
// 🔒 TRACKING SHEET HANDLER
// ═══════════════════════════════════════════════════
function handleTrackingSheet_(sheet, row, col, newValue, e) {
  if (row < 2) return;

  const CALL_COL = 19;
  const MSG_COL = 13;
  const WA_STATUS_COL = 24;

  if (col !== CALL_COL && col !== MSG_COL && col !== WA_STATUS_COL) return;

  const cell = sheet.getRange(row, col);
  const props = PropertiesService.getDocumentProperties();
  const normalized = newValue.toLowerCase();

  // ── CALL COLUMN (19) ──
  if (col === CALL_COL) {
    const callKey = "call_row_" + row;
    Logger.log("📞 Call column edit | Row " + row);

    if (normalized !== "completed" && props.getProperty(callKey) === "Completed") {
      cell.setValue("Completed");
      SpreadsheetApp.getActiveSpreadsheet().toast("🔒 Call is locked — already Completed", "Locked", 3);
      Logger.log("  ↳ BLOCKED: already Completed");
      return;
    }

    if (normalized === "completed") {
      props.setProperty(callKey, "Completed");
      cell.setValue("Completed");
      Logger.log("  ↳ Locked as Completed");
    } else if (normalized === "pending") {
      cell.setValue("Pending");
    } else if (normalized === "unreachable") {
      cell.setValue("Unreachable");
    }
    return;
  }

  // ── MESSAGE COLUMN (13) ──
  if (col === MSG_COL) {
    const msgKey = "msg_row_" + row;
    Logger.log("📨 Message column edit | Row " + row);

    if (normalized !== "sent" && props.getProperty(msgKey) === "Sent") {
      cell.setValue("Sent");
      SpreadsheetApp.getActiveSpreadsheet().toast("🔒 Message is locked — already Sent", "Locked", 3);
      Logger.log("  ↳ BLOCKED: already Sent");
      return;
    }

    if (normalized === "sent") {
      props.setProperty(msgKey, "Sent");
      cell.setValue("Sent");
      Logger.log("  ↳ Locked as Sent");
    } else if (normalized === "pending") {
      cell.setValue("Pending");
    }
    return;
  }
  // ── WHATSAPP STATUS COLUMN (24) ──
  if (col === WA_STATUS_COL) {
    const waKey = "wa_msg_" + row;
    Logger.log("📲 WhatsApp status edit | Row " + row);

    if (normalized !== "sent" && props.getProperty(waKey) === "Sent") {
      cell.setValue("Sent");
      SpreadsheetApp.getActiveSpreadsheet().toast("🔒 WhatsApp status locked — already Sent", "Locked", 3);
      Logger.log("  ↳ BLOCKED: already Sent");
      return;
    }

    if (normalized === "sent") {
      props.setProperty(waKey, "Sent");
      cell.setValue("Sent");
      Logger.log("  ↳ Locked as Sent");
    } else if (normalized === "pending") {
      cell.setValue("Pending");
    }
    return;
  }
}

// ═══════════════════════════════════════════════════
// 🚀 LEAD SHEET HANDLER (LeadCallMsg & followup)
// ═══════════════════════════════════════════════════
function handleLeadSheet_(ss, sheet, sheetName, row, col, newValue, oldValue, e) {
  const START_ROW = 3;
  const PHONE_COL = 5;    // E
  const WHATSAPP_COL = 8; // H
  const CALLING_COL = 9;  // I
  const STATUS_COL = 10;  // J

  // Ignore header rows
  if (row < START_ROW) return;

  // Route to the correct sheet handler
  if (sheetName === "LeadCallMsg") {
    handleLeadCallMsg_(ss, sheet, row, col, newValue, oldValue, e, START_ROW, PHONE_COL, WHATSAPP_COL, CALLING_COL, STATUS_COL);
  } else if (sheetName === "followup") {
    handleFollowupSheet_(ss, sheet, row, col, newValue, oldValue, e, START_ROW, PHONE_COL, STATUS_COL);
  }
}

// ═══════════════════════════════════════════════════
// 📋 LEADCALLMSG HANDLER
// Statuses: Follow up, Closed, Pending
// H/I edit → auto-set "Follow up" (NO copy to followup sheet)
// ═══════════════════════════════════════════════════
function handleLeadCallMsg_(ss, sheet, row, col, newValue, oldValue, e, START_ROW, PHONE_COL, WHATSAPP_COL, CALLING_COL, STATUS_COL) {

  // Only react to WhatsApp (H), Calling (I), or Status (J)
  if (col !== WHATSAPP_COL && col !== CALLING_COL && col !== STATUS_COL) return;

  // Get full row data
  const lastCol = sheet.getLastColumn();
  const range = sheet.getRange(row, 1, 1, lastCol);
  const rowData = range.getValues()[0];
  const formulas = range.getFormulas()[0];
  const phone = String(rowData[PHONE_COL - 1]).trim();

  if (!phone) {
    Logger.log("⚠️ Skipped: No phone number in row " + row);
    return;
  }

  Logger.log("📱 Phone: " + phone + " | Sheet: LeadCallMsg");

  const historySheet = ss.getSheetByName("History");
  if (!historySheet) {
    SpreadsheetApp.getActiveSpreadsheet().toast("⚠️ 'History' sheet not found!", "Error", 5);
    return;
  }

  // ─────────────────────────────────────────
  // CASE 1: WhatsApp (H) or Calling (I) edited
  // → Only set status to "Follow up", NO copy to followup sheet
  // ─────────────────────────────────────────
  if (col === WHATSAPP_COL || col === CALLING_COL) {
    Logger.log("📞 WhatsApp/Calling edit → auto-setting Follow up (no copy)");

    sheet.getRange(row, STATUS_COL).setValue("Follow up");

    // Remove from History if it was there
    removeRowsByPhone_(historySheet, phone, PHONE_COL, START_ROW);

    SpreadsheetApp.getActiveSpreadsheet().toast("📋 Status → Follow up", "Updated", 3);
    return;
  }

  // ─────────────────────────────────────────
  // CASE 2: Status column (J) edited
  // ─────────────────────────────────────────
  if (col === STATUS_COL) {
    Logger.log("🔄 LeadCallMsg status: [" + oldValue + "] → [" + newValue + "]");

    // ── Follow up ──
    if (newValue === "Follow up") {
      Logger.log("  ↳ Branch: Follow up (status set, no copy)");

      // Remove from History if present
      removeRowsByPhone_(historySheet, phone, PHONE_COL, START_ROW);

      SpreadsheetApp.getActiveSpreadsheet().toast("📋 Status set to Follow up", "Follow up", 3);
      return;
    }

    // ── Closed ──
    if (newValue === "Closed") {
      Logger.log("  ↳ Branch: Closed");

      // VALIDATION: Columns A (1) through I (9) must all be filled
      for (let c = 0; c < 9; c++) {
        const val = rowData[c];
        if (val === "" || val === null || val === undefined) {
          e.range.setValue(oldValue || "");
          SpreadsheetApp.getActiveSpreadsheet().toast(
            "❌ Fill all fields A → I before marking Closed (Column " + String.fromCharCode(65 + c) + " is empty)",
            "Validation Error",
            5
          );
          Logger.log("  ↳ BLOCKED: Column " + String.fromCharCode(65 + c) + " is empty");
          return;
        }
      }

      // Copy to History (no duplicates)
      copyRowToSheet_(historySheet, rowData, formulas, phone, PHONE_COL, START_ROW, lastCol);

      SpreadsheetApp.getActiveSpreadsheet().toast("✅ Closed → Saved to History", "Closed", 3);
      return;
    }

    // ── Pending ──
    if (newValue === "Pending") {
      Logger.log("  ↳ Branch: Pending");

      // Remove from History only
      removeRowsByPhone_(historySheet, phone, PHONE_COL, START_ROW);

      SpreadsheetApp.getActiveSpreadsheet().toast("🔄 Pending → Removed from History", "Pending", 3);
      return;
    }

    // ── Closed → anything else (reopen) ──
    if (oldValue === "Closed" && newValue !== "Closed") {
      Logger.log("  ↳ Branch: Reopened from Closed");

      removeRowsByPhone_(historySheet, phone, PHONE_COL, START_ROW);

      SpreadsheetApp.getActiveSpreadsheet().toast("↩️ Reopened → Removed from History", "Reopened", 3);
      return;
    }

    Logger.log("  ↳ No matching rule for status [" + newValue + "]");
    return;
  }
}

// ═══════════════════════════════════════════════════
// 📂 FOLLOWUP SHEET HANDLER
// Only 2 statuses: "Follow up" (stay) and "Closed" (→ History + delete)
// No WhatsApp/Calling triggers on this sheet
// ═══════════════════════════════════════════════════
function handleFollowupSheet_(ss, sheet, row, col, newValue, oldValue, e, START_ROW, PHONE_COL, STATUS_COL) {

  // Only react to Status column (J) edits
  if (col !== STATUS_COL) return;

  // Get full row data
  const lastCol = sheet.getLastColumn();
  const range = sheet.getRange(row, 1, 1, lastCol);
  const rowData = range.getValues()[0];
  const formulas = range.getFormulas()[0];
  const phone = String(rowData[PHONE_COL - 1]).trim();

  if (!phone) {
    Logger.log("⚠️ Skipped: No phone number in row " + row);
    return;
  }

  Logger.log("📱 Phone: " + phone + " | Sheet: followup");

  const historySheet = ss.getSheetByName("History");
  if (!historySheet) {
    SpreadsheetApp.getActiveSpreadsheet().toast("⚠️ 'History' sheet not found!", "Error", 5);
    return;
  }

  Logger.log("🔄 followup status: [" + oldValue + "] → [" + newValue + "]");

  // ── Closed → validate, copy to History, delete from followup ──
  if (newValue === "Closed") {
    Logger.log("  ↳ Branch: Closed (followup)");

    // VALIDATION: Columns A (1) through I (9) must all be filled
    for (let c = 0; c < 9; c++) {
      const val = rowData[c];
      if (val === "" || val === null || val === undefined) {
        e.range.setValue(oldValue || "Follow up");
        SpreadsheetApp.getActiveSpreadsheet().toast(
          "❌ Fill all fields A → I before marking Closed (Column " + String.fromCharCode(65 + c) + " is empty)",
          "Validation Error",
          5
        );
        Logger.log("  ↳ BLOCKED: Column " + String.fromCharCode(65 + c) + " is empty");
        return;
      }
    }

    // Copy to History (no duplicates) — row stays in followup with "Closed" status
    copyRowToSheet_(historySheet, rowData, formulas, phone, PHONE_COL, START_ROW, lastCol);

    SpreadsheetApp.getActiveSpreadsheet().toast("✅ Closed → Saved to History", "Closed", 3);
    return;
  }

  // ── Closed → Follow up (retrieve back from History) ──
  if (newValue === "Follow up" && oldValue === "Closed") {
    Logger.log("  ↳ Branch: Closed → Follow up (retrieving from History)");

    removeRowsByPhone_(historySheet, phone, PHONE_COL, START_ROW);

    SpreadsheetApp.getActiveSpreadsheet().toast("↩️ Retrieved back from History", "Reopened", 3);
    return;
  }

  // ── Follow up (no previous Closed) → do nothing, stay in followup ──
  if (newValue === "Follow up") {
    Logger.log("  ↳ Branch: Follow up — staying in followup (no action)");
    return;
  }

  // ── Any other value → revert to Follow up (only 2 statuses allowed) ──
  Logger.log("  ↳ Invalid status on followup sheet: [" + newValue + "] → reverting to Follow up");
  e.range.setValue("Follow up");
  SpreadsheetApp.getActiveSpreadsheet().toast("⚠️ Only 'Follow up' or 'Closed' allowed here", "Invalid", 3);
}

// ═══════════════════════════════════════════════════
// 📦 HELPER: Copy a row to a target sheet (no dups)
// ═══════════════════════════════════════════════════
function copyRowToSheet_(targetSheet, rowData, formulas, phone, phoneCol, startRow, totalCols) {
  Logger.log("  📥 copyRowToSheet_ → " + targetSheet.getName() + " | Phone: " + phone);

  // Check for existing entry with same phone
  const tLastRow = targetSheet.getLastRow();

  if (tLastRow >= startRow) {
    const existingPhones = targetSheet
      .getRange(startRow, phoneCol, tLastRow - startRow + 1, 1)
      .getValues()
      .flat()
      .map(function(p) { return String(p).trim(); });

    if (existingPhones.indexOf(phone) !== -1) {
      Logger.log("  ⏭️ Duplicate found — skipping copy to " + targetSheet.getName());
      return;
    }
  }

  // Append the row
  var destRow = targetSheet.getLastRow() + 1;
  targetSheet.getRange(destRow, 1, 1, totalCols).setValues([rowData]);

  // Restore formulas (preserves HYPERLINK etc.)
  for (var i = 0; i < formulas.length; i++) {
    if (formulas[i]) {
      targetSheet.getRange(destRow, i + 1).setFormula(formulas[i]);
    }
  }

  // Remove data validations (dropdowns) from copied row
  targetSheet.getRange(destRow, 1, 1, totalCols).clearDataValidations();

  Logger.log("  ✅ Row copied to " + targetSheet.getName() + " at row " + destRow);
}

// ═══════════════════════════════════════════════════
// 🗑️ HELPER: Remove ALL rows matching phone number
// ═══════════════════════════════════════════════════
function removeRowsByPhone_(targetSheet, phone, phoneCol, startRow) {
  var sheetName = targetSheet.getName();
  Logger.log("  🗑️ removeRowsByPhone_ → " + sheetName + " | Phone: " + phone);

  var tLastRow = targetSheet.getLastRow();
  if (tLastRow < startRow) {
    Logger.log("  ℹ️ No data rows in " + sheetName);
    return;
  }

  var phoneData = targetSheet
    .getRange(startRow, phoneCol, tLastRow - startRow + 1, 1)
    .getValues()
    .flat();

  var deletedCount = 0;

  // Delete from bottom to top to avoid shifting issues
  for (var i = phoneData.length - 1; i >= 0; i--) {
    if (String(phoneData[i]).trim() === phone) {
      targetSheet.deleteRow(i + startRow);
      deletedCount++;
    }
  }

  Logger.log("  🗑️ Deleted " + deletedCount + " row(s) from " + sheetName);
}
function UpdateIncentive() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const startRow = 3;
  const lastRow = sheet.getLastRow();

  // 👉 Pull full row (we need Order Status also now)
  const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 12).getValues();

  const monthMap = {};

  // =========================
  // 🧠 GROUP + FILTER DELIVERED ONLY
  // =========================
  data.forEach((row, index) => {
    const date = row[0];        // Column A
    const amount = Number(row[8]) || 0;  // Net Amount (Column I)
    const status = String(row[11]).toLowerCase(); // Column L = Order Status

    // ❌ Skip if not delivered
    if (!status.includes("deliver")) return;

    if (!date) return;

    const d = new Date(date);
    const key = Utilities.formatDate(d, Session.getScriptTimeZone(), "MMM yy");

    if (!monthMap[key]) {
      monthMap[key] = {
        total: 0,
        firstRow: startRow + index
      };
    }

    // ✅ Only delivered sales counted
    monthMap[key].total += amount;
  });

  // =========================
  // 🧹 CLEAR OLD BOXES
  // =========================
  if (lastRow > 1) {
    sheet.getRange(2, 27, lastRow - 1, 4).clearContent();
  }

  // =========================
  // 📅 SORT MONTHS
  // =========================
  const sortedMonths = Object.keys(monthMap).sort((a, b) => {
    return new Date("1 " + a) - new Date("1 " + b);
  });

  // =========================
  // ✍️ WRITE BLOCKS
  // =========================
  sortedMonths.forEach(month => {
    const totalSales = monthMap[month].total;
    const writeRow = Math.max(2, monthMap[month].firstRow);

    let earning1 = 0;
    let earning2 = 0;
    let bonus = 0;

    if (totalSales <= 50000) {
      earning1 = totalSales * 0.025;
    } else {
      earning1 = 50000 * 0.025;
      earning2 = (totalSales - 50000) * 0.03;
    }

    if (totalSales > 150000) {
      bonus = 500;
    }

    const totalPayable = earning1 + earning2 + bonus;

    // =========================
    // 🧱 UI BLOCK
    // =========================

    const titleCell = sheet.getRange(writeRow, 27);
    titleCell.setValue(`Priya ${month} Incentive`)
      .setFontWeight("bold")
      .setFontSize(12)
      .setBackground("#d9e1f2");

    sheet.getRange(writeRow + 1, 27)
      .setValue(Math.round(totalSales))
      .setFontWeight("bold");

    sheet.getRange(writeRow + 2, 27, 1, 4).setValues([
      ["S No", "Sales", "Bonus %", "Sales Earning"]
    ])
      .setFontWeight("bold")
      .setBackground("#ddebf7")
      .setHorizontalAlignment("center");

    sheet.getRange(writeRow + 3, 27, 1, 4).setValues([[
      1,
      Math.min(totalSales, 50000),
      "2.50%",
      Math.round(earning1)
    ]]);

    sheet.getRange(writeRow + 4, 27, 1, 4).setValues([[
      2,
      totalSales > 50000 ? totalSales - 50000 : 0,
      "3%",
      Math.round(earning2)
    ]]);

    sheet.getRange(writeRow + 5, 27, 1, 4).setValues([[
      3,
      "Bonus",
      "",
      bonus
    ]]);

    sheet.getRange(writeRow + 6, 29)
      .setValue("Total Payable")
      .setFontWeight("bold");

    sheet.getRange(writeRow + 6, 30)
      .setValue(Math.round(totalPayable))
      .setFontWeight("bold");

    sheet.getRange(writeRow + 3, 27, 3, 4)
      .setHorizontalAlignment("center");

    sheet.getRange(writeRow, 27, 7, 4)
      .setBorder(true, true, true, true, true, true);
  });

  SpreadsheetApp.getActiveSpreadsheet().toast("✅ Incentive updated (Delivered only)");
}
function runFullSync() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  
  const leadSheet = ss.getSheetByName("LeadCallMsg");
  const incentiveSheet = ss.getSheetByName("Incentive_Report");

  if (!leadSheet || !incentiveSheet){
    SpreadsheetApp.getUi().alert("⚠️ Missing sheet");
    return;
  }

  Logger.log("🚀 MASTER SYNC STARTED");

  const finalData = [];
  const columnGLinks = [];

  // =========================================
  // 🔵 1. LOAD LEADCALLMSG
  // =========================================
  const leadLastRow = leadSheet.getLastRow();

  if (leadLastRow >= 3) {
    const leadData = leadSheet.getRange(3, 1, leadLastRow - 2, 10).getValues();

    for (let i = 0; i < leadData.length; i++) {
      const date = leadData[i][0] || new Date();
      const name = leadData[i][3];
      let phone = String(leadData[i][4] || "").replace(/\D/g, "");

      if (phone.length === 10) phone = "91" + phone;
      if (!phone) continue;

      const status =
        String(leadData[i][5]).toLowerCase().trim() === "sent"
          ? "Sent"
          : "Pending";

      // Column F empty for Lead
      finalData.push([date, "LeadCallMsg", name, phone, status, ""]);

      // Column G → WhatsApp link
      const formula =
        '=HYPERLINK("https://web.whatsapp.com/send?phone=' +
        phone +
        '","Whatsapp")';

      columnGLinks.push([formula]);
    }

    Logger.log("✅ LeadCallMsg Loaded");
  }

  // =========================================
  // 🟢 2. LOAD INCENTIVE REPORT
  // =========================================
  const incLastRow = incentiveSheet.getLastRow();

  if (incLastRow >= 2) {
    const incRange = incentiveSheet.getRange(2, 1, incLastRow - 1, 25);

    const incValues = incRange.getValues();
    const incFormulas = incRange.getFormulas();

    for (let i = 0; i < incValues.length; i++) {
      const date = incValues[i][0];
      const name = incValues[i][2];
      let phone = String(incValues[i][4] || "").replace(/\D/g, "");

      if (phone.length === 10) phone = "91" + phone;
      if (!phone) continue;

      const status = incValues[i][23] || "Pending";

      // 🔥 ORIGINAL MESSAGE (Column F stays SAME)
      const originalFormula = incFormulas[i][22]; // W column
      const originalValue = incValues[i][22];

      const finalMessage = originalFormula ? originalFormula : originalValue;

      finalData.push([
        date,
        "Incentive_Report",
        name,
        phone,
        status,
        finalMessage, // KEEP "Send WhatsApp"
      ]);

      // =========================================
      // 🔁 CREATE RENAMED VERSION FOR COLUMN G
      // =========================================
      let newFormula = "";

      if (originalFormula) {
        const match = originalFormula.match(/"(https?:\/\/[^"]+)"/);

        if (match && match[1]) {
          const url = match[1];

          // ✅ SAME LINK but renamed text
          newFormula = '=HYPERLINK("' + url + '","Whatsapp")';
        }
      }

      columnGLinks.push([newFormula]);
    }

    Logger.log("✅ Incentive_Report Loaded (renamed only in Column G)");
  }

  // =========================================
  // 🧹 3. CLEAR OLD DATA
  // =========================================

}
