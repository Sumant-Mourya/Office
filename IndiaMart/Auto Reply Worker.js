/******************** CONFIG ********************/
const DAILY_REPLY_LIMIT = 95;
const PER_RUN_LIMIT = 5;            // max replies to send per trigger run
const BACKFILL_START_DATE = "2026/2/24"; // fetch all IndiaMart mails from this date
const START_KEYWORDS = ["enquiry for", "re: enquiry for"];
const DRIVE_PARENT_FOLDER_ID = "1Kuu88B1BQd3SGAcWJEOWyItfVKRTEJD-";
const ENABLE_AUTO_REPLY_LOG = true; // true = Auto Reply Log ON, false = OFF
/************************************************/

/******************** LOGGER ********************/
// Named uniquely so the Backlog worker's writeLog cannot override this one
function writeAutoReplyLog(message) {
  if (!ENABLE_AUTO_REPLY_LOG) return; // 🚫 logging disabled

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName("Auto Reply Log");
  if (!logSheet) {
    logSheet = ss.insertSheet("Auto Reply Log");
    logSheet.appendRow(["Time", "Log"]);
  }
  // prefix ' to avoid formula issue
  logSheet.appendRow([new Date(), "'" + message]);
}
/************************************************/

/******************** PARSER ********************/
function parseNameAndEmail(raw) {
  if (!raw) return { name: "", email: "" };

  const m = raw.match(/^(.*?)(?:\s*<(.+?)>)$/);
  if (m) {
    return {
      name: m[1].trim(),
      email: m[2].trim()
    };
  }
  return {
    name: "",
    email: raw.trim()
  };
}
/************************************************/

/******************** DRIVE HELPERS ********************/
function findFolderByKeyword(parentFolderId, keyword) {
  const parent = DriveApp.getFolderById(parentFolderId);
  const target = keyword.toLowerCase().trim();

  const folders = parent.getFolders();
  while (folders.hasNext()) {
    const f = folders.next();
    if (f.getName().toLowerCase().trim() === target) {
      return f;
    }
  }
  return null;
}

function getImageBlobsFromFolder(folder) {
  const blobs = [];
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType().startsWith("image/")) {
      blobs.push(file.getBlob());
    }
  }
  return blobs;
}
/************************************************/

/******************** SHARED DRIVE KEYWORD HELPERS ********************
 * Both Auto Reply Worker and Backlog Worker call these.
 * Keywords come ONLY from Drive subfolder names — no sheet rules.
 **********************************************************************/
function getDriveKeywords() {
  try {
    const parent  = DriveApp.getFolderById(DRIVE_PARENT_FOLDER_ID);
    const folders = parent.getFolders();
    const names   = [];
    while (folders.hasNext()) {
      const name = folders.next().getName().toLowerCase().trim();
      if (name === "default reply") continue; // skip the catch-all folder
      names.push(name);
    }
    writeAutoReplyLog("Drive keywords loaded (" + names.length + "): " + names.join(", "));
    return names;
  } catch (e) {
    writeAutoReplyLog("Could not load Drive keywords: " + e.message);
    return [];
  }
}

/**
 * Returns the LONGEST Drive folder name found inside subjectLower,
 * or null.  Longest-first prevents short words shadowing deeper ones.
 */
function findBestDriveKeyword(subjectLower, driveKeywords) {
  const sorted = driveKeywords.slice().sort((a, b) => b.length - a.length);
  for (let i = 0; i < sorted.length; i++) {
    if (subjectLower.includes(sorted[i])) return sorted[i];
  }
  return null;
}

/** Returns true when the error is a Gmail/email send-quota error. */
function isEmailQuotaError(e) {
  const msg = (e.message || "").toLowerCase();
  return msg.includes("service invoked too many times") ||
         msg.includes("over quota") ||
         msg.includes("gmail sending quota");
}
/**********************************************************************/

/******************** EMAIL TEMPLATE ********************/
function buildEmailBody(customerName, keyword) {
  const SALES_PERSON = "Gaurav";

  if ((keyword || "").toLowerCase().includes("kapila pashu aahar")) {
    return buildKapilaMailBody(customerName);
  }
  if ((keyword || "").toLowerCase().includes("amul cattel feed")) {
    return buildAmulMailBody(customerName);
  }
  return (
    "Hi " + (customerName || "there") + ",\n" +
    "This is " + SALES_PERSON + " from 55Carat.\n" +
    "Thank you for your enquiry regarding " + keyword + 
    ". Please find the attached images for your reference.\n" +
    "Kindly review and let us know if you need pricing or any further details.\n" +
    "WHATSAPP / CALL: " +
    "9389028195 | 7983524823\n\n" +
    "Regards,\n" +
    SALES_PERSON + "\n" +
    "55Carat"
  );
}

function buildKapilaMailBody(customerName) {
  const SALES_PERSON = "krishna";
  return (
    "Hi " + (customerName || "there") + ",\n" +
    "This is " + SALES_PERSON + " from 55Carat.\n" +
    "Thank you for your enquiry regarding Kapila Pashu Aahar.\n" +
    "Products:\n" +
    "***Kapila Super Pellet\n" +
    "   Protein: 18% | Fat: 2-2.5%\n" +
    "***Daily Special Bypass\n" +
    "   Protein: 20% | Fat: 3%\n" +
    "***Kapila Buff Pro\n" +
    "   Protein: 20% | Fat: 3%\n" +
    "***Kapila Hi Pro\n" +
    "   Protein: 22% | Fat: 7%\n" +
    "***Kapila HPF\n" +
    "   Protein: 24% | Fat: 5%\n" +
    "***Kapila Super Six\n" +
    "   Protein: 26% | Fat: 6%\n" +
    "Kindly review and let us know if you need pricing or further details.\n" +
    "WHATSAPP / CALL: 9520991800\n" +
    "Regards,\n" +
    SALES_PERSON + "\n" +
    "55Carat"
  );
}

function buildAmulMailBody(customerName) {
  const SALES_PERSON = "krishna";
  return (
    "Hi " + (customerName || "there") + ",\n" +
    "This is " + SALES_PERSON + " from 55Carat.\n" +
    "Thank you for your enquiry regarding Amul Cattle Feed.\n" +
    "Products:\n" +
    "***Amul Nutri Plus\n" +
    "   Protein: 18% | Fat: 2%\n" +
    "***Amul Power Dan\n" +
    "   Protein: 20% | Fat: 3%\n" +
    "***Amul Super Dan\n" +
    "   Protein: 22% | Fat: 4%\n" +
    "***Nutri Power Pallet\n" +
    "   Protein: 25% | Fat: 7%\n" +
    "***Amul Power Mixer\n" +
    "   Protein: 20% | Fat: 3%\n" +
    "***Amul Buffelo\n" +
    "   Protein: 22% | Fat: 7%\n" +
    "Kindly review and let us know if you need pricing or further details.\n" +
    "WHATSAPP / CALL: 9520991800\n" +
    "Regards,\n" +
    SALES_PERSON + "\n" +
    "55Carat"
  );
}


// Default reply body builder
function buildDefaultEmailBody(customerName) {
  const SALES_PERSON = "Gaurav";
  return (
    "Hi " + (customerName || "there") + ",\n" +
    "This is " + SALES_PERSON + " from 55Carat.\n" +
    "Thank you for your enquiry.\n" +
    ". Please Contact on Give Number for Further information.\n" +
    "WHATSAPP / CALL: " +
    "9389028195 | 7983524823\n\n" +
    "Regards,\n" +
    SALES_PERSON + "\n" +
    "55Carat"
  );
}
/*******************************************************/


/**
 * Scans Gmail for IndiaMart mails since BACKFILL_START_DATE and logs any
 * thread not yet in the sheet. Does NOT send emails — that is deferred to
 * processPendingNoReplies and the history loop.
 */
function backfillOldMails(replyLogSheet, loggedThreadIds, MY_EMAIL) {
  const allowedSenders = [
    "buyershelpdesk@indiamart.com",
    "buyershelp+enq@indiamart.com",
    "buyleads@indiamart.com"
  ];
  const query =
    "(" + allowedSenders.map(function(s) { return "from:" + s; }).join(" OR ") + ")" +
    " after:" + BACKFILL_START_DATE;
  writeAutoReplyLog("Backfill: searching Gmail with: " + query);

  // Ensure Default column header exists
  if (replyLogSheet.getLastColumn() < 7) {
    replyLogSheet.insertColumnAfter(6);
    replyLogSheet.getRange(1, 7).setValue("Default(No Keyword Match)");
  }

  let start   = 0;
  const PAGE  = 100;
  let added   = 0;

  while (true) {
    let threads;
    try {
      threads = GmailApp.search(query, start, PAGE);
    } catch (e) {
      writeAutoReplyLog("Backfill: Gmail search error: " + e.message);
      break;
    }
    if (!threads || threads.length === 0) break;

    for (let t = 0; t < threads.length; t++) {
      const thread   = threads[t];
      const threadId = thread.getId();
      if (loggedThreadIds.has(threadId)) continue; // already in sheet

      const messages = thread.getMessages();
      if (!messages || messages.length === 0) continue;

      // Use the first IndiaMart message for metadata
      let firstMsg = null;
      for (let m = 0; m < messages.length; m++) {
        const from = (messages[m].getFrom() || "").toLowerCase();
        if (allowedSenders.some(function(addr) { return from.includes(addr); })) {
          firstMsg = messages[m];
          break;
        }
      }
      if (!firstMsg) continue;
      if ((firstMsg.getFrom() || "").toLowerCase().includes(MY_EMAIL)) continue;

      const replyToRaw = firstMsg.getReplyTo() || firstMsg.getFrom();
      const parsed     = parseNameAndEmail(replyToRaw);
      const mailDate   = firstMsg.getDate();
      const subject    = firstMsg.getSubject();

      replyLogSheet.appendRow([
        mailDate,
        threadId,
        parsed.name,
        parsed.email,
        subject,
        "NO",
        ""
      ]);
      loggedThreadIds.add(threadId);
      added++;
      writeAutoReplyLog("Backfill logged: " + threadId + " | " + subject);
    }

    start += threads.length;
    if (threads.length < PAGE) break; // reached last page
  }

  writeAutoReplyLog("Backfill complete: " + added + " new threads logged.");
}

function processPendingNoReplies(
  replyLogSheet,
  driveKeywords,
  todayCount,
  replySheet,
  runState,     // { sent: number }
  kwRepliedIds, // Set of threadIds already sent keyword replies — duplicate guard
  defRepliedIds // Set of threadIds already sent default replies — duplicate guard
) {
  const lastRow = replyLogSheet.getLastRow();
  if (lastRow < 2) return todayCount;

  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const quotaHitToday = props.getProperty("EMAIL_QUOTA_HIT_DATE") === today;
  if (quotaHitToday) {
    writeAutoReplyLog("processPendingNoReplies: quota was hit today, skipping all sends.");
    return todayCount;
  }

  // Add Default column if missing
  if (replyLogSheet.getLastColumn() < 7) {
    replyLogSheet.insertColumnAfter(6);
    replyLogSheet.getRange(1, 7).setValue("Default(No Keyword Match)");
  }

  const data = replyLogSheet.getRange(2, 1, lastRow - 1, 7).getValues();

  /* ---- Pass 1: keyword / Kapila / Amul rows first ---- */
  for (let i = 0; i < data.length; i++) {
    if (todayCount >= DAILY_REPLY_LIMIT) break;
    if (runState.sent >= PER_RUN_LIMIT) break;

    const [mailDate, threadId, name, replyTo, subject, keyword_matched] = data[i];
    if (keyword_matched === "YES") continue;
    if (keyword_matched !== "NO")  continue;

    const subjectLower = (subject || "").toLowerCase();
    const isKapila     = subjectLower.includes("kapila pashu aahar");
    const isAmul       = subjectLower.includes("amul cattel feed");
    const drivekw      = (!isKapila && !isAmul)
                         ? findBestDriveKeyword(subjectLower, driveKeywords)
                         : null;

    if (!isKapila && !isAmul && !drivekw) continue; // no keyword — handled in pass 2

    // ---- Duplicate guard: skip if keyword reply already sent ----
    if (kwRepliedIds.has(threadId)) {
      writeAutoReplyLog("Pass1 duplicate guard: keyword already replied to " + threadId + ". Marking sheet YES.");
      replyLogSheet.getRange(i + 2, 6).setValue("YES");
      continue;
    }

    if (isKapila || isAmul) {
      writeAutoReplyLog("Reprocessing special keyword. ThreadId: " + threadId);
      const specialKeyword = isKapila ? "kapila pashu aahar" : "amul cattel feed";
      const folder         = findFolderByKeyword(DRIVE_PARENT_FOLDER_ID, specialKeyword);
      const attachments    = folder ? getImageBlobsFromFolder(folder) : [];
      const body           = isKapila ? buildKapilaMailBody(name) : buildAmulMailBody(name);
      const subjectLine    = isKapila ? "Kapila Pashu Aahar - Details" : "Amul Cattel Feed - Details";
      try {
        const thread = GmailApp.getThreadById(threadId);
        if (thread) {
          thread.reply(body, { subject: subjectLine, to: replyTo, attachments: attachments.length ? attachments : undefined });
        } else {
          GmailApp.sendEmail(replyTo, subjectLine, body, attachments.length ? { attachments } : {});
        }
        replyLogSheet.getRange(i + 2, 6).setValue("YES");
        kwRepliedIds.add(threadId);
        todayCount++;
        runState.sent++;
        replySheet.getRange("H1").setValue(todayCount);
        writeAutoReplyLog("Special reply sent to: " + replyTo);
        if (runState.sent >= PER_RUN_LIMIT) {
          writeAutoReplyLog("Per-run send limit reached inside special replies. Stopping further sends this run.");
          return todayCount;
        }
      } catch (e) {
        if (isEmailQuotaError(e)) {
          writeAutoReplyLog("Email quota hit (special). Stopping sends for today.");
          todayCount = DAILY_REPLY_LIMIT;
          replySheet.getRange("H1").setValue(todayCount);
          props.setProperty("EMAIL_QUOTA_HIT_DATE", today);
          break;
        }
        writeAutoReplyLog("Error (special): " + e.message);
      }
    } else if (drivekw) {
      writeAutoReplyLog("Reprocessing Drive keyword '" + drivekw + "'. ThreadId: " + threadId);
      const folder        = findFolderByKeyword(DRIVE_PARENT_FOLDER_ID, drivekw);
      const attachments   = folder ? getImageBlobsFromFolder(folder) : [];
      const prettyKeyword = drivekw.replace(/\b\w/g, c => c.toUpperCase());
      const subjectLine   = "Message from 55Carat Regarding " + prettyKeyword;
      const body          = buildEmailBody(name, prettyKeyword);
      try {
        const thread = GmailApp.getThreadById(threadId);
        if (thread) {
          thread.reply(body, { subject: subjectLine, to: replyTo, attachments: attachments.length ? attachments : undefined });
        } else {
          GmailApp.sendEmail(replyTo, subjectLine, body, attachments.length ? { attachments } : {});
        }
        replyLogSheet.getRange(i + 2, 6).setValue("YES");
        kwRepliedIds.add(threadId);
        todayCount++;
        runState.sent++;
        replySheet.getRange("H1").setValue(todayCount);
        writeAutoReplyLog("Drive-keyword reply sent to: " + replyTo);
        if (runState.sent >= PER_RUN_LIMIT) {
          writeAutoReplyLog("Per-run send limit reached inside drive replies. Stopping further sends this run.");
          return todayCount;
        }
      } catch (e) {
        if (isEmailQuotaError(e)) {
          writeAutoReplyLog("Email quota hit (Drive keyword). Stopping sends for today.");
          todayCount = DAILY_REPLY_LIMIT;
          replySheet.getRange("H1").setValue(todayCount);
          props.setProperty("EMAIL_QUOTA_HIT_DATE", today);
          break;
        }
        writeAutoReplyLog("Error (Drive keyword): " + e.message);
      }
    }
  }

  /* ---- Pass 2: default (no keyword) rows ---- */
  if (todayCount < DAILY_REPLY_LIMIT && runState.sent < PER_RUN_LIMIT) {
    const defaultFolder = findFolderByKeyword(DRIVE_PARENT_FOLDER_ID, "Default Reply");
    for (let i = 0; i < data.length; i++) {
      if (todayCount >= DAILY_REPLY_LIMIT) break;
      if (runState.sent >= PER_RUN_LIMIT) break;

      const [mailDate, threadId, name, replyTo, subject, keyword_matched, defaultStatus] = data[i];
      if (keyword_matched === "YES") continue;
      if (keyword_matched !== "NO")  continue;
      if (defaultStatus   === "YES") continue;

      const subjectLower = (subject || "").toLowerCase();
      const isKapila     = subjectLower.includes("kapila pashu aahar");
      const isAmul       = subjectLower.includes("amul cattel feed");
      const drivekw      = (!isKapila && !isAmul)
                           ? findBestDriveKeyword(subjectLower, driveKeywords)
                           : null;
      if (isKapila || isAmul || drivekw) continue; // handled in pass 1

      // ---- Duplicate guard ----
      if (defRepliedIds.has(threadId)) {
        writeAutoReplyLog("Pass2 duplicate guard: default already replied to " + threadId + ". Marking sheet.");
        replyLogSheet.getRange(i + 2, 7).setValue("YES");
        continue;
      }

      writeAutoReplyLog("Sending default reply. ThreadId: " + threadId);
      const attachments = defaultFolder ? getImageBlobsFromFolder(defaultFolder) : [];
      const body        = buildDefaultEmailBody(name);
      try {
        const thread = GmailApp.getThreadById(threadId);
        if (thread) {
          thread.reply(body, { subject: "Message from 55Carat", to: replyTo, attachments: attachments.length ? attachments : undefined });
        } else {
          GmailApp.sendEmail(replyTo, "Message from 55Carat", body, attachments.length ? { attachments } : {});
        }
        replyLogSheet.getRange(i + 2, 7).setValue("YES");
        defRepliedIds.add(threadId);
        todayCount++;
        runState.sent++;
        replySheet.getRange("H1").setValue(todayCount);
        writeAutoReplyLog("Default reply sent to: " + replyTo);
        if (runState.sent >= PER_RUN_LIMIT) {
          writeAutoReplyLog("Per-run send limit reached inside default replies. Stopping further sends this run.");
          return todayCount;
        }
      } catch (e) {
        if (isEmailQuotaError(e)) {
          writeAutoReplyLog("Email quota hit (default). Stopping sends for today.");
          todayCount = DAILY_REPLY_LIMIT;
          replySheet.getRange("H1").setValue(todayCount);
          props.setProperty("EMAIL_QUOTA_HIT_DATE", today);
          break;
        }
        writeAutoReplyLog("Error (default): " + e.message);
      }
    }
  }

  return todayCount;
}


function autoReplyWorker() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) { 
    writeAutoReplyLog("Auto Worker: Another instance is locking the script. Skipping.");
    return;
  }

  const props = PropertiesService.getScriptProperties();

  // Handle overlapping triggers: store running flag with timestamp and a skip counter
  const runningVal = props.getProperty("AUTO_WORKER_RUNNING");
  let skipCount = Number(props.getProperty("AUTO_WORKER_SKIP_COUNT") || "0");
  if (runningVal) {
    const parts = runningVal.split("|");
    const ts = Number(parts[1] || "0");
    // If previous run started recently (5 minutes), treat it as running
    if (ts && (Date.now() - ts) < 5 * 60 * 1000) {
      skipCount++;
      props.setProperty("AUTO_WORKER_SKIP_COUNT", String(skipCount));
      writeAutoReplyLog("Previous Auto Worker still running. Skip count: " + skipCount);
      if (skipCount >= 3) {
        writeAutoReplyLog("Skip count exceeded. Clearing previous running flag to restart worker.");
        props.deleteProperty("AUTO_WORKER_RUNNING");
        props.deleteProperty("AUTO_WORKER_SKIP_COUNT");
        // continue to start a fresh worker
      } else {
        lock.releaseLock();
        return;
      }
    } else {
      // stale flag or old timestamp — clear and continue
      writeAutoReplyLog("Found stale AUTO_WORKER_RUNNING. Clearing and continuing.");
      props.deleteProperty("AUTO_WORKER_RUNNING");
      props.deleteProperty("AUTO_WORKER_SKIP_COUNT");
    }
  }

  // Mark as running with timestamp and reset skip counter
  props.setProperty("AUTO_WORKER_RUNNING", "YES|" + Date.now());
  props.setProperty("AUTO_WORKER_SKIP_COUNT", "0");

  try {
    writeAutoReplyLog("========== WORKER START ==========");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const MY_EMAIL = Session.getActiveUser().getEmail().toLowerCase();

    /* ================= AUTO REPLY SHEET ================= */
    const replySheet = ss.getSheetByName("Auto Reply");
    if (!replySheet || replySheet.getLastRow() < 2) {
      writeAutoReplyLog("Auto Reply sheet missing or empty. Exit.");
      return;
    }

    let lastHistoryId = replySheet.getRange("G1").getValue();
    let todayCount = Number(replySheet.getRange("H1").getValue() || 0);

    const today = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );
    
    // Get sheet date safely and reset only when sheet date is older than today
    let sheetDateRaw = replySheet.getRange("I1").getValue();
    let sheetDate = "";
    if (sheetDateRaw instanceof Date) {
      sheetDate = Utilities.formatDate(
        sheetDateRaw,
        Session.getScriptTimeZone(),
        "yyyy-MM-dd"
      );
    } else {
      sheetDate = sheetDateRaw.toString().trim();
    }

    if (!sheetDate || sheetDate < today) {
      todayCount = 0;
      replySheet.getRange("H1").setValue(0);
      replySheet.getRange("I1").setValue(today);
      writeAutoReplyLog("Daily reply count reset (sheet date older than today)");
    }

    // Skip all sends if quota was already exhausted today
    if (props.getProperty("EMAIL_QUOTA_HIT_DATE") === today) {
      writeAutoReplyLog("Email quota was already hit today. Skipping sends until tomorrow.");
      return;
    }

    /* ================= Auto Reply Data ================= */
    let replyLogSheet = ss.getSheetByName("Auto Reply Data");
    if (!replyLogSheet) {
      replyLogSheet = ss.insertSheet("Auto Reply Data");
      replyLogSheet.appendRow([
        "Mail Date",
        "Thread ID",
        "Name",
        "Reply-To",
        "Subject",
        "Keyword_Matched"
      ]);
      writeAutoReplyLog("Auto Reply Data sheet created");
    }

    // Ensure Default column header exists
    if (replyLogSheet.getLastColumn() < 7) {
      replyLogSheet.insertColumnAfter(6);
      replyLogSheet.getRange(1, 7).setValue("Default(No Keyword Match)");
    }

    const loggedThreadIds = new Set();
    const kwRepliedIds    = new Set(); // threads already sent keyword replies — duplicate guard
    const defRepliedIds   = new Set(); // threads already sent default replies — duplicate guard
    if (replyLogSheet.getLastRow() > 1) {
      const allLogData = replyLogSheet
        .getRange(2, 1, replyLogSheet.getLastRow() - 1, 7)
        .getValues();
      allLogData.forEach(function(row) {
        const tId       = row[1];
        const kwMatched = row[5];
        const defStatus = row[6];
        if (tId) loggedThreadIds.add(tId);
        if (tId && kwMatched === "YES") kwRepliedIds.add(tId);
        if (tId && defStatus === "YES") defRepliedIds.add(tId);
      });
    }
    writeAutoReplyLog("Logged thread IDs: " + loggedThreadIds.size + " | kw replied: " + kwRepliedIds.size + " | def replied: " + defRepliedIds.size);

    /* ================= FIRST RUN ================= */
    if (!lastHistoryId) {
      const profile = Gmail.Users.getProfile("me");
      replySheet.getRange("G1").setValue(profile.historyId);
      writeAutoReplyLog("First run detected. HistoryId saved: " + profile.historyId);
      // Don't return — still run backfill and pending replies below
    }

    // Always load Drive keywords (needed for backfill + pending + history)
    const driveKeywords = getDriveKeywords();

    /* ================= BACKFILL OLD MAILS ================= */
    backfillOldMails(replyLogSheet, loggedThreadIds, MY_EMAIL);

    // ===== REPROCESS PENDING NO REPLIES (SHEET BASED) =====
    const runState = { sent: 0 };
    todayCount = processPendingNoReplies(
      replyLogSheet,
      driveKeywords,
      todayCount,
      replySheet,
      runState,
      kwRepliedIds,
      defRepliedIds
    );

    /* ================= FETCH HISTORY (paginated) ================= */
    if (!lastHistoryId) {
      // First run — nothing to fetch in history yet
      writeAutoReplyLog("========== WORKER END (first run) ==========");
      return;
    }

    writeAutoReplyLog("Fetching Gmail history from ID: " + lastHistoryId);
    let allHistoryItems = [];
    let latestHistoryId = lastHistoryId;
    let pageToken = null;
    do {
      const params = {
        startHistoryId: lastHistoryId,
        historyTypes:   ["messageAdded"],
        maxResults:     500
      };
      if (pageToken) params.pageToken = pageToken;
      let histResp;
      try {
        histResp = Gmail.Users.History.list("me", params);
      } catch (e) {
        writeAutoReplyLog("History fetch error: " + e.message);
        break;
      }
      if (histResp.history) allHistoryItems = allHistoryItems.concat(histResp.history);
      if (histResp.historyId) latestHistoryId = histResp.historyId;
      pageToken = histResp.nextPageToken || null;
    } while (pageToken);

    if (allHistoryItems.length === 0) {
      writeAutoReplyLog("No new history events found");
      replySheet.getRange("G1").setValue(latestHistoryId);
      writeAutoReplyLog("HistoryId updated: " + latestHistoryId);
      writeAutoReplyLog("========== WORKER END ==========");
      return;
    }

    writeAutoReplyLog("History events found: " + allHistoryItems.length);

    /* ================= PROCESS HISTORY ================= */
    const histories = allHistoryItems;
    let stopProcessing = false;
    for (let hi = 0; hi < histories.length && !stopProcessing; hi++) {
      const h = histories[hi];
      if (!h.messagesAdded) continue;
      for (let mi = 0; mi < h.messagesAdded.length; mi++) {
        const item = h.messagesAdded[mi];

        if (todayCount >= DAILY_REPLY_LIMIT) {
          writeAutoReplyLog("Daily reply limit reached. Skipping further replies.");
          stopProcessing = true;
          break;
        }
        if (runState.sent >= PER_RUN_LIMIT) {
          writeAutoReplyLog("Per-run send limit reached. Skipping further replies this run.");
          stopProcessing = true;
          break;
        }

        let msg;
        try {
          msg = Gmail.Users.Messages.get("me", item.message.id, {
            format: "metadata",
            metadataHeaders: ["Subject", "From", "Reply-To", "Date"]
          });
        } catch (e) {
          writeAutoReplyLog("Skipped message (not found): " + item.message.id);
          continue;
        }

        const headers = {};
        msg.payload.headers.forEach(hd => headers[hd.name] = hd.value);

        /* ===== SELF-SENT & NON-INBOX FILTER ===== */
        if ((headers["From"] || "").toLowerCase().includes(MY_EMAIL)) {
          writeAutoReplyLog("Skipped self-sent mail: " + msg.id);
          continue;
        }
        if (!msg.labelIds || !msg.labelIds.includes("INBOX")) {
          writeAutoReplyLog("Skipped non-inbox mail: " + msg.id);
          continue;
        }

        const subject = (headers.Subject || "").trim();
        const subjectLower = subject.toLowerCase();


        // Only process if mail is from one of the specified IndiaMart addresses
        const fromEmail = (headers["From"] || "").toLowerCase();
        const allowedSenders = [
          "buyershelpdesk@indiamart.com",
          "buyershelp+enq@indiamart.com",
          "buyleads@indiamart.com"
        ];
        if (!allowedSenders.some(addr => fromEmail.includes(addr))) {
          writeAutoReplyLog("Skipped non-Indiamart sender: " + fromEmail);
          continue;
        }

        const threadId = msg.threadId;
        if (loggedThreadIds.has(threadId)) {
          writeAutoReplyLog("Skipped already logged thread: " + threadId);
          continue;
        }

        /* ===== NAME + EMAIL ===== */
        const parsed = parseNameAndEmail(
          headers["Reply-To"] || headers["From"] || ""
        );

        /* ===== LOG FIRST ===== */
        const row = replyLogSheet.getLastRow() + 1;
        replyLogSheet.appendRow([
          new Date(headers.Date || Date.now()),
          threadId,
          parsed.name,
          parsed.email,
          subject,
          "NO"
        ]);
        loggedThreadIds.add(threadId);

        writeAutoReplyLog("New mail logged. ThreadId: " + threadId);

        /* ===== KEYWORD MATCH (Drive folders only) ===== */
        const isKapila = subjectLower.includes("kapila pashu aahar");
        const isAmul   = subjectLower.includes("amul cattel feed");
        const drivekw  = (!isKapila && !isAmul)
                         ? findBestDriveKeyword(subjectLower, driveKeywords)
                         : null;

        if (!isKapila && !isAmul && !drivekw) {
          // ---- No keyword — send default reply immediately ----
          writeAutoReplyLog("No keyword matched. Sending default reply. ThreadId: " + threadId);

          // Duplicate guard: only send if not already replied
          if (defRepliedIds.has(threadId)) {
            writeAutoReplyLog("Default duplicate guard: default already replied to " + threadId + ". Marking col7 YES.");
            replyLogSheet.getRange(row, 7).setValue("YES");
            continue;
          }

          const defaultFolder  = findFolderByKeyword(DRIVE_PARENT_FOLDER_ID, "Default Reply");
          const defAttachments = defaultFolder ? getImageBlobsFromFolder(defaultFolder) : [];
          const defBody        = buildDefaultEmailBody(parsed.name);
          try {
            const defThread = GmailApp.getThreadById(threadId);
            if (defThread) {
              defThread.reply(defBody, { subject: "Message from 55Carat", to: parsed.email, attachments: defAttachments.length ? defAttachments : undefined });
            } else {
              GmailApp.sendEmail(parsed.email, "Message from 55Carat", defBody, defAttachments.length ? { attachments: defAttachments } : {});
            }
            replyLogSheet.getRange(row, 7).setValue("YES");
            defRepliedIds.add(threadId);
            todayCount++;
            runState.sent++;
            replySheet.getRange("H1").setValue(todayCount);
            writeAutoReplyLog("Default reply sent to " + parsed.email + " | runSent: " + runState.sent);
            if (runState.sent >= PER_RUN_LIMIT) {
              writeAutoReplyLog("Per-run send limit reached after default reply. Stopping this run.");
              stopProcessing = true;
              break;
            }
          } catch (e) {
            if (isEmailQuotaError(e)) {
              writeAutoReplyLog("Email quota hit (default in history). Stopping sends for today.");
              todayCount = DAILY_REPLY_LIMIT;
              replySheet.getRange("H1").setValue(todayCount);
              props.setProperty("EMAIL_QUOTA_HIT_DATE", today);
              stopProcessing = true;
              break;
            }
            writeAutoReplyLog("Error sending default reply: " + e.message);
          }
          continue;
        }

        // ---- Duplicate guard: never send keyword reply twice ----
        if (kwRepliedIds.has(threadId)) {
          writeAutoReplyLog("History duplicate guard: keyword already replied to " + threadId + ". Marking sheet YES.");
          replyLogSheet.getRange(row, 6).setValue("YES");
          continue;
        }

        let matchedKeyword, body, customSubject;

        if (isKapila || isAmul) {
          matchedKeyword = isKapila ? "kapila pashu aahar" : "amul cattel feed";
          body           = isKapila ? buildKapilaMailBody(parsed.name) : buildAmulMailBody(parsed.name);
          customSubject  = isKapila ? "Kapila Pashu Aahar - Details" : "Amul Cattel Feed - Details";
        } else {
          matchedKeyword = drivekw;
          const prettyKeyword = drivekw.replace(/\b\w/g, c => c.toUpperCase());
          body          = buildEmailBody(parsed.name, prettyKeyword);
          customSubject = "Message from 55Carat Regarding " + prettyKeyword;
        }

        writeAutoReplyLog("Keyword matched: " + matchedKeyword);

        /* ===== DRIVE IMAGES ===== */
        const folder = findFolderByKeyword(DRIVE_PARENT_FOLDER_ID, matchedKeyword);
        const attachments = folder ? getImageBlobsFromFolder(folder) : [];
        writeAutoReplyLog("Drive images found: " + attachments.length);

        /* ===== SEND EMAIL ===== */
        try {
          const thread = GmailApp.getThreadById(threadId);
          if (thread) {
            thread.reply(body, { subject: customSubject, to: parsed.email, attachments: attachments.length ? attachments : undefined });
          } else {
            GmailApp.sendEmail(parsed.email, customSubject, body, attachments.length ? { attachments } : {});
          }
          todayCount++;
          runState.sent++;
          kwRepliedIds.add(threadId);
          replySheet.getRange("H1").setValue(todayCount);
          replyLogSheet.getRange(row, 6).setValue("YES");
          writeAutoReplyLog("Reply sent to " + parsed.email + " | Subject: " + customSubject + " | images: " + attachments.length + " | runSent: " + runState.sent);
          if (runState.sent >= PER_RUN_LIMIT) {
            writeAutoReplyLog("Per-run send limit reached in history processing. Stopping further sends this run.");
            stopProcessing = true;
            break;
          }
        } catch (e) {
          if (isEmailQuotaError(e)) {
            writeAutoReplyLog("Email quota hit. Stopping sends for today.");
            todayCount = DAILY_REPLY_LIMIT;
            replySheet.getRange("H1").setValue(todayCount);
            props.setProperty("EMAIL_QUOTA_HIT_DATE", today);
            stopProcessing = true;
            break;
          } else {
            writeAutoReplyLog("Error sending reply: " + e.message);
          }
        }
      }
    }

    /* ================= SAVE HISTORY ID ================= */
    if (latestHistoryId && latestHistoryId !== lastHistoryId) {
      replySheet.getRange("G1").setValue(latestHistoryId);
      writeAutoReplyLog("HistoryId updated to: " + latestHistoryId);
    } else {
      writeAutoReplyLog("HistoryId unchanged: " + latestHistoryId);
    }

    writeAutoReplyLog("========== WORKER END ==========");

  } catch (err) {
    writeAutoReplyLog("ERROR: " + err.message);
    throw err;
  } finally {
    // Clear running flag and skip counter so future triggers can start
    const finalProps = PropertiesService.getScriptProperties();
    finalProps.deleteProperty("AUTO_WORKER_RUNNING");
    finalProps.deleteProperty("AUTO_WORKER_SKIP_COUNT");
    lock.releaseLock();
  }
}