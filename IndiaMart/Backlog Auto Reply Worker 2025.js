/***********************************************************************
 * BACKLOG AUTO REPLY WORKER — BEFORE 2025-04-10
 *
 * Fetches old IndiaMart messages strictly before 2025-04-10 and
 * saves them to a new sheet in the same format as the existing
 * backlog sheet. This worker fetches up to 500 messages per page
 * and processes up to 4 pages per run (with short sleeps between
 * pages). It logs progress to a separate log sheet and removes
 * the trigger once all pages have been fetched.
 ***********************************************************************/

/* ===== CONFIG ===== */
const OLD_BATCH_PAGE_SIZE       = 500; // messages per Gmail page
const OLD_PAGES_PER_RUN         = 4;   // pages to fetch per run
const OLD_BACKLOG_SHEET_NAME    = "Backlog Auto Reply Data (<=2025-04-10)";
const OLD_BACKLOG_LOG_SHEET     = "Backlog Old Log Report";
const OLD_BACKLOG_CUTOFF       = new Date("2025-04-10T00:00:00.000Z"); // strictly before

const OLD_BACKLOG_SEARCH_QUERY =
  "from:(buyershelpdesk@indiamart.com OR buyershelp+enq@indiamart.com OR buyleads@indiamart.com) before:2025/04/10";

const OLD_ALLOWED_SENDERS = [
  "buyershelpdesk@indiamart.com",
  "buyershelp+enq@indiamart.com",
  "buyleads@indiamart.com"
];

const ENABLE_OLD_BACKLOG_LOG = true;

function writeOldBacklogLog(msg) {
  if (!ENABLE_OLD_BACKLOG_LOG) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName(OLD_BACKLOG_LOG_SHEET);
  if (!logSheet) {
    logSheet = ss.insertSheet(OLD_BACKLOG_LOG_SHEET);
    logSheet.appendRow(["Time", "Log"]);
  }
  // Safely stringify the message so we never append undefined/null values
  let text;
  if (msg === undefined || msg === null) {
    text = "(no message)";
  } else if (typeof msg === 'object') {
    try { text = JSON.stringify(msg); } catch (e) { text = String(msg); }
  } else {
    text = String(msg);
  }
  // Prepend a single quote so Sheets won't interpret it as a formula
  try {
    logSheet.appendRow([new Date(), "'" + text]);
  } catch (e) {
    // If writing to sheet fails, fallback to server log
    try { Logger.log(new Date() + " - Log write failed: " + e.message + " | msg: " + text); } catch (e2) { /* ignore */ }
  }
}

function simpleParseNameAndEmail(inp) {
  const res = { name: "", email: "" };
  if (!inp) return res;
  const m = inp.match(/([^<]*)<([^>]+)>/);
  if (m) {
    res.name = (m[1] || "").trim();
    res.email = (m[2] || "").trim();
  } else if (inp.includes("@")) {
    res.email = inp.trim();
  } else {
    res.name = inp.trim();
  }
  return res;
}

function backlogAutoReplyWorker2025() {
  // Manual start/stop control: running flag stored in script properties
  const props = PropertiesService.getScriptProperties();

  const alreadyStarted = props.getProperty("BACKLOG_2025_STARTED") === "true";
  if (alreadyStarted) {
    // Second invocation -> request stop and clear started flag
    props.setProperty("BACKLOG_2025_STOP_FLAG", "true");
    props.setProperty("BACKLOG_2025_STARTED", "false");
    writeOldBacklogLog("Manual stop requested. Worker will stop shortly. ===== OLD BACKLOG MANUAL STOP =====");
    return;
  }

  // Not already started -> mark started and clear any stop flag
  props.setProperty("BACKLOG_2025_STARTED", "true");
  props.setProperty("BACKLOG_2025_STOP_FLAG", "false");
  writeOldBacklogLog("Manual start requested. ===== OLD BACKLOG MANUAL START =====");

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) {
    writeOldBacklogLog("Another instance is running. Clearing started flag and skipping this run.");
    props.setProperty("BACKLOG_2025_STARTED", "false");
    return;
  }

  // Confirm lock acquired
  writeOldBacklogLog("Lock acquired. Proceeding with run.");

  try {
    if (props.getProperty("BACKLOG_2025_STOP_FLAG") === "true") {
      writeOldBacklogLog("Stop flag set. Skipping run.");
      props.setProperty("BACKLOG_2025_STARTED", "false");
      lock.releaseLock();
      return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Ensure sheet exists and header format matches original backlog
    let backlogSheet = ss.getSheetByName(OLD_BACKLOG_SHEET_NAME);
    if (!backlogSheet) {
      backlogSheet = ss.insertSheet(OLD_BACKLOG_SHEET_NAME);
      backlogSheet.appendRow([
        "Mail Date",
        "Thread ID",
        "Name",
        "Reply-To",
        "Subject",
        "Keyword_Matched",
        "Default(No Keyword Match)"
      ]);
      writeOldBacklogLog("Created sheet: " + OLD_BACKLOG_SHEET_NAME);
    }

    // Load already-logged thread IDs to avoid duplicates
    const loggedThreadIds = new Set();
    if (backlogSheet.getLastRow() > 1) {
      backlogSheet
        .getRange(2, 2, backlogSheet.getLastRow() - 1, 1)
        .getValues()
        .flat()
        .forEach(id => { if (id) loggedThreadIds.add(String(id)); });
    }
    writeOldBacklogLog("Known thread IDs in sheet: " + loggedThreadIds.size);

    // Pagination state
    let pageToken = props.getProperty("BACKLOG_2025_PAGE_TOKEN") || null;
    const fetchDone = props.getProperty("BACKLOG_2025_FETCH_DONE") === "true";

    if (fetchDone) {
      writeOldBacklogLog("Fetch already completed previously. Exiting.");
      return;
    }

    // First-run detection: no saved page token => first run
    const isFirstRun = !pageToken;
    if (isFirstRun) {
      writeOldBacklogLog("===== OLD BACKLOG FIRST RUN START =====");
    } else {
      writeOldBacklogLog("===== OLD BACKLOG WORKER START (resume) =====");
    }

    // Loop up to OLD_PAGES_PER_RUN pages
    for (let page = 0; page < OLD_PAGES_PER_RUN; page++) {
      if (props.getProperty("BACKLOG_2025_STOP_FLAG") === "true") {
        writeOldBacklogLog("Stop flag detected. Halting pagination loop.");
        break;
      }

      // Request one page
      const listParams = { q: OLD_BACKLOG_SEARCH_QUERY, maxResults: OLD_BATCH_PAGE_SIZE };
      if (pageToken) listParams.pageToken = pageToken;

      let listResponse;
      try {
        listResponse = Gmail.Users.Messages.list("me", listParams);
      } catch (e) {
        writeOldBacklogLog("Gmail list error: " + e.message);
        break;
      }

      const messages = listResponse.messages || [];
      writeOldBacklogLog("Page " + (page + 1) + ": messages on this page: " + messages.length);

      if (messages.length === 0) {
        // No messages left — mark fetch done and clear page token
        props.deleteProperty("BACKLOG_2025_PAGE_TOKEN");
        props.setProperty("BACKLOG_2025_FETCH_DONE", "true");
        writeOldBacklogLog("No messages in page — marking fetch done. ===== OLD BACKLOG FINAL RUN STOP =====");
        break;
      }

      // Process each message metadata-only and save to sheet
      for (let i = 0; i < messages.length; i++) {
        if (props.getProperty("BACKLOG_2025_STOP_FLAG") === "true") {
          writeOldBacklogLog("Stop flag detected inside message loop. Exiting.");
          break;
        }

        let msg;
        try {
          msg = Gmail.Users.Messages.get("me", messages[i].id, {
            format: "metadata",
            metadataHeaders: ["Subject", "From", "Reply-To", "Date"]
          });
        } catch (e) {
          writeOldBacklogLog("Skipped message (not found): " + messages[i].id);
          continue;
        }

        const headers = {};
        (msg.payload.headers || []).forEach(hd => { headers[hd.name] = hd.value; });

        // Skip self-sent
        const MY_EMAIL = Session.getActiveUser().getEmail().toLowerCase();
        if ((headers["From"] || "").toLowerCase().includes(MY_EMAIL)) continue;

        // Skip non-inbox
        if (!msg.labelIds || !msg.labelIds.includes("INBOX")) continue;

        // Only allowed senders
        const fromEmail = (headers["From"] || "").toLowerCase();
        if (!OLD_ALLOWED_SENDERS.some(addr => fromEmail.includes(addr))) continue;

        // Date guard: strictly before cutoff
        const mailDate = new Date(headers["Date"] || 0);
        if (mailDate >= OLD_BACKLOG_CUTOFF) {
          writeOldBacklogLog("Skipped mail on/after cutoff: " + mailDate.toISOString());
          continue;
        }

        const threadId = msg.threadId;
        if (loggedThreadIds.has(threadId)) continue;

        const subject = (headers["Subject"] || "").trim();
        const parsed = simpleParseNameAndEmail(headers["Reply-To"] || headers["From"] || "");

        // Append in same format
        backlogSheet.appendRow([
          mailDate,
          threadId,
          parsed.name,
          parsed.email,
          subject,
          "NO",
          ""
        ]);
        loggedThreadIds.add(threadId);
        writeOldBacklogLog("Saved: " + threadId + " | " + subject);
      }

      // Save pagination state
      if (listResponse.nextPageToken) {
        pageToken = listResponse.nextPageToken;
        props.setProperty("BACKLOG_2025_PAGE_TOKEN", pageToken);
        writeOldBacklogLog("Next page token saved.");
        // Sleep briefly so we pace across the minute (approx)
        if (page < OLD_PAGES_PER_RUN - 1) {
          Utilities.sleep(15000); // 15s pause between pages
        }
      } else {
        // No more pages — mark fetch done
        props.deleteProperty("BACKLOG_2025_PAGE_TOKEN");
        props.setProperty("BACKLOG_2025_FETCH_DONE", "true");
        writeOldBacklogLog("Last Gmail page processed. Fetch marked done. ===== OLD BACKLOG FINAL RUN STOP =====");
        break;
      }
    }

    // If fetch done now, delete any triggers that run this worker
    if (props.getProperty("BACKLOG_2025_FETCH_DONE") === "true") {
      writeOldBacklogLog("All pages fetched. ===== OLD BACKLOG FINAL RUN STOP ===== Deleting trigger and stopping.");
      ScriptApp.getProjectTriggers().forEach(t => {
        if (t.getHandlerFunction() === "backlogAutoReplyWorker2025") {
          ScriptApp.deleteTrigger(t);
        }
      });
      writeOldBacklogLog("✅ Old backlog processing COMPLETE. Trigger removed.");
    } else {
      writeOldBacklogLog("Run completed for now; pageToken saved for next run.");
    }

    writeOldBacklogLog("===== OLD WORKER END =====");
    // clear started flag (run finished)
    try { props.setProperty("BACKLOG_2025_STARTED", "false"); } catch (e) { /* ignore */ }

  } catch (err) {
    writeOldBacklogLog("FATAL ERROR: " + (err && err.message ? err.message : err));
    throw err;
  } finally {
    try { props.setProperty("BACKLOG_2025_STARTED", "false"); } catch (e) { /* ignore */ }
    lock.releaseLock();
  }
}

/* Toggle helper for stop flag */
function toggleOldBacklogStopFlag() {
  const props = PropertiesService.getScriptProperties();
  const cur = props.getProperty("BACKLOG_2025_STOP_FLAG") === "true";
  props.setProperty("BACKLOG_2025_STOP_FLAG", (!cur).toString());
  writeOldBacklogLog("Toggled STOP_FLAG to " + (!cur));
}
