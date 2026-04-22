/***********************************************************************
 * BACKLOG AUTO REPLY WORKER
 * Processes emails received BEFORE 23-Feb-2026 that were never replied.
 *
 * NOTE: This file shares the same GAS project scope as Auto Reply Worker.js
 *       so it reuses: DAILY_REPLY_LIMIT, DRIVE_PARENT_FOLDER_ID,
 *       writeLog(), parseNameAndEmail(), findFolderByKeyword(),
 *       getImageBlobsFromFolder(), buildEmailBody(), buildKapilaMailBody(),
 *       buildAmulMailBody(), buildDefaultEmailBody()
 ***********************************************************************/

/******************** BACKLOG CONFIG ********************/
const BACKLOG_BATCH_SEND_LIMIT = 10;        // max successful sends per trigger run
const BACKLOG_MAX_FETCH_PER_RUN = 1000;     // max messages inspected in Phase 2 per run
const BACKLOG_SHEET_NAME       = "Backlog Auto Reply Data";

// Only emails received strictly BEFORE this date
const BACKLOG_CUTOFF = new Date("2026-02-23T00:00:00.000Z");

// Gmail search query — IndiaMart senders, before 23-Feb-2026
const BACKLOG_SEARCH_QUERY =
  "from:(buyershelpdesk@indiamart.com OR buyershelp+enq@indiamart.com OR buyleads@indiamart.com) before:2026/02/23";

// Same IndiaMart sender list used in main worker
const BACKLOG_ALLOWED_SENDERS = [
  "buyershelpdesk@indiamart.com",
  "buyershelp+enq@indiamart.com",
  "buyleads@indiamart.com"
];
/********************************************************/


/***********************************************************************
 * NOTE: getDriveKeywords(), findBestDriveKeyword(), isEmailQuotaError()
 * are defined in Auto Reply Worker.js and shared across this project.
 ***********************************************************************/

/******************** BACKLOG LOGGER ********************/
const ENABLE_BACKLOG_LOG = true; // true = Backlog log ON, false = OFF

function writeBacklogLog(message) {
  if (!ENABLE_BACKLOG_LOG) return; // 🚫 logging disabled

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName("Backlog Log Report");
  if (!logSheet) {
    logSheet = ss.insertSheet("Backlog Log Report");
    logSheet.appendRow(["Time", "Log"]);
  }
  logSheet.appendRow([new Date(), "'" + message]);
}
/********************************************************/

/***********************************************************************
 * PHASE-1 HELPER
 * Re-processes rows in "Backlog Auto Reply Data" sheet where
 * Keyword_Matched == "NO".
 * Returns updated { todayCount, sentThisRun }.
 ***********************************************************************/
function backlogProcessPendingNoReplies(backlogSheet, driveKeywords, todayCount, replySheet, sentThisRun) {

  const lastRow = backlogSheet.getLastRow();
  if (lastRow < 2) return { todayCount, sentThisRun };

  // Ensure 7th column (Default) exists
  if (backlogSheet.getLastColumn() < 7) {
    backlogSheet.insertColumnAfter(6);
    backlogSheet.getRange(1, 7).setValue("Default(No Keyword Match)");
  }

  const data = backlogSheet.getRange(2, 1, lastRow - 1, 7).getValues();
  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  /* ===========================================================
   * PASS 1 — Keyword / Kapila / Amul matches ONLY
   * We exhaust keyword-matched rows first so Kapila / Amul mails
   * never get blocked by default-only rows consuming the budget.
   * =========================================================== */
  writeBacklogLog("[BACKLOG] Phase-1 Pass-1: keyword/Kapila/Amul rows");
  for (let i = 0; i < data.length; i++) {

    if (props.getProperty("BACKLOG_STOP_FLAG") === "true") {
      writeBacklogLog("[BACKLOG] Stop flag detected in Phase-1 Pass-1. Exiting.");
      return { todayCount, sentThisRun };
    }
    if (todayCount >= DAILY_REPLY_LIMIT || sentThisRun >= BACKLOG_BATCH_SEND_LIMIT) break;

    const [mailDate, threadId, name, replyTo, subject, keyword_matched] = data[i];

    if (keyword_matched === "YES") continue;
    if (keyword_matched !== "NO")  continue;

    const subjectLower = (subject || "").toLowerCase();
    const isKapila     = subjectLower.includes("kapila pashu aahar");
    const isAmul       = subjectLower.includes("amul cattel feed");
    const drivekw      = (!isKapila && !isAmul)
                         ? findBestDriveKeyword(subjectLower, driveKeywords)
                         : null;

    // Pass 1 only handles rows that have any keyword hit
    if (!isKapila && !isAmul && !drivekw) continue;

    /* ----- Special templates: Kapila / Amul ----- */
    if (isKapila || isAmul) {
      writeBacklogLog("[BACKLOG] Phase-1 Pass-1 special keyword. ThreadId: " + threadId);
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
        backlogSheet.getRange(i + 2, 6).setValue("YES");
        todayCount++; sentThisRun++;
        replySheet.getRange("H1").setValue(todayCount);
        writeBacklogLog("[BACKLOG] Special reply sent to: " + replyTo);
      } catch (e) {
        if (isEmailQuotaError(e)) {
          writeBacklogLog("[BACKLOG] Email quota hit (special). Stopping sends.");
          todayCount = DAILY_REPLY_LIMIT;
          replySheet.getRange("H1").setValue(todayCount);
          props.setProperty("EMAIL_QUOTA_HIT_DATE", today);
          return { todayCount, sentThisRun };
        }
        writeBacklogLog("[BACKLOG] Error (special): " + e.message);
      }
      continue;
    }

    /* ----- Drive folder keyword match ----- */
    if (drivekw) {
      writeBacklogLog("[BACKLOG] Phase-1 Pass-1 Drive keyword matched: " + drivekw + " | ThreadId: " + threadId);
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
        backlogSheet.getRange(i + 2, 6).setValue("YES");
        todayCount++; sentThisRun++;
        replySheet.getRange("H1").setValue(todayCount);
        writeBacklogLog("[BACKLOG] Drive-keyword reply sent to: " + replyTo);
      } catch (e) {
        if (isEmailQuotaError(e)) {
          writeBacklogLog("[BACKLOG] Email quota hit (Drive keyword). Stopping sends.");
          todayCount = DAILY_REPLY_LIMIT;
          replySheet.getRange("H1").setValue(todayCount);
          props.setProperty("EMAIL_QUOTA_HIT_DATE", today);
          return { todayCount, sentThisRun };
        }
        writeBacklogLog("[BACKLOG] Error (Drive keyword): " + e.message);
      }
    }
  }

  /* ===========================================================
   * PASS 2 — Default (no keyword) rows – only if budget remains
   * =========================================================== */
  if (sentThisRun < BACKLOG_BATCH_SEND_LIMIT && todayCount < DAILY_REPLY_LIMIT) {
    writeBacklogLog("[BACKLOG] Phase-1 Pass-2: default rows");
    const defaultFolder = findFolderByKeyword(DRIVE_PARENT_FOLDER_ID, "Default Reply");

    for (let i = 0; i < data.length; i++) {

      if (props.getProperty("BACKLOG_STOP_FLAG") === "true") {
        writeBacklogLog("[BACKLOG] Stop flag detected in Phase-1 Pass-2. Exiting.");
        return { todayCount, sentThisRun };
      }
      if (todayCount >= DAILY_REPLY_LIMIT || sentThisRun >= BACKLOG_BATCH_SEND_LIMIT) break;

      const [mailDate, threadId, name, replyTo, subject, keyword_matched, defaultStatus] = data[i];

      if (keyword_matched === "YES") continue;   // fully done
      if (keyword_matched !== "NO")  continue;   // unexpected
      if (defaultStatus   === "YES") continue;   // default already sent

      const subjectLower = (subject || "").toLowerCase();
      const isKapila     = subjectLower.includes("kapila pashu aahar");
      const isAmul       = subjectLower.includes("amul cattel feed");
      const drivekw      = (!isKapila && !isAmul)
                           ? findBestDriveKeyword(subjectLower, driveKeywords)
                           : null;

      // Pass 2 only handles rows with truly NO keyword match at all
      if (isKapila || isAmul || drivekw) continue;

      writeBacklogLog("[BACKLOG] Phase-1 Pass-2 default reply. ThreadId: " + threadId);
      const attachments = defaultFolder ? getImageBlobsFromFolder(defaultFolder) : [];
      const body        = buildDefaultEmailBody(name);
      try {
        const thread = GmailApp.getThreadById(threadId);
        if (thread) {
          thread.reply(body, { subject: "Message from 55Carat", to: replyTo, attachments: attachments.length ? attachments : undefined });
        } else {
          GmailApp.sendEmail(replyTo, "Message from 55Carat", body, attachments.length ? { attachments } : {});
        }
        backlogSheet.getRange(i + 2, 7).setValue("YES"); // Mark Default = YES, keep Keyword_Matched = NO
        todayCount++; sentThisRun++;
        replySheet.getRange("H1").setValue(todayCount);
        writeBacklogLog("[BACKLOG] Default reply sent to: " + replyTo);
      } catch (e) {
        if (isEmailQuotaError(e)) {
          writeBacklogLog("[BACKLOG] Email quota hit (default). Stopping sends.");
          todayCount = DAILY_REPLY_LIMIT;
          replySheet.getRange("H1").setValue(todayCount);
          props.setProperty("EMAIL_QUOTA_HIT_DATE", today);
          return { todayCount, sentThisRun };
        }
        writeBacklogLog("[BACKLOG] Error (default): " + e.message);
      }
    }
  }

  return { todayCount, sentThisRun };
}


/***********************************************************************
 * MAIN ENTRY POINT — triggered by "toggleBacklogAutoReply" in
 * start and stop.js
 ***********************************************************************/
function backlogAutoReplyWorker() {

  // Prevent concurrent executions — only one instance at a time
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) {
    writeBacklogLog("[BACKLOG] Another instance is already running. Skipping this run.");
    return;
  }

  const props = PropertiesService.getScriptProperties();

  // Honor the stop flag set by toggleBacklogAutoReply
  if (props.getProperty("BACKLOG_STOP_FLAG") === "true") {
    writeBacklogLog("[BACKLOG] Stop flag is set. Skipping run.");
    return;
  }

  writeBacklogLog("========== BACKLOG WORKER START ==========");

  try {
    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const MY_EMAIL  = Session.getActiveUser().getEmail().toLowerCase();

    /* ==================== SHARED AUTO REPLY SHEET ==================== */
    const replySheet = ss.getSheetByName("Auto Reply");
    if (!replySheet || replySheet.getLastRow() < 2) {
      writeBacklogLog("[BACKLOG] Auto Reply sheet missing or empty. Exit.");
      return;
    }

    // Shared daily counter (same H1 cell as Auto Reply Worker uses)
    let todayCount = Number(replySheet.getRange("H1").getValue() || 0);
    const today    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

    if (props.getProperty("REPLY_COUNT_DATE") !== today) {
      todayCount = 0;
      replySheet.getRange("H1").setValue(0);
      props.setProperty("REPLY_COUNT_DATE", today);
      props.deleteProperty("EMAIL_QUOTA_HIT_DATE"); // new day — clear quota flag
      writeBacklogLog("[BACKLOG] Daily reply count reset.");
    }

    // If quota was already exhausted today, skip Phase 1 sends but still allow Phase 2 to SAVE rows
    const quotaHitToday = props.getProperty("EMAIL_QUOTA_HIT_DATE") === today;
    if (quotaHitToday) {
      writeBacklogLog("[BACKLOG] Email quota was already hit today. Phase 1 sends skipped until tomorrow.");
    }

    if (todayCount >= DAILY_REPLY_LIMIT) {
      writeBacklogLog("[BACKLOG] Daily limit already reached (" + DAILY_REPLY_LIMIT + "). Exit.");
      return;
    }

    /* ==================== DRIVE KEYWORDS (sole source) ==================== */
    // Keywords come only from Drive subfolder names — no sheet rules needed
    const driveKeywords = getDriveKeywords();
    writeBacklogLog("[BACKLOG] Drive keywords loaded: " + driveKeywords.length);

    /* ==================== BACKLOG SHEET ==================== */
    let backlogSheet = ss.getSheetByName(BACKLOG_SHEET_NAME);
    if (!backlogSheet) {
      backlogSheet = ss.insertSheet(BACKLOG_SHEET_NAME);
      backlogSheet.appendRow([
        "Mail Date",
        "Thread ID",
        "Name",
        "Reply-To",
        "Subject",
        "Keyword_Matched",
        "Default(No Keyword Match)"
      ]);
      writeBacklogLog("[BACKLOG] Sheet '" + BACKLOG_SHEET_NAME + "' created.");
    }

    // Ensure column 7 header exists
    if (backlogSheet.getLastColumn() < 7) {
      backlogSheet.insertColumnAfter(6);
      backlogSheet.getRange(1, 7).setValue("Default(No Keyword Match)");
    }

    /* ==================== LOAD KNOWN THREAD IDs ==================== */
    const loggedThreadIds = new Set();
    if (backlogSheet.getLastRow() > 1) {
      backlogSheet
        .getRange(2, 2, backlogSheet.getLastRow() - 1, 1)
        .getValues()
        .flat()
        .forEach(id => { if (id) loggedThreadIds.add(String(id)); });
    }
    writeBacklogLog("[BACKLOG] Known thread IDs in sheet: " + loggedThreadIds.size);

    let sentThisRun = 0;

    /* ================================================================
     * PHASE 1 – Re-process existing rows with Keyword_Matched == "NO"
     * Only runs when quota is NOT already exhausted.
     * ============================================================== */
    if (!quotaHitToday) {
      writeBacklogLog("[BACKLOG] === Phase 1: Reprocessing Keyword_Matched==NO rows ===");
      const p1 = backlogProcessPendingNoReplies(
        backlogSheet, driveKeywords, todayCount, replySheet, sentThisRun
      );
      todayCount  = p1.todayCount;
      sentThisRun = p1.sentThisRun;
      writeBacklogLog("[BACKLOG] Phase 1 done. sentThisRun=" + sentThisRun + " todayCount=" + todayCount);
    } else {
      writeBacklogLog("[BACKLOG] Phase 1 skipped (quota hit today).");
    }

    /* ================================================================
     * PHASE 2 – Fetch old emails from Gmail (if budget left)
     * ============================================================== */
    const fetchDone = props.getProperty("BACKLOG_FETCH_DONE") === "true";

    if (!fetchDone && todayCount < DAILY_REPLY_LIMIT) {

      writeBacklogLog("[BACKLOG] === Phase 2: Fetching old Gmail messages ===");

      let fetchedThisRun = 0; // total messages inspected this run

      const savedPageToken = props.getProperty("BACKLOG_PAGE_TOKEN") || null;

      const listParams = {
        q         : BACKLOG_SEARCH_QUERY,
        maxResults: 50   // smaller pages so stop flag is checked more often
      };
      if (savedPageToken) {
        listParams.pageToken = savedPageToken;
        writeBacklogLog("[BACKLOG] Resuming from saved page token.");
      }

      let listResponse;
      try {
        listResponse = Gmail.Users.Messages.list("me", listParams);
      } catch (e) {
        writeBacklogLog("[BACKLOG] Gmail list error: " + e.message);
        return;
      }

      const messages = listResponse.messages || [];
      writeBacklogLog("[BACKLOG] Messages on this page: " + messages.length);

      if (messages.length === 0) {
        // Nothing returned – mark fetch done
        props.deleteProperty("BACKLOG_PAGE_TOKEN");
        props.setProperty("BACKLOG_FETCH_DONE", "true");
        writeBacklogLog("[BACKLOG] No messages returned. Marking fetch done.");
      }

      for (let i = 0; i < messages.length; i++) {

        // Enforce per-run fetch cap
        if (fetchedThisRun >= BACKLOG_MAX_FETCH_PER_RUN) {
          writeBacklogLog("[BACKLOG] Per-run fetch limit (" + BACKLOG_MAX_FETCH_PER_RUN + ") reached. Pausing until next run.");
          break;
        }
        fetchedThisRun++;

        // Check stop flag on every message so a manual stop takes effect quickly
        if (props.getProperty("BACKLOG_STOP_FLAG") === "true") {
          writeBacklogLog("[BACKLOG] Stop flag detected inside fetch loop. Halting.");
          return;
        }

        /* ---- Fetch message metadata ---- */
        let msg;
        try {
          msg = Gmail.Users.Messages.get("me", messages[i].id, {
            format         : "metadata",
            metadataHeaders: ["Subject", "From", "Reply-To", "Date"]
          });
        } catch (e) {
          writeBacklogLog("[BACKLOG] Skipped message (not found): " + messages[i].id);
          continue;
        }

        const headers = {};
        msg.payload.headers.forEach(hd => { headers[hd.name] = hd.value; });

        /* ---- Skip self-sent ---- */
        if ((headers["From"] || "").toLowerCase().includes(MY_EMAIL)) continue;

        /* ---- Skip non-inbox ---- */
        if (!msg.labelIds || !msg.labelIds.includes("INBOX")) continue;

        /* ---- Only IndiaMart senders ---- */
        const fromEmail = (headers["From"] || "").toLowerCase();
        if (!BACKLOG_ALLOWED_SENDERS.some(addr => fromEmail.includes(addr))) continue;

        /* ---- Date guard: strictly before BACKLOG_CUTOFF ---- */
        const mailDate = new Date(headers["Date"] || 0);
        if (mailDate >= BACKLOG_CUTOFF) {
          writeBacklogLog("[BACKLOG] Skipped mail on/after cutoff: " + mailDate.toISOString());
          continue;
        }

        /* ---- Duplicate thread guard ---- */
        const threadId = msg.threadId;
        if (loggedThreadIds.has(threadId)) continue;

        /* ---- Parse sender ---- */
        const subject  = (headers["Subject"] || "").trim();
        const subjectLower = subject.toLowerCase();
        const parsed   = parseNameAndEmail(headers["Reply-To"] || headers["From"] || "");

        /* ---- Save to sheet immediately (Keyword_Matched = NO) ---- */
        const row = backlogSheet.getLastRow() + 1;
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
        writeBacklogLog("[BACKLOG] Saved: " + threadId + " | " + subject);

        /* ---- Attempt to send reply if budget allows ---- */
        if (!quotaHitToday && sentThisRun < BACKLOG_BATCH_SEND_LIMIT && todayCount < DAILY_REPLY_LIMIT) {

          const isKapila = subjectLower.includes("kapila pashu aahar");
          const isAmul   = subjectLower.includes("amul cattel feed");
          const drivekw  = (!isKapila && !isAmul)
                           ? findBestDriveKeyword(subjectLower, driveKeywords)
                           : null;

          /* --- Special: Kapila / Amul (with images) --- */
          if (isKapila || isAmul) {
            const specialKeyword = isKapila ? "kapila pashu aahar" : "amul cattel feed";
            const folder         = findFolderByKeyword(DRIVE_PARENT_FOLDER_ID, specialKeyword);
            const attachments    = folder ? getImageBlobsFromFolder(folder) : [];
            const body           = isKapila ? buildKapilaMailBody(parsed.name) : buildAmulMailBody(parsed.name);
            const subjectLine    = isKapila ? "Kapila Pashu Aahar - Details" : "Amul Cattel Feed - Details";
            try {
              const thread = GmailApp.getThreadById(threadId);
              if (thread) {
                thread.reply(body, {
                  subject    : subjectLine,
                  to         : parsed.email,
                  attachments: attachments.length ? attachments : undefined
                });
              } else {
                GmailApp.sendEmail(parsed.email, subjectLine, body, attachments.length ? { attachments } : {});
              }
              backlogSheet.getRange(row, 6).setValue("YES");
              todayCount++; sentThisRun++;
              replySheet.getRange("H1").setValue(todayCount);
              writeBacklogLog("[BACKLOG] Special reply sent → " + parsed.email + " | " + specialKeyword + " | images: " + attachments.length);
            } catch (e) {
              if (isEmailQuotaError(e)) {
                writeBacklogLog("[BACKLOG] Email quota hit (special). Stopping sends for today.");
                todayCount = DAILY_REPLY_LIMIT;
                replySheet.getRange("H1").setValue(todayCount);
                props.setProperty("EMAIL_QUOTA_HIT_DATE", today);
                sentThisRun = BACKLOG_BATCH_SEND_LIMIT;
              } else {
                writeBacklogLog("[BACKLOG] Error (special): " + e.message);
              }
            }

          /* --- Drive folder keyword (with images) --- */
          } else if (drivekw) {
            const folder        = findFolderByKeyword(DRIVE_PARENT_FOLDER_ID, drivekw);
            const attachments   = folder ? getImageBlobsFromFolder(folder) : [];
            const prettyKeyword = drivekw.replace(/\b\w/g, c => c.toUpperCase());
            const subjectLine   = "Message from 55Carat Regarding " + prettyKeyword;
            const body          = buildEmailBody(parsed.name, prettyKeyword);
            try {
              const thread = GmailApp.getThreadById(threadId);
              if (thread) {
                thread.reply(body, {
                  subject    : subjectLine,
                  to         : parsed.email,
                  attachments: attachments.length ? attachments : undefined
                });
              } else {
                GmailApp.sendEmail(parsed.email, subjectLine, body, attachments.length ? { attachments } : {});
              }
              backlogSheet.getRange(row, 6).setValue("YES");
              todayCount++; sentThisRun++;
              replySheet.getRange("H1").setValue(todayCount);
              writeBacklogLog("[BACKLOG] Drive-keyword reply sent → " + parsed.email + " | " + drivekw + " | images: " + attachments.length);
            } catch (e) {
              if (isEmailQuotaError(e)) {
                writeBacklogLog("[BACKLOG] Email quota hit (Drive kw). Stopping sends for today.");
                todayCount = DAILY_REPLY_LIMIT;
                replySheet.getRange("H1").setValue(todayCount);
                props.setProperty("EMAIL_QUOTA_HIT_DATE", today);
                sentThisRun = BACKLOG_BATCH_SEND_LIMIT;
              } else {
                writeBacklogLog("[BACKLOG] Error (Drive keyword): " + e.message);
              }
            }

          /* --- No keyword found — leave as NO, Phase-1 handles next run --- */
          } else {
            writeBacklogLog("[BACKLOG] No keyword matched for thread " + threadId + ". Queued for Phase-1.");
          }
        }
        // Save continues even when budget is exhausted or quota hit
      }

      /* ---- Save pagination position ---- */
      if (listResponse.nextPageToken) {
        props.setProperty("BACKLOG_PAGE_TOKEN", listResponse.nextPageToken);
        writeBacklogLog("[BACKLOG] Next page token saved.");
      } else {
        props.deleteProperty("BACKLOG_PAGE_TOKEN");
        props.setProperty("BACKLOG_FETCH_DONE", "true");
        writeBacklogLog("[BACKLOG] Last Gmail page processed. Fetch marked as done.");
      }

    } else if (fetchDone) {
      writeBacklogLog("[BACKLOG] All Gmail pages already fetched. Skipping Phase 2.");
    } else {
      writeBacklogLog("[BACKLOG] Phase 2 skipped (budget full). sentThisRun=" + sentThisRun);
    }

    /* ================================================================
     * PHASE 3 – Check if everything is completely done → stop trigger
     * ============================================================== */
    const isFetchDoneNow = props.getProperty("BACKLOG_FETCH_DONE") === "true";

    if (isFetchDoneNow) {
      const finalLastRow = backlogSheet.getLastRow();
      let anyPending     = false;

      if (finalLastRow > 1) {
        const statusCols = backlogSheet.getRange(2, 6, finalLastRow - 1, 2).getValues();
        for (let i = 0; i < statusCols.length; i++) {
          const kw  = statusCols[i][0]; // Keyword_Matched
          const def = statusCols[i][1]; // Default(No Keyword Match)
          // Pending = not keyword matched YES AND not default YES
          if (kw !== "YES" && def !== "YES") {
            anyPending = true;
            break;
          }
        }
      }

      if (!anyPending) {
        writeBacklogLog("[BACKLOG] All emails processed. Deleting trigger and stopping.");
        ScriptApp.getProjectTriggers().forEach(t => {
          if (t.getHandlerFunction() === "backlogAutoReplyWorker") {
            ScriptApp.deleteTrigger(t);
          }
        });
        writeBacklogLog("[BACKLOG] ✅ Backlog processing COMPLETE. Trigger removed.");
      } else {
        writeBacklogLog("[BACKLOG] Fetch done but pending rows still exist. Waiting for next trigger run.");
      }
    }

    writeBacklogLog(
      "[BACKLOG] ========== WORKER END | SentThisRun=" + sentThisRun +
      " | TodayTotal=" + todayCount + " =========="
    );

  } catch (err) {
    writeBacklogLog("[BACKLOG] FATAL ERROR: " + err.message);
    throw err;
  } finally {
    lock.releaseLock();
  }
}
