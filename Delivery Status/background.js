const SHEET_NAME = "All Orders";
const MAX_PER_CYCLE = 200;

let isProcessing = false;

console.log("Background Loaded");

let shouldStop = false;

function todayString() {
    return (new Date()).toISOString().slice(0,10);
}


/* Restore running on startup */
chrome.runtime.onStartup.addListener(restoreRunning);
chrome.runtime.onInstalled.addListener(restoreRunning);

function restoreRunning() {
    chrome.storage.local.get("running", (data) => {
        if (data.running) {
            chrome.alarms.create("loop", { periodInMinutes: 1 });
            log("Restored Running State");
        }
    });

    // Auto-start when Chrome profile launches (if not completed/stopped for today).
    maybeAutoStartOnStartup();
}

function startProcessingFromTrigger(triggerLabel = "manual") {
    shouldStop = false;
    const isManualStart = String(triggerLabel || "").toLowerCase().includes("manual");

    // Respect per-day completion: if we've already completed today and no progress marker, don't start.
    chrome.storage.local.get(["lastCompleteDate", "lastProcessedRow", "currentSheetIndex"], (data) => {
        const today = todayString();
        const lastCompleteDate = data.lastCompleteDate || "";
        const lastProcessedRow = Number(data.lastProcessedRow) || 0;
        const currentSheetIndex = Number(data.currentSheetIndex) || 0;

        // Auto/background starts should respect daily completion, manual starts should not.
        if (!isManualStart && lastCompleteDate === today && lastProcessedRow === 0 && currentSheetIndex === 0) {
            log("Start requested but already completed today — skipping start");
            chrome.storage.local.set({ running: false, runState: "completed", lastRunDate: today });
            return;
        }

        const now = Date.now();
        const startState = {
            running: true,
            lastRunTime: now,
            lastRunDate: today,
            runState: "running"
        };

        // Manual start intentionally clears completion marker so user can re-run same day.
        if (isManualStart) {
            startState.lastCompleteDate = "";
        }

        chrome.storage.local.set(startState, () => {
            chrome.alarms.clear("loop", () => {
                chrome.alarms.create("loop", { periodInMinutes: 1 });
                log(`Process Started (${triggerLabel})`);
                processSheet();
            });
        });
    });
}

function maybeAutoStartOnStartup() {
    chrome.storage.local.get([
        "running",
        "sheetId",
        "sheetIds",
        "lastCompleteDate"
    ], async (data) => {
        const today = todayString();
        const lastCompleteDate = data.lastCompleteDate || "";

        if (data.running) return;
        if (lastCompleteDate === today) return;

        let sheetIds = data.sheetIds || [];
        if (!Array.isArray(sheetIds) || sheetIds.length === 0) {
            if (data.sheetId) {
                sheetIds = [data.sheetId];
            }
        }

        if (sheetIds.length === 0) return;

        try {
            // Startup runs without user gesture, so only silent token retrieval is allowed.
            await getToken({ interactive: false });
            startProcessingFromTrigger("startup");
        } catch (err) {
            if (err && (err.code === "LOGIN_REQUIRED" || err.message === "LOGIN_REQUIRED")) {
                log("Auto-start paused: login required. Open extension and click Start.");
            } else if (err && err.code === "BAD_CLIENT_ID") {
                chrome.storage.local.set({ running: false, runState: "oauth-config-error" });
                log("OAuth configuration error: bad client id for this extension ID. Reconfigure Google OAuth client for this extension.");
            } else {
                log("Auto-start failed: " + (err?.message || "Unknown error"));
            }
        }
    });
}

/* Update Sheet cell AB (status) helper */
async function updateSheet(row, status, token, sheetId) {
    const updateRange = `${SHEET_NAME}!AB${row}`;
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${updateRange}?valueInputOption=RAW`;

    await fetch(url, {
        method: "PUT",
        headers: {
            "Authorization": `Bearer ${token}`,
            "Content-Type": "application/json"
        },
        body: JSON.stringify({ values: [[status]] })
    });

    log("Updated Row " + row);
}

/* Update arbitrary single cell helper */
async function updateCell(row, colLetter, value, token, sheetId) {
    if (value === undefined || value === null) return;
    const updateRange = `${SHEET_NAME}!${colLetter}${row}`;
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${updateRange}?valueInputOption=RAW`;

    await fetch(url, {
        method: "PUT",
        headers: {
            "Authorization": `Bearer ${token}`,
            "Content-Type": "application/json"
        },
        body: JSON.stringify({ values: [[value]] })
    });

    log(`Updated Row ${row} Col ${colLetter}`);
}

/* Start / Stop */
chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {

    if (msg.action === "authenticate") {
        getToken({ interactive: true })
            .then(() => sendResponse({ ok: true }))
            .catch((err) => sendResponse({ ok: false, code: err?.code || "AUTH_FAILED", error: err?.message || "Authentication failed" }));
        return true;
    }

    if (msg.action === "start") {
        startProcessingFromTrigger(msg.source || "manual");
        sendResponse({ ok: true });
        return true;
    }

    if (msg.action === "stop") {
        shouldStop = true;

        // mark stopped and reset markers so next startup will process from start again on same day
        chrome.storage.local.set({
            running: false,
            currentSheetIndex: 0,
            lastProcessedRow: 0,
            lastRunTime: 0,
            lastCompleteDate: "",
            manualStopDate: "",
            runState: "manual-stopped"
        }, () => {
            chrome.alarms.clear("loop");
            log("Process Stopped and progress reset");
        });
        sendResponse({ ok: true });
        return true;
    }

    return true;
});


/* Alarm loop */
chrome.alarms.onAlarm.addListener((alarm) => {
    if (alarm.name === "loop") {
        chrome.storage.local.get("running", (data) => {
            if (data.running && !isProcessing) {
                processSheet();
            }
        });
    }
});

/* Logging */
function log(text) {
    const time = new Date().toLocaleTimeString();
    const message = `[${time}] ${text}`;

    chrome.storage.local.get("logs", (data) => {
        let logs = data.logs || [];
        logs.push(message);
        if (logs.length > 200) logs.shift();
        chrome.storage.local.set({ logs: logs });
    });

    chrome.runtime.sendMessage({ type: "log", text: message }).catch(() => {});
}

/* OAuth */
async function getToken(options = {}) {
    const interactive = options.interactive === true;
    return new Promise((resolve, reject) => {
        chrome.identity.getAuthToken({ interactive: interactive }, (token) => {
            if (chrome.runtime.lastError || !token) {
                const rawMessage = (chrome.runtime.lastError?.message || "").toString();
                const normalized = rawMessage.toLowerCase();

                if (normalized.includes("bad client id") || normalized.includes("invalid_client")) {
                    const badClientErr = new Error("OAuth client is not linked to this extension ID. Configure Google OAuth for this extension on this PC.");
                    badClientErr.code = "BAD_CLIENT_ID";
                    reject(badClientErr);
                    return;
                }

                if (!interactive) {
                    const loginErr = new Error("LOGIN_REQUIRED");
                    loginErr.code = "LOGIN_REQUIRED";
                    reject(loginErr);
                } else {
                    const authErr = new Error(rawMessage || "Authentication failed");
                    authErr.code = "AUTH_FAILED";
                    reject(authErr);
                }
                return;
            }
            resolve(token);
        });
    });
}

/* Main Process */
async function processSheet() {

    if (isProcessing) return;
    isProcessing = true;

    // check whether we already completed processing today
    chrome.storage.local.get(["lastCompleteDate", "lastProcessedRow", "sheetId", "sheetIds", "currentSheetIndex", "orderIdCol", "orderStatusCol", "siteList"], async (data) => {

        let sheetIds = data.sheetIds || [];
        if (!Array.isArray(sheetIds) || sheetIds.length === 0) {
            if (data.sheetId) {
                sheetIds = [data.sheetId];
            }
        }

        if (sheetIds.length === 0) {
            log("No Sheet IDs Found");
            chrome.storage.local.set({ running: false, runState: "idle" });
            isProcessing = false;
            return;
        }

        const currentSheetIndex = Number(data.currentSheetIndex) || 0;
        const today = todayString();
        
        if (currentSheetIndex >= sheetIds.length) {
            log("All sheets already completed today — stopping");
            chrome.storage.local.set({
                running: false,
                currentSheetIndex: 0,
                lastCompleteDate: today,
                lastRunDate: today,
                runState: "completed"
            }, () => {
                chrome.alarms.clear("loop");
                isProcessing = false;
            });
            return;
        }

        const sheetId = sheetIds[currentSheetIndex];
        const lastCompleteDate = data.lastCompleteDate || "";
        const lastProcessedRow = Number(data.lastProcessedRow) || 0;

        if (lastCompleteDate === today && lastProcessedRow === 0 && currentSheetIndex === 0) {
            // already completed for today — stop and don't run again until next day
            chrome.storage.local.set({ running: false, runState: "completed", lastRunDate: today }, () => {
                chrome.alarms.clear("loop");
                log("Already completed today — skipping run");
                isProcessing = false;
            });
            return;
        }

        try {

            log(`Reading Sheet ${currentSheetIndex + 1}/${sheetIds.length}: ${sheetId.slice(0, 8)}...`);
            const token = await getToken({ interactive: false });

            // load user mappings, site-list
            const orderIdColLetter = (data.orderIdCol || "A").toUpperCase();
            const orderStatusColLetter = (data.orderStatusCol || "D").toUpperCase();
            const rawSiteList = (data.siteList || "").toString().trim();

            // default include sites if none provided
            const defaultSites = [
                "Amazon.com.be",
                "Amazon.pl",
                "Amazon.it",
                "Amazon.nl",
                "Amazon.se",
                "Amazon.com.tr",
                "Amazon.com.mx",
                "Amazon.com.jp",
                "Amazon.ca",
                "Amazon.co.uk",
                "Amazon.com",
                "Amazon.de",
                "Amazon.es",
                "Amazon.fr"
            ];
            const siteLines = rawSiteList ? rawSiteList.split(/\r?\n/) : defaultSites;
            const includedSet = new Set(siteLines.map(s => (s || "").toString().trim().toLowerCase()).filter(s => s));


            // Detect last used row by checking key columns (A = order id, I = sales channel, AA = tracking)
            const headers = { "Authorization": `Bearer ${token}` };
            const urlA = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${SHEET_NAME}!A3:A`;
            const urlAA = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${SHEET_NAME}!AA3:AA`;
            const urlI = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${SHEET_NAME}!I3:I`;

            const [ja, jaa, ji] = await Promise.all([
                fetch(urlA, { headers }).then(r => r.json()).catch(() => ({})),
                fetch(urlAA, { headers }).then(r => r.json()).catch(() => ({})),
                fetch(urlI, { headers }).then(r => r.json()).catch(() => ({}))
            ]);

            const lenA = (ja.values || []).length;
            const lenAA = (jaa.values || []).length;
            const lenI = (ji.values || []).length;

            const maxLen = Math.max(lenA, lenAA, lenI);
            if (maxLen === 0) {
                log("No rows found in this sheet — treating as complete");
                // if sheet is empty, proceed to next
                isProcessing = false;
                completeCurrentSheet(currentSheetIndex, sheetIds);
                return;
            }

            // add 2 because we started at row 3
            const endRow = maxLen + 2;
            const range = `${SHEET_NAME}!A3:AB${endRow}`;
            const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${range}`;

            const res = await fetch(url, { headers });
            const json = await res.json();
            const rows = json.values || [];

            log("Rows Found: " + rows.length);

            let processed = 0;
            let loopCompleted = true; // becomes false if we exit early (rate limit, stop, or error)

            // helper indices
            const orderIdIndex = colLetterToIndex(orderIdColLetter); // usually 0 for A
            const salesChannelIndex = colLetterToIndex("I"); // column I = Sales Channel
            const trackingIndex = colLetterToIndex("AA"); // 26
            const courierIndex = colLetterToIndex("Z"); // column Z = courier name
            const writeStatusIndex = colLetterToIndex("AB"); // 27 (we write 17track status here)

            // determine start index based on lastProcessedRow (resume after that sheet row)
            const startIndex = lastProcessedRow > 0 ? Math.max(0, lastProcessedRow + 1 - 3) : 0;
            
            // iterate rows and persist progress so we can resume on errors/restarts
            for (let i = startIndex; i < rows.length; i++) {

                if (processed >= MAX_PER_CYCLE) { loopCompleted = false; break; }

                if (shouldStop) { log("Stopped Mid Cycle"); loopCompleted = false; break; }

                const row = rows[i];

                const orderId = (row[orderIdIndex] || "").toString().trim();
                const currentStatus = (row[writeStatusIndex] || "").toString().trim();

                // read Sales Channel column and only process rows matching the include whitelist (robust match)
                const salesChannel = (row[salesChannelIndex] || "").toString().trim();
                if (includedSet.size > 0 && !matchesInclude(salesChannel, siteLines)) {
                    // not in include list — skip this row but persist progress
                    const skipRowNumber = i + 3;
                    chrome.storage.local.set({ lastProcessedRow: skipRowNumber });
                    continue;
                }

                if (!orderId) {
                    const skipRowNumber = i + 3;
                    chrome.storage.local.set({ lastProcessedRow: skipRowNumber });
                    continue;
                }

                if (!currentStatus || !currentStatus.toLowerCase().includes("delivered")) {

                    // always check 17track (use AA column for tracking number)
                    const tracking = (row[trackingIndex] || "").toString().trim();
                    const rowNumber = i + 3;
                    if (!tracking) {
                        log(orderId + " has no tracking value — skipping");
                        chrome.storage.local.set({ lastProcessedRow: rowNumber });
                        continue;
                    }

                    const courier = (row[courierIndex] || "").toString().trim();
                    const isRoyal = courier.toString().toLowerCase() === 'royal mail';

                    let resultObj;
                    if (isRoyal) {
                        log("Checking (Royal Mail): " + tracking);
                        resultObj = await fetchRoyalMailStatus(tracking);
                    } else {
                        log("Checking (17track): " + tracking);
                        resultObj = await fetchTrackingStatus(tracking, salesChannel);
                    }

                    // extract text from result (support object with statusText/xpathText)
                    let resultText = "";
                    let xpathText = null;
                    if (resultObj && typeof resultObj === 'object') {
                        resultText = (resultObj.statusText || "").toString();
                        xpathText = resultObj.xpathText || null;
                    } else {
                        resultText = (resultObj || "").toString();
                    }

                    // if tracking result didn't find the status, skip writing to sheet but persist progress
                    const lowerResult = resultText.toLowerCase();
                    if (!resultText || lowerResult.includes("not found") || lowerResult.includes("status not found") || lowerResult.includes("notfound")) {
                        log(tracking + " → Not found — skipping write");
                        chrome.storage.local.set({ lastProcessedRow: rowNumber });
                        continue;
                    }

                    const normResult = normalizeStatus(resultText);
                    log(tracking + " → " + resultText + " => normalized: '" + normResult + "'");

                    await updateSheet(rowNumber, normResult, token, sheetId);

                    // write XPath extracted data into column AC if available (skip for Royal Mail)
                    if (!isRoyal && xpathText && xpathText.toString().trim() !== "") {
                        await updateCell(rowNumber, 'AC', xpathText.toString().trim(), token, sheetId);
                    }

                    if (normResult.toLowerCase().includes("delivered")) {
                        await colorRowGreen(rowNumber, token, sheetId);
                    }

                    processed++;
                    // persist progress after handling this row
                    chrome.storage.local.set({ lastProcessedRow: rowNumber });
                }
            }

            // if we iterated through all fetched rows without breaking early, treat as complete
            if (loopCompleted) {
                isProcessing = false;
                completeCurrentSheet(currentSheetIndex, sheetIds);
            } else {
                log("Cycle Complete (paused/resumed later)");
                isProcessing = false;
            }

        } catch (err) {
            if (err && (err.code === "LOGIN_REQUIRED" || err.message === "LOGIN_REQUIRED")) {
                chrome.storage.local.set({ running: false, runState: "auth-required" }, () => {
                    chrome.alarms.clear("loop");
                });
                log("Login required. Open extension popup and click Start.");
                chrome.runtime.sendMessage({ type: "auth-required" }).catch(() => {});
                isProcessing = false;
                return;
            }

            if (err && err.code === "BAD_CLIENT_ID") {
                chrome.storage.local.set({ running: false, runState: "oauth-config-error" }, () => {
                    chrome.alarms.clear("loop");
                });
                log("OAuth configuration error: bad client id. This extension ID is not authorized in Google Cloud OAuth settings.");
                chrome.runtime.sendMessage({ type: "auth-config-error", text: err.message }).catch(() => {});
                isProcessing = false;
                return;
            }

            log("Error: " + err.message);
            isProcessing = false;
        }
    });
}

/**
 * Handles transitioning to next sheet or finishing all sheets
 */
function completeCurrentSheet(currentIndex, sheetIds) {
    const nextIndex = currentIndex + 1;
    if (nextIndex < sheetIds.length) {
        log(`Sheet ${currentIndex + 1} Complete. Moving to Sheet ${nextIndex + 1}...`);
        chrome.storage.local.set({ lastProcessedRow: 0, currentSheetIndex: nextIndex, runState: "running" }, () => {
            // Trigger processing of the next sheet immediately
            setTimeout(processSheet, 1000);
        });
    } else {
        const nowComplete = Date.now();
        const todayComplete = todayString();
        chrome.storage.local.set({ 
            lastProcessedRow: 0, 
            currentSheetIndex: 0, 
            running: false, 
            lastRunTime: nowComplete, 
            lastRunDate: todayComplete,
            lastCompleteDate: todayComplete,
            runState: "completed"
        }, () => {
            chrome.alarms.clear("loop");
            log("All sheets processed — stopping until manually restarted");
        });
    }
}

/* Tracking Fetch with XPath */
function fetchTrackingStatus(tracking, salesChannel) {

    return new Promise((resolve) => {

        let url = `https://t.17track.net/en#nums=${tracking}`;
        // If the sales channel is Shopify, append fc parameter required by 17track
        try {
            if (salesChannel && salesChannel.toString().toLowerCase().includes('shopify')) {
                url += '&fc=100055';
            }
        } catch (e) {}

        chrome.tabs.create({ url: url, active: false }, (tab) => {

            if (!tab?.id) {
                resolve({ statusText: 'Status Not Found', xpathText: null });
                return;
            }

            const tabId = tab.id;

            setTimeout(() => {

                chrome.scripting.executeScript({
                    target: { tabId: tabId },
                    func: (injectedSalesChannel) => {

                        return new Promise((res) => {

                            let attempts = 0;

                            const interval = setInterval(() => {

                                const el = document.querySelector(
                                    "div.text-sm.text-text-primary.flex.items-center.gap-1"
                                );

                                if (el && el.innerText.trim()) {

                                    clearInterval(interval);
                                    const statusText = el.innerText.trim();
                                    // attempt to fetch XPath value if present
                                    let xpathText = null;
                                    try {
                                        // Try Shopify-specific XPath first when sales channel indicates Shopify
                                        const shopifyXpath = '/html/body/div[3]/div/section/div[2]/div/div/div/div/div/div/div/div[1]/div[1]/div[3]/div[1]/div[2]/span';
                                        const defaultXpath = '/html/body/div[3]/div/section/div[3]/div/div/div/div/div/div/div/div[1]/div/div[3]/div[1]/div[2]/span'

                                        const xpaths = [];
                                        try {
                                            if (injectedSalesChannel && injectedSalesChannel.toString().toLowerCase().includes('shopify')) {
                                                xpaths.push(shopifyXpath);
                                            }
                                        } catch (e) {}
                                        xpaths.push(defaultXpath);

                                        for (let xpath of xpaths) {
                                            try {
                                                const xp = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
                                                if (xp && xp.singleNodeValue && xp.singleNodeValue.textContent && xp.singleNodeValue.textContent.trim()) {
                                                    xpathText = xp.singleNodeValue.textContent.trim();
                                                    break;
                                                }
                                            } catch (e) {}
                                        }
                                    } catch (e) {}

                                    res({ statusText, xpathText });
                                    return;
                                }

                                attempts++;

                                if (attempts > 20) {
                                    clearInterval(interval);
                                    res({ statusText: 'Status Not Found', xpathText: null });
                                }

                            }, 1000);
                        });
                    },
                    args: [salesChannel]
                }, (results) => {

                    let statusObj = { statusText: 'Status Not Found', xpathText: null };

                    if (!chrome.runtime.lastError && results?.[0]) {
                        statusObj = results[0].result;
                    }

                    chrome.tabs.remove(tabId);
                    resolve(statusObj);
                });

            }, 6000);
        });
    });
}


/* Royal Mail tracking: open Royal Mail track page and wait until provided XPath contains text (no timeout)
   Will return { statusText: '...', xpathText: null }
*/
function fetchRoyalMailStatus(tracking) {

    return new Promise((resolve) => {

        const url = `https://www.royalmail.com/track-your-item#/tracking-results/${tracking}`;

        chrome.tabs.create({ url: url, active: false }, (tab) => {

            if (!tab?.id) {
                resolve({ statusText: 'Status Not Found', xpathText: null });
                return;
            }

            const tabId = tab.id;

            // small initial wait to allow page navigation
            setTimeout(() => {

                // execute script that polls the XPath until text appears or timeout (ms) expires
                chrome.scripting.executeScript({
                    target: { tabId: tabId },
                    func: (xpath, timeoutMs) => {
                        return new Promise((res) => {
                            const interval = setInterval(() => {
                                try {
                                    const xp = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
                                    if (xp && xp.singleNodeValue) {
                                        const node = xp.singleNodeValue;
                                        const txt = (node.innerText || node.textContent || '').trim();
                                        if (txt) {
                                            clearInterval(interval);
                                            if (to) clearTimeout(to);
                                            res({ statusText: txt, xpathText: null });
                                        }
                                    }
                                } catch (e) {}
                            }, 1000);

                            // fallback timeout to stop waiting after timeoutMs milliseconds
                            const to = setTimeout(() => {
                                clearInterval(interval);
                                res({ statusText: 'Status Not Found', xpathText: null });
                            }, timeoutMs || 30000);
                        });
                    },
                    args: ['//*[@id="rml_track_and_trace"]/div/div/section/div/div/div/div/div/div[2]/div[2]/h2', 30000]
                }, (results) => {

                    let statusObj = { statusText: 'Status Not Found', xpathText: null };

                    if (!chrome.runtime.lastError && results?.[0]) {
                        statusObj = results[0].result;
                    }

                    chrome.tabs.remove(tabId);
                    resolve(statusObj);
                });

            }, 3000);
        });
    });
}

async function colorRowGreen(row, token, sheetId) {

    const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}:batchUpdate`;

    const body = {
        requests: [
            {
                repeatCell: {
                    range: {
                        sheetId: 0, // ⚠️ if sheetId index different tell me
                        startRowIndex: row - 1,
                        endRowIndex: row,
                    },
                    cell: {
                        userEnteredFormat: {
                            backgroundColor: {
                                red: 0,
                                green: 0.5,
                                blue: 0
                            }
                        }
                    },
                    fields: "userEnteredFormat.backgroundColor"
                }
            }
        ]
    };

    await fetch(url, {
        method: "POST",
        headers: {
            "Authorization": `Bearer ${token}`,
            "Content-Type": "application/json"
        },
        body: JSON.stringify(body)
    });

    log("Row " + row + " Colored Green");
}

/* Helpers */
function colLetterToIndex(letter) {
    // Convert column letters (A, B, Z, AA, AB ...) to zero-based index
    letter = (letter || "").toUpperCase().replace(/[^A-Z]/g, "");
    let index = 0;
    for (let i = 0; i < letter.length; i++) {
        index = index * 26 + (letter.charCodeAt(i) - 64);
    }
    return index - 1;
}

function normalizeKey(s) {
    if (!s && s !== 0) return "";
    // keep only alphanumeric and uppercase for stable matching
    return String(s).toUpperCase().replace(/[^A-Z0-9]/g, "").trim();
}

// uploaded-file parsing removed

function normalizeStatus(raw) {
    if (!raw && raw !== 0) return "";
    let s = String(raw).trim();

    if (s.toLowerCase().includes("alert")) {
        return s;
    }

    // remove parenthetical notes e.g. "Delivered (7 Days)" -> "Delivered"
    s = s.replace(/\([^)]*\)/g, "").trim();

    const lower = s.toLowerCase();

    if (/delivered/i.test(lower)) return "Delivered";
    if (/picked\s*up|pickedup/i.test(lower)) return "Picked Up";
    if (/out\s*for\s*delivery|outfor|out\s*for/i.test(lower)) return "Out For Delivery";

    // Keep certain statuses as-is (Cancelled, Pending, etc.)
    // If none of the special cases, return original trimmed value
    return s;
}

function parseCsvLine(line) {
    const res = [];
    let cur = '';
    let inQuotes = false;
    for (let i = 0; i < line.length; i++) {
        const ch = line[i];
        if (ch === '"') {
            if (inQuotes && line[i+1] === '"') {
                cur += '"';
                i++; // skip escaped quote
            } else {
                inQuotes = !inQuotes;
            }
            continue;
        }
        if (ch === ',' && !inQuotes) {
            res.push(cur.trim());
            cur = '';
            continue;
        }
        cur += ch;
    }
    if (cur.length || line.endsWith(',')) res.push(cur.trim());
    return res.filter(c => c !== "");
}

// Note: uploaded-file parsing removed — related helper `buildLookupFromFile` removed.

function matchesInclude(salesChannel, siteLines) {
    if (!salesChannel && salesChannel !== 0) return false;
    const norm = (s) => (s || "").toString().toLowerCase().replace(/[^a-z0-9\.]/g, '');
    const salesNorm = norm(salesChannel);
    if (!salesNorm) return false;

    for (let entry of siteLines) {
        if (!entry) continue;
        const entryNorm = norm(entry);
        if (!entryNorm) continue;
        if (salesNorm === entryNorm) return true;
        if (salesNorm.includes(entryNorm)) return true;
        if (entryNorm.includes(salesNorm)) return true;
    }
    return false;
}
