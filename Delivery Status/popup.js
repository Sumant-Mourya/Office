const sheetIdsContainer = document.getElementById("sheetIdsContainer");
const addSheetBtn = document.getElementById("addSheetBtn");
const startBtn = document.getElementById("startBtn");
const stopBtn = document.getElementById("stopBtn");
const logsDiv = document.getElementById("logs");
const statusDiv = document.getElementById("status");
const settingsBtn = document.getElementById("settingsBtn");
const backBtn = document.getElementById("backBtn");
const mainScreen = document.getElementById("mainScreen");
const settingsScreen = document.getElementById("settingsScreen");
const orderIdCol = document.getElementById("orderIdCol");
const orderStatusCol = document.getElementById("orderStatusCol");
const siteList = document.getElementById("siteList");
const saveSitesBtn = document.getElementById("saveSitesBtn");
const saveStatus = document.getElementById("saveStatus");
const clearLogsBtn = document.getElementById("clearLogsBtn");

function requestInteractiveLogin() {
    return new Promise((resolve, reject) => {
        chrome.runtime.sendMessage({ action: "authenticate" }, (response) => {
            if (chrome.runtime.lastError) {
                reject(new Error(chrome.runtime.lastError.message));
                return;
            }
            if (response && response.ok) {
                resolve();
                return;
            }
            const err = new Error(response?.error || "Google login failed");
            err.code = response?.code || "AUTH_FAILED";
            reject(err);
        });
    });
}

function createSheetRow(value = "") {
    const row = document.createElement("div");
    row.className = "sheet-row";
    
    const input = document.createElement("input");
    input.type = "text";
    input.className = "sheetIdInput";
    input.placeholder = "Enter Sheet ID";
    input.value = value;
    
    const removeBtn = document.createElement("button");
    removeBtn.className = "btn-icon remove-sheet-btn";
    removeBtn.title = "Remove Sheet";
    removeBtn.innerHTML = `
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <line x1="5" y1="12" x2="19" y2="12"></line>
        </svg>
    `;
    
    removeBtn.addEventListener("click", () => {
        row.remove();
    });
    
    row.appendChild(input);
    row.appendChild(removeBtn);
    sheetIdsContainer.appendChild(row);
}

document.addEventListener("DOMContentLoaded", () => {
    chrome.storage.local.get(["sheetId", "sheetIds", "logs", "running", "orderIdCol", "orderStatusCol", "siteList", "lastCompleteDate", "manualStopDate"], (data) => {

        let loadedSheetIds = [];
        if (data.sheetIds && Array.isArray(data.sheetIds) && data.sheetIds.length > 0) {
            loadedSheetIds = data.sheetIds;
        } else if (data.sheetId) {
            // legacy migration
            loadedSheetIds = [data.sheetId];
        }

        if (loadedSheetIds.length > 0) {
            loadedSheetIds.forEach(id => createSheetRow(id));
        } else {
            createSheetRow(); // Show at least one empty input
        }

        if (addSheetBtn) {
            addSheetBtn.addEventListener("click", () => {
                createSheetRow();
            });
        }

        if (data.running) {
            statusDiv.innerText = "Running";
            statusDiv.className = "status-badge status-running";
        } else {
            statusDiv.innerText = "Stopped";
            statusDiv.className = "status-badge status-stopped";
        }

        if (data.logs) {
            logsDiv.innerHTML = "";
            data.logs.forEach(l => {
                const div = document.createElement("div");
                div.innerText = l;
                logsDiv.appendChild(div);
            });
            logsDiv.scrollTop = logsDiv.scrollHeight;
        }

        // mapping defaults
        if (orderIdCol) orderIdCol.value = data.orderIdCol || "A";
        if (orderStatusCol) orderStatusCol.value = data.orderStatusCol || "D";

        // site list defaults (multiline). Added Shopify as requested.
        const defaultSites = [
            "Shopify",
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
        ].join("\n");
        if (data.hasOwnProperty("siteList")) {
            siteList.value = data.siteList;
        } else {
            siteList.value = defaultSites;
        }

        // upload UI removed; no saved-file status shown
    });

    // switch to settings screen
    settingsBtn.addEventListener("click", () => {
        mainScreen.style.display = "none";
        settingsScreen.style.display = "block";
    });

    // switch to main screen
    if (backBtn) {
        backBtn.addEventListener("click", () => {
            settingsScreen.style.display = "none";
            mainScreen.style.display = "block";
        });
    }

    // clear logs
    if (clearLogsBtn) {
        clearLogsBtn.addEventListener("click", () => {
            chrome.storage.local.set({ logs: [] }, () => {
                logsDiv.innerHTML = "";
            });
        });
    }

    // upload functionality removed

    // persist mapping inputs
    if (orderIdCol) {
        orderIdCol.addEventListener("input", () => {
            chrome.storage.local.set({ orderIdCol: orderIdCol.value });
        });
    }
    if (orderStatusCol) {
        orderStatusCol.addEventListener("input", () => {
            chrome.storage.local.set({ orderStatusCol: orderStatusCol.value });
        });
    }

    // persist site list
    siteList.addEventListener("input", () => {
        chrome.storage.local.set({ siteList: siteList.value });
    });

    // Save button for site list (explicit save)
    if (saveSitesBtn) {
        saveSitesBtn.addEventListener("click", () => {
            chrome.storage.local.set({ siteList: siteList.value }, () => {
                if (saveStatus) {
                    saveStatus.innerText = "Saved";
                    setTimeout(() => { saveStatus.innerText = ""; }, 2000);
                }
            });
        });
    }
});

startBtn.addEventListener("click", () => {
    const inputs = document.querySelectorAll(".sheetIdInput");
    const sheetIds = [];
    inputs.forEach(input => {
        const val = input.value.trim();
        if (val) sheetIds.push(val);
    });
    
    if (sheetIds.length === 0) {
        return alert("Enter at least one Sheet ID");
    }

    // Always restart processing from the first sheet when 'Start' is clicked manually
    chrome.storage.local.set({ 
        sheetIds: sheetIds, 
        currentSheetIndex: 0,
        lastProcessedRow: 0, 
        lastRunTime: 0, 
        lastCompleteDate: "",
        manualStopDate: "", // Clear manual stop flag so it can run
        sheetId: sheetIds[0] // fallback for any old scripts
    }, async () => {
        try {
            await requestInteractiveLogin();
            chrome.runtime.sendMessage({ action: "start", source: "manual-popup" });
            statusDiv.innerText = "Running";
            statusDiv.className = "status-badge status-running";
        } catch (err) {
            statusDiv.innerText = "Stopped";
            statusDiv.className = "status-badge status-stopped";

            if (err && err.code === "BAD_CLIENT_ID") {
                alert(
                    "OAuth setup mismatch on this PC.\n\n" +
                    "Extension ID: " + chrome.runtime.id + "\n\n" +
                    "Fix: In Google Cloud Console, create/update OAuth client for Chrome Extension with this Extension ID, then replace manifest oauth2.client_id if needed."
                );
                return;
            }

            alert("Google login is required before starting. " + (err?.message || ""));
        }
    });
});

stopBtn.addEventListener("click", () => {
    // Mark stopped and reset run markers. Keep auto-start eligible on next Chrome launch.
    chrome.storage.local.set({ running: false, currentSheetIndex: 0, lastProcessedRow: 0, lastRunTime: 0, lastCompleteDate: "", manualStopDate: "", logs: [] }, () => {
        chrome.runtime.sendMessage({ action: "stop" });
        statusDiv.innerText = "Stopped";
        statusDiv.className = "status-badge status-stopped";
        // clear logs in UI
        logsDiv.innerHTML = "";
    });
});


chrome.runtime.onMessage.addListener((msg) => {
    if (msg.type === "log") {
        const div = document.createElement("div");
        div.innerText = msg.text;
        logsDiv.appendChild(div);
        logsDiv.scrollTop = logsDiv.scrollHeight;
    }

    if (msg.type === "auth-required") {
        statusDiv.innerText = "Stopped";
        statusDiv.className = "status-badge status-stopped";
    }

    if (msg.type === "auth-config-error") {
        statusDiv.innerText = "Stopped";
        statusDiv.className = "status-badge status-stopped";
        alert(
            "OAuth setup mismatch (bad client id).\n\n" +
            "Extension ID: " + chrome.runtime.id + "\n\n" +
            (msg.text || "Configure Google OAuth client for this extension ID.")
        );
    }
});
