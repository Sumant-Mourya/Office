/*************************************************
 * MENU – SINGLE TOGGLE BUTTON
 *************************************************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Auto Reply Control")
    .addItem("🔁 Start / Stop Auto Reply", "toggleAutoReply")
    .addItem("🔁 Start / Stop Backlog Auto Reply", "toggleBacklogAutoReply")
    .addToUi();
}

/*************************************************
 * TOGGLE AUTO REPLY
 *************************************************/
function toggleAutoReply() {

  const handlerName = "autoReplyWorker";
  const triggers = ScriptApp.getProjectTriggers();

  // 🔍 check if trigger already exists
  const existing = triggers.find(t =>
    t.getHandlerFunction() === handlerName
  );

  // 🟥 IF RUNNING → STOP
  if (existing) {
    ScriptApp.deleteTrigger(existing);
    SpreadsheetApp.getUi().alert("⏹ Auto Reply STOPPED");
    return;
  }

  // 🟩 IF NOT RUNNING → START
  ScriptApp.newTrigger(handlerName)
    .timeBased()
    .everyMinutes(1) // testing mode
    .create();

  SpreadsheetApp.getUi().alert("▶ Auto Reply STARTED");
}

/**
 * TOGGLE BACKLOG AUTO REPLY (single-button)
 */
function toggleBacklogAutoReply() {

  const handlerName = "backlogAutoReplyWorker";
  const triggers = ScriptApp.getProjectTriggers();

  // 🔍 check if trigger already exists
  const existing = triggers.find(t =>
    t.getHandlerFunction() === handlerName
  );

  // 🟥 IF RUNNING → STOP
  if (existing) {
    ScriptApp.deleteTrigger(existing);
    // Set stop flag so the currently running execution exits at next check
    PropertiesService.getScriptProperties().setProperty("BACKLOG_STOP_FLAG", "true");
    SpreadsheetApp.getUi().alert("⏹ Backlog Auto Reply STOPPED");
    return;
  }

  // 🟩 IF NOT RUNNING → START
  // Clear stop flag so the worker is allowed to run
  PropertiesService.getScriptProperties().setProperty("BACKLOG_STOP_FLAG", "false");
  ScriptApp.newTrigger(handlerName)
    .timeBased()
    .everyMinutes(1) // same testing interval as main toggle
    .create();

  SpreadsheetApp.getUi().alert("▶ Backlog Auto Reply STARTED");
}