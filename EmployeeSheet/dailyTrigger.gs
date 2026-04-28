/**
 * Trigger and scheduler functions.
 * This keeps trigger setup separate from main sync logic.
 */

/**
 * Single daily runner that executes both jobs in order.
 */
function runDaily8amJobs() {
  syncLateSolvedOrders();
  syncClearBoxOrders();
}

/**
 * Run this manually once to reset and register a single 8 AM trigger.
 * If already registered, it will be unregistered first, then re-registered.
 */
function registerOrResetMorningTrigger() {
  const functionName = 'runDaily8amJobs';
  const existing = ScriptApp.getProjectTriggers();

  let removed = 0;
  existing.forEach((t) => {
    if (t.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });

  ScriptApp.newTrigger(functionName)
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  if (removed > 0) {
    Logger.log('Removed %s existing trigger(s), then registered %s at 8 AM daily.', removed, functionName);
  } else {
    Logger.log('No existing trigger found. Registered %s at 8 AM daily.', functionName);
  }
}
