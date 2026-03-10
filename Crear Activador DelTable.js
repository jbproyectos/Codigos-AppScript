//////////////////////////////

function creaActivadorDelTable() {
  const allTriggers = ScriptApp.getProjectTriggers();
  allTriggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'DelTable') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('DelTable') // Activador para cada fin de dia
    .timeBased()
    .everyDays(1)
    .atHour(23)
    .create();
}

//////////////////////////////
