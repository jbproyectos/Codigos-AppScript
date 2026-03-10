//////////////////////////////

function registrarTriggerCreado(){

  const controlSheetId = `https://docs.google.com/spreadsheets/d/11IcA7OSlVT5qnpG8F6PHjhEM5hW5bF3etr4p3Y3v3Ck/edit?gid=0#gid=0`;

  const archivoActual = SpreadsheetApp.getActiveSpreadsheet().getUrl();

  const sheetControl = SpreadsheetApp
    .openById(controlSheetId)
    .getSheetByName(`CONTROL`);

  const data = sheetControl.getRange("J:J").getValues();

  for (let i = 0; i < data.length; i++) {

    if (data[i][0] === archivoActual) {

      sheetControl.getRange(i + 1, 12).setValue(true);
      break;

    }

  }

}

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

function creaActivador() {
  const allTriggers = ScriptApp.getProjectTriggers();
  allTriggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'CopiaDiasLaborales') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('CopiaDiasLaborales') // Activador para cada fin de dia
    .timeBased()
    .everyDays(1)
    .atHour(21)
    .create();
  
  registrarTriggerCreado();
}
