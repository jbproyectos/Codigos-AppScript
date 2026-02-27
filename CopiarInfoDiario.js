  //  Constantes globales
const SSID = SpreadsheetApp.getActiveSpreadsheet().getId();
const SH_NAME = "HOY";
const CONCENTRADO = `CONCENTRADO`;

function copiarHoja() {
  var ss = SpreadsheetApp.openById(SSID);
  var sheet = ss.getSheetByName(SH_NAME);
  var name = 
    Utilities.formatDate(
      new Date(sheet.getRange(`A5`).getValue()),
      Session.getScriptTimeZone(),
      "dd/MM/yyyy"
    );
  var nuevaHoja = ss.getSheetByName(name);
  (nuevaHoja)?ss.deleteSheet(nuevaHoja):0;
  sheet.copyTo(ss).setName(name);
  nuevaHoja = ss.getSheetByName(name);
  var nuevoRango = nuevaHoja.getRange(`D6:I463`);
  var nuevosValores = nuevoRango.getDisplayValues();
    // Metodo para eliminar el IMPORTRANGE y dejar datos planos
  nuevoRango.setValues(nuevosValores);
  nuevaHoja.deleteRow(5);
  sheet.getRange(`B1`).setValue(sheet.getRange(`B3`).getDisplayValue());
}

//////////////////////////////

function mandarAlConcentrado(){
  var ss = SpreadsheetApp.openById(SSID);
  var sheet = ss.getSheetByName(SH_NAME);
  var concentrado = ss.getSheetByName(CONCENTRADO);
  var valoresSheet = sheet.getRange(`F6`).getDataRegion().getValues();
  var valoresFiltrados = valoresSheet.filter(fila => 
    fila[0] !== `` &&
    fila[0] !== null &&
    fila[0] !== `SALDO AL DÍA` &&
    fila[0] !== `MONTO` &&
    fila[0] !== `SUMA`
    );
  var valoresConc = concentrado.getRange(`F4`).getDataRegion().getValues();
  var valoresAnteriores = valoresConc.filter(fila => 
    fila[0] !== `` &&
    fila[0] !== null &&
    fila[0] !== `SALDO AL DÍA` &&
    fila[0] !== `MONTO` &&
    fila[0] !== `SUMA`
    );
  var concRange = concentrado.getRange(valoresAnteriores.length+5,4,valoresFiltrados.length,6);
  concentrado.insertRowsAfter(valoresAnteriores.length+5,valoresFiltrados.length);
  concRange.setValues(valoresFiltrados);
}
