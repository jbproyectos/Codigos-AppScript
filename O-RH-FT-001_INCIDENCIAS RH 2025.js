function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('💼 | Menu')
    .addItem('1. Historial PERMISO/VACACIONES, RETARDO/FALTA  | 📑', 'copiadosPgadosDosHojas')
    .addItem('1. Borrar WORKY, BIOTIME, PERMISO/VACACIONES W  | 🗳', 'borradoHojas')
    .addToUi();
}

/////////////////////////
function copiadosPgadosDosHojas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var objetoHojas = [
    {origen: "PERMISO/VACACIONES", destino: "HISTORICO PERMISO/VACACIONES", rango: "A2:K", colInicial: 1},
    {origen: "RETARDO/FALTA", destino: "HISTORICO RETARDO/FALTA", rango: "A2:H", colInicial: 1}
  ];

  objetoHojas.forEach(function(hoja){
    var hojaOrigen = ss.getSheetByName(hoja.origen);
    var hojaDestino = ss.getSheetByName(hoja.destino);

    var rango = hojaOrigen.getRange(hoja.rango).getValues();
    var ultimaColumnaDestino = obtenerUltimaFila(hojaDestino, 1, 2);

    
    hojaDestino.getRange(ultimaColumnaDestino, hoja.colInicial, rango.length, rango[0].length).setValues(rango);

  });
}

function obtenerUltimaFila(hoja, columna, filaInicio) {
  const lastRow = hoja.getLastRow();

  if (lastRow < filaInicio) {
    return filaInicio; // hoja vacía desde filaInicio
  }

  const datos = hoja
    .getRange(filaInicio, columna, lastRow - filaInicio + 1)
    .getValues();

  for (let i = datos.length - 1; i >= 0; i--) {
    if (datos[i][0] !== "") {
      return filaInicio + i + 1; // ⬅️ siguiente fila libre
    }
  }

  return filaInicio; // si no hay datos, empieza en filaInicio
}

//////////////////
function borradoHojas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var objetoBorrado = [
      { origen: "PERMISO/VACACIONES W", rango: "A6:P"},
      { origen: "BIOTIME", rango: "A2:M"},
      { origen: "WORKY", rango: "A2:AB"}  
  ];

   // Borrar rangos
  objetoBorrado.forEach(hojaO => {
    const hoja = ss.getSheetByName(hojaO.origen);
    if (hoja) {
      hoja.getRange(hojaO.rango).clearContent();
    }
  });

}
