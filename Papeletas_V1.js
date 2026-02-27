// SCRIPT PARA TOMAR COO REFERENCIA PARA PAPELETAS OPERACIONES

const AR_PERSONAL = "19fionVnXuVOe2Ex5WthuF5te0YrZbPz-HN8YvV9XwuA"; // OFICIAL
const AR_TEMPORAL = "18SOk6PCHpIxbL7oEfXK8MHnr8yzWGzJWNf_HYCmrmGk"; // OFICIAL
const SH_PAPELETAS_10R = "PAPELETAS_10R"; // OFICIAL
const SH_PAPELETAS_5R = "PAPELETAS_5R"; // OFICIAL

// const AR_PERSONAL = "1b0AV9lN3m_6Dj5DoVsuOtMaBXwfmxbc_W2iysrhQ7Cg"; // PRUEBA
// const AR_TEMPORAL = "1zmzm63nKPLD3qTmR0jJClKxtShXEILHQtxy2Mlsx284"; // PRUEBA
// const SH_PAPELETAS_10R = "PAPELETAS_10R_Prueba"; // PRUEBA
// const SH_PAPELETAS_5R = "PAPELETAS_5R_Prueba"; // PRUEBA

const SHEET_PER = "S.Gastos Personales";
const SHEET_DIR = "SOLICITUD GASTOS TEMPORAL - CONCATENADO";

//////////////////////////////

function papeletasInfoPer(SSID) { // De archivo 003 a 10R
  const hoja = SpreadsheetApp.openById(AR_PERSONAL).getSheetByName(SHEET_PER);
  const ultimaFila = hoja.getLastRow();
  const cantidadFilas = ultimaFila - 5;
  const datos = hoja.getRange(6, 1, cantidadFilas, hoja.getLastColumn()).getValues();
  const datosFiltrados = datos.filter((fila, index) => {
    return (fila[17] == "EFECTIVO" && fila[28] == "NUEVO"); // Valida si es "NUEVO" y tipo "EFECTIVO"
  });
  var identificador = 0;    // A
  var quienSolEmail = 2;    // C
  var dptoSol = 3;          // D
  var areaCliente = 6;      // G
  var desc = 15;            // P
  var formaPago = 17;       // R
  var comentarios = 19;     // T
  var titularContact = 22;  // W
  var monto = 23;           // X
  const agrupados = {};
  const email = {
    "CARLOS_CAVAZOS": "logistica3@kabzo.org",
    "EMILIO_TORRES": "logistica4@kabzo.org",
    "ALONDRA_MARIBEL": "administracion.empresarial1@rhfinder.com; juridico1@kabzo.org",
    "GABRIELA_LEAL": "administracionempresarial4@rhfinder.com; imss@kabzo.org; direccion@rhfinder.com",
    "CINTHIA_OVIEDO": "juridico15@kabzo.org",
    "MIKEL_GARZA": "juridico9@kabzo.org",
    "SARAI_BELLO": "projectmanager@kabzo.org",
    "FATIMA_MARTINEZ": "administracionempresarial3@rhfinder.com",
    "FRIDA_PIÑA": "administracion.empresarial6@rhfinder.com",
    "NAYELI_LUNA": "administracion.empresarial5@rhfinder.com",
    "JORGE_TAMEZ": "sistemas3@kabzo.org",
    "VALERIA_VARGAS": "administracion.empresarial1@rhfinder.com",
    "EVAMARIA_SUAREZ": "asistente2.admin@kabzo.org",
    "CYNTHIA_GONZALEZ": "disenografico@produccionesdobleb.com",
    "BRENDA_SALDIVAR": "administracion.empresarial6@rhfinder.com; juridico7@kabzo.org",
    "ALEJANDRA_JASSO": "juridico5@kabzo.org",
    "SAMANTHA_REYNA": "facturacion@kabzo.org",
    "PERLA_ESPINOZA": "administracionempresarial4@rhfinder.com; dir.contable@kabzo.org",
    "ISAIRA_MORANTES": "contabilidad1@kabzo.org",
    "ISAMAR_BALDERAS": "administracionempresarial4@rhfinder.com; imss@kabzo.org",
    "SERGIO_SALAS": "tesorerias3@kabzo.org",
    "GABRIELA_GONZALEZ": "tesoreria1@kabzo.org",
    "FERNANDO_SALDAÑA": "receptor@kabzo.org",
    "AZAEL_RANGEL": "verificador@kabzo.org",
    "DAMARIS_TORRES": "verificador2@kabzo.org",
    "ANGELES_TREVIÑO": "administracion@kabzo.org",
    "KEVIN_MORENO": "cobranza@kabzo.org",
    "NATALIE_REYNA": "administracion.empresarial@rhfinder.com",
    "LIZBETH_MARTINEZ": "cobranza3@kabzo.org",
    "JUAN_BAUTISTA": "desarrollo@kabzo.org",
    "JORGE_MORAN": "auxproyectos@kabzo.org",
    "EMELY_CARMONA": "gastos@kabzo.org",
    "CAROLINA_VARGAS": "",
    "DIONICIO": "",
    "DIEGO": "logistica8@kabzo.org",
    "DARLIN_APONTE": "asistente.admin@kabzo.org",
  };

  datosFiltrados.forEach(fila => {  // Acomoda la informacion en la columna agrupando por ID
    const id = fila[identificador];
    const nuevaDesc = fila[desc]?.trim() || "";
    const nuevoTexto = fila[comentarios]?.trim() || "";
    if (!agrupados[id]) {
      // asesor = , envio = , pr = 
      agrupados[id] = {
        identificador: id,
        asesor: "CONSULTORIA",
        fechaCaptura: "=TODAY()",
        envio: "DIRECCION",
        pr: "CONSULTORIA",
        quienSolEmail: email[fila[quienSolEmail]],
        dptoSol: fila[dptoSol],
        areaCliente: fila[areaCliente],
        vacio: "",
        descs: new Set(nuevaDesc ? [nuevaDesc] : []),
        textos: new Set(nuevoTexto ? [nuevoTexto] : []),
        formaPago: fila[formaPago],
        titularContact: fila[titularContact],
        monto: (parseFloat(fila[monto]) * -1) || 0
      };
    } else {
      agrupados[id].monto += (parseFloat(fila[monto]) * -1) || 0;
      if (nuevaDesc) agrupados[id].descs.add(nuevaDesc);
      if (nuevoTexto) agrupados[id].textos.add(nuevoTexto);
    }
  });

  const salida = Object.values(agrupados).map(obj => [
    obj.asesor,
    obj.fechaCaptura,
    obj.envio,
    obj.pr,
    obj.quienSolEmail,
    obj.areaCliente,
    obj.titularContact,
    obj.vacio,
    obj.vacio,
    [...obj.descs, ...obj.textos].join(" // "),
    obj.identificador,
    Math.round(obj.monto)
  ]);

  if (salida.length === 0) {
    SpreadsheetApp.getUi().alert("No hay datos para escribir.");  // En caso de no encontrar datos se muestra mensaje
    return;
  }
  const hojaDestino = SpreadsheetApp.openById(SSID).getSheetByName(SH_PAPELETAS_10R);
  if (!hojaDestino) {
    SpreadsheetApp.getUi().alert("La hoja de Papeletas no existe.");  // En caso de no encontrar hoja con ese nombre
  }
  const columnasPorBloque = 14; // 14 columnas: B (2) hasta AO (41) en saltos de 3
  const saltoFilas = 25;        // Cada bloque nuevo baja 25 filas
  const columnasEspaciadas = 3; // Espacio entre columnas

  salida.forEach((colData, colIndex) => {
    const bloque = Math.floor(colIndex / columnasPorBloque);
    const posicionEnBloque = colIndex % columnasPorBloque;

    const filaDestino = 4 + bloque * saltoFilas;
    const columnaDestino = 2 + posicionEnBloque * columnasEspaciadas;

    hojaDestino.getRange(filaDestino, columnaDestino, colData.length, 1).setValues( // Inserta los valores de la papeleta en los intervalos correctos
      colData.map(valor => [valor])
    );
  });
}

//////////////////////////////

function papeletasInfoDir(SSID) { // De archivo Temporal a 5R
  const hoja = SpreadsheetApp.openById(AR_TEMPORAL).getSheetByName(SHEET_DIR);
  const ultimaFila = hoja.getLastRow();
  const cantidadFilas = ultimaFila - 5;
  const datos = hoja.getRange(6, 1, cantidadFilas, hoja.getLastColumn()).getValues();
  const datosFiltrados = datos.filter((fila, index) => {
    return (fila[17] == "EFECTIVO" && fila[28] == "NUEVO"); // Valida si es "NUEVO" y tipo "EFECTIVO"
  });
  var identificador = 0;    // A
  var fechaCaptura = 1;     // B
  var quienSolEmail = 2;    // C
  var dptoSol = 3;          // D
  var areaCliente = 6;      // G
  var desc = 15;            // P
  var formaPago = 17;       // R
  var comentarios = 19;     // T
  var titularContact = 22;  // W
  var monto = 23;           // X
  const agrupados = {};
  const email = {
    "CARLOS_CAVAZOS": "logistica3@kabzo.org",
    "EMILIO_TORRES": "logistica4@kabzo.org",
    "ALONDRA_MARIBEL": "administracion.empresarial1@rhfinder.com; juridico1@kabzo.org",
    "GABRIELA_LEAL": "administracionempresarial4@rhfinder.com; imss@kabzo.org; direccion@rhfinder.com",
    "CINTHIA_OVIEDO": "juridico15@kabzo.org",
    "MIKEL_GARZA": "juridico9@kabzo.org",
    "SARAI_BELLO": "projectmanager@kabzo.org",
    "FATIMA_MARTINEZ": "administracionempresarial3@rhfinder.com",
    "FRIDA_PIÑA": "administracion.empresarial6@rhfinder.com",
    "NAYELI_LUNA": "administracion.empresarial5@rhfinder.com",
    "JORGE_TAMEZ": "sistemas3@kabzo.org",
    "VALERIA_VARGAS": "administracion.empresarial1@rhfinder.com",
    "EVAMARIA_SUAREZ": "asistente2.admin@kabzo.org",
    "CYNTHIA_GONZALEZ": "disenografico@produccionesdobleb.com",
    "BRENDA_SALDIVAR": "administracion.empresarial6@rhfinder.com; juridico7@kabzo.org",
    "ALEJANDRA_JASSO": "juridico5@kabzo.org",
    "SAMANTHA_REYNA": "facturacion@kabzo.org",
    "PERLA_ESPINOZA": "administracionempresarial4@rhfinder.com; dir.contable@kabzo.org",
    "ISAIRA_MORANTES": "contabilidad1@kabzo.org",
    "ISAMAR_BALDERAS": "administracionempresarial4@rhfinder.com; imss@kabzo.org",
    "SERGIO_SALAS": "tesorerias3@kabzo.org",
    "GABRIELA_GONZALEZ": "tesoreria1@kabzo.org",
    "FERNANDO_SALDAÑA": "receptor@kabzo.org",
    "AZAEL_RANGEL": "verificador@kabzo.org",
    "DAMARIS_TORRES": "verificador2@kabzo.org",
    "ANGELES_TREVIÑO": "administracion@kabzo.org",
    "KEVIN_MORENO": "cobranza@kabzo.org",
    "NATALIE_REYNA": "administracion.empresarial@rhfinder.com",
    "LIZBETH_MARTINEZ": "cobranza3@kabzo.org",
    "JUAN_BAUTISTA": "desarrollo@kabzo.org",
    "JORGE_MORAN": "auxproyectos@kabzo.org",
    "EMELY_CARMONA": "gastos@kabzo.org"
  };

  datosFiltrados.forEach(fila => {  // Acomoda la informacion en la columna agrupando por ID
    const id = fila[identificador];
    const nuevaDesc = fila[desc]?.trim() || "";
    const nuevoTexto = fila[comentarios]?.trim() || "";
    if (!agrupados[id]) {
      // asesor = , envio = , pr = 
      agrupados[id] = {
        identificador: id,
        asesor: "GASTOS",
        fechaCaptura: "=TODAY()",
        envio: "DIRECCION",
        pr: "GASTOS",
        quienSolEmail: email[fila[quienSolEmail]],
        dptoSol: fila[dptoSol],
        areaCliente: fila[areaCliente],
        vacio: "",
        descs: new Set(nuevaDesc ? [nuevaDesc] : []),
        textos: new Set(nuevoTexto ? [nuevoTexto] : []),
        formaPago: fila[formaPago],
        titularContact: fila[titularContact],
        monto: (parseFloat(fila[monto]) * -1) || 0
      };
    } else {
      agrupados[id].monto += (parseFloat(fila[monto]) * -1) || 0;
      if (nuevaDesc) agrupados[id].descs.add(nuevaDesc);
      if (nuevoTexto) agrupados[id].textos.add(nuevoTexto);
    }
  });

  const salida = Object.values(agrupados).map(obj => [
    obj.asesor,
    obj.fechaCaptura,
    obj.envio,
    obj.pr,
    obj.quienSolEmail,
    obj.areaCliente,
    obj.titularContact,
    obj.vacio,
    obj.vacio,
    [...obj.descs, ...obj.textos].join(" // "),
    obj.identificador,
    Math.round(obj.monto)
  ]);
  
  if (salida.length === 0) {
    SpreadsheetApp.getUi().alert("No hay datos para escribir.");  // En caso de no encontrar datos se muestra mensaje
    return;
  }
  const hojaDestino = SpreadsheetApp.openById(SSID).getSheetByName(SH_PAPELETAS_5R);
  if (!hojaDestino) {
    SpreadsheetApp.getUi().alert("La hoja de Papeletas no existe.");  // En caso de no encontrar hoja con ese nombre
  }
  const columnasPorBloque = 14; // 14 columnas: B (2) hasta AO (41) en saltos de 3
  const saltoFilas = 25;        // Cada bloque nuevo baja 25 filas
  const columnasEspaciadas = 3; // Espacio entre columnas

  salida.forEach((colData, colIndex) => {
    const bloque = Math.floor(colIndex / columnasPorBloque);
    const posicionEnBloque = colIndex % columnasPorBloque;

    const filaDestino = 4 + bloque * saltoFilas;
    const columnaDestino = 2 + posicionEnBloque * columnasEspaciadas;

    hojaDestino.getRange(filaDestino, columnaDestino, colData.length, 1).setValues( // Inserta los valores de la papeleta en los intervalos correctos
      colData.map(valor => [valor])
    );
  });
}

//////////////////////////////

function customTranspose(matrix) {
  return matrix[0].map((_, colIndex) => matrix.map(row => row[colIndex]));
}

//////////////////////////////

function eraseColumns5R(SSID) {
  const hojaDestino = SpreadsheetApp.openById(SSID).getSheetByName(SH_PAPELETAS_5R);
  for (i = 0; i < 14; i++) {
    hojaDestino.getRange(4, (2 + i * 3), 423, 1).clearContent();
  }
}

//////////////////////////////

function eraseColumns10R(SSID) {
  const hojaDestino = SpreadsheetApp.openById(SSID).getSheetByName(SH_PAPELETAS_10R);
  for (i = 0; i < 14; i++) {
    hojaDestino.getRange(4, (2 + i * 3), 423, 1).clearContent();
  }
}
