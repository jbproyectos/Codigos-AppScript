function onOpen() {

  const ui = SpreadsheetApp.getUi();

  ui.createMenu(' ➠ Transferir Datos')

        .addItem('Iniciar Envío', 'enviarInfo')

        .addItem('Envio Provisionadas', 'enviarProv')

        .addItem('Check test', 'btns')

        // .addItem('Insertar Fecha', 'fecha')

  

    .addToUi();

}

  

//////////////////////////////

  

function onEdit(e) {

  var celdaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell();

  var fila = celdaActiva.getRow() //fila activa que se bloquea

  var columna = celdaActiva.getColumn() //fila activa que se bloquea

  var hoja = celdaActiva.getSheet();

  var valor = celdaActiva.getValue();

  

  const correoEditor = Session.getActiveUser().getEmail();

  const fechaModificacion = new Date();

  

  if (hoja.getName() === "DOCUMENTOS FACTURACION" && columna === 26) {

    (valor)?hoja.getRange(fila, columna+1, 1, 2).setValues([[correoEditor, fechaModificacion]]):hoja.getRange(fila, columna+1, 1, 2).clearContent();

  }

  

  if (hoja.getName() === "MOVS" && columna === 19) {

    (valor)?hoja.getRange(fila, columna+1, 1, 2).setValues([[correoEditor, fechaModificacion]]):hoja.getRange(fila, columna+1, 1, 2).clearContent();

  }

}

  

//////////////////////////////

  

function verificarFacturas() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const hojaMovs = ss.getSheetByName("MOVS");

  const hojaDocs = ss.getSheetByName("DOCUMENTOS FACTURACION");

  

  if (!hojaMovs || !hojaDocs) {

    Logger.log("Asegúrate de que ambas hojas existan: MOVS y DOCUMENTOS FACTURACION");

    return;

  }

  

  // Obtener datos de las hojas

  const datosMovs = hojaMovs.getDataRange().getValues();

  const datosDocs = hojaDocs.getDataRange().getValues();

  

  const resultados = []; // Almacena los resultados

  

  // Recorrer las filas de MOVS (empezar en la fila 2 para omitir encabezados)

  for (let i = 1; i < datosMovs.length; i++) {

    try {

      Logger.log(`--- Procesando MOVS fila ${i + 1} ---`); // Seguimiento de progreso

  

      const mesaMovs = String(datosMovs[i][1]).trim(); // Columna B

      const promotorMovs = String(datosMovs[i][4]).trim(); // Columna E

      const idMovs = String(datosMovs[i][5]).trim(); // Columna F

      const empresaMovs = String(datosMovs[i][8]).trim(); // Columna I

      const totalMovs = parseFloat(datosMovs[i][11]); // Columna L (Total)

  

      if (!totalMovs) {

        Logger.log(`MOVS fila ${i + 1}: Sin total, saltando.`);

        continue; // Saltar si el total está vacío

      }

  

      let sumaFacturas = 0;

      const filasRelacionadas = [];

  

      // Buscar coincidencias en DOCUMENTOS FACTURACION

      for (let j = 1; j < datosDocs.length; j++) {

        const mesaDocs = String(datosDocs[j][18]).trim(); // Columna S

        const promotorDocs = String(datosDocs[j][17]).trim(); // Columna R

        const vendedorDocs = String(datosDocs[j][11]).trim(); // Columna L

        const empresaDocs = String(datosDocs[j][1]).trim(); // Columna B

        const factura = parseFloat(datosDocs[j][5]); // Columna F (Total Factura)

  

        // Comparar columnas especificadas (sin fecha)

        if (

          mesaMovs === mesaDocs &&

          promotorMovs === promotorDocs &&

          idMovs === vendedorDocs &&

          empresaMovs === empresaDocs

        ) {

          sumaFacturas += factura || 0; // Sumar factura si existe

          filasRelacionadas.push(j + 1); // Guardar fila (en formato humano)

  

          // Imprimir solo coincidencias

          Logger.log(

            `DOCUMENTOS fila ${j + 1}: Mesa=${mesaDocs}, Promotor=${promotorDocs}, Vendedor=${vendedorDocs}, Empresa=${empresaDocs}, Factura=${factura}`

          );

        }

      }

  

      // Comparar la suma de facturas con el total de la hoja MOVS

      if (sumaFacturas === totalMovs) {

        resultados.push(

          `Movimiento fila ${i + 1}: Facturas encontradas en las filas ${filasRelacionadas.join(", ")}. Total MOVS: ${totalMovs}, Suma Facturas: ${sumaFacturas}`

        );

      } else if (filasRelacionadas.length > 0) {

        resultados.push(

          `Movimiento fila ${i + 1}: Facturas encontradas en las filas ${filasRelacionadas.join(", ")}, pero los totales no coinciden. Total MOVS: ${totalMovs}, Suma Facturas: ${sumaFacturas}`

        );

      } else {

        Logger.log(`MOVS fila ${i + 1}: No se encontraron coincidencias en DOCUMENTOS.`);

      }

    } catch (error) {

      Logger.log(`Error procesando la fila ${i + 1} de MOVS: ${error.message}`);

    }

  }

  

  // Escribir los resultados en el log

  if (resultados.length > 0) {

    Logger.log("Resultados:\n" + resultados.join("\n"));

  } else {

    Logger.log("No se encontraron coincidencias.");

  }

}

  

//////////////////////////////

  

function enviarInfo() {

   // Verifica si ya se ha enviado la informacion

  if(SpreadsheetApp.getActive().getSheetByName("DOCUMENTOS FACTURACION").getRange("AJ1").getValue() != "V"){

    info();

     // En la celda 'AJ1' inserta el valor de V para marcar que ya ha sido enviada la informacion

    SpreadsheetApp.getActive().getSheetByName("DOCUMENTOS FACTURACION").getRange("AJ1").setValue("V");

    // SpreadsheetApp.getUi().alert("Informacion Enviada Exitosamente.");

  }else{SpreadsheetApp.getUi().alert("Ya se habia enviado la informacion anteriormente.")};

}

  

//////////////////////////////

  

function enviarProv() {

  if(SpreadsheetApp.getActive().getSheetByName("MOVS").getRange("X1").getValue() != "P"){

    prov();

     // En la celda 'AE1' inserta el valor de V para marcar que ya ha sido enviada la informacion

    SpreadsheetApp.getActive().getSheetByName("MOVS").getRange("X1").setValue("P");

    // SpreadsheetApp.getUi().alert("Provisionadas, Dolares y Canceladas Enviadas Exitosamente.");

  }else{SpreadsheetApp.getUi().alert("Ya se habia enviado la informacion anteriormente.")};

}

  

//////////////////////////////

  

function fecha(){

  var rangeDF = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DOCUMENTOS FACTURACION").getRange("AE2");

  var rangeM = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MOVS").getRange("W2");

  rangeDF.setValue(DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getDateCreated());

  rangeM.setValue(DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getDateCreated());

}

  

//////////////////////////////

  

function btns(){

  upDocFact2();

  upMovs2();

}

  

function enviarInfoPrueba() {

    // Archivo Origen

  var ssOrigen = SpreadsheetApp.getActiveSpreadsheet();

  var sheetODF = ssOrigen.getSheetByName("DOCUMENTOS FACTURACION");

    // Obtiene los datos de la hoja "DOCUMENTOS FACTURACION"

  const lastRowDocs = sheetODF.getRange("A1").getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();

  var rangoDatosODF = sheetODF.getRange("A1:AE"+lastRowDocs).getValues();

  var sheetOM = ssOrigen.getSheetByName("MOVS");

  const lastRowMovs = sheetOM.getRange("A1").getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();

    // Obtiene los datos de la hoja "MOVS"

  var rangoDatosOM = sheetOM.getRange("A1:V"+lastRowMovs).getValues();

    // Verifica si hay registros para copiar

  if (rangoDatosODF.length > 0 && rangoDatosOM.length > 0){

    var tempVDF = [[]], tempVM = [[]], tempFDF = [[]], tempFM = [[]];

      // Guarda encabezados

    tempVDF[0] = rangoDatosODF[0]; tempVM[0] = rangoDatosOM[0];

    tempFDF[0] = rangoDatosODF[0]; tempFM[0] = rangoDatosOM[0];

    for(i=1;i<rangoDatosODF.length;i++){

           // Validacion en DOC FACT en columnas 'U' (20) y 'W' (22)

      if(rangoDatosODF[i][20]!="PROVISIONADA" && rangoDatosODF[i][20]!="DOLARES" && rangoDatosODF[i][22]!="CANCELADA"){

        if (rangoDatosODF[i][0] || rangoDatosODF[i+1][0]){

           // Divide los registros de la hoja "DOCUMENTOS FACTURACION" (Columna 'Z' numero 25)

         (rangoDatosODF[i][25]==true)?tempVDF.push(rangoDatosODF[i]):tempFDF.push(rangoDatosODF[i]);

        }else{break;}

      }

    }

    for(i=1;i<rangoDatosOM.length;i++){

           // Validacion en MOVS en columnas 'N' (13) y 'P' (15)

      if(rangoDatosOM[i][13]!="PROVISIONADA" && rangoDatosOM[i][13]!="DOLARES" && rangoDatosOM[i][15]!="CANCELADA"){

        if (rangoDatosOM[i][0] || rangoDatosOM[i+1][0]) {

           // Divide los registros de la hoja "MOVS" (Columna 'S' numero 18)

          (rangoDatosOM[i][18]==true)?tempVM.push(rangoDatosOM[i]):tempFM.push(rangoDatosOM[i]);

        }else{break;}

      }

    }

      // Envia la Informacion al Archivo de Verdaderos

    var ssV = SpreadsheetApp.openById(CONCENTRADO_ID); // Archivo Concentrado MAYO

    var ssVDF = ssV.getSheetByName("DOCUMENTOS FACTURACION");

    var ssVM = ssV.getSheetByName("MOVS");

      // Crea el Archivo de Falsos

    var ssF = ssOrigen.copy(ssOrigen.getName() + " - FALSOS");

    var ssFDF = ssF.getSheetByName("DOCUMENTOS FACTURACION");

    var ssFM = ssF.getSheetByName("MOVS");

      // Inserta los datos

    tempVDF.shift(); tempVM.shift();

     // Verifica si hay informacion para mandar

    var ssVDFLastRow = (ssVDF.getRange(`A1:A`).getValues().filter(fila => fila[0] !== ``).flat()

    .length)+1;

    var ssVMLastRow = (ssVM.getRange(`A1:A`).getValues().filter(fila => fila[0] !== ``).flat()

    .length)+1;

    Logger.log(`${ssVDFLastRow} y ${ssVMLastRow}`);

  }

}

## Relacionados

- [[Concentrado Conciliacion Oper-Fact/Cuadre Diario/ScriptDias.gs]] — mismo archivo/proyecto: aquí vive el menú (`enviarInfo`) que dispara la lógica pesada de ese script
