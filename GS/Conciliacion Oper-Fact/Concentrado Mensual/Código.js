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

  

function onOpen() {

  const ui = SpreadsheetApp.getUi();

  ui.createMenu(' ➠ Transferir Datos')

    .addItem('Iniciar Envío', 'enviarInfo')

        .addItem('Check test', 'btns')

        .addItem('📊 Ver diagrama de flujo', 'mostrarDiagramaFlujo')

  

    .addToUi();

}

  
  
  

function enviarInfo(){

  EnviarInfo.enviarInfo();

}

  
  

function onEdit(e) {

  const hoja = e.source.getActiveSheet();

  const fila = e.range.getRow();

  const columna = e.range.getColumn();

  const valor = e.value;

  

  // Obtiene el correo del editor y la fecha y hora actual

  const correoEditor = Session.getActiveUser().getEmail();

  const fechaModificacion = new Date();

  

  // Configuración para la hoja "DOCUMENTOS FACTURACION"

  if (hoja.getName() === "DOCUMENTOS FACTURACION" && columna === 26 && valor === "TRUE") {

    // Columna X es 24, Columna Y es 25, Columna Z es 26

    hoja.getRange(fila, 27).setValue(correoEditor); // Columna Y

    hoja.getRange(fila, 28).setValue(fechaModificacion); // Columna Z

  }

  

  // Configuración para la hoja "MOVS"

  if (hoja.getName() === "MOVS" && columna === 19 && valor === "TRUE") {

    // Columna Q es 17, Columna R es 18, Columna S es 19

    hoja.getRange(fila, 20).setValue(correoEditor); // Columna R

    hoja.getRange(fila,21).setValue(fechaModificacion); // Columna S

  }

}

  
  
  

function btns(){

  EnviarInfo.upDocFact()

  EnviarInfo.upMovs()

}

  
  

function upDocFact() {

  // Abre la hoja activa

  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DOCUMENTOS FACTURACION');

  // Define el rango de la columna V (valores a verificar) y la columna W (checkboxes)

  const rangoValores = hoja.getRange('x2:y1520'); // Cambia este rango según tus datos

  const valores = rangoValores.getValues(); // Obtén los valores de la columna V

  // Define el rango de checkboxes en la columna W

  const rangoCheckboxes = hoja.getRange('z2:z1520');

  const checkboxes = [];

  // Define el rango de las columnas X (Capturado por código) y Y (Fecha de ejecución)

  const rangoCapturadoPorCodigo = hoja.getRange('aa2:aa1520');

  const rangoFechaEjecucion = hoja.getRange('ab2:ab1520');

  const capturadoPorCodigo = [];

  const fechaEjecucion = [];

  // Obtén la fecha de ejecución (solo una vez, la misma para todas las filas)

  const fechaActual = new Date();

  // Recorre los valores de la columna V

  for (let i = 0; i < valores.length; i++) {

    // Si el valor en la columna V es 1, marca el checkbox (TRUE), si no, desmárcalo (FALSE)

    checkboxes.push([(valores[i][0] === 1)||(valores[i][1] === 1)]);

    // Si el valor es 1, coloca "Capturado por código" en la columna X y la fecha de ejecución en la columna Y

    if ((valores[i][0] === 1)||(valores[i][1] === 1)) {

      capturadoPorCodigo.push(["Capturado por código"]);

      fechaEjecucion.push([fechaActual]);

    } else {

      capturadoPorCodigo.push([null]); // Deja vacío si no es 1

      fechaEjecucion.push([null]); // Deja vacío si no es 1

    }

  }

  // Actualiza los checkboxes en la columna W

  rangoCheckboxes.setValues(checkboxes);

  // Actualiza la columna X (Capturado por código)

  rangoCapturadoPorCodigo.setValues(capturadoPorCodigo);

  // Actualiza la columna Y (Fecha de ejecución)

  rangoFechaEjecucion.setValues(fechaEjecucion);

}

  
  

function upMovs() {

  // Abre la hoja activa

  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MOVS');

  // Define el rango de la columna L (valores a verificar) y la columna M (checkboxes)

  const rangoValores = hoja.getRange('q2:r1520'); // Cambia este rango según tus datos

  const valores = rangoValores.getValues(); // Obtén los valores de la columna L

  // Define el rango de checkboxes en la columna M

  const rangoCheckboxes = hoja.getRange('s2:s1520');

  const checkboxes = [];

  // Define el rango de las columnas P (Capturado por código) y Q (Fecha de ejecución)

  const rangoCapturadoPorCodigo = hoja.getRange('t2:t1520');

  const rangoFechaEjecucion = hoja.getRange('u2:u1520');

  const capturadoPorCodigo = [];

  const fechaEjecucion = [];

  // Obtén la fecha de ejecución (solo una vez, la misma para todas las filas)

  const fechaActual = new Date();

  // Recorre los valores de la columna L

  for (let i = 0; i < valores.length; i++) {

    // Si el valor en la columna L es 1, marca el checkbox (TRUE), si no, desmárcalo (FALSE)

    checkboxes.push([(valores[i][0] === 1)||(valores[i][1] === 1)]);

    // Si el valor es 1, coloca "Capturado por código" en la columna P y la fecha de ejecución en la columna Q

    if ((valores[i][0] === 1)||(valores[i][1] === 1)) {

      capturadoPorCodigo.push(["Capturado por código"]);

      fechaEjecucion.push([fechaActual]);

    } else {

      capturadoPorCodigo.push([null]); // Deja vacío si no es 1

      fechaEjecucion.push([null]); // Deja vacío si no es 1

    }

  }

  // Actualiza los checkboxes en la columna M

  rangoCheckboxes.setValues(checkboxes);

  // Actualiza la columna P (Capturado por código)

  rangoCapturadoPorCodigo.setValues(capturadoPorCodigo);

  // Actualiza la columna Q (Fecha de ejecución)

  rangoFechaEjecucion.setValues(fechaEjecucion);

}

  

/**

  

 * Muestra el diagrama de flujo del proceso completo en un modal.

  

 */

  

function mostrarDiagramaFlujo() {

  

  try {

  

    const html = HtmlService

      .createHtmlOutputFromFile("DiagramaFlujo")

      .setWidth(1000)

      .setHeight(800);

  

    SpreadsheetApp

      .getUi()

      .showModalDialog(

        html,

        "Flujo del proceso — Concentrado Conciliación Operación-Facturación"

      );

  

  } catch (error) {

  

    Logger.log(`Error: ${error.message}`);

  

  }

  

}

## Relacionados

- [[Concentrado Conciliacion Oper-Fact/Concentrado Mensual/CONCENTRADO.gs]] — mismo archivo Concentrado Mensual: la limpieza/organización mensual
