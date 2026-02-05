function onOpen() { 
    var ui = SpreadsheetApp.getUi();
  var mensaje = "Recuerda que esta plantilla contiene listas anidadas y recibe información de otros archivos:"
    + "\n- 🚫 No agregar o quitar columnas y filas."
    + "\n- 🚫 No alterar fórmulas."
    + "\n- 🚫 No modificar la posición de las tablas o el rango."
    + "\n- 🔷 ASEGÚRATE DE LLENAR TODAS LAS COLUMNAS IDENTIFICADAS EN COLOR AZÚL"
    + "\n- ✅ Para un uso adecuado del archivo consulta tu instrucción de trabajo P-PS-IT-001_ SOLICITUD DE GASTOS CEOA REV 0.1"
    + "\n- ☎︎ Contacta a 'Optimización' para realizar modificaciones. V21";
  
  ui.alert(mensaje);


    ui.createMenu('📅 | Backup')
    .addItem('1. Copiado 003 - GASTOS PERSONALES', 'metodoConTablaGastos')
    .addItem('2. Backup del 10-R', 'allFunct')
    .addToUi();

    ui.
    createMenu("PAPELETAS"). // 10R
    // addItem("Papeletas 5R","mandarInfoPapeletasDir").
    addItem("Papeletas 10R","mandarInfoPapeletasPer").
    // addItem("Borrar 5R","borrar5R").
    addItem("Borrar 10R","borrar10R").
    addToUi();

     //boton para tarjetas
    ui.
    createMenu("TARJETAS"). // 5R
    addItem("CARGO DE TARJETAS","accion").
    addToUi();
}

function metodoConTablaGastos(){
  ejemploFuncion();
}

function allFunct() {
  copiarMasterA10R();//SE MANDA AL MASTER LA INFORMACION 10/06/2025
  copiarArchivosG2(); //implementado 03/06/2024 
  copiarFormatoAGoogleDrive();
  //mover y bloqueo de usuarios
  bloquearTodasLasHojas();
  moverArchivo();
}

function copiaYpegarDatos_SD(hojaOrigen, hojaDestino, rangoOrigen, columnaInicio, rangoLetras) { //quedo 4:37, G1,G2, G3 de SD
  var datos = hojaOrigen.getRange(rangoOrigen).getValues();
  let ultimaFila = encontrarUltimaFilaEnColumna(hojaDestino, rangoLetras);
  hojaDestino.getRange(ultimaFila + 1, columnaInicio, datos.length, datos[0].length).setValues(datos);
}

function encontrarUltimaFilaEnColumna(hojaDestino, rangoLetras) {
  var valores = hojaDestino.getRange(rangoLetras).getValues();
  var ultimaFila = 0;

  for (var i = valores.length - 1; i >= 0; i--) {
    if (valores[i].some(cell => cell !== "")) {
      ultimaFila = i + 1;
      break;
    }
  }

  return ultimaFila;
}


function copiarYpegarDatos_FT12(hojaOrigen, hojaDestino, rangoOrigen, columnaInicio) { //FUNCIONA 2:32
  // Obtener los datos desde la hoja de origen (D5:N407)
  var datos = hojaOrigen.getRange(rangoOrigen).getDisplayValues(); //correccion 29/04/2024

  // Encontrar la última fila con valores en la hoja de destino (columna C)
  var ultimaFilaDestino = hojaDestino.getLastRow();

  // Pegar los datos en la hoja de destino (C:M) después de la última fila con valores
  hojaDestino.getRange(ultimaFilaDestino + 1, columnaInicio, datos.length, datos[0].length).setValues(datos);
}

function copiarFormatoAGoogleDrive() {
  try {
    var hojaDeCalculo = SpreadsheetApp.getActiveSpreadsheet();// Obtén la hoja de cálculo activa
    var currentDate = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd");
    var nombreArchivo = hojaDeCalculo.getName();// Obtén el nombre de la hoja de cálculo
    var nuevaHojaDeCalculo = hojaDeCalculo.copy('copia ' + nombreArchivo + currentDate); // Crea una nueva hoja de cálculo
    var idNuevoArchivo = nuevaHojaDeCalculo.getId();// Obtén la ID del archivo de la nueva hoja de cálculo
    var nuevoNombre = 'Copia de ' + nombreArchivo; // Cambia el nombre del archivo copiado
    DriveApp.getFileById(idNuevoArchivo).setName(nuevoNombre); // Puedes ajustar el nuevo nombre según tus necesidades
    var carpetaDestino = DriveApp.getFolderById('14kMb7oELKRTEEcb1DzB9hYVD2aVT4i5T'); // Reemplaza 'ID_DE_LA_CARPETA' con la ID de la carpeta destino 10HcEi2RlaT1U_BwBsWzcEzW0JBJfQb8q //carpeta mia: 1NB8_H0vuuGaxXzn0n2Wi1FlBqDPYxc7e
    DriveApp.getFileById(idNuevoArchivo).moveTo(carpetaDestino); // Mueve el nuevo archivo a la carpeta de destino
    Logger.log('Copia de formato creada y guardada en la carpeta destino. Nombre del archivo: ' + nuevoNombre); // Registra el nombre del archivo en el registro

    /*modificacion 23/10/2024 */
    var hojaOrigen = hojaDeCalculo.getSheetByName("G2 - GASTOS ABBY (Principal)");
    var hojaDestino = nuevaHojaDeCalculo.getSheetByName("G2 - GASTOS ABBY (Principal)");
    var columnas = ['D', 'F', 'H', 'J', 'L', 'N', 'P', 'R', 'T', 'V', 'X', 'Z', 'AB'];
    columnas.forEach(function(col) {
      var datos = hojaOrigen.getRange(`${col}68`).getValues();
      hojaDestino.getRange(`${col}5`).setValues(datos);
    }); 

     var hojasSD = [
     /* { origen: "G2 - GASTOS ABBY (Principal)", destino: "ACUMULADO GPA", rango: "D77:AK577", columnaInicio: 2, rangoLetras: "B:AB" },//=  COPIO*/
      { origen: "ENTRECUENTAS", destino: "SD", rango: "O2:Q50", columnaInicio: 3, rangoLetras: "C:E" }//modificado 11/09/2024 = COPIA
    ];

    hojasSD.forEach(function(hoja) {
      var hojaOrigen = hojaDeCalculo.getSheetByName(hoja.origen);
      var hojaDestino = nuevaHojaDeCalculo.getSheetByName(hoja.destino);
      copiaYpegarDatos_SD(hojaOrigen, hojaDestino, hoja.rango, hoja.columnaInicio,hoja.rangoLetras);
    });

    /*g1 Y g1 FONDEO DE TARJETAS */
     var hojasDatosFT = [
      { origen: "ENTRECUENTAS", destino: "FONDEO DE TARJETAS", rango: "P93:U126", columnaInicio: 3 }//modificado 11/09/2024 C-H = COPIO
    ];

    hojasDatosFT.forEach(function(hoja) {
      var hojaOrigen = hojaDeCalculo.getSheetByName(hoja.origen);
      var hojaDestino = nuevaHojaDeCalculo.getSheetByName(hoja.destino);
      copiarYpegarDatos_FT12(hojaOrigen, hojaDestino, hoja.rango, hoja.columnaInicio);
    });

    

    limpiarCeldasEnHojas(nuevaHojaDeCalculo);
    eliminarRangosEnHojaProtegida(nuevaHojaDeCalculo);

  } catch (error) {
    Logger.log('Error: ' + error.toString());
  }
} 


function limpiarCeldasEnHojas(nuevaHojaDeCalculo) {
  var hojas = [
    { nombre: "G2 - GASTOS ABBY (Principal)", rangos: ["D6:AC67", "E77:AN464"] },//, "AA74"
    { nombre: "ENTRECUENTAS", rangos: ["B4:H300", "I3:N110", "P3:U50", "P54:AG89", "P93:AA126"] }
  ];

  hojas.forEach(function(hoja) {
    var sheet = nuevaHojaDeCalculo.getSheetByName(hoja.nombre);
    hoja.rangos.forEach(function(rango) {
      sheet.getRange(rango).clearContent();
    });
  });
}

//funciona 29/04/2025////
function eliminarRangosEnHojaProtegida(nuevaHojaDeCalculo) {
  try {
    var hoja = nuevaHojaDeCalculo.getSheetByName("HistorialEjecuciones"); // Nombre de la hoja protegida
    if (!hoja) {
      Logger.log("La hoja 'HistorialEjecuciones' no existe.");
      return;
    }

    // Buscar la protección de la hoja
    var protecciones = hoja.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    var proteccion = protecciones.length > 0 ? protecciones[0] : null;

    // Guardar los editores actuales si hay protección
    var editores = proteccion ? proteccion.getEditors() : [];

    // Si hay protección, eliminarla temporalmente
    if (proteccion) {
      proteccion.remove();
      Logger.log("Protección eliminada temporalmente.");
    } else {
      Logger.log("La hoja no tiene protección activa.");
    }

    // Limpiar los rangos especificados
    var rangos = ["A1:E1", "A2:E2", "A3:E3"];
    rangos.forEach(function (rango) {
      hoja.getRange(rango).clearContent();
    });
    Logger.log("Celdas limpiadas correctamente.");

    // Restaurar la protección si existía
    if (proteccion) {
      var nuevaProteccion = hoja.protect();
      nuevaProteccion.setWarningOnly(false); // La protección es estricta
      editores.forEach(function (editor) {
        nuevaProteccion.addEditor(editor); // Restablecer los editores originales
      });
      Logger.log("Protección restaurada.");
    }

  } catch (error) {
    Logger.log("Error: " + error.toString());
  }
}

function copiarArchivosG2() { //saca a una copia de g2 y de ENTRECUENTAS ==funciona
  var hojaDeCalculo = SpreadsheetApp.getActiveSpreadsheet();
  var currentDate = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd");
  var nombreArchivo = hojaDeCalculo.getName();
  var hojasDatos = ["G2 - GASTOS ABBY (Principal)", "ENTRECUENTAS"];
    
  var carpetaBackup = DriveApp.getFolderById("1u4n5zAO3Dsp9Uwxz2DuFC_XDcJmasz2x");//id de la carpeta a depositar. //carpeta mia id:1kez8C5PfEDB4PHH0I6fEnMje-N76YCPX

  //Crear un nuevo archivo donde se copiaran las hojas
  var nombreBackup = 'Backup - ' + nombreArchivo + ' - ' + currentDate;
  nuevaHojaDeCalculo = SpreadsheetApp.create(nombreBackup);

  hojasDatos.forEach(function(hojaNombre){
    var hojaOrigen = hojaDeCalculo.getSheetByName(hojaNombre);
    if (!hojaOrigen) {
      Logger.log('No se encontró la hoja con el nombre: ' + hojaNombre);
      return;
    }

    // Copiar la hoja al archivo nuevo
    var hojaNueva = hojaOrigen.copyTo(nuevaHojaDeCalculo);
    hojaNueva.setName(hojaNombre);
  });

  // Eliminar la hoja inicial creada al momento de crear el nuevo archivo
  var hojaInicial = nuevaHojaDeCalculo.getSheets()[0];
  nuevaHojaDeCalculo.deleteSheet(hojaInicial);
  
  // Mover el archivo a la carpeta de respaldo
  var idNuevoArchivo = nuevaHojaDeCalculo.getId();
  DriveApp.getFileById(idNuevoArchivo).moveTo(carpetaBackup);
}

/////////////////////////////////////MAESTER ya esta31/01/2025
function ejemploFuncion() {//principal
  try {
    // Lógica de tu función
    Logger.log("Ejecutando función ejemplo...");
    
    // Registro exitoso de la ejecución
    registrarEjecucion('ejemploFuncion', 'Éxito');
  } catch (error) {
    // Registro en caso de fallo
    registrarEjecucion('ejemploFuncion', 'Error: ' + error.message);
  }
}

function registrarEjecucion(funcionNombre, resultado) {
  var hojaHistorial = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('HistorialEjecuciones'); // Obtén la hoja llamada 'HistorialEjecuciones'
  var ui = SpreadsheetApp.getUi(); //obtener la interfaz del usuario para mostrar alertas
  
  if (!hojaHistorial) {
    hojaHistorial = SpreadsheetApp.getActiveSpreadsheet().insertSheet('HistorialEjecuciones');
    hojaHistorial.appendRow(['Fecha', 'Hora', 'Función', 'Usuario', 'Resultado']);
  }

  var fechaActual = new Date();
  var fechaFormato = Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  // Verificar si ya hay un registro de esta función en el día actual
  var datos = hojaHistorial.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    var fechaEnHoja = datos[i][0]; // Fecha de la hoja
    // Si la fecha en la hoja no está formateada correctamente, intenta formatearla
    if (fechaEnHoja instanceof Date) {
      fechaEnHoja = Utilities.formatDate(fechaEnHoja, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    // Comparar la fecha formateada de la hoja con la fecha actual
    if (fechaEnHoja === fechaFormato && datos[i][2] === funcionNombre) {
      Logger.log("La función " + funcionNombre + " ya se ejecutó hoy.");
      ui.alert("La función '" + funcionNombre + "' ya ha sido registrada hoy."); // Mostrar alerta al usuario
      return; // Si ya existe un registro de esta función en el día actual, no hacemos nada
    }
  }

   // Si no existe un registro para hoy, se agrega uno nuevo
    var usuario = Session.getActiveUser().getEmail(); // Obtener el correo del usuario que ejecuta el script
    
    // Añadir un nuevo registro en la hoja de historial
    hojaHistorial.appendRow([
      fechaFormato, // Solo la fecha, no la hora
      Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), 'HH:mm:ss'), // La hora de ejecución
      funcionNombre,
      usuario || "Anónimo",
      resultado
    ]);

    concentrado(); //si no esta registrada el metodo.
}

function concentrado(){//conectado al hijo.
  var libroOrigen = SpreadsheetApp.openById('19fionVnXuVOe2Ex5WthuF5te0YrZbPz-HN8YvV9XwuA'); // P-PS-FT-003_Rev.2_SOLICITUD DE GASTOS PERSONALES 2025 == 19fionVnXuVOe2Ex5WthuF5te0YrZbPz-HN8YvV9XwuA
  var libroDestino = SpreadsheetApp.getActiveSpreadsheet(); // Obtener el libro activo //P-AE-FT-007_10-R.

  var hojaG2 = [
    { origen: "S.Gastos Personales",  destino: "G2 - GASTOS ABBY (Principal)", rango: "A:AJ"} /* letra 25 = AA  => AC = 27(FECHA DE PAGO)*/ //A:AH a A:AI 28 a 29(FECHA DE PAGO)
  ];

  hojaG2.forEach(function(hoja) {
    var hojaOrigen = libroOrigen.getSheetByName(hoja.origen);
    var hojaDestino = libroDestino.getSheetByName(hoja.destino);
    metodo10R(hojaOrigen, hojaDestino, hoja.rango);
  });
}

/*funcion para el baciado para 10 R */
function metodo10R(hojaOrigen, hojaDestino, rango) {//funciona original == 19/09/2024
  var today = new Date();
  
  // Formatear la fecha de hoy en DD/MM/YY
  var fomateoTaday = Utilities.formatDate(today, Session.getScriptTimeZone(), 'dd/MM/yy');

  // Rango de pegado en la hoja destino (D77:AD254) //D77:AK577 //E77:AL262
  var filaInicioDestino = 77; // Fila 77 en la hoja destino
  var filaFinalDestino = 364; // Fila 254 en la hoja destino
  var columnaInicioDestino = 5; // Columna E en la hoja destino
  var columnaFinalDestino = 40; // Columna  29 = AD => 37 = AK en la hoja destino 6 AK // AL => 38 // AM 39

  var dataValues = hojaOrigen.getRange(rango).getValues();

  // Encontrar la última fila con datos en el rango D77:AD254  === D77:AK577en la hoja destino
  var datosDestino = hojaDestino.getRange(filaInicioDestino, columnaInicioDestino, filaFinalDestino - filaInicioDestino + 1, columnaFinalDestino - columnaInicioDestino + 1).getValues();
  
  var ultimaFilaDestino = filaInicioDestino;
  //var ultimaFilaDestino1 = encontrarUltimaFilaEnColumna(hojaDestino);
  
  // Verificar cuál es la última fila ocupada en el rango D77:AD254
  for (var i = 0; i < datosDestino.length; i++) {
    var fila = datosDestino[i];
    // Si la fila no está vacía (es decir, tiene algún valor), entonces es una fila ocupada
    if (fila.some(function (cell) { return cell !== "" && cell !== null; })) {
      ultimaFilaDestino = filaInicioDestino + i + 1;
    }
  }
  
  // Controlar la fila actual para pegar datos
  var filaDestino = ultimaFilaDestino;

  // Recorrer todas las filas de datos de la hoja origen
  for (var i = 0; i < dataValues.length && filaDestino <= filaFinalDestino; i++) {
     var dataFecha = dataValues[i][29]; // Columna AC (fecha de pago) =rangoLetra 28 a 29

    // Intentar convertir la fecha si no es un objeto Date
    if (typeof dataFecha === 'string') {
      dataFecha = new Date(dataFecha); // Convertir cadena a fecha
    }

    // Verificar si dataFecha es un objeto Date válido
    if (dataFecha instanceof Date && !isNaN(dataFecha.getTime())) {
      // Formatear la fecha obtenida
      var fomateoFecha = Utilities.formatDate(dataFecha, Session.getScriptTimeZone(), 'dd/MM/yy');
      
      // Comparar con la fecha de hoy
      if (fomateoFecha === fomateoTaday) { // Verificar si la celda A tiene la fecha de hoy
        Logger.log("La fecha de hoy es: " + fomateoTaday);

        // Verificar si la celda 24 = Z  => 26 AD es "PAGADO" o "PAGADO Y COMPROBANTE EN CARPETA"
        if (dataValues[i][28] === "PAGADO" || dataValues[i][28] === "PAGADO Y COMPROBANTE EN CARPETA") { //28 a 29
          // Pegar datos en la hoja destino dentro del rango D77:AD254
          hojaDestino.getRange(filaDestino, columnaInicioDestino, 1, columnaFinalDestino - columnaInicioDestino + 1)
            .setValues([dataValues[i].slice(0, columnaFinalDestino - columnaInicioDestino + 1)]); // Pegar solo hasta la col.AK(para poder agrenda el rango de pegado a la hoja destino debes de modificar lo que es la variable "columnaFinalDestino" y contar bien es que posicion esta la posicion)
          
          filaDestino++; // Incrementar la fila destino para la siguiente inserción
        }
      }
    } else {
      Logger.log("La celda en la fila " + (i + 1) + " no contiene una fecha válida.");
    }
  }

  if (filaDestino > filaFinalDestino) {
    Logger.log("Se alcanzó el límite de filas en el rango de destino.");
  } else {
    Logger.log("Pegado finalizado, datos hasta la fila: " + (filaDestino - 1));
  }
}
//////////copiado10R_CONCENTRADO///////////
function copiarMasterA10R() {//copiado = condicion para que copie fecha de hoy y todos los status menos nuevo y vacio.
  var libroOrigen = SpreadsheetApp.getActiveSpreadsheet();
  //var libroDestino = SpreadsheetApp.openById('11FniTDx4I_FtuysdLikwMbRhbC_wOaZMuE5mfdhrPYc'); // Master = 19eYrBuMNHkFySoPkYwGrst842lEOIHVBK6E0ozyb2SY
  var libroDestino = SpreadsheetApp.openById('1eQHolmm_3BPFSkG4jZvRgACWLxiLkXsNNJv2AubijOQ'); // Master = 19eYrBuMNHkFySoPkYwGrst842lEOIHVBK6E0ozyb2SY
    
  var hojaOrigen = libroOrigen.getSheetByName("G2 - GASTOS ABBY (Principal)");
  var hojaDestino = libroDestino.getSheetByName("ANUAL");

  // Obtener la fecha de ayer
  var ayer = new Date();
  // Restar 1 día (24 horas)
  ayer.setDate(ayer.getDate() - 1);
  var fomateoAyer = Utilities.formatDate(ayer, Session.getScriptTimeZone(), 'dd/MM/yy');
    
  // Obtener la fecha actual formateada
  var today = new Date();
  var fomateoToday = Utilities.formatDate(today, Session.getScriptTimeZone(), 'dd/MM/yy');

   // Obtener los valores de la hoja origen
  var datos = hojaOrigen.getRange("E77:AN464").getValues();//A:AO a A:AP

  // Preparar un arreglo para las filas que cumplen las condiciones
  var filasParaPegar = [];

  for (var i = 0; i < datos.length; i++) {
    var dataFecha = datos[i][29]; // Columna AB (índice 27) //fecha de pago AD 29

    // Validar si el dato en la columna AB es una fecha válida
    if (dataFecha instanceof Date && !isNaN(dataFecha.getTime())) {
      var fomateoFecha = Utilities.formatDate(dataFecha, Session.getScriptTimeZone(), 'dd/MM/yy');

      // Verificar si coincide con la fecha de hoy
      if (fomateoFecha === fomateoToday || fomateoFecha === fomateoAyer) {
        // Verificar condiciones en la columna Z (índice 26) col. Status AC 28
        if (datos[i][28] === "PAGADO Y COMPROBANTE EN CARPETA") {
          filasParaPegar.push(datos[i]); // Añadir fila para pegar
        }
      }
    }
  }

  // Pegar todas las filas que cumplen las condiciones en la hoja destino
  if (filasParaPegar.length > 0) {
    var ultimaFilaDestino = hojaDestino.getLastRow();
    hojaDestino.getRange(ultimaFilaDestino + 1, 1, filasParaPegar.length, filasParaPegar[0].length)
      .setValues(filasParaPegar);
    Logger.log(filasParaPegar.length + " filas copiadas a la hoja destino.");
  } else {
    Logger.log("No se encontraron filas que cumplan las condiciones para copiar.");
  }
}
///Papeletas//-> Gustavo
const SSID = SpreadsheetApp.getActiveSpreadsheet().getId();

function mandarInfoPapeletasDir(){ // 5R
  Papeletas.papeletasInfoDir(SSID);
}

function mandarInfoPapeletasPer(){ // 10R
  Papeletas.papeletasInfoPer(SSID);
}

function borrar5R() {
  Papeletas.eraseColumns5R(SSID);
}

function borrar10R() {
  Papeletas.eraseColumns10R(SSID);
}

////////////nueva implementacion 17/10/2025/////////
function bloquearTodasLasHojas() {
  var libroOrigen = SpreadsheetApp.getActiveSpreadsheet();
  var hojas = libroOrigen.getSheets();
  
  // Lista de correos que SÍ tendrán permiso de editar
  var usuariosPermitidos = [
    "verificador@kabzo.org",
    //"optimizacion@kabzo.org",
    //"analistaprocesos2@kabzo.org"
  ];
  
  hojas.forEach(function(hoja) {
    try {
      Logger.log("Intentando proteger hoja: " + hoja.getName());
      
      // Crear o actualizar protección en la hoja
      var proteccion = hoja.protect().setDescription("Protección automática: " + hoja.getName());
      
      // Quitar todos los editores actuales
      var editores = proteccion.getEditors();
      if (editores.length > 0) {
        proteccion.removeEditors(editores);
      }
      
      // Permitir solo a estos usuarios
      if (usuariosPermitidos.length > 0) {
        proteccion.addEditors(usuariosPermitidos);
      }
      
      // Desactivar edición por dominio (si aplica)
      if (proteccion.canDomainEdit()) {
        proteccion.setDomainEdit(false);
      }

      Logger.log("✅ Protección aplicada correctamente a: " + hoja.getName());

    } catch (e) {
      Logger.log("❌ Error en hoja: " + hoja.getName() + " → " + e.message);
    }
  });
  
  Logger.log("Proceso finalizado. Revisa el registro (Ctrl + Enter o Ver > Registro).");
}

function moverArchivo() {//mueve 29/09/2025 -> unidad compartida
  // Reemplaza 'ID_CARPETA_DESTINO' con el ID de la carpeta a la que deseas mover el archivo.
  var idCarpetaDestino = '1NNgMCbCQioqTvp35nN9XTLy2xB62EYeg';//carpeta de backUp de azael

  // Si el script está vinculado a la hoja de cálculo que quieres mover,
  // puedes obtener su ID automáticamente.
  var archivoActual = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  
  // O, si prefieres mover un archivo específico por su ID, usa esta línea en su lugar:
  // var archivoMover = DriveApp.getFileById('ID_DEL_ARCHIVO_A_MOVER');

  var carpetaDestino = DriveApp.getFolderById(idCarpetaDestino);
  
  // Mueve el archivo a la carpeta de destino.
  archivoActual.moveTo(carpetaDestino);
}

/////////////////CARGO DE TARJETAS/////////////////
function accion(){
  var libroOrigen = SpreadsheetApp.openById('1IBgDOupqxGimF0a-SLR9vKKLPE0tWqlHHYanczLEm8o'); // CARGO DE TARJETAS
  var hojaOrigen = libroOrigen.getSheetByName("CARGOS");

  BLOQUEO.agrandarRangoBloqueoCondicionalV2(hojaOrigen);
  tarjetas01();
}

function tarjetas01() { //boton para llevar la informacion.
  var libroOrigen = SpreadsheetApp.openById('1IBgDOupqxGimF0a-SLR9vKKLPE0tWqlHHYanczLEm8o'); // CARGO DE TARJETAS
  var hojaOrigen = libroOrigen.getSheetByName("CARGOS");

  var libroDestino = SpreadsheetApp.getActiveSpreadsheet(); // G1 = 5R
  var hojaDestino = libroDestino.getSheetByName("ENTRECUENTAS");


  var rango = hojaOrigen.getRange("B4:K").getValues();


  //fecha de comparacion
  var today = new Date();
  var fomateoToday = Utilities.formatDate(today, Session.getScriptTimeZone(), 'dd/MM/yy');

  // Preparar un arreglo para las filas que cumplen las condiciones
  var filasParaPegar = [];

  for (var i = 0; i < rango.length; i++) {
    var dataFecha = rango[i][6]; // Columna AB (índice 27) //28 a 29

    // Validar si el dato en la columna AB es una fecha válida
    if (dataFecha instanceof Date && !isNaN(dataFecha.getTime())) {
      var fomateoFecha = Utilities.formatDate(dataFecha, Session.getScriptTimeZone(), 'dd/MM/yy');
        //if (fomateoFecha === fomateoToday) { //27 a 28
        if (fomateoFecha === fomateoToday) { //27 a 28
          // Verificar condiciones en la columna Z (índice 26)
          if (rango[i][7] === "CONSULTORIA") {
            filasParaPegar.push(rango[i]); // Añadir fila para pegar
          }
        }
      
    }
  }

  if (filasParaPegar.length > 0) {

    //P93:U126 
    var inicioFila = 93;
    var maxFilas = 126 - 93 + 1; 

    var datosDestino = hojaDestino.getRange(93, 16, 126 - 93 + 1, 126).getValues();

    var ultimaFilaDestino = 93;

    for (var i = 0; i < datosDestino.length; i++) {
      var fila = datosDestino[i];
      if (fila.some(function (cell) { return cell !== "" && cell !== null; })) {
        ultimaFilaDestino = inicioFila + i + 1;
      }
    }

    var filaDestino = ultimaFilaDestino;

    if (filasParaPegar.length > maxFilas) {
      filasParaPegar = filasParaPegar.slice(0, maxFilas);
      Logger.log("Se truncaron los datos para ajustarse al rango permitido.");
    }

    var forTodayFol = Utilities.formatDate(today,Session.getScriptTimeZone(), 'ddMMyy');

    const filasParaPegar2 = filasParaPegar.map(r => {//negativo //numero  NL1912250013852
     //consecutivo++; // aumenta por cada fila

      //secuencia por polimorfismo.
      var secuencia = generarFolioCentral();
      
      var secuenciaStr = String(secuencia).padStart(7, '0');

      //var folio = 'NL' + forTodayFol + secuenciaStr;
      var folio = 'VV' + forTodayFol + secuenciaStr;// 30/01/2026 modificacion

      const s = String(r[0] ?? '');
      const sClean = s.replace(/^-?/, ''); // quita signo si existe
      return [ '-' + sClean, r[1], r[2], r[3], r[6], folio];
    });

    const filasParaPegar3 = filasParaPegar.map(r => { //positivo
      return [ r[0], r[1], r[2], r[3], r[6], r[5]  ];
    });

    //negativo
    const rangoTransformado2 = hojaDestino.getRange(filaDestino, 16, filasParaPegar2.length, filasParaPegar2[0].length);
    rangoTransformado2.setValues(filasParaPegar2);
    

    //ultima dila destino
    var ultimaColDest = ultimaFilaNoVaciaV10(hojaDestino);

    //positivo
    const rangoTransformado3 = hojaDestino.getRange(ultimaColDest + 1, 2, filasParaPegar3.length, filasParaPegar3[0].length);
    rangoTransformado3.setValues(filasParaPegar3);
  }

  //var fechaReporte = //FECHA EN QUE SE REPORTA 
}

function ultimaFilaNoVaciaV10(hoja) {
 // const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SOLICITUDES 2024");
  if (!hoja) {
    Logger.log("La hoja "+ hoja + " no existe.");
    return;
  }
  
  const columna = hoja.getRange("C:C").getValues(); // Obtiene todos los valores de la columna B
  let ultimaFila = 0;

  // Iterar desde el final hacia arriba para encontrar la última fila con datos
  for (let i = columna.length - 1; i >= 0; i--) {
    if (columna[i][0] !== "") {
      ultimaFila = i + 1; // +1 porque los índices comienzan en 0
      break;
    }
  }

  return ultimaFila;
 // Logger.log(`La última fila con datos en la columna A de 'solicitudes' es: ${ultimaFila}`);
}


function generarFolioCentral() {
  var url = 'https://script.google.com/macros/s/AKfycbyYlSe04wahYBFAyApCeCPlL2r0QBrO0yg8V2P0lSQKuB9ARPLQXp5auSPOTK9XBLzE/exec'; // /exec

  var response = UrlFetchApp.fetch(url, {
    method: 'post',
    muteHttpExceptions: true
  });

  var codigo = response.getResponseCode();
  var texto = response.getContentText();

  if (codigo !== 200) {
    throw new Error(
      'Error al generar folio. Código: ' + codigo + 
      ' Respuesta: ' + texto
    );
  }

  return Number(texto);
}
