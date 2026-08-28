```/**

 * ============================================================================

 * CREAR NUEVO MES - CONCENTRADO CUADRE DE INGRESOS

 * ----------------------------------------------------------------------------

 * Flujo:

 * 1. Lee la hoja "IDs" del archivo actual.

 * 2. Obtiene la carpeta raíz y la carpeta del mes anterior.

 * 3. Crea la carpeta del mes actual.

 * 4. Duplica el concentrado mensual.

 * 5. Limpia la información del nuevo archivo.

 * 6. Crea la carpeta DIAS.

 * 7. Mueve el archivo de Provisionadas.

 * 8. Actualiza la hoja IDs del nuevo archivo.

 * ============================================================================

 */

  

const MESES = [

  "ENERO",

  "FEBRERO",

  "MARZO",

  "ABRIL",

  "MAYO",

  "JUNIO",

  "JULIO",

  "AGOSTO",

  "SEPTIEMBRE",

  "OCTUBRE",

  "NOVIEMBRE",

  "DICIEMBRE"

];

  

function crearMesNuevo() {

  

  try {

  

    Logger.log("========== INICIO DEL PROCESO ==========");

  

    const ssIDs = SpreadsheetApp.openById("1uuvJ0VWmCcG_5S0dLPI1PwZ_lIZh2gVPnfP9Inast9I");

    const hojaIDs = ssIDs.getSheetByName("IDs PRUEBA");

  

    if (!hojaIDs) {

      throw new Error('No existe la hoja "IDs".');

    }

  

    const datos = hojaIDs.getRange("A2:C6").getValues();

  

    const configuracion = _obtenerConfiguracion(datos);

  

    const hoy = new Date();

  

    const indiceMesActual = hoy.getMonth();

    const indiceMesAnterior = indiceMesActual === 0 ? 11 : indiceMesActual - 1;

  

    const anioActual = hoy.getFullYear();

  

    const mesActual = MESES[indiceMesActual];

    const mesAnterior = MESES[indiceMesAnterior];

  

    Logger.log("Mes actual: " + mesActual);

    Logger.log("Mes anterior: " + mesAnterior);

  

    const carpetaRaiz = DriveApp.getFolderById(configuracion.carpetaRaizId);

  

    const carpetaMesAnterior = DriveApp.getFolderById(

      configuracion.carpetaMesId

    );

  

    const carpetaMesActual = _crearCarpetaMes(

      carpetaRaiz,

      indiceMesActual,

      mesActual,

      anioActual

    );

  

    Logger.log("Carpeta del mes creada correctamente.");

    Logger.log("Buscando concentrado del mes anterior...");

  

    const concentradoAnterior = _buscarArchivoEnCarpeta(

      carpetaMesAnterior,

      `CONCENTRADO CUADRE DE INGRESOS ${mesAnterior} ${anioActual}`

    );

  

    if (!concentradoAnterior) {

      throw new Error(

        `No se encontró el archivo CONCENTRADO CUADRE DE INGRESOS ${mesAnterior} ${anioActual}`

      );

    }

  

    Logger.log("Concentrado encontrado.");

  

    const nombreNuevoConcentrado =

      `CONCENTRADO CUADRE DE INGRESOS ${mesActual} ${anioActual}`;

  

    Logger.log("Realizando copia del concentrado...");

  

    const nuevoArchivo = concentradoAnterior.makeCopy(

      nombreNuevoConcentrado,

      carpetaMesActual

    );

  

    Logger.log("Copia creada.");

  

    const nuevoSpreadsheet = SpreadsheetApp.openById(

      nuevoArchivo.getId()

    );

  

    Logger.log("Limpiando hojas...");

  

    _limpiarHojas(nuevoSpreadsheet);

  

    Logger.log("Creando carpeta DIAS...");

  

    const carpetaDias = carpetaMesActual.createFolder("DIAS");

  

    Logger.log("Buscando archivo Provisionadas...");

  

    const provisionadas = _buscarArchivoEnCarpeta(

      carpetaMesAnterior,

      "CUADRE DE INGRESOS CONCENTRADO PROVISIONADAS"

    );

  

    if (!provisionadas) {

      throw new Error(

        "No se encontró el archivo CUADRE DE INGRESOS CONCENTRADO PROVISIONADAS."

      );

    }

  

    Logger.log("Moviendo archivo Provisionadas...");

  

    provisionadas.moveTo(carpetaMesActual);

  

    Logger.log("Actualizando hoja IDs del nuevo archivo...");

  

    _actualizarIDs(

      nuevoSpreadsheet,

      carpetaRaiz.getName(),

      configuracion.carpetaRaizId,

      carpetaMesActual.getName(),

      carpetaMesActual.getId(),

      carpetaDias.getId(),

      nombreNuevoConcentrado,

      nuevoArchivo.getId(),

      provisionadas.getId(),

      mesActual

    );

  

    Logger.log("========== PROCESO FINALIZADO ==========");

  

  } catch (error) {

  

    Logger.log(error);

  

    throw error;

  

  }

  

}

  

/**

 * Obtiene la configuración almacenada en la hoja IDs.

 */

function _obtenerConfiguracion(datos) {

  

  if (datos.length < 5) {

    throw new Error("La hoja IDs no contiene la información suficiente.");

  }

  

  return {

    carpetaRaizNombre: datos[0][0],

    carpetaRaizId: datos[0][1],

  

    carpetaMesNombre: datos[1][0],

    carpetaMesId: datos[1][1],

  

    carpetaDiasNombre: datos[2][0],

    carpetaDiasId: datos[2][1],

  

    concentradoNombre: datos[3][0],

    concentradoId: datos[3][1],

  

    provisionadasNombre: datos[4][0],

    provisionadasId: datos[4][1]

  };

  

}

  

/**

 * Crea la carpeta correspondiente al mes actual.

 */

function _crearCarpetaMes(

  carpetaRaiz,

  indiceMes,

  mes,

  anio

) {

  

  const nombreCarpeta =

    `${indiceMes + 1}. ${mes} ${anio}`;

  

  const carpetas = carpetaRaiz.getFoldersByName(

    nombreCarpeta

  );

  

  if (carpetas.hasNext()) {

    throw new Error(

      `La carpeta "${nombreCarpeta}" ya existe.`

    );

  }

  

  Logger.log(

    `Creando carpeta ${nombreCarpeta}...`

  );

  

  return carpetaRaiz.createFolder(

    nombreCarpeta

  );

  

}

  

/**

 * Busca un archivo por nombre exacto dentro de una carpeta.

 */

function _buscarArchivoEnCarpeta(

  carpeta,

  nombreArchivo

) {

  

  const archivos =

    carpeta.getFilesByName(nombreArchivo);

  

  if (!archivos.hasNext()) {

    return null;

  }

  

  return archivos.next();

  

}

  

/**

 * Limpia la información del archivo recién creado.

 */

function _limpiarHojas(ss) {

  

  Logger.log("Limpiando hoja DOCUMENTOS FACTURACION...");

  

  const hojaFacturacion = ss.getSheetByName(

    "DOCUMENTOS FACTURACION"

  );

  

  if (!hojaFacturacion) {

    throw new Error(

      'No existe la hoja "DOCUMENTOS FACTURACION".'

    );

  }

  

  hojaFacturacion.getRange(

    "A2:AI"

  ).clearContent();

  

  Logger.log("Limpiando hoja MOVS...");

  

  const hojaMovs = ss.getSheetByName(

    "MOVS"

  );

  

  if (!hojaMovs) {

    throw new Error(

      'No existe la hoja "MOVS".'

    );

  }

  

  hojaMovs.getRange(

    "A2:V"

  ).clearContent();

  

}

  

/**

 * Actualiza la hoja IDs del nuevo archivo.

 */

function _actualizarIDs(

  ss,

  nombreCarpetaRaiz,

  idCarpetaRaiz,

  nombreCarpetaMes,

  idCarpetaMes,

  idCarpetaDias,

  nombreConcentrado,

  idConcentrado,

  idProvisionadas,

  mesActual

) {

  

  const hoja = ss.getSheetByName("IDs");

  

  if (!hoja) {

    throw new Error(

      'No existe la hoja "IDs".'

    );

  }

  

  hoja.getRange("A2:C6").setValues(

    [

      [nombreCarpetaRaiz, idCarpetaRaiz, mesActual],

      [nombreCarpetaMes, idCarpetaMes, mesActual],

      ["DIAS", idCarpetaDias, mesActual],

      [nombreConcentrado, idConcentrado, mesActual],

      ["CUADRE DE INGRESOS CONCENTRADO PROVISIONADAS", idProvisionadas, mesActual]

    ]

  );

  

  SpreadsheetApp.flush();

  

  Logger.log("Hoja IDs actualizada correctamente.");

  

}
