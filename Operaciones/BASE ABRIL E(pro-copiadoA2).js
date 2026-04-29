/**
 * MENÚ PERSONALIZADO
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🎬 ACCIONES")
    .addItem("💃🏻 ENVIAR INFO SEMANAL", "copiarHojasActivas")
    .addToUi();
}


/**
 * CONFIGURACIÓN
 */
//const DESTINO_ID = '1PWwVK6pYUYyRR4CauZcYcPwwDGq4OX9dkN468dTIxvc'; //LINK SEMANAL

function extraerIdDeUrl(url) {
  const match = url.match(/[-\w]{25,}/);
  if (!match) {
    throw new Error('URL inválida en B2');
  }
  return match[0];
}

const CONFIG = {
  patrones: [
    'MOVS DIARIOS',
    'EC SD',
    'CONCENTRADO',
    'REP',
    'MOV',
    'SD',
    'OPS',
    'EC',
    'USD',
    'CONCENTRADO E',
    'FONDO',
    'INTS',
    'PROMENS'
  ]
};


/**
 * FUNCIÓN PRINCIPAL
 */
function copiarHojasActivas() {
  const log = [];

  try {
    // Obtener URL desde B2
    var activo = SpreadsheetApp.getActiveSpreadsheet();
    var hojaConfig = activo.getSheetByName('LINK SEMANAL'); // hoja donde está B2
    var urlDestino = hojaConfig.getRange('B2').getValue();

    // Extraer ID del spreadsheet desde la URL
    var idDestino = extraerIdDeUrl(urlDestino);

    // Abrir archivo destino
    var destino = SpreadsheetApp.openById(idDestino);
    //const destino = SpreadsheetApp.openById(DESTINO_ID);

    const todasLasHojas = activo.getSheets();
    const hojasEncontradas = filtrarHojasPorPatron(todasLasHojas, log);

    if (hojasEncontradas.length === 0) {
      activo.toast('⚠️ No se encontraron hojas', 'ERROR', 5);
      return;
    }

    activo.toast(`🚀 Copiando ${hojasEncontradas.length} hojas...`, 'PROCESO', 3);

    const resultados = {
      copiadas: [],
      errores: []
    };

    for (const hojaOrigen of hojasEncontradas) {
      try {
        const nombreOriginal = hojaOrigen.getName();

        log.push(`📋 Copiando: ${nombreOriginal}`);

        // Copiar hoja al archivo destino
        const hojaCopiada = hojaOrigen.copyTo(destino);

        // Generar nombre único
        const nombreFinal = generarNombreUnico(destino, nombreOriginal);
        hojaCopiada.setName(nombreFinal);

        log.push(`✅ Guardada como: ${nombreFinal}`);
        resultados.copiadas.push(nombreFinal);

      } catch (error) {
        const errorMsg = `${hojaOrigen.getName()}: ${error.message}`;
        log.push(`❌ ERROR: ${errorMsg}`);
        resultados.errores.push(errorMsg);
      }
    }

    let mensaje = `✅ Copiadas: ${resultados.copiadas.length}\n`;
    mensaje += `❌ Errores: ${resultados.errores.length}`;

    activo.toast(mensaje, 'FINALIZADO', 10);
    mostrarLogCompleto(log, activo);

    //ana
     const hojaEliminar = destino.getSheetByName('Hoja 1') 
                    || destino.getSheetByName('Sheet1');

      if (hojaEliminar && destino.getSheets().length > 1) {
        destino.deleteSheet(hojaEliminar);
        log.push('🗑️ Hoja inicial eliminada');
      } else {
        log.push('ℹ️ No se eliminó hoja inicial');
      }

  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet()
      .toast(`❌ Error: ${error.message}`, 'ERROR', 5);
  }
}


/**
 * FILTRAR HOJAS POR PATRÓN
 * Busca si el nombre contiene cualquiera de los textos
 */
function filtrarHojasPorPatron(hojas, log) {
  const hojasEncontradas = [];

  for (const hoja of hojas) {
    const nombre = hoja.getName().toUpperCase();

    for (const patron of CONFIG.patrones) {
      const patronUpper = patron.toUpperCase();

      if (nombre.includes(patronUpper)) {
        hojasEncontradas.push(hoja);
        log.push(`✅ ${hoja.getName()}`);
        break;
      }
    }
  }

  return hojasEncontradas;
}


/**
 * GENERAR NOMBRE ÚNICO
 */
function generarNombreUnico(destino, nombreBase) {
  const hojas = destino.getSheets().map(h => h.getName());

  if (!hojas.includes(nombreBase)) {
    return nombreBase;
  }

  let contador = 1;
  let nuevoNombre = `${nombreBase} (${contador})`;

  while (hojas.includes(nuevoNombre)) {
    contador++;
    nuevoNombre = `${nombreBase} (${contador})`;
  }

  return nuevoNombre;
}


/**
 * CREAR LOG
 */
function mostrarLogCompleto(log, activo) {
  try {
    let hojaLog = activo.getSheetByName('📋 LOG_COPIA_HOJAS');

    if (hojaLog) {
      activo.deleteSheet(hojaLog);
    }

    hojaLog = activo.insertSheet('📋 LOG_COPIA_HOJAS');

    log.forEach((linea, i) => {
      hojaLog.getRange(i + 1, 1).setValue(linea);
    });

    hojaLog.setColumnWidth(1, 700);

  } catch (error) {
    Logger.log(error);
  }
}
