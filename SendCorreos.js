/**
 * Función principal que se ejecutará por el trigger temporal - modificable para agregar mas correos a leer
 */
function obtenerCorreosDeTodasLasCuentas() {
  const cuentas = [
    'desarrolladorjr@kabzo.org',
    'desarrollo@kabzo.org',
    'optimizacion@kabzo.org',
    'projectmanager@kabzo.org',
    'sistemas3@kabzo.org'
  ];
  
  cuentas.forEach(cuenta => {
    try {
      Logger.log(`\n====== INICIANDO PROCESAMIENTO PARA: ${cuenta} ======`);
      obtenerCorreosDeCuenta(cuenta);
      Utilities.sleep(2000); // Espera entre cuentas
    } catch (e) {
      Logger.log(`ERROR GLOBAL procesando cuenta ${cuenta}: ${e.toString()}`);
      MailApp.sendEmail({
        to: 'desarrollo@kabzo.org',
        subject: `Error grave en cuenta ${cuenta}`,
        body: `Error: ${e.toString()}\n\nStack Trace:\n${e.stack}`
      });
    }
  });
  Logger.log("====== PROCESAMIENTO COMPLETADO ======");
}

/**
 * Procesa los correos de una cuenta específica
 */
function obtenerCorreosDeCuenta(cuenta) {
  const startTime = new Date();
  Logger.log(`Inicio de procesamiento para ${cuenta} a las ${startTime}`);
  
  const url = "http://38.65.143.27/correos/SendMail.php";
  const hoy = new Date();
  const fechaHoy = Utilities.formatDate(hoy, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  
  // Query para buscar correos del día actual (sin marcar como leídos)
  const query = `to:${cuenta} after:${fechaHoy} before:${fechaHoy.replace(/\d{2}$/, (d) => ('0' + (+d + 1)).slice(-2))} -from:noreply-apps-scripts-notifications@google.com`;
  
  try {
    // Obtener hilos sin afectar estado de lectura
    const threads = GmailApp.search(query, 0, 500);
    Logger.log(`Hilos encontrados para ${cuenta}: ${threads.length}`);
    
    // Obtener mensajes sin marcarlos como leídos
    const mensajes = [];
    threads.forEach(thread => {
      const msgs = thread.getMessages();
      msgs.forEach(msg => {
        // Verificar si el correo es relevante para la cuenta
        if (esCorreoRelevante(msg, cuenta)) {
          mensajes.push(msg);
        }
      });
    });
    
    Logger.log(`Mensajes a procesar para ${cuenta}: ${mensajes.length}`);
    
    // Procesar por lotes
    const BATCH_SIZE = 10;
    let resultados = {
      nuevos: 0,
      duplicados: 0,
      errores: 0
    };
    
    for (let i = 0; i < mensajes.length; i += BATCH_SIZE) {
      const batch = mensajes.slice(i, i + BATCH_SIZE);
      const resultado = procesarBatch(batch, cuenta, url);
      resultados.nuevos += resultado.nuevos;
      resultados.duplicados += resultado.duplicados;
      resultados.errores += resultado.errores;
      Utilities.sleep(1000); // Pausa entre lotes
    }
    
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    
    Logger.log(`Resumen para ${cuenta}:
      - Nuevos registros: ${resultados.nuevos}
      - Duplicados: ${resultados.duplicados}
      - Errores: ${resultados.errores}
      - Tiempo total: ${duration} segundos`);
    
    // Notificar solo si hay errores reales
    if (resultados.errores > 0) {
      MailApp.sendEmail({
        to: 'desarrollo@kabzo.org',
        subject: `Errores en procesamiento ${cuenta}`,
        body: `Se encontraron ${resultados.errores} errores al procesar ${cuenta}`
      });
    }
    
  } catch (e) {
    Logger.log(`Error al procesar cuenta ${cuenta}: ${e.toString()}`);
    throw e;
  }
}

/**
 * Determina si un correo es relevante para la cuenta
 */
function esCorreoRelevante(mensaje, cuenta) {
  try {
    // Verificar si la cuenta es destinataria directa
    const destinatarios = mensaje.getTo();
    if (destinatarios && destinatarios.includes(cuenta)) {
      return true;
    }
    
    // Opcional: verificar también en CC si lo deseas
    const cc = mensaje.getCc();
    if (cc && cc.includes(cuenta)) {
      return true;
    }
    
    return false;
  } catch (e) {
    Logger.log(`Error al verificar relevancia: ${e.toString()}`);
    return false;
  }
}

/**
 * Procesa un lote de mensajes
 */
function procesarBatch(batch, cuenta, url) {
  let resultado = {
    nuevos: 0,
    duplicados: 0,
    errores: 0
  };
  
  batch.forEach(mensaje => {
    try {
      // Obtener datos sin marcar como leído
      const datosCorreo = {
        message_id: mensaje.getId(),
        remitente: mensaje.getFrom(),
        asunto: mensaje.getSubject(),
        mensaje: extraerCuerpoSinMarcar(mensaje),
        fecha: Utilities.formatDate(mensaje.getDate(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        hora: Utilities.formatDate(mensaje.getDate(), Session.getScriptTimeZone(), 'HH:mm:ss'),
        dequecorreoviene: cuenta
      };
      
      Logger.log(`Procesando mensaje ID: ${datosCorreo.message_id}`);
      Logger.log(`Asunto: ${datosCorreo.asunto}`);
      
      const opciones = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(datosCorreo),
        muteHttpExceptions: true,
        timeout: 30000
      };
      
      const respuesta = UrlFetchApp.fetch(url, opciones);
      const responseCode = respuesta.getResponseCode();
      const contentText = respuesta.getContentText().trim();
      
      // Interpretar respuesta según tu PHP
      if (responseCode === 200) {
        const respuestaJson = JSON.parse(contentText);
        
        if (respuestaJson.status === 'success') {
          resultado.nuevos++;
          Logger.log(`Nuevo registro exitoso para ID: ${datosCorreo.message_id}`);
        } 
        else if (respuestaJson.status === 'duplicate') {
          resultado.duplicados++;
          Logger.log(`Correo duplicado (ID existente): ${datosCorreo.message_id}`);
        }
        else {
          throw new Error(`Respuesta inesperada: ${contentText}`);
        }
      } else {
        throw new Error(`Error HTTP ${responseCode}: ${contentText}`);
      }
      
    } catch (e) {
      resultado.errores++;
      Logger.log(`ERROR al procesar mensaje: ${e.toString()}`);
    }
  });
  
  return resultado;
}

/**
 * Extrae el cuerpo del mensaje sin marcarlo como leído
 */
function extraerCuerpoSinMarcar(mensaje) {
  try {
    // Usar getRawContent() para evitar marcar como leído
    const rawContent = mensaje.getRawContent();
    
    // Implementación simple para extraer texto plano
    const textParts = rawContent.split(/\r?\n/);
    let inBody = false;
    let body = '';
    
    for (const line of textParts) {
      if (line.startsWith('Content-Type: text/plain')) {
        inBody = true;
        continue;
      }
      if (inBody && line === '') {
        inBody = false;
        break;
      }
      if (inBody) {
        body += line + '\n';
      }
    }
    
    return body.trim().substring(0, 10000); // Limitar tamaño
  } catch (e) {
    Logger.log(`Error al extraer cuerpo: ${e.toString()}`);
    return "No se pudo extraer el contenido del mensaje";
  }
}
