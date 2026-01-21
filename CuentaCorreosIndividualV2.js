// CORREO PROPIETARIO "correo@dominio.com" - NOMBRE APELLIDO v2
const CORREO = `correo@dominio.com`
const SSID = SpreadsheetApp.getActiveSpreadsheet().getId();

function onOpen() {
  var menu = SpreadsheetApp.getUi().createMenu("📑 Reporte Correo");
  menu.addSeparator();
  menu.addItem("⏰| Crear Activador ", "creaActivador");
  menu.addSeparator();
  menu.addItem("🔄| Actualizar datos ", "CopiaDiasLaborales");
  menu.addSeparator();
  menu.addToUi();
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
}

//////////////////////////////

function getData() {

  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB");
  correo = CORREO + ">";
  const nombre = CorreosKabzo.CORREOS_KABZO[correo].NOMBRE || "INVALIDO";
  if (nombre === "INVALIDO") return;

  // Fechas y arrays precalculados
  const hoy = new Date();
  const fechaHace3Meses = new Date(hoy);
  fechaHace3Meses.setMonth(fechaHace3Meses.getMonth() - 2);
  const zona = "GMT-6";
  const formatoFecha = "yyyy-M-dd";
  const buscafecha = Utilities.formatDate(hoy, zona, formatoFecha);
  const before = Utilities.formatDate(hoy, zona, "yyyy-12-31");
  const query = `after:${buscafecha} before:${before} in:anywhere`;

  // Horarios para clasificar
  const maniana = ["12 AM","1 AM","2 AM","3 AM","4 AM","5 AM","6 AM","7 AM","8 AM","9 AM","10 AM","11 AM"];
  const mediodia = ["12 PM","1 PM","2 PM","3 PM"];
  const tarde = ["4 PM","5 PM","6 PM","7 PM"];
  const fuera = ["8 PM","9 PM","10 PM","11 PM"];

  hoja.getRange("A1:H1")
    .setBackgroundColor("#46bdc6")
    .setFontColor("#ffffff")
    .setValues([["REMITENTE","ASUNTO","MENSAJE","FECHA","HORA","GRUPO","COLABORADOR","TIEMPO"]]);

  const cadenas = GmailApp.search(query);
  Logger.log(cadenas.length);
  const datos = [];

  cadenas.forEach(cadena => {
    const mensajes = cadena.getMessages();
    const correoMsg = mensajes[0];
    var arreglos = [...CorreosKabzo.ARR_KABZO, ...CorreosKabzo.ARR_PROMOTORES];
    const remitente = correoMsg.getFrom().split("<")[1];
    if (!arreglos.includes(remitente)) return; // Validar remitente permitido
    var directorios = Object.assign({}, CorreosKabzo.DIR_KABZO, CorreosKabzo.DIR_PROMOTORES);

    const asunto = cadena.getFirstMessageSubject();
    const cuerpo = correoMsg.getPlainBody();
    const fechaCorreo = correoMsg.getDate();
    const fechaStr = Utilities.formatDate(fechaCorreo, zona, formatoFecha);

    const horaoriginal = Utilities.formatDate(fechaCorreo, zona, "h a");

    // Clasificar horario
    let tiempo = "";
    if (maniana.includes(horaoriginal)) tiempo = "MAÑANA";
    else if (mediodia.includes(horaoriginal)) tiempo = "MEDIO DIA";
    else if (tarde.includes(horaoriginal)) tiempo = "TARDE";
    else if (fuera.includes(horaoriginal)) tiempo = "FUERA DE HORARIO";

    // Semana del año
    const weekNum = getWeekNumber(fechaCorreo);

    // Texto personalizado (mes-num y abreviado)
    const mesNum = fechaCorreo.getMonth() + 1; // 0-11
    const mesAbr = fechaCorreo.toLocaleString("es-MX", { month: "short" }).toUpperCase();
    const anio = fechaCorreo.getFullYear();
    const mesTexto = `${mesNum} - ${mesAbr} ${anio}`;

    datos.push([
      remitente,
      asunto,
      cuerpo,
      fechaStr,
      horaoriginal,
      directorios[remitente],  // Nombre de los directorios (remitente)
      nombre, // Nombre del que recibe los correos
      tiempo,
      weekNum,
      mesTexto,
      CorreosKabzo.CORREOS_KABZO[correo].PUESTO
    ]);

    GmailApp.markMessagesRead(mensajes);
  });

  Logger.log(datos.length);
  if (datos.length) {
    Logger.log(datos.length);
    hoja.getRange(hoja.getLastRow() + 1, 1, datos.length, datos[0].length).setValues(datos);
    hoja.getRange("J1").setValue(new Date());
  }
}

//////////////////////////////

function getWeekNumber(date) {
  const tempDate = new Date(date.getTime());
  tempDate.setHours(0, 0, 0, 0);
  tempDate.setDate(tempDate.getDate() + 4 - (tempDate.getDay() || 7)); // jueves de esa semana
  const yearStart = new Date(tempDate.getFullYear(), 0, 1);
  const weekNo = Math.ceil((((tempDate - yearStart) / 86400000) + 1) / 7);
  return weekNo;
}

//////////////////////////////

function DelTable() {
  const hoja = SpreadsheetApp.openById(SSID).getSheets()[0]
  hoja.getRange("A2:K").clearContent();
  hoja.getRange("J1").setValue(new Date());
}

//////////////////////////////

function CopiaDiasLaborales() {
  const dia = new Date();
  const numeroDia = new Date(dia).getDay();
  if(numeroDia != 6 && numeroDia != 0) getData();
}
