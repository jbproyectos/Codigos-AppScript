  //  Constantes Globales con MAYUSCULAS
  //  Variables y funciones con camelCase
  //  Archivo: P-PS-FT-005_Rev.4_SOLICITUD DE GASTOS 2026

const SSID = SpreadsheetApp.getActiveSpreadsheet().getId();
const TABLAS_SHEET = "SOLICITUD_NOMINA";
const PLANTILLA_NAME = `Plantilla_Tablas`;

const SSID_A2 = `18OWh6mXY0-o5MRCAZEMQrznQmMJURt9n5bnolfF8ERU`; // Archivo A2 PRUEBA
const ID_DIRECTORIO = `1NZBsJOLjnP6aojinaPUaLMnliDHYnNqVJKMYq8VhTJE`; // V0.2
const ID_MASTER_GASTOS = `178M33EaTbv6rT6CA2XkA_csJlMoBI9Ej3s1T_7hq0no`;

function tablaNomina() {
  const regla = SpreadsheetApp.newDataValidation()
    .requireCheckbox("TRUE", "FALSE")
    .build();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const PLANTILLA = ss.getSheetByName(PLANTILLA_NAME);
  const SHEET = ss.getSheetByName(TABLAS_SHEET);

  var totalRange = (SHEET.getRange(1001,11,1000,6).getValues().slice(1)).filter(fila =>
    fila[0] != "" && 
    fila[0] != null && 
    fila[0] != `USUARIO FINAL`);

  var nomInfo = totalRange.filter(fila => 
    fila[5] == `NOMINA SEMANAL` &&
    fila[2] != 0).map(fila => [fila[0], fila[2], fila[1]]);

  totalRange = totalRange.map(fila => [fila[0], fila[5], fila[2]]);

  var afiInfo = totalRange.filter(fila => 
    (fila[1] == `IMPUESTO SOBRE LA NOMINA` 
    || fila[1] == `AGUINALDOS`
    || fila[1] == `CREDITO INFONAVIT`
    || fila[1] == `CREDITO PERSONAL`
    ) &&
    fila[2] != 0);

  var bonoInfo = totalRange.filter(fila => 
    (fila[1] == `BONO LEALTAD` || 
    fila[1] == `BONO DESPENSA` || 
    fila[1] == `BONO TRANSPORTE`) &&
    fila[2] != 0);
     
  const veces = SHEET.getRange(`J2`).getValue();
  var empInfoCruda = totalRange.filter(fila => 
    fila[1] == `BONO GRATIFICACION` && veces > 0);
    // fila[1] == `BONO GRATIFICACION` &&
    // fila[2] != 0);
  var empInfo = empInfoCruda.flatMap(fila =>
    Array.from({ length: veces }, () => [...fila])
  );

  var kpiInfo = totalRange.filter(fila => 
    fila[1] == `BONO MENSUAL` &&
    fila[2] != 0);


  // Validacion para Bonos segun semana
  var diasBono = 22; // 22 Cuarto Viernes para Bonos Bienestar 🟢
  // var diasBono = 1; // 1 Primer Viernes (PRUEBA)
  (getFridayOfMonth(diasBono))?afiInfo = [...afiInfo, ...bonoInfo]:0;

  var diasBono = 8; // 8 Segundo Viernes para Bono KPIs 🔴
  // var diasBono = 1; // 1 Primer Viernes (PRUEBA)
  (getFridayOfMonth(diasBono))?afiInfo = [...afiInfo, ...kpiInfo]:0;

  var diasBono = 15; // 15 Tercer Viernes para Bono Empleado del Mes 🟡
  // var diasBono = 1; // 1 Primer Viernes (PRUEBA)
  (getFridayOfMonth(diasBono))?afiInfo = [...afiInfo, ...empInfo]:0;

  //  afiInfo contiene todos los datos que se muestran en la segunda tabla.
  //  V2 no se muestran datos de Aguinaldos ni Impuesto sobre la Nomina
  //  V3 SI se muestran datos de Aguinaldos e Impuesto sobre la Nomina, asi como los Creditos correspondientes
  //  Incluir tabla aparte de Horas Extras

  const plantEncabezados2 = PLANTILLA.getRange(1,1,7,16);
  const plantNomFormatRange = PLANTILLA.getRange(8,1,nomInfo.length,5);
  const avisoRange = PLANTILLA.getRange(`M34:P39`);
  avisoRange.copyFormatToRange(SHEET, 13, 16, 39, 44);
  SHEET.getRange(`M39`).setValue(avisoRange.getValue());
  plantEncabezados2.copyFormatToRange(SHEET, 1, 11, 6, 12);
  SHEET.getRange(`A6:P12`).setValues(plantEncabezados2.getValues());
  plantNomFormatRange.copyFormatToRange(SHEET, 1, 5, 13, nomInfo.length+12);

  SHEET.getRange(`A7`).setFormula(PLANTILLA.getRange("A2").getFormula());
  SHEET.getRange(`G7`).setFormula(PLANTILLA.getRange("G2").getFormula());  
  SHEET.getRange(13,5,nomInfo.length,1).setFormulas(PLANTILLA.getRange(8,5,nomInfo.length,1).getFormulas());
  SHEET.getRange(13,1,nomInfo.length,3).setValues(nomInfo);

  if(afiInfo.length>0){
    const plantAllFormatRange = PLANTILLA.getRange(8,7,afiInfo.length,5);
    plantAllFormatRange.copyFormatToRange(SHEET, 7, 11, 13, afiInfo.length+12);
    SHEET.getRange(13,11,afiInfo.length,1).setFormulas(PLANTILLA.getRange(8,11,afiInfo.length,1).getFormulas());
    try {
      const offset = afiInfo.filter(
        fila => fila[1] === 'AGUINALDOS' || fila[1] === 'IMPUESTO SOBRE LA NOMINA' || fila[1] === 'CREDITO INFONAVIT' || fila[1] === 'CREDITO PERSONAL'
      ).length;
      const startRow = 13 + offset;
      const numRows = afiInfo.filter(
        fila => fila[1] != 'AGUINALDOS' && fila[1] != 'IMPUESTO SOBRE LA NOMINA' && fila[1] != 'CREDITO INFONAVIT' && fila[1] != 'CREDITO PERSONAL'
      ).length;
      if (numRows > 0) {
        SHEET
          .getRange(startRow, 10, numRows, 1)
          .setDataValidation(regla);
      }
    } catch (err) {
      Logger.log('Error al aplicar validación: ' + err.message);
      SpreadsheetApp.getActiveSpreadsheet().toast('Error al aplicar validación: ' + err.message)
    }
    SHEET.getRange(13,7,afiInfo.length,3).setValues(afiInfo);
    if (SHEET.getRange(afiInfo.length+12,8,1,1).getValue()==`BONO GRATIFICACION`){
      SHEET.getRange(afiInfo.length+12-(veces-1),7,veces,1)
        .setDataValidation(SHEET.getRange(`K1402`).getDataValidation());
      SHEET.getRange(afiInfo.length+12-(veces-1),9,veces,1).setBackground(`white`);
    }
    return
  }
  SHEET.getRange(`G10:K12`).clear();
}

//////////////////////////////

// Función actualizada para cuarto viernes del mes (day = 5)
function getFridayOfMonth(diasBono) {
// function getFridayOfMonth() {
  // var diasBono = 15;
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABLAS_SHEET);
  if(hoja.getRange(`R5`).getValue()){
    var pruebaFecha = hoja.getRange(`R2`).getValue();
    var year = (new Date(pruebaFecha)).getFullYear();
    var month = (new Date(pruebaFecha)).getMonth();
    var date = (new Date(pruebaFecha)).getDate();
    const fF = new Date(year, month, diasBono); // 0-Primera, 7-Segunda, 14-Tercera, 21-Cuarta
    const thisFriday = new Date(year, month, date+2);
    while (fF.getDay() !== 5) fF.setDate(fF.getDate() + 1); // 5 = viernes
    // console.log(`${getWeekNumber(fF)} - ${getWeekNumber(thisFriday)}`);
    return getWeekNumber(fF)==getWeekNumber(thisFriday);
  }

  var year = (new Date()).getFullYear();
  var month = (new Date()).getMonth();
  var date = (new Date()).getDate();
  const fF = new Date(year, month, diasBono); // 0-Primera, 7-Segunda, 14-Tercera, 21-Cuarta
  const thisFriday = new Date(year, month, date+2);
  while (fF.getDay() !== 5) fF.setDate(fF.getDate() + 1); // 5 = viernes
  // console.log(getWeekNumber(fF)==getWeekNumber(thisFriday));
  return getWeekNumber(fF)==getWeekNumber(thisFriday);
}

//////////////////////////////

function resetAfi(){
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABLAS_SHEET);
  const nombres = hoja.getRange("K1602:K1700").getValues().flat().filter(String);
  var values = Array(nombres.length).fill([0]);
  hoja.getRange(1602, 13, nombres.length, 1).setValues(values);
  hoja.getRange(1702, 13, nombres.length, 1).setValues(values);
  hoja.getRange(1802, 13, nombres.length, 1).setValues(values);
  hoja.getRange(1902, 13, nombres.length, 1).setValues(values);
}

//////////////////////////////

function getWeekNumber(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 5; // Viernes = 5
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  const weekNum = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  return weekNum;
}

//////////////////////////////

function probarSemana() {
  var fecha = new Date("2025-09-19"); // ejemplo
  var numSemana = getWeekNumber(fecha);
  Logger.log("Semana ISO: " + numSemana);
}

//////////////////////////////

function delTable() {
  SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(TABLAS_SHEET)
    .getRange("A4:P1000")
    .clearContent()
    .clearFormat()
    .clearDataValidations();
}

//////////////////////////////

function dataValidation(){
  const regla = SpreadsheetApp.newDataValidation()
    .requireCheckbox("TRUE", "FALSE")
    .build();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PLANTILLA_NAME);
    var range = sheet.getRange("J8");
    range.setDataValidation(regla);
}

//////////////////////////////

function permisos(){
  SpreadsheetApp.getActiveSpreadsheet();
  DriveApp.getRootFolder();
  console.log("Permisos verificados correctamente");
}

//////////////////////////////

function mandarSolicitudBoton(){
  if(mandarSolicitud()){
    rebajes();
    delTable();
    return
  }
  
}

//////////////////////////////

function mandarSolicitud() {
  var nominaHoja = SpreadsheetApp.openById(SSID).getSheetByName(TABLAS_SHEET);
  var nominaSupCompleta = nominaHoja.getRange(12,1,200,16).getValues();
  if (nominaHoja.getRange(`A13`).getValue() == `` || nominaHoja.getRange(`A13`).getValue() == null){
    SpreadsheetApp.getActiveSpreadsheet().toast(`🔴NO HAY DATOS POR MANDAR.🔴`);
    return false
  }
  if ((nominaHoja.getRange(`G13:G300`).getValues().filter(fila => fila[0] === `NOMBRE`)).flat().length >= 1 ){
    SpreadsheetApp.getActiveSpreadsheet().toast(`🟡SELECCIONA UN EMPLEADO DEL MES.🟡`);
    return false
  }
  if (nominaHoja.getRange(`L2`).getValue() < 0){
    SpreadsheetApp.getActiveSpreadsheet().toast(`🔴SOBREPASÓ EL MONTO MÁXIMO DEL BONO GRATIFICACIÓN.🔴`);
    return false
  }
  var fechaDinamica = new Date();

  (nominaHoja.getRange(`R5`).getValue())?fechaDinamica = nominaHoja.getRange(`R2`).getValue():0;

  var nominaSemanal = (nominaSupCompleta).filter(fila => 
    fila[0] !== "" && fila[0] !== null && fila[3] !== `HORAS EXTRAS`)
    .map(fila => 
      [fila[0], fila[1], fila[2], `NOMINA`, fila[4]] 
    ).slice(1).filter(fila => 
    fila[0] !== "" && fila[0] !== null && fila[3] !== `HORAS EXTRAS`);
  
  var horasExtra = (nominaSupCompleta).slice(1).filter(fila =>   //  Cambiar el intervalo a la tercer tabla
    fila[12] !== "" && fila[12] !== null && fila[12] !== `AL DARLE CLICK A "MANDAR SOLICITUD" ESTA CONFIRMANDO QUE LOS DATOS SON CORRECTOS.`)
    .map(fila => 
      [fila[12], fila[13], fila[14], `HORAS EXTRAS`, fila[15]] 
    );

  var nominaBonos = (nominaSupCompleta)
    .map(fila => 
      [fila[6],fila[8],1,fila[7], fila[10]]
    ).slice(1).filter(fila => 
    fila[0] !== "" && fila[0] !== null && fila[4] !== 0);

  var solicitudSuperior = [...nominaSemanal, ...horasExtra, ...nominaBonos];

  var datos = nominaHoja.getRange(1002,1,95,29).getValues()
    .filter(fila => fila[10] != `` && fila[10] != null);
  var datosObj = personaObject(datos);
  var today = Utilities.formatDate(fechaDinamica,Session.getScriptTimeZone(), `dd/MM/yyyy`);

  // Logger.log(JSON.stringify(datosObj));
  // return;
  const a2Sheet = SpreadsheetApp.openById(SSID_A2).getSheetByName(`S.Gastos CICLICOS INTERNO PS A2`);
  var consecutivo = a2Sheet.getRange(`A6:A`).getValues();
  solicitudSuperior = solicitudSuperior.map(fila => {
    try {
      const nombre = fila[0];
  
      return [
        generarIdentificador(
          datosObj[nombre].AREA_APLICA,
          datosObj[nombre].CATEGORIA,
          fila[3],
          nombre,
          consecutivo
        ) || `SIN DATOS`,
        today,
        datosObj[nombre].QUIEN_SOL || `SIN DATOS`,
        datosObj[nombre].DPTO_SOL || `SIN DATOS`,
        datosObj[nombre].USO || `SIN DATOS`,
        ((fila[3]==`NOMINA`)?`SEMANAL`:(fila[3]==`HORAS EXTRAS`)?`ÚNICO`:`MENSUAL`) || `SIN DATOS`,
        datosObj[nombre].AREA_APLICA || `SIN DATOS`,
        datosObj[nombre].CATEGORIA || `SIN DATOS`,
        fila[3] || `SIN DATOS`,
        (fila[3] == `BONO GRATIFICACION`)
          ? `EMPLEADO DEL MES`
          : (datosObj[nombre].DETALLE || `SIN DATOS`),
        nombre || `SIN DATOS`,
        fila[2] || `SIN DATOS`,
        (fila[3]!=`CREDITO INFONAVIT`&&fila[3]!=`CREDITO PERSONAL`)?((fila[1]>0)?(-1*fila[1]):fila[1]):((fila[1]>0)?fila[1]:(-1*fila[1]))
          || `SIN DATOS`,
        `N/A`,
        `N/A`,
        (fila[3]!=`NOMINA`&&fila[3]!=`HORAS EXTRAS`)
          ? mesNomina(fila[3],today)
          : semanaDelMesNominaSemanal(fila[3],today) || `SIN DATOS`,
        `SERVICIO`,
        `TRANSFERENCIA`,
        `NACIONAL`,
        datosObj[nombre].EMPRESA,
        datosObj[nombre].DESTINO || `SIN DATOS`,
        datosObj[nombre].CUENTA_CLABE || `SIN DATOS`,
        datosObj[nombre].TITULAR || `SIN DATOS`,
        (fila[3]!=`CREDITO INFONAVIT`&&fila[3]!=`CREDITO PERSONAL`)?((fila[4]>0)?(-1*fila[4]):fila[4]):((fila[4]>0)?fila[4]:(-1*fila[4]))
          || `SIN DATOS`,
        `AZAEL_RANGEL`,
        `N/A`,
        `SIN TICKET`,
        `N/A`
      ];
  
    } catch (error) {
      Logger.log(`❌ Error con nombre: "${fila[0]}"`);
      SpreadsheetApp.getUi().alert(`❌ Error con nombre: "${fila[0]}"`);
      Logger.log(`Fila completa: ${JSON.stringify(fila)}`);
      Logger.log(`Mensaje: ${error.message}`);
  
      // 👉 opcional: regresar fila "marcada"
      return ["ERROR", ...fila];
    }
  });

  var sheet13 = SpreadsheetApp.openById(SSID_A2).getSheetByName(`S.Gastos CICLICOS INTERNO PS A2`);
  var lastRow13 = (sheet13.getRange(`C1:C`).getValues().filter(fila => fila[0]!="").flat()).length+3;

  // Logger.log(`Arreglo concat:
  // ${solicitudSuperior}`);

    //  Insertar datos en A2
  sheet13.getRange(lastRow13,1,solicitudSuperior.length,solicitudSuperior[0].length).setValues(solicitudSuperior);
  // return false
  return true
}

//////////////////////////////

function semanaDelMesNominaSemanal(subcatego,fechaStr) {
  const [dia, mes, anio] = fechaStr.split('/').map(n => parseInt(n, 10));
  const fecha = new Date(anio, mes - 1, dia);
  const PRIMER_DIA_SEMANA = 5;
  function obtenerViernesSemana(fecha) {
    const d = new Date(fecha);
    const day = d.getDay();
    const diff = (PRIMER_DIA_SEMANA - day + 7) % 7;
    d.setDate(d.getDate() + diff);
    return d;
  }
  const viernesSemana = obtenerViernesSemana(fecha);
  const mesViernes = viernesSemana.getMonth();
  const anioViernes = viernesSemana.getFullYear();
  const inicioMes = new Date(anioViernes, mesViernes, 1);
  const offset = (inicioMes.getDay() - PRIMER_DIA_SEMANA + 7) % 7;
  const numeroDia = viernesSemana.getDate();
  const semana = Math.floor((numeroDia + offset - 1) / 7);
  var numSemana = ``;
  switch (semana){
    case 1: numSemana = `1RA`; break;
    case 2: numSemana = `2DA`; break;
    case 3: numSemana = `3RA`; break;
    case 4: numSemana = `4TA`; break;
    default: numSemana = `5TA`; break;
  }
  const meses = [
    "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
    "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
  ];

  var mesNombre = meses[mesViernes] || ``;
  return (`${subcatego} ${numSemana} SEMANA ${mesNombre} ${anioViernes}`);
}

//////////////////////////////

function mesNomina(subcatego,fechaStr) {
  const [dia, mes, anio] = fechaStr.split('/').map(n => parseInt(n, 10));
  const fecha = new Date(anio, mes - 1, dia);
  const PRIMER_DIA_SEMANA = 5;
  function obtenerViernesSemana(fecha) {
    const d = new Date(fecha);
    const day = d.getDay();
    const diff = (PRIMER_DIA_SEMANA - day + 7) % 7;
    d.setDate(d.getDate() + diff);
    return d;
  }
  const viernesSemana = obtenerViernesSemana(fecha);
  const mesViernes = viernesSemana.getMonth();
  const anioViernes = viernesSemana.getFullYear();
  const meses = [
    "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
    "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
  ];

  var mesNombre = meses[mesViernes] || ``;
  return (`${subcatego} ${mesNombre} ${anioViernes}`);
}

//////////////////////////////

function rebajes(){
  const nominaHoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABLAS_SHEET);
  var rebajesSemana = (nominaHoja.getRange("M3").getValue())*1;
  var rebajesTotalViejo = nominaHoja.getRange("O3").getValue()*1;
  var rebajesTotalNuevo = (rebajesTotalViejo+rebajesSemana)*1;
  // Logger.log(rebajesTotalNuevo);
  nominaHoja.getRange(`O3`).setValue(rebajesTotalNuevo);
}

//////////////////////////////

function generarIdentificador(area,categoria,subcatego,nombre,consecutivo){
   var archivo;
   (categoria == `NOMINAS`)?archivo = `A2`:0; // SWITCH para obtener que archivo es con la categoria
   const directorioSheet = SpreadsheetApp.openById(ID_DIRECTORIO).getSheetByName(`RESTRUCTURACION`);
   const masterGastosSheet = SpreadsheetApp.openById(ID_MASTER_GASTOS).getSheetByName(`DIR-CAT-SUBCAT`);
   const directorioArreglo = directorioSheet.getRange(`B1:C`).getValues()
    .filter(fila => fila[1] != `` && fila[1] != null).map(fila => [fila[1], fila[0]]);
   const masterGArreglo = masterGastosSheet.getRange(`D1:G`).getValues()
    .filter(fila => fila[1] != `` && fila[1] != null).map(fila => [fila[0], fila[3]]);
   const areasArreglo = masterGastosSheet.getRange(`P1:Q`).getValues()
    .filter(fila => fila[1] != `` && fila[1] != null);
  const directorioObjeto = arrayToObject(directorioArreglo);
  const masterGObjeto = arrayToObject(masterGArreglo);
  const areasObjeto = arrayToObject(areasArreglo);
  var numEmpleado = cerosAntes(directorioObjeto[nombre]);
  (subcatego==`BONO GRATIFICACION`)?subcatego=`BONO GRATIFICACION`:0;
  var numSubcatego = masterGObjeto[subcatego];
  var numArea = areasObjeto[area];

   consecutivo = consecutivo.filter(fila => typeof fila[0] === `string` && fila[0].startsWith(`${numArea}-${archivo}-${numEmpleado}-${numSubcatego}`)).flat().length+1;
    const folio = seisCerosAntes(consecutivo);
  return validarDatos(archivo, numArea, numEmpleado, numSubcatego, folio);
}

//////////////////////////////

function arrayToObject(data) {
  return data.reduce((acc, row) => {
    const clave = row[0];
    const valor = row[1];
    acc[clave] = valor;
    return acc;
  }, {});
}

//////////////////////////////

function personaObject(data) {
  return data.reduce((acc, row) => {
    const nombre = row[10];
    acc[nombre] = {
      QUIEN_SOL: row[2],
      DPTO_SOL: row[3],
      USO: row[4],
      AREA_APLICA: row[6],
      CATEGORIA: row[7],
      DETALLE: row[9],
      EMPRESA: row[19],
      DESTINO: row[20],
      CUENTA_CLABE: row[21],
      TITULAR: row[22]
    };
    return acc;
  }, {});
}

//////////////////////////////

function cerosAntes(numero){
  try{
    numeroStr = JSON.stringify(numero);
    switch (numeroStr.length){
    case 0: numeroStr = `0000`+numeroStr; break;
    case 1: numeroStr = `000`+numeroStr; break;
    case 2: numeroStr = `00`+numeroStr; break;
    case 3: numeroStr = `0`+numeroStr; break;
    case 4: numeroStr = numeroStr; break;
    default: numeroStr = `0000`;
    }
  }catch(err){
    return undefined;
  }
  return numeroStr;
}

//////////////////////////////

function seisCerosAntes(numero){
    numeroStr = JSON.stringify(numero);
    switch (numeroStr.length){
    case 0: numeroStr = `000000`+numeroStr; break;
    case 1: numeroStr = `00000`+numeroStr; break;
    case 2: numeroStr = `0000`+numeroStr; break;
    case 3: numeroStr = `000`+numeroStr; break;
    case 4: numeroStr = `00`+numeroStr; break;
    case 5: numeroStr = `0`+numeroStr; break;
    case 6: numeroStr = numeroStr; break;
    default: numeroStr = `000000`;
  }
  return numeroStr;
}


function validarDatos(archivo, numArea, numEmpleado, numSubcatego, folio) {
  const variables = {
    archivo,
    numArea,
    numEmpleado,
    numSubcatego,
    folio
  };
  const faltantes = Object.entries(variables)
    .filter(([_, valor]) => valor === undefined)
    .map(([nombre]) => nombre);
  if (faltantes.length > 0) {
    return `Identificador inválido. Falta: ${faltantes.join(', ')}`;
  }
  return`${numArea}-${archivo}-${numEmpleado}-${numSubcatego}-${folio}`;
}

//////////////////////////////

function horasXtra() {
  const ss = SpreadsheetApp.openById(SSID);
  const SHEET = ss.getSheetByName(TABLAS_SHEET);
  const PLANTILLA = ss.getSheetByName(PLANTILLA_NAME);
  const lastRow = SHEET.getRange(`M12`).getDataRegion().getLastRow()+1;
  const formula = 
  `=IF(O${lastRow}*1>8,
  IF(O${lastRow}*1>16,(N${lastRow}*8)+(N${lastRow}*16)+(N${lastRow}*(O${lastRow}-16)*3),(N${lastRow}*8)+(N${lastRow}*(O${lastRow}-8)*2)),
N${lastRow}*O${lastRow})`;
  const precioUnitFormula = 
    `=IFERROR(VLOOKUP($M${lastRow}, $K$1002:$M$1100, 3, FALSE)/8, "")`;

  PLANTILLA.getRange(8,13,1,4).copyFormatToRange(SHEET,13,16,lastRow,lastRow);
  SHEET.getRange(`M${lastRow}`)
      .setDataValidation(SHEET.getRange(`K1402`).getDataValidation());
  SHEET.getRange(lastRow,13,1,4).setValues([[`NOMBRE`,precioUnitFormula,0,formula]]);
}

//////////////////////////////

function arrToObject(data) {
  let obj = {};
  data.forEach(fila => {
    obj[fila[0]] = fila[1];
  });
  return obj;
}

function comsiones(data) {
  const ss = SpreadsheetApp.openById(SSID);
  const SHEET = ss.getSheetByName(TABLAS_SHEET);
  const PLANTILLA = ss.getSheetByName(PLANTILLA_NAME);
  const lastRow = SHEET.getRange(`G12`).getDataRegion().getLastRow()+1;
  var arrNom = SHEET.getRange(`K1002:M1100`).getValues().filter(fila => fila[0] != "" && fila[0] != null)
    .map(fila => [fila[0], fila[2]]);
  var objNom = arrToObject(arrNom);
  var puNom = objNom[data.name];    //  Precio Unitario Nominas (Obtener dependiendo del data.name)
    // data.name
    // data.horas
  var formula = 
  `=IF(O${lastRow}*1>8,
  IF(O${lastRow}*1>16,(N${lastRow}*8)+(N${lastRow}*16)+(N${lastRow}*(O${lastRow}-16)*3),(N${lastRow}*8)+(N${lastRow}*(O${lastRow}-8)*2)),
N${lastRow}*O${lastRow})`;

  var arrXtra = [[
    data.name,
    (puNom/8),   //  (Precio Unitario Nomina Semanal)/8 (?)
    data.horas,
    formula
  ]];

  PLANTILLA.getRange(8,13,1,4).copyFormatToRange(SHEET,13,16,lastRow,lastRow);
  // SHEET.getRange(lastRow,13,1,1).setValues(PLANTILLA.getRange(8,13,1,1).getValues());
  SHEET.getRange(lastRow,13,1,4).setValues(arrXtra);
}

//////////////////////////////

function arrToObject(data) {
  let obj = {};
  data.forEach(fila => {
    obj[fila[0]] = fila[1];
  });
  return obj;
}

//////////////////////////////

function mandarKPIs() {
  const nomSS = SpreadsheetApp.getActiveSpreadsheet();
  const tablaKPI = nomSS.getActiveSheet().getRange("P4").getDataRegion();
  const tablaKPIVal = tablaKPI.getValues();
  const sheetNom = nomSS.getSheetByName("SOLICITUD_NOMINA");
  const tablaNom = sheetNom.getRange("K1502").getDataRegion();
  const rows = tablaNom.getNumRows();
  const tablaNomVal = (tablaNom.getValues()).map(fila => fila[10]);
  tablaNomVal.shift();
  tablaKPIVal.shift();
  const mapaKPI = new Map(tablaKPIVal.map(fila => [fila[0], fila]));
  const tablaKPIOrdenada = tablaNomVal.map(nombre => mapaKPI.get(nombre) || [nombre, "", "", ""]);
  const tablaFinal = tablaKPIOrdenada.map(fila => [fila[0], 1, fila[2]]);
  sheetNom.getRange(1502, 11, rows-1, 3).setValues(tablaFinal);
}

//////////////////////////////

function mandarSolicitudCheck(){
  if(SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(`SOLICITUD_NOMINA`)
    .getRange(`N1`)
    .getValue()==true){
      mandarSolicitudBoton();
    }else{
      Logger.log(`No se mandó información.`);
    }
}
