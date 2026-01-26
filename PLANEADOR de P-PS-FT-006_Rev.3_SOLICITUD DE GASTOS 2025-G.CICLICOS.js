function ciclicosBoton() {
  try {
    // Optimizado: A1 estaba duplicado — ya no.
    var hojasDatos = [
      //prueba 19/12/2025
      /*{ link: "1XO-Ddfi7omPYlejbFiCIfrWdmJJrYKR5YMc4Vf9NC90", destino: "S.Gastos CICLICOS INTERNO PS A6", origen: "Base de Datos Despacho" },
      { link: "1OJsP9XyP6T4FxDG90Tbd71_9Wucx9U6b2sI1z4fHp0k", destino: "S.Gastos CICLICOS INTERNO PS A5", origen: "Base de Datos Despacho" },
      { link: "11MiUy67oMXIPwyb5kPV9E8-PQU9WZQqJiGm2HgoACCc", destino: "S.Gastos CICLICOS INTERNO PS A4", origen: "Base de Datos Despacho" },
      { link: "1ZkdKx6z5ibd6HEYwoa0oeKgziPNZ7itZFcYG9Uy55iU", destino: "S.Gastos CICLICOS INTERNO PS A3", origen: "Base de Datos Despacho" },
      { link: "108glr8-NjL3MnlhwEJgbEi8d6aCw5nZOWmL5h68NfbU", destino: "S.Gastos CICLICOS INTERNO PS A2", origen: "Base de Datos Despacho" },//modif 18/12/2025 nominas
      { link: "15orxBAT0fqlZ1R2f1w78PBPkS4IlfiLXRqubvh9M2lM", destino: "S.Gastos CICLICOS INTERNO PS A1", origen: "Base de Datos Despacho" },//modif 18/12/2025 que no sean nominas
      { link: "1ecfRYOMajfyuui4w5iAIwct0gJVUPg4MjbmGRfPpRT4", destino: "S.Gastos CICLICOS INTERNO PS A0", origen: "Base de Datos Despacho" },
      { link: "10ShS-qAuQvmfthOaN326EgbXQfKMB4mI9znPnwFGVag", destino: "S.Gastos Personales", origen: "Base de Datos Personal" }//agregar el 003 gastos personales.*/
      
      //originales
      { link: "1havjYfhnJ-Qe5DyDg0duLPAX7BN7veffhysscsG9jPc", destino: "S.Gastos CICLICOS INTERNO PS A6", origen: "Base de Datos Despacho" },
      { link: "1ngHul195CohXo7eFB6lvDOhgNxAP9pwOgKnt27th8UI", destino: "S.Gastos CICLICOS INTERNO PS A5", origen: "Base de Datos Despacho" },
      { link: "1lOQ7p4H4pfqpADV5pDQBFKLv-_Jaf9aS6OBjwi0zGos", destino: "S.Gastos CICLICOS INTERNO PS A4", origen: "Base de Datos Despacho" },
      { link: "1PaQdKfVk51UiMNnKVS-Jo0YiDSs3mU-76_zweQy1M6c", destino: "S.Gastos CICLICOS INTERNO PS A3", origen: "Base de Datos Despacho" },
      { link: "1CalZsgEqEhWPJGloUGBSZwUXz9uPaVK2VfwbvRQuTms", destino: "S.Gastos CICLICOS INTERNO PS A2", origen: "Base de Datos Despacho" },//modif 18/12/2025 nominas
      { link: "1HueYpVVHTSL6bJpBF8y_PEF-C5MGIt_o2wgPeYRJd7I", destino: "S.Gastos CICLICOS INTERNO PS A1", origen: "Base de Datos Despacho" },//modif 18/12/2025 que no sean nominas
      { link: "18S6lqUMLJ07QB4QEWvh6Gppb7PSjEWnXPbhY57sbZhM", destino: "S.Gastos CICLICOS INTERNO PS A0", origen: "Base de Datos Despacho" },
      { link: "19fionVnXuVOe2Ex5WthuF5te0YrZbPz-HN8YvV9XwuA", destino: "S.Gastos Personales", origen: "Base de Datos Personal"}//agregar el 003 gastos personales.
    ];

    hojasDatos.forEach(function (hoja) {
      try {
        envioInfoCiclico_rapidoV3(hoja);
      } catch (err) {
        Logger.log(`❌ Error procesando ${hoja.destino}: ${err.message}`);
      }
    });

  } catch (e) {
    Logger.log("❌ Error general en ciclicosBoton: " + e.message);
  }
}

function envioInfoCiclico_rapidoV3(hojaInfo) {

  var libroOrigen = SpreadsheetApp.getActiveSpreadsheet();
  var hojaOrigen = libroOrigen.getSheetByName(hojaInfo.origen);

  var ultimaFila = hojaOrigen.getLastRow();
  if (ultimaFila < 2) return; // No hay datos

  // Solo lee lo que existe
  //var datos = hojaOrigen.getRange(2, 2, ultimaFila - 1, 28).getValues();
  //fila 2, columna 1 = A, fila final, columna final
  var datos = hojaOrigen.getRange(2, 1, ultimaFila - 1, 29).getValues();

  // Abrir UNA VEZ archivo de destino
  var libroDestino = SpreadsheetApp.openById(hojaInfo.link);
  var hojaDestino = libroDestino.getSheetByName(hojaInfo.destino);

  var fechaHoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yy');

  // Filtro inteligente en una sola pasada
  var filas = datos.filter(function (fila) {

    // Fecha
    //var fecha = fila[0];
    var fecha = fila[1];
    if (!(fecha instanceof Date)) return false;
    //if (Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'dd/MM/yy') !== fechaHoy)
    if (Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'dd/MM/yy') !== fechaHoy)
      return false;

    fecha === Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm:ss')

    var persona = fila[2];
    var area = fila[4]; //uso
    var categoria = fila[7];

    switch (hojaInfo.destino) {
      case "S.Gastos CICLICOS INTERNO PS A0":
        return persona === "NATALIE_REYNA" && area === "DESPACHO";

      /*las demas categorias deben de caer aqui*/
      case "S.Gastos CICLICOS INTERNO PS A1":
        return persona === "VALERIA_VARGAS" && area === "DESPACHO" && categoria !== "NOMINAS";//que no sean nomina, que sean las demas categorias.

      /*solo nominas cairan en el A2 */
      case "S.Gastos CICLICOS INTERNO PS A2":
        return persona === "VALERIA_VARGAS" && area === "DESPACHO" && categoria === "NOMINAS"; //solo nominas

      case "S.Gastos CICLICOS INTERNO PS A3":
        return persona === "FATIMA_MARTINEZ" && area === "DESPACHO";

      case "S.Gastos CICLICOS INTERNO PS A4":
        return persona === "NADIA_ELIZONDO" && area === "DESPACHO";

      case "S.Gastos CICLICOS INTERNO PS A5":
        return persona === "NAYELI_LUNA" && area === "DESPACHO";

      case "S.Gastos CICLICOS INTERNO PS A6":
        return persona === "FRIDA_PIÑA" && area === "DESPACHO";
      
      case "S.Gastos Personales":
        return persona === "VALERIA_VARGAS" && area === "PERSONAL";

      default:
        return true; // fallback
    }
  });

  if (filas.length === 0) {
    Logger.log("ℹ️ No hay filas para " + hojaInfo.destino);
    return;
  }

  //var inicioPegado = hojaDestino.getLastRow() + 1;
  var inicioPegado = ultimaFilaNoVaciaV1(hojaDestino);

  //hojaDestino.getRange(inicioPegado + 1, 2, filas.length, 28)
  hojaDestino.getRange(inicioPegado + 1, 1, filas.length, 29)
    .setValues(filas);

  Logger.log(`✅ ${hojaInfo.destino}: Pegadas ${filas.length} filas.`);
}

function ultimaFilaNoVaciaV1(hoja) {
 // const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SOLICITUDES 2024");
  if (!hoja) {
    Logger.log("La hoja "+ hoja + " no existe.");
    return;
  }
  
  //const columna = hoja.getRange("A:A").getValues(); // Obtiene todos los valores de la columna B
  const columna = hoja.getRange("D:D").getValues(); // Obtiene todos los valores de la columna B
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


/*ANUAL 22/12/2025*/
function generarAnual() {//con la actualizacion 03/12/2025
  var libroOrigen = SpreadsheetApp.getActiveSpreadsheet();
  var hojaOrigen = libroOrigen.getSheetByName("S.Gastos CICLICOS INTERNO PS");
  //var hojaDestino = libroOrigen.getSheetByName("Planeador Despacho");

  //var hojaOrigen = libroOrigen.getSheetByName("Copia de S.Gastos CICLICOS INTERNO PS(Personal)");
  //var hojaDestino = libroOrigen.getSheetByName("Planeador Personal");
  var hojaDestino = libroOrigen.getSheetByName("planeador Anual");
  //var hojaDestino = libroOrigen.getSheetByName("planeador Despacho");
  
  var datos = hojaOrigen.getRange("A:AE").getValues();

  var ultimaFilaDestino = hojaDestino.getLastRow();

  var anioInicio = 2025, mesInicio = 8;  // septiembre 2025
  var anioFin = 2026, mesFin = 11;       // diciembre 2026


  var periodicidades = {
    "17 ABRIL": function (anio) { return obtenerAnualPorMedioDia(anio, 3, 17); }, //3 es abril
    "17 JUNIO": function (anio) { return obtenerAnualPorMedioDia(anio, 5, 17); }, // 5 junio
    "21 MAYO": function (anio) { return obtenerAnualPorMedioDia(anio, 4, 21); }, // Mayo = 4
    "23 JUNIO": function (anio) { return obtenerAnualPorMedioDia(anio, 5, 23); }, // 5 junio
    "29 NOVIEMBRE": function (anio) { return obtenerAnualPorMedioDia(anio, 10, 29); }, // Noviembre = 10
    "3 MARZO": function (anio) { return obtenerAnualPorMedioDia(anio, 2, 3); }, // Marzo = 2
    "3 NOVIEMBRE": function (anio) { return obtenerAnualPorMedioDia(anio, 10, 3); }, // Noviembre = 10
    
    "3ER MIERCOLES DE JULIO": function (anio) { return obtenerAnual(anio, 6, 14, 10); }, // Julio = 6
    "4TO LUNES DE ABRIL": function (anio) { return obtenerAnual(anio, 3, 21, 8); }, // Abril = 3
    "4TO LUNES DE AGOSTO": function (anio) { return obtenerAnual(anio, 7, 21, 8); }, // Agosto = 7 //no viene
    "4TO LUNES DE DICIEMBRE": function (anio) { return obtenerAnual(anio, 11, 21, 8); }, // Diciembre = 11
    "4TO LUNES DE ENERO": function (anio) { return obtenerAnual(anio, 0, 21, 8); }, // Enero = 0
    "4TO LUNES DE FEBRERO": function (anio) { return obtenerAnual(anio, 1, 21, 8); }, // Febrero = 1
    "4TO LUNES DE JULIO": function (anio) { return obtenerAnual(anio, 6, 21, 8); }, // Julio = 6
    "4TO LUNES DE JUNIO": function (anio) { return obtenerAnual(anio, 5, 21, 8); }, // Junio = 5
    "4TO LUNES DE MARZO": function (anio) { return obtenerAnual(anio, 2, 21, 8); }, // Marzo = 2
    "4TO LUNES DE MAYO": function (anio) { return obtenerAnual(anio, 4, 21, 8); }, // Mayo = 4
    "4TO LUNES DE NOVIEMBRE": function (anio) { return obtenerAnual(anio, 10, 21, 8); }, // Noviembre = 10
    "4TO LUNES DE SEPTIEMBRE": function (anio) { return obtenerAnual(anio, 8, 21, 8); }, // Septiembre = 8
    
    "6 JULIO": function (anio) { return obtenerAnualPorMedioDia(anio, 6, 6); }, // Julio = 6
    "6 NOVIEMBRE": function (anio) { return obtenerAnualPorMedioDia(anio, 10, 6); }, // Noviembre = 10
  };


  var salida = [];

  for (var i = 5; i < datos.length; i++) {
    var periodicidad = (datos[i][30] || "").toString().trim().toUpperCase();
    var funcion = periodicidades[periodicidad];
    if (!funcion) continue;

    // Tomar columnas C:AA (índices 2 a 26)
    var filaDatos = datos[i].slice(2, 28);

    // en vez de un solo año, recorre todos
    for (var anio = anioInicio; anio <= anioFin; anio++) {
      var fecha = funcion(anio);
      fecha = ajustarPorFestivoAnual(fecha);

      // Solo guardar si cae dentro del rango
      if (fecha >= new Date(anioInicio, mesInicio, 1) && fecha <= new Date(anioFin, mesFin, 28)) {
      //if (fecha >= new Date(anioInicio, mesInicio, 1) && fecha <= new Date(anioFin, mesFin, 29)) {

      
        var fechaFormateada = "'" + formatearFecha(fecha);  // <- apostrofe antes
        var nuevaFila = [fechaFormateada].concat(filaDatos);



        // asegura col AB = "NUEVO"
        while (nuevaFila.length < 26) { 
        //while (nuevaFila.length < 27) { 
          nuevaFila.push(""); 
        }
        nuevaFila.push("NUEVO");

        salida.push(nuevaFila);
        Logger.log("Periodicidad encontrada: " + periodicidad);
        Logger.log("Fecha generada: " + fecha);
      }
    }
  }

  // escribe a partir de col B
  if (salida.length > 0) {
    hojaDestino.getRange(ultimaFilaDestino + 1, 2, salida.length, salida[0].length).setValues(salida);
    
    // ✅ Formatear la columna de fechas (col B)
    hojaDestino.getRange(ultimaFilaDestino + 1, 2, salida.length, 1)
              .setNumberFormat("dd/MM/yyyy");
  }

}

function formatearFecha(fecha) {
  if (!fecha) return "";
  var dia = fecha.getDate();
  var mes = fecha.getMonth() + 1;
  var anio = fecha.getFullYear();

  // Asegura que día y mes tengan 2 dígitos
  var diaStr = (dia < 10 ? "0" : "") + dia;
  var mesStr = (mes < 10 ? "0" : "") + mes;

  return diaStr + "/" + mesStr + "/" + anio;
}


// Función para obtener el tercer lunes de un mes dado
//"3ER MIERCOLES DE JULIO": function (anio) { return obtenerAnual(anio, 6, 14, 10); }, //julio
                    //anio, 6, 14, 8
function obtenerAnual(anio, mes, sumterCuar, dias) {
  var fecha = new Date(anio, mes, 1);//busca el primer dia del mes
  var diaSemana = fecha.getDay(); //sacamos el primer dia de la semana
  var diasHastaLunes = (dias - diaSemana) % 7; //cualcula cuantos dias hasta el lunes
  var tercerLunes = 1 + diasHastaLunes + sumterCuar; //suma 14 dias para, el tercer lunes del mes o 21 para cuato dia del mes
  return new Date(anio, mes, tercerLunes);
} 

/*
  buscar las fechas:
    17 ABRIL
    17 JUNIO
    21 MAYO
    23 JUNIO
    29 NOVIEMBRE
    3 MARZO
    3 NOVIEMBRE
    6 JULIO
    6 NOVIEMBRE
    si cae en fin de semana, mover al dia habil anterior
*/
function obtenerAnualPorMedioDia(anio, mes, dia) {
  var fecha = new Date(anio, mes, dia);
  // Si cae en fin de semana, mover al día hábil anterior (viernes o antes)
  while (fecha.getDay() === 0 || fecha.getDay() === 6) { // 0=Dom, 6=Sab
    fecha.setDate(fecha.getDate() - 1);
  }
  return fecha;
}

//si cae en un dias festivo ara lo siguiente: 
//si cae lunes dias festivo que lo mueva para el Viernes de la semana anterior
//si cae un dia festivo que sea un dia antes pero vigente.
function ajustarPorFestivoAnual(fecha) {
  var festivosFijos = [
    { mes: 0, dia: 1 },   // 1 Enero
    { mes: 1, dia: 2 },   // 5 Febrero //Primer lunes de febrero
    //{ mes: 1, dia: 5 },   // 5 Febrero ==arreglar el codigo para esto.
    { mes: 2, dia: 16 },  // 21 Marzo //tercer lunes de marzo
    //{ mes: 2, dia: 21 },  // 21 Marzo ==arreglar el codigo para este.
    { mes: 4, dia: 1 },   // 1 Mayo
    { mes: 8, dia: 16 },  // 16 Septiembre
    { mes: 11, dia: 12 }, // 12 Diciembre
    { mes: 11, dia: 25 }  // 25 Diciembre
  ];

  function obtenerTercerLunesNoviembre(year) {
    var fecha = new Date(year, 10, 1); // 1 Noviembre
    var primerDia = fecha.getDay(); // 0=Dom, 1=Lun...
    var primerLunes = primerDia === 1 ? 1 : (8 - primerDia);
    var tercerLunes = primerLunes + 14; // sumo 14 días (dos semanas más)
    return new Date(year, 10, tercerLunes);
  }

  function esFestivo(d) {
    // Revisa los festivos fijos
    var esFijo = festivosFijos.some(f => d.getMonth() === f.mes && d.getDate() === f.dia);
    // Revisa si es tercer lunes de noviembre
    var tercerLunes = obtenerTercerLunesNoviembre(d.getFullYear());
    var esTercerLunesNov = d.getMonth() === 10 && d.getDate() === tercerLunes.getDate();
    return esFijo || esTercerLunesNov;
  }

  if (!esFestivo(fecha)) return fecha; // si no es festivo, regresa igual

  var dia = fecha.getDay(); // 0=Dom, 1=Lun...

  // lunes festivo → viernes de la semana pasada
    if (dia === 1) {
      var viernes = new Date(fecha); 
      viernes.setDate(fecha.getDate() - 3);//viernes de la semana pasada
      if (!esFestivo(viernes) && viernes.getDay() !== 0 && viernes.getDay() !== 6) {//no domingo, ni sabados
        return viernes;
      }
    }

    // Martes -> Lunes de esa semana
    if (dia === 2) {
      var viernes = new Date(fecha); 
      viernes.setDate(fecha.getDate() - 1);//lunes de la semana 
      if (!esFestivo(viernes) && viernes.getDay() !== 0 && viernes.getDay() !== 6) {//no domingo, ni sabados
        return viernes;
      }
    }

    // miércoles festivo → martes
    if (dia === 3) {
      var martes = new Date(fecha);
      martes.setDate(fecha.getDate() - 1); //martes de la semana 
      if (!esFestivo(martes) && martes.getDay() !== 0 && martes.getDay() !== 6) {
        return martes;
      }
    }

    // Jueves -> Miercoles de esa semana
    if (dia === 4) {
      var viernes = new Date(fecha); 
      viernes.setDate(fecha.getDate() - 1);//miercoles de la semana 
      if (!esFestivo(viernes) && viernes.getDay() !== 0 && viernes.getDay() !== 6) {//no domingo, ni sabados
        return viernes;
      }
    }

    // viernes -> jueves de esa semana
    if (dia === 5) {
      var viernes = new Date(fecha); 
      viernes.setDate(fecha.getDate() - 1);//jueves de la semana 
      if (!esFestivo(viernes) && viernes.getDay() !== 0 && viernes.getDay() !== 6) {//no domingo, ni sabados
        return viernes;
      }
    }

  return fecha; // si nada aplica, regresa la original
}

/*BIMESTRAL 22/12/2025*/
function generarBimestrales() {
  var libroOrigen = SpreadsheetApp.getActiveSpreadsheet();
  var hojaOrigen = libroOrigen.getSheetByName("S.Gastos CICLICOS INTERNO PS");
  //var hojaDestino = libroOrigen.getSheetByName("Planeador Despacho");
  var hojaDestino = libroOrigen.getSheetByName("Planeador Bimestral");
  //var hojaDestino = libroOrigen.getSheetByName("hojaPrueba");
  

  //var hojaOrigen = libroOrigen.getSheetByName("Copia de S.Gastos CICLICOS INTERNO PS(Personal)");
  //var hojaDestino = libroOrigen.getSheetByName("Planeador Personal");
  

  var datos = hojaOrigen.getRange("A:AE").getValues();

  var ultimaFilaDestino = hojaDestino.getLastRow();

  var anioInicio = 2025, mesInicio = 8;  // septiembre 2025
  var anioFin = 2026, mesFin = 11;       // diciembre 2026

  // Meses bimestrales
  var mesEneNovBiM = [0, 2, 4, 6, 8, 10]; //4TO LUNES / 3ER MIERCOLES DE ENE, MAR, MAY, JUL, SEP, NOV
  var mesFebDicBiMF = [1, 3, 5, 7, 9, 11]; //4TO LUNES DE FEB / 3ER MIERCOLES, FEB, ABR, JUN, AGO, OCT, DIC

  var periodicidades = {

    "1ER DIA HABIL DE ENE, MAR, MAY, JUL, SEP, NOV": function (anio) { return obtenerBimestral(anio, mesEneNovBiM, 0, 1); },
    "1ER DIA HABIL DE FEB, ABR, JUN, AGO, OCT, DIC": function (anio) { return obtenerBimestral(anio, mesFebDicBiMF, 0, 1); },
    "2DO MIERCOLES DE ENE, MAR, MAY, JUL, SEP, NOV": function (anio) { return obtenerBimestral(anio, mesEneNovBiM, 7, 3); },
    "2DO MIERCOLES DE FEB, ABR, JUN, AGO, OCT, DIC": function (anio) { return obtenerBimestral(anio, mesFebDicBiMF, 7, 3); },
    "3ER MIERCOLES DE ENE, MAR, MAY, JUL, SEP, NOV": function (anio) { return obtenerBimestral(anio, mesEneNovBiM, 14, 10); },
    "3ER MIERCOLES DE FEB, ABR, JUN, AGO, OCT, DIC": function (anio) { return obtenerBimestral(anio, mesFebDicBiMF, 14, 10); },
    "4TO LUNES DE ENE, MAR, MAY, JUL, SEP, NOV": function (anio) { return obtenerBimestral(anio, mesEneNovBiM, 21, 8); }

  };

  var salida = [];

  for (var i = 5; i < datos.length; i++) {
    //var periodicidad = (datos[i][29] || "").toString().trim().toUpperCase();// gastos personales
    var periodicidad = (datos[i][30] || "").toString().trim().toUpperCase();//gastos Despacho
    var funcion = periodicidades[periodicidad];
    if (!funcion) continue;

    // Tomar columnas C:AA (índices 2 a 26)
    var filaDatos = datos[i].slice(2, 27);

    // recorrer años dentro del rango
    for (var anio = anioInicio; anio <= anioFin; anio++) {
      var fechas = funcion(anio); // arreglo de fechas bimestrales //Invalid Date??? esta vacio, no hay fecha.
      fechas = ajustarPorFestivoBime(fechas); // ajusta lunes/miércoles festivos


      fechas.forEach(function(fecha) {
        if (fecha >= new Date(anioInicio, mesInicio, 1) && fecha <= new Date(anioFin, mesFin, 28)) {
          var fechaFormateada = formatearFechaB(fecha);
          var nuevaFila = [fechaFormateada].concat(filaDatos);

          // asegura que llegue hasta col AB
          //while (nuevaFila.length < 26) {
          while (nuevaFila.length < 27) {
            nuevaFila.push("");
          }
          nuevaFila.push("NUEVO");

          salida.push(nuevaFila);
          Logger.log("Periodicidad encontrada: " + periodicidad);
          Logger.log("Fecha generada: " + fecha);
        }
      });
    }
  }


  // Escribe la salida en la hoja destino
  if (salida.length > 0) {
    hojaDestino.getRange(ultimaFilaDestino + 1, 2, salida.length, salida[0].length).setValues(salida); 
    //solo se cambio de 1 a 2 porque empieza la col 2 osea B
    // ✅ Formatear la columna de fechas (col B)
    hojaDestino.getRange(ultimaFilaDestino + 1, 2, salida.length, 1)
             .setNumberFormat("dd/MM/yyyy");
  }
}

/*function primerDiaHabilDelMes(anio, mes) {
  var fecha = new Date(anio, mes, 1);
  while (fecha.getDay() === 0 || fecha.getDay() === 6) {
    fecha.setDate(fecha.getDate() + 1);
  }
  return fecha;
}*/


function formatearFechaB(fecha) {
  if (!fecha) return "";
  var dia = fecha.getDate();
  var mes = fecha.getMonth() + 1;
  var anio = fecha.getFullYear();
  return dia + "/" + mes + "/" + anio;
}

// Esta función debe devolver un arreglo de fechas, una por cada mes del arreglo
function obtenerBimestral(anio, mesesArr, sumterCuar, dias) {
  var fechas = [];
  mesesArr.forEach(function (mes) {
    var fecha = obtenerBimes(anio, mes, sumterCuar, dias);
    fechas.push(fecha);
  });
  return fechas;
}

// Función para obtener el lunes/miércoles según parámetros
/*function obtenerBimes(anio, mes, sumterCuar, dias) {
  var fecha = new Date(anio, mes, 1);
  var diaSemana = fecha.getDay(); // 0=Dom, 1=Lun, 2=Mar, 3=Mié... 4=Jueves 5=viernes 6=sabado
  var diasHastaDia = (dias - diaSemana) % 7;
  var diaFinal = 1 + diasHastaDia + sumterCuar;
  return new Date(anio, mes, diaFinal);
}*/

function obtenerBimes(anio, mes, sumterCuar, diaObjetivo) {
  var fecha = new Date(anio, mes, 1);
  var primerDia = fecha.getDay(); // día de la semana del 1 del mes

  // calcular cuántos días faltan para llegar al día de la semana objetivo
  var offset = (diaObjetivo - primerDia + 7) % 7;

  // primer occurrence
  var primerDiaObjetivo = 1 + offset;

  // luego sumas semanas
  var diaFinal = primerDiaObjetivo + sumterCuar;

  return new Date(anio, mes, diaFinal);
}


/* dias festivos**
Si cae en día festivo, ajustar la fecha:
Lunes -> viernes de la semana pasada
Martes -> Lunes de esa semana
Miercoles -> Martes de esa semana
Jueves -> Miercoles de esa semana
viernes -> Jueves de esa semana
No domingos, ni sabados
*/

function ajustarPorFestivoBime(fechas) {//actualizado 03/12/2025 '' si queda igual se iguala a los demas metodos.
  var festivosFijos = [
    { mes: 0, dia: 1 },   // 1 Enero
    { mes: 1, dia: 2 },   // 5 Febrero //Primer lunes de febrero
    //{ mes: 1, dia: 5 },   // 5 Febrero ==arreglar el codigo para esto.
    { mes: 2, dia: 16 },  // 21 Marzo //tercer lunes de marzo
    //{ mes: 2, dia: 21 },  // 21 Marzo ==arreglar el codigo para este.
    { mes: 4, dia: 1 },   // 1 Mayo
    { mes: 8, dia: 16 },  // 16 Septiembre
    { mes: 11, dia: 12 }, // 12 Diciembre
    { mes: 11, dia: 25 }  // 25 Diciembre
  ];

  /*El tercer lunes de noviembre lo tengo representado asi, porque no tiene un dia en espesifico sino que es variable. */
  // --- calcular el 3er lunes de noviembre dinámico ---
  function obtenerTercerLunesNoviembre(year) {
    var fecha = new Date(year, 10, 1); // 1 Noviembre
    var primerDia = fecha.getDay();    // 0=Dom, 1=Lun...
    var primerLunes = primerDia === 1 ? 1 : (8 - primerDia);
    var tercerLunes = primerLunes + 14; // dos semanas más
    return new Date(year, 10, tercerLunes);
  }

  function esFestivo(d) {
    // primero los fijos
    var fijo = festivosFijos.some(f => d.getMonth() === f.mes && d.getDate() === f.dia);
    // ahora el 3er lunes de noviembre
    var tercerLunes = obtenerTercerLunesNoviembre(d.getFullYear());
    var esTercerLunesNov = d.getMonth() === 10 && d.getDate() === tercerLunes.getDate();
    return fijo || esTercerLunesNov;
  }

  // recorrer todas las fechas
  return fechas.map(function(fecha) {
    if (!fecha) return fecha;
    if (!(fecha instanceof Date)) return fecha;

    if (!esFestivo(fecha)) return fecha; // si no es festivo, se queda igual

    var dia = fecha.getDay(); // 0=Dom, 1=Lun, 2=Mar, 3=Mié... 4=Jueves 5=viernes 6=sabado

    // lunes festivo → viernes de la semana pasada
    if (dia === 1) {
      var viernes = new Date(fecha); 
      viernes.setDate(fecha.getDate() - 3);//viernes de la semana pasada
      if (!esFestivo(viernes) && viernes.getDay() !== 0 && viernes.getDay() !== 6) {//no domingo, ni sabados
        return viernes;
      }
    }

    // Martes -> Lunes de esa semana
    if (dia === 2) {
      var viernes = new Date(fecha); 
      viernes.setDate(fecha.getDate() - 1);//lunes de la semana 
      if (!esFestivo(viernes) && viernes.getDay() !== 0 && viernes.getDay() !== 6) {//no domingo, ni sabados
        return viernes;
      }
    }

    // miércoles festivo → martes
    if (dia === 3) {
      var martes = new Date(fecha);
      martes.setDate(fecha.getDate() - 1); //martes de la semana 
      if (!esFestivo(martes) && martes.getDay() !== 0 && martes.getDay() !== 6) {
        return martes;
      }
    }

    // Jueves -> Miercoles de esa semana
    if (dia === 4) {
      var viernes = new Date(fecha); 
      viernes.setDate(fecha.getDate() - 1);//miercoles de la semana 
      if (!esFestivo(viernes) && viernes.getDay() !== 0 && viernes.getDay() !== 6) {//no domingo, ni sabados
        return viernes;
      }
    }

    // viernes -> jueves de esa semana
    if (dia === 5) {
      var viernes = new Date(fecha); 
      viernes.setDate(fecha.getDate() - 1);//jueves de la semana 
      if (!esFestivo(viernes) && viernes.getDay() !== 0 && viernes.getDay() !== 6) {//no domingo, ni sabados
        return viernes;
      }
    }
    

    return fecha; // si no aplica, devolver la original
  });
}

/*ELIMINAR*/
function eliminarFilasDespacho() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = libro.getSheetByName("planeador trimestral");
  //var hoja = libro.getSheetByName("Planeador Bimestral");
  if (!hoja) {
    Logger.log("La hoja 'despacho' no existe.");
    return;
  }
  var numFilas = hoja.getLastRow();
  if (numFilas > 0) {
    hoja.deleteRows(2, numFilas - 1); // Elimina desde la fila 2 hasta la última (deja encabezados)
  }
}

/*MENSUAL*/
function mensual() { //funciona bien = no tocar
  var libroOrigen = SpreadsheetApp.getActiveSpreadsheet();
  var hojaOrigen = libroOrigen.getSheetByName("S.Gastos CICLICOS INTERNO PS");
  //var hojaDestino = libroOrigen.getSheetByName("Planeador Despacho");
  var hojaDestino = libroOrigen.getSheetByName("planeador mensual");

  //var hojaOrigen = libroOrigen.getSheetByName("Copia de S.Gastos CICLICOS INTERNO PS(Personal)");
  //var hojaDestino = libroOrigen.getSheetByName("Planeador Personal");
  

  var ultimaFilaOrigen = hojaOrigen.getLastRow();
  var ultimaColumnaOrigen = hojaOrigen.getLastColumn();
  var datos = hojaOrigen.getRange(1, 1, ultimaFilaOrigen, ultimaColumnaOrigen).getValues();

  var ultimaFilaDestino = hojaDestino.getLastRow();

  var anioInicio = 2025, mesInicio = 8;  // septiembre
  var anioFin = 2026, mesFin = 11;       // diciembre

  var periodicidades = {
    "15 O 16 DE CADA MES" : quinceO16DiaHabilDelMes,
    "1ER DIA HABIL DEL MES" : primerDiaHabilDelMes,
    "1ER LUNES DE CADA MES" : primerLunesDelMes,
    "1ER MIERCOLES DE CADA MES" : primerMiercolesDelMes,
    "2DO LUNES DE CADA MES" : segundoLunesDelMes,
    "2DO MIERCOLES DE CADA MES" : segundoMiercolesDelMes,
    "2DO VIERNES DE CADA MES" : segundoViernesDelMes, //nuevo funciona
    "3ER DIA HABIL DEL MES" : tercerDiaHabilDelMes,//nuevo funciona
    "3ER LUNES DE CADA MES" : tercerLunesDelMes,
    "3ER MIERCOLES DE CADA MES" : tercerMiercolesDelMes,
    "3ER VIERNES DE CADA MES" : tercerViernesDelMes,//nuevo funciona
    "4TO LUNES DE CADA MES" : cuartoLunesDelMes,
    "4TO VIERNES DE CADA MES" : cuartoViernesDelMes,
    "9 DE CADA MES" : nueveDiaHabilDelMes,
    "11 DE CADA MES" : onceDiaHabilDelMes
  };

  var salida = [];

  for (var i = 5; i < datos.length; i++) {
    //var periodicidad = (datos[i][29] || "").toString().trim().toUpperCase(); //gastos personales
    var periodicidad = (datos[i][30] || "").toString().trim().toUpperCase();//gastos despacho
    var funcion = periodicidades[periodicidad];
    if (!funcion) continue;

    //var filaDatos = datos[i].slice(2, 30); // columnas C:AD
    var filaDatos = datos[i].slice(2, 27); // columnas C:AD

    // Reemplazar AB por "NUEVO" si está vacía o tiene cualquier valor
    filaDatos[26] = "NUEVO"; // posición 25 = columna AB

    for (var anio = anioInicio; anio <= anioFin; anio++) {
      var mesInicial = (anio === anioInicio ? mesInicio : 0);
      var mesFinal = (anio === anioFin ? mesFin : 11);

      for (var mes = mesInicial; mes <= mesFinal; mes++) {
        var fecha = funcion(anio, mes);
        fecha = ajustarPorFestivo(fecha, periodicidad);

        var fechaFormateada = formatearFechaM(fecha);

        // Insertar la fecha en columna B
        var nuevaFila = [ , fechaFormateada].concat(filaDatos); // A vacío, B=fecha, C:AE = filaDatos
        salida.push(nuevaFila);
      }
    }
  }

  if (salida.length > 0) {
   // hojaDestino.getRange(ultimaFilaDestino + 1, 1, salida.length, salida[0].length).setValues(salida);
    hojaDestino.getRange(ultimaFilaDestino + 1, 1, salida.length, salida[0].length).setValues(salida);
  }
}



//
// -------- FUNCIONES DE PERIODICIDAD --------
//

function quinceO16DiaHabilDelMes(anio, mes) {//fallo ==aqui me quedo
  var fecha = new Date(anio, mes, 15);// 0=Dom, 1=Lun, 2=Mar, 3=Mié... 4=Jueves 5=viernes 6=sabado
  if (fecha.getDay() === 6) {//sabados 
    fecha.setDate(fecha.getDate() - 1);
  }else if(fecha.getDay() === 0){ //y ni domingos
      fecha.setDate(fecha.getDate() - 2);
  }
  return fecha;
}

function primerDiaHabilDelMes(anio, mes) {
  var fecha = new Date(anio, mes, 1);
  while (fecha.getDay() === 0 || fecha.getDay() === 6) {//sabados y ni domingos
    fecha.setDate(fecha.getDate() + 1);
  }
  return fecha;
}

function primerLunesDelMes(anio, mes) {
  var fecha = new Date(anio, mes, 1);
  while (fecha.getDay() !== 1) {
    fecha.setDate(fecha.getDate() + 1);
  }
  return fecha;
}

function primerMiercolesDelMes(anio, mes) {
  var fecha = new Date(anio, mes, 1);
  while (fecha.getDay() !== 3) {
    fecha.setDate(fecha.getDate() + 1);
  }
  return fecha;
}

function segundoLunesDelMes(anio, mes) {
  var fecha = primerLunesDelMes(anio, mes);
  fecha.setDate(fecha.getDate() + 7);
  return fecha;
}

function segundoMiercolesDelMes(anio, mes) {
  var fecha = primerMiercolesDelMes(anio, mes);
  fecha.setDate(fecha.getDate() + 7);
  return fecha;
}

function segundoViernesDelMes(anio, mes) {//nuevo
  var fecha = new Date(anio, mes, 1);
  var count = 0;
  while (count < 2) {
    if (fecha.getDay() === 5) count++;
    if (count < 2) fecha.setDate(fecha.getDate() + 1);
  }
  return fecha;
}

function tercerDiaHabilDelMes(anio, mes) {//fallo ==aqui me quedo
  //var fecha = new Date(anio, mes, 1);
  var fecha = new Date(anio, mes, 3);// 0=Dom, 1=Lun, 2=Mar, 3=Mié... 4=Jueves 5=viernes 6=sabado
  if (fecha.getDay() === 6) {//sabados 
  //while (fecha.getDay() === 0 || fecha.getDay() === 6) {//sabados y ni domingos
    //fecha.setDate(fecha.getDate() + 3);
    fecha.setDate(fecha.getDate() - 1);
  }else if(fecha.getDay() === 0){ //y ni domingos
      fecha.setDate(fecha.getDate() - 2);
  }
  return fecha;
}

function tercerLunesDelMes(anio, mes) {
  var fecha = primerLunesDelMes(anio, mes);
  fecha.setDate(fecha.getDate() + 14);
  return fecha;
}

function tercerMiercolesDelMes(anio, mes) {
  var fecha = primerMiercolesDelMes(anio, mes);
  fecha.setDate(fecha.getDate() + 14);
  return fecha;
}

function tercerViernesDelMes(anio, mes) {//nuevo
  var fecha = new Date(anio, mes, 1);
  var count = 0;
  while (count < 3) {
    if (fecha.getDay() === 5) count++;
    if (count < 3) fecha.setDate(fecha.getDate() + 1);
  }
  return fecha;
}

function cuartoLunesDelMes(anio, mes) {
  var fecha = primerLunesDelMes(anio, mes);
  fecha.setDate(fecha.getDate() + 21);
  return fecha;
}

function cuartoViernesDelMes(anio, mes) {
  var fecha = new Date(anio, mes, 1);
  var count = 0;
  while (count < 4) {
    if (fecha.getDay() === 5) count++; //// 0=Dom, 1=Lun, 2=Mar, 3=Mié... 4=Jueves 5=viernes 6=sabado
    if (count < 4) fecha.setDate(fecha.getDate() + 1);
  }
  return fecha;
}

function ultimoViernesDelMes(anio, mes) {
  var fecha = new Date(anio, mes + 1, 0);
  while (fecha.getDay() !== 5) {
    fecha.setDate(fecha.getDate() - 1);
  }
  return fecha;
}

function nueveDiaHabilDelMes(anio, mes) {//fallo ==aqui me quedo
  var fecha = new Date(anio, mes, 9);// 0=Dom, 1=Lun, 2=Mar, 3=Mié... 4=Jueves 5=viernes 6=sabado
  if (fecha.getDay() === 6) {//sabados 
    fecha.setDate(fecha.getDate() - 1);
  }else if(fecha.getDay() === 0){ //y ni domingos
      fecha.setDate(fecha.getDate() - 2);
  }
  return fecha;
}

function onceDiaHabilDelMes(anio, mes) {//fallo ==aqui me quedo
  var fecha = new Date(anio, mes, 11);// 0=Dom, 1=Lun, 2=Mar, 3=Mié... 4=Jueves 5=viernes 6=sabado
  if (fecha.getDay() === 6) {//sabados 
    fecha.setDate(fecha.getDate() - 1);
  }else if(fecha.getDay() === 0){ //y ni domingos
      fecha.setDate(fecha.getDate() - 2);
  }
  return fecha;
}


//
// --------- Ajuste por días festivos ---------
//si cae en un dias festivo ara lo siguiente: 
//si es miércoles el dia festivo, mover al martes de ese semana
//cuando sea esta periosidad 1ER DIA HABIL DEL MES y caiga en dia festivo solo mover un dia habil de esta misma semana, si no mover a la siguiente semana pero primer dia habil de esa semana
//si caae un dia lunes el dia festivo mover al dia martes
//si cae un viernes mover al miercoles de la semana A sepcion la periosidad, porque si tiene periosidad 1ER DIA HABIL DEL MES:
//buscar el primer dia habil, ejemplo dia destivo 01/05/82026 es viernes pero su periosidad trae 1ER DIA HABIL DEL MES, si es asi debe buscsra el primer dia habil como 04/05/26
function ajustarPorFestivo(fecha, periodicidad) {
  var festivos = [
    { mes: 0, dia: 1 },   // 1 Enero
    { mes: 1, dia: 2 },   // 5 Febrero //Primer lunes de febrero
    //{ mes: 1, dia: 5 },   // 5 Febrero ==arreglar el codigo para esto.
    { mes: 2, dia: 16 },  // 21 Marzo //tercer lunes de marzo
    //{ mes: 2, dia: 21 },  // 21 Marzo ==arreglar el codigo para este.
    { mes: 4, dia: 1 },   // 1 Mayo
    { mes: 8, dia: 16 },  // 16 Septiembre
    { mes: 11, dia: 12 }, // 12 Diciembre
    { mes: 11, dia: 25 }  // 25 Diciembre
  ];

  function tercerLunesNoviembre(year) {
    var fecha = new Date(year, 10, 1); // 1 de noviembre
    var primerDia = fecha.getDay();
    var primerLunes = (primerDia === 1) ? 1 : ((8 - primerDia) % 7) + 1;
    return { mes: 10, dia: primerLunes + 14 }; // tercer lunes
  }

  // Agregar el tercer lunes de noviembre del año de la fecha
  festivos.push(tercerLunesNoviembre(fecha.getFullYear()));

  function esFestivo(d) {
    return festivos.some(f => d.getMonth() === f.mes && d.getDate() === f.dia);
  }

  if (!esFestivo(fecha)) return fecha;

  var dia = fecha.getDay();

  // lunes festivo → viernes de la semana pasada
    if (dia === 1) {
      var viernes = new Date(fecha); 
      viernes.setDate(fecha.getDate() - 3);//viernes de la semana pasada
      if (!esFestivo(viernes) && viernes.getDay() !== 0 && viernes.getDay() !== 6) {//no domingo, ni sabados
        return viernes;
      }
    }

    // Martes -> Lunes de esa semana
    if (dia === 2) {
      var viernes = new Date(fecha); 
      viernes.setDate(fecha.getDate() - 1);//lunes de la semana 
      if (!esFestivo(viernes) && viernes.getDay() !== 0 && viernes.getDay() !== 6) {//no domingo, ni sabados
        return viernes;
      }
    }

    // miércoles festivo → martes
    if (dia === 3) {
      var martes = new Date(fecha);
      martes.setDate(fecha.getDate() - 1); //martes de la semana 
      if (!esFestivo(martes) && martes.getDay() !== 0 && martes.getDay() !== 6) {
        return martes;
      }
    }

    // Jueves -> Miercoles de esa semana
    if (dia === 4) {
      var viernes = new Date(fecha); 
      viernes.setDate(fecha.getDate() - 1);//miercoles de la semana 
      if (!esFestivo(viernes) && viernes.getDay() !== 0 && viernes.getDay() !== 6) {//no domingo, ni sabados
        return viernes;
      }
    }

    // viernes -> jueves de esa semana
    if (dia === 5) {
      var viernes = new Date(fecha); 
      viernes.setDate(fecha.getDate() - 1);//jueves de la semana 
      if (!esFestivo(viernes) && viernes.getDay() !== 0 && viernes.getDay() !== 6) {//no domingo, ni sabados
        return viernes;
      }
    }

  return fecha;
}



//
// ---- Formato de fecha
//
function formatearFechaM(fecha) {
  return Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

/*QUINCENAL*/
/*
  codigo que de la hoja horigen "S.Gastos CICLICOS INTERNO PS" sacaras la periosdad cuando sea en la columna 29 === DIAS 15 Y DIAS 30 y cuando encuentes este cadena vas a sacar la fecha de cada mes 15 dias: osea en enero 15 miercoles y 30 jueves pero adcesion de fedrero la primera quincena va ser el 15 y (si el 15 sale sabado o domingo moverlo al viernes de preferencia) como en febrero no hay 30 hay que moverlo al ultimo dia vigente del mes (viegente de dias lunes a viernes). y esas fechas se pegaran en la columna A de la hojas destino : "Fechas 2025" y despes de la columna A ira el rango de C:AA de la hoja origen corespondiente perioridad que salga "DIAS 15 Y DIAS 30", el rango a sacar de la hoja destino es "A:AD".
*/
function quincenal() {//funciona no mover
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hojaOrigen = libro.getSheetByName("S.Gastos CICLICOS INTERNO PS");
  var hojaDestino = libro.getSheetByName("planeador quincenal");
  //var hojaDestino = libro.getSheetByName("Base");

  //var hojaOrigen = libro.getSheetByName("Copia de S.Gastos CICLICOS INTERNO PS(Personal)");
  //var hojaDestino = libro.getSheetByName("Planeador Personal");
  //var hojaDestino = libro.getSheetByName("hojaPrueba");

  var datos = hojaOrigen.getRange("A:AE").getValues(); // Incluye todo hasta col 30
  var salida = [];

  // Rango de fechas personalizable
  var anioInicio = 2025, mesInicio = 8;  // Septiembre 2025 (0=Enero)
  var anioFin = 2026, mesFin = 11;       // Diciembre 2026

  for (var i = 1; i < datos.length; i++) { // desde la fila 2
    var periodicidad = (datos[i][30] || "").toString().trim().toUpperCase();

    if (periodicidad === "DIAS 15 Y DIAS 30") {
      //var filaDatos = datos[i].slice(2, 27); // columnas C:AA
      var filaDatos = datos[i].slice(2, 28); // columnas C:AA

      for (var anio = anioInicio; anio <= anioFin; anio++) {
        var mesStart = (anio === anioInicio) ? mesInicio : 0;
        var mesEnd = (anio === anioFin) ? mesFin : 11;

        for (var mes = mesStart; mes <= mesEnd; mes++) {
          // ---- Primera quincena ----
          var fecha1 = ajustarSiFinDeSemana(new Date(anio, mes, 15));
          var nuevaFila1 = [formatearFechaQ(fecha1)].concat(filaDatos);
          nuevaFila1 = ajustarPorFestivoQuin(nuevaFila1); // ajusta lunes/miércoles festivos


          // Completa hasta col AB y agrega "NUEVO"
          //while (nuevaFila1.length < 26) nuevaFila1.push("");
          while (nuevaFila1.length < 26) nuevaFila1.push("");
          nuevaFila1.push("NUEVO");
          salida.push(nuevaFila1);

          // ---- Segunda quincena ----
          var fecha2;
          if (mes === 1) { // Febrero
            fecha2 = ultimoDiaHabilDelMes(anio, mes);
          } else {
            fecha2 = ajustarSiFinDeSemana(new Date(anio, mes, 30));
          }

          var nuevaFila2 = [formatearFechaQ(fecha2)].concat(filaDatos);
          nuevaFila2 = ajustarPorFestivoQuin(nuevaFila2); // ajusta lunes/miércoles festivos
          //while (nuevaFila2.length < 26) nuevaFila2.push("");
          while (nuevaFila2.length < 26) nuevaFila2.push("");
          nuevaFila2.push("NUEVO");
          salida.push(nuevaFila2);// salida agregar mes que salga con numero

          Logger.log("Periodicidad encontrada: " + periodicidad);
          Logger.log("Fechas generadas: " + fecha1 + ", " + fecha2);
        }
      }
    }
  }

  // Pegar en hoja destino desde columna B
 // Pegar en hoja destino desde columna B
  if (salida.length > 0) {
    var ultimaFilaDestino = hojaDestino.getLastRow();
    var numCols = salida[0].length; // columnas reales en la fila
    hojaDestino.getRange(ultimaFilaDestino + 1, 2, salida.length, numCols).setValues(salida);

    // ✅ Formatear la columna de fechas (col B)
    hojaDestino.getRange(ultimaFilaDestino + 1, 2, salida.length, 1)
              .setNumberFormat("dd/MM/yyyy");
  }

}

function ajustarSiFinDeSemana(fecha) {
  var diaSemana = fecha.getDay(); // 0=Domingo, 6=Sábado
  if (diaSemana === 0) fecha.setDate(fecha.getDate() - 2);
  else if (diaSemana === 6) fecha.setDate(fecha.getDate() - 1);
  return fecha;
}

function ultimoDiaHabilDelMes(anio, mes) {
  var fecha = new Date(anio, mes + 1, 0); // último día del mes
  while (fecha.getDay() === 0 || fecha.getDay() === 6) {
    fecha.setDate(fecha.getDate() - 1);
  }
  return fecha;
}

function formatearFechaQ(fecha) {
  var dia = fecha.getDate();
  var mes = fecha.getMonth() + 1;
  var anio = fecha.getFullYear();
  return dia + "/" + mes + "/" + anio;
}

function ajustarPorFestivoQuin(fila) {
  // fila[0] es la fecha en formato "d/m/yyyy"
  var festivos = [
    { mes: 0, dia: 1 },   // 1 Enero
    { mes: 1, dia: 2 },   // 5 Febrero //Primer lunes de febrero
    //{ mes: 1, dia: 5 },   // 5 Febrero ==arreglar el codigo para esto.
    { mes: 2, dia: 16 },  // 21 Marzo //tercer lunes de marzo
    //{ mes: 2, dia: 21 },  // 21 Marzo ==arreglar el codigo para este.
    { mes: 4, dia: 1 },   // 1 Mayo
    { mes: 8, dia: 16 },  // 16 Septiembre
    { mes: 11, dia: 12 }, // 12 Diciembre
    { mes: 11, dia: 25 }  // 25 Diciembre
  ];

  // Calcular tercer lunes de noviembre para un año dado
  function tercerLunesNoviembre(year) {
    var fecha = new Date(year, 10, 1); // 1 de noviembre
    var primerDia = fecha.getDay();    // 0=Dom,...,6=Sab
    // calcular primer lunes
    var primerLunes = (primerDia === 1) ? 1 : ((8 - primerDia) % 7) + 1;
    // sumar 14 días para llegar al tercer lunes
    return { mes: 10, dia: primerLunes + 14 };
  }

  function esFestivo(d) {
    var year = d.getFullYear();
    // agregar el tercer lunes de noviembre dinámico
    var festivoMovil = tercerLunesNoviembre(year);
    var todosFestivos = festivos.concat(festivoMovil);

    return todosFestivos.some(f => d.getMonth() === f.mes && d.getDate() === f.dia);
  }

  // convertir la fecha de la fila a objeto Date
  var partes = fila[0].split("/");
  var fecha = new Date(Number(partes[2]), Number(partes[1]) - 1, Number(partes[0]));

  if (esFestivo(fecha)) {
    var dia = fecha.getDay();

     // lunes festivo → viernes de la semana pasada
    if (dia === 1) {
      var lunes = new Date(fecha); 
      lunes.setDate(fecha.getDate() - 3);//viernes de la semana pasada
      if (!esFestivo(lunes) && lunes.getDay() !== 0 && lunes.getDay() !== 6) {//no domingo, ni sabados
        fila[0] = formatearFechaQ(lunes);
        return fila;
      }
    }

     // Martes -> Lunes de esa semana
    if (dia === 2) {
      var martes = new Date(fecha); 
      martes.setDate(fecha.getDate() - 1);//lunes de la semana 
      if (!esFestivo(martes) && martes.getDay() !== 0 && martes.getDay() !== 6) {//no domingo, ni sabados
        fila[0] = formatearFechaQ(martes);
        return fila;
      }
    }

    // miércoles festivo → martes
    if (dia === 3) {
      var miercoles = new Date(fecha);
      miercoles.setDate(fecha.getDate() - 1); //martes de la semana 
      if (!esFestivo(miercoles) && miercoles.getDay() !== 0 && miercoles.getDay() !== 6) {
        fila[0] = formatearFechaQ(miercoles);
        return fila;
      }
    }

    // Jueves -> Miercoles de esa semana
    if (dia === 4) {
      var jueves = new Date(fecha); 
      jueves.setDate(fecha.getDate() - 1);//miercoles de la semana 
      if (!esFestivo(jueves) && jueves.getDay() !== 0 && jueves.getDay() !== 6) {//no domingo, ni sabados
        fila[0] = formatearFechaQ(jueves);
        return fila;
      }
    }

    // viernes -> jueves de esa semana
    if (dia === 5) {
      var viernes = new Date(fecha); 
      viernes.setDate(fecha.getDate() - 1);//jueves de la semana 
      if (!esFestivo(viernes) && viernes.getDay() !== 0 && viernes.getDay() !== 6) {//no domingo, ni sabados
        fila[0] = formatearFechaQ(viernes);
        return fila;
      }
    }

  }

  return fila;
}
/* SEMANAL*/
/*
codigo que de la hoja horigen "S.Gastos CICLICOS INTERNO PS" sacaras la periosdad cuando sea en la columna 29 === "CADA LUNES || CADA VIERNES ||
DIARIO" y cuando encuentes este cadena vas a sacar laS fechaS: cada lunes de cada mes y viernes y diario sin contar sabado o domigos y esas fechas se pegaran en la columna A de la hojas destino : "Fechas 2025" y despes de la columna A ira el rango de C:AA de la hoja origen corespondiente perioridad que salga "CADA LUNES || CADA VIERNES ||
DIARIO", el rango a sacar de la hoja destino es "A:AD".
/*
    Extrae de la hoja origen "S.Gastos CICLICOS INTERNO PS" las filas donde la columna 29 (AC) sea "CADA LUNES", "CADA VIERNES" o "DIARIO".
    Para cada coincidencia, genera todas las fechas de 2025 según la periodicidad:
        - "CADA LUNES": todos los lunes del año (sin sábados ni domingos)
        - "CADA VIERNES": todos los viernes del año (sin sábados ni domingos)
        - "DIARIO": todos los días del año excepto sábados y domingos
    Por cada fecha, pega en la hoja destino "Fechas 2025" en la columna A la fecha, y de la B a la AA los datos de la fila origen (C:AA).
    El rango de salida es "A:AD".
*/
function generarSemanal() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hojaOrigen = libro.getSheetByName("S.Gastos CICLICOS INTERNO PS");
  var hojaDestino = libro.getSheetByName("planeador semanal");
  //var hojaDestino = libro.getSheetByName("hojaPrueba");

  //var hojaOrigen = libro.getSheetByName("Copia de S.Gastos CICLICOS INTERNO PS(Personal)");
  //var hojaDestino = libro.getSheetByName("Planeador Personal");

  var datos = hojaOrigen.getRange("A:AE").getValues();
  var ultimaFilaDestino = hojaDestino.getLastRow();

  // Rango de fechas personalizable
  var anioInicio = 2025, mesInicio = 8;  // Septiembre 2025 (0=Enero)
  var anioFin = 2026, mesFin = 11;       // Diciembre 2026
  /*var fechaInicio = new Date(anioInicio, mesInicio, 1);
  var fechaFin = new Date(anioFin, mesFin + 1, 0);*/

  var salida = [];

  for (var i = 1; i < datos.length; i++) {
    //var periodicidad = (datos[i][29] || "").toString().trim().toUpperCase();
    var periodicidad = (datos[i][30] || "").toString().trim().toUpperCase();
    //if (periodicidad !== "CADA LUNES" && periodicidad !== "CADA VIERNES" && periodicidad !== "DIARIO") continue;
    if (periodicidad !== "CADA VIERNES" && periodicidad !== "TODOS LOS LUNES HABILES") continue;
  
    //var filaDatos = datos[i].slice(2, 27);
    var filaDatos = datos[i].slice(2, 28);
  
    for (var anio = anioInicio; anio <= anioFin; anio++) {
      // Determina el mes de inicio y fin para el primer y último año
      var mesIni = (anio === anioInicio) ? mesInicio : 0;
      var mesFinLoop = (anio === anioFin) ? mesFin : 11;
  
      for (var mes = mesIni; mes <= mesFinLoop; mes++) {
        var fechaIniMes = new Date(anio, mes, 1);
        var fechaFinMes = new Date(anio, mes + 1, 0);
  
        var fechas = [];
        //if (periodicidad === "CADA LUNES") fechas = obtenerDiasPorSemanaRango(fechaIniMes, fechaFinMes, 1);
        if (periodicidad === "CADA VIERNES") fechas = obtenerDiasPorSemanaRango(fechaIniMes, fechaFinMes, 5);
        else if (periodicidad === "TODOS LOS LUNES HABILES") fechas = todosLunesHabiles(fechaIniMes, fechaFinMes, 1); // 0=Dom, 1=Lun, 2=Mar, 3=Mié... 4=Jueves 5=viernes 6=sabado
        //else if (periodicidad === "DIARIO") fechas = obtenerDiasHabilesRango(fechaIniMes, fechaFinMes);
  
        fechas.forEach(function(fecha) {
          fecha = ajustarPorFestivoSem(fecha);
          var fila = [formatearFechaS(fecha)].concat(filaDatos);
          fila.push("NUEVO");
          salida.push(fila);
          Logger.log("Periodicidad: " + periodicidad + " | Fecha generada: " + fecha);
        });
      }
    }
  }

  // Escribir en hoja destino desde la columna B
  if (salida.length > 0) {
    hojaDestino.getRange(ultimaFilaDestino + 1, 2, salida.length, salida[0].length).setValues(salida);
    // ✅ Formatear la columna de fechas (col B)
    hojaDestino.getRange(ultimaFilaDestino + 1, 2, salida.length, 1)
             .setNumberFormat("dd/MM/yyyy");
  }
}

// Días hábiles dentro de un rango
/*function obtenerDiasHabilesRango(fechaInicio, fechaFin) {
  var fechas = [];
  var f = new Date(fechaInicio);
  while (f <= fechaFin) {
    if (f.getDay() !== 0 && f.getDay() !== 6) fechas.push(new Date(f));
    f.setDate(f.getDate() + 1);
  }
  return fechas;
}*/

// Días de semana (lunes=1, viernes=5) dentro de un rango
function obtenerDiasPorSemanaRango(fechaInicio, fechaFin, diaSemana) { // 0=Dom, 1=Lun, 2=Mar, 3=Mié... 4=Jueves 5=viernes 6=sabado
  var fechas = [];
  var f = new Date(fechaInicio);
  while (f <= fechaFin) {
    if (f.getDay() === diaSemana) fechas.push(new Date(f));
      f.setDate(f.getDate() + 1);
  }
  return fechas;
}

// Días de semana (lunes=1, viernes=5) dentro de un rango
function todosLunesHabiles(fechaInicio, fechaFin, diaSemana) { // 0=Dom, 1=Lun, 2=Mar, 3=Mié... 4=Jueves 5=viernes 6=sabado
  var fechas = [];
  var f = new Date(fechaInicio);
  while (f <= fechaFin) {
    if (f.getDay() === diaSemana) fechas.push(new Date(f));
    f.setDate(f.getDate() + 1);
  }
  return fechas;
}

// Formato de fecha
function formatearFechaS(fecha) {
  //return fecha.getDate() + "/" + (fecha.getMonth() + 1) + "/" + fecha.getFullYear();
  return Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}


/*
si cae lune dia festivo se pasa a martes de eesa semana
si cae viernes dia festivo se pasa a miercoles de esa semana
si cae diario seria los siguientes:
  si cae lunes dia festivo que lo mueva para el martes de ese semana
  si cae martes dia festivo que lo mueva para el miercoles de esa semana
  si cae miercoles dia festivo se pasa a martes de esa semana
  si cae jueves dia festivo que lo mueva para el miercoles de esa semana
  si cae viernes dia festivo que lo mueva para el miercoles de esa semana
*/

function ajustarPorFestivoSem(fecha) {
  // Festivos fijos
  var festivos = [
    { mes: 0, dia: 1 },   // 1 Enero
    { mes: 1, dia: 2 },   // 5 Febrero //Primer lunes de febrero
    //{ mes: 1, dia: 5 },   // 5 Febrero ==arreglar el codigo para esto.
    { mes: 2, dia: 16 },  // 21 Marzo //tercer lunes de marzo
    //{ mes: 2, dia: 21 },  // 21 Marzo ==arreglar el codigo para este.
    { mes: 4, dia: 1 },   // 1 Mayo
    { mes: 8, dia: 16 },  // 16 Septiembre
    { mes: 11, dia: 12 }, // 12 Diciembre
    { mes: 11, dia: 25 }  // 25 Diciembre
  ];

  // Calcular tercer lunes de noviembre para el año dado
  function tercerLunesNoviembre(year) {
    var fecha = new Date(year, 10, 1); // 1 de noviembre
    var primerDia = fecha.getDay();    // 0=Dom,...,6=Sab
    // calcular primer lunes
    var primerLunes = (primerDia === 1) ? 1 : ((8 - primerDia) % 7) + 1;
    // sumar 14 días para llegar al tercer lunes
    return { mes: 10, dia: primerLunes + 14 };
  }

  function esFestivo(d) {
    var year = d.getFullYear();
    var festivoMovil = tercerLunesNoviembre(year);
    var todosFestivos = festivos.concat(festivoMovil);

    return todosFestivos.some(f => d.getMonth() === f.mes && d.getDate() === f.dia);
  }

  if (!fecha) return fecha;

  var dia = fecha.getDay(); // 0=Dom, 1=Lun,...,5=Vie

  if (!esFestivo(fecha)) return fecha;

  // lunes festivo → viernes de la semana pasada
    if (dia === 1) {
      var viernes = new Date(fecha); 
      viernes.setDate(fecha.getDate() - 3);//viernes de la semana pasada
      if (!esFestivo(viernes) && viernes.getDay() !== 0 && viernes.getDay() !== 6) {//no domingo, ni sabados
        return viernes;
      }
    }

    // Martes -> Lunes de esa semana
    if (dia === 2) {
      var viernes = new Date(fecha); 
      viernes.setDate(fecha.getDate() - 1);//lunes de la semana 
      if (!esFestivo(viernes) && viernes.getDay() !== 0 && viernes.getDay() !== 6) {//no domingo, ni sabados
        return viernes;
      }
    }

    // miércoles festivo → martes
    if (dia === 3) {
      var martes = new Date(fecha);
      martes.setDate(fecha.getDate() - 1); //martes de la semana 
      if (!esFestivo(martes) && martes.getDay() !== 0 && martes.getDay() !== 6) {
        return martes;
      }
    }

    // Jueves -> Miercoles de esa semana
    if (dia === 4) {
      var viernes = new Date(fecha); 
      viernes.setDate(fecha.getDate() - 1);//miercoles de la semana 
      if (!esFestivo(viernes) && viernes.getDay() !== 0 && viernes.getDay() !== 6) {//no domingo, ni sabados
        return viernes;
      }
    }

    // viernes -> jueves de esa semana
    if (dia === 5) {
      var viernes = new Date(fecha); 
      viernes.setDate(fecha.getDate() - 1);//jueves de la semana 
      if (!esFestivo(viernes) && viernes.getDay() !== 0 && viernes.getDay() !== 6) {//no domingo, ni sabados
        return viernes;
      }
    }

  // Si no aplica ninguna regla, devuelve la fecha original
  return fecha;
}

/*TRIMESTRAL*/ 
function generarTrimestral() {
  var libroOrigen = SpreadsheetApp.getActiveSpreadsheet();
  var hojaOrigen = libroOrigen.getSheetByName("S.Gastos CICLICOS INTERNO PS");
  var hojaDestino = libroOrigen.getSheetByName("planeador trimestral");
  //var hojaDestino = libroOrigen.getSheetByName("hojaPrueba");
  
  //var hojaOrigen = libroOrigen.getSheetByName("Copia de S.Gastos CICLICOS INTERNO PS(Personal)");
  //var hojaDestino = libroOrigen.getSheetByName("Planeador Personal");
  
  var datos = hojaOrigen.getRange("A:AE").getValues();

  var ultimaFilaDestino = hojaDestino.getLastRow();

  // Rango de fechas personalizable
  var anioInicio = 2025, mesInicio = 8;  // Septiembre 2025 (0=Enero)
  var anioFin = 2026, mesFin = 11;       // Diciembre 2026

  //2DO LUNES DE FEBRERO, MAYO, AGOSTO, NOVIEMBRE
  var segundoLunes = [1, 4, 7, 10]; //FEBRERO, MAYO, AGOSTO, NOVIEMBRE

  var periodicidades = {
    
    "2DO LUNES DE FEBRERO, MAYO, AGOSTO, NOVIEMBRE": function(anio) { //NUEVO
      return obtenerTrimestral(anio, segundoLunes, "2DO_LUNES"); 
    }
  };


  var salida = [];

  for (var i = 5; i < datos.length; i++) {
    //var periodicidad = (datos[i][29] || "").toString().trim().toUpperCase();
    var periodicidad = (datos[i][30] || "").toString().trim().toUpperCase();
    var funcion = periodicidades[periodicidad];
    if (!funcion) continue;

    // Tomar columnas C:AA (índices 2 a 26)
    var filaDatos = datos[i].slice(2, 27);
    //var filaDatos = datos[i].slice(2, 28);

    // Recorrer los años dentro del rango
    for (var anio = anioInicio; anio <= anioFin; anio++) {
      var fechas = funcion(anio); // Devuelve un arreglo de fechas trimestrales
      fechas = ajustarPorFestivoTrime(fechas); // ajusta lunes/miércoles festivos

      fechas.forEach(function(fecha) {
        // Solo fechas dentro del rango
        if (fecha >= new Date(anioInicio, mesInicio, 1) && fecha <= new Date(anioFin, mesFin, 28)) {
          var fechaFormateada = formatearFechaT(fecha);
          var nuevaFila = [fechaFormateada].concat(filaDatos);

          // Asegura que llegue hasta la columna AB
          while (nuevaFila.length < 27) {
            nuevaFila.push("");
          }
          nuevaFila.push("NUEVO");

          salida.push(nuevaFila);

          Logger.log("Periodicidad encontrada: " + periodicidad);
          Logger.log("Fecha generada: " + fecha);
        }
      });
    }
  }

  // Escribe la salida en la hoja destino desde la columna B
  if (salida.length > 0) {
    hojaDestino.getRange(ultimaFilaDestino + 1, 2, salida.length, salida[0].length).setValues(salida);
    // ✅ Formatear la columna de fechas (col B)
    hojaDestino.getRange(ultimaFilaDestino + 1, 2, salida.length, 1)
             .setNumberFormat("dd/MM/yyyy");
  }
}

// Formato de fecha
function formatearFechaT(fecha) {
  if (!fecha) return "";
  var dia = fecha.getDate();
  var mes = fecha.getMonth() + 1;
  var anio = fecha.getFullYear();
  return dia + "/" + mes + "/" + anio;
}

function primerLunesDelMes(anio, mes) {
  var fecha = new Date(anio, mes, 1);
  while (fecha.getDay() !== 1) {
    fecha.setDate(fecha.getDate() + 1);
  }
  return fecha;
}

function segundoLunesDelMes(anio, mes) {
  var fecha = primerLunesDelMes(anio, mes);
  fecha.setDate(fecha.getDate() + 7);
  return fecha;
}

/*function segundoDiaHabilDelMes(anio, mes) {//fallo ==aqui me quedo
  var fecha = new Date(anio, mes, 1);// 0=Dom, 1=Lun, 2=Mar, 3=Mié... 4=Jueves 5=viernes 6=sabado
  if (fecha.getDay() === 6) {//sabados 
    fecha.setDate(fecha.getDate() - 1);
  }else if(fecha.getDay() === 0){ //y ni domingos
      fecha.setDate(fecha.getDate() - 2);
  }
  return fecha;
}*/

// Devuelve fechas trimestrales correctas
function obtenerTrimestral(anio, mesesArr, tipo) {
  var fechas = [];
  mesesArr.forEach(function(mes) {
    if (tipo === "2DO_LUNES") {
      fechas.push(segundoLunesDelMes(anio, mes));
    }
  });
  return fechas;
}

//dias festivos
/*
si cae un dia festivo en miercoles que pase al mastes de esa semana
si cae lunes festivo que pase al mastes de esa semana
y si no pasa nada pasa
 */
/*function ajustarPorFestivoTrime(fechas) {
  var festivos = [
    { mes: 0, dia: 1 },   // 1 Enero
    { mes: 1, dia: 5 },   // 5 Febrero
    { mes: 2, dia: 21 },  // 21 Marzo
    { mes: 4, dia: 1 },   // 1 Mayo
    { mes: 8, dia: 16 },  // 16 Septiembre
    { mes: 10, dia: 20 }, // 20 Noviembre
    { mes: 11, dia: 12 },  // 12 Diciembre
    { mes: 11, dia: 25 }  // 25 Diciembre
  ];

  function esFestivo(d) {
    return festivos.some(f => d.getMonth() === f.mes && d.getDate() === f.dia);
  }

  // Recorrer todas las fechas y ajustar
  return fechas.map(function(fecha) {
    if (!fecha) return fecha;

    if (!esFestivo(fecha)) return fecha; // si no es festivo, regresar igual

    var dia = fecha.getDay(); // 0=Dom, 1=Lun, 3=Mié

    // Miércoles festivo → martes
    if (dia === 3) {
      var martes = new Date(fecha);
      martes.setDate(fecha.getDate() - 1);
      if (!esFestivo(martes) && martes.getDay() !== 0 && martes.getDay() !== 6) {
        return martes;
      }
    }

    // Lunes festivo → martes
    if (dia === 1) {
      var martes = new Date(fecha);
      martes.setDate(fecha.getDate() + 1);
      if (!esFestivo(martes) && martes.getDay() !== 0 && martes.getDay() !== 6) {
        return martes;
      }
    }

    return fecha; // si no aplica, regresar fecha original
  });
}*/
function ajustarPorFestivoTrime(fechas) {
  var festivosFijos = [
    { mes: 0, dia: 1 },   // 1 Enero
    { mes: 1, dia: 2 },   // 5 Febrero //Primer lunes de febrero
    //{ mes: 1, dia: 5 },   // 5 Febrero ==arreglar el codigo para esto.
    { mes: 2, dia: 16 },  // 21 Marzo //tercer lunes de marzo
    //{ mes: 2, dia: 21 },  // 21 Marzo ==arreglar el codigo para este.
    { mes: 4, dia: 1 },   // 1 Mayo
    { mes: 8, dia: 16 },  // 16 Septiembre
    { mes: 11, dia: 12 }, // 12 Diciembre
    { mes: 11, dia: 25 }  // 25 Diciembre
  ];
  

  // 👉 calcular dinámicamente el 3er lunes de noviembre
  function obtenerTercerLunesNoviembre(year) {
    var fecha = new Date(year, 10, 1); // 1 Noviembre
    var primerDia = fecha.getDay();    // 0=Dom, 1=Lun...
    var primerLunes = primerDia === 1 ? 1 : (8 - primerDia);
    var tercerLunes = primerLunes + 14; // tercer lunes = primer lunes + 14 días
    return new Date(year, 10, tercerLunes);
  }

  function esFestivo(d) {
    // festivos fijos
    var esFijo = festivosFijos.some(f => d.getMonth() === f.mes && d.getDate() === f.dia);

    // tercer lunes de noviembre
    var tercerLunes = obtenerTercerLunesNoviembre(d.getFullYear());
    var esTercerLunes = d.getMonth() === 10 && d.getDate() === tercerLunes.getDate();

    return esFijo || esTercerLunes;
  }

  // Recorrer todas las fechas y ajustar
  return fechas.map(function (fecha) {
    if (!fecha) return fecha;

    if (!esFestivo(fecha)) return fecha; // si no es festivo, regresar igual

    var dia = fecha.getDay(); // 0=Dom, 1=Lun, 2=Mar, 3=Mié...


    // lunes festivo → viernes de la semana pasada
    if (dia === 1) {
      var lunes = new Date(fecha); 
      lunes.setDate(fecha.getDate() - 3);//viernes de la semana pasada
      if (!esFestivo(lunes) && lunes.getDay() !== 0 && lunes.getDay() !== 6) {//no domingo, ni sabados
        return lunes;
      }
    }

    // Martes -> Lunes de esa semana
    if (dia === 2) {
      var martes = new Date(fecha); 
      martes.setDate(fecha.getDate() - 1);//lunes de la semana 
      if (!esFestivo(martes) && martes.getDay() !== 0 && martes.getDay() !== 6) {//no domingo, ni sabados
        return martes;
      }
    }

    // miércoles festivo → martes
    if (dia === 3) {
      var miercoles = new Date(fecha);
      miercoles.setDate(fecha.getDate() - 1); //martes de la semana 
      if (!esFestivo(miercoles) && miercoles.getDay() !== 0 && miercoles.getDay() !== 6) {
        return miercoles;
      }
    }

    // Jueves -> Miercoles de esa semana
    if (dia === 4) {
      var jueves = new Date(fecha); 
      jueves.setDate(fecha.getDate() - 1);//miercoles de la semana 
      if (!esFestivo(jueves) && jueves.getDay() !== 0 && jueves.getDay() !== 6) {//no domingo, ni sabados
        return jueves;
      }
    }

    // viernes -> jueves de esa semana
    if (dia === 5) {
      var viernes = new Date(fecha); 
      viernes.setDate(fecha.getDate() - 1);//jueves de la semana 
      if (!esFestivo(viernes) && viernes.getDay() !== 0 && viernes.getDay() !== 6) {//no domingo, ni sabados
        return viernes;
      }
    }

    // Miércoles festivo → martes
    /*if (dia === 3) {
      var martes = new Date(fecha);
      martes.setDate(fecha.getDate() - 1);
      if (!esFestivo(martes) && martes.getDay() !== 0 && martes.getDay() !== 6) {
        return martes;
      }
    }

    // Lunes festivo → martes
    if (dia === 1) {
      var martes = new Date(fecha);
      martes.setDate(fecha.getDate() + 1);
      if (!esFestivo(martes) && martes.getDay() !== 0 && martes.getDay() !== 6) {
        return martes;
      }
    }*/

    return fecha; // si no aplica, regresar fecha original
  });
}

/*CUATRIMESTRAL*/
function generarCuatrimestral() {//funciono
  var libroOrigen = SpreadsheetApp.getActiveSpreadsheet();
  var hojaOrigen = libroOrigen.getSheetByName("S.Gastos CICLICOS INTERNO PS");
  var hojaDestino = libroOrigen.getSheetByName("planeador cuatrimestral");
  //var hojaDestino = libroOrigen.getSheetByName("hojaPrueba");
  
  //var hojaOrigen = libroOrigen.getSheetByName("Copia de S.Gastos CICLICOS INTERNO PS(Personal)");
  //var hojaDestino = libroOrigen.getSheetByName("Planeador Personal");
  
  var datos = hojaOrigen.getRange("A:AE").getValues();

  var ultimaFilaDestino = hojaDestino.getLastRow();

  // Rango de fechas personalizable
  var anioInicio = 2025, mesInicio = 8;  // Septiembre 2025 (0=Enero)
  var anioFin = 2026, mesFin = 11;       // Diciembre 2026

   //1ER LUNES DE ENE, MAYO Y SEP
  var primerLunes = [0, 4, 8]; //FEBRERO, ENE, MAYO Y SEP

  var periodicidades = {
    
    "1ER LUNES DE ENE, MAYO Y SEP": function(anio) { //NUEVO
      return obtenerCuatri(anio, primerLunes); 
    }
  };


  var salida = [];

  for (var i = 5; i < datos.length; i++) {
    //var periodicidad = (datos[i][29] || "").toString().trim().toUpperCase();
    var periodicidad = (datos[i][30] || "").toString().trim().toUpperCase();
    var funcion = periodicidades[periodicidad];
    if (!funcion) continue;

    // Tomar columnas C:AA (índices 2 a 26)
    var filaDatos = datos[i].slice(2, 27);
    //var filaDatos = datos[i].slice(2, 28);

    // Recorrer los años dentro del rango
    for (var anio = anioInicio; anio <= anioFin; anio++) {
      var fechas = funcion(anio); // Devuelve un arreglo de fechas trimestrales
      fechas = ajustarPorFestivoCuatri(fechas, periodicidad); // ajusta lunes/miércoles festivos

      fechas.forEach(function(fecha) {
        // Solo fechas dentro del rango
        if (fecha >= new Date(anioInicio, mesInicio, 1) && fecha <= new Date(anioFin, mesFin, 28)) {
          var fechaFormateada = formatearFechaC(fecha);
          var nuevaFila = [fechaFormateada].concat(filaDatos);

          // Asegura que llegue hasta la columna AB
          while (nuevaFila.length < 27) {
            nuevaFila.push("");
          }
          nuevaFila.push("NUEVO");

          salida.push(nuevaFila);

          Logger.log("Periodicidad encontrada: " + periodicidad);
          Logger.log("Fecha generada: " + fecha);
        }
      });
    }
  }

  // Escribe la salida en la hoja destino desde la columna B
  if (salida.length > 0) {
    hojaDestino.getRange(ultimaFilaDestino + 1, 2, salida.length, salida[0].length).setValues(salida);
    // ✅ Formatear la columna de fechas (col B)
    hojaDestino.getRange(ultimaFilaDestino + 1, 2, salida.length, 1)
             .setNumberFormat("dd/MM/yyyy");
  }
}

// Formato de fecha
function formatearFechaC(fecha) {
  if (!fecha) return "";
  var dia = fecha.getDate();
  var mes = fecha.getMonth() + 1;
  var anio = fecha.getFullYear();
  return dia + "/" + mes + "/" + anio;
}

function primerLunes(anio, mes) {
  var fecha = new Date(anio, mes, 1);
  while (fecha.getDay() !== 1) {
    fecha.setDate(fecha.getDate() + 1);
  }
  return fecha;
}

// Devuelve fechas trimestrales correctas
function obtenerCuatri(anio, mesesArr) {
  var fechas = [];
  mesesArr.forEach(function(mes) {
      fechas.push(primerLunes(anio, mes));
    
  });
  return fechas;
}

//dias festivos
/*
si cae un dia festivo en miercoles que pase al mastes de esa semana
si cae lunes festivo que pase al mastes de esa semana
y si no pasa nada pasa
 */

function ajustarPorFestivoCuatri(fechas, periodicidad) {
  var festivosFijos = [
    { mes: 0, dia: 1 },   // 1 Enero
    { mes: 1, dia: 2 },   // 5 Febrero //Primer lunes de febrero
    //{ mes: 1, dia: 5 },   // 5 Febrero ==arreglar el codigo para esto.
    { mes: 2, dia: 16 },  // 21 Marzo //tercer lunes de marzo
    //{ mes: 2, dia: 21 },  // 21 Marzo ==arreglar el codigo para este.
    { mes: 4, dia: 1 },   // 1 Mayo
    { mes: 8, dia: 16 },  // 16 Septiembre
    { mes: 11, dia: 12 }, // 12 Diciembre
    { mes: 11, dia: 25 }  // 25 Diciembre
  ];
  

  // 👉 calcular dinámicamente el 3er lunes de noviembre
  function obtenerTercerLunesNoviembre(year) {
    var fecha = new Date(year, 10, 1); // 1 Noviembre
    var primerDia = fecha.getDay();    // 0=Dom, 1=Lun...
    var primerLunes = primerDia === 1 ? 1 : (8 - primerDia);
    var tercerLunes = primerLunes + 14; // tercer lunes = primer lunes + 14 días
    return new Date(year, 10, tercerLunes);
  }

  function esFestivo(d) {
    // festivos fijos
    var esFijo = festivosFijos.some(f => d.getMonth() === f.mes && d.getDate() === f.dia);

    // tercer lunes de noviembre
    var tercerLunes = obtenerTercerLunesNoviembre(d.getFullYear());
    var esTercerLunes = d.getMonth() === 10 && d.getDate() === tercerLunes.getDate();

    return esFijo || esTercerLunes;
  }

  // Recorrer todas las fechas y ajustar
  return fechas.map(function (fecha) {
    if (!fecha) return fecha;

    if (!esFestivo(fecha)) return fecha; // si no es festivo, regresar igual

    var dia = fecha.getDay(); // 0=Dom, 1=Lun, 2=Mar, 3=Mié...


    // lunes festivo → viernes de la semana pasada
    if (dia === 1) {
      var lunes = new Date(fecha); 
      lunes.setDate(fecha.getDate() - 3);//viernes de la semana pasada
      if (!esFestivo(lunes) && lunes.getDay() !== 0 && lunes.getDay() !== 6) {//no domingo, ni sabados
        return lunes;
      }
    }

    // Martes -> Lunes de esa semana
    if (dia === 2) {
      var martes = new Date(fecha); 
      martes.setDate(fecha.getDate() - 1);//lunes de la semana 
      if (!esFestivo(martes) && martes.getDay() !== 0 && martes.getDay() !== 6) {//no domingo, ni sabados
        return martes;
      }
    }

    // miércoles festivo → martes
    if (dia === 3) {
      var miercoles = new Date(fecha);
      miercoles.setDate(fecha.getDate() - 1); //martes de la semana 
      if (!esFestivo(miercoles) && miercoles.getDay() !== 0 && miercoles.getDay() !== 6) {
        return miercoles;
      }
    }

    // Jueves -> Miercoles de esa semana
    if (dia === 4) {
      var jueves = new Date(fecha); 
      jueves.setDate(fecha.getDate() - 1);//miercoles de la semana 
      if (!esFestivo(jueves) && jueves.getDay() !== 0 && jueves.getDay() !== 6) {//no domingo, ni sabados
        return jueves;
      }
    }

    // viernes -> jueves de esa semana
    if (dia === 5) {
      var viernes = new Date(fecha); 
      viernes.setDate(fecha.getDate() - 1);//jueves de la semana 
      if (!esFestivo(viernes) && viernes.getDay() !== 0 && viernes.getDay() !== 6) {//no domingo, ni sabados
        return viernes;
      }
    }



    return fecha; // si no aplica, regresar fecha original
  });
}
