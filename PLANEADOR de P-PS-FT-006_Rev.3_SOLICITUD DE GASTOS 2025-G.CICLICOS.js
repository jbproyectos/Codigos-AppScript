function ciclicosBoton() {
  try {
    var hojasDatos = [
      {
        link: "11fwocVoPzNY6IuKPMtjXZRLOkdGQP5eDTStTq0NnEls",
        //link: "11MiUy67oMXIPwyb5kPV9E8-PQU9WZQqJiGm2HgoACCc",
        destino: "S.Gastos CICLICOS INTERNO PS A4",
        origen: "Base de Datos Despacho"
      },

      {
        link: "1dNBBBTgpPPXMQ4Jc7E_sdTrwhuHPk0JQH83k6sst5Kw",
        //link: "1ZkdKx6z5ibd6HEYwoa0oeKgziPNZ7itZFcYG9Uy55iU",
        destino: "S.Gastos CICLICOS INTERNO PS A3",
        origen: "Base de Datos Despacho"
      },

      {
        link: "1s9bacMC_D5STQ2BHhSry9aVcIpQtoLCgh8G5MA9at4Y",
        //link: "1ecfRYOMajfyuui4w5iAIwct0gJVUPg4MjbmGRfPpRT4",
        destino: "S.Gastos CICLICOS INTERNO PS A0",
        origen: "Base de Datos Despacho"
      },

      {
        link: "1Ic3r0-q88eRPL9PraWt8JjCRiu5Umn1ViKbqoq8SBkg",
        //link: "10ShS-qAuQvmfthOaN326EgbXQfKMB4mI9znPnwFGVag",
        destino: "S.Gastos Personales",
        origen: "Base de Datos Personal"
      }
    ];

    hojasDatos.forEach(function (hoja) {

      try {

        envioInfoCiclico_rapidoV3(hoja);

      } catch (err) {

        Logger.log(
          `❌ Error procesando ${hoja.destino}: ${err.message}`
        );
      }
    });

  } catch (e) {

    Logger.log(
      "❌ Error general en ciclicosBoton: " + e.message
    );
  }
}

function envioInfoCiclico_rapidoV3(hojaInfo) {

  var libroOrigen = SpreadsheetApp.getActiveSpreadsheet();

  var hojaOrigen =
    libroOrigen.getSheetByName(hojaInfo.origen);

  var ultimaFila = hojaOrigen.getLastRow();

  if (ultimaFila < 2) return;

  var datos = hojaOrigen.getRange(2, 1, ultimaFila - 1, 29).getValues();

  var libroDestino = SpreadsheetApp.openById(hojaInfo.link);

  var hojaDestino = libroDestino.getSheetByName(hojaInfo.destino);

  var fechaHoy = Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd/MM/yy');

  var generarF;

 
  // COLUMNAS
  const COL_CATEGORIA = 7;
  const COL_SUBCATEGORIA = 8;
  const COL_DESCRIPCION = 10;

 
  // FUNCION VALIDACION
  function coincideRegla(fila, config) {

    // tipo
    if (config.tipo && fila[4] !== config.tipo) { return false; }

    // categorias
    if (config.categorias && !config.categorias.includes( fila[COL_CATEGORIA] )) {  return false;}

    // subcategorias
    if (config.subcategorias && !config.subcategorias.includes( fila[COL_SUBCATEGORIA] )) { return false; }

    // descripcion
    if (config.descripcionContiene) {//va condecir con las primeras letras
      let descripcion = String(fila[COL_DESCRIPCION] ).toUpperCase();
      let coincide = config.descripcionContiene.some(texto => descripcion.includes( texto.toUpperCase() ));

      if (!coincide) { return false; }
    }

    return true;
  }

  
  // REGLAS
  var reglas = {

    // A0
    "S.Gastos CICLICOS INTERNO PS A0": fila =>
      coincideRegla(fila, {
        tipo: "DESPACHO",
        categorias: [  "REPRESENTANTES" ]
      })

      ||

      coincideRegla(fila, {
        tipo: "DESPACHO",
        categorias: [  "SOFTWARE" ],
        subcategorias: [
          "ZOOM"
        ],
        descripcionContiene: [
          "ZOOM PRO"
        ]
      })

      ||

      coincideRegla(fila, {
        tipo: "DESPACHO",
        categorias: [  "SOFTWARE" ],
        subcategorias: [
          "ICLOUD",
          "GOOGLE DOMAINS",
          "CANVA",
          "CORREO",
          "WORKY"
        ]
      }),

    
    // A3
    "S.Gastos CICLICOS INTERNO PS A3": fila =>
      coincideRegla(fila, {

        tipo: "DESPACHO",

        categorias: [
          "SERVICIOS EXTERNOS",
          "INSUMOS"
        ]
      }),

    
    // A4
    "S.Gastos CICLICOS INTERNO PS A4": fila =>
      coincideRegla(fila, {
        tipo: "DESPACHO",
        categorias: [
          "SOFTWARE"
        ],
        subcategorias: [
          "MICROSIP"
        ]
      })

      ||

      coincideRegla(fila, {
        tipo: "DESPACHO",
        categorias: [
          "IMPUESTOS",
          "TELEFONOS CELULARES",
          "VEHICULOS",
          "RENTAS DE OFICINAS",
          "SERVICIOS DE OFICINA",
          "MANTENIMIENTO DE OFICINAS"
        ]
      }),


    // personal
    "S.Gastos Personales": fila =>
      coincideRegla(fila, {
        tipo: "PERSONAL",
      }),

  };

  
  // FILTRADO
  var filas = datos.filter(function (fila) {

    var fecha = fila[1];

    if (!(fecha instanceof Date)) {
      return false;
    }

    if (Utilities.formatDate(fecha,Session.getScriptTimeZone(),'dd/MM/yy') !== fechaHoy) {  return false;}

    var ahora = new Date();

    var hora = Utilities.formatDate(ahora, Session.getScriptTimeZone(), "HH:mm:ss" );

    var fecha2 = Utilities.formatDate( ahora, Session.getScriptTimeZone(), 'dd/MM/yy'  );

    generarF = fecha2 + " " + hora;

    var status = fila[28];

    if (status !== "NUEVO") {return false;}

    return reglas[hojaInfo.destino] ? reglas[hojaInfo.destino](fila)  : true;
  });

  
  // SI NO HAY FILAS
  if (filas.length === 0) {
    Logger.log( "ℹ️ No hay filas para " +  hojaInfo.destino);
    return;
  }

  
  // PEGAR
  var inicioPegado =
    ultimaFilaNoVaciaV1(hojaDestino);

  hojaDestino.getRange(inicioPegado + 1, 1, filas.length, 29).setValues(filas);

  hojaDestino.getRange( inicioPegado + 1, 2, filas.length, 1).setValue(generarF);

  Logger.log(
    `✅ ${hojaInfo.destino}: Pegadas ${filas.length} filas.`
  );
}

function ultimaFilaNoVaciaV1(hoja) {

  if (!hoja) {
    Logger.log(  "La hoja " + hoja + " no existe.");
    return;
  }

  const columna = hoja.getRange("D:D").getValues();

  let ultimaFila = 0;

  for (let i = columna.length - 1; i >= 0; i-- ) {
    if (columna[i][0] !== "") {
      ultimaFila = i + 1;
      break;
    }
  }

  return ultimaFila;
}
