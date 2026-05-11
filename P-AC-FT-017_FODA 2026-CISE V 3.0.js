function onOpen() { 
    var ui = SpreadsheetApp.getUi();

    ui.createMenu('📅 | EnvioProductividad/Bloqueos (Exclusivo CEO A)')
    .addItem('1. Limpieza COMENTARIOS y envio en FODA + BloqueoCalificaciones| 📄', 'unionMetodos')
    .addItem('2. Envio a PRODUCTIVIDAD desde FODA| 📄', 'botonEnvioAProctividad')
    .addToUi();
}


function unionMetodos(){//onOpen
  copiadoFoda();
  corridaBloqueos();
}


function botonEnvioAProctividad() { //Ana 06/03/2026 //ya lo hace bien.
  try {

    var hojasDatos = [

      // ORIGINALES
      { link: "1vTanMlBVg7tF9JoJ2Af7l0RRKQnHBVSjn_7Pd0gcxOE", destino: "PRODUCTIVIDAD_AGENDA" },
      { link: "1vTanMlBVg7tF9JoJ2Af7l0RRKQnHBVSjn_7Pd0gcxOE", destino: "PRODUCTIVIDAD_PERSONAL_DOMESTICO" },

      { link: "1qyRAlrV6UkAQJLrNKVdgCaa2flgjIvRRfvCORUEUn5c", destino: "PRODUCTIVIDAD BANCOS" },
      { link: "1qyRAlrV6UkAQJLrNKVdgCaa2flgjIvRRfvCORUEUn5c", destino: "PRODUCTIVIDAD DOMICILIOS" },

      { link: "1EA93mvWa19QeynSqFJEk-AOOrmyKWMRiYjsgPRjS4qE", destino: "PRODUCTIVIDAD_CONTABILIDAD" },
      { link: "1EA93mvWa19QeynSqFJEk-AOOrmyKWMRiYjsgPRjS4qE", destino: "PRODUCTIVIDAD_AFILIACIONES" },
      { link: "1EA93mvWa19QeynSqFJEk-AOOrmyKWMRiYjsgPRjS4qE", destino: "PRODUCTIVIDAD_FACTURACION" },

      { link: "13owDqwxzUC9Czr2guQTaNVZ56Nx5yXHD1oCubUIG87g", destino: "PRODUCTIVIDAD_JURIDICO" },
      { link: "13owDqwxzUC9Czr2guQTaNVZ56Nx5yXHD1oCubUIG87g", destino: "PRODUCTIVIDAD_RH" },

      { link: "1jKYdpdagz38r51D47nKx9L6qpB5IMa9lba9fw_mNOEw", destino: "PRODUCTIVIDAD_LOGISTICA" },

      { link: "1GaISjrMli1OKpVsqj1hzYnGkOZJgvc67MN6YoVp5NAM", destino: "PRODUCTIVIDAD_OPERACIONES" },

      { link: "1an8CQihNBsiiizC3aenyVsq2oaG1sppTgF6hdpbL2Jg", destino: "PRODUCTIVIDAD_PRESUPUESTOS" },
      { link: "1an8CQihNBsiiizC3aenyVsq2oaG1sppTgF6hdpbL2Jg", destino: "PRODUCTIVIDAD_VERIFICACIONES" },

      // CORREGIDOS (tenían _ y la hoja probablemente tiene espacio)
      { link: "16O3ad4e4Nxw0W9RkozA9WesKulZg8vlX7jI0DEG9TYo", destino: "PRODUCTIVIDAD PROYECTOS" },
      { link: "16O3ad4e4Nxw0W9RkozA9WesKulZg8vlX7jI0DEG9TYo", destino: "PRODUCTIVIDAD SISTEMAS" },

      { link: "1GKaum0yFCzVIBN4z98SI2wGgg2AFcUjn9KP9ExC336Y", destino: "PRODUCTIVIDAD_TESORERIA" }

    ];

    hojasDatos.forEach(function (hoja) {

      try {

        Logger.log("Procesando hoja: " + hoja.destino);

        envioInf(hoja);

      } catch (err) {

        Logger.log("❌ Error procesando " + hoja.destino + ": " + err.message);

      }

    });

  } catch (e) {

    Logger.log("❌ Error general en botonEnvioAProctividad: " + e.message);

  }
}


function envioInf(hojaInf) {

  // =========================
  // HOJA ORIGEN
  // =========================
  var libro = SpreadsheetApp.getActiveSpreadsheet();

  var hojaOri = libro.getSheetByName("FODA 2026");

  /*const hojaOri = libro.getSheets().find(s =>
    s.getName().startsWith("FODA")
  );*/

  if (!hojaOri) {

    Logger.log("❌ La hoja 'FODA 2026' no existe.");
    return;

  }

  // =========================
  // LIBRO DESTINO
  // =========================
  var libroDestino = SpreadsheetApp.openById(hojaInf.link);

  if (!libroDestino) {

    Logger.log("❌ No se pudo abrir el archivo destino.");
    return;

  }

  // =========================
  // HOJA DESTINO
  // =========================
  var hojaDes = libroDestino.getSheetByName(hojaInf.destino);

  if (!hojaDes) {

    Logger.log("❌ La hoja '" + hojaInf.destino + "' no existe.");
    return;

  }

  // =========================
  // FILAS
  // =========================
  var ultimaFilaDi = obtenerUltimaFilaNoVaciaSolicitudes(hojaDes);

  var nuevaFilaDestino = ultimaFilaDi + 1;

  var ultimaFilaOrigen = obtenerUltimaFilaNoVaciaSolicitudes(hojaOri);

  var filaInicio = 2;

  var rango = hojaOri
    .getRange(filaInicio, 1, ultimaFilaOrigen - filaInicio + 1, 8)
    .getValues();

  // =========================
  // FILTRO NUEVOS
  // =========================
  var filasFiltradas = rango.filter(function(fila) {

    return fila[6] &&
           fila[6].toString().trim() === "NUEVO" &&
           String(fila[7] || "").trim() === "";

  });

  // =========================
  // FILTRO POR ÁREA
  // =========================
  var filas = filasFiltradas.filter(function (fila){

    var area = String(fila[2] || "").trim();
    var subarea = String(fila[3] || "").trim();

    var nombreHoja = hojaDes.getName();

    switch (nombreHoja){

      case "PRODUCTIVIDAD_AGENDA":
        return (
          area === "ASISTENCIA EJECUTIVA" &&
          (
            subarea === "ASISTENCIA EJECUTIVA" ||
            subarea === "PRESUPUESTOS PERSONAL"
          )
        );

      case "PRODUCTIVIDAD_PERSONAL_DOMESTICO":
        return area === "ASISTENCIA EJECUTIVA" &&
               subarea === "PERSONAL DOMESTICO";

      case "PRODUCTIVIDAD BANCOS":
        return area === "BANCOS" &&
               subarea === "BANCOS";

      case "PRODUCTIVIDAD DOMICILIOS":
        return area === "DOMICILIOS" &&
               subarea === "DOMICILIOS";

      case "PRODUCTIVIDAD_CONTABILIDAD":
        return area === "CONTABILIDAD" &&
               subarea === "CONTABILIDAD";

      case "PRODUCTIVIDAD_AFILIACIONES":
        return area === "CONTABILIDAD" &&
               subarea === "AFILIACIONES";

      case "PRODUCTIVIDAD_FACTURACION":
        return area === "FACTURACIÓN" &&
               subarea === "FACTURACIÓN";

      case "PRODUCTIVIDAD_JURIDICO":
        return area === "JURÍDICO" &&
               subarea === "JURÍDICO";

      case "PRODUCTIVIDAD_RH":
        return area === "RECURSOS HUMANOS" &&
               subarea === "RECURSOS HUMANOS";

      case "PRODUCTIVIDAD_LOGISTICA":
        return area === "LOGÍSTICA" &&
               subarea === "LOGÍSTICA";

      case "PRODUCTIVIDAD_OPERACIONES":
        return area === "OPERACIÓN" &&
               subarea === "OPERACIÓN";

      case "PRODUCTIVIDAD_PRESUPUESTOS":
        return area === "PRESUPUESTOS" &&
               subarea === "PRESUPUESTOS";

      case "PRODUCTIVIDAD_VERIFICACIONES":
        return area === "VERIFICACIÓN" &&
               subarea === "VERIFICACIÓN";

      case "PRODUCTIVIDAD PROYECTOS":
        return area === "PROYECTOS" &&
               subarea !== "SISTEMAS";

      case "PRODUCTIVIDAD SISTEMAS":
        return area === "PROYECTOS" &&
               subarea === "SISTEMAS";

      case "PRODUCTIVIDAD_TESORERIA":
        return area === "TESORERÍA" &&
               subarea === "TESORERÍA";

      default:

        Logger.log("⚠️ No existe CASE para: " + nombreHoja);

        return false;
    }

  });

  // =========================
  // ACTUALIZAR H
  // =========================
  filas.forEach(function(filaFiltrada) {

    rango.forEach(function(fila, i) {

      if (fila[4] === filaFiltrada[4]) {

        hojaOri
          .getRange(filaInicio + i, 8)
          .setValue("Trasladar al archivo de Productividad");

      }

    });

  });

  Logger.log('✅ Se actualizó columna H en FODA');

  // =========================
  // PEGADO
  // =========================
  if (filas.length > 0) {

    var datosFinales = filas.map(function(f) {

      var filaNueva = new Array(14).fill("");

      filaNueva[1] = f[5];
      filaNueva[2] = "INTERNO";
      filaNueva[3] = f[2];
      filaNueva[4] = "FODA";
      filaNueva[5] = "Actividad bimestral - máx 60 días";
      filaNueva[13] = f[4];

      return filaNueva.slice(1);

    });

    hojaDes
      .getRange(nuevaFilaDestino, 2, datosFinales.length, 13)
      .setValues(datosFinales);

    Logger.log("✅ Datos copiados en: " + hojaInf.destino);

  } else {

    Logger.log("⚠️ No hay filas para copiar en: " + hojaInf.destino);

  }

}

function obtenerUltimaFilaNoVaciaSolicitudes(hoja) {
  if (!hoja) {
    Logger.log("La hoja "+ hoja + " no existe.");
    return;
  }
  
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
}
//bloqueo
const CONFIG_BLOQUEO = {//ya lo hace 08/05/2026

  BLOQUEO_AIP: {//A:I -> T
    rangos: [
      { filaInicio: 3, columnaInicio: 1, columnaFin: 9 },  // A:I
      //{ filaInicio: 3, columnaInicio: 20, columnaFin: 20 } // O
    ],
    usuariosBloqueados: [
      'dlav.agenda_ejecutiva@grupo-cise.com',
      'mdlatg.agenda_ejecutiva@grupo-cise.com',
      'aasl.bancos@grupo-cise.com',
      'ajgm.bancos@grupo-cise.com',
      'bss.bancos@grupo-cise.com',
      'etg.bancos@grupo-cise.com',
      'jald.bancos@grupo-cise.com',
      'kmr.cobranza@grupo-cise.com',
      'lums.cobranza@grupo-cise.com',
      'aarv.contabilidad@grupo-cise.com',
      'asdg.contabilidad@grupo-cise.com',
      'bas.contabilidad@grupo-cise.com',
      'egc.contabilidad@grupo-cise.com',
      'fanl.contabilidad@grupo-cise.com',
      'gga.contabilidad@grupo-cise.com',
      'hamm.contabilidad@grupo-cise.com',
      'imm.contabilidad@grupo-cise.com',
      'jdjlp.contabilidad@grupo-cise.com',
      'jcrl.contabilidad@grupo-cise.com',
      'mgv.contabilidad@grupo-cise.com',
      'rrf.contabilidad@grupo-cise.com',
      'var.contabilidad@grupo-cise.com',
      'jodz.domicilios@grupo-cise.com',
      'mfco.domicilios@grupo-cise.com',
      'mgc.domicilios@grupo-cise.com',
      'tmc.tesoreria@grupo-cise.com',
      'agv.facturacion@grupo-cise.com',
      'aact.facturacion@grupo-cise.com',
      'anco.facturacion@grupo-cise.com',
      'bfg.facturacion@grupo-cise.com',
      'dsmm.facturacion@grupo-cise.com',
      'deev.facturacion@grupo-cise.com',
      'ers.facturacion@grupo-cise.com',
      'era.facturacion@grupo-cise.com',
      'feg.facturacion@grupo-cise.com',
      'gjar.facturacion@grupo-cise.com',
      'has.facturacion@grupo-cise.com',
      'jimt.facturacion@grupo-cise.com',
      'jaag.facturacion@grupo-cise.com',
      'jjar.facturacion@grupo-cise.com',
      'jam.facturacion@grupo-cise.com',
      'rrs.facturacion@grupo-cise.com',
      'rlt.facturacion@grupo-cise.com',
      'srf.facturacion@grupo-cise.com',
      'xrf.facturacion@grupo-cise.com',
      'amvt.juridico@grupo-cise.com',
      'lmsi.juridico@grupo-cise.com',
      'ccl.logistica@grupo-cise.com',
      'egts.logistica@grupo-cise.com',
      'jlsc.logistica@grupo-cise.com',
      'amf.operacion@grupo-cise.com',
      'asre.operacion@grupo-cise.com',
      'effc.operacion@grupo-cise.com',
      'ercv.operacion@grupo-cise.com',
      'ears.operacion@grupo-cise.com',
      'fsdlr.operacion@grupo-cise.com',
      'fcl.operacion@grupo-cise.com',
      'hav.operacion@grupo-cise.com',
      'iarc.operacion@grupo-cise.com',
      'jepl.operacion@grupo-cise.com',
      'jgr.operacion@grupo-cise.com',
      'ljss.operacion@grupo-cise.com',
      'lvaa.operacion@grupo-cise.com',
      'mott.operacion@grupo-cise.com',
      'marb.operacion@grupo-cise.com',
      'mfrr.operacion@grupo-cise.com',
      'nvsm.operacion@grupo-cise.com',
      'rsm.operacion@grupo-cise.com',
      'revl.operacion@grupo-cise.com',
      'rvb.operacion@grupo-cise.com',
      'sevg.operacion@grupo-cise.com',
      'smgb.operacion@grupo-cise.com',
      'uffc.operacion@grupo-cise.com',
      'vrr.operacion@grupo-cise.com',
      'ylrc.operacion@grupo-cise.com',
      'dahs.personal_domestico@grupo-cise.com',
      'esb.personal_domestico@grupo-cise.com',
      'cnrr.presupuestos@grupo-cise.com',
      'flmr.presupuestos@grupo-cise.com',
      'fdpg.presupuestos@grupo-cise.com',
      'niet.presupuestos@grupo-cise.com',
      'nnla.presupuestos@grupo-cise.com',
      'vavg.presupuestos@grupo-cise.com',
      'cglg.rrhh@grupo-cise.com',
      'jovs.sistemas@grupo-cise.com',
      'aavv.tesoreria@grupo-cise.com',
      'cggg.tesoreria@grupo-cise.com',
      'sss.tesoreria@grupo-cise.com',
      'mdv.tesoreria@grupo-cise.com',
      'yvv.tesoreria@grupo-cise.com',
      'jllc.tesoreria@grupo-cise.com',
      'arr.verificacion@grupo-cise.com',
      'dngt.verificacion@grupo-cise.com',
      'ftmg.verificacion@grupo-cise.com',
      'jlg.verificacion@grupo-cise.com',
      'lgpb.verificacion@grupo-cise.com',
      'mem.verificacion@grupo-cise.com',
      'mavr.verificacion@grupo-cise.com',
      'nyjm.verificacion@grupo-cise.com',
      'psb.verificacion@grupo-cise.com',
      'aat.direccion_gral@grupo-cise.com',
      'reportes@kabzo.org',
      'abbydobbleb.99@gmail.com',
      'dirgeneral@kubicspaces.com',
      'grupoviaya@gmail.com',
      'lars.ceos@grupo-cise.com',
      'ycl.ceos@grupo-cise.com',
      'yetp.ceos@grupo-cise.com'
    ],
    propietarios: [//SIN CEOS
      'ft.proyectos@grupo-cise.com',
      'gs.proyectos@grupo-cise.com',
      'bs.proyectos@grupo-cise.com',
      'jlmv.proyectos@grupo-cise.com',
      'jb.proyectos@grupo-cise.com',
      'ap.proyectos@grupo-cise.com',
      'ma.proyectos@grupo-cise.com',
      'sb.proyectos@grupo-cise.com'
    ]
  },

  BLOQUEO_JNP: {   // 👈 SEGUNDA CORRIDA //J:N -> U == J:P
    rangos: [
      //{ filaInicio: 3, columnaInicio: 10, columnaFin: 14 }, // J:N
      { filaInicio: 3, columnaInicio: 10, columnaFin: 16 }, // J:P
      { filaInicio: 3, columnaInicio: 21, columnaFin: 21 }  // U
    ],
    usuariosBloqueados: [
        'dlav.agenda_ejecutiva@grupo-cise.com',
        'mdlatg.agenda_ejecutiva@grupo-cise.com',
        'aasl.bancos@grupo-cise.com',
        'ajgm.bancos@grupo-cise.com',
        'bss.bancos@grupo-cise.com',
        'etg.bancos@grupo-cise.com',
        'jald.bancos@grupo-cise.com',
        'kmr.cobranza@grupo-cise.com',
        'lums.cobranza@grupo-cise.com',
        'aarv.contabilidad@grupo-cise.com',
        'asdg.contabilidad@grupo-cise.com',
        'bas.contabilidad@grupo-cise.com',
        'egc.contabilidad@grupo-cise.com',
        'fanl.contabilidad@grupo-cise.com',
        'gga.contabilidad@grupo-cise.com',
        'hamm.contabilidad@grupo-cise.com',
        'imm.contabilidad@grupo-cise.com',
        'jdjlp.contabilidad@grupo-cise.com',
        'jcrl.contabilidad@grupo-cise.com',
        'mgv.contabilidad@grupo-cise.com',
        'rrf.contabilidad@grupo-cise.com',
        'var.contabilidad@grupo-cise.com',
        'jodz.domicilios@grupo-cise.com',
        'mfco.domicilios@grupo-cise.com',
        'mgc.domicilios@grupo-cise.com',
        'tmc.tesoreria@grupo-cise.com',
        'agv.facturacion@grupo-cise.com',
        'aact.facturacion@grupo-cise.com',
        'anco.facturacion@grupo-cise.com',
        'bfg.facturacion@grupo-cise.com',
        'dsmm.facturacion@grupo-cise.com',
        'deev.facturacion@grupo-cise.com',
        'ers.facturacion@grupo-cise.com',
        'era.facturacion@grupo-cise.com',
        'feg.facturacion@grupo-cise.com',
        'gjar.facturacion@grupo-cise.com',
        'has.facturacion@grupo-cise.com',
        'jimt.facturacion@grupo-cise.com',
        'jaag.facturacion@grupo-cise.com',
        'jjar.facturacion@grupo-cise.com',
        'jam.facturacion@grupo-cise.com',
        'rrs.facturacion@grupo-cise.com',
        'rlt.facturacion@grupo-cise.com',
        'srf.facturacion@grupo-cise.com',
        'xrf.facturacion@grupo-cise.com',
        'amvt.juridico@grupo-cise.com',
        'lmsi.juridico@grupo-cise.com',
        'ccl.logistica@grupo-cise.com',
        'egts.logistica@grupo-cise.com',
        'jlsc.logistica@grupo-cise.com',
        'amf.operacion@grupo-cise.com',
        'asre.operacion@grupo-cise.com',
        'effc.operacion@grupo-cise.com',
        'ercv.operacion@grupo-cise.com',
        'ears.operacion@grupo-cise.com',
        'fsdlr.operacion@grupo-cise.com',
        'fcl.operacion@grupo-cise.com',
        'hav.operacion@grupo-cise.com',
        'iarc.operacion@grupo-cise.com',
        'jepl.operacion@grupo-cise.com',
        'jgr.operacion@grupo-cise.com',
        'ljss.operacion@grupo-cise.com',
        'lvaa.operacion@grupo-cise.com',
        'mott.operacion@grupo-cise.com',
        'marb.operacion@grupo-cise.com',
        'mfrr.operacion@grupo-cise.com',
        'nvsm.operacion@grupo-cise.com',
        'rsm.operacion@grupo-cise.com',
        'revl.operacion@grupo-cise.com',
        'rvb.operacion@grupo-cise.com',
        'sevg.operacion@grupo-cise.com',
        'smgb.operacion@grupo-cise.com',
        'uffc.operacion@grupo-cise.com',
        'vrr.operacion@grupo-cise.com',
        'ylrc.operacion@grupo-cise.com',
        'dahs.personal_domestico@grupo-cise.com',
        'esb.personal_domestico@grupo-cise.com',
        'cnrr.presupuestos@grupo-cise.com',
        'flmr.presupuestos@grupo-cise.com',
        'fdpg.presupuestos@grupo-cise.com',
        'niet.presupuestos@grupo-cise.com',
        'nnla.presupuestos@grupo-cise.com',
        'vavg.presupuestos@grupo-cise.com',
        'cglg.rrhh@grupo-cise.com',
        'jovs.sistemas@grupo-cise.com',
        'aavv.tesoreria@grupo-cise.com',
        'cggg.tesoreria@grupo-cise.com',
        'sss.tesoreria@grupo-cise.com',
        'mdv.tesoreria@grupo-cise.com',
        'yvv.tesoreria@grupo-cise.com',
        'jllc.tesoreria@grupo-cise.com',
        'arr.verificacion@grupo-cise.com',
        'dngt.verificacion@grupo-cise.com',
        'ftmg.verificacion@grupo-cise.com',
        'jlg.verificacion@grupo-cise.com',
        'lgpb.verificacion@grupo-cise.com',
        'mem.verificacion@grupo-cise.com',
        'mavr.verificacion@grupo-cise.com',
        'nyjm.verificacion@grupo-cise.com',
        'psb.verificacion@grupo-cise.com'
    ],
    propietarios: [//con CEOS
      'ft.proyectos@grupo-cise.com',
      'gs.proyectos@grupo-cise.com',
      'bs.proyectos@grupo-cise.com',
      'jlmv.proyectos@grupo-cise.com',
      'jb.proyectos@grupo-cise.com',
      'ap.proyectos@grupo-cise.com',
      'ma.proyectos@grupo-cise.com',
      'sb.proyectos@grupo-cise.com',
      'aat.direccion_gral@grupo-cise.com',
      'reportes@kabzo.org',
      'abbydobbleb.99@gmail.com',
      'dirgeneral@kubicspaces.com',
      'grupoviaya@gmail.com',
      'lars.ceos@grupo-cise.com',
      'ycl.ceos@grupo-cise.com',
      'yetp.ceos@grupo-cise.com'
    ]
  }

};

function desbloquearTodo() {

  const nombrehojadeseada = "CALIFICACION TRIMESTRAL 2026";
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombrehojadeseada);

  if (!hoja) {
    Logger.log("La hoja no existe");
    return;
  }

  const proteccionesRango = hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  const proteccionesHoja = hoja.getProtections(SpreadsheetApp.ProtectionType.SHEET);

  // eliminar protecciones de rango
  proteccionesRango.forEach(p => p.remove());

  // eliminar protecciones de hoja
  proteccionesHoja.forEach(p => p.remove());

  Logger.log("Todas las protecciones fueron eliminadas.");
}

function ejecutarBloqueo(config) {

  const nombrehojadeseada = "CALIFICACION TRIMESTRAL 2026";
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombrehojadeseada);


  if (!hoja) {
    Logger.log(`La hoja '${nombrehojadeseada}' no existe.`);
    return;
  }

  config.rangos.forEach(r => {

    const { filaInicio, columnaInicio, columnaFin } = r;

    const ultimaFila = hoja.getLastRow();

    const rango = hoja.getRange(
      filaInicio,
      columnaInicio,
      ultimaFila - filaInicio + 1,
      columnaFin - columnaInicio + 1
    );

    const proteccion = rango.protect().setDescription("🚫 Bloqueo automático");

    proteccion.addEditors(config.propietarios);

    const editores = proteccion.getEditors();

    editores.forEach(user => {
      if (config.usuariosBloqueados.includes(user.getEmail())) {
        proteccion.removeEditor(user);
      }
    });

  });
}

function corridaBloqueos() {
  desbloquearTodo();
  ejecutarBloqueo(CONFIG_BLOQUEO.BLOQUEO_AIP);
  ejecutarBloqueo(CONFIG_BLOQUEO.BLOQUEO_JNP);
}



///calidficacion a foda
function copiadoFoda(){
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var libroOrigen = libro.getSheetByName("CALIFICACION TRIMESTRAL 2026");
  var libroDestino = libro.getSheetByName("FODA 2026");
 

  //col. F subArea 
  // columnas G + H + I + P
  //E-p -> e=5 , p = 16 fila =3

  //var ultimaFilaP = ultimaFilaNoVacia(libroOrigen, "P:P");
  
  var ultimaFilaP = ultimaFilaNoVacia(libroOrigen, "U:U");//seria p -> u

  var ultimaFilaE = ultimaFilaNoVacia(libroDestino, "E:E");

  //var rango = libroOrigen.getRange(3, 5, ultimaFilaP, 16).getValues();
  var rango = libroOrigen.getRange(3, 6, ultimaFilaP, 16).getValues();
  //var rango = libroOrigen.getRange("E5:P864").getValues();

  var filasPegar = [];
  var columnaEliminarP = []; //que elimine solo la colmana P 

  for(var i=0; i<rango.length; i++){
    if(rango[i][15] !== ""){
      filasPegar.push(rango[i]); //se pegan los que tenga en la columna P
      columnaEliminarP.push(rango[i]); //para borrar lo de esa columna P.
    }
  }
  ///new
  if (filasPegar.length > 0) {

    var hojaDestino = libroDestino;
    var filaInicio = ultimaFilaE + 1;

    var colE = [];
    var colP = [];

    filasPegar.forEach(function(fila){

      // columna E
      colE.push([fila[0]]);

      // texto para columna P
      var texto = fila[1] + " " + fila[2] + " " + fila[3] + " " + fila[15];

      var rich = SpreadsheetApp.newRichTextValue()
        .setText(texto)
        .setTextStyle(
          texto.length - fila[15].length,
          texto.length,
          SpreadsheetApp.newTextStyle()
            .setBold(true) // ← negritas
            .setForegroundColor("black")
            .build()
        )
        .build();

      colP.push([rich]);

    });

    // pegar en columna E
    hojaDestino.getRange(filaInicio, 4, colE.length, 1).setValues(colE);

    // pegar en columna P
    hojaDestino.getRange(filaInicio, 5, colP.length, 1).setRichTextValues(colP);

  }
    

  //borrar la columna P posicion = 11
  if (columnaEliminarP.length > 0){
    for (var i = 0; i < rango.length; i++) {

      if (rango[i][20] !== "") { // si tenía dato en U
        libroOrigen.getRange(i + 3, 21).clearContent(); 
      }

    }
  }
}

//ultima fila
function ultimaFilaNoVacia(hoja, rango) {
  if (!hoja) {
    Logger.log("La hoja "+ hoja + " no existe.");
    return;
  }
  
  const columna = hoja.getRange(rango).getValues(); // Obtiene todos los valores de la columna B
  let ultimaFila = 0;

  // Iterar desde el final hacia arriba para encontrar la última fila con datos
  for (let i = columna.length - 1; i >= 0; i--) {
    if (columna[i][0] !== "") {
      ultimaFila = i + 1; // +1 porque los índices comienzan en 0
      break;
    }
  }

  return ultimaFila;
}
