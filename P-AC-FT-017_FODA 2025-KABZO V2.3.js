//envio a los archivos de productivadad
function onOpen() { 
    var ui = SpreadsheetApp.getUi();

    ui.createMenu('📅 | Copiado/Bloqueos')
    .addItem('1. Copiados/Bloqueos| 📄', 'unionMetodos')
    .addToUi();
}

function unionMetodos(){//onOpen
  botonEnvioAProctividad();
  corridaBloqueos();
  copiadoFoda();
}

function botonEnvioAProctividad() { //Ana 06/03/2026
  try {
    var hojasDatos = [ //modificar esto con los nuevos de jonathan.
      //originales
      { link: "1OfPWTbh9o-Bn-gakr8IqW1MiLzsBxcEKHaBxrpI50QU", destino: "PRODUCTIVIDAD_AGENDA" },
      { link: "1OfPWTbh9o-Bn-gakr8IqW1MiLzsBxcEKHaBxrpI50QU", destino: "PRODUCTIVIDAD_PERSONAL_DOMESTICO" },
      { link: "1s9JptP88LXb3jrBgIV34EZ-8tCBaQMdPjy0-zEmPHgc", destino: "PRODUCTIVIDAD_BANCOS" },
      { link: "18ZkE0mLJXH6k6jE-VyQkz9JmVOWPwwINnxuF7yY3wAg", destino: "PRODUCTIVIDAD_COBRANZA" },
      { link: "1ryVvXUmcxfZTfRv95DHW1W1qbU9hMrdtmCZEW99fUtA", destino: "PRODUCTIVIDAD_CONTABILIDAD" },
      { link: "1LKl2IpSsHrckfj4Re51NKkFFAVeSsRduVDElkK1UAx0", destino: "PRODUCTIVIDAD_DOMICILIOS" },
      { link: "11FjRxtVLx0aLhQIPlFoq2ZLWAY0H8AjYChcScnrwmhk", destino: "PRODUCTIVIDAD_FACTURACION" },
      { link: "1g9FRpnBj3Zvd3LjQYKdL0pTZfcYAL7nmJnZGOC9bkQk", destino: "PRODUCTIVIDAD_JURIDICO" },
      { link: "15vwvYNGZ1iJtFSLLVRZoUeN3dX6jXgesRRtWomuQFpI", destino: "PRODUCTIVIDAD_LOGISTICA" },
      { link: "12tdaUfycmasY1DKT4FQMnUfrni0WgLk5BZINgll-3_w", destino: "PRODUCTIVIDAD_OPERACIONES" },
      { link: "1zviMwxVNxU_wY7r0yKBZTl3fo6mDMBfnpAgLbmuEv7Y", destino: "PRODUCTIVIDAD_PRESUPUESTOS" },
      { link: "1zviMwxVNxU_wY7r0yKBZTl3fo6mDMBfnpAgLbmuEv7Y", destino: "PRODUCTIVIDAD_VERIFICACION" },
      { link: "1NXPJQ0vP85hDVAGBxERpaKWvDmtAHKewLA2_B4kjR0g", destino: "PRODUCTIVIDAD_PROYECTOS" },
      { link: "1RbWTBkxyPPC91gBwaIE1HrEaOpIkUpkrn-VqsmUid_A", destino: "PRODUCTIVIDAD_RRHH" },
      { link: "1GKaum0yFCzVIBN4z98SI2wGgg2AFcUjn9KP9ExC336Y", destino: "PRODUCTIVIDAD_TESORERIA" },
      /*{ link: "", destino: "PRODUCTIVIDAD_CEOS" }*///se cancelo porque no debe de ir un proceso.
    ];

    hojasDatos.forEach(function (hoja) {
      try {
        envioInf(hoja);
      } catch (err) {
        Logger.log(`❌ Error procesando ${hoja.destino}: ${err.message}`);
      }
    });

  } catch (e) {
    Logger.log("❌ Error general en ciclicosBoton: " + e.message);
  }
}

function envioInf(hojaInf) {//nuevo temporal y 11 Dir. En Proceso.

  //HOJA ORIGEN FODA
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hojaOri = libro.getSheetByName("FODA 2025"); //obtener la hoja de calculo activa


  //HOJA DESTINO ARCHIVO PROCTIVIDAD
  var libroDestino = SpreadsheetApp.openById(hojaInf.link);
  var hojaDes = libroDestino.getSheetByName(hojaInf.destino);  

  var ultimaFilaDi = obtenerUltimaFilaNoVaciaSolicitudes(hojaDes);
  var nuevaFilaDestino = ultimaFilaDi + 1;

  ////////////////
  //hoja origen
   var ultimaFilaOrigen = obtenerUltimaFilaNoVaciaSolicitudes(hojaOri);

  // Variables para recopilar datos
  var filaInicio = 2; // Fila de inicio para copiar y pegar 
  var ultimafilaOrigen = ultimaFilaOrigen;
  var lastRowDestino = nuevaFilaDestino;

  var rango = hojaOri.getRange(filaInicio, 1, ultimafilaOrigen - filaInicio + 1, 8).getValues(); // Columnas A a H


  // Filtrar las filas donde la columna 7 (G) tenga el valor "NUEVO"
  var filasFiltradas = rango.filter(function(fila) {
    //return fila[6] && fila[6].toString().trim() === "NUEVO" && (!fila[7] || fila[7].toString().trim() === "");//Validacion que col.G venga nuevo y la col.H venga vacia.
    return fila[6] && fila[6].toString().trim() === "NUEVO" && String(fila[7] || "").trim() === "";//Validacion que col.G venga nuevo y la col.H venga vacia.
  });

  // Verificar si hay filas para copiar
  var filas = filasFiltradas.filter(function (fila){

      var area = fila[2];
      var subarea = fila[3];
      var nombreHoja = hojaDes.getName();

      switch (nombreHoja){
        case "PRODUCTIVIDAD_AGENDA":
          return area === "ASISTENCIA EJECUTIVA" && subarea === "ASISTENCIA EJECUTIVA" || 
                 area === "ASISTENCIA EJECUTIVA" && subarea === "PRESUPUESTOS PERSONAL";

        case "PRODUCTIVIDAD_PERSONAL_DOMESTICO":
          return area === "ASISTENCIA EJECUTIVA" && subarea === "PERSONAL DOMESTICO";

        case "PRODUCTIVIDAD_BANCOS":
          return area === "BANCOS" && subarea === "BANCOS";

        case "PRODUCTIVIDAD_COBRANZA":
          return area === "COBRANZA" && subarea === "COBRANZA";

        case "PRODUCTIVIDAD_CONTABILIDAD":
          return area === "CONTABILIDAD" && subarea === "CONTABILIDAD" ||
                 area === "CONTABILIDAD" && subarea === "AFILIACIONES";

        case "PRODUCTIVIDAD_DOMICILIOS":
          return area === "DOMICILIOS" && subarea === "DOMICILIOS";

        case "PRODUCTIVIDAD_FACTURACION":
          return area === "FACTURACIÓN" && subarea === "FACTURACIÓN";

        case "PRODUCTIVIDAD_JURIDICO":
          return area === "JURÍDICO" && subarea === "JURÍDICO";

        case "PRODUCTIVIDAD_LOGISTICA":
          return area === "LOGÍSTICA" && subarea === "LOGÍSTICA";

        case "PRODUCTIVIDAD_OPERACIONES":
          return area === "OPERACIÓN" && subarea === "OPERACIÓN";

        case "PRODUCTIVIDAD_PRESUPUESTOS":
          return area === "PRESUPUESTOS" && subarea === "PRESUPUESTOS";

        case "PRODUCTIVIDAD_PROYECTOS":
          return area === "PROYECTOS" && subarea !== "SISTEMAS";//para que cai todos los demas que no sean sistemas

        case "PRODUCTIVIDAD_SISTEMAS":
          return area === "PROYECTOS" && subarea === "SISTEMAS";

        case "PRODUCTIVIDAD_RRHH":
          return area === "RECURSOS HUMANOS" && subarea === "RECURSOS HUMANOS";

        case "PRODUCTIVIDAD_TESORERIA":
          return area === "TESORERÍA" && subarea === "TESORERÍA";
        
        case "PRODUCTIVIDAD_CEOS":
          return area === "CEOS" && subarea === "CEOS";

        default:
          return false;
      }
  
  });

     // Cambiar el estatus en la hoja de origen a "EN PROCESO"
    filas.forEach(function(filaFiltrada) {
      rango.forEach(function(fila, i) {
        if (fila[4] === filaFiltrada[4]) { // Comparar por el Descripcion único (columna E)
          hojaOri.getRange(filaInicio + i, 8).setValue("Trasladar al archivo de Productividad"); // Actualizar columna 8 (Acción inmediata col. H)
        }
      });
    });
    Logger.log('Se cambio en Foda col. H "Trasladar al archivo de Productividad" ');


    if (filas.length > 0) {//for
      var datosFinales = filas.map(f => {//el vector datosFinales trae un espacio en posicion 0 que no debe de pegar en hoja destino, debe de pegar desde la posicion 1 en adelante
        var filaNueva = new Array(14).fill("");

        filaNueva[1] = f[5]; //col.c
        filaNueva[2] = "INTERNO"; //col.c
        filaNueva[3] = f[2];//col.D
        filaNueva[4] = "FODA";//col.E
        filaNueva[5] = "Actividad bimestral - máx 60 días";//col.F
        filaNueva[13] = f[4];//cl.N
        
        return filaNueva.slice(1); // elimina la posición 0

      });

      //copiado completo.
      hojaDes.getRange(lastRowDestino, 2, datosFinales.length, 13).setValues(datosFinales); 
      //hojaDes.getRange(lastRowDestino, 1, datosFinales.length, 14).setValues(datosFinales); 


      Logger.log('Datos copiados exitosamente a partir de la columna D y columna N');

  } else {
    Logger.log('No hay filas con los estatus especificados en la columna AA para copiar.');
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

/////////////bloqueo de usuarios
const CONFIG_BLOQUEO = {

  BLOQUEO_AIP: {
    rangos: [
      { filaInicio: 3, columnaInicio: 1, columnaFin: 9 },  // A:I
      { filaInicio: 3, columnaInicio: 15, columnaFin: 15 } // O
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

  BLOQUEO_JNP: {   // 👈 SEGUNDA CORRIDA
    rangos: [
      { filaInicio: 3, columnaInicio: 10, columnaFin: 14 }, // J:N
      { filaInicio: 3, columnaInicio: 16, columnaFin: 16 }  // P
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

  const nombrehojadeseada = "CALIFICACION TRIMESTRAL 2025";
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

  const nombrehojadeseada = "CALIFICACION TRIMESTRAL 2025";
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
//////////copiado de hojas a filas del mismo foda
function copiadoFoda(){
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var libroOrigen = libro.getSheetByName("CALIFICACION TRIMESTRAL 2025");
  var libroDestino = libro.getSheetByName("FODA 2025");

  //col. F subArea 
  // columnas G + H + I + P
  //E-p -> e=5 , p = 16 fila =3

  var ultimaFilaP = ultimaFilaNoVacia(libroOrigen, "P:P");

  var ultimaFilaE = ultimaFilaNoVacia(libroDestino, "E:E");

  var rango = libroOrigen.getRange(3, 5, ultimaFilaP, 16).getValues();
  //var rango = libroOrigen.getRange("E5:P864").getValues();

  var filasPegar = [];
  var columnaEliminarP = []; //que elimine solo la colmana P 

  for(var i=0; i<rango.length; i++){
    if(rango[i][11] !== ""){
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
      colE.push([fila[1]]);

      // texto para columna P
      var texto = fila[2] + " " + fila[3] + " " + fila[4] + " " + fila[11];

      var rich = SpreadsheetApp.newRichTextValue()
        .setText(texto)
        .setTextStyle(
          texto.length - fila[11].length,
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

      if (rango[i][11] !== "") { // si tenía dato en P
        libroOrigen.getRange(i + 3, 16).clearContent(); 
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
