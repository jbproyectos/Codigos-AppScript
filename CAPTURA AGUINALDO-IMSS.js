  //  Constantes Globales con MAYUSCULAS
  //  Variables y funciones con camelCase
  
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu(`⭕ REBAJES/IMSS`)
    .addItem(`► Mandar IMSS a Nominas`, `mandarInfoNominas`)
    .addToUi();
}

  //  Archivo: CAPTURA AGUINALDO/IMSS 
const DATOS_SSID = SpreadsheetApp.getActiveSpreadsheet().getId();

const NOM_SSID = [  //  Arreglo con IDs de los archivos de Captura Nomina
  `1zc6xunIz8J3B52QVu5sXWYkkxV6VhCq_SKJO_vBgpwM`, // DIR SARAI
  `1TWO5okVpZ2b1_qnmPnbx1orlUZ54fncx-5PigqwT6Wc`, // DIR AZAEL
  `14ihddZEAUzAx8srGEE8ap65RlT_sBhf0BTrZpfUZFzk`, // DIR YESSICA
  `19CKv9L-zCCy1FqW2oCHX5tWyoOGmHBG8ST5HdQJH1LE`, // LINEA CEO ARMANDO
  `1btNYC0d-fE24_BecXBYZeNHFfQHQ86ZrqGfMBSAiEng`, // DIR CARLOS
  `1ceN-yAsV3R7XujZ6yHBHt_AIHpOGd6C2G2MeNeWBqmE`  // DIR ANGIE
];

  //  Nombres de las hojas
const DATOS_SHEET = "DATOS."; //  Hoja en archivo Captura Aguinaldos/Imss
const NOM_SHEET = [  // Hojas en archivos Captura Nominas
  "Formato Nomina Ejemplo",
  "Formato Nomina Ejemplo RRHH",
  "Formato Nomina Ejemplo Contabilidad",
  "Formato Nomina Ejemplo Facturacion",
  "Formato Nomina Ejemplo Juridico",
  "Formato Nomina Ejemplo Domicilios",
  "Formato Nomina Ejemplo Bancos",
  "Formato Nomina Ejemplo Cobranza"
];

function mandarInfoNominas(){
    //  Intentar abrir para insertar datos en lista de archivos
  NOM_SSID.forEach(archNomID => {
    try{
    var ssDatos = SpreadsheetApp.openById(DATOS_SSID);
    var ssNom = SpreadsheetApp.openById(archNomID);
    var sheetDatos = ssDatos.getSheetByName(DATOS_SHEET);
    NOM_SHEET.forEach(nombreSheet => {
      try{
      var sheetNom = ssNom.getSheetByName(nombreSheet);

      var datosCompletos = sheetDatos.getRange("A4:U").getValues(); // Nombres empiezan en A4
      var nombresNom = sheetNom.getRange("K1002:K1100").getValues().
      filter(fila => fila[0] != "" && fila[0] != null);
      datos = datosCompletos.filter(fila => 
        fila[0] != "" && fila[0] != null).map(fila => 
          [fila[0], fila[16], fila[17], fila[20]]);

      var mapaDatos = {};
      datos.forEach(fila => {
        mapaDatos[fila[0]] = fila;
      });

      var datosOrdenados = nombresNom.map(n => mapaDatos[n[0]] || [n[0], ""]);
      var imss = datosOrdenados.map(fila => [fila[1]]);
      var aguinaldo = datosOrdenados.map(fila => [fila[2]]);
      var infonavit = datosOrdenados.map(fila => [fila[3]]);
      infonavit = infonavit.map(fila => {
        (fila[0]==""||fila[0]==null)?fila[0]=0:0;
        return fila;
      });
      aguinaldo = aguinaldo.map(fila => {
        (fila[0]==""||fila[0]==null)?fila[0]=0:0;
        return fila;
      });
      imss = imss.map(fila => {
        (fila[0]==""||fila[0]==null)?fila[0]=0:0;
        return fila;
      });

      // SpreadsheetApp.getUi().alert(`IMSS: ${imss}
      // AGUINALDO: ${aguinaldo}`);
      sheetNom.getRange(1602,13,imss.length,1).setValues(imss);
      sheetNom.getRange(1702,13,aguinaldo.length,1).setValues(aguinaldo);
      sheetNom.getRange(1802,13,infonavit.length,1).setValues(infonavit);
      // Logger.log(JSON.stringify(imss));
      }catch{
        Logger.log(`No se encontro la hoja ${nombreSheet} en archivo ${ssNom.getName()}`)
      }
    })
    } catch (err) {
      Logger.log(`Error con archivo ${archNomID}: ${err.message}`);
    }
  });
}
