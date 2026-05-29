  //  Constantes Globales con MAYUSCULAS
  //  Variables y funciones con camelCase
  //  Archivo: CAPTURA AGUINALDO/IMSS 
const DATOS_SSID = SpreadsheetApp.getActiveSpreadsheet().getId();
const NOM_SHEET = ["SOLICITUD_NOMINA"]; // Hoja en archivos SOLICITUD_NOMINA
const NOM_SSID_01 = [  //  Arreglo con IDs de los archivos de Captura Nomina
  `16UpUYJEK7lUg_shhoe_ViTCLYElDlGaDSda6Yd5ES_E`,// ASISTENCIA_EJECUTIVA
  `1wIrDKJAAVHIoeCoffZ7rcrUy0cCjaBbbJSQBZTFQh4E`,// BANCOS
  `1FSOsSstW1EzEBKWWh3TbVIqmeoesiKaW3jcrS6Nh1Po`,// CONTABILIDAD
  `1_Kfbscpj3x9RbMxM85PDbQecH9U4EkaPUkcKXktsO8Y`,// DOMICILIOS
  `12Hdwv_Xqst8TVE3ZIog-4a5C1H1026CBWasB74jdURk`,// FACTURACIÓN
  `15BVa3ZqDwX3xefU9cEgBBeDZlrDvMcLjg0QDanfmaiQ`,// JURÍDICO
  `1nebC_hYSG-SPX5Re4mLvytolOkbXNZFCZWMGVgi6F6w`,// LOGÍSTICA
];
const NOM_SSID_02 = [  //  Arreglo con IDs de los archivos de Captura Nomina
  `1V2-NUr54kjQawmyjhwWV0_4neh6DggtHJsXTtsGEw_s`,// OPERACIÓN
  `1SA8r_Ul-IJh6aAarQAWJtC1zFXPQKT7oot32pXLJrNs`,// PRESUPUESTOS
  `170jSCHbe6LRG_oJQe0H_Cx8DM_94xNZWFkiskCOWjZE`,// PROYECTOS
  `1NZb9dOdmkf9EWkM7k4EG3exoC6FUK2fWKmlb7pgiRQk`,// RECURSOS_HUMANOS
  `1oM2Hb42_73sCXWSE4m1tbYrqEa6Oq0ewkQefkwZwSWs`,// TESORERÍA
  `1e7c43FrPAIEh_ijgfYLtkGRdZqvo5l9XGkOyRUQrJ6I`,// VERIFICACIÓN
];

// const NOM_SSID = [  //  Arreglo con IDs de los archivos de Captura Nomina
//   `1zc6xunIz8J3B52QVu5sXWYkkxV6VhCq_SKJO_vBgpwM`, // PROYECTOS
//   `1ceN-yAsV3R7XujZ6yHBHt_AIHpOGd6C2G2MeNeWBqmE`, // AGENDA_EJECUTIVA
//   `1peUhkz4vu2MYSKHLQUEHDWt5SnkBbKnYoS7rqvLHh8g`, // TESORERIA
//   `1puzijIGnQyu5jFegnljOBB5Vlxuw8cm4zRZRi9FdTNo`, // VERIFICACION
//   `1TWO5okVpZ2b1_qnmPnbx1orlUZ54fncx-5PigqwT6Wc`, // PRESUPUESTOS
//   `1cYsMyp_5cumH-Uw1Yu0WD8iUrDENMvjvqnXnEEmKog8`, // RRHH
//   `1mcPLIjPpgVqqx3V81Z9jQyPqm_3Tlv1DRIU8Vu2YeZA`, // JURIDICO
//   `1PTJ3YqL8zyHR2G1SBEuUmUrDmJiHII-Pamshgv5rsWU`, // CONTABILIDAD
//   `1U_kaB5pqWy1GHEEfJLmSBxcaTUzix1k79YumjE-0trM`, // DOMICILIOS
//   `1btNYC0d-fE24_BecXBYZeNHFfQHQ86ZrqGfMBSAiEng`, // LOGISTICA
//   `1Abimu3FGoFeJY4pHAxRsBWe974266DEKM8B5cVfskME`, // BANCOS
//   `1e4fiO-H-tJL68up_FLPOCVU0PD_L7w0MJyQPog9VLLk`, // FACTURACION
//   `14ihddZEAUzAx8srGEE8ap65RlT_sBhf0BTrZpfUZFzk`, // OPERACION
//   // `1V2fSfM3j6sAYq3qan81HhujjivEF5MVpvgT4QnNDT6s`, // COBRANZA
// ];
  
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu(`⭕ REBAJES/IMSS`)
    // .addItem(`► Mandar Información a Nóminas`, `mandarInfoNominas`)
    .addItem(`► Borrar Información`, `eraseData`)
    .addToUi();
}

function mandarInfoNominas01(){
    //  Intentar abrir para insertar datos en lista de archivos
  NOM_SSID_01.forEach(archNomID => {
    try{
    var ssNom = SpreadsheetApp.openById(archNomID);
    var sheetDatos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(`CAPTURA`);
    NOM_SHEET.forEach(nombreSheet => {
      try{
      var sheetNom = ssNom.getSheetByName(nombreSheet);

      var datosCompletos = sheetDatos.getRange("A6:J").getDisplayValues(); // Nombres empiezan en A3 (FILAS PENDIENTES)
      var nombresNom = sheetNom.getRange("K1002:K1100").getValues().
      filter(fila => fila[0] != "" && fila[0] != null);
      datos = datosCompletos.filter(fila =>
        fila[0] != "" && fila[0] != null).map(fila => 
          [fila[0], fila[6], fila[7], fila[8], fila[9]]);

      var mapaDatos = {};
      datos.forEach(fila => {
        mapaDatos[fila[0]] = fila;
      });

      var datosOrdenados = nombresNom.map(n => mapaDatos[n[0]] || [n[0], ""]);
      Logger.log(`Datos: ${datosOrdenados}`);
      var imss = datosOrdenados.map(fila => [fila[1]]);
      var aguinaldo = datosOrdenados.map(fila => [fila[2]]);
      var infonavit = datosOrdenados.map(fila => [fila[3]]);
      var personal = datosOrdenados.map(fila => [fila[4]]);
      personal = personal.map(fila => {
        (fila[0]==""||fila[0]==null)?fila[0]=0:0;
        return fila;
      });
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
      sheetNom.getRange(1902,13,personal.length,1).setValues(personal);
      // Logger.log(JSON.stringify(imss));
      }catch(error){
        Logger.log(`No se encontro la hoja ${nombreSheet} en archivo ${ssNom.getName()}`)
        GmailApp.sendEmail(
          'gs.proyectos@grupo-cise.com,sb.proyectos@grupo-cise.com,fdpg.presupuestos@grupo-cise.com,cnrr.presupuestos@grupo-cise.com',
          '⚠️ Error',
          'Advertencia',
          {
            htmlBody: generarCorreoError(
              'mandarInfoNominas',
              error,
              ssNom.getName(),
              nombreSheet
            )
          }
        );
      }
    })
    } catch (error) {
      Logger.log(`Error con archivo ${archNomID}: ${error.message}`);
      // enviarCorreo(archNomID);
    }
  });
}

function mandarInfoNominas02(){
    //  Intentar abrir para insertar datos en lista de archivos
  NOM_SSID_02.forEach(archNomID => {
    try{
    var ssNom = SpreadsheetApp.openById(archNomID);
    var sheetDatos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(`CAPTURA`);
    NOM_SHEET.forEach(nombreSheet => {
      try{
      var sheetNom = ssNom.getSheetByName(nombreSheet);

      var datosCompletos = sheetDatos.getRange("A6:J").getDisplayValues(); // Nombres empiezan en A3 (FILAS PENDIENTES)
      var nombresNom = sheetNom.getRange("K1002:K1100").getValues().
      filter(fila => fila[0] != "" && fila[0] != null);
      datos = datosCompletos.filter(fila =>
        fila[0] != "" && fila[0] != null).map(fila => 
          [fila[0], fila[6], fila[7], fila[8], fila[9]]);

      var mapaDatos = {};
      datos.forEach(fila => {
        mapaDatos[fila[0]] = fila;
      });

      var datosOrdenados = nombresNom.map(n => mapaDatos[n[0]] || [n[0], ""]);
      Logger.log(`Datos: ${datosOrdenados}`);
      var imss = datosOrdenados.map(fila => [fila[1]]);
      var aguinaldo = datosOrdenados.map(fila => [fila[2]]);
      var infonavit = datosOrdenados.map(fila => [fila[3]]);
      var personal = datosOrdenados.map(fila => [fila[4]]);
      personal = personal.map(fila => {
        (fila[0]==""||fila[0]==null)?fila[0]=0:0;
        return fila;
      });
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
      sheetNom.getRange(1902,13,personal.length,1).setValues(personal);
      // Logger.log(JSON.stringify(imss));
      }catch(error){
        Logger.log(`No se encontro la hoja ${nombreSheet} en archivo ${ssNom.getName()}`)
        GmailApp.sendEmail(
          'gs.proyectos@grupo-cise.com,sb.proyectos@grupo-cise.com,fdpg.presupuestos@grupo-cise.com,cnrr.presupuestos@grupo-cise.com',
          '⚠️ Error',
          'Advertencia',
          {
            htmlBody: generarCorreoError(
              'mandarInfoNominas',
              error,
              ssNom.getName(),
              nombreSheet
            )
          }
        );
      }
    })
    } catch (error) {
      Logger.log(`Error con archivo ${archNomID}: ${error.message}`);
      // enviarCorreo(archNomID);
    }
  });
}

function eraseData() {
  SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(`DATOS.`)
    .getRange(`E6:H`)
    .clearContent();
}



function generarCorreoError(funcion, error, archivo = '', hoja = '') {
  return `
  <!DOCTYPE html>
  <html>
  <body style="font-family:Arial,sans-serif;background:#f5f5f5;padding:20px">
    <div style="max-width:700px;margin:auto;background:#fff">
      <div style="background:#d32f2f;color:#fff;padding:15px">
        <h2>⚠️ Error en Automatización</h2>
      </div>
      <div style="padding:20px">
        <p>
          Se detectó un error durante la ejecución de una macro.
        </p>
        <table>
          <tr>
            <td><b>Función:</b></td>
            <td>${funcion}</td>
          </tr>
          <tr>
            <td><b>Fecha:</b></td>
            <td>${new Date()}</td>
          </tr>
          <tr>
            <td><b>Archivo:</b></td>
            <td>${archivo}</td>
          </tr>
          <tr>
            <td><b>Hoja:</b></td>
            <td>${hoja}</td>
          </tr>
        </table>
        <h3 style="color:#d32f2f">Mensaje de Error</h3>
        <div style="
          background:#fff3f3;
          padding:15px;
          border-left:4px solid #d32f2f;
        ">
          ${error.message}
        </div>
        <h3>Stack Trace</h3>
        <pre style="
          background:#f5f5f5;
          padding:15px;
          overflow:auto;
        ">${error.stack || 'No disponible'}</pre>
      </div>
    </div>
  </body>
  </html>
  `;
}
