//////////////////////////////

function enviarCorreosParalelo() {
  const contactos = obtenerContactos();
  // const contactos = {
  //   "gs.proyectos@grupo-cise.com": {
  //     nombre: "Gustavo Sanchez",
  //     sheet: "https://docs.google.com/spreadsheets/d/1xZkdgBbxjmRcO2DkUI_Pn0BxrCq_eEm6xgxtGWZqAW0/edit?gid=1338152542#gid=1338152542"
  //   },
  //   "ap.proyectos@grupo-cise.com": {
  //     nombre: "Adriana Perez",
  //     sheet: "https://docs.google.com/spreadsheets/d/11IcA7OSlVT5qnpG8F6PHjhEM5hW5bF3etr4p3Y3v3Ck/edit?gid=0#gid=0"
  //   }
  // };
  const token = ScriptApp.getOAuthToken();
  const url = "https://gmail.googleapis.com/gmail/v1/users/me/messages/send";
  const requests = [];
  for (const email in contactos) {
    const nombre = contactos[email].nombre;
    const linkSheet = contactos[email].sheet;
    const asunto = "CUENTA CORREOS";
    const html = plantillaCorreo(nombre, email, linkSheet);
    const mensaje =
`To: ${email}
Subject: ${asunto}
MIME-Version: 1.0
Content-Type: text/html; charset=UTF-8

${html}`;

    // const raw = Utilities.base64EncodeWebSafe(mensaje);
    const raw = Utilities.base64EncodeWebSafe(
      Utilities.newBlob(mensaje).getBytes()
    );
    requests.push({
      url: url,
      method: "post",
      contentType: "application/json",
      headers: {
        Authorization: "Bearer " + token
      },
      payload: JSON.stringify({
        raw: raw
      }),
      muteHttpExceptions: true
    });
  }
  const responses = UrlFetchApp.fetchAll(requests);

  responses.forEach(r => {
    Logger.log(r.getResponseCode());
    Logger.log(r.getContentText());
  });
}

//////////////////////////////

function obtenerContactos(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(`CONTROL`);
  const data = sheet.getRange("C2:J13").getValues(); // empieza en fila 2 para evitar encabezados
  const contactos = {};
  data.forEach(row => {
    const nombre = row[0]; // columna A
    const email = row[3];  // columna B
    const link = row[7];   // columna C
    if (email) {
      contactos[email] = {
        nombre: nombre,
        sheet: link
      };
    }
  });
  return contactos;
}

//////////////////////////////

function plantillaCorreo(nombre, email, linkSheet){
  return `
  <html>
    <body style="font-family:Arial">
      <h2>Buen dia, ${nombre}! 👋</h2>
      
      <p>
      Entrega de Archivos de Gestión de Correos 📩
      <br>
      Área: Facturación.
      <br>
      <br>

      Este es un mensaje personalizado enviado a:
      <br>
      <b>${email}</b>
      <br>
      </p>
      <p>
      Se les informa que se procederá con la entrega de los archivos individuales por colaborador al área de FACTURACION
      <br>
      Así mismo, les comparto el video tutorial. En él se explica paso a paso cómo activar correctamente su generador de correos. 
      <br>
      Este proceso es obligatorio para cada colaborador de cada área con el motivo de que sus correos diarios se registren y cuantifiquen.
      <br>
      Les recordamos que cada archivo es personal y debe estar vinculado correctamente a su cuenta institucional para ejecutar la cuantificación de sus correos. 
      <br>
      Agradecemos su atención. Estamos al pendiente de cualquier duda o comentario. 
      </p>
      <center>
      <p>
      LINK DEL ARCHIVO:
      <p>
      <center>
      <center>
      <a href="${linkSheet}" 
      style="background:#1a73e8;color:white;padding:10px 16px;text-decoration:none;border-radius:5px;">
      Abrir Archivo de Correos
      </a>
      <center>
      </p>
      </p>
      <p>
      <br>
      VIDEOTUTORIAL: 
      https://drive.google.com/file/d/1Vm32qKM8d1_3lSqJfohkyKsef4yFglzm/view?usp=drive_link
      
      <br>
      AYUDA VISUAL: 
      https://drive.google.com/file/d/1v_z2pUItHVfeojF5d2txYj5gJKeqcGcw/view?usp=drive_link
      </p>

      <hr>
      <small>Mensaje enviado automáticamente</small>
    </body>
  </html>
  `;
}

//////////////////////////////

function pruebaCorreo(){
  GmailApp.sendEmail(
    "gs.proyectos@grupo-cise.com",
    "Prueba",
    "Correo de prueba"
  );
}
