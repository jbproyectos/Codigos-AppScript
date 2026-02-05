function doPost(e) {//este es un Apps Script

    var props = PropertiesService.getScriptProperties();
    var consecutivo = Number(props.getProperty('CONSECUTIVO_FOLIO')) || 138137 ;//NL03022600 138137
    consecutivo++;

    props.setProperty('CONSECUTIVO_FOLIO', consecutivo);

    return ContentService
      .createTextOutput(String(consecutivo))
      .setMimeType(ContentService.MimeType.TEXT);

}
