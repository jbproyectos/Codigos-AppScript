function doPost(e) {//numero sucesor del 10R del folio de gastos negativo
  
    var props = PropertiesService.getScriptProperties();
    var consecutivo = Number(props.getProperty('CONSECUTIVO_FOLIO')) || 4111;//VV300126 00004111
    consecutivo++;

    props.setProperty('CONSECUTIVO_FOLIO', consecutivo);

    return ContentService
      .createTextOutput(String(consecutivo))
      .setMimeType(ContentService.MimeType.TEXT);

}
