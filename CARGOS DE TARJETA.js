function agrandarRangoBloqueoCondicionalV2(hojaOrigen) {
  if(!hojaOrigen){
    Logger.log("la hoja no existe");
  }

  const filaInicio = 4; // Fila inicial del bloqueo
  const columnaInicio = 1; // Columna A
  const columnaFinUltimaFila = 14; // Columna N
  const columnaFin = 11; // Columna k
  //const columnaFin = 15; // Columna O
  const usuariosBloqueados = [
    'analistaprocesos2@kabzo.org',
    'auditorinterno@kabzo.org',
    'auxproyectos@kabzo.org'
  ];

  //A-O
  // Eliminar protecciones existentes en el rango A-AA
  const protecciones = hojaOrigen.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  protecciones.forEach(function (proteccion) {
    const rango = proteccion.getRange();
    if (rango.getColumn() === columnaInicio && rango.getLastColumn() === columnaFin && rango.getRow() >= filaInicio) {
      proteccion.remove();
      Logger.log(`Protección eliminada: ${rango.getA1Notation()}`);
    }
  });

  // Obtener la última fila con valores en la columna AA
  const ultimaFila = obtenerUltimaFilaConValoresEnColumnaV002(hojaOrigen, columnaFinUltimaFila, filaInicio);

  // Verificar si la columna Z está completamente vacía
  if (ultimaFila === filaInicio - 1) { // Si no hay datos, ultimaFila regresa filaInicio - 1
    Logger.log(`La columna AA está vacía desde la fila ${filaInicio}. No se realizará el bloqueo.`);
    return;
  }

  // Definir el rango que se bloqueará
  const rangoActual = hojaOrigen.getRange(filaInicio, columnaInicio, ultimaFila - filaInicio + 1, columnaFin);

  // Crear una nueva protección para el rango
  const nuevaProteccion = rangoActual.protect().setDescription('🚫 | SOLO PROPIETARIO');

  //const propietarioEmail = "optimizacion@kabzo.org";
  const propietarioEmail = "ma.proyectos@grupo-cise.com";

  // Bloquear usuarios específicos
  nuevaProteccion.getEditors().forEach(function (editor) {
    const editorEmail = editor.getEmail();
    if (usuariosBloqueados.includes(editorEmail) || editorEmail !== propietarioEmail) {
      nuevaProteccion.removeEditor(editor);
    }
  });

  // Asegurarse de que el propietario tenga acceso
  nuevaProteccion.addEditor(propietarioEmail);

  if (nuevaProteccion.canDomainEdit()) {
    nuevaProteccion.setDomainEdit(false);
  }

  Logger.log(`Nueva protección creada en el rango: ${rangoActual.getA1Notation()}`);
}

// Función para obtener la última fila con datos en una columna específica
function obtenerUltimaFilaConValoresEnColumnaV002(hoja, columna, filaInicio) {
  const datos = hoja.getRange(filaInicio, columna, hoja.getLastRow() - filaInicio + 1).getValues();

  for (let i = datos.length - 1; i >= 0; i--) {
    if (datos[i][0] !== "") { // Verifica si hay datos en la celda
      return filaInicio + i;
    }
  }

  return filaInicio - 1; // Retorna filaInicio - 1 si no hay datos en la columna
}
