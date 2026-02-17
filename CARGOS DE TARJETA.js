function agrandarRangoBloqueoCondicionalV3(hoja) {//A hasta AB
  //const nombrehojadeseada = "CARGOS";
  //const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombrehojadeseada);

  if (!hoja) {
    Logger.log(`La hoja '${nombrehojadeseada}' no existe.`);
    return;
  }

  const filaInicio = 4; // Fila inicial del bloqueo
  const columnaInicio = 1; // Columna A
  const columnaFinUltimaFila = 14; // Columna N
  const columnaFin = 11; // Columna k
  const usuariosBloqueados = [
    'bs.proyectos@grupo-cise.com',
    'administracion.empresarial@rhfinder.com',
    'gs.proyectos@grupo-cise.com',
    'cnrr.presupuestos@grupo-cise.com',
    'dngt.verificacion@grupo-cise.com',
    'dlav.agenda_ejecutiva@grupo-cise.com',
    'yetp.ceos@grupo-cise.com',
    'esb.personal_domestico@grupo-cise.com',
    'flmr.presupuestos@grupo-cise.com',
    'ftmg.verificacion@grupo-cise.com',
    'jlg.verificacion@grupo-cise.com',
    'lgpb.verificacion@grupo-cise.com',
    'mem.verificacion@grupo-cise.com',
    'mavr.verificacion@grupo-cise.com',
    'niet.presupuestos@grupo-cise.com',
    'nnla.presupuestos@grupo-cise.com',
    'nyjm.verificacion@grupo-cise.com',
    'psb.verificacion@grupo-cise.com',
    'reportes@kabzo.org',
    'verificador9@consultoriavrf.com'
  ];

  // Eliminar protecciones existentes en el rango A-AB
  const protecciones = hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  protecciones.forEach(function (proteccion) {
    const rango = proteccion.getRange();
    if (rango.getColumn() === columnaInicio && rango.getLastColumn() === columnaFin && rango.getRow() >= filaInicio) {
      proteccion.remove();
      Logger.log(`Protección eliminada: ${rango.getA1Notation()}`);
    }
  });

  // Obtener la última fila con valores en la columna Z
  const ultimaFila = obtenerUltimaFilaConValoresEnColumnaV2(hoja, columnaFinUltimaFila, filaInicio);
  //const ultimaFila = obtenerUltimaFilaConValoresEnColumnaV002(hojaOrigen, columnaFinUltimaFila, filaInicio);

  // Verificar si la columna Z está completamente vacía
  if (ultimaFila === filaInicio - 1) { // Si no hay datos, ultimaFila regresa filaInicio - 1
    Logger.log(`La columna AB está vacía desde la fila ${filaInicio}. No se realizará el bloqueo.`);
    return;
  }

  // Definir el rango que se bloqueará
  const rangoActual = hoja.getRange(filaInicio, columnaInicio, ultimaFila - filaInicio + 1, columnaFin);

  // Crear una nueva protección para el rango
  const nuevaProteccion = rangoActual.protect().setDescription('🚫 | SOLO PROPIETARIO');

  // Lista de propietarios que deben mantener permisos de edición
  const propietarioEmail = [
    'ma.proyectos@grupo-cise.com',
    'arr.verificacion@grupo-cise.com',
    'bs.proyectos@grupo-cise.com',
    'gs.proyectos@grupo-cise.com',
    'fdpg.presupuestos@grupo-cise.com',
    'vavg.presupuestos@grupo-cise.com',
    'abbydobbleb.99@gmail.com',
    'sb.proyectos@grupo-cise.com',
    'grupoviaya@gmail.com',
    'ft.proyectos@grupo-cise.com',
    'jlmv.proyectos@grupo-cise.com',
    'jb.proyectos@grupo-cise.com',
    'ap.proyectos@grupo-cise.com',
    'vavg.presupuestos@grupo-cise.com'
  ];

  // Remover editores que no sean propietarios o que estén en la lista de bloqueados
  nuevaProteccion.getEditors().forEach(function (editor) {
    const editorEmail = editor.getEmail();
    if (usuariosBloqueados.includes(editorEmail) || !propietarioEmail.includes(editorEmail)) {
      nuevaProteccion.removeEditor(editor);
    }
  });

  // Asegurar que los propietarios tengan acceso
  propietarioEmail.forEach(email => {
    nuevaProteccion.addEditor(email);
    Logger.log("propietarios "+ email)
  });

  // Deshabilitar edición a todo el dominio si está habilitada
  if (nuevaProteccion.canDomainEdit()) {
    nuevaProteccion.setDomainEdit(false);
  }

  Logger.log(`Nueva protección creada en el rango: ${rangoActual.getA1Notation()}`);
}

// Función para obtener la última fila con datos en una columna específica
function obtenerUltimaFilaConValoresEnColumnaV2(hoja, columna, filaInicio) {
  const datos = hoja.getRange(filaInicio, columna, hoja.getLastRow() - filaInicio + 1).getValues();

  for (let i = datos.length - 1; i >= 0; i--) {
    if (datos[i][0] !== "") { // Verifica si hay datos en la celda
      return filaInicio + i;
    }
  }

  return filaInicio - 1; // Retorna filaInicio - 1 si no hay datos en la columna
}
