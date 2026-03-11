function onEdit(e) {

  const hoja = e.range.getSheet();
  const nombre = hoja.getName();

  if (nombre !== "PRODUCTIVIDAD DOMICILIOS" 
    && nombre !== "PRODUCTIVIDAD BANCOS"
    && nombre !== "PRODUCTIVIDAD MANTENIMIENTO") return;

  if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) return;

  const fila = e.range.getRow();
  const col = e.range.getColumn();
  const valor = e.value;

  if (fila === 1 || !valor) return;

  const email = Session.getActiveUser().getEmail();

  if (col === 4) {
    hoja.getRange(fila, 2).setValue(email);
    return;
  }

  if (col === 17 && valor === "TERMINADO") {
    hoja.getRange(fila, 20).setValue(email);
  }

}
