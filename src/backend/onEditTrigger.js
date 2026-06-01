function onEdit(e) {
  const range = e.range;
  const hoja = range.getSheet();
  if (hoja.getName() !== NOMBRE_HOJA) return;

  const startRow = range.getRow();
  const endRow = range.getLastRow();
  const startCol = range.getColumn();
  const endCol = range.getLastColumn();

  if (startRow === 1) return;

  for (let fila = startRow; fila <= endRow; fila++) {
    if (startCol <= COLUMNAS.IMPORTE + 1 && endCol >= COLUMNAS.IMPORTE + 1) {
      const importe = hoja.getRange(fila, COLUMNAS.IMPORTE + 1).getValue();
      if (typeof importe === "number" && !isNaN(importe)) {
        actualizarImporteLetra(fila, importe);
      }
    }
    procesarFilaSiCompleta(hoja, fila);
  }
}

function procesarFilaSiCompleta(hoja, fila) {
  const cliente = hoja.getRange(fila, COLUMNAS.CLIENTE + 1).getValue();
  const importe = hoja.getRange(fila, COLUMNAS.IMPORTE + 1).getValue();
  const concepto = hoja.getRange(fila, COLUMNAS.CONCEPTO + 1).getValue();
  const fechaPago = hoja.getRange(fila, COLUMNAS.FECHA_PAGO + 1).getValue();
  const folioActual = hoja.getRange(fila, COLUMNAS.FOLIO + 1).getValue();

  if (folioActual) return;
  if (!cliente || !importe || !concepto || !fechaPago) return;

  const fecha = new Date();
  const fechaActual = Utilities.formatDate(
    fecha,
    Session.getScriptTimeZone(),
    "yyyyMMdd",
  );
  const horaActual = Utilities.formatDate(
    fecha,
    Session.getScriptTimeZone(),
    "HHmmss",
  );
  const random = Math.floor(Math.random() * 900) + 100;

  const folio = `GLR-${fechaActual}${horaActual}-${random}`;
  asignarFolioYFecha(fila, folio, fecha);
}
