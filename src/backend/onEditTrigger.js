function onEdit(e) {
  if (!e || !e.range) return;
  const range = e.range;
  const hoja = range.getSheet();
  const nombreHojaEditada = hoja.getName();

  // Validar si la hoja editada es operativa
  const config = obtenerConfiguracion();
  const esHojaRecibos = config.plantillas.some(
    (p) => p.nombreHoja === nombreHojaEditada,
  );
  if (!esHojaRecibos) return;

  const startRow = range.getRow();
  const endRow = range.getLastRow();
  const startCol = range.getColumn();
  const endCol = range.getLastColumn();

  if (startRow === 1) return; // Ignorar encabezados

  // CONVERTIR A MAYÚSCULAS AUTOMÁTICAMENTE
  if (e.value && typeof e.value === "string" && !e.value.startsWith("=")) {
    if (startCol <= 4) {
      // Solo columnas manuales
      const upperValue = e.value.toUpperCase();
      if (e.value !== upperValue) range.setValue(upperValue);
    }
  }

  // Obtener la plantilla actual para sacar el prefijo
  const plantillaActual = config.plantillas.find(
    (p) => p.nombreHoja === nombreHojaEditada,
  );

  for (let fila = startRow; fila <= endRow; fila++) {
    if (startCol <= COLUMNAS.IMPORTE + 1 && endCol >= COLUMNAS.IMPORTE + 1) {
      const importe = hoja.getRange(fila, COLUMNAS.IMPORTE + 1).getValue();
      if (typeof importe === "number" && !isNaN(importe)) {
        actualizarImporteLetra(fila, importe, nombreHojaEditada);
      }
    }
    // Pasamos el prefijo como último parámetro
    procesarFilaSiCompleta(
      hoja,
      fila,
      nombreHojaEditada,
      plantillaActual.prefijoFolio,
    );
  }
}

function procesarFilaSiCompleta(hoja, fila, nombreHoja, prefijoFolio) {
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

  // FOLIO DINÁMICO DESDE CONFIGURACIÓN Y POR FILA
  const folio = `${prefijoFolio}-${fechaActual}${horaActual}-F0${fila}`;

  asignarFolioYFecha(fila, folio, fecha, nombreHoja);
}
