function onEdit(e) {
  if (!e || !e.range) return;
  const range = e.range;
  const hoja = range.getSheet();
  const nombreHojaEditada = hoja.getName();

  // Validar si la hoja editada es una de las hojas de Recibos activas
  const config = obtenerConfiguracion();
  const esHojaRecibos = config.plantillas.some(
    (p) => p.nombreHoja === nombreHojaEditada,
  );

  if (!esHojaRecibos) return;

  const startRow = range.getRow();
  const endRow = range.getLastRow();
  const startCol = range.getColumn();
  const endCol = range.getLastColumn();

  if (startRow === 1) return;

  // Evaluamos si editaron la columna Eliminar y eligieron la opción correcta
  if (startCol === COLUMNAS.ELIMINAR + 1 && e.value === "🗑️ ELIMINAR RECIBO") {
    const urlDoc = hoja.getRange(startRow, COLUMNAS.ARCHIVO + 1).getValue();

    // 1. Mover documento a la papelera de Drive
    if (urlDoc) {
      const match = urlDoc.match(/[-\w]{25,}/);
      if (match) {
        try {
          DriveApp.getFileById(match[0]).setTrashed(true);
        } catch (err) {
          console.error("No se pudo enviar a la papelera: " + err);
        }
      }
    }

    // 2. Eliminar la fila completa en Sheets
    hoja.deleteRow(startRow);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Recibo y documento eliminados correctamente.",
      "🗑️ Eliminado",
    );

    // 3. Detenemos la ejecución aquí (no genera mayúsculas, ni folios)
    return;
  }

  // CONVERTIR A MAYÚSCULAS AUTOMÁTICAMENTE
  // Si el valor ingresado es un texto y NO es una fórmula
  if (e.value && typeof e.value === "string" && !e.value.startsWith("=")) {
    // Solo aplicamos mayúsculas a las columnas de llenado manual (de Cliente a FechaPago)
    if (startCol <= 4) {
      const upperValue = e.value.toUpperCase();
      // Verificamos si el texto cambió para evitar bucles infinitos
      if (e.value !== upperValue) {
        range.setValue(upperValue);
      }
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
