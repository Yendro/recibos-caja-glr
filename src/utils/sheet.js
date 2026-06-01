/**
 * Devuelve un arreglo con los objetos de las hojas válidas de recibos y su ID de plantilla.
 */
function obtenerHojasDeRecibos() {
  const config = obtenerConfiguracion();
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const hojas = [];

  config.plantillas.forEach((p) => {
    const hoja = spreadSheet.getSheetByName(p.nombreHoja);
    if (hoja)
      hojas.push({
        hoja: hoja,
        idPlantilla: p.idPlantilla,
        nombre: p.nombreHoja,
      });
  });

  return hojas;
}

function actualizarImporteLetra(fila, importe, nombreHoja) {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = spreadSheet.getSheetByName(nombreHoja);
  const letras = numeroALetras(importe);
  hoja.getRange(fila, COLUMNAS.IMPORTE_LETRA + 1).setValue(letras);
}

function obtenerFilasPendientes() {
  const hojasData = obtenerHojasDeRecibos();
  const pendientes = [];

  // Iteramos sobre todas las hojas de recibos generadas
  hojasData.forEach((data) => {
    const hoja = data.hoja;
    const datos = hoja.getDataRange().getValues();

    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      if (fila[COLUMNAS.STATUS] === "Pendiente") {
        pendientes.push({
          rowIndex: i + 1,
          nombreHoja: data.nombre,
          idPlantilla: data.idPlantilla, // Lo pasamos para saber qué plantilla clonar
          cliente: fila[COLUMNAS.CLIENTE] || "",
          importe: fila[COLUMNAS.IMPORTE] || 0,
          importeLetra: fila[COLUMNAS.IMPORTE_LETRA] || "",
          concepto: fila[COLUMNAS.CONCEPTO] || "",
          fechaPago: fila[COLUMNAS.FECHA_PAGO],
          folio: fila[COLUMNAS.FOLIO] || "",
          fechaCreacion: fila[COLUMNAS.FECHA_CREACION],
          archivo: fila[COLUMNAS.ARCHIVO] || "",
        });
      }
    }
  });
  return pendientes;
}

function actualizarEstadoArchivo(rowIndex, status, urlDoc = null, nombreHoja) {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = spreadSheet.getSheetByName(nombreHoja);
  if (status) {
    hoja.getRange(rowIndex, COLUMNAS.STATUS + 1).setValue(status);
  }
  if (urlDoc) {
    hoja.getRange(rowIndex, COLUMNAS.ARCHIVO + 1).setValue(urlDoc);
  }
}

function asignarFolioYFecha(rowIndex, folio, fechaCreacion, nombreHoja) {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = spreadSheet.getSheetByName(nombreHoja);

  hoja.getRange(rowIndex, COLUMNAS.FOLIO + 1).setValue(folio);
  hoja.getRange(rowIndex, COLUMNAS.FECHA_CREACION + 1).setValue(fechaCreacion);
  hoja.getRange(rowIndex, COLUMNAS.ESTADO_FIRMA + 1).setValue("PENDIENTE");
  hoja.getRange(rowIndex, COLUMNAS.ESTADO_CORREO + 1).setValue("PENDIENTE");
  hoja.getRange(rowIndex, COLUMNAS.ESTATUS_FOTO + 1).setValue("PENDIENTE");
  hoja.getRange(rowIndex, COLUMNAS.STATUS + 1).setValue("PENDIENTE");
}
