function obtenerHoja() {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = spreadSheet.getSheetByName(NOMBRE_HOJA);
  if (!hoja) {
    hoja = spreadSheet.insertSheet(NOMBRE_HOJA);
    const nombresColumnas = [
      "Cliente",
      "Importe",
      "Importe en Letra",
      "Concepto",
      "FechaPago",
      "EstadoFirma",
      "Folio",
      "FechaCreacion",
      "Status",
      "Archivo",
      "EstadoCorreo",
      "Destinatarios",
    ];
    hoja.getRange(1, 1, 1, nombresColumnas.length).setValues([nombresColumnas]);
    hoja.setFrozenRows(1);
  }
  return hoja;
}

function actualizarImporteLetra(fila, importe) {
  const hoja = obtenerHoja();
  const letras = numeroALetras(importe);
  hoja.getRange(fila, COLUMNAS.IMPORTE_LETRA + 1).setValue(letras);
}

function obtenerFilasPendientes() {
  const hoja = obtenerHoja();
  const datos = hoja.getDataRange().getValues();
  const pendientes = [];

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    if (fila[COLUMNAS.STATUS] === "Pendiente") {
      pendientes.push({
        rowIndex: i + 1,
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
  return pendientes;
}

function actualizarEstadoArchivo(rowIndex, status, urlDoc = null) {
  const hoja = obtenerHoja();
  if (status) {
    hoja.getRange(rowIndex, COLUMNAS.STATUS + 1).setValue(status);
  }
  if (urlDoc) {
    hoja.getRange(rowIndex, COLUMNAS.ARCHIVO + 1).setValue(urlDoc);
  }
}

function asignarFolioYFecha(rowIndex, folio, fechaCreacion) {
  const hoja = obtenerHoja();
  hoja.getRange(rowIndex, COLUMNAS.FOLIO + 1).setValue(folio);
  hoja.getRange(rowIndex, COLUMNAS.FECHA_CREACION + 1).setValue(fechaCreacion);
  hoja.getRange(rowIndex, COLUMNAS.ESTADO_FIRMA + 1).setValue("No firmado");
  hoja.getRange(rowIndex, COLUMNAS.ESTADO_CORREO + 1).setValue("No enviado");
  hoja.getRange(rowIndex, COLUMNAS.STATUS + 1).setValue("Pendiente");
}
