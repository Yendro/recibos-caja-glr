// Punto de entrada de la Web App
function doGet(e) {
  const html =
    HtmlService.createTemplateFromFile("src/frontend/index").evaluate();
  html.setTitle("Firma de Recibos GLR");
  // Permite que se embeba si fuera necesario y ajusta la escala para móviles/tablets
  html.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  html.addMetaTag(
    "viewport",
    "width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no",
  );
  return html;
}

function obtenerDatosParaWeb() {
  const config = obtenerConfiguracion(); // Obtenemos correos dinámicos
  const hojasData = obtenerHojasDeRecibos(); // Función que creamos en sheet.js
  const recibos = [];

  // Iteramos sobre TODAS las hojas de recibos
  hojasData.forEach((data) => {
    const hoja = data.hoja;
    const nombreHoja = data.nombre;
    const datos = hoja.getDataRange().getValues();

    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      const status = fila[COLUMNAS.STATUS];
      const estadoFirma = fila[COLUMNAS.ESTADO_FIRMA] || "PENDIENTE";
      const estadoCorreo = fila[COLUMNAS.ESTADO_CORREO] || "PENDIENTE";

      if (
        status === "GENERADO" &&
        !(estadoFirma === "FIRMADO" && estadoCorreo === "ENVIADO")
      ) {
        recibos.push({
          rowIndex: i + 1,
          nombreHoja: nombreHoja, // Le decimos al frontend de qué hoja viene
          cliente: fila[COLUMNAS.CLIENTE] || "",
          importe: fila[COLUMNAS.IMPORTE] || 0,
          concepto: fila[COLUMNAS.CONCEPTO] || "",
          folio: fila[COLUMNAS.FOLIO] || "",
          urlDoc: fila[COLUMNAS.ARCHIVO] || "",
          estadoFirma: estadoFirma,
          estadoCorreo: estadoCorreo,
          estatusFoto: fila[COLUMNAS.ESTATUS_FOTO] || "PENDIENTE",
          destinatarios: fila[COLUMNAS.DESTINATARIOS] || "",
        });
      }
    }
  });

  return {
    recibos: recibos.reverse(),
    correosPredefinidos: config.directorioCorreos, // Mandamos los objetos {nombre, correo}
  };
}

function eliminarRecibosMasivo(listaAEliminar) {
  try {
    const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Enviar documentos a papelera
    listaAEliminar.forEach((r) => {
      if (r.urlDoc) {
        const match = r.urlDoc.match(/[-\w]{25,}/);
        if (match) {
          try {
            DriveApp.getFileById(match[0]).setTrashed(true);
          } catch (e) {
            console.warn("Archivo ya no existe:", e);
          }
        }
      }
    });

    // 2. Agrupar por hoja para borrar en bloque
    const porHoja = {};
    listaAEliminar.forEach((r) => {
      if (!porHoja[r.nombreHoja]) porHoja[r.nombreHoja] = [];
      porHoja[r.nombreHoja].push(r.rowIndex);
    });

    // 3. Borrar filas DE ABAJO HACIA ARRIBA
    for (const hojaNombre in porHoja) {
      const hoja = spreadSheet.getSheetByName(hojaNombre);
      if (hoja) {
        // Orden Descendente vital para no desplazar las filas al borrar
        const filasDescendentes = porHoja[hojaNombre].sort((a, b) => b - a);
        filasDescendentes.forEach((filaIndex) => hoja.deleteRow(filaIndex));
      }
    }

    return { success: true };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}
