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

// Función que llamará el frontend para obtener la lista de trabajo
// Nueva función que carga los recibos filtrados y los correos predefinidos
function obtenerDatosParaWeb() {
  const hoja = obtenerHoja();
  const datos = hoja.getDataRange().getValues();
  const recibos = [];

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const status = fila[COLUMNAS.STATUS];
    const estadoFirma = fila[COLUMNAS.ESTADO_FIRMA] || "No firmado";
    const estadoCorreo = fila[COLUMNAS.ESTADO_CORREO] || "No enviado";

    // LÓGICA: Solo incluir recibos 'Generados' que NO estén completamente terminados (Firmado Y Enviado)
    if (
      status === "Generado" &&
      !(estadoFirma === "Firmado" && estadoCorreo === "Enviado")
    ) {
      recibos.push({
        rowIndex: i + 1,
        cliente: fila[COLUMNAS.CLIENTE] || "",
        importe: fila[COLUMNAS.IMPORTE] || 0,
        concepto: fila[COLUMNAS.CONCEPTO] || "",
        folio: fila[COLUMNAS.FOLIO] || "",
        urlDoc: fila[COLUMNAS.ARCHIVO] || "",
        estadoFirma: estadoFirma,
        estadoCorreo: estadoCorreo,
        destinatarios: fila[COLUMNAS.DESTINATARIOS] || "",
      });
    }
  }

  return {
    recibos: recibos.reverse(),
    correosPredefinidos: CORREOS_PREDEFINIDOS,
  };
}
