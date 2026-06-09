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
  const solicitudes = []; // Arreglo para las solicitudes

  // 1. OBTENER RECIBOS PENDIENTES
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

  // 2. OBTENER SOLICITUDES PENDIENTES
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const hojaSolicitudes = spreadSheet.getSheetByName(NOMBRE_HOJA_SOLICITUDES);

  if (hojaSolicitudes) {
    const datosSol = hojaSolicitudes.getDataRange().getValues();

    for (let i = 1; i < datosSol.length; i++) {
      const fila = datosSol[i];
      const estadoConfirmacion = fila[5] || "PENDIENTE";
      const estadoCorreoSol = fila[6] || "PENDIENTE";

      // Solo traemos las que no han sido totalmente procesadas
      if (
        estadoConfirmacion !== "CONFIRMADO" ||
        estadoCorreoSol !== "ENVIADO"
      ) {
        // Convertir el objeto de Fecha a Texto (ISO) para que el servidor no colapse
        let fechaLimpia = fila[0];
        if (fechaLimpia instanceof Date) {
          fechaLimpia = fechaLimpia.toISOString();
        } else if (fechaLimpia) {
          fechaLimpia = fechaLimpia.toString();
        }

        solicitudes.push({
          rowIndex: i + 1,
          marcaTemporal: fechaLimpia, // Ahora viaja como texto seguro
          cliente: fila[1] || "",
          correoSolicitante: fila[2] || "",
          concepto: fila[3] || "",
          importe: fila[4] || 0,
          estadoConfirmacion: estadoConfirmacion,
          estadoCorreo: estadoCorreoSol,
          destinatarios: fila[7] || "",
        });
      }
    }
  }

  // 3. RETORNAR TODO AL FRONTEND
  return {
    recibos: recibos.reverse(),
    solicitudes: solicitudes.reverse(), // Las más recientes arriba
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

function confirmarSolicitudesMasivo(filasAConfirmar) {
  try {
    const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = spreadSheet.getSheetByName(NOMBRE_HOJA_SOLICITUDES);
    if (!hoja) throw new Error("Hoja de solicitudes no encontrada");

    // Recorrer las filas y marcar la Columna F (Confirmación Caja, que es índice 6 en base 1)
    filasAConfirmar.forEach((rowIndex) => {
      hoja
        .getRange(rowIndex, COLUMNAS_SOLICITUDES.CONFIRMACION + 1)
        .setValue("CONFIRMADO");
    });

    return { success: true };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}
