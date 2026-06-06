// Recibimos la hoja pre-cargada para no buscarla
function generarRecibo(rowData, hojaCache) {
  const {
    rowIndex,
    cliente,
    importe,
    importeLetra,
    concepto,
    fechaPago,
    folio,
    fechaCreacion,
    nombreHoja, // Identifica en qué hoja escribir
    idPlantilla, // Identifica qué plantilla copiar
  } = rowData;

  try {
    const fechaObj =
      fechaCreacion instanceof Date ? fechaCreacion : new Date(fechaCreacion);
    const carpetaDestino = crearCarpertaPorFecha(fechaObj, nombreHoja);

    // Uso del ID de plantilla dinámico
    const plantilla = DriveApp.getFileById(idPlantilla);
    const clienteSeguro = cliente
      .trim()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/\s+/g, "")
      .replace(/[\\/:*?"<>|]/g, "");

    const nombreArchivo = `${folio}-${clienteSeguro}.gdoc`;
    const copia = plantilla.makeCopy(nombreArchivo, carpetaDestino);
    const documento = DocumentApp.openById(copia.getId());
    const body = documento.getBody();

    const importeFormateado =
      "$ " +
      Number(importe).toLocaleString("es-MX", {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
      }) +
      " MXN";
    const fechaPagoFormateada = Utilities.formatDate(
      new Date(fechaPago),
      Session.getScriptTimeZone(),
      "dd/MM/yyyy",
    );

    body.replaceText("{{Cliente}}", cliente);
    body.replaceText("{{Importe}}", importeFormateado);
    body.replaceText("{{ImporteLetra}}", importeLetra);
    body.replaceText("{{Concepto}}", concepto);
    body.replaceText("{{FechaPago}}", fechaPagoFormateada);
    body.replaceText("{{Folio}}", folio);

    documento.saveAndClose();
    const url = documento.getUrl();

    // Usamos la escritura en lote y le pasamos la hoja en caché
    actualizarEstadoArchivo(rowIndex, "GENERADO", url, hojaCache);

    return { success: true, url };
  } catch (error) {
    console.error(
      `Error fila ${rowIndex} hoja ${nombreHoja}: ${error.message}`,
    );
    actualizarEstadoArchivo(rowIndex, "ERROR", null, hojaCache);
    return { success: false, error: error.message };
  }
}

function generarTodosPendientes() {
  // 1. Pedimos la llave del documento
  const lock = LockService.getDocumentLock();

  // 2. Intentamos entrar (espera hasta 10 segundos). Si alguien más lo está usando, rebotamos.
  if (!lock.tryLock(10000)) {
    return [
      {
        success: false,
        error:
          "⚠️ Otro usuario está generando recibos en este momento. Por favor, espera unos segundos e intenta de nuevo.",
      },
    ];
  }

  try {
    const pendientesTotales = obtenerFilasPendientes();
    if (pendientesTotales.length === 0) return [];

    // Límite de seguridad para evitar Timeouts de Google
    const MAX_PROCESAMIENTO = 15;
    const pendientes = pendientesTotales.slice(0, MAX_PROCESAMIENTO);

    // Caché de Hojas de cálculo
    const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    const hojasCache = {};

    const resultados = [];

    for (let i = 0; i < pendientes.length; i++) {
      const pend = pendientes[i];

      // Si la hoja no está en caché, la buscamos y la guardamos
      if (!hojasCache[pend.nombreHoja]) {
        hojasCache[pend.nombreHoja] = spreadSheet.getSheetByName(
          pend.nombreHoja,
        );
      }

      // Le pasamos la hoja pre-cargada a la función
      const res = generarRecibo(pend, hojasCache[pend.nombreHoja]);
      resultados.push(res);

      // SpreadsheetApp.flush();
    }

    // Si había más de 15, le avisa al frontend inyectando un mensaje
    if (pendientesTotales.length > MAX_PROCESAMIENTO) {
      resultados[0].warning = `Se generaron los primeros ${MAX_PROCESAMIENTO} recibos por seguridad. Quedan ${pendientesTotales.length - MAX_PROCESAMIENTO} en cola. Vuelve a hacer clic en Generar.`;
    }

    return resultados;
  } catch (error) {
    console.error(`[ERROR CRÍTICO] ${error.message}`);
    return [
      { success: false, error: `Error crítico: ${error.message}`, fatal: true },
    ];
  } finally {
    lock.releaseLock();
  }
}
