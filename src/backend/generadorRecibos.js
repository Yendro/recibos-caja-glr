function generarRecibo(rowData) {
  const {
    rowIndex,
    cliente,
    importe,
    importeLetra,
    concepto,
    fechaPago,
    folio,
    fechaCreacion,
  } = rowData;

  try {
    const fechaObj =
      fechaCreacion instanceof Date ? fechaCreacion : new Date(fechaCreacion);
    const carpetaDestino = crearCarpertaPorFecha(fechaObj);

    const plantilla = DriveApp.getFileById(ID_PLANTILLA_RECIBO);
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
    actualizarEstadoArchivo(rowIndex, "Generado", url);

    return { success: true, url };
  } catch (error) {
    console.error(
      `Error al generar recibo para fila ${rowIndex}: ${error.message}`,
    );
    actualizarEstadoArchivo(rowIndex, "Error");
    return { success: false, error: error.message };
  }
}

function generar() {}

function generarTodosPendientes() {
  console.log("[INICIO] Iniciando generación masiva de recibos pendientes");
  try {
    const pendientes = obtenerFilasPendientes();
    console.log(
      `[DEBUG] Se encontraron ${pendientes.length} recibos pendientes`,
    );

    if (pendientes.length === 0) {
      console.log("[INFO] No hay recibos pendientes para generar");
      return [];
    }

    const resultados = [];
    let exitosos = 0;
    let fallidos = 0;

    for (let i = 0; i < pendientes.length; i++) {
      const pend = pendientes[i];
      console.log(`[DEBUG] Procesando recibo ${i + 1} de ${pendientes.length}`);

      const res = generarRecibo(pend);
      resultados.push(res);

      if (res.success) {
        exitosos++;
      } else {
        fallidos++;
      }
    }

    console.log(
      `[RESUMEN] Generación completada - Exitosos: ${exitosos}, Fallidos: ${fallidos}`,
    );
    return resultados;
  } catch (error) {
    console.error(
      `[ERROR CRÍTICO] Error en generarTodosPendientes: ${error.message}`,
    );
    console.error(`[ERROR] Stack: ${error.stack}`);
    return [
      {
        success: false,
        error: `Error crítico: ${error.message}`,
        fatal: true,
      },
    ];
  }
}
