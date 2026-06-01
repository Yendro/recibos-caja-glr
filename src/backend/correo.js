function enviarReciboPorCorreo(rowIndex, destinatariosArr) {
  try {
    const hoja = obtenerHoja();
    const urlDoc = hoja
      .getRange(rowIndex, COLUMNAS.ARCHIVO + 1)
      .getValue()
      .toString();
    const folio = hoja.getRange(rowIndex, COLUMNAS.FOLIO + 1).getValue();
    const cliente = hoja.getRange(rowIndex, COLUMNAS.CLIENTE + 1).getValue();

    const match = urlDoc.match(/[-\w]{25,}/);
    if (!match) throw new Error("No se encontró un documento válido.");
    const docId = match[0];

    // CONVERSIÓN NATIVA: Más rápida, segura y no requiere UrlFetchApp
    const pdfBlob = DriveApp.getFileById(docId)
      .getAs(MimeType.PDF)
      .setName(`${folio}_${cliente}.pdf`);

    // Preparar la plantilla HTML del correo
    const htmlTemplate = HtmlService.createTemplateFromFile(
      "src/frontend/correo",
    );
    htmlTemplate.cliente = cliente;
    htmlTemplate.folio = folio;
    const bodyCorreo = htmlTemplate.evaluate().getContent();

    const correosDestino = destinatariosArr.join(",");

    // Envío por MailApp utilizando los alcances de la cuenta activa
    MailApp.sendEmail({
      to: correosDestino,
      subject: `Recibo Confirmado - ${folio}`,
      htmlBody: bodyCorreo,
      attachments: [pdfBlob],
    });

    // Actualizamos los estatus en el Sheets
    hoja.getRange(rowIndex, COLUMNAS.ESTADO_CORREO + 1).setValue("Enviado");
    hoja
      .getRange(rowIndex, COLUMNAS.DESTINATARIOS + 1)
      .setValue(correosDestino);

    return { success: true, message: "Enviado con éxito" };
  } catch (error) {
    console.error("Error en enviarReciboPorCorreo:", error);
    return { success: false, message: error.toString() };
  }
}
