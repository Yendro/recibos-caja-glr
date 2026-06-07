function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("⚙️ Configuración")
    .addItem("Configurar e Iniciar Sistema", "inicializarSistema")
    .addSeparator()
    .addItem("Configurar Permisos de Google", "forzarPermisos")
    .addToUi();
}

function forzarPermisos() {
  // Al invocar esta línea, Google fuerza la revisión del appsscript.json
  GmailApp.getAliases();

  // Si pasa la línea anterior sin detenerse por permisos, mostramos éxito
  SpreadsheetApp.getUi().alert(
    "Permisos Configurados",
    "Si no te apareció la pantalla de advertencia de Google, significa que tu cuenta ya tiene todos los permisos necesarios aprobados.\n\nYa puedes utilizar el sistema con normalidad.",
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
}

function obtenerEstadoPendientes() {
  const pendientes = obtenerFilasPendientes();
  return pendientes.map((p) => ({
    cliente: p.cliente,
    folio: p.folio,
    rowIndex: p.rowIndex,
  }));
}
