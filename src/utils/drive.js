function crearCapetaRaiz() {
  const iter = DriveApp.getFoldersByName(CARPETA_RAIZ);
  if (iter.hasNext()) return iter.next();
  return DriveApp.createFolder(CARPETA_RAIZ);
}

function crearCarpetaRecibos() {
  const raiz = crearCapetaRaiz();
  const iter = raiz.getFoldersByName(CARPETA_RECIBOS);
  if (iter.hasNext()) return iter.next();
  return raiz.createFolder(CARPETA_RECIBOS);
}

function crearCarpertaPorFecha(fecha) {
  const nombreCarpeta = Utilities.formatDate(
    fecha,
    Session.getScriptTimeZone(),
    FORMATO_FECHA_CARPETA,
  );
  const carpetaRecibos = crearCarpetaRecibos();
  const iter = carpetaRecibos.getFoldersByName(nombreCarpeta);
  if (iter.hasNext()) return iter.next();
  return carpetaRecibos.createFolder(nombreCarpeta);
}
