const FORMATO_FECHA_CARPETA = "yyyy-MM-dd";
const NOMBRE_HOJA_CONFIG = "Config";
const NOMBRE_HOJA_SOLICITUDES = "Solicitudes";

const COLUMNAS = {
  CLIENTE: 0,
  IMPORTE: 1,
  CONCEPTO: 2,
  FECHA_PAGO: 3,
  IMPORTE_LETRA: 4,
  FOLIO: 5,
  FECHA_CREACION: 6,
  STATUS: 7,
  ARCHIVO: 8,
  ESTADO_FIRMA: 9,
  ESTADO_CORREO: 10,
  DESTINATARIOS: 11,
  ESTATUS_FOTO: 12,
};

const COLUMNAS_SOLICITUDES = {
  MARCA_TEMPORAL: 0,
  CLIENTE: 1,
  CORREO_SOLICITANTE: 2,
  CONCEPTO: 3,
  IMPORTE: 4,
  CONFIRMACION: 5,
  ESTADO_CORREO: 6,
  DESTINATARIOS: 7,
};

/**
 * Lee la hoja de 'Config' y construye un objeto con las plantillas y correos
 * @returns {Object} { plantillas: Array, directorioCorreos: Array }
 */
function obtenerConfiguracion() {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const hojaConfig = spreadSheet.getSheetByName(NOMBRE_HOJA_CONFIG);

  if (!hojaConfig) {
    throw new Error(
      "La hoja de Configuración no existe. Ejecuta 'Inicializar Sistema' desde el menú.",
    );
  }

  // Leer Variables Generales (Drive y Correo)
  const carpetaRaiz =
    hojaConfig.getRange("H2").getValue() || "Recibos-Generales";
  const aliasCorreo = hojaConfig.getRange("H5").getValue() || "";
  const nombreRemitente =
    hojaConfig.getRange("H6").getValue() || "Sistema de Recibos";

  // Leer el ID del formulario
  const idFormulario = hojaConfig.getRange("H7").getValue() || "";

  const datosPlantillas = hojaConfig.getRange("A3:D10").getValues();
  const plantillas = [];
  for (const fila of datosPlantillas) {
    if (fila[0] && fila[1]) {
      plantillas.push({
        nombreHoja: fila[0].toString().trim(),
        idPlantilla: fila[1].toString().trim(),
        editores: fila[2]
          ? fila[2]
              .toString()
              .split(",")
              .map((e) => e.trim())
          : [],
        prefijoFolio: fila[3] ? fila[3].toString().trim().toUpperCase() : "REC",
      });
    }
  }

  const datosCorreos = hojaConfig.getRange("E3:F30").getValues();
  const directorioCorreos = [];
  for (const fila of datosCorreos) {
    if (fila[0] && fila[1]) {
      directorioCorreos.push({
        nombre: fila[0].toString().trim(),
        correo: fila[1].toString().trim(),
      });
    }
  }

  return {
    carpetaRaiz: carpetaRaiz,
    aliasCorreo: aliasCorreo.toString().trim(),
    nombreRemitente: nombreRemitente.toString().trim(),
    idFormulario: idFormulario.toString().trim(),
    plantillas: plantillas,
    directorioCorreos: directorioCorreos,
  };
}
