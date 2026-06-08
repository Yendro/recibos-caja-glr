function inicializarSistema() {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  let hojaConfig = spreadSheet.getSheetByName("Config");

  // PASO 1: CREAR HOJA CONFIG COMO PLANTILLA GENÉRICA (Si no existe)
  if (!hojaConfig) {
    hojaConfig = spreadSheet.insertSheet("Config");

    // Formato de Plantillas
    hojaConfig
      .getRange("A1:D1")
      .setValues([["CONFIGURACIÓN DE PLANTILLAS Y HOJAS", "", "", ""]])
      .merge()
      .setBackground("#fa8072")
      .setFontColor("white")
      .setFontWeight("bold");
    hojaConfig
      .getRange("A2:D2")
      .setValues([
        [
          "Nombre Hoja",
          "ID de la Plantilla de Docs",
          "Correos Autorizados (separados por coma)",
          "Prefijo Folio",
        ],
      ])
      .setBackground("#f1f5f9")
      .setFontWeight("bold");
    hojaConfig.getRange("A3:D4").setValues([
      [
        "RecibosMarcaA",
        "PEGAR_ID_PLANTILLA_AQUI",
        "admin@tuempresa.com,ejemplo@separado.com",
        "MARCAA",
      ],
      [
        "RecibosMarcaB",
        "PEGAR_ID_PLANTILLA_AQUI",
        "admin@tuempresa.com",
        "MARCAB",
      ],
    ]);

    // Formato de Directorio de Correos
    hojaConfig
      .getRange("E1:F1")
      .setValues([["DIRECTORIO DE CORREOS", ""]])
      .merge()
      .setBackground("#0075c9")
      .setFontColor("white")
      .setFontWeight("bold");
    hojaConfig
      .getRange("E2:F2")
      .setValues([["Nombre de la Persona", "Correo Electrónico"]])
      .setBackground("#f1f5f9")
      .setFontWeight("bold");
    hojaConfig.getRange("E3:F4").setValues([
      ["Juan Perez", "juan@tuempresa.com"],
      ["Maria Gomez", "maria@tuempresa.com"],
    ]);

    // Formato de Configuración de Drive
    hojaConfig
      .getRange("G1:H1")
      .setValues([["CONFIGURACIÓN DRIVE", ""]])
      .merge()
      .setBackground("#fa8072")
      .setFontColor("white")
      .setFontWeight("bold");
    hojaConfig
      .getRange("G2:H2")
      .setValues([["Nombre Carpeta Raíz", "/Carpeta donde está este archivo"]])
      .setBackground("#f1f5f9")
      .setFontWeight("bold");

    // Formato de Configuración de Correo
    hojaConfig
      .getRange("G4:H4")
      .setValues([["CONFIGURACIÓN DE CORREO", ""]])
      .merge()
      .setBackground("#0075c9")
      .setFontColor("white")
      .setFontWeight("bold");
    hojaConfig
      .getRange("G5:H5")
      .setValues([
        ["Correo que se verá al enviar el recibo", "correo@recibo.com"],
      ])
      .setBackground("#f1f5f9")
      .setFontWeight("bold");
    hojaConfig
      .getRange("G6:H6")
      .setValues([
        ["Nombre que se verá al enviar el recibo", "Caja NombreEmpresa"],
      ])
      .setBackground("#f1f5f9")
      .setFontWeight("bold");

    // Aplicar Formato General a la hoja Config (Alineación, Centrado y Ajuste de texto)
    const rangoTotalConfig = hojaConfig.getRange(1, 1, 30, 8); // Abarcar todo el área de configuración
    rangoTotalConfig.setWrap(true);
    rangoTotalConfig.setVerticalAlignment("middle");
    rangoTotalConfig.setHorizontalAlignment("left");

    // Anchos de columna considerables para que todo sea legible
    hojaConfig.setColumnWidth(1, 180); // Nombre Hoja
    hojaConfig.setColumnWidth(2, 350); // ID Plantilla
    hojaConfig.setColumnWidth(3, 250); // Correos Autorizados
    hojaConfig.setColumnWidth(4, 120); // Prefijo Folio
    hojaConfig.setColumnWidth(5, 200); // Nombre Persona
    hojaConfig.setColumnWidth(6, 250); // Correo Persona
    hojaConfig.setColumnWidth(7, 260); // Textos descriptivos (G)
    hojaConfig.setColumnWidth(8, 250); // Valores ingresados (H)

    SpreadsheetApp.getUi().alert(
      "Paso 1 Completado",
      "Se ha creado la hoja 'Config' con datos de ejemplo.\n\nPor favor, llena esta hoja con los nombres de tus hojas, IDs de plantillas y correos reales.\n\nCuando termines, vuelve a hacer clic en 'Inicializar Sistema' para generar las hojas operativas.",
      SpreadsheetApp.getUi().ButtonSet.OK,
    );

    return;
  }

  // PASO 2: SI YA EXISTE 'CONFIG', LEER DATOS REALES Y CREAR/PROTEGER HOJAS DE RECIBOS
  const config = obtenerConfiguracion();
  let hojasCreadas = 0;
  let hojasProtegidas = 0;

  config.plantillas.forEach((plantilla) => {
    // Ignorar si el usuario olvidó borrar el placeholder
    if (
      plantilla.idPlantilla === "PEGAR_ID_PLANTILLA_AQUI" ||
      plantilla.nombreHoja === ""
    )
      return;

    let hojaRecibos = spreadSheet.getSheetByName(plantilla.nombreHoja);

    // --- A. CREACIÓN DE LA HOJA (Si no existe) ---
    if (!hojaRecibos) {
      hojaRecibos = spreadSheet.insertSheet(plantilla.nombreHoja);

      const nombresColumnas = [
        "Cliente",
        "Importe",
        "Concepto",
        "FechaPago",
        "ImporteLetra",
        "Folio",
        "FechaCreacion",
        "Status",
        "Archivo",
        "EstadoFirma",
        "EstadoCorreo",
        "Destinatarios",
        "EstatusFoto",
      ];

      // 1. Asignar los nombres de encabezado
      hojaRecibos
        .getRange(1, 1, 1, nombresColumnas.length)
        .setValues([nombresColumnas])
        .setFontColor("white")
        .setFontWeight("bold");

      // 2. Colores Bicolor para los encabezados
      hojaRecibos.getRange(1, 1, 1, 4).setBackground("#fa8072"); // Naranja para las de ingreso manual
      hojaRecibos
        .getRange(1, 5, 1, nombresColumnas.length - 4)
        .setBackground("#0075c9"); // Azul para las automáticas

      // 3. Fondos suaves para diferenciar celdas manuales de automáticas
      hojaRecibos
        .getRange(2, 1, hojaRecibos.getMaxRows() - 1, 4)
        .setBackground("#fff2ee"); // Naranja suave
      hojaRecibos
        .getRange(
          2,
          5,
          hojaRecibos.getMaxRows() - 1,
          nombresColumnas.length - 4,
        )
        .setBackground("#eaf4fc"); // Azul suave

      // 4. Inmovilizar paneles
      hojaRecibos.setFrozenRows(1);
      hojaRecibos.setFrozenColumns(4);

      // Columna Importe (B / 2) -> Moneda
      hojaRecibos
        .getRange(2, 2, hojaRecibos.getMaxRows() - 1, 1)
        .setNumberFormat("$#,##0.00");

      // Columna FechaCreacion (G / 7) -> Fecha y Hora
      hojaRecibos
        .getRange(2, 7, hojaRecibos.getMaxRows() - 1, 1)
        .setNumberFormat("dd/MM/yyyy HH:mm:ss");

      // FORMATO GENERAL DE TEXTO
      const rangoDatos = hojaRecibos.getRange(
        2,
        1,
        hojaRecibos.getMaxRows() - 1,
        nombresColumnas.length,
      );
      rangoDatos.setWrap(true); // Ajustar texto
      rangoDatos.setVerticalAlignment("middle"); // Centrado vertical a altura media
      rangoDatos.setHorizontalAlignment("left"); // Alineación horizontal a la izquierda

      // --- ANCHO DE COLUMNAS (En pixeles) ---
      hojaRecibos.setColumnWidth(1, 200); // Cliente
      hojaRecibos.setColumnWidth(2, 120); // Importe
      hojaRecibos.setColumnWidth(3, 250); // Concepto
      hojaRecibos.setColumnWidth(4, 120); // FechaPago
      hojaRecibos.setColumnWidth(5, 300); // ImporteLetra (Largo)
      hojaRecibos.setColumnWidth(6, 220); // Folio
      hojaRecibos.setColumnWidth(7, 150); // FechaCreacion
      hojaRecibos.setColumnWidth(8, 120); // Status
      hojaRecibos.setColumnWidth(9, 350); // Archivo (URL)
      hojaRecibos.setColumnWidth(10, 120); // EstadoFirma
      hojaRecibos.setColumnWidth(11, 120); // EstadoCorreo
      hojaRecibos.setColumnWidth(12, 250); // Destinatarios
      hojaRecibos.setColumnWidth(13, 120); // EstatusFoto

      // Validación de Fechas
      const reglaFecha = SpreadsheetApp.newDataValidation()
        .requireDate()
        .setAllowInvalid(false)
        .setHelpText(
          "Por favor, introduce una fecha válida o haz doble clic para abrir el calendario.",
        )
        .build();
      hojaRecibos
        .getRange(2, 4, hojaRecibos.getMaxRows() - 1, 1)
        .setDataValidation(reglaFecha);

      const numColumnasAutomaticas = nombresColumnas.length - 4;
      const rangoAutomaticas = hojaRecibos.getRange(
        2,
        5,
        hojaRecibos.getMaxRows() - 1,
        numColumnasAutomaticas,
      );
      const proteccionAviso = rangoAutomaticas
        .protect()
        .setDescription("Columnas Automáticas");
      proteccionAviso.setWarningOnly(true);

      hojasCreadas++;
    }

    // --- B. PROTECCIÓN DE LA HOJA (Se ejecuta siempre que inicializas) ---
    // Esto asegura que si agregas un nuevo correo en la Config, se aplique a hojas existentes

    // 1. Limpiar protecciones viejas (para evitar duplicados o "capas" de protección)
    const proteccionesAnteriores = hojaRecibos.getProtections(
      SpreadsheetApp.ProtectionType.SHEET,
    );
    for (let i = 0; i < proteccionesAnteriores.length; i++) {
      proteccionesAnteriores[i].remove();
    }

    // 2. Proteger la hoja completa
    const proteccionGlobal = hojaRecibos
      .protect()
      .setDescription("Protección Automatizada: " + plantilla.nombreHoja);

    // 3. Excluir la primera fila (encabezados) para que los filtros funcionen
    const rangoEncabezado = hojaRecibos.getRange("A1:N1");
    proteccionGlobal.setUnprotectedRanges([rangoEncabezado]);

    // 4. Asegurar que tú (Owner) siempre tienes acceso
    const me = Session.getEffectiveUser();
    proteccionGlobal.addEditor(me);

    // 5. Remover a todos los demás editores genéricos del archivo
    proteccionGlobal.removeEditors(proteccionGlobal.getEditors());

    // 6. Inyectar exclusivamente los correos de la Configuración
    if (plantilla.editores.length > 0) {
      // Usamos un bloque Try/Catch por si alguien escribe mal un correo en Config
      try {
        proteccionGlobal.addEditors(plantilla.editores);
      } catch (e) {
        console.error(
          "Error al añadir editores a " + plantilla.nombreHoja + ": " + e,
        );
      }
    }

    hojasProtegidas++;
  });

  // --- C. PROTECCIÓN DE LA HOJA DE CONFIGURACIÓN ---
  const proteccionesConfig = hojaConfig.getProtections(
    SpreadsheetApp.ProtectionType.SHEET,
  );
  for (let i = 0; i < proteccionesConfig.length; i++) {
    proteccionesConfig[i].remove();
  }

  const proteccionConfig = hojaConfig
    .protect()
    .setDescription("Bloqueo de Configuración");
  proteccionConfig.addEditor(Session.getEffectiveUser()); // Solo el owner del archivo puede editar ahora
  proteccionConfig.removeEditors(proteccionConfig.getEditors());

  hojaConfig.hideSheet();

  SpreadsheetApp.getUi().alert(
    "Auditoría Finalizada",
    `Sistema inicializado y asegurado.\n\n` +
      `- Hojas nuevas creadas: ${hojasCreadas}\n` +
      `- Hojas re-protegidas: ${hojasProtegidas}\n\n` +
      `Los permisos de edición ahora son exclusivos para los correos listados en la hoja Config.`,
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
}
