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
          "Nombre Hoja (Destino)",
          "ID de la Plantilla de Docs",
          "Correos Autorizados (separados por coma)",
          "Prefijo Folio",
        ],
      ])
      .setBackground("#f1f5f9")
      .setFontWeight("bold");

    // Datos de EJEMPLO (Placeholders genéricos)
    hojaConfig.getRange("A3:D4").setValues([
      [
        "RecibosMarcaA",
        "PEGAR_ID_PLANTILLA_AQUI",
        "admin@tuempresa.com",
        "MARCAA",
      ],
      [
        "RecibosMarcaB",
        "PEGAR_ID_PLANTILLA_AQUI",
        "admin@tuempresa.com",
        "MARCAB",
      ],
    ]);

    // Formato de Correos
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

    // Datos de EJEMPLO
    hojaConfig.getRange("E3:F4").setValues([
      ["Juan Perez", "juan@tuempresa.com"],
      ["Maria Gomez", "maria@tuempresa.com"],
    ]);

    hojaConfig.autoResizeColumns(1, 6);

    SpreadsheetApp.getUi().alert(
      "Paso 1 Completado",
      "Se ha creado la hoja 'Config' con datos de ejemplo.\n\nPor favor, llena esta hoja con los nombres de tus hojas, IDs de plantillas y correos reales.\n\nCuando termines, vuelve a hacer clic en 'Inicializar Sistema' para generar las hojas operativas.",
      SpreadsheetApp.getUi().ButtonSet.OK,
    );

    // Detenemos la ejecución aquí para que el usuario pueda llenar los datos
    return;
  }

  // PASO 2: SI YA EXISTE 'CONFIG', LEER DATOS REALES Y CREAR HOJAS DE RECIBOS
  const config = obtenerConfiguracion();
  let hojasCreadas = 0;

  config.plantillas.forEach((plantilla) => {
    // Ignorar si el usuario olvidó borrar el placeholder
    if (
      plantilla.idPlantilla === "PEGAR_ID_PLANTILLA_AQUI" ||
      plantilla.nombreHoja === ""
    )
      return;

    let hojaRecibos = spreadSheet.getSheetByName(plantilla.nombreHoja);

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

      // 3. Fondo suave a las columnas de ingreso manual
      hojaRecibos
        .getRange(2, 1, hojaRecibos.getMaxRows() - 1, 4)
        .setBackground("#fff2ee");

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

      // 5. Agregar Validación de Fecha (Calendario) en la columna FechaPago (Columna 4)
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

      // 6. Proteger las columnas automáticas
      const numColumnasAutomaticas = nombresColumnas.length - 4;
      const proteccion = hojaRecibos
        .getRange(1, 5, hojaRecibos.getMaxRows(), numColumnasAutomaticas)
        .protect();
      proteccion.setDescription(
        "Columnas Protegidas - " + plantilla.nombreHoja,
      );
      proteccion.setWarningOnly(true);

      hojasCreadas++;
    }
  });

  // Ocultamos la hoja de configuración por seguridad solo después de generar las hojas
  hojaConfig.hideSheet();

  SpreadsheetApp.getUi().alert(
    "Éxito",
    `Sistema inicializado correctamente.\nSe verificaron/crearon ${hojasCreadas} hojas de recibos nuevas.\n\n- Las columnas manuales tienen diferente color.\n- La columna FechaPago tiene calendario interactivo.\n- Las columnas automáticas están protegidas.`,
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
}
