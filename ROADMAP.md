# 🗺️ Roadmap Técnico: Recibos GLR

**Objetivo Principal V.1.1.0:** Eliminar datos _hardcodeados_, soportar múltiples plantillas (GLR/Playaviva) mediante una configuración dinámica en Sheets, y refinar el flujo operativo (folios, fotos y seguridad).

---

## 🛠️ Fase 1: Arquitectura y Configuración Base (Completado)

_El sistema debe saber de dónde leer la configuración antes de operar._

- [x] **Inicializador de Configuración:** Crear función en el menú de Sheets que genere una hoja oculta llamada `Config`.
- [x] **Estructura de `Config`:** Debe almacenar Plantillas, Directorio de Contactos y Directorio de Drive.
- [x] **Refactor de `config.js`:** Actualizar el código para que lea dinámicamente la hoja `Config` en lugar de usar constantes fijas.
- [x] **Generador Multi-Hoja:** Modificar la inicialización para que cree hojas separadas (ej. `RecibosGLR`, `RecibosPlayaviva`).
- [x] **Protección Visual:** Aplicar inmovilización de primera fila y proteger (bloquear edición manual) a las columnas automáticas.

## ⚙️ Fase 2: Lógica de Captura y Backend (Completado)

_Mejorar la fiabilidad de los datos ingresados en Sheets._

- [x] **Texto a Mayúsculas:** Modificar `onEditTrigger.js` para convertir a MAYÚSCULAS el texto ingresado.
- [x] **Folios Precisos:** Actualizar la generación del Folio para que utilice el número de fila.
- [x] **Generación Incremental:** Guardar la URL del documento en la celda **inmediatamente** después de generarlo.

## 📸 Fase 3: Nuevo Flujo de Foto y Correo (Completado)

_Evitar que el documento se rompa antes de tiempo y cargar correos dinámicos._

- [x] **Nueva Columna:** Añadir la columna `EstatusFoto` al inicializar las hojas.
- [x] **Retraso de Borrado de Foto:** **NO** borrar el tag `{{IdentificacionCliente}}` si el usuario no tomó foto de inmediato.
- [x] **Limpieza Final:** Borrar la etiqueta vacía _justo antes_ de convertir el documento a PDF.
- [x] **Endpoint de Correos:** Enviar al frontend la lista dinámica de contactos leída de la hoja `Config`.

## 💻 Fase 4: Interfaz Web App Frontend (Completado)

_Reflejar los cambios del backend en la pantalla del usuario._

- [x] **Filtros de Vista:** Añadir controles UI (Pill buttons) para filtrar los recibos por su hoja de origen.
- [x] **Contexto en Tarjetas:** Mostrar el número de índice/fila del Sheet en la tarjeta del recibo (Badge Superior).
- [x] **Bloqueo de Foto:** Deshabilitar el botón de capturar e indicar "Capturada" si ya existe foto.
- [x] **Mejoras del Canvas:** Aumentar resolución base a 500px y presentar visualización UI en grid de 2 columnas para correos.

## 🛡️ Fase 5: Permisos y Limpieza (Completado Parcialmente)

_Proteger el sistema de errores humanos._

- [x] **Columna ELIMINAR:** Añadir columna con un Smart Chip / Dropdown de confirmación.
- [x] **Trigger de Borrado:** Trigger `onEdit` que mueve el Doc a la papelera de Drive y elimina la fila.
- [ ] **Permisos por Hoja Automáticos:** _(Decidido estratégicamente: Movido a la versión V.2.0.0 para evitar conflictos con el Scope de ejecución. Se manejará manualmente por interfaz nativa temporalmente)._

---

## 🚑 Hotfix V.1.1.1: Configuración de Remitente (Completado)

_Enmascaramiento de remitente para evitar exposición de cuenta personal._

- [x] **Campos de Interfaz:** Añadir campos de "Alias de Correo" y "Nombre de Remitente" a la sección de Configuración de `inicializador.js`.
- [x] **Refactor de Envío:** Cambiar la API de `MailApp` a `GmailApp` para soportar parámetros avanzados.
- [x] **Fallbacks:** Inyectar "Caja" como prevención de errores si el usuario no configura la celda.
