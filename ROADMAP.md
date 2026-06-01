# 🗺️ Roadmap Técnico: Recibos GLR V.1.1.0

**Rama de trabajo:** `feat/v1.1.0-config-base`

**Objetivo Principal:** Eliminar datos _hardcodeados_, soportar múltiples plantillas (GLR/Playaviva) mediante una configuración dinámica en Sheets, y refinar el flujo operativo (folios, fotos y seguridad).

---

## 🛠️ Fase 1: Arquitectura y Configuración Base

_El sistema debe saber de dónde leer la configuración antes de operar._

- [ ] **Inicializador de Configuración:** Crear función en el menú de Sheets que genere una hoja oculta llamada `Config`.
- [ ] **Estructura de `Config`:** Debe almacenar:
  - Plantillas (Nombre, ID del Doc, Nombre de la Hoja destino).
  - Directorio de Contactos (Nombre de la persona, Correo).
  - Responsables de edición por hoja.
- [ ] **Refactor de `config.js`:** Actualizar el código para que lea dinámicamente la hoja `Config` en lugar de usar constantes fijas.
- [ ] **Generador Multi-Hoja:** Modificar la inicialización para que cree hojas separadas (ej. `RecibosGLR`, `RecibosPlayaviva`) en base a `Config`.
- [ ] **Protección Visual:** Aplicar inmovilización de primera fila y proteger (bloquear edición manual) a las columnas automáticas en las hojas generadas.

## ⚙️ Fase 2: Lógica de Captura y Backend

_Mejorar la fiabilidad de los datos ingresados en Sheets._

- [ ] **Texto a Mayúsculas:** Modificar `onEditTrigger.js` para convertir a MAYÚSCULAS el texto ingresado (excluyendo fórmulas, URLs y correos).
- [ ] **Folios Precisos:** Actualizar la generación del Folio para que utilice el número de fila (`range.getRow()`) en lugar de un número aleatorio.
- [ ] **Generación Incremental:** Modificar `generadorRecibos.js` para que guarde la URL del documento en la celda **inmediatamente** después de generarlo, no al finalizar todo el bucle.

## 📸 Fase 3: Nuevo Flujo de Foto y Correo

_Evitar que el documento se rompa antes de tiempo y cargar correos dinámicos._

- [ ] **Nueva Columna:** Añadir la columna `EstatusFoto` al inicializar las hojas.
- [ ] **Retraso de Borrado de Foto:** Modificar `firma.js` para que **NO** borre el tag `{{IdentificacionCliente}}` si el usuario no tomó foto.
- [ ] **Limpieza Final:** Modificar `correo.js` para que evalúe si existe el tag `{{IdentificacionCliente}}` y lo borre _justo antes_ de convertir el documento a PDF.
- [ ] **Endpoint de Correos:** Modificar `webapp.js` para que envíe al frontend la lista dinámica de contactos (Nombres + Correos) leída de la hoja `Config`.

## 💻 Fase 4: Interfaz Web App (Frontend)

_Reflejar los cambios del backend en la pantalla del usuario._

- [ ] **Filtros de Vista:** Añadir controles en la UI para filtrar los recibos por su hoja de origen (GLR o Playaviva).
- [ ] **Contexto en Tarjetas:** Mostrar el número de índice/fila del Sheet en la tarjeta del recibo.
- [ ] **Bloqueo de Foto:** Leer `EstatusFoto`. Si ya hay foto, deshabilitar el botón de capturar e indicar que ya existe.
- [ ] **Mejoras del Canvas:** Aumentar la resolución base del canvas de la foto para que se inserte en un tamaño mayor en el Doc. Mostrar mensaje de éxito al guardar.

## 🛡️ Fase 5: Permisos y Limpieza

_Proteger el sistema de errores humanos._

- [ ] **Columna ELIMINAR:** Añadir columna con un Smart Chip / Dropdown de confirmación.
- [ ] **Trigger de Borrado:** Crear un `onEdit` que escuche la columna ELIMINAR; si se activa, debe mover el Doc a la papelera de Drive y limpiar/borrar la fila en Sheets.
- [ ] **Permisos por Hoja:** Crear un script que proteja `RecibosGLR` y `RecibosPlayaviva` para que solo sean editables por los correos asignados a cada una en la hoja `Config`.
