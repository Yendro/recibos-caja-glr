# 🗺️ Roadmap Técnico: Recibos GLR (Versión 2.0.0)

**Objetivo Principal V.2.0.0:** Consolidar el sistema en una Single Page Application (SPA) modular, automatizar la seguridad granular en Sheets, y construir un módulo de atención de solicitudes conectando Google Forms con la Web App.

---

## 🧩 Fase 1: Modularización y Migración del Generador

_Preparar el terreno en el código y trasladar la generación al frontend centralizado._

- [ ] **Refactorización de Scripts:** Dividir el monolítico `script.html` en múltiples archivos lógicos (ej. `script_core`, `script_ui`, `script_modals`) para mantener el proyecto mantenible.
- [ ] **Componente Sidebar (Offcanvas):** Diseñar e integrar un menú lateral deslizable con Bootstrap Offcanvas (`sidebar.html`) que oscurezca el fondo al abrirse, adaptado al diseño UI de la app.
- [ ] **Migración de Generación:** Trasladar la lógica del botón "Generar Pendientes" del entorno de Sheets al nuevo Sidebar del SPA. Esto asegura que la auditoría y ejecución corran exclusivamente bajo la cuenta administradora.

## 🔐 Fase 2: Seguridad y Permisos de Hojas

_Aislar las áreas de trabajo para evitar errores operativos entre responsables._

- [ ] **Protección por Código:** Actualizar `inicializador.js` para que, al crear una hoja, bloquee automáticamente toda su cuadrícula.
- [ ] **Inyección de Editores:** Leer la columna "Correos Autorizados" de la hoja `Config` y otorgar permisos de edición exclusivos a esos correos en sus respectivas hojas.

## 🗑️ Fase 3: Módulo de Administración (Eliminación Segura)

_Manejar la eliminación de registros desde el SPA para garantizar el borrado físico del PDF._

- [ ] **Limpieza de Triggers:** Remover la lógica de eliminación (`🗑️ ELIMINAR RECIBO`) del `onEditTrigger.js` en Sheets para evitar fallos de permisos delegados.
- [ ] **UI de Administración:** Crear una vista/modal en el SPA que permita seleccionar múltiples recibos mediante checkboxes.
- [ ] **Backend de Borrado:** Programar función que reciba los recibos seleccionados, mueva sus respectivos Google Docs a la papelera de Drive, y elimine las filas en Sheets con permisos de Owner.

## 📝 Fase 4: Módulo de Solicitudes (Integración Google Forms)

_Permitir a los clientes consultar el estatus de su dinero mediante un flujo automatizado._

- [ ] **Configuración de Hoja Base:** Añadir al inicializador la creación de la hoja oficial "Solicitud de Recibos".
- [ ] **Trigger Puente (`onFormSubmit`):** Crear script que escuche el envío del Google Form nativo, extraiga los datos y los inserte "limpios" en la hoja oficial con los estados iniciales (`PENDIENTE`).
- [ ] **Vista SPA "Pendientes por Confirmar":** Desarrollar una nueva sección/pestaña en la interfaz que renderice exclusivamente las solicitudes que aún no han sido confirmadas.
- [ ] **Flujo de Acciones (Confirmar y Enviar):** - Botón **Confirmar**: Cambia el estado en Sheets a `CONFIRMADO` (con modal de advertencia previo).
  - Botón **Enviar**: Habilita el envío de un correo de notificación al cliente (usando el directorio dinámico o el correo proporcionado) avisando que su dinero está en Caja. Al enviarse, la solicitud desaparece de la vista de pendientes.
