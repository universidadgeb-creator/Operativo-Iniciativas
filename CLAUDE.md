# GEB Iniciativas Humanas — Panel CRM

Panel interno para el equipo de GEB University / Iniciativas Humanas. Es HTML/JS estático desplegado en **GitHub Pages** (sin backend propio, sin build step, sin frameworks — todo vanilla JS en archivos `.html` autocontenidos). La única fuente de verdad es un **Google Sheet**, y la única forma de escribir en él es un **Google Apps Script Web App** centralizado.

Repo: `universidadgeb-creator/Operativo-Iniciativas` (local: esta carpeta). Dueña del proyecto: Cecilia (coordinadora de GEB University), no técnica — reporta bugs por síntomas visuales o del Sheet, no por stack traces.

## Arquitectura

- **Lectura**: cada página descarga sus datos en vivo con el endpoint público de Google `gviz/tq?tqx=out:json&sheet=<tab>` (sin API key, sin login). El Sheet debe seguir compartido como "Cualquier persona con el enlace puede ver".
- **Escritura**: todo botón que guarda algo (registrar, marcar, dar de alta, dar de baja...) manda un `fetch(WEBAPP_URL, {method:"POST", body: JSON.stringify(payload)})` con un campo `tipo`. El único código que escribe en el Sheet es `apps-script.gs`, desplegado como Web App con un `doPost(e)` que rutea por `payload.tipo` a un handler `handleXxx(p)`. **Ningún botón del panel escribe directo al Sheet.**
- **Sheet único unificado**: desde el 2026-07-04/05 existe un solo Sheet (`SHEET_NUEVO_ID`, constante en `apps-script.gs` y en cada `.html`) que es la fuente de datos de las 10 iniciativas. El Sheet histórico anterior quedó desconectado por completo — no lo uses como referencia de columnas.
- **Redeploy manual obligatorio**: editar `apps-script.gs` en este repo **no actualiza nada por sí solo**. Hay que copiar el archivo completo al editor de Apps Script (Extensiones → Apps Script del Sheet) y crear una **nueva versión de implementación** (Implementar → Administrar implementaciones → lápiz → Nueva versión → Implementar). Recuérdaselo siempre al usuario después de tocar ese archivo.
- **Biblioteca Virtual externa**: `biblioteca-geb.html`/`apps-script.gs` también leen/escriben una pestaña "Fisicos" en un Sheet **distinto** (`SHEET_BIBLIOTECA_ID`) — es más frágil porque usa columnas por posición, no por encabezado.

### Regla de oro: columnas por posición, no por nombre

Cada pestaña tiene un array `SCHEMA`/`COLS` hardcodeado en el `<script>` de cada página, con las columnas en el mismo orden que en el Sheet real. **Nunca se mapea por `cols[].label` de gviz** porque ese label llega vacío en pestañas con fórmulas QUERY. Si alguien reordena, inserta o borra una columna a mano en el Sheet, el panel sigue leyendo por posición y muestra (o escribe) el dato equivocado sin ningún error visible. Esta ha sido la causa más común de bugs del proyecto — ver `INSTRUCCIONES_CRM.md` sección 4 para el historial de corrupciones ya resueltas. Antes de tocar el `SCHEMA` de una pestaña, verifica el orden real contra el Sheet.

### Otras convenciones de código que ya están validadas y no deben romperse

- Filtrado de filas: se descartan filas vacías, filas cuyo primer valor empieza con `#` (errores tipo `#REF!`), y filas cuya primera celda coincide con un nombre de columna del propio SCHEMA (subencabezados que a veces derrama gviz).
- Teléfonos: llegan del Sheet como número entero; se formatean a `+52XXXXXXXXXX` con `formatTel()` (quita `.0`, quita no-dígitos, antepone `+52` si son 10 dígitos) antes de construir cualquier link `wa.me`.
- WhatsApp es **siempre 1:1** — links `wa.me` individuales con texto prellenado. Nunca envío masivo; el operador hace clic persona por persona. Los mensajes grupales (ej. en Retiro Espiritual) usan botón "copiar al portapapeles", no `wa.me`, porque van a un grupo de WhatsApp, no a un número.
- Fechas: gviz devuelve fechas como texto `"Date(y,m,d,...)"` en la mayoría de los casos, pero a veces como texto `dd/mm/yyyy` (cuando vienen de un Form o captura manual). La función `gvizDateToJs()` (repetida en varias páginas) maneja ambos casos — nunca uses `new Date(stringDelSheet)` directo.
- Copys de WhatsApp editables viven en pestañas `*_Copys` del Sheet (columnas `Momento, Copy_Texto, Numero_WA, Activo`), nunca hardcodeados en el HTML. El texto sustituye variables tipo `{Nombre}`, `{Mes}`, `{Fechas}` antes de armar el link. Los emojis se quitaron de todos los copys porque no siempre llegaban bien por WhatsApp (ver `limpiarEmojisCopys()` en `apps-script.gs`) — no reintroducir emojis en copys nuevos.
- `Requiere_Seguimiento`: el significado varía por iniciativa — en Atención Psicológica ya no se usa para nada (el botón de seguimiento siempre está disponible); en otras sí determina si algo aparece resaltado/pendiente. No asumas una regla uniforme; revisa la página específica.
- **Nunca usar `localStorage` ni `sessionStorage`** en ninguna página del panel.
- Bajas: existe un concepto de "baja de empresa" (la persona deja de trabajar en GEB) que dispara una cascada (`handleBajaEmpresa`) marcando Estado/Requiere_Seguimiento en las 7 iniciativas restringidas por baja de empresa: `AP_Inscritos, IG_Inscritos, BE_Inscritos, EA_Lideres, EG_Inscritos, RA_Inscritos, RE_Interesados, CS_Inscritos`. Salvando Vidas y Biblioteca quedan fuera de esa cascada a propósito (alguien puede seguir donando/prestando libros aunque ya no sea colaborador).
- "Reactivar" deshace una baja individual sin tocar las demás iniciativas — patrón repetido como `handleXxReactivar(p)`.

## Diseño visual — sistema compartido en todas las páginas

Referencia canónica original: `salvando-vidas.html` (todas las páginas nuevas deben ser visualmente indistinguibles de ella en proporciones y estructura).

- **Tipografía**: `Cormorant Garamond` (500/600/700) para títulos/headings, `DM Sans` (400–700) para todo el body. Cargadas desde Google Fonts en cada `<head>`.
- **Paleta** (variables CSS `:root`, repetidas en cada archivo):
  - `--blue: #0B2545` (navy, topbar y headings)
  - `--gold: #C49B4C` (acento — borde inferior del page-header, dot del topbar, botones "activo")
  - `--bg: #F7F5F0` (fondo general, beige claro)
  - `--card: #FFFFFF`, `--text: #1C2A3A`, `--text-soft: #5C6B7A`, `--border: #E3E1D9`
  - Pares fondo/texto por estado: `--green-bg/--green-text` (activo/confirmado), `--amber-bg/--amber-text` (seguimiento/pendiente), `--gray-bg/--gray-text` (inactivo), `--red-bg/--red-text` (alertas), `--b-bg/--b-text` (azul info)
  - Tipos de sangre con sus propios pares (`--o-bg/--o-text`, `--a-bg/--a-text`, etc.) en Salvando Vidas
- **Layout repetido en cada página de iniciativa**:
  - `.topbar` navy fijo arriba con el nombre "GEB University — Iniciativas Humanas" + dot dorado
  - `.back-link` "← volver al panel general" hacia `index.html`
  - `header.page-header` con `<h1>` (Cormorant, 34px) + subtítulo + borde inferior dorado de 2px
  - `.metric-grid` (tarjetas blancas con borde, valor grande en navy + label pequeño gris)
  - `.section` + `.section-title` (Cormorant, 20px, navy) para cada bloque
  - Tarjetas de seguimiento individual con `.avatar` circular de iniciales, `.pill` de estado, botones `.wa-btn` (fondo gris claro, ícono verde WhatsApp) que abren un panel con el copy y botón "Abrir en WhatsApp"
  - `.inline-form` (fondo `#FAF9F5`, oculto por default, toggle con clase `.open`) para altas/registros/ediciones — mismo patrón de `form-row` + `form-actions` (`.btn-guardar` navy, `.btn-cancelar` outline) + `.form-msg` de éxito/error
  - Tablas con `.table-wrap` + `.filters` (buscador + selects) arriba
  - `.footnote` gris al pie de sección con aclaraciones ("datos y copys desde Google Sheets…")
- **Journey map SVG**: cada página de iniciativa tiene un diagrama de flujo del colaborador en SVG puro (`viewBox="0 0 680 270"` en las páginas más recientes; algunas antiguas usan `630 200`), cajas redondeadas ~72×40 con `<title>`/`<desc>` accesibles, flechas con `marker` reutilizable, y colores por caja tomados de los pares fondo/texto de arriba.
- `index.html` (panel general) reutiliza el mismo sistema pero con su propio layout de "áreas" (columnas por enfoque) + buscador global de colaboradores + gráficas de barras apiladas simples (sin librería, divs con `width:%`) para el desglose por iniciativa.

## Estructura del repo

| Archivo | Qué es |
|---|---|
| `index.html` | Panel general: métricas globales, buscador de colaboradores (lee `Colaboradores`), desglose por iniciativa, grid de 10 iniciativas agrupadas en 4 "enfoques" (Bienestar Integral, Impacto Comunitario, Movilidad Social, Habilidades para la Vida), botón "+ Dar de alta" |
| `altas.html` | Alta unificada: un solo formulario con datos generales de la persona + checkboxes de "¿a qué iniciativas entra?" (cada una revela sus campos propios); manda todo en un solo POST (`alta_unificada`) que crea filas en cada pestaña marcada |
| `atencion-psicologica.html` | Iniciativa Atención Psicológica (ELA e Impulso43): inscritos, sesiones, pase de lista en lote, diagnóstico inicial/final, bajas/reactivaciones |
| `retiro-espiritual.html` | Retiro del Espíritu Santo: ediciones (alta y edición), pase de lista en lote, habitaciones editables en el panel, lista de espera/interesados, historial por edición — **y una sección adicional para Camino de Santiago** (comparte página, es la 10ª iniciativa sin card propia en `index.html`) |
| `salvando-vidas.html` | Donadores de sangre: guardias del mes, historial de donaciones, marcar atendido/baja |
| `eco-accion.html` | Eco-Acción: líderes y registro de cilindros (`EA_Lideres`, `EA_Cilindros`) |
| `escuela-geb.html` | Escuela GEB (INEA/UVL): inscritos, asistencias, exámenes |
| `impulso-geb.html` | Impulso GEB: inscritos, paso de generación |
| `beca-educativa.html` | Beca Educativa GEB: inscritos, revisión de inscripción |
| `reto-ahorro.html` | Reto Ahorro: semana actual (individual y en lote), copy editable desde el panel, alertas automáticas de atraso |
| `biblioteca-geb.html` | Biblioteca Comunitaria: préstamos, donaciones, devoluciones, recibos y etiquetas en PDF, sync con Sheet externo de físicos |
| `apps-script.gs` | Único código con permiso de escritura. `doPost(e)` rutea por `tipo` a un handler por acción; también contiene utilidades de desarrollo (`borrarDatosPruebaTodos`, `limpiarEmojisCopys`) colgadas del menú personalizado "GEB CRM" del Sheet (`onOpen()`) |
| `INSTRUCCIONES_CRM.md` | Documento de referencia para el equipo (no técnico): qué pestaña llena quién, bugs ya corregidos, reglas para no romper el Sheet. Consúltalo antes de tocar el mapeo de columnas de cualquier pestaña — suele estar más actualizado que la memoria de una sesión vieja |
| `GEB_Sheet_Nuevo_Unificado.xlsx` | Copia local de referencia del Sheet nuevo unificado (no está en control de versiones activo por decisión del usuario, no lo genera ni lo consume el panel) |

## Preferencias del usuario

- Cecilia no es técnica: reporta bugs describiendo lo que ve mal en el panel o en el Sheet, no con errores de consola. Explica los hallazgos en esos términos.
- Prefiere que se le muestre primero un mockup en texto de las secciones de una página nueva antes de escribir el HTML, y dar su aprobación antes de construir.
- Da feedback de diseño concreto y espera que se itere sobre lo ya construido en vez de rehacerlo desde cero (ej. pidió cambiar de "registrar asistencia una por una" a "pase de lista en lote" sobre un mockup ya aprobado).
- Pide explícitamente el commit/push a GitHub cuando quiere que el cambio quede desplegado — no asumir que hay que hacerlo automáticamente sin que lo pida.
- Este entorno de trabajo no tiene Python ni Node instalados — no se puede levantar el servidor de preview local (`.claude/launch.json` usa `python -m http.server`) ni correr scripts de esas plataformas; la verificación de cambios de código a veces solo puede hacerse por lectura/lint manual, no ejecutando el sitio.
