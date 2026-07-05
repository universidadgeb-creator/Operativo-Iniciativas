# Instrucciones del CRM GEB Iniciativas Humanas

Documento de referencia para todo el equipo que usa o administra el panel y el Sheet maestro. Última revisión: 2026-07-03.

## 1. Cómo funciona la conexión Portal ↔ Sheets

El panel (las páginas `.html` en GitHub Pages) **nunca modifica el Sheet directamente**. Funciona en dos vías separadas:

- **Lectura**: cada página descarga los datos del Sheet en tiempo real usando un enlace público de Google llamado "gviz" (no requiere contraseña ni clave). Por eso el Sheet debe seguir compartido como "Cualquier persona con el enlace puede ver".
- **Escritura**: cuando alguien usa un botón del portal (ej. "Registrar donación", "Guardar reacción", "Pase de lista"), el portal manda esa información a un programa intermedio llamado **Google Apps Script Web App** (el archivo `apps-script.gs`), y ese programa es el único que efectivamente escribe en el Sheet.

**Regla de oro:** el código del portal asume que cada columna está en una posición fija (columna 1, 2, 3...) en cada pestaña. Si alguien **inserta, borra, mueve o reordena una columna a mano** en una pestaña que el portal usa, el portal seguirá leyendo por posición y mostrará (o escribirá) el dato equivocado — sin ningún aviso de error. Esto ya pasó varias veces (ver sección 4) y es la causa más común de "se ve mal" o "se movió la información".

Si necesitas reordenar o insertar una columna, **avisa antes** a quien mantiene el código, para actualizar el portal al mismo tiempo.

## 2. Qué se edita en el portal y qué se edita en Sheets, por iniciativa

| Iniciativa | Pestaña | Quién la llena | Notas |
|---|---|---|---|
| **Atención Psicológica** | `AP_Inscritos` | Híbrida — alta y datos de inscripción son manuales; el portal solo actualiza Estado y Sesiones_Tomadas | Columna `Activo_Programa` (D, reemplaza a `Fecha_Registro` que nunca se usó) — **100% manual, la llena el equipo con "Activo", "Inactivo" o "Lista de espera"**. Decide si esa persona cuenta en las vistas/métricas (licencias ocupadas, activos, espera...); "Inactivo" o vacío la oculta de todo. Columna `Estado` (H) queda **exclusiva del script** — solo puede tener vacío, "Sesión pendiente" o "Completó"; nadie debe escribirla a mano nunca. "Contactado"/"En seguimiento" de Impulso43 ya no se guardan en ningún lado — el portal los calcula solo mirando si ya se registraron sesiones. Columna `Requiere_Seguimiento` (L) **ya no se usa para nada** — el copy de seguimiento siempre está disponible como botón manual en el panel; puedes dejarla vacía siempre (no se borró del Sheet para no correr Notas de lugar). **Hay que renombrar la columna D del Sheet real a `Activo_Programa` y llenarla (Activo/Inactivo/Lista de espera) para cada inscrito, o todos quedarán invisibles en las vistas operativas.** |
| | `AP_Sesiones` | Híbrida — el portal registra asistencia/cambios de modalidad; Fecha_Programada y Hora_Programada son manuales | |
| | `AP_Copys` | 100% manual | Textos de WhatsApp, edítalos directo en el Sheet |
| **Retiro Espiritual** | `RE_Ediciones` | Solo portal (alta de nuevas ediciones) | |
| | `RE_Asistencias` | Portal (alta + pase de lista) | El campo `edicion_label` nunca se llena, ignóralo |
| | `RE_Habitaciones` | Solo portal | Cada vez que se guarda, se reescribe TODO el acomodo de esa edición — no edites a mano mientras alguien está armando habitaciones en el portal |
| | `RE_Copys` | 100% manual | |
| **Salvando Vidas** | `SV_Donadores` | 100% manual | ⚠️ Ver hallazgo urgente en sección 4 |
| | `SV_Historial` | Híbrida — el portal registra donaciones nuevas | |
| | `SV_Copys` | 100% manual | |
| **Escuela GEB** | `EG_Inscritos` | 100% manual | ⚠️ Ver hallazgo urgente en sección 4 |
| | `EG_Asistencias` | Solo portal | |
| | `EG_Examenes` | Solo portal | Aún sin registros |
| | `EG_Copys` | 100% manual | |
| **Reto Ahorro** | `RA_Inscritos` | Híbrida — alta manual; el portal solo actualiza Semana_Actual | ⚠️ Ver hallazgo urgente en sección 4 |
| | `RA_Semanas` | Solo portal | Aún sin registros |
| | `RA_Copys` | 100% manual | El texto de "Copy semanal" trae la semana y el monto escritos a mano — hay que actualizarlo cada semana |
| **Biblioteca GEB** | `BIB_Prestamos` | Híbrida — el alta llega de un Google Form externo; el cierre de devolución lo hace el portal | |
| | `BIB_Donaciones` | Híbrida — alta por portal; recibo/etiqueta generados por portal | |
| | `BIB_Devoluciones` | Híbrida — reporte llega de un Google Form externo; el portal marca "procesado" | |
| | `BIB_Copys` | 100% manual | |
| | Hoja externa "Fisicos" | Portal (préstamos/disponibilidad) | Sheet distinto al maestro — más frágil ante reordenamientos manuales, evítalos ahí |
| **Concentrado** | `Concentrado` | 100% manual, por nombre | Ya reducida a columnas Sí/No + datos básicos de roster (19 columnas). Ver sección 3 |

## 3. La pestaña Concentrado

Es el censo general de colaboradores con una columna Sí/No por iniciativa. Estructura real actual (19 columnas, A→S):

`Nombre, Rol, Sucursal, Unidad, Correo, FechaNac, Género, Tel, Turno, Psic, Retiro, Impulso, Diagnóstico educativo, Escuela GEB, Donadores, Ahorro, Biblioteca Donante, Biblioteca Receptor, Camino Santiago`

- La columna "Escuela GEB" puede tener 3 valores: Sí / No / **Historico** (alguien que ya no está activo pero se quiere conservar su diagnóstico educativo).
- Para reordenar Concentrado sin desligar las columnas Sí/No de la persona correcta, usa el menú **GEB CRM → 🔒 Reordenar Concentrado por Nombre (seguro)** dentro del propio Sheet (Extensiones → Apps Script). Nunca ordenes seleccionando solo algunas columnas — eso es lo que descuadra todo.

## 4. Corrupción de altas (RA_Inscritos / SV_Donadores / EG_Inscritos / AP_Inscritos) — RESUELTO

Se encontró el mismo patrón de corrupción en 4 pestañas: `AP_Inscritos`, `RA_Inscritos`, `SV_Donadores`, `EG_Inscritos`. Las columnas Nombre/Sucursal/Teléfono se llenaban con una fórmula QUERY que jalaba de Concentrado (derramada sobre varias filas), mientras que las columnas manuales (Estado, Urgencia, Requiere_Seguimiento, etc.) se llenan a mano ancladas a un número de fila fijo. Cada vez que el orden de Concentrado cambiaba, la fórmula recorría Nombre/Sucursal/Teléfono pero las columnas manuales se quedaban pegadas a la persona equivocada. En `RA_Inscritos`, `SV_Donadores` y `EG_Inscritos` el problema era peor: sus fórmulas referenciaban columnas de Concentrado que ya ni existen (se perdieron cuando Concentrado se simplificó), mostrando `#VALUE!` directo.

**Arreglo aplicado (`apps-script.gs`):** en las 4 pestañas, Nombre/Sucursal/Teléfono ahora son texto fijo que nunca se mueve. La fórmula QUERY (corregida, apuntando al flag vigente de Concentrado: Psic=J para AP, Ahorro=P para Reto Ahorro, Donadores=O para Salvando Vidas, EscuelaGEB=N para Escuela GEB) vive en una columna oculta (P:R de cada pestaña), y una función revisa cada hora esa columna oculta contra los nombres que ya existen, agregando los nuevos **al final** (nunca en medio) con `Requiere_Seguimiento = "Sí"`.

**Acción pendiente del equipo:** entrar al Sheet y dar clic, una sola vez cada uno, en los 4 ítems del menú **GEB CRM → 🔧 Configurar sincronización … (1 vez)** (uno por AP, Reto Ahorro, Salvando Vidas, Escuela GEB) — después de que alguien redespliegue `apps-script.gs` en el editor de Apps Script. En Reto Ahorro/Salvando Vidas/Escuela GEB no había datos operativos que rescatar (se confirmó con el equipo), así que ese botón además agrega de inmediato a todas las personas que hoy cumplen el flag correspondiente en Concentrado, marcadas con Requiere_Seguimiento="Sí" para que se revisen una por una.

## 5. Bugs de código ya corregidos en esta revisión

- `apps-script.gs`: la asistencia de Retiro Espiritual escribía "Sí" y el id de edición en las columnas equivocadas de Concentrado (Impulso GEB y Diagnóstico educativo en vez de Retiro), corrompiendo esos datos cada vez que se pasaba lista. Corregido.
- `index.html`, `escuela-geb.html`, `retiro-espiritual.html`: tenían el esquema de columnas de Concentrado desactualizado (de cuando tenía más columnas), causando que varias tarjetas/badges mostraran el dato de la iniciativa vecina. Corregido y verificado contra el Sheet real.
- `escuela-geb.html`: la sección "Prospectos" mostraba datos de Biblioteca Donante/Receptor disfrazados de prospectos de Beca Educativa (esas columnas ya no existen). Ahora muestra honestamente "Sin prospectos" — pendiente decidir si se elimina esa sección por completo (ya no tiene fuente de datos).
- `escuela-geb.html`: el contador de asistencias en las tarjetas de INEA usaba un nombre de campo (`Asistentes`) que no existe; siempre marcaba 0/Y. Corregido a `Asistio`.
- `escuela-geb.html`: el esquema de `EG_Copys` no contemplaba la columna real `Numero_WA`, lo que probablemente hacía que el botón de WhatsApp enviara el texto "INEA"/"UVL"/"Ambas" en vez del mensaje real. Corregido.
- `biblioteca-geb.html`: el esquema de `BIB_Prestamos` tenía una columna de más (`fecha_prestamo`, que no existe como columna separada), lo que corría todos los campos siguientes una posición y rompía en silencio el cálculo de préstamos atrasados/por vencer. Corregido.

## 6. Hallazgos menores (no bloqueantes, para revisar cuando haya tiempo)

- **Salvando Vidas**: la fecha de donación en `SV_Historial` se muestra como número crudo (ej. "46185") en vez de fecha legible — falta aplicarle la misma conversión que ya tiene Biblioteca.
- **Salvando Vidas**: inconsistencia en el manejo de `Requiere_Seguimiento` vacío — en unas partes del código cuenta como "Sí" y en otras no.
- **Atención Psicológica**: el registro de sesiones de Impulso43 no avanza el `Estado` automáticamente (sí lo hace el de ELA) — hay que cambiarlo a mano en Sheets.
- **Atención Psicológica**: `AP_Inscritos` columnas A-C (Nombre/Sucursal/Teléfono) ya NO son una fórmula QUERY en vivo — se congelaron como texto fijo porque la fórmula se desalineaba con las columnas manuales (D-M) cada vez que Concentrado cambiaba de orden. La fórmula ahora vive en la columna oculta P, y la función `sincronizarNuevosAP` (dispara cada hora) agrega los nombres nuevos al final de la tabla automáticamente, marcando `Requiere_Seguimiento = "Sí"`. Ver menú **GEB CRM → Configurar sincronización AP (1 vez)** para (re)activar esto si algún día se rompe el disparador.
- **Reto Ahorro**: el copy "Anuncio de bono" está disponible en cualquier semana, aunque el texto menciona la semana 25 — hay que recordar no enviarlo antes de tiempo.
- **Biblioteca GEB**: algunas filas antiguas de `BIB_Donaciones` tienen texto tipo "📄 Ver recibo" en vez de una URL real en el link — hay que regenerarlas una vez con el botón existente.
- **Biblioteca GEB**: la hoja externa "Fisicos" usa columnas por posición (no por nombre de encabezado, a diferencia del resto de Biblioteca) — más frágil si alguien la reordena a mano.
- **Retiro Espiritual**: si el nombre en `RE_Asistencias` no coincide exacto con el de Concentrado, el filtro de género al asignar habitaciones no se aplica.

## 7. Reglas para el equipo (varias personas usando el mismo Sheet)

1. Nunca insertes, borres ni reordenes columnas en una pestaña que el portal usa sin avisar antes — revisa la tabla de la sección 2 para saber si una pestaña es "solo portal" (más sensible) o "100% manual" (más segura de tocar).
2. Si necesitas reordenar filas en Concentrado, usa el menú **GEB CRM → Reordenar Concentrado por Nombre (seguro)** en vez de "Ordenar rango" a mano.
3. Si vas a llenar una columna de "Requiere_Seguimiento" (Salvando Vidas, Reto Ahorro, Atención Psicológica), nunca la dejes vacía a mano sin razón — escribe "No" explícito si no aplica. En Atención Psicológica el script la administra solo (Sí en altas nuevas, vacío cuando ya se dio seguimiento desde el panel, No al marcar Baja); no la edites a mano salvo para corregir un caso puntual.
4. Si el panel muestra algo "vacío" o "en cero" que debería tener datos, antes de asumir que es un bug de código, revisa si la columna A de esa pestaña tiene un error (`#VALUE!`, `#REF!`) — es la señal de que se rompió una fórmula o se corrió una columna.
5. Cualquier cambio a `apps-script.gs` requiere que alguien lo copie al editor de Apps Script y cree una nueva versión de implementación (Implementar → Administrar implementaciones → editar el lápiz → Nueva versión → Implementar) — el archivo del repo por sí solo no actualiza nada.
