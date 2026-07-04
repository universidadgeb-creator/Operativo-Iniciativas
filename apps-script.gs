// ================================================================
// GEB Iniciativas Humanas — Web App
// Despliega como: Ejecutar como "Yo", Acceso "Cualquier persona"
// ================================================================

const SS = SpreadsheetApp.openById("1fS5qeoB1ViuCUP4HOQ1zTJgh4tfWUkObvb5LMr3to5A");

// Biblioteca Virtual externa (solo lectura/escritura de la pestaña "Fisicos")
const SHEET_BIBLIOTECA_ID = "1FDZB3aR-YAyVMsiAjo92PuUdH0iTfBmreCWv5DvCpdM";
const LOGO_BIBLIO_ID = "1NqoFmESlsTP4dpFscYglcs9o6THQP4TR";
const LOGO_UGEB_ID = "1mBIHoKyngoa7cBiSvHCVY0x_pIkFyARJ";
const COLOR_PRIMARIO = "#185FA5";

// ── Menú de utilidades para evitar el desligue de filas en Concentrado ──
// Causa típica: alguien selecciona solo algunas columnas (ej. A:I) al ordenar
// o filtrar en Sheets, y las columnas J:S (Sí/No por iniciativa) no se mueven
// con el resto de la fila, quedando la información de otra persona. Este
// menú ordena SIEMPRE el rango completo de la fila, así nunca se desliga.
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("GEB CRM")
    .addItem("🔒 Reordenar Concentrado por Nombre (seguro)", "reordenarConcentradoPorNombre")
    .addItem("🔧 Configurar sincronización AP (1 vez)", "configurarSincronizacionAP")
    .addItem("🔧 Configurar sincronización Reto Ahorro (1 vez)", "configurarSincronizacionRA")
    .addItem("🔧 Configurar sincronización Salvando Vidas (1 vez)", "configurarSincronizacionSV")
    .addItem("🔧 Configurar sincronización Escuela GEB (1 vez)", "configurarSincronizacionEG")
    .addItem("🔧 Configurar alertas de atraso Reto Ahorro (1 vez)", "configurarAlertasAtrasoRA")
    .addToUi();
}

// ── Atención Psicológica: Requiere_Seguimiento (col L de AP_Inscritos) ──
// Flujo: vacío por default. Cuando alguien captura un registro nuevo a mano
// (llena la col A=Nombre de una fila que antes estaba vacía), este trigger
// simple lo marca "Sí" automáticamente para que el panel lo resalte como
// pendiente. Se vuelve a vaciar desde las acciones del panel (asistencia,
// cambio de modalidad, sesiones Impulso43) porque eso ya cuenta como
// seguimiento dado. Se marca "No" solo cuando alguien se da de baja
// (ver handleApBaja) para diferenciar "nunca se le dio seguimiento" de
// "ya no aplica seguimiento".
function onEdit(e) {
  try {
    var sheet = e.range.getSheet();
    if (sheet.getName() !== "AP_Inscritos") return;
    if (e.range.getColumn() !== 1) return; // solo cuando se llena Nombre (col A)
    var row = e.range.getRow();
    if (row === 1) return; // encabezado
    if (!e.range.getValue()) return; // se borró el nombre, no es alta
    if (e.oldValue) return; // ya tenía nombre antes: no es fila nueva
    var seguimientoCell = sheet.getRange(row, 12); // Requiere_Seguimiento (col L)
    if (!seguimientoCell.getValue()) seguimientoCell.setValue("Sí");
  } catch (err) {
    // onEdit no debe romper la hoja si algo falla
  }
}

// ── Sincronización automática de altas (AP, Reto Ahorro, Salvando Vidas, Escuela GEB) ──
// Las 4 pestañas comparten el mismo problema de fondo: Nombre/Sucursal/Teléfono
// (siempre las primeras 3 columnas, A:C) se llenaban con una fórmula QUERY que
// jala de Concentrado (derramada sobre varias filas), mientras que las columnas
// siguientes (Estado, Urgencia, Semana_Actual, Requiere_Seguimiento, etc.) se
// llenan a mano ancladas a un número de fila fijo. Cada vez que el orden de
// Concentrado cambiaba, o la fórmula quedaba apuntando a columnas que ya no
// existen (Concentrado se ha restructurado varias veces), A:C se desalineaba de
// esas columnas manuales o de plano se rompía mostrando #VALUE!. Un recálculo de
// fórmula tampoco dispara onEdit, así que Requiere_Seguimiento nunca se marcaba
// "Sí" para altas nuevas. IMPORTANTE: las 4 pestañas resultaron tener datos
// operativos reales en sus columnas manuales (incluso RA_Inscritos y
// EG_Inscritos, que parecían vacías a simple vista) — nunca asumir que una
// pestaña "no tiene nada que perder" sin confirmarlo primero con el equipo.
//
// Arreglo: en cada pestaña, el bloque de "identidad" que viene de Concentrado
// (Nombre/Sucursal/Teléfono — y en Escuela GEB también Diagnóstico Educativo,
// col D) se congela como texto fijo que nunca se mueve solo. La fórmula QUERY
// (corregida, apuntando al flag vigente de Concentrado: Psic=J, Ahorro=P,
// Donadores=O, EscuelaGEB=N) vive en una columna oculta (P en adelante), y una
// función corre cada hora comparando esa columna contra los nombres que ya
// existen en col A. Los nombres nuevos se agregan como fila nueva AL FINAL
// (nunca en medio), con Requiere_Seguimiento = "Sí".
var SYNC_AP = {
  sheet: "AP_Inscritos", numCols: 13, identityCols: 3, reqCol: 11, helperCol: 16,
  triggerFnName: "sincronizarNuevosAP",
  formula: '=QUERY(Concentrado!A1:S989,"SELECT A,C,H WHERE J=\'Sí\'",1)',
};
// RA_Inscritos real solo tiene 6 columnas propias (Nombre,Sucursal,Telefono_WA,
// Semana_Actual,Fecha_Inscripcion,Requiere_Seguimiento) — reqCol:6 apuntaba a una
// 7ª columna que nunca existió (una "Estado" fantasma se había insertado a la
// mitad en el SCHEMA del frontend, no aquí). Corregido a reqCol:5 (Requiere_
// Seguimiento, índice real) y numCols:8 para reservar Estado/Fecha_Ultima_Semana,
// agregadas al final — ver reto-ahorro.html.
var SYNC_RA = {
  sheet: "RA_Inscritos", numCols: 8, identityCols: 3, reqCol: 5, helperCol: 16,
  triggerFnName: "sincronizarNuevosRA",
  formula: '=QUERY(Concentrado!A1:S989,"SELECT A,C,H WHERE P=\'Sí\'",1)',
};
var SYNC_SV = {
  sheet: "SV_Donadores", numCols: 7, identityCols: 3, reqCol: 6, helperCol: 16,
  triggerFnName: "sincronizarNuevosSV",
  formula: '=QUERY(Concentrado!A1:S989,"SELECT A,C,H WHERE O=\'Sí\'",1)',
};
var SYNC_EG = {
  // EscuelaGEB (Sí/No/Historico) vive en la columna N de Concentrado. El
  // encabezado de esa columna en el Sheet dice "" y el de M dice "ESCUELA GEB"
  // por un rótulo mal puesto — verificado contra los DATOS reales (N trae
  // Sí/No limpios, M trae diagnóstico educativo tipo "Licenciatura"/"Primaria").
  // Diagnostico_Edu (col D de EG_Inscritos) también viene de Concentrado (col M),
  // así que se incluye como 4ª columna de identidad a proteger y sincronizar.
  sheet: "EG_Inscritos", numCols: 9, identityCols: 4, reqCol: 8, helperCol: 16,
  triggerFnName: "sincronizarNuevosEG",
  formula: '=QUERY(Concentrado!A1:S989,"SELECT A,C,H,M WHERE N=\'Sí\'",1)',
};

// El número de columnas de identidad varía por pestaña (config.identityCols):
// 3 en AP/RA/SV (Nombre,Sucursal,Telefono), 4 en EG (+ Diagnostico_Edu).
function _sincronizarAltas(config) {
  var ws = SS.getSheetByName(config.sheet);
  if (!ws) return;

  var lastRow = ws.getLastRow();
  var nombresExistentes = {};
  if (lastRow >= 2) {
    var actuales = ws.getRange(2, 1, lastRow - 1, 1).getValues();
    actuales.forEach(function (r) {
      var n = String(r[0] || "").trim().toLowerCase();
      if (n) nombresExistentes[n] = true;
    });
  }

  var filasFuente = Math.max(ws.getLastRow() - 1, 0);
  if (filasFuente === 0) return;
  var fuente = ws.getRange(2, config.helperCol, filasFuente, config.identityCols).getValues();

  fuente.forEach(function (r) {
    var nombre = String(r[0] || "").trim();
    if (!nombre) return;
    var key = nombre.toLowerCase();
    if (nombresExistentes[key]) return;

    var fila = new Array(config.numCols).fill("");
    for (var i = 0; i < config.identityCols; i++) fila[i] = r[i] || "";
    fila[config.reqCol] = "Sí"; // Requiere_Seguimiento: alta nueva
    ws.appendRow(fila);
    nombresExistentes[key] = true;
  });
}

function _asegurarTrigger(fnName) {
  var yaExiste = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === fnName;
  });
  if (!yaExiste) ScriptApp.newTrigger(fnName).timeBased().everyHours(1).create();
}

function sincronizarNuevosAP() { _sincronizarAltas(SYNC_AP); }
function sincronizarNuevosRA() { _sincronizarAltas(SYNC_RA); }
function sincronizarNuevosSV() { _sincronizarAltas(SYNC_SV); }
function sincronizarNuevosEG() { _sincronizarAltas(SYNC_EG); }

// Configuración de una sola vez, igual para las 4 pestañas: SIEMPRE congela lo
// que hoy se ve en el bloque de identidad (A:C, o A:D en Escuela GEB) —sea
// resultado de fórmula o ya texto plano— antes de tocar nada, así nunca se
// arriesga a perder datos reales por asumir que una pestaña estaba vacía. Se
// captura y se vuelve a escribir el mismo bloque completo (incluyendo la fila
// 1, por si el encabezado también era parte de una fórmula derramada) —
// funciona sin importar si la fórmula estaba anclada en la fila 1 o en la fila 2.
// Ejecutar desde el menú "GEB CRM → Configurar sincronización … (1 vez)".
function _configurarSincronizacion(config, etiqueta) {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    "Configurar sincronización " + etiqueta,
    "Esto va a: 1) fijar los datos actuales de identidad (Nombre/Sucursal/Teléfono" +
      (config.identityCols > 3 ? "/Diagnóstico Educativo" : "") + ") de " + etiqueta + " como texto fijo, " +
      "2) mover la fórmula QUERY (corregida) a una columna oculta, y 3) activar la revisión automática " +
      "cada hora, agregando SOLO los nombres nuevos al final. Ninguna otra columna (Estado, " +
      "Requiere_Seguimiento, Semana_Actual, etc.) se toca. ¿Continuar?",
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  var ws = SS.getSheetByName(config.sheet);
  if (!ws) { ui.alert("No se encontró la pestaña " + config.sheet + "."); return; }

  var lastRow = ws.getLastRow();
  if (lastRow >= 1) {
    var rango = ws.getRange(1, 1, lastRow, config.identityCols); // incluye posible encabezado
    var valoresActuales = rango.getValues();
    rango.clearContent(); // quita cualquier fórmula derramada (ancle en fila 1 o 2)
    rango.setValues(valoresActuales); // los vuelve a poner como texto fijo, igual a como estaban
  }

  ws.getRange(2, config.helperCol).setFormula(config.formula);
  ws.hideColumns(config.helperCol, config.identityCols);
  _asegurarTrigger(config.triggerFnName);

  ui.alert("Listo. Revisa que los datos de identidad de " + etiqueta + " se sigan viendo bien. Si alguna " +
    "alta reciente ya existía antes de este cambio, agrégale \"Sí\" a mano en Requiere_Seguimiento (la " +
    "sincronización no la detecta porque su nombre ya existe).");
}

function configurarSincronizacionAP() { _configurarSincronizacion(SYNC_AP, "AP_Inscritos"); }
function configurarSincronizacionRA() { _configurarSincronizacion(SYNC_RA, "RA_Inscritos"); }
function configurarSincronizacionSV() { _configurarSincronizacion(SYNC_SV, "SV_Donadores"); }
function configurarSincronizacionEG() { _configurarSincronizacion(SYNC_EG, "EG_Inscritos"); }

function reordenarConcentradoPorNombre() {
  const ws = SS.getSheetByName("Concentrado");
  if (!ws) { SpreadsheetApp.getUi().alert("No se encontró la pestaña Concentrado."); return; }
  const lastRow = ws.getLastRow();
  const lastCol = ws.getLastColumn();
  if (lastRow < 3) return;
  // Ordena A2:<lastCol><lastRow> como un solo bloque por la columna A (Nombre),
  // así todas las columnas Sí/No de cada persona viajan siempre con su fila.
  ws.getRange(2, 1, lastRow - 1, lastCol).sort({ column: 1, ascending: true });
  SpreadsheetApp.getUi().alert("Concentrado reordenado por Nombre sin perder las columnas Sí/No.");
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const tipo    = payload.tipo;

    if (tipo === "asistencia")         return handleAsistencia(payload);
    if (tipo === "examen")             return handleExamen(payload);
    if (tipo === "donacion")           return handleDonacion(payload);
    if (tipo === "ap_asistencia")      return handleApAsistencia(payload);
    if (tipo === "ap_cambioModalidad") return handleApCambioModalidad(payload);
    if (tipo === "ap_sesionesI43")     return handleApSesionesI43(payload);
    if (tipo === "ap_baja")           return handleApBaja(payload);
    if (tipo === "ap_diagnostico")    return handleApDiagnostico(payload);
    if (tipo === "sv_atendido")       return handleSvAtendido(payload);
    if (tipo === "sv_baja")           return handleSvBaja(payload);
    if (tipo === "ra_actualizarSemana") return handleRaActualizarSemana(payload);
    if (tipo === "ra_marcarAtendido")   return handleRaMarcarAtendido(payload);
    if (tipo === "re_altaEdicion")           return handleReAltaEdicion(payload);
    if (tipo === "re_registro")             return handleReRegistro(payload);
    if (tipo === "re_asistenciaLote")        return handleReAsistenciaLote(payload);
    if (tipo === "re_guardarHabitaciones")   return handleReGuardarHabitaciones(payload);
    if (tipo === "bib_altaDonacion")        return handleBibAltaDonacion(payload);
    if (tipo === "bib_procesarDevolucion")  return handleBibProcesarDevolucion(payload);
    if (tipo === "bib_extenderPrestamo")    return handleBibExtenderPrestamo(payload);
    if (tipo === "bib_generarRecibo")       return handleBibGenerarRecibo(payload);
    if (tipo === "bib_generarEtiqueta")     return handleBibGenerarEtiqueta(payload);
    if (tipo === "bib_sincronizarFisicos")  return handleBibSincronizarFisicos(payload);

    return resp({ ok: false, error: "Tipo desconocido: " + tipo });
  } catch(err) {
    return resp({ ok: false, error: err.message });
  }
}

// ── Asistencia INEA ─────────────────────────────────────────────
// Cols: A=Nombre, B=Fecha_Clase, C=Tema, D=Asistio, E=Notas
function handleAsistencia(p) {
  const ws = SS.getSheetByName("EG_Asistencias");
  if (!ws) return resp({ ok: false, error: "Pestaña EG_Asistencias no encontrada" });
  const nombres = (p.asistentes || []);
  nombres.forEach(function(nombre) {
    ws.appendRow([
      nombre,
      p.fecha  || "",
      p.tema   || "",
      "Sí",
      p.notas  || ""
    ]);
  });
  return resp({ ok: true });
}

// ── Examen ──────────────────────────────────────────────────────
// Cols: A=Nombre, B=Modalidad, C=Tipo, D=Tema_Materia, E=Fecha_Examen, F=Resultado, G=Calificacion, H=Notas
function handleExamen(p) {
  const ws = SS.getSheetByName("EG_Examenes");
  if (!ws) return resp({ ok: false, error: "Pestaña EG_Examenes no encontrada" });
  ws.appendRow([
    p.nombre       || "",
    p.modalidad    || "",
    p.tipoExamen   || "",
    p.tema         || "",
    p.fecha        || "",
    p.resultado    || "",
    p.calificacion || "",
    p.notas        || ""
  ]);
  return resp({ ok: true });
}

// ── Donación ────────────────────────────────────────────────────
// Cols: A=Nombre, B=Mes, C=Año, D=Dono, E=Fecha_Donacion, F=Talon_Recibido, G=Notas
function handleDonacion(p) {
  const ws = SS.getSheetByName("SV_Historial");
  if (!ws) return resp({ ok: false, error: "Pestaña SV_Historial no encontrada" });
  ws.appendRow([
    p.nombre        || "",
    p.mes           || "",
    p.año           || "",
    p.dono          || "",
    p.fechaDonacion || "",
    p.talon         || "",
    p.notas         || ""
  ]);
  return resp({ ok: true });
}

// ── Salvando Vidas: Requiere_Seguimiento (col G de SV_Donadores) ──
// Mismo patrón que AP_Inscritos: "Sí" lo pone _sincronizarAltas cuando es
// alta nueva. Se limpia (marcarAtendido) cuando ya se le dio seguimiento
// desde el panel. "No" se pone solo al dar de baja (handleSvBaja).
function handleSvAtendido(p) {
  var ws = SS.getSheetByName("SV_Donadores");
  if (!ws) return resp({ ok: false, error: "Pestaña SV_Donadores no encontrada" });

  var nombre = (p.nombre || "").trim();
  var data   = ws.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
      ws.getRange(i + 1, 7).setValue(""); // Requiere_Seguimiento (col G)
      return resp({ ok: true });
    }
  }

  return resp({ ok: false, error: "No se encontró a " + nombre + " en SV_Donadores" });
}

// ── Salvando Vidas: Marcar de baja ───────────────────────────────
// El donador ya no participa pero se conserva su historial de donaciones
// para indicadores. Estado (col F) pasa a "Inactivo" y Requiere_Seguimiento
// (col G) se marca "No".
function handleSvBaja(p) {
  var ws = SS.getSheetByName("SV_Donadores");
  if (!ws) return resp({ ok: false, error: "Pestaña SV_Donadores no encontrada" });

  var nombre = (p.nombre || "").trim();
  var data   = ws.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
      ws.getRange(i + 1, 6).setValue("Inactivo"); // Estado (col F)
      ws.getRange(i + 1, 7).setValue("No");        // Requiere_Seguimiento (col G)
      return resp({ ok: true });
    }
  }

  return resp({ ok: false, error: "No se encontró a " + nombre + " en SV_Donadores" });
}

// ── Atención Psicológica: Registrar asistencia ──────────────────
// Cols AP_Sesiones: A=Nombre, B=Num_Sesion, C=Fecha_Programada, D=Hora_Programada,
//                   E=Modalidad, F=Asistio, G=Cambio_Modalidad, H=Notas_Sesion
function handleApAsistencia(p) {
  var ws = SS.getSheetByName("AP_Sesiones");
  if (!ws) return resp({ ok: false, error: "Pestaña AP_Sesiones no encontrada" });

  var nombre    = (p.nombre || "").trim();
  var numSesion = String(p.numSesion || "").trim();
  var data      = ws.getDataRange().getValues();
  var found     = false;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()
        && String(data[i][1]).trim() === numSesion) {
      ws.getRange(i + 1, 6).setValue(p.asistio || "");
      ws.getRange(i + 1, 8).setValue(p.notas || "");
      found = true;
      break;
    }
  }

  if (!found) {
    ws.appendRow([nombre, numSesion, "", "", "", p.asistio || "", "", p.notas || ""]);
  }

  // Actualizar Sesiones_Tomadas y Estado en AP_Inscritos
  var wsIns = SS.getSheetByName("AP_Inscritos");
  if (wsIns) {
    var insData = wsIns.getDataRange().getValues();
    for (var j = 1; j < insData.length; j++) {
      if (String(insData[j][0]).trim().toLowerCase() === nombre.toLowerCase()) {
        // Contar sesiones con Asistio = Sí
        var sesData = ws.getDataRange().getValues();
        var count = 0;
        for (var k = 1; k < sesData.length; k++) {
          if (String(sesData[k][0]).trim().toLowerCase() === nombre.toLowerCase()) {
            var a = String(sesData[k][5]).trim().toLowerCase();
            if (a === "sí" || a === "si") count++;
          }
        }
        wsIns.getRange(j + 1, 11).setValue(count); // Sesiones_Tomadas (col K, índice 10)

        var totalSesiones = parseInt(insData[j][9]) || 5;
        if (count >= totalSesiones) {
          wsIns.getRange(j + 1, 8).setValue("Completó");
        } else if (p.asistio === "No") {
          wsIns.getRange(j + 1, 8).setValue("Sesión pendiente");
        }
        wsIns.getRange(j + 1, 12).setValue(""); // Requiere_Seguimiento: ya se dio seguimiento
        break;
      }
    }
  }

  return resp({ ok: true });
}

// ── Atención Psicológica: Cambio de modalidad ───────────────────
function handleApCambioModalidad(p) {
  var ws = SS.getSheetByName("AP_Sesiones");
  if (!ws) return resp({ ok: false, error: "Pestaña AP_Sesiones no encontrada" });

  var nombre    = (p.nombre || "").trim();
  var numSesion = String(p.numSesion || "").trim();
  var data      = ws.getDataRange().getValues();
  var found     = false;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()
        && String(data[i][1]).trim() === numSesion) {
      ws.getRange(i + 1, 5).setValue(p.modalidad || "");
      ws.getRange(i + 1, 7).setValue("Sí");
      found = true;
      break;
    }
  }

  if (!found) {
    ws.appendRow([nombre, numSesion, "", "", p.modalidad || "", "", "Sí", ""]);
  }

  // Ya se dio seguimiento a esta persona desde el panel
  var wsIns2 = SS.getSheetByName("AP_Inscritos");
  if (wsIns2) {
    var insData2 = wsIns2.getDataRange().getValues();
    for (var m = 1; m < insData2.length; m++) {
      if (String(insData2[m][0]).trim().toLowerCase() === nombre.toLowerCase()) {
        wsIns2.getRange(m + 1, 12).setValue("");
        break;
      }
    }
  }

  return resp({ ok: true });
}

// ── Atención Psicológica: Sesiones Impulso43 ────────────────────
// Estado (col H) es exclusivo del script: solo lo mueve a "Completó" al llegar
// al total de sesiones, igual que handleApAsistencia hace para ELA. El resto del
// tiempo se deja vacío — el portal decide "Contactado" vs "En seguimiento" a partir
// de Sesiones_Tomadas, no hace falta que nadie lo escriba a mano.
function handleApSesionesI43(p) {
  var ws = SS.getSheetByName("AP_Inscritos");
  if (!ws) return resp({ ok: false, error: "Pestaña AP_Inscritos no encontrada" });

  var nombre = (p.nombre || "").trim();
  var data   = ws.getDataRange().getValues();
  var sesionesTomadas = parseInt(p.sesionesTomadas) || 0;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
      ws.getRange(i + 1, 11).setValue(sesionesTomadas);
      var totalSesiones = parseInt(data[i][9]) || 5;
      if (sesionesTomadas >= totalSesiones) {
        ws.getRange(i + 1, 8).setValue("Completó");
      }
      ws.getRange(i + 1, 12).setValue(""); // Requiere_Seguimiento: ya se dio seguimiento
      break;
    }
  }

  return resp({ ok: true });
}

// ── Atención Psicológica: Marcar de baja ────────────────────────
// La persona ya no está en la empresa pero se conserva su registro para
// indicadores históricos. Activo_Programa (col D) pasa a "Inactivo" y
// Requiere_Seguimiento (col L) se marca "No" (no aplica más seguimiento).
function handleApBaja(p) {
  var ws = SS.getSheetByName("AP_Inscritos");
  if (!ws) return resp({ ok: false, error: "Pestaña AP_Inscritos no encontrada" });

  var nombre = (p.nombre || "").trim();
  var data   = ws.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
      ws.getRange(i + 1, 4).setValue("Inactivo");  // Activo_Programa (col D)
      ws.getRange(i + 1, 12).setValue("No");        // Requiere_Seguimiento (col L)
      return resp({ ok: true });
    }
  }

  return resp({ ok: false, error: "No se encontró a " + nombre + " en AP_Inscritos" });
}

// ── Atención Psicológica: Diagnóstico inicial/final ─────────────
// Cols N/O de AP_Inscritos (agregadas al final para no correr el resto de
// columnas). Escala numérica de severidad — el panel colorea en verde/
// amarillo/rojo comparando ambos valores, esto solo los guarda tal cual.
function handleApDiagnostico(p) {
  var ws = SS.getSheetByName("AP_Inscritos");
  if (!ws) return resp({ ok: false, error: "Pestaña AP_Inscritos no encontrada" });

  var nombre = (p.nombre || "").trim();
  var data   = ws.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
      ws.getRange(i + 1, 14).setValue(p.diagnosticoInicial || ""); // Diagnostico_Inicial (col N)
      ws.getRange(i + 1, 15).setValue(p.diagnosticoFinal || "");   // Diagnostico_Final (col O)
      return resp({ ok: true });
    }
  }

  return resp({ ok: false, error: "No se encontró a " + nombre + " en AP_Inscritos" });
}

// ── Reto Ahorro: Editar semana actual desde el panel ────────────
// Cols RA_Inscritos: D=Semana_Actual, F=Requiere_Seguimiento, G=Estado,
// H=Fecha_Ultima_Semana (las 2 últimas agregadas al final, ver reto-ahorro.html).
// Actualizar la semana cuenta como seguimiento dado: limpia el flag y reinicia
// el reloj de atraso que usa revisarAtrasosRA().
function handleRaActualizarSemana(p) {
  var ws = SS.getSheetByName("RA_Inscritos");
  if (!ws) return resp({ ok: false, error: "Pestaña RA_Inscritos no encontrada" });

  var nombre = (p.nombre || "").trim();
  var data   = ws.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
      ws.getRange(i + 1, 4).setValue(parseInt(p.semana, 10) || 0); // Semana_Actual (col D)
      ws.getRange(i + 1, 6).setValue("");                          // Requiere_Seguimiento (col F)
      ws.getRange(i + 1, 8).setValue(new Date());                  // Fecha_Ultima_Semana (col H)
      return resp({ ok: true });
    }
  }

  return resp({ ok: false, error: "No se encontró a " + nombre + " en RA_Inscritos" });
}

// ── Reto Ahorro: Marcar atendido (limpia el flag sin cambiar semana) ──
function handleRaMarcarAtendido(p) {
  var ws = SS.getSheetByName("RA_Inscritos");
  if (!ws) return resp({ ok: false, error: "Pestaña RA_Inscritos no encontrada" });

  var nombre = (p.nombre || "").trim();
  var data   = ws.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
      ws.getRange(i + 1, 6).setValue(""); // Requiere_Seguimiento (col F)
      return resp({ ok: true });
    }
  }

  return resp({ ok: false, error: "No se encontró a " + nombre + " en RA_Inscritos" });
}

// ── Reto Ahorro: marcar atraso automático (>3 semanas sin actualizar) ──
// Corre diario (ver configurarAlertasAtrasoRA). Si un inscrito no dado de baja
// no ha tenido su Semana_Actual actualizada (Fecha_Ultima_Semana, col H) en más
// de 21 días —usando Fecha_Inscripcion (col E) como respaldo si nunca se ha
// actualizado— se marca Requiere_Seguimiento="Sí". Se limpia automáticamente en
// cuanto alguien actualiza su semana desde el panel o se marca "atendido".
function revisarAtrasosRA() {
  var ws = SS.getSheetByName("RA_Inscritos");
  if (!ws) return;
  var data = ws.getDataRange().getValues();
  var ahora = new Date();
  var LIMITE_MS = 21 * 24 * 60 * 60 * 1000;

  for (var i = 1; i < data.length; i++) {
    var estado = String(data[i][6] || "").trim().toLowerCase(); // Estado (col G)
    if (estado === "baja") continue;

    var baseline = data[i][7] || data[i][4]; // Fecha_Ultima_Semana (H) o Fecha_Inscripcion (E)
    if (!baseline) continue;
    var fechaBase = new Date(baseline);
    if (isNaN(fechaBase.getTime())) continue;

    var yaMarcado = String(data[i][5] || "").trim().toLowerCase() === "sí"; // Requiere_Seguimiento (F)
    if (!yaMarcado && (ahora - fechaBase) > LIMITE_MS) {
      ws.getRange(i + 1, 6).setValue("Sí"); // Requiere_Seguimiento (col F)
    }
  }
}

function _asegurarTriggerDiario(fnName) {
  var yaExiste = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === fnName;
  });
  if (!yaExiste) ScriptApp.newTrigger(fnName).timeBased().everyDays(1).atHour(8).create();
}

function configurarAlertasAtrasoRA() {
  _asegurarTriggerDiario("revisarAtrasosRA");
  SpreadsheetApp.getUi().alert("Listo. Cada día revisará si algún inscrito de Reto Ahorro (no dado de baja) lleva más de 3 semanas sin actualizar su semana, y marcará Requiere_Seguimiento en automático.");
}

// ── Retiro del Espíritu Santo: Alta de nueva edición ────────────
// Cols RE_Ediciones: A=id_edicion, B=nombre_edicion, C=fechas_texto,
//                    D=fecha_inicio, E=fecha_fin, F=lugar, G=estado, H=notas
function handleReAltaEdicion(p) {
  const ws = SS.getSheetByName("RE_Ediciones");
  if (!ws) return resp({ ok: false, error: "Pestaña RE_Ediciones no encontrada" });

  const anio = p.fechaInicio ? new Date(p.fechaInicio).getFullYear() : new Date().getFullYear();
  const datos = ws.getDataRange().getValues();
  let maxNum = 0;
  const patron = new RegExp("^RE-" + anio + "-(\\d+)$");
  for (let i = 1; i < datos.length; i++) {
    const m = String(datos[i][0] || "").match(patron);
    if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
  }
  const idEdicion = "RE-" + anio + "-" + String(maxNum + 1).padStart(2, "0");

  ws.appendRow([
    idEdicion,
    p.nombreEdicion || "",
    p.fechasTexto || "",
    p.fechaInicio ? new Date(p.fechaInicio) : "",
    p.fechaFin ? new Date(p.fechaFin) : "",
    p.lugar || "",
    p.estado || "Próxima",
    p.notas || ""
  ]);

  return resp({ ok: true, idAsignado: idEdicion });
}

// ── Retiro del Espíritu Santo: Alta manual de un participante ──
// Cols RE_Asistencias: A=id_edicion, B=nombre, C=edicion_label, D=asistio, E=dio_testimonio, F=notas
function handleReRegistro(p) {
  const ws = SS.getSheetByName("RE_Asistencias");
  if (!ws) return resp({ ok: false, error: "Pestaña RE_Asistencias no encontrada" });

  const nombre = (p.nombre || "").trim();
  const idEdicion = (p.idEdicion || "").trim();
  if (!nombre || !idEdicion) return resp({ ok: false, error: "Nombre y edición son obligatorios." });

  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === idEdicion && String(data[i][1]).trim().toLowerCase() === nombre.toLowerCase()) {
      return resp({ ok: false, error: "Este participante ya está registrado en esta edición." });
    }
  }

  ws.appendRow([idEdicion, nombre, "", "Pendiente", "—", p.notas || ""]);
  return resp({ ok: true });
}

// ── Retiro del Espíritu Santo: Pase de lista en lote ────────────
// Actualiza RE_Asistencias y, si asistio="Sí", también Concentrado (col K=Retiro).
// Concentrado real (verificado 2026-07-03, 19 columnas): ...I=Turno, J=Psic, K=Retiro,
// L=Impulso GEB, M=Diagnóstico educativo, N=Escuela GEB, O=Donadores, P=Ahorro,
// Q=Biblioteca Donante, R=Biblioteca Receptor, S=Camino Santiago. Ya no existe RetiroFecha.
function handleReAsistenciaLote(p) {
  const ws = SS.getSheetByName("RE_Asistencias");
  if (!ws) return resp({ ok: false, error: "Pestaña RE_Asistencias no encontrada" });

  const idEdicion = (p.idEdicion || "").trim();
  const registros = p.registros || [];
  const data = ws.getDataRange().getValues();

  registros.forEach(function(reg) {
    const nombre = (reg.nombre || "").trim();
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === idEdicion && String(data[i][1]).trim().toLowerCase() === nombre.toLowerCase()) {
        ws.getRange(i + 1, 4).setValue(reg.asistio || "");
        ws.getRange(i + 1, 5).setValue(reg.dioTestimonio || "");
        found = true;
        break;
      }
    }
    if (!found) {
      ws.appendRow([idEdicion, nombre, "", reg.asistio || "", reg.dioTestimonio || "", ""]);
    }
  });

  // Sincroniza Concentrado para quienes asistieron (col 11=K=Retiro)
  const wsConc = SS.getSheetByName("Concentrado");
  if (wsConc) {
    const concData = wsConc.getDataRange().getValues();
    registros.forEach(function(reg) {
      if (reg.asistio !== "Sí") return;
      const nombre = (reg.nombre || "").trim().toLowerCase();
      for (let j = 1; j < concData.length; j++) {
        if (String(concData[j][0]).trim().toLowerCase() === nombre) {
          wsConc.getRange(j + 1, 11).setValue("Sí");
          break;
        }
      }
    });
  }

  return resp({ ok: true });
}

// ── Retiro del Espíritu Santo: Guardar acomodo de habitaciones ──
// Cols RE_Habitaciones: A=id_edicion, B=habitacion, C=genero, D-G=nombre_1..4, H=notas
// Borra y reescribe todas las filas de la edición (el panel manda el estado completo).
function handleReGuardarHabitaciones(p) {
  const ws = SS.getSheetByName("RE_Habitaciones");
  if (!ws) return resp({ ok: false, error: "Pestaña RE_Habitaciones no encontrada" });

  const idEdicion = (p.idEdicion || "").trim();
  const habitaciones = p.habitaciones || [];

  const data = ws.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim() === idEdicion) ws.deleteRow(i + 1);
  }

  habitaciones.forEach(function(h) {
    const oc = h.ocupantes || [];
    ws.appendRow([idEdicion, h.habitacion || "", h.genero || "", oc[0]||"", oc[1]||"", oc[2]||"", oc[3]||"", h.notas || ""]);
  });

  return resp({ ok: true });
}

// ================================================================
// BIBLIOTECA COMUNITARIA
// ================================================================

const COLS_BIB = {
  prestamos: {
    timestamp: ["timestamp", "Marca temporal"],
    nombre: ["nombre", "Nombre completo"],
    correo: ["correo", "Correo de contacto", "Correo corporativo"],
    whatsapp: ["whatsapp", "WhatsApp de contacto", "WhatsApp"],
    idLibro: ["id_libro", "ID del libro"],
    titulo: ["titulo", "Título del libro"],
    fechaPrestamo: ["fecha_prestamo"],
    fechaCompromiso: ["fecha_compromiso", "Fecha compromiso de devolución"],
    disclosure: ["disclosure_aceptado", "Acepto los siguientes términos"],
    devuelto: ["devuelto"],
    fechaDevolucionReal: ["fecha_devolucion_real"],
    condicionDevolucion: ["condicion_devolucion"],
    comentarios: ["comentarios"],
    depto: ["depto", "Departamento"]
  },
  donaciones: {
    timestamp: ["timestamp", "Marca temporal"],
    donante: ["donante"],
    correo: ["correo"],
    whatsapp: ["whatsapp"],
    titulo: ["titulo"],
    autor: ["autor"],
    tipo: ["tipo"],
    fechaIngreso: ["fecha_ingreso"],
    fechaRetorno: ["fecha_retorno"],
    idLibro: ["id_libro"],
    reciboEnviado: ["recibo_enviado"],
    etiquetaGenerada: ["etiqueta_generada"],
    linkRecibo: ["link_recibo"],
    linkEtiqueta: ["link_etiqueta"]
  },
  devoluciones: {
    timestamp: ["timestamp", "Marca temporal"],
    nombre: ["nombre", "Nombre completo"],
    idLibro: ["id_libro", "ID del libro"],
    titulo: ["titulo", "Título del libro"],
    condicion: ["condicion", "¿En qué condiciones lo devuelves?"],
    descripcionDano: ["descripcion_dano", "Si hay daño"],
    recomendaria: ["recomendaria", "¿Recomendarías"],
    comentarios: ["comentarios", "Comentarios o reseña"],
    procesado: ["procesado", "Procesado"]
  }
};

function formatearFechaBib_(fecha) {
  if (!fecha) return "";
  var fechaObjeto = new Date(fecha);
  if (isNaN(fechaObjeto.getTime())) return String(fecha);
  return Utilities.formatDate(fechaObjeto, Session.getScriptTimeZone(), "dd/MM/yyyy");
}

function formatearFechaCortaBib_(fecha) {
  if (!fecha) return "";
  var fechaObjeto = new Date(fecha);
  if (isNaN(fechaObjeto.getTime())) return String(fecha);
  return Utilities.formatDate(fechaObjeto, Session.getScriptTimeZone(), "dd/MM/yy");
}

function buscarIndiceBib_(headers, nombresAlternativos) {
  const normalizar = (s) => String(s || "")
    .toLowerCase()
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .trim()
    .replace(/\s+/g, " ");
  const headersNorm = headers.map(normalizar);
  for (const nombre of nombresAlternativos) {
    const nombreNorm = normalizar(nombre);
    let idx = headersNorm.indexOf(nombreNorm);
    if (idx >= 0) return idx;
    idx = headersNorm.findIndex(h => h.startsWith(nombreNorm) || (nombreNorm.startsWith(h) && h.length > 3));
    if (idx >= 0) return idx;
  }
  return -1;
}

function obtenerIndicesBib_(headers, mapeoColumnas) {
  const idx = {};
  for (const [key, alternativas] of Object.entries(mapeoColumnas)) {
    idx[key] = buscarIndiceBib_(headers, alternativas);
  }
  return idx;
}

function obtenerImagenBase64Bib_(driveFileId) {
  try {
    const archivo = DriveApp.getFileById(driveFileId);
    const blob = archivo.getBlob();
    const base64 = Utilities.base64Encode(blob.getBytes());
    const mimeType = blob.getContentType();
    return `data:${mimeType};base64,${base64}`;
  } catch (e) {
    return null;
  }
}

function obtenerOCrearCarpetaBib_(nombreCarpeta) {
  const carpetas = DriveApp.getFoldersByName(nombreCarpeta);
  if (carpetas.hasNext()) return carpetas.next();
  return DriveApp.createFolder(nombreCarpeta);
}

// Núcleo de matching reutilizado por el barrido completo (procesarDevolucionesBib_)
// y por el handler individual del Web App.
function matchearYProcesarDevolucionBib_(hojaDev, hojaPre, idxD, idxP, datosPre, filaDev, filaDevNum) {
  const nombreBuscado = String(filaDev[idxD.nombre] || "").trim().toLowerCase();
  const idBuscado = String(filaDev[idxD.idLibro] || "").trim().toLowerCase();
  const fechaDevolucion = filaDev[idxD.timestamp];
  const condicion = idxD.condicion >= 0 ? filaDev[idxD.condicion] : "";

  if (!nombreBuscado || !idBuscado) return { match: false, intentado: false };

  let mejorMatch = -1;
  let mejorFecha = null;

  for (let j = 1; j < datosPre.length; j++) {
    const filaPre = datosPre[j];
    const nombrePre = String(filaPre[idxP.nombre] || "").trim().toLowerCase();
    const idPre = String(filaPre[idxP.idLibro] || "").trim().toLowerCase();
    const tituloPre = String(filaPre[idxP.titulo] || "").trim().toLowerCase();
    const devuelto = filaPre[idxP.devuelto];

    if (devuelto === "Sí" || devuelto === "SI" || devuelto === true) continue;
    if (nombrePre !== nombreBuscado) continue;

    const matchLibro = (idPre === idBuscado) ||
                       (tituloPre && tituloPre.includes(idBuscado)) ||
                       (idBuscado && idBuscado.includes(tituloPre) && tituloPre.length > 3);
    if (!matchLibro) continue;

    const fechaPrestamo = filaPre[idxP.timestamp] ? new Date(filaPre[idxP.timestamp]) : null;
    if (!mejorFecha || (fechaPrestamo && fechaPrestamo > mejorFecha)) {
      mejorMatch = j + 1;
      mejorFecha = fechaPrestamo;
    }
  }

  if (mejorMatch > 0) {
    if (idxP.devuelto >= 0) hojaPre.getRange(mejorMatch, idxP.devuelto + 1).setValue("Sí");
    if (idxP.fechaDevolucionReal >= 0 && fechaDevolucion)
      hojaPre.getRange(mejorMatch, idxP.fechaDevolucionReal + 1).setValue(new Date(fechaDevolucion));
    if (idxP.condicionDevolucion >= 0 && condicion)
      hojaPre.getRange(mejorMatch, idxP.condicionDevolucion + 1).setValue(condicion);
    if (idxD.procesado >= 0)
      hojaDev.getRange(filaDevNum, idxD.procesado + 1).setValue("✓ Cerrado " + formatearFechaBib_(new Date()));
    return { match: true, intentado: true, filaPrestamo: mejorMatch };
  } else {
    if (idxD.procesado >= 0)
      hojaDev.getRange(filaDevNum, idxD.procesado + 1).setValue("⚠️ Revisar manualmente");
    return { match: false, intentado: true };
  }
}

// Barrido completo de devoluciones no procesadas (respaldo manual desde el editor de Apps Script).
function procesarDevolucionesBib_() {
  const hojaDev = SS.getSheetByName("BIB_Devoluciones");
  const hojaPre = SS.getSheetByName("BIB_Prestamos");
  if (!hojaDev || !hojaPre) return { procesados: 0, sinMatch: 0 };

  const datosDev = hojaDev.getDataRange().getValues();
  if (datosDev.length < 2) return { procesados: 0, sinMatch: 0 };
  const idxD = obtenerIndicesBib_(datosDev[0], COLS_BIB.devoluciones);

  const datosPre = hojaPre.getDataRange().getValues();
  if (datosPre.length < 2) return { procesados: 0, sinMatch: 0 };
  const idxP = obtenerIndicesBib_(datosPre[0], COLS_BIB.prestamos);

  let procesados = 0, sinMatch = 0;
  for (let i = 1; i < datosDev.length; i++) {
    const filaDev = datosDev[i];
    if (idxD.procesado >= 0 && filaDev[idxD.procesado]) continue;
    const resultado = matchearYProcesarDevolucionBib_(hojaDev, hojaPre, idxD, idxP, datosPre, filaDev, i + 1);
    if (resultado.match) procesados++;
    else if (resultado.intentado) sinMatch++;
  }
  return { procesados, sinMatch };
}

function generarReciboPorFilaBib_(fila) {
  const hoja = SS.getSheetByName("BIB_Donaciones");
  if (!hoja) throw new Error("Pestaña BIB_Donaciones no encontrada");

  const headers = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  const datos = hoja.getRange(fila, 1, 1, hoja.getLastColumn()).getValues()[0];
  const idx = obtenerIndicesBib_(headers, COLS_BIB.donaciones);

  const donante = datos[idx.donante];
  const titulo = datos[idx.titulo];
  const autor = idx.autor >= 0 ? datos[idx.autor] : "";
  const tipo = datos[idx.tipo];
  const fechaIngreso = formatearFechaBib_(datos[idx.fechaIngreso]);
  const fechaRetorno = idx.fechaRetorno >= 0 && datos[idx.fechaRetorno]
    ? formatearFechaBib_(datos[idx.fechaRetorno])
    : "N/A (donación permanente)";
  const folio = "BIB-" + fila + "-" + new Date().getFullYear();

  const logoBiblioData = obtenerImagenBase64Bib_(LOGO_BIBLIO_ID);
  const logoUgebData = obtenerImagenBase64Bib_(LOGO_UGEB_ID);

  const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><style>
    body { font-family: Georgia, serif; padding: 40px; color: #2C2C2A; }
    .logos { text-align: center; margin-bottom: 20px; }
    .logos img { height: 80px; margin: 0 20px; vertical-align: middle; }
    .header { text-align: center; border-bottom: 2px solid ${COLOR_PRIMARIO}; padding-bottom: 20px; }
    h1 { color: ${COLOR_PRIMARIO}; margin: 10px 0 5px; }
    .folio { color: #718096; font-size: 13px; }
    .contenido { margin-top: 30px; line-height: 1.8; font-size: 15px; }
    .dato { margin: 12px 0; }
    .label { font-weight: bold; color: ${COLOR_PRIMARIO}; }
    .gracias { margin-top: 40px; padding: 20px; background: #EBF2FB; border-left: 4px solid ${COLOR_PRIMARIO}; }
    .firma { margin-top: 60px; text-align: center; color: #888780; font-size: 12px; }
  </style></head><body>
    <div class="header">
      <div class="logos">
        ${logoBiblioData ? `<img src="${logoBiblioData}" alt="Biblio">` : ''}
        ${logoUgebData ? `<img src="${logoUgebData}" alt="GEB University">` : ''}
      </div>
      <h1>Recibo de ${tipo === "Permanente" ? "donación" : "préstamo a biblioteca"}</h1>
      <p class="folio">Folio ${folio}</p>
    </div>
    <div class="contenido">
      <div class="dato"><span class="label">Donante:</span> ${donante}</div>
      <div class="dato"><span class="label">Libro:</span> "${titulo}"${autor ? " — " + autor : ""}</div>
      <div class="dato"><span class="label">Tipo:</span> ${tipo}</div>
      <div class="dato"><span class="label">Fecha de ingreso:</span> ${fechaIngreso}</div>
      <div class="dato"><span class="label">Fecha de retorno:</span> ${fechaRetorno}</div>
    </div>
    <div class="gracias">
      <strong>¡Gracias por contribuir a nuestra biblioteca comunitaria!</strong><br>
      ${tipo === "Permanente"
        ? "Tu libro pasa a formar parte permanente del acervo y estará disponible para todos los colaboradores."
        : "Tu libro estará disponible para préstamo durante 6 meses. En la fecha de retorno te contactaremos para devolvértelo."}
    </div>
    <div class="firma">
      Documento generado automáticamente — Biblioteca Comunitaria GEB University<br>
      ${formatearFechaBib_(new Date())}
    </div>
  </body></html>`;

  const blob = Utilities.newBlob(html, "text/html", `Recibo_${folio}.html`).getAs("application/pdf");
  blob.setName(`Recibo_${folio}.pdf`);

  const carpeta = obtenerOCrearCarpetaBib_("Recibos Biblioteca");
  const archivo = carpeta.createFile(blob);
  archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const url = archivo.getUrl();

  if (idx.reciboEnviado >= 0)
    hoja.getRange(fila, idx.reciboEnviado + 1).setValue("Generado " + formatearFechaBib_(new Date()));
  if (idx.linkRecibo >= 0)
    hoja.getRange(fila, idx.linkRecibo + 1).setValue(url).setFontColor("#1D9E75");

  return { folio, url };
}

function construirHtmlEtiquetasBib_(etiquetas) {
  const logoBiblioData = obtenerImagenBase64Bib_(LOGO_BIBLIO_ID);
  const logoUgebData = obtenerImagenBase64Bib_(LOGO_UGEB_ID);

  let html = `<!DOCTYPE html><html><head><meta charset="utf-8"><style>
    @page { size: letter; margin: 0.4in; }
    body { margin: 0; padding: 0; font-family: 'Helvetica', Arial, sans-serif; }
    .hoja { display: grid; grid-template-columns: 1fr 1fr; grid-template-rows: 1fr 1fr; gap: 8px; page-break-after: always; height: 10.2in; }
    .hoja:last-child { page-break-after: auto; }
    .etiqueta { border: 1px solid #D5D5D0; border-radius: 6px; padding: 18px 16px; display: flex; flex-direction: column; box-sizing: border-box; page-break-inside: avoid; }
    .logos { text-align: center; margin-bottom: 12px; }
    .logos img { height: 60px; margin: 0 12px; vertical-align: middle; object-fit: contain; }
    .divisor { border-top: 1px solid #C8C8C0; margin: 0 30px 14px; }
    .header-text { text-align: center; font-size: 9px; letter-spacing: 2px; color: #888780; margin-bottom: 14px; }
    .titulo { text-align: center; font-style: italic; font-size: 13px; color: #2C2C2A; margin-bottom: 16px; min-height: 18px; }
    .label { text-align: center; font-size: 12px; color: #5F5E5A; margin-bottom: 4px; }
    .donante { text-align: center; font-size: 18px; font-weight: bold; color: ${COLOR_PRIMARIO}; margin-bottom: 20px; }
    .pie { text-align: center; font-size: 11px; color: #888780; margin-top: auto; }
  </style></head><body>`;

  for (let i = 0; i < etiquetas.length; i += 4) {
    html += `<div class="hoja">`;
    for (let j = 0; j < 4; j++) {
      const e = etiquetas[i + j];
      if (e) {
        const tituloTexto = e.titulo ? `"${e.titulo}"` : "";
        const idFecha = [e.fecha, e.idLibro].filter(Boolean).join(" · ");
        html += `<div class="etiqueta">
          <div class="logos">
            ${logoBiblioData ? `<img src="${logoBiblioData}" alt="">` : ''}
            ${logoUgebData ? `<img src="${logoUgebData}" alt="">` : ''}
          </div>
          <div class="divisor"></div>
          <div class="header-text">FORMA PARTE DE NUESTRA BIBLIOTECA</div>
          <div class="titulo">${tituloTexto}</div>
          <div class="label">Donado por</div>
          <div class="donante">${e.donante}</div>
          <div class="pie">${idFecha}</div>
        </div>`;
      } else {
        html += `<div class="etiqueta" style="border: 1px dashed #E5E5E0;"></div>`;
      }
    }
    html += `</div>`;
  }
  html += `</body></html>`;
  return html;
}

function generarEtiquetaPorFilasBib_(filas) {
  const hoja = SS.getSheetByName("BIB_Donaciones");
  if (!hoja) throw new Error("Pestaña BIB_Donaciones no encontrada");

  const headers = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  const idx = obtenerIndicesBib_(headers, COLS_BIB.donaciones);

  const etiquetas = [];
  const filasValidas = [];
  filas.forEach(fila => {
    const datos = hoja.getRange(fila, 1, 1, hoja.getLastColumn()).getValues()[0];
    if (!datos[idx.donante]) return;
    etiquetas.push({
      donante: datos[idx.donante],
      titulo: datos[idx.titulo] || "",
      fecha: formatearFechaCortaBib_(datos[idx.fechaIngreso]),
      idLibro: datos[idx.idLibro] || ""
    });
    filasValidas.push(fila);
  });

  if (etiquetas.length === 0) throw new Error("No se encontraron datos válidos en las filas indicadas");

  const html = construirHtmlEtiquetasBib_(etiquetas);
  const blob = Utilities.newBlob(html, "text/html", "etiquetas.html").getAs("application/pdf");
  const fecha = Utilities.formatDate(new Date(), "GMT-6", "yyyy-MM-dd_HHmm");
  blob.setName(`Etiquetas_${fecha}.pdf`);

  const carpeta = obtenerOCrearCarpetaBib_("Etiquetas Biblioteca");
  const archivo = carpeta.createFile(blob);
  archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const url = archivo.getUrl();

  filasValidas.forEach(f => {
    if (idx.etiquetaGenerada >= 0)
      hoja.getRange(f, idx.etiquetaGenerada + 1).setValue("Generada " + formatearFechaBib_(new Date()));
    if (idx.linkEtiqueta >= 0)
      hoja.getRange(f, idx.linkEtiqueta + 1).setValue(url).setFontColor("#1D9E75");
  });

  return { url, cantidad: etiquetas.length };
}

function sincronizarFisicosBib_() {
  const ssBiblio = SpreadsheetApp.openById(SHEET_BIBLIOTECA_ID);
  const hojaFisicos = ssBiblio.getSheetByName("Fisicos");
  if (!hojaFisicos) throw new Error("No se encontró la pestaña 'Fisicos' en la Biblioteca Virtual.");

  const fisicosActuales = hojaFisicos.getDataRange().getValues();
  const idsExistentes = new Set(
    fisicosActuales.slice(1).map(r => String(r[3] || "").trim()).filter(id => id !== "")
  );

  let nextId = fisicosActuales.length;
  const nuevasFilas = [];

  const hojaDonaciones = SS.getSheetByName("BIB_Donaciones");
  if (hojaDonaciones) {
    const donaciones = hojaDonaciones.getDataRange().getValues();
    if (donaciones.length > 1) {
      const idxD = obtenerIndicesBib_(donaciones[0], COLS_BIB.donaciones);
      for (let i = 1; i < donaciones.length; i++) {
        const fila = donaciones[i];
        const titulo  = idxD.titulo >= 0 ? String(fila[idxD.titulo] || "").trim() : "";
        const autor   = idxD.autor >= 0 ? String(fila[idxD.autor] || "").trim() : "";
        const idLibro = idxD.idLibro >= 0 ? String(fila[idxD.idLibro] || "").trim() : "";
        if (!titulo || !idLibro) continue;
        if (idsExistentes.has(idLibro)) continue;
        nextId++;
        nuevasFilas.push([nextId, titulo, autor, idLibro, "", "", true, "", "", ""]);
        idsExistentes.add(idLibro);
      }
    }
  }

  if (nuevasFilas.length > 0) {
    hojaFisicos.getRange(hojaFisicos.getLastRow() + 1, 1, nuevasFilas.length, nuevasFilas[0].length).setValues(nuevasFilas);
  }

  const hojaPrestamos = SS.getSheetByName("BIB_Prestamos");
  if (hojaPrestamos) {
    const prestamos = hojaPrestamos.getDataRange().getValues();
    if (prestamos.length > 1) {
      const idxP = obtenerIndicesBib_(prestamos[0], COLS_BIB.prestamos);
      for (let i = 1; i < prestamos.length; i++) {
        const fila = prestamos[i];
        const nombre  = idxP.nombre >= 0 ? String(fila[idxP.nombre] || "").trim() : "";
        const idLibro = idxP.idLibro >= 0 ? String(fila[idxP.idLibro] || "").trim() : "";
        let fecha = "";
        if (idxP.fechaPrestamo >= 0 && fila[idxP.fechaPrestamo]) fecha = fila[idxP.fechaPrestamo];
        else if (idxP.timestamp >= 0 && fila[idxP.timestamp]) fecha = fila[idxP.timestamp];
        const devuelto = idxP.devuelto >= 0 ? fila[idxP.devuelto] : "";
        if (devuelto === "Sí" || devuelto === "SI" || devuelto === true) continue;
        if (!idLibro) continue;

        const filasFisicos = hojaFisicos.getDataRange().getValues();
        for (let j = 1; j < filasFisicos.length; j++) {
          const idFisico = String(filasFisicos[j][3] || "").trim();
          if (idFisico === idLibro) {
            hojaFisicos.getRange(j + 1, 7).setValue(false);
            hojaFisicos.getRange(j + 1, 8).setValue(nombre);
            hojaFisicos.getRange(j + 1, 10).setValue(fecha ? Utilities.formatDate(new Date(fecha), "GMT-6", "yyyy-MM-dd") : "");
            break;
          }
        }
      }
    }
  }

  return {
    nuevos: nuevasFilas.length,
    mensaje: nuevasFilas.length > 0
      ? `${nuevasFilas.length} libro(s) nuevo(s) agregado(s) y préstamos activos sincronizados por ID.`
      : "Préstamos activos sincronizados por ID. No había libros nuevos que agregar."
  };
}

// ── Biblioteca: Alta de donación ────────────────────────────────
// Recibe: {tipo:"bib_altaDonacion", donante, correo, whatsapp, titulo, autor, modalidad}
function handleBibAltaDonacion(p) {
  const ws = SS.getSheetByName("BIB_Donaciones");
  if (!ws) return resp({ ok: false, error: "Pestaña BIB_Donaciones no encontrada" });

  const headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0];
  const idx = obtenerIndicesBib_(headers, COLS_BIB.donaciones);

  const datos = ws.getDataRange().getValues();
  let maxNum = 0;
  if (idx.idLibro >= 0) {
    for (let i = 1; i < datos.length; i++) {
      const id = String(datos[i][idx.idLibro] || "");
      const m = id.match(/DON-(\d+)/);
      if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
    }
  }
  const nuevoId = "DON-" + String(maxNum + 1).padStart(4, "0");

  const fila = new Array(headers.length).fill("");
  if (idx.timestamp >= 0)    fila[idx.timestamp]    = new Date();
  if (idx.donante >= 0)      fila[idx.donante]      = p.donante || "";
  if (idx.correo >= 0)       fila[idx.correo]       = p.correo || "";
  if (idx.whatsapp >= 0)     fila[idx.whatsapp]     = p.whatsapp || "";
  if (idx.titulo >= 0)       fila[idx.titulo]       = p.titulo || "";
  if (idx.autor >= 0)        fila[idx.autor]        = p.autor || "";
  if (idx.tipo >= 0)         fila[idx.tipo]         = p.modalidad || "";
  if (idx.fechaIngreso >= 0) fila[idx.fechaIngreso] = new Date();
  if (idx.idLibro >= 0)      fila[idx.idLibro]      = nuevoId;

  ws.appendRow(fila);

  return resp({ ok: true, idAsignado: nuevoId, fila: ws.getLastRow() });
}

// ── Biblioteca: Procesar una devolución específica ──────────────
// Recibe: {tipo:"bib_procesarDevolucion", filaDevolucion}
function handleBibProcesarDevolucion(p) {
  const hojaDev = SS.getSheetByName("BIB_Devoluciones");
  const hojaPre = SS.getSheetByName("BIB_Prestamos");
  if (!hojaDev || !hojaPre) return resp({ ok: false, error: "Pestañas BIB_Devoluciones/BIB_Prestamos no encontradas" });

  const filaDevNum = parseInt(p.filaDevolucion, 10);
  if (!filaDevNum || filaDevNum < 2) return resp({ ok: false, error: "filaDevolucion inválida" });

  const headersDev = hojaDev.getRange(1, 1, 1, hojaDev.getLastColumn()).getValues()[0];
  const idxD = obtenerIndicesBib_(headersDev, COLS_BIB.devoluciones);
  const datosPre = hojaPre.getDataRange().getValues();
  const idxP = obtenerIndicesBib_(datosPre[0], COLS_BIB.prestamos);

  const filaDev = hojaDev.getRange(filaDevNum, 1, 1, hojaDev.getLastColumn()).getValues()[0];

  const resultado = matchearYProcesarDevolucionBib_(hojaDev, hojaPre, idxD, idxP, datosPre, filaDev, filaDevNum);

  if (resultado.match) {
    return resp({ ok: true, match: true, filaPrestamo: resultado.filaPrestamo });
  } else {
    return resp({ ok: true, match: false, mensaje: "No se encontró préstamo abierto que coincida — revisar manualmente" });
  }
}

// ── Biblioteca: Extender préstamo +14 días ──────────────────────
// Recibe: {tipo:"bib_extenderPrestamo", nombre, idLibro}
function handleBibExtenderPrestamo(p) {
  const ws = SS.getSheetByName("BIB_Prestamos");
  if (!ws) return resp({ ok: false, error: "Pestaña BIB_Prestamos no encontrada" });

  const datos = ws.getDataRange().getValues();
  const idx = obtenerIndicesBib_(datos[0], COLS_BIB.prestamos);

  const nombreBuscado = String(p.nombre || "").trim().toLowerCase();
  const idBuscado = String(p.idLibro || "").trim().toLowerCase();

  for (let i = 1; i < datos.length; i++) {
    const nombreFila = String(datos[i][idx.nombre] || "").trim().toLowerCase();
    const idFila = String(datos[i][idx.idLibro] || "").trim().toLowerCase();
    if (nombreFila === nombreBuscado && idFila === idBuscado) {
      const fechaActual = datos[i][idx.fechaCompromiso] ? new Date(datos[i][idx.fechaCompromiso]) : new Date();
      const nuevaFecha = new Date(fechaActual.getTime() + 14 * 24 * 60 * 60 * 1000);
      ws.getRange(i + 1, idx.fechaCompromiso + 1).setValue(nuevaFecha);
      return resp({ ok: true, nuevaFechaCompromiso: formatearFechaBib_(nuevaFecha) });
    }
  }

  return resp({ ok: false, error: "No se encontró el préstamo (nombre + idLibro)" });
}

// ── Biblioteca: Generar recibo de donación (PDF) ────────────────
// Recibe: {tipo:"bib_generarRecibo", fila}
function handleBibGenerarRecibo(p) {
  try {
    const resultado = generarReciboPorFilaBib_(parseInt(p.fila, 10));
    return resp({ ok: true, folio: resultado.folio, url: resultado.url });
  } catch (err) {
    return resp({ ok: false, error: err.message });
  }
}

// ── Biblioteca: Generar etiqueta(s) de libro donado (PDF) ───────
// Recibe: {tipo:"bib_generarEtiqueta", filas: [n, n, ...]}
function handleBibGenerarEtiqueta(p) {
  try {
    const filas = (p.filas || []).map(f => parseInt(f, 10)).filter(f => f >= 2);
    if (filas.length === 0) return resp({ ok: false, error: "No se recibieron filas válidas" });
    const resultado = generarEtiquetaPorFilasBib_(filas);
    return resp({ ok: true, url: resultado.url, cantidad: resultado.cantidad });
  } catch (err) {
    return resp({ ok: false, error: err.message });
  }
}

// ── Biblioteca: Sincronizar libros físicos con Biblioteca Virtual ──
// Recibe: {tipo:"bib_sincronizarFisicos"}
function handleBibSincronizarFisicos(p) {
  try {
    const resultado = sincronizarFisicosBib_();
    return resp({ ok: true, nuevos: resultado.nuevos, mensaje: resultado.mensaje });
  } catch (err) {
    return resp({ ok: false, error: err.message });
  }
}

// ── Respuesta CORS ───────────────────────────────────────────────
function resp(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// GET para verificar que el Web App está activo
function doGet() {
  return resp({ ok: true, msg: "GEB Web App activo" });
}
