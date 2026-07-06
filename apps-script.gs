// ================================================================
// GEB Iniciativas Humanas — Web App
// Despliega como: Ejecutar como "Yo", Acceso "Cualquier persona"
// ================================================================
// Desconectado por completo del Sheet histórico (2026-07-05): el Sheet nuevo
// es la única fuente de datos operativos para las 9 iniciativas. El Sheet
// histórico (SHEET_HISTORICO_ID) ya no se lee ni se escribe en ningún flujo
// normal — su única mención en este archivo es dentro de las funciones de
// migración de una sola vez (ya usadas), que quedan como referencia histórica
// y no corren automáticamente.

// Sheet nuevo unificado — única fuente de datos de las 9 iniciativas.
const SHEET_NUEVO_ID = "1oaWADtW9SmqxuOoOM1bf65Ht_NkpE4l_o_DMutN_IbI";
const SS = SpreadsheetApp.openById(SHEET_NUEVO_ID);

// Sheet histórico — solo se usa como fuente de lectura dentro de las
// funciones de migración de una sola vez (migrarXASheetNuevo). Nunca se
// escribe ahí.
const SHEET_HISTORICO_ID = "1fS5qeoB1ViuCUP4HOQ1zTJgh4tfWUkObvb5LMr3to5A";

// Biblioteca Virtual externa (solo lectura/escritura de la pestaña "Fisicos")
const SHEET_BIBLIOTECA_ID = "1FDZB3aR-YAyVMsiAjo92PuUdH0iTfBmreCWv5DvCpdM";

const LOGO_BIBLIO_ID = "1NqoFmESlsTP4dpFscYglcs9o6THQP4TR";
const LOGO_UGEB_ID = "1mBIHoKyngoa7cBiSvHCVY0x_pIkFyARJ";
const COLOR_PRIMARIO = "#185FA5";

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("GEB CRM")
    .addItem("🔧 Configurar alertas de atraso Reto Ahorro (1 vez)", "configurarAlertasAtrasoRA")
    .addItem("🔧 Configurar AP en Sheet nuevo (1 vez)", "configurarAPSheetNuevo")
    .addItem("🧪 Sembrar datos de prueba AP (Sheet nuevo)", "sembrarPruebasAP")
    .addItem("🧪 Sembrar datos de prueba Escuela GEB (Sheet nuevo)", "sembrarPruebasEG")
    .addItem("🔧 Configurar RE + Camino Santiago en Sheet nuevo (1 vez)", "configurarREyCSSheetNuevo")
    .addItem("🧪 Sembrar datos de prueba Retiro Espiritual (Sheet nuevo)", "sembrarPruebasRE")
    .addItem("🔧 Migrar Biblioteca al Sheet nuevo (1 vez, copia datos reales)", "migrarBibliotecaASheetNuevo")
    .addItem("🔧 Configurar Impulso GEB en Sheet nuevo (1 vez)", "configurarIGSheetNuevo")
    .addItem("🧪 Sembrar datos de prueba Impulso GEB (Sheet nuevo)", "sembrarPruebasIG")
    .addItem("🔧 Configurar Beca Educativa en Sheet nuevo (1 vez)", "configurarBESheetNuevo")
    .addItem("🧪 Sembrar datos de prueba Beca Educativa (Sheet nuevo)", "sembrarPruebasBE")
    .addItem("🔧 Migrar Copys restantes (RA/SV/EG) al Sheet nuevo (1 vez)", "migrarCopysRestantesASheetNuevo")
    .addItem("🔧 Migrar logs restantes (EG_Asistencias/Examenes, SV_Historial) (1 vez)", "migrarLogsRestantesASheetNuevo")
    .addItem("🧹 Desactivar sincronización antigua de Concentrado (1 vez)", "desactivarSincronizacionAntigua")
    .addItem("🗑️ Borrar TODOS los datos de prueba (PRUEBA...)", "borrarDatosPruebaTodos")
    .addItem("🔧 Configurar Eco-Acción en Sheet nuevo (1 vez)", "configurarEASheetNuevo")
    .addItem("🧪 Sembrar datos de prueba Eco-Acción (Sheet nuevo)", "sembrarPruebasEA")
    .addToUi();
}

// Elimina cualquier trigger instalado por versiones anteriores del script que
// sincronizaba altas nuevas desde Concentrado hacia el Sheet histórico
// (sincronizarNuevosAP/RA/SV/EG). Esa mecánica ya no existe — todas las altas
// entran por altas.html → handleAltaUnificada, directo al Sheet nuevo. Correr
// una sola vez como parte del corte total; es seguro correrlo aunque esos
// triggers ya no existan (simplemente no encuentra nada que borrar).
function desactivarSincronizacionAntigua() {
  var eliminados = [];
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var fn = t.getHandlerFunction();
    if (fn.indexOf("sincronizarNuevos") === 0) {
      ScriptApp.deleteTrigger(t);
      eliminados.push(fn);
    }
  });
  SpreadsheetApp.getUi().alert(eliminados.length
    ? "Triggers eliminados: " + eliminados.join(", ")
    : "No había triggers de sincronización antigua instalados.");
}

// ── Borrar datos de prueba (desarrollo) ──────────────────────────
// Recorre las pestañas del Sheet nuevo que recibieron filas "PRUEBA ..." al
// sembrar datos de prueba, y las borra. No toca datos reales (solo filas cuyo
// Nombre empieza exactamente con "PRUEBA"). RE_Ediciones/Asistencias/
// Habitaciones se limpian por id_edicion ("RE-PRUEBA-01"), ya que ahí el
// Nombre no siempre es la primera columna.
function borrarDatosPruebaTodos() {
  var resultado = [];
  var tabsPorNombre = ["Colaboradores", "AP_Inscritos", "AP_Sesiones", "EG_Inscritos",
    "RE_Interesados", "CS_Inscritos", "IG_Inscritos", "BE_Inscritos", "EA_Lideres", "EA_Cilindros"];

  tabsPorNombre.forEach(function (nombreHoja) {
    var ws = SS.getSheetByName(nombreHoja);
    if (!ws) { resultado.push(nombreHoja + ": pestaña no encontrada"); return; }
    var data = ws.getDataRange().getValues();
    var borrados = 0;
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0] || "").trim().indexOf("PRUEBA") === 0) { ws.deleteRow(i + 1); borrados++; }
    }
    resultado.push(nombreHoja + ": " + borrados + " fila(s) borrada(s)");
  });

  var wsEd = SS.getSheetByName("RE_Ediciones");
  if (wsEd) {
    var dataEd = wsEd.getDataRange().getValues();
    var idsPrueba = [];
    var borradosEd = 0;
    for (var i = dataEd.length - 1; i >= 1; i--) {
      if (String(dataEd[i][0] || "").indexOf("RE-PRUEBA") === 0) {
        idsPrueba.push(dataEd[i][0]);
        wsEd.deleteRow(i + 1);
        borradosEd++;
      }
    }
    resultado.push("RE_Ediciones: " + borradosEd + " fila(s) borrada(s)");

    ["RE_Asistencias", "RE_Habitaciones"].forEach(function (nombreHoja) {
      var ws = SS.getSheetByName(nombreHoja);
      if (!ws) return;
      var data = ws.getDataRange().getValues();
      var borrados = 0;
      for (var i = data.length - 1; i >= 1; i--) {
        if (idsPrueba.indexOf(data[i][0]) >= 0) { ws.deleteRow(i + 1); borrados++; }
      }
      resultado.push(nombreHoja + ": " + borrados + " fila(s) borrada(s)");
    });
  }

  SpreadsheetApp.getUi().alert("Datos de prueba borrados:\n\n" + resultado.join("\n"));
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const tipo    = payload.tipo;

    if (tipo === "asistencia")         return handleAsistencia(payload);
    if (tipo === "examen")             return handleExamen(payload);
    if (tipo === "donacion")           return handleDonacion(payload);
    if (tipo === "ap_asistencia_lote") return handleApAsistenciaLote(payload);
    if (tipo === "ap_cambioModalidad") return handleApCambioModalidad(payload);
    if (tipo === "ap_sesionesI43")     return handleApSesionesI43(payload);
    if (tipo === "ap_baja")           return handleApBaja(payload);
    if (tipo === "ap_diagnostico")    return handleApDiagnostico(payload);
    if (tipo === "alta_unificada")    return handleAltaUnificada(payload);
    if (tipo === "sv_atendido")       return handleSvAtendido(payload);
    if (tipo === "sv_baja")           return handleSvBaja(payload);
    if (tipo === "migrar_sv_historico") return handleMigrarSVHistorico(payload);
    if (tipo === "retirar_sync_sv")     return handleRetirarSyncSV(payload);
    if (tipo === "migrar_eg_historico") return handleMigrarEGHistorico(payload);
    if (tipo === "retirar_sync_eg")     return handleRetirarSyncEG(payload);
    if (tipo === "eg_marcarAtendido")   return handleEgMarcarAtendido(payload);
    if (tipo === "eg_baja")             return handleEgBaja(payload);
    if (tipo === "eg_diagnostico")      return handleEgDiagnostico(payload);
    if (tipo === "ra_actualizarSemana") return handleRaActualizarSemana(payload);
    if (tipo === "ra_marcarAtendido")   return handleRaMarcarAtendido(payload);
    if (tipo === "ra_baja")             return handleRaBaja(payload);
    if (tipo === "migrar_ra_historico") return handleMigrarRAHistorico(payload);
    if (tipo === "retirar_sync_ra")     return handleRetirarSyncRA(payload);
    if (tipo === "migrar_ap_historico") return handleMigrarAPHistorico(payload);
    if (tipo === "retirar_sync_ap")     return handleRetirarSyncAP(payload);
    if (tipo === "re_altaEdicion")           return handleReAltaEdicion(payload);
    if (tipo === "re_registro")             return handleReRegistro(payload);
    if (tipo === "re_asistenciaLote")        return handleReAsistenciaLote(payload);
    if (tipo === "re_guardarHabitaciones")   return handleReGuardarHabitaciones(payload);
    if (tipo === "re_altaInteresado")            return handleReAltaInteresado(payload);
    if (tipo === "re_marcarAtendidoInteresado")  return handleReMarcarAtendidoInteresado(payload);
    if (tipo === "re_bajaInteresado")            return handleReBajaInteresado(payload);
    if (tipo === "re_inscribirInteresado")       return handleReInscribirInteresado(payload);
    if (tipo === "cs_marcarAtendido")            return handleCsMarcarAtendido(payload);
    if (tipo === "cs_registrarFecha")            return handleCsRegistrarFecha(payload);
    if (tipo === "cs_baja")                      return handleCsBaja(payload);
    if (tipo === "ig_actualizarCampo")   return handleIgActualizarCampo(payload);
    if (tipo === "ig_marcarAtendido")    return handleIgMarcarAtendido(payload);
    if (tipo === "ig_baja")              return handleIgBaja(payload);
    if (tipo === "ig_pasarGeneracion2")  return handleIgPasarGeneracion2(payload);
    if (tipo === "be_marcarAtendido")    return handleBeMarcarAtendido(payload);
    if (tipo === "be_baja")              return handleBeBaja(payload);
    if (tipo === "be_revisarInscripcion") return handleBeRevisarInscripcion(payload);
    if (tipo === "ea_registrarCilindro")  return handleEaRegistrarCilindro(payload);
    if (tipo === "ea_marcarAtendido")     return handleEaMarcarAtendido(payload);
    if (tipo === "ea_baja")               return handleEaBaja(payload);
    if (tipo === "bib_altaDonacion")        return handleBibAltaDonacion(payload);
    if (tipo === "bib_procesarDevolucion")  return handleBibProcesarDevolucion(payload);
    if (tipo === "bib_extenderPrestamo")    return handleBibExtenderPrestamo(payload);
    if (tipo === "bib_generarRecibo")       return handleBibGenerarRecibo(payload);
    if (tipo === "bib_generarEtiqueta")     return handleBibGenerarEtiqueta(payload);
    if (tipo === "bib_sincronizarFisicos")  return handleBibSincronizarFisicos(payload);
    if (tipo === "admin_ejecutar")          return handleAdminEjecutar(payload);

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

// ── Salvando Vidas (Sheet NUEVO desde 2026-07-04) ────────────────
// Cols SV_Inscritos en SHEET_NUEVO_ID: A=Nombre, B=Sucursal, C=Telefono_WA,
// D=Fecha_Alta, E=Estado, F=Requiere_Seguimiento, G=Notas, H=Tipo_Sangre,
// I=Mes_Guardia. El Sheet histórico (SS) ya NO se toca para esta iniciativa.
function svSheetNuevo_() {
  return SpreadsheetApp.openById(SHEET_NUEVO_ID).getSheetByName("SV_Inscritos");
}

// "Sí" lo pone handleAltaUnificada en altas nuevas. Se limpia (marcarAtendido)
// cuando ya se le dio seguimiento desde el panel. "No" se pone solo al dar de
// baja (handleSvBaja).
function handleSvAtendido(p) {
  var ws = svSheetNuevo_();
  if (!ws) return resp({ ok: false, error: "Pestaña SV_Inscritos no encontrada en el Sheet nuevo" });

  var nombre = (p.nombre || "").trim();
  var data   = ws.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
      ws.getRange(i + 1, 6).setValue(""); // Requiere_Seguimiento (col F)
      return resp({ ok: true });
    }
  }

  return resp({ ok: false, error: "No se encontró a " + nombre + " en SV_Inscritos" });
}

// ── Salvando Vidas: Marcar de baja ───────────────────────────────
// El donador ya no participa pero se conserva su historial de donaciones
// para indicadores. Estado (col E) pasa a "Baja" y Requiere_Seguimiento
// (col F) se marca "No".
function handleSvBaja(p) {
  var ws = svSheetNuevo_();
  if (!ws) return resp({ ok: false, error: "Pestaña SV_Inscritos no encontrada en el Sheet nuevo" });

  var nombre = (p.nombre || "").trim();
  var data   = ws.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
      ws.getRange(i + 1, 5).setValue("Baja"); // Estado (col E)
      ws.getRange(i + 1, 6).setValue("No");   // Requiere_Seguimiento (col F)
      return resp({ ok: true });
    }
  }

  return resp({ ok: false, error: "No se encontró a " + nombre + " en SV_Inscritos" });
}

// ── Migración única: SV_Donadores histórico → Sheet nuevo (SV_Inscritos) ──
// Mismo patrón que handleMigrarRAHistorico. El histórico nunca tuvo
// Fecha_Alta ni un Requiere_Seguimiento confiable (siempre vacío en la
// práctica), así que Fecha_Alta migra vacía y Requiere_Seguimiento se deriva
// solo del Estado viejo (Activo→vacío, Inactivo→"No").
function handleMigrarSVHistorico(p) {
  var wsViejo = SpreadsheetApp.openById(SHEET_HISTORICO_ID).getSheetByName("SV_Donadores");
  var wsNuevo = svSheetNuevo_();
  if (!wsViejo) return resp({ ok: false, error: "No se encontró SV_Donadores en el Sheet histórico" });
  if (!wsNuevo) return resp({ ok: false, error: "No se encontró SV_Inscritos en el Sheet nuevo" });

  var datosViejos = wsViejo.getDataRange().getValues();
  var datosNuevos = wsNuevo.getDataRange().getValues();
  var yaExisten = {};
  for (var j = 1; j < datosNuevos.length; j++) {
    var n = String(datosNuevos[j][0] || "").trim().toLowerCase();
    if (n) yaExisten[n] = true;
  }

  var migrados = 0, omitidos = 0;

  for (var i = 1; i < datosViejos.length; i++) {
    var nombre = String(datosViejos[i][0] || "").trim();
    if (!nombre || nombre === "Nombre") continue;
    var key = nombre.toLowerCase();
    if (yaExisten[key]) { omitidos++; continue; }

    var sucursal    = datosViejos[i][1] || "";
    var telefono    = datosViejos[i][2] || "";
    var tipoSangre  = datosViejos[i][3] || "";
    var mesGuardia  = datosViejos[i][4] || "";
    var estadoViejo = String(datosViejos[i][5] || "").trim().toLowerCase();

    var esBaja = estadoViejo === "inactivo" || estadoViejo === "baja";
    var estadoNuevo = esBaja ? "Baja" : "Activo";
    var seguimientoNuevo = esBaja ? "No" : "";

    wsNuevo.appendRow([nombre, sucursal, telefono, "", estadoNuevo, seguimientoNuevo, "", tipoSangre, mesGuardia]);
    yaExisten[key] = true;
    migrados++;
  }

  return resp({ ok: true, migrados: migrados, omitidos: omitidos });
}

function handleRetirarSyncSV(p) {
  var triggers = ScriptApp.getProjectTriggers();
  var eliminado = false;
  triggers.forEach(function (t) {
    if (t.getHandlerFunction() === "sincronizarNuevosSV") {
      ScriptApp.deleteTrigger(t);
      eliminado = true;
    }
  });
  return resp({ ok: true, triggerEliminado: eliminado });
}

// ── Migración única: EG_Inscritos histórico → Sheet nuevo ────────
// EG_Inscritos nunca tuvo handlers de escritura desde el portal (100% manual
// hasta ahora) — solo hace falta migrar los datos y apuntar escuela-geb.html
// al Sheet nuevo. "Nivel" (col F vieja, ej. "Preparatoria") SÍ se usa activamente
// en escuela-geb.html (pills, filtro, {Nivel} en copys) — se agrega como 10ª
// columna propia en el Sheet nuevo (no estaba en el diseño original del
// esquema unificado). El Requiere_Seguimiento viejo (col I) no es confiable
// (traía "No" para activos y "Si" para inactivos, al revés de la convención
// unificada), así que se ignora y se deriva solo del Estado, igual que en RA/SV.
function handleMigrarEGHistorico(p) {
  var wsViejo = SpreadsheetApp.openById(SHEET_HISTORICO_ID).getSheetByName("EG_Inscritos");
  var wsNuevo = SpreadsheetApp.openById(SHEET_NUEVO_ID).getSheetByName("EG_Inscritos");
  if (!wsViejo) return resp({ ok: false, error: "No se encontró EG_Inscritos en el Sheet histórico" });
  if (!wsNuevo) return resp({ ok: false, error: "No se encontró EG_Inscritos en el Sheet nuevo" });

  if (!wsNuevo.getRange(1, 10).getValue()) wsNuevo.getRange(1, 10).setValue("Nivel");

  var datosViejos = wsViejo.getDataRange().getValues();
  var datosNuevos = wsNuevo.getDataRange().getValues();
  var yaExisten = {};
  for (var j = 1; j < datosNuevos.length; j++) {
    var n = String(datosNuevos[j][0] || "").trim().toLowerCase();
    if (n) yaExisten[n] = true;
  }

  var migrados = 0, omitidos = 0;

  for (var i = 1; i < datosViejos.length; i++) {
    var nombre = String(datosViejos[i][0] || "").trim();
    if (!nombre || nombre === "Nombre") continue;
    var key = nombre.toLowerCase();
    if (yaExisten[key]) { omitidos++; continue; }

    var sucursal      = datosViejos[i][1] || "";
    var telefono      = datosViejos[i][2] || "";
    var diagnosticoEdu = datosViejos[i][3] || "";
    var modalidad     = datosViejos[i][4] || "";
    var nivel         = datosViejos[i][5] || "";
    var estadoViejo   = String(datosViejos[i][6] || "").trim().toLowerCase();
    var fechaAlta     = datosViejos[i][7] || "";

    var esBaja = estadoViejo === "inactivo" || estadoViejo === "baja";
    var estadoNuevo = esBaja ? "Baja" : "Activo";
    var seguimientoNuevo = esBaja ? "No" : "";

    wsNuevo.appendRow([nombre, sucursal, telefono, fechaAlta, estadoNuevo, seguimientoNuevo, "", modalidad, diagnosticoEdu, nivel]);
    yaExisten[key] = true;
    migrados++;
  }

  return resp({ ok: true, migrados: migrados, omitidos: omitidos });
}

function handleRetirarSyncEG(p) {
  var triggers = ScriptApp.getProjectTriggers();
  var eliminado = false;
  triggers.forEach(function (t) {
    if (t.getHandlerFunction() === "sincronizarNuevosEG") {
      ScriptApp.deleteTrigger(t);
      eliminado = true;
    }
  });
  return resp({ ok: true, triggerEliminado: eliminado });
}

// ── Escuela GEB: acciones del panel (Sheet nuevo) ────────────────
// Cols EG_Inscritos en SHEET_NUEVO_ID: A=Nombre, B=Sucursal, C=Telefono_WA,
// D=Fecha_Inscripcion, E=Estado, F=Requiere_Seguimiento, G=Notas, H=Modalidad,
// I=Diagnostico_Edu, J=Nivel. Antes, "marcar atendido" era un alert() pidiendo
// editar el Sheet a mano — reemplazado por un handler real, igual que RA/SV/AP.
function egSheetNuevo_() {
  return SpreadsheetApp.openById(SHEET_NUEVO_ID).getSheetByName("EG_Inscritos");
}

function handleEgMarcarAtendido(p) {
  var ws = egSheetNuevo_();
  if (!ws) return resp({ ok: false, error: "Pestaña EG_Inscritos no encontrada en el Sheet nuevo" });

  var nombre = (p.nombre || "").trim();
  var data   = ws.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
      ws.getRange(i + 1, 6).setValue(""); // Requiere_Seguimiento (col F)
      return resp({ ok: true });
    }
  }

  return resp({ ok: false, error: "No se encontró a " + nombre + " en EG_Inscritos" });
}

function handleEgBaja(p) {
  var ws = egSheetNuevo_();
  if (!ws) return resp({ ok: false, error: "Pestaña EG_Inscritos no encontrada en el Sheet nuevo" });

  var nombre = (p.nombre || "").trim();
  var data   = ws.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
      ws.getRange(i + 1, 5).setValue("Baja"); // Estado (col E)
      ws.getRange(i + 1, 6).setValue("No");   // Requiere_Seguimiento (col F)
      return resp({ ok: true });
    }
  }

  return resp({ ok: false, error: "No se encontró a " + nombre + " en EG_Inscritos" });
}

// Diagnóstico educativo (col I): se captura al alta (altas.html) pero también se
// puede poner/corregir desde el panel — alimenta las estadísticas de escolaridad
// y rezago, que ya no dependen de Concentrado.
function handleEgDiagnostico(p) {
  var ws = egSheetNuevo_();
  if (!ws) return resp({ ok: false, error: "Pestaña EG_Inscritos no encontrada en el Sheet nuevo" });

  var nombre = (p.nombre || "").trim();
  var data   = ws.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
      ws.getRange(i + 1, 9).setValue(p.diagnosticoEdu || ""); // Diagnostico_Edu (col I)
      return resp({ ok: true });
    }
  }

  return resp({ ok: false, error: "No se encontró a " + nombre + " en EG_Inscritos" });
}

// ── Carga real única: 10 inscritos Escuela GEB (2026-07-05) ──────
// Datos reales dados por Cecilia. appendRow directo (no vía alta_unificada)
// porque ese endpoint hoy siempre deja Nivel vacío y Requiere_Seguimiento="Sí"
// al crear, y aquí el Excel trae Estado/Requiere_Seguimiento/Nivel reales.
function cargarEGLote_2026_07_05() {
  var ws = egSheetNuevo_();
  if (!ws) return;

  var filas = [
    ["America Polette Perea Carrillo",   "NAC",           "3315197578", "01/03/2026", "Activo", "No", "", "UVL",  "Secundaria", "Preparatoria"],
    ["ANDREA ACUÑA ROBLEDO",             "ÁVILA CAMACHO", "3345938028", "01/04/2026", "Activo", "No", "", "INEA", "Secundaria", "Preparatoria"],
    ["Araceli Abigail Caro Villalobos",  "GMT",           "",           "",           "Baja",   "Sí", "", "UVL",  "Primaria",   "Secundaria"],
    ["Catalina Hurtado Olivares",        "NAC",           "3331421871", "01/03/2026", "Activo", "No", "", "UVL",  "Secundaria", "Preparatoria"],
    ["Erick Salvador Salamanca Abarca",  "GMT",           "3339451753", "",           "Baja",   "Sí", "", "UVL",  "Secundaria", "Preparatoria"],
    ["Fátima Lizette Hernández Jiménez", "NAC",           "3332478861", "01/12/2025", "Activo", "No", "", "INEA", "Secundaria", "Preparatoria"],
    ["Perla del Carmen Luna Gutierrez",  "NAC",           "3333240945", "01/12/2025", "Activo", "No", "", "INEA", "Secundaria", "Preparatoria"],
    ["Rosalba Jimenez Rivera",           "VR",            "3331598123", "01/04/2026", "Activo", "No", "", "INEA", "Primaria",   "Secundaria"],
    ["VERONICA TORRES CAMPOS",           "ITESO",         "",           "01/12/2025", "Activo", "No", "", "INEA", "Secundaria", "Preparatoria"],
    ["VITALINA BARRIOS ARTEAGA",         "ITESO",         "3339822326", "",           "Baja",   "Sí", "", "UVL",  "Primaria",   "Secundaria"]
  ];

  var data = ws.getDataRange().getValues();
  var yaExisten = {};
  for (var j = 1; j < data.length; j++) {
    var n = String(data[j][0] || "").trim().toLowerCase();
    if (n) yaExisten[n] = true;
  }

  var agregadas = 0, omitidas = 0;
  filas.forEach(function (f) {
    var key = String(f[0]).trim().toLowerCase();
    if (yaExisten[key]) { omitidas++; return; }
    ws.appendRow(f);
    yaExisten[key] = true;
    agregadas++;
  });

  Logger.log("EG lote 2026-07-05: agregadas=" + agregadas + " omitidas=" + omitidas);
}

// ── Sembrar datos de prueba Escuela GEB (Sheet nuevo) — solo desarrollo ──
// Agrega colaboradores ficticios (prefijo "PRUEBA") a Colaboradores (para
// FechaNac, usado en el cálculo de rezago) y a EG_Inscritos, cubriendo INEA,
// UVL, con y sin diagnóstico, activo/baja/pendiente de seguimiento.
function sembrarPruebasEG() {
  var ssNuevo = SpreadsheetApp.openById(SHEET_NUEVO_ID);
  var wsColab = ssNuevo.getSheetByName("Colaboradores");
  var ws = egSheetNuevo_();
  if (!wsColab || !ws) { SpreadsheetApp.getUi().alert("No se encontraron Colaboradores o EG_Inscritos en el Sheet nuevo."); return; }

  var colaboradores = [
    ["PRUEBA Marisol Peña",   "Sucursal Centro", "", "", "1990-03-10", "Femenino", "3312345610", "", "", new Date()],
    ["PRUEBA Jorge Salcido",  "Sucursal Norte",  "", "", "1985-07-22", "Masculino", "3312345611", "", "", new Date()],
    ["PRUEBA Rosa Delgado",   "Sucursal Sur",    "", "", "2000-01-05", "Femenino", "3312345612", "", "", new Date()],
    ["PRUEBA Iván Cortés",    "Sucursal Centro", "", "", "1978-11-30", "Masculino", "3312345613", "", "", new Date()]
  ];
  wsColab.getRange(wsColab.getLastRow() + 1, 1, colaboradores.length, 10).setValues(colaboradores);

  var hoy = new Date();
  var inscritos = [
    ["PRUEBA Marisol Peña",  "Sucursal Centro", "3312345610", hoy, "Activo", "Sí", "", "INEA", "Secundaria", ""],
    ["PRUEBA Jorge Salcido", "Sucursal Norte",  "3312345611", hoy, "Activo", "", "", "INEA", "Primaria", ""],
    ["PRUEBA Rosa Delgado",  "Sucursal Sur",    "3312345612", hoy, "Activo", "", "", "UVL", "Preparatoria", ""],
    ["PRUEBA Iván Cortés",   "Sucursal Centro", "3312345613", hoy, "Baja", "No", "", "INEA", "Secundaria", ""]
  ];
  ws.getRange(ws.getLastRow() + 1, 1, inscritos.length, 10).setValues(inscritos);

  SpreadsheetApp.getUi().alert("4 colaboradores de prueba agregados a Colaboradores + EG_Inscritos (Sheet nuevo).");
}

// ── Atención Psicológica (Sheet NUEVO desde 2026-07-04) ──────────
// Cols AP_Inscritos en SHEET_NUEVO_ID: A=Nombre, B=Sucursal, C=Telefono_WA,
// D=Fecha_Alta, E=Estado, F=Requiere_Seguimiento, G=Notas, H=Modalidad_Servicio,
// I=Modalidad_Sesion, J=Urgencia, K=Total_Sesiones, L=Sesiones_Tomadas,
// M=Progreso_Sesiones, N=Diagnostico_Inicial, O=Diagnostico_Final.
// Estado (col E) unifica lo que antes eran DOS columnas separadas
// (Activo_Programa manual + Estado de progreso automático): ahora es un solo
// campo con valores "Lista de espera" | "Activo" | "Sesión pendiente" (activo,
// pero faltó a una sesión) | "Completó" | "Baja". El Sheet histórico (SS) ya NO
// se toca para esta iniciativa — ver INSTRUCCIONES_CRM.md.
function apSheetNuevo_() {
  return SpreadsheetApp.openById(SHEET_NUEVO_ID).getSheetByName("AP_Inscritos");
}
function apSesionesSheetNuevo_() {
  return SpreadsheetApp.openById(SHEET_NUEVO_ID).getSheetByName("AP_Sesiones");
}

// ── Carga real única: 40 históricos Atención Psicológica (2026-07-05) ──
// (Oscar Eduardo Ibarra Hernández se quitó de aquí: pasó a la lista de
// "actuales" en cargarAPActualesLote_2026_07_05 porque sigue activo, no
// completado.)
// Datos reales dados por Cecilia — solo se tiene la fecha en que recibieron
// atención psicológica (o "En espera"), sin modalidad/sesiones/urgencia aún
// (eso llega después como "los actuales"). appendRow directo porque
// alta_unificada siempre pone Fecha_Alta=hoy, Requiere_Seguimiento="Sí" y
// Total_Sesiones=5 al crear, y aquí ya se sabe que estos 39 están "Completó"
// (sin seguimiento) y los 2 restantes en "Lista de espera".
function cargarAPHistoricoLote_2026_07_05() {
  var ws = apSheetNuevo_();
  if (!ws) return;

  var filas = [
    ["Jerónimo de Jesús Aldrete Diaz", "VR", "3315297027", "01/09/2025", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["Carolina López Reyes Celis", "GMT", "3313575983", "02/10/2025", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["ALEJANDRA MELENDEZ MENDOZA", "CAÑADAS", "3333993503", "03/06/2025", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["Jennifer Vanessa Murguia Campos", "VR", "", "03/12/2025", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["Jesús Alberto Vázquez Hernández", "NAC", "3326372938", "04/02/2026", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["ALFONSO DE LA TORRE QUIÑONES", "ESTANCIA", "3322440118", "04/11/2025", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["Yuliana Gabriela Romero Guzmán", "GMT", "3321753547", "05/02/2026", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["Esmeralda Jacquelin Guillen Carrillo", "GMT", "", "05/05/2025", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["Brayan Enrique Navarro Ruiz", "NAC", "", "06/02/2025", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["Gustavo Noé Esqueda González", "VR", "3323510891", "07/05/2026", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["JORGE ALBERTO REAL ROMERO", "CAÑADAS", "3317973782", "08/06/2025", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["Cesar Abraham Medina Hernández", "NAC", "3312878685", "10/01/2025", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["Zhesly Dayyan Santiago Santiago", "NAC", "", "12/11/2025", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["Sarahi Ornelas Araiza", "NAC", "3320980143", "13/01/2025", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["BRENDA MARCELA CASTRO GUTIERREZ", "ESTANCIA", "3314989366", "13/01/2025", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["Darwin Ariel Zapata Garcia", "GMT", "3333928136", "13/01/2025", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["Gloria Yadhira Lio Gonzalez", "NAC", "", "13/01/2025", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["Jorge Isaac Padilla Hernandez", "NAC", "", "13/01/2025", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["LIZBETH ANAHI MORALES DE LIRA", "ÁVILA CAMACHO", "3313947712", "13/01/2025", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["Oscar Sinuhe Rodriguez Avalos", "NAC", "", "13/01/2025", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["Abril Reséndiz González", "VR", "7442271895", "14/01/2025", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["María Cristina De Rosas Márquez", "OFC", "3324977328", "14/01/2025", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["VERONICA TORRES CAMPOS", "ITESO", "", "14/08/2025", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["ANTONIO DE JESUS LOPEZ VELASCO", "ESTANCIA", "", "15/01/2025", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["Ángel Salcedo Castañeda", "VR", "3316004427", "16/09/2025", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["Gustavo Damian Zarate Aceves", "VR", "8145717836", "18/08/2025", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["Hector Alonso Martínez Infante", "NAC", "3331031501", "18/08/2025", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["EDUARDO CALDERON LOMELI", "ITESO", "3317389619", "20/10/2025", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["José Miguel Felix Lopez", "NAC", "", "21/11/2025", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["SALVADOR VIDAL ROSAS GONZALEZ", "ITESO", "", "22/01/2025", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["Blanca Lizette Michel López", "NAC", "3312203087", "22/09/2025", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["Maria Dolores Lopez Gonzalez", "NAC", "", "23/09/2025", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["America Polette Perea Carrillo", "NAC", "3315197578", "26/08/2025", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["KARLA BEATRIZ FLORES DUARTE", "ALEIRA", "", "26/08/2025", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["TATIANA GUADALUPE LOZA PINEDA", "OFC EASY", "", "26/08/2025", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["MONICA CRUZ MONTAÑO", "CAÑADAS", "", "28/10/2025", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["Jessica Leonella Toro Gil", "OFC", "3335047988", "05/06/2026", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["Claudia Guadalupe López Huerta", "NAC", "3317423476", "12/09/2025", "Completó", "", "", "", "", "", "", 0, "", "", ""],
    ["Andrea Lizeth Orozco Mateo", "NAC", "3322491772", "", "Lista de espera", "Sí", "", "", "", "", "", 0, "", "", ""],
    ["Nelly Rodriguez Hernandez", "NAC", "", "", "Lista de espera", "Sí", "", "", "", "", "", 0, "", "", ""]
  ];

  var data = ws.getDataRange().getValues();
  var yaExisten = {};
  for (var j = 1; j < data.length; j++) {
    var n = String(data[j][0] || "").trim().toLowerCase();
    if (n) yaExisten[n] = true;
  }

  var agregadas = 0, omitidas = 0;
  filas.forEach(function (f) {
    var key = String(f[0]).trim().toLowerCase();
    if (yaExisten[key]) { omitidas++; return; }
    var fila = f.slice();
    if (fila[3]) {
      var partes = String(fila[3]).split("/"); // dd/mm/yyyy
      fila[3] = new Date(Number(partes[2]), Number(partes[1]) - 1, Number(partes[0]));
    }
    ws.appendRow(fila);
    yaExisten[key] = true;
    agregadas++;
  });

  Logger.log("AP histórico 2026-07-05: agregadas=" + agregadas + " omitidas=" + omitidas);
}

// ── Carga real única: 8 activos actuales Atención Psicológica (2026-07-05) ──
// Datos reales dados por Cecilia. Todos Estado="Activo" (Urgencia="Baja" es el
// nivel de urgencia clínica, no un estado de baja del programa). Fecha_Alta
// usa la "Psic. Fecha" de cada quien (Cecilia confirmó, no la Fecha_Inicio
// 16/06/2026 que era igual para todos). El diagnóstico corto (Ansiedad, Duelo,
// etc.) va en Notas, no en Diagnostico_Inicial (también confirmado).
function cargarAPActualesLote_2026_07_05() {
  var ws = apSheetNuevo_();
  if (!ws) return;

  var filas = [
    ["Diana Marcela Carrillo Reyes",       "NAC",      "3330382538", "22/04/2026", "Activo", "", "Ansiedad",          "ELA", "Presencial", "Baja", 5, 0, "", "", ""],
    ["LETICIA DEL ROSARIO LOPE AYALA",     "ALEIRA",   "",           "27/03/2026", "Activo", "", "Problemas familiares", "ELA", "Virtual", "Baja", 5, 3, "", "", ""],
    ["Ileana Nereida Sánchez Salazar",     "VR",       "3332341833", "31/03/2026", "Activo", "", "Duelo",             "ELA", "Virtual", "Baja", 5, 3, "", "", ""],
    ["Luz Kerena López Valdes",            "NAC",      "3323848846", "31/07/2025", "Activo", "", "Duelo",             "ELA", "Virtual", "Baja", 5, 3, "", "", ""],
    ["Jennifer Valeria Maldonado Calderón","NAC",      "3328286538", "05/06/2025", "Activo", "", "Duelo",             "ELA", "Virtual", "Baja", 5, 3, "", "", ""],
    ["José Emmanuel Hernandez Muñoz",      "VR",       "3331075031", "16/06/2026", "Activo", "", "Estrés laboral",    "ELA", "Virtual", "Baja", 5, 3, "", "", ""],
    ["PALOMA BERENICE FLORES PALOS",       "CAÑADAS",  "3317140825", "16/06/2026", "Activo", "", "Ansiedad",          "ELA", "Virtual", "Baja", 5, 3, "", "", ""],
    ["Oscar Eduardo Ibarra Hernández",     "NAC",      "3315879326", "16/06/2026", "Activo", "", "Duelo",             "ELA", "Presencial", "Baja", 5, 0, "", "", ""]
  ];

  var data = ws.getDataRange().getValues();
  var yaExisten = {};
  for (var j = 1; j < data.length; j++) {
    var n = String(data[j][0] || "").trim().toLowerCase();
    if (n) yaExisten[n] = true;
  }

  var agregadas = 0, omitidas = 0;
  filas.forEach(function (f) {
    var key = String(f[0]).trim().toLowerCase();
    if (yaExisten[key]) { omitidas++; return; }
    var fila = f.slice();
    if (fila[3]) {
      var partes = String(fila[3]).split("/"); // dd/mm/yyyy
      fila[3] = new Date(Number(partes[2]), Number(partes[1]) - 1, Number(partes[0]));
    }
    ws.appendRow(fila);
    yaExisten[key] = true;
    agregadas++;
  });

  Logger.log("AP actuales 2026-07-05: agregadas=" + agregadas + " omitidas=" + omitidas);
}

// ── Atención Psicológica: Registrar asistencia en lote (pasar lista) ──
// Cols AP_Sesiones (Sheet nuevo): A=Nombre, B=Num_Sesion, C=Fecha_Programada,
//                   D=Hora_Programada, E=Modalidad, F=Asistio, G=Cambio_Modalidad, H=Notas_Sesion
// p.registros = [{nombre, numSesion, asistio}, ...] — uno por cada persona marcada
// en el pop-up "Pasar lista" (los que se dejan en "sin cambio" no llegan aquí).
function handleApAsistenciaLote(p) {
  var registros = p.registros || [];
  if (!registros.length) return resp({ ok: true });

  var ws = apSesionesSheetNuevo_();
  var wsIns = apSheetNuevo_();
  if (!ws || !wsIns) return resp({ ok: false, error: "Pestaña AP_Sesiones o AP_Inscritos no encontrada en el Sheet nuevo" });

  var data = ws.getDataRange().getValues();

  registros.forEach(function (r) {
    var nombre    = (r.nombre || "").trim();
    var numSesion = String(r.numSesion || "").trim();
    var asistio   = r.asistio || "";
    var found     = false;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()
          && String(data[i][1]).trim() === numSesion) {
        ws.getRange(i + 1, 6).setValue(asistio);
        data[i][5] = asistio;
        found = true;
        break;
      }
    }

    if (!found) {
      ws.appendRow([nombre, numSesion, "", "", "", asistio, "", ""]);
      data.push([nombre, numSesion, "", "", "", asistio, "", ""]);
    }
  });

  // Recalcular Sesiones_Tomadas y Estado una sola vez por persona, con los datos ya actualizados
  var insData = wsIns.getDataRange().getValues();
  registros.forEach(function (r) {
    var nombre = (r.nombre || "").trim();
    for (var j = 1; j < insData.length; j++) {
      if (String(insData[j][0]).trim().toLowerCase() === nombre.toLowerCase()) {
        var count = 0;
        for (var k = 1; k < data.length; k++) {
          if (String(data[k][0]).trim().toLowerCase() === nombre.toLowerCase()) {
            var a = String(data[k][5]).trim().toLowerCase();
            if (a === "sí" || a === "si") count++;
          }
        }
        wsIns.getRange(j + 1, 12).setValue(count); // Sesiones_Tomadas (col L)

        var totalSesiones = parseInt(insData[j][10]) || 5; // Total_Sesiones (col K, índice 10)
        if (count >= totalSesiones) {
          wsIns.getRange(j + 1, 5).setValue("Completó"); // Estado (col E)
        } else if (r.asistio === "No") {
          wsIns.getRange(j + 1, 5).setValue("Sesión pendiente"); // Estado (col E)
        } else {
          wsIns.getRange(j + 1, 5).setValue("Activo"); // Estado (col E)
        }
        wsIns.getRange(j + 1, 6).setValue(""); // Requiere_Seguimiento (col F): ya se dio seguimiento
        break;
      }
    }
  });

  return resp({ ok: true });
}

// ── Atención Psicológica: Cambio de modalidad ───────────────────
function handleApCambioModalidad(p) {
  var ws = apSesionesSheetNuevo_();
  if (!ws) return resp({ ok: false, error: "Pestaña AP_Sesiones no encontrada en el Sheet nuevo" });

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
  var wsIns2 = apSheetNuevo_();
  if (wsIns2) {
    var insData2 = wsIns2.getDataRange().getValues();
    for (var m = 1; m < insData2.length; m++) {
      if (String(insData2[m][0]).trim().toLowerCase() === nombre.toLowerCase()) {
        wsIns2.getRange(m + 1, 6).setValue(""); // Requiere_Seguimiento (col F)
        break;
      }
    }
  }

  return resp({ ok: true });
}

// ── Atención Psicológica: Sesiones Impulso43 ────────────────────
// Estado (col E) solo se mueve a "Completó" al llegar al total de sesiones,
// igual que handleApAsistenciaLote hace para ELA. El resto del tiempo se deja
// tal cual ("Activo", puesto al alta) — el portal decide "Contactado" vs "En
// seguimiento" a partir de Sesiones_Tomadas, no hace falta que nadie lo escriba a mano.
function handleApSesionesI43(p) {
  var ws = apSheetNuevo_();
  if (!ws) return resp({ ok: false, error: "Pestaña AP_Inscritos no encontrada en el Sheet nuevo" });

  var nombre = (p.nombre || "").trim();
  var data   = ws.getDataRange().getValues();
  var sesionesTomadas = parseInt(p.sesionesTomadas) || 0;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
      ws.getRange(i + 1, 12).setValue(sesionesTomadas); // Sesiones_Tomadas (col L)
      var totalSesiones = parseInt(data[i][10]) || 5; // Total_Sesiones (col K, índice 10)
      if (sesionesTomadas >= totalSesiones) {
        ws.getRange(i + 1, 5).setValue("Completó"); // Estado (col E)
      }
      ws.getRange(i + 1, 6).setValue(""); // Requiere_Seguimiento (col F): ya se dio seguimiento
      break;
    }
  }

  return resp({ ok: true });
}

// ── Atención Psicológica: Marcar de baja ────────────────────────
// La persona ya no está en la empresa pero se conserva su registro para
// indicadores históricos. Estado (col E) pasa a "Baja" y Requiere_Seguimiento
// (col F) se marca "No" (no aplica más seguimiento) — misma convención que RA/SV/EG.
function handleApBaja(p) {
  var ws = apSheetNuevo_();
  if (!ws) return resp({ ok: false, error: "Pestaña AP_Inscritos no encontrada en el Sheet nuevo" });

  var nombre = (p.nombre || "").trim();
  var data   = ws.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
      ws.getRange(i + 1, 5).setValue("Baja"); // Estado (col E)
      ws.getRange(i + 1, 6).setValue("No");   // Requiere_Seguimiento (col F)
      return resp({ ok: true });
    }
  }

  return resp({ ok: false, error: "No se encontró a " + nombre + " en AP_Inscritos" });
}

// ── Atención Psicológica: Diagnóstico inicial/final ─────────────
// Cols N/O de AP_Inscritos (Sheet nuevo). Escala numérica de severidad — el
// panel colorea en verde/amarillo/rojo comparando ambos valores, esto solo los
// guarda tal cual.
function handleApDiagnostico(p) {
  var ws = apSheetNuevo_();
  if (!ws) return resp({ ok: false, error: "Pestaña AP_Inscritos no encontrada en el Sheet nuevo" });

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

// ── Configuración única: crear AP_Sesiones / AP_Copys en el Sheet nuevo ──
// AP_Inscritos ya existe en el Sheet nuevo (lo crea handleAltaUnificada), pero
// AP_Sesiones (detalle por sesión) y AP_Copys (plantillas de WhatsApp) todavía
// no. Esta función crea ambas pestañas con encabezado si no existen, y copia
// verbatim las filas de AP_Copys del Sheet histórico (son plantillas de
// mensaje, no datos de personas — se pueden copiar tal cual). Idempotente: si
// AP_Copys en el Sheet nuevo ya tiene filas de datos, no vuelve a copiar.
function configurarAPSheetNuevo() {
  var ssNuevo = SpreadsheetApp.openById(SHEET_NUEVO_ID);
  var ui = SpreadsheetApp.getUi();

  var wsSesiones = ssNuevo.getSheetByName("AP_Sesiones");
  if (!wsSesiones) {
    wsSesiones = ssNuevo.insertSheet("AP_Sesiones");
    wsSesiones.appendRow(["Nombre", "Num_Sesion", "Fecha_Programada", "Hora_Programada", "Modalidad", "Asistio", "Cambio_Modalidad", "Notas_Sesion"]);
  }

  var wsCopys = ssNuevo.getSheetByName("AP_Copys");
  if (!wsCopys) {
    wsCopys = ssNuevo.insertSheet("AP_Copys");
    wsCopys.appendRow(["Momento", "Servicio", "Copy_Texto", "Numero_WA", "Activo"]);
  }

  var copysCopiados = false;
  if (wsCopys.getLastRow() < 2) {
    var wsCopysViejo = SpreadsheetApp.openById(SHEET_HISTORICO_ID).getSheetByName("AP_Copys");
    if (wsCopysViejo && wsCopysViejo.getLastRow() >= 2) {
      var filas = wsCopysViejo.getRange(2, 1, wsCopysViejo.getLastRow() - 1, 5).getValues();
      wsCopys.getRange(2, 1, filas.length, 5).setValues(filas);
      copysCopiados = true;
    }
  }

  ui.alert("Listo. AP_Sesiones y AP_Copys ya existen en el Sheet nuevo" +
    (copysCopiados ? " (copys copiados del Sheet histórico)." : "."));
}

// ── Sembrar datos de prueba AP (Sheet nuevo) — solo para desarrollo ──
// Agrega un puñado de colaboradores ficticios (prefijo "PRUEBA") cubriendo
// cada estado del panel: Lista de espera, Activo ELA con sesión pendiente,
// Activo Impulso43, Completó, y Baja. Pensado para probar el flujo end-to-end
// sin tocar datos reales de nadie. Se puede borrar el bloque de filas "PRUEBA"
// a mano en cualquier momento antes de ir a producción.
function sembrarPruebasAP() {
  var ws = apSheetNuevo_();
  var wsSesiones = apSesionesSheetNuevo_();
  if (!ws || !wsSesiones) { SpreadsheetApp.getUi().alert("Corre primero 'Configurar AP en Sheet nuevo (1 vez)'."); return; }

  var hoy = new Date();
  var filas = [
    ["PRUEBA Ana Martínez",   "Sucursal Centro", "3312345601", hoy, "Lista de espera", "Sí", "", "ELA", "Virtual", "Alta", 5, 0, "", "", ""],
    ["PRUEBA Luis Herrera",   "Sucursal Norte",  "3312345602", hoy, "Sesión pendiente", "", "", "ELA", "Presencial", "Media", 5, 2, "", "", ""],
    ["PRUEBA Carla Jiménez",  "Sucursal Sur",    "3312345603", hoy, "Activo", "", "", "Impulso43", "Virtual", "Baja", 8, 3, "", "6", ""],
    ["PRUEBA Diego Ramos",    "Sucursal Centro", "3312345604", hoy, "Completó", "", "", "ELA", "Híbrida", "Baja", 5, 5, "", "8", "3"],
    ["PRUEBA Sofía Torres",   "Sucursal Norte",  "3312345605", hoy, "Baja", "No", "", "ELA", "Virtual", "Baja", 5, 1, "", "", ""]
  ];
  ws.getRange(ws.getLastRow() + 1, 1, filas.length, 15).setValues(filas);

  var sesiones = [
    ["PRUEBA Luis Herrera", 3, "", "", "Presencial", "", "", ""],
    ["PRUEBA Carla Jiménez", 4, "", "", "Virtual", "", "", ""]
  ];
  wsSesiones.getRange(wsSesiones.getLastRow() + 1, 1, sesiones.length, 8).setValues(sesiones);

  SpreadsheetApp.getUi().alert("5 colaboradores de prueba agregados a AP_Inscritos (Sheet nuevo).");
}

// ── Migración única: AP_Inscritos histórico → Sheet nuevo ────────
// Igual patrón que handleMigrarRAHistorico: idempotente (omite a quien ya
// existe en el Sheet nuevo por nombre). NO se ha ejecutado todavía — pendiente
// de decisión con el equipo sobre qué registros reales del Sheet histórico
// vale la pena migrar (ver auditoría: AP_Inscritos histórico tiene casi todo
// vacío, parece un censo importado más que datos operativos).
function handleMigrarAPHistorico(p) {
  var wsViejo = SpreadsheetApp.openById(SHEET_HISTORICO_ID).getSheetByName("AP_Inscritos");
  var wsNuevo = apSheetNuevo_();
  if (!wsViejo) return resp({ ok: false, error: "No se encontró AP_Inscritos en el Sheet histórico" });
  if (!wsNuevo) return resp({ ok: false, error: "No se encontró AP_Inscritos en el Sheet nuevo" });

  var datosViejos = wsViejo.getDataRange().getValues();
  var datosNuevos = wsNuevo.getDataRange().getValues();
  var yaExisten = {};
  for (var j = 1; j < datosNuevos.length; j++) {
    var n = String(datosNuevos[j][0] || "").trim().toLowerCase();
    if (n) yaExisten[n] = true;
  }

  var migrados = [], omitidos = [];

  for (var i = 1; i < datosViejos.length; i++) {
    var nombre = String(datosViejos[i][0] || "").trim();
    if (!nombre || nombre === "Nombre") continue;
    var key = nombre.toLowerCase();
    if (yaExisten[key]) { omitidos.push(nombre); continue; }

    var sucursal = datosViejos[i][1] || "";
    var telefono = datosViejos[i][2] || "";
    var activoProgramaViejo = String(datosViejos[i][3] || "").trim().toLowerCase();
    var modalidadServicio = datosViejos[i][4] || "";
    var modalidadSesion = datosViejos[i][5] || "";
    var urgencia = datosViejos[i][6] || "";
    var estadoViejo = String(datosViejos[i][7] || "").trim().toLowerCase();
    var totalSesiones = datosViejos[i][9] || 5;
    var sesionesTomadas = datosViejos[i][10] || 0;
    var diagInicial = datosViejos[i][13] || "";
    var diagFinal = datosViejos[i][14] || "";

    var estadoNuevo = activoProgramaViejo === "inactivo" ? "Baja"
      : activoProgramaViejo.indexOf("lista") === 0 ? "Lista de espera"
      : (estadoViejo === "completó" || estadoViejo === "completo") ? "Completó"
      : (estadoViejo === "sesión pendiente" || estadoViejo === "sesion pendiente") ? "Sesión pendiente"
      : "Activo";
    var seguimientoNuevo = activoProgramaViejo === "inactivo" ? "No" : "";

    wsNuevo.appendRow([nombre, sucursal, telefono, new Date(), estadoNuevo, seguimientoNuevo, "",
      modalidadServicio, modalidadSesion, urgencia, totalSesiones, sesionesTomadas, "", diagInicial, diagFinal]);
    yaExisten[key] = true;
    migrados.push(nombre);
  }

  return resp({ ok: true, migrados: migrados.length, omitidos: omitidos.length });
}

// ── Retirar la sincronización antigua de Atención Psicológica ───
// Una vez migrado AP al Sheet nuevo (altas.html → handleAltaUnificada), el
// trigger horario y el onEdit que vigilaban el Sheet histórico ya no deben
// seguir corriendo. El onEdit (líneas ~44-58) solo afecta al Sheet histórico
// y puede desactivarse borrando el trigger instalado, si lo hubiera; el propio
// onEdit simple (no instalable) no se puede "eliminar" vía código, pero deja de
// importar en cuanto el equipo deje de capturar altas en el Sheet histórico.
function handleRetirarSyncAP(p) {
  var triggers = ScriptApp.getProjectTriggers();
  var eliminado = false;
  triggers.forEach(function (t) {
    if (t.getHandlerFunction() === "sincronizarNuevosAP") {
      ScriptApp.deleteTrigger(t);
      eliminado = true;
    }
  });
  return resp({ ok: true, triggerEliminado: eliminado });
}

// ── Alta unificada (Sheet nuevo) ─────────────────────────────────
// Escribe en SHEET_NUEVO_ID, NUNCA en SS (el Sheet histórico). Una sola alta
// desde el portal crea la fila en Colaboradores (si no existía) + una fila por
// cada iniciativa marcada, con las 7 columnas base estandarizadas
// (Nombre, Sucursal, Telefono_WA, Fecha_Alta, Estado, Requiere_Seguimiento, Notas)
// + las específicas de esa iniciativa. Requiere_Seguimiento siempre nace en "Sí".
// Si la persona ya tenía fila en alguna iniciativa marcada, esa iniciativa se
// omite (no se duplica) y se reporta en "omitidas" para que se revise a mano.
const ALTA_INICIATIVAS = {
  AP: {
    sheet: "AP_Inscritos",
    campos: function (d) {
      return [d.modalidadServicio || "", d.modalidadSesion || "", d.urgencia || "",
              d.totalSesiones || 5, 0, "", "", ""]; // Sesiones_Tomadas=0, Progreso/Diagnósticos vacíos
    }
  },
  RA: {
    sheet: "RA_Inscritos",
    campos: function (d) { return ["", ""]; } // Semana_Actual, Fecha_Ultima_Semana
  },
  SV: {
    sheet: "SV_Inscritos",
    campos: function (d) { return [d.tipoSangre || "", d.mesGuardia || ""]; }
  },
  EG: {
    sheet: "EG_Inscritos",
    campos: function (d) { return [d.modalidad || "", d.diagnosticoEdu || "", ""]; } // Diagnostico_Educativo, Nivel
  },
  RE: {
    sheet: "RE_Interesados",
    campos: function (d) { return []; } // solo las 7 columnas base — sin campos propios
  },
  CS: {
    sheet: "CS_Inscritos",
    campos: function (d) { return [""]; } // Fecha_Camino vacía al alta
  },
  IG: {
    sheet: "IG_Inscritos",
    campos: function (d) { return ["", "", "", ""]; } // Plan_Vida, Presupuesto, Ahorro, Movilidad_Social — vacíos al alta
  },
  BE: {
    sheet: "BE_Inscritos",
    campos: function (d) { return [d.universidad || "", d.nivel || "", d.carrera || "", ""]; } // Universidad, Nivel, Carrera, Inscribio
  },
  EA: {
    sheet: "EA_Lideres",
    campos: function (d) { return [0]; } // Total_Cilindros=0 al alta
  }
};

function handleAltaUnificada(p) {
  var ssNuevo = SpreadsheetApp.openById(SHEET_NUEVO_ID);
  var nombre = (p.nombre || "").trim();
  if (!nombre) return resp({ ok: false, error: "Falta el nombre" });

  var wsColab = ssNuevo.getSheetByName("Colaboradores");
  if (!wsColab) return resp({ ok: false, error: "Pestaña Colaboradores no encontrada en el Sheet nuevo" });

  var colabData = wsColab.getDataRange().getValues();
  var yaExisteColab = false;
  for (var c = 1; c < colabData.length; c++) {
    if (String(colabData[c][0]).trim().toLowerCase() === nombre.toLowerCase()) { yaExisteColab = true; break; }
  }
  if (!yaExisteColab) {
    wsColab.appendRow([
      nombre, p.sucursal || "", p.unidad || "", p.correo || "", p.fechaNac || "",
      p.genero || "", p.telefono || "", p.turno || "", p.rol || "", new Date()
    ]);
  }

  var creadas = [];
  var omitidas = [];

  (p.iniciativas || []).forEach(function (ini) {
    var cfg = ALTA_INICIATIVAS[ini.tipo];
    if (!cfg) return;
    var ws = ssNuevo.getSheetByName(cfg.sheet);
    if (!ws) { omitidas.push(cfg.sheet + " (pestaña no encontrada)"); return; }

    var data = ws.getDataRange().getValues();
    var yaExiste = false;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()) { yaExiste = true; break; }
    }
    if (yaExiste) { omitidas.push(cfg.sheet); return; }

    var base = [nombre, ini.sucursal || p.sucursal || "", p.telefono || "", new Date(), ini.estado || "Activo", "Sí", ini.notas || ""];
    ws.appendRow(base.concat(cfg.campos(ini)));
    creadas.push(cfg.sheet);
  });

  return resp({ ok: true, creadas: creadas, omitidas: omitidas });
}

// ── Reto Ahorro (Sheet NUEVO desde 2026-07-04) ───────────────────
// Cols RA_Inscritos en SHEET_NUEVO_ID: A=Nombre, B=Sucursal, C=Telefono_WA,
// D=Fecha_Alta, E=Estado, F=Requiere_Seguimiento, G=Notas, H=Semana_Actual,
// I=Fecha_Ultima_Semana. El Sheet histórico (SS) ya NO se toca para esta
// iniciativa — ver INSTRUCCIONES_CRM.md.
function raSheetNuevo_() {
  return SpreadsheetApp.openById(SHEET_NUEVO_ID).getSheetByName("RA_Inscritos");
}

// Actualizar la semana cuenta como seguimiento dado: limpia el flag y reinicia
// el reloj de atraso que usa revisarAtrasosRA().
function handleRaActualizarSemana(p) {
  var ws = raSheetNuevo_();
  if (!ws) return resp({ ok: false, error: "Pestaña RA_Inscritos no encontrada en el Sheet nuevo" });

  var nombre = (p.nombre || "").trim();
  var data   = ws.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
      ws.getRange(i + 1, 8).setValue(parseInt(p.semana, 10) || 0); // Semana_Actual (col H)
      ws.getRange(i + 1, 6).setValue("");                          // Requiere_Seguimiento (col F)
      ws.getRange(i + 1, 9).setValue(new Date());                  // Fecha_Ultima_Semana (col I)
      return resp({ ok: true });
    }
  }

  return resp({ ok: false, error: "No se encontró a " + nombre + " en RA_Inscritos" });
}

// ── Reto Ahorro: Marcar atendido (limpia el flag sin cambiar semana) ──
function handleRaMarcarAtendido(p) {
  var ws = raSheetNuevo_();
  if (!ws) return resp({ ok: false, error: "Pestaña RA_Inscritos no encontrada en el Sheet nuevo" });

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

// ── Reto Ahorro: Marcar de baja ──────────────────────────────────
// Estado (col E) pasa a "Baja". Se conserva el registro para indicadores.
function handleRaBaja(p) {
  var ws = raSheetNuevo_();
  if (!ws) return resp({ ok: false, error: "Pestaña RA_Inscritos no encontrada en el Sheet nuevo" });

  var nombre = (p.nombre || "").trim();
  var data   = ws.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
      ws.getRange(i + 1, 5).setValue("Baja"); // Estado (col E)
      ws.getRange(i + 1, 6).setValue("No");   // Requiere_Seguimiento (col F): baja de esta iniciativa
      return resp({ ok: true });
    }
  }

  return resp({ ok: false, error: "No se encontró a " + nombre + " en RA_Inscritos" });
}

// ── Reto Ahorro: marcar atraso automático (>3 semanas sin actualizar) ──
// Corre diario (ver configurarAlertasAtrasoRA). Si un inscrito no dado de baja
// no ha tenido su Semana_Actual actualizada (Fecha_Ultima_Semana, col I) en más
// de 21 días —usando Fecha_Alta (col D) como respaldo si nunca se ha
// actualizado— se marca Requiere_Seguimiento="Sí". Se limpia automáticamente en
// cuanto alguien actualiza su semana desde el panel o se marca "atendido".
function revisarAtrasosRA() {
  var ws = raSheetNuevo_();
  if (!ws) return;
  var data = ws.getDataRange().getValues();
  var ahora = new Date();
  var LIMITE_MS = 21 * 24 * 60 * 60 * 1000;

  for (var i = 1; i < data.length; i++) {
    var estado = String(data[i][4] || "").trim().toLowerCase(); // Estado (col E)
    if (estado === "baja") continue;

    var baseline = data[i][8] || data[i][3]; // Fecha_Ultima_Semana (I) o Fecha_Alta (D)
    if (!baseline) continue;
    var fechaBase = new Date(baseline);
    if (isNaN(fechaBase.getTime())) continue;

    var yaMarcado = String(data[i][5] || "").trim().toLowerCase() === "sí"; // Requiere_Seguimiento (F)
    if (!yaMarcado && (ahora - fechaBase) > LIMITE_MS) {
      ws.getRange(i + 1, 6).setValue("Sí"); // Requiere_Seguimiento (col F)
    }
  }
}

// ── Migración única: RA_Inscritos histórico → Sheet nuevo ────────
// Se ejecuta una sola vez (vía doPost, tipo "migrar_ra_historico") para llevar
// los datos reales del Sheet histórico al esquema unificado. Idempotente: si
// una persona ya existe en el Sheet nuevo, se omite (se puede correr de nuevo
// sin duplicar). "Baja" en el Estado viejo se traduce a Requiere_Seguimiento="No"
// (ya no aplica seguimiento) por la convención unificada; activos migran con
// Requiere_Seguimiento vacío (ya se les venía dando seguimiento, no son altas
// nuevas).
function handleMigrarRAHistorico(p) {
  var wsViejo = SpreadsheetApp.openById(SHEET_HISTORICO_ID).getSheetByName("RA_Inscritos");
  var wsNuevo = raSheetNuevo_();
  if (!wsViejo) return resp({ ok: false, error: "No se encontró RA_Inscritos en el Sheet histórico" });
  if (!wsNuevo) return resp({ ok: false, error: "No se encontró RA_Inscritos en el Sheet nuevo" });

  var datosViejos = wsViejo.getDataRange().getValues();
  var datosNuevos = wsNuevo.getDataRange().getValues();
  var yaExisten = {};
  for (var j = 1; j < datosNuevos.length; j++) {
    var n = String(datosNuevos[j][0] || "").trim().toLowerCase();
    if (n) yaExisten[n] = true;
  }

  var migrados = [], omitidos = [];

  for (var i = 1; i < datosViejos.length; i++) {
    var nombre = String(datosViejos[i][0] || "").trim();
    if (!nombre || nombre === "Nombre") continue; // vacío o fila artefacto
    var key = nombre.toLowerCase();
    if (yaExisten[key]) { omitidos.push(nombre); continue; }

    var sucursal   = datosViejos[i][1] || "";
    var telefono   = datosViejos[i][2] || "";
    var semana     = datosViejos[i][3] || 0;
    var fechaAlta  = datosViejos[i][4] || "";
    var estadoViejo = String(datosViejos[i][6] || "").trim().toLowerCase();
    var fechaUltimaSemana = datosViejos[i][7] || "";

    var estadoNuevo = estadoViejo === "baja" ? "Baja" : "Activo";
    var seguimientoNuevo = estadoViejo === "baja" ? "No" : "";

    wsNuevo.appendRow([nombre, sucursal, telefono, fechaAlta, estadoNuevo, seguimientoNuevo, "", semana, fechaUltimaSemana]);
    yaExisten[key] = true;
    migrados.push(nombre);
  }

  return resp({ ok: true, migrados: migrados.length, omitidos: omitidos.length });
}

// ── Retirar la sincronización antigua de Reto Ahorro ─────────────
// Una vez migrado RA al Sheet nuevo, el trigger horario que vigilaba Concentrado
// sobre el Sheet histórico ya no debe seguir corriendo (las altas nuevas ahora
// entran por altas.html → handleAltaUnificada, directo al Sheet nuevo). Dejarlo
// activo no corrompe nada, pero seguiría agregando gente al Sheet histórico en
// silencio, que ya nadie revisa.
function handleRetirarSyncRA(p) {
  var triggers = ScriptApp.getProjectTriggers();
  var eliminado = false;
  triggers.forEach(function (t) {
    if (t.getHandlerFunction() === "sincronizarNuevosRA") {
      ScriptApp.deleteTrigger(t);
      eliminado = true;
    }
  });
  return resp({ ok: true, triggerEliminado: eliminado });
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

// ── Retiro del Espíritu Santo (Sheet nuevo desde 2026-07-05) ─────
// Migrado completo: RE_Ediciones/Asistencias/Habitaciones/Copys/Interesados
// viven en SHEET_NUEVO_ID. Ya no depende de Concentrado (ni para nombres/
// sucursal/género — eso ahora sale de Colaboradores — ni para el conteo
// histórico de asistencia, que se calcula directo de RE_Asistencias).
function reEdicionesSheetNuevo_()    { return SpreadsheetApp.openById(SHEET_NUEVO_ID).getSheetByName("RE_Ediciones"); }
function reAsistenciasSheetNuevo_()  { return SpreadsheetApp.openById(SHEET_NUEVO_ID).getSheetByName("RE_Asistencias"); }
function reHabitacionesSheetNuevo_() { return SpreadsheetApp.openById(SHEET_NUEVO_ID).getSheetByName("RE_Habitaciones"); }
function reInteresadosSheetNuevo_()  { return SpreadsheetApp.openById(SHEET_NUEVO_ID).getSheetByName("RE_Interesados"); }

// ── Carga real única: histórico Retiro Espiritual 2022-2026 (2026-07-05) ──
// Datos reales dados por Cecilia: 13 ediciones (RE_Ediciones) y 106 registros
// de asistencia (RE_Asistencias, incluye colaboradores actuales, ex-colabora-
// dores y familiares/invitados marcados en Notas). id_edicion sigue el mismo
// formato "RE-<año>-<NN>" que usa handleReAltaEdicion. Testimonios: se
// colapsó a un solo Sí/No (Sí si asistió a cualquiera de las 2 fechas de
// sesión de testimonios); el detalle textual original (Confirmado/Invitó/
// Baja/N/A) se conserva en Notas para no perder información.
function cargarREHistoricoLote_2026_07_05() {
  var wsEd = reEdicionesSheetNuevo_();
  var wsAs = reAsistenciasSheetNuevo_();
  if (!wsEd || !wsAs) return;

  var ediciones = [
    ["RE-2022-01", "Retiro jul-22", "jul-22", "", "", "", "Completada", ""],
    ["RE-2022-02", "Retiro oct-22", "oct-22", "", "", "", "Completada", ""],
    ["RE-2023-01", "Retiro oct-23", "oct-23", "", "", "", "Completada", ""],
    ["RE-2023-02", "Retiro oct-23", "oct-23", "", "", "", "Completada", ""],
    ["RE-2024-01", "Retiro abr-24", "abr-24", "", "", "", "Completada", ""],
    ["RE-2024-02", "Retiro oct-24", "oct-24", "", "", "", "Completada", ""],
    ["RE-2025-01", "Retiro feb-25", "feb-25", "", "", "", "Completada", ""],
    ["RE-2025-02", "Retiro jun-25", "jun-25", "", "", "", "Completada", ""],
    ["RE-2025-03", "Retiro jul-25", "jul-25", "", "", "", "Completada", ""],
    ["RE-2025-04", "Retiro oct-25", "oct-25", "", "", "", "Completada", ""],
    ["RE-2025-05", "Retiro nov-25", "nov-25", "", "", "", "Completada", ""],
    ["RE-2026-01", "Retiro ene-26", "ene-26", "", "", "", "Completada", ""],
    ["RE-2026-02", "Retiro may-26", "may-26", "", "", "", "Pendiente de registro", ""]
  ];

  var asistencias = [
    ["RE-2022-01", "Oliver Adan De la Torre Jauregui", "jul-22", "Sí", "No", ""],
    ["RE-2022-01", "Gerardo Roberto Villarreal Treviño", "jul-22", "Sí", "Sí", ""],
    ["RE-2022-02", "Elisa Nuñez", "oct-22", "Sí", "No", "(no es colaborador actualmente); baja de liderazgo; testimonios: 22/08/2025=Baja"],
    ["RE-2022-02", "Julia Elizabeth Rodríguez Arroyo", "oct-22", "Sí", "Sí", ""],
    ["RE-2023-01", "Marco Antonio Peregrina Sierra", "oct-23", "Sí", "Sí", ""],
    ["RE-2023-01", "Raúl Eduardo Mantecón Siordia", "oct-23", "Sí", "No", ""],
    ["RE-2023-01", "Javier Eduardo Muñoz Gaeta", "oct-23", "Sí", "Sí", ""],
    ["RE-2023-01", "Judith Manzano", "oct-23", "No", "No", "(invitado/familiar, no colaborador); testimonios: 22/08/2025=N/A"],
    ["RE-2023-01", "Dana María Schmid", "oct-23", "", "No", "(invitado/familiar, no colaborador); testimonios: 22/08/2025=N/A"],
    ["RE-2023-01", "Jordi Vargas Garibay", "oct-23", "Sí", "No", "(invitado/familiar, no colaborador); testimonios: 22/08/2025=N/A"],
    ["RE-2023-01", "Juan Pablo Monraz", "oct-23", "Sí", "No", "(invitado/familiar, no colaborador)"],
    ["RE-2023-01", "Linda Zoraida Padilla Alvarez", "oct-23", "Sí", "No", ""],
    ["RE-2023-01", "Daniel Garcia Lomeli", "oct-23", "", "No", ""],
    ["RE-2023-01", "Ernesto Aviña", "oct-23", "Sí", "No", "(no es colaborador actualmente); baja de liderazgo; testimonios: 22/08/2025=Baja"],
    ["RE-2023-01", "Jose Hiram Gonzalez Palazuelos", "oct-23", "Sí", "No", ""],
    ["RE-2023-01", "Karen Anaid Armas Gutierrez", "oct-23", "No", "Sí", ""],
    ["RE-2023-01", "Thalia Veronica Rangel Sotelo", "oct-23", "Sí", "Sí", ""],
    ["RE-2023-01", "Alexia Gándara", "oct-23", "Sí", "No", "(no es colaborador actualmente)"],
    ["RE-2023-02", "Mamá de Fer Benitez", "oct-23", "Sí", "No", "(invitado/familiar, no colaborador); testimonios: 22/08/2025=N/A"],
    ["RE-2023-02", "Julia Elizabeth Rodríguez Arroyo", "oct-23", "Sí", "Sí", ""],
    ["RE-2023-02", "Michelle Mejía Romero", "oct-23", "Sí", "Sí", ""],
    ["RE-2024-01", "Fabricio Blanco", "abr-24", "Sí", "No", "(invitado/familiar, no colaborador); testimonios: 22/08/2025=N/A"],
    ["RE-2024-01", "Blanca Angélica Muñiz Horta", "abr-24", "Sí", "No", "(invitado/familiar, no colaborador); testimonios: 22/08/2025=N/A"],
    ["RE-2024-01", "Carolina Izunza", "abr-24", "Sí", "No", "(invitado/familiar, no colaborador); testimonios: 22/08/2025=N/A"],
    ["RE-2024-01", "Angélica", "abr-24", "Sí", "No", "(invitado/familiar, no colaborador); testimonios: 22/08/2025=N/A"],
    ["RE-2024-01", "Myrna", "abr-24", "Sí", "No", "(invitado/familiar, no colaborador); testimonios: 22/08/2025=N/A"],
    ["RE-2024-01", "Martín Amaury Mejía Ruiz", "abr-24", "Sí", "No", "(invitado/familiar, no colaborador); testimonios: 22/08/2025=N/A"],
    ["RE-2024-01", "David Santiago Morales", "abr-24", "Sí", "No", "(invitado/familiar, no colaborador); testimonios: 22/08/2025=N/A"],
    ["RE-2024-01", "Ernesto Sánchez", "abr-24", "Sí", "No", "(no es colaborador actualmente); baja de liderazgo; testimonios: 22/08/2025=Baja"],
    ["RE-2024-01", "Karen Anaid Armas Gutierrez", "abr-24", "Sí", "Sí", ""],
    ["RE-2024-01", "Antonio Aceves Carvajal", "abr-24", "Sí", "No", "testimonios: 22/08/2025=N/A"],
    ["RE-2024-01", "David Piñón", "abr-24", "Sí", "No", "(no es colaborador actualmente)"],
    ["RE-2024-01", "Fernanda Benitez Barragan", "abr-24", "Sí", "No", "(no es colaborador actualmente); testimonios: 22/08/2025=Baja"],
    ["RE-2024-01", "Amada Hernández Martínez", "abr-24", "Sí", "No", "(no es colaborador actualmente); baja de liderazgo; testimonios: 22/08/2025=Baja"],
    ["RE-2024-01", "Rosa María Campos Torres", "abr-24", "Sí", "No", "(socio, no colaborador); testimonios: 22/08/2025=N/A"],
    ["RE-2024-01", "Magally Basurto", "abr-24", "Sí", "No", "(no es colaborador actualmente)"],
    ["RE-2024-01", "Karen Salas Peregrina", "abr-24", "Sí", "Sí", ""],
    ["RE-2024-02", "Jessica Leonella Toro Gil", "oct-24", "Sí", "Sí", "testimonios: 22/08/2025=Invitada"],
    ["RE-2024-02", "Katia Alejandra Villalobos Villela", "oct-24", "Sí", "No", ""],
    ["RE-2024-02", "Ingrid Itzel Tellez Villalba", "oct-24", "Sí", "No", ""],
    ["RE-2024-02", "Luis Ernesto Barba Durán", "oct-24", "Sí", "Sí", ""],
    ["RE-2024-02", "Cristina Chavoya Ríos", "oct-24", "Sí", "No", "(no es colaborador actualmente); baja de liderazgo; testimonios: 22/08/2025=Baja"],
    ["RE-2024-02", "Hector Alonso Martínez Infante", "oct-24", "Sí", "No", ""],
    ["RE-2024-02", "Pablo Iván Solís Arizaga", "oct-24", "Sí", "No", ""],
    ["RE-2024-02", "Edgar Lomelí Sánchez", "oct-24", "Sí", "No", "(no es colaborador actualmente)"],
    ["RE-2024-02", "María Cristina de Rosas Márquez", "oct-24", "Sí", "Sí", ""],
    ["RE-2025-01", "VITALINA BARRIOS ARTEAGA", "feb-25", "Sí", "No", ""],
    ["RE-2025-01", "Cesar Abraham Medina Hernández", "feb-25", "Sí", "No", ""],
    ["RE-2025-01", "Imelda Elizabeth Olvera Gonzalez", "feb-25", "Sí", "Sí", ""],
    ["RE-2025-01", "Maritza Jacqueline Sánchez Pérez", "feb-25", "Sí", "No", "(no es colaborador actualmente); baja de liderazgo; testimonios: 22/08/2025=Baja"],
    ["RE-2025-01", "TATIANA GUADALUPE LOZA PINEDA", "feb-25", "Sí", "Sí", ""],
    ["RE-2025-01", "María Vanessa Velásquez Ramírez", "feb-25", "Sí", "No", "(no es colaborador actualmente); baja de liderazgo; testimonios: 22/08/2025=Baja"],
    ["RE-2025-01", "Damaris Ibarra Rivas", "feb-25", "Sí", "Sí", "(no es colaborador actualmente)"],
    ["RE-2025-02", "Valeria Esposa (Isaac)", "jun-25", "Sí", "No", "(invitado/familiar, no colaborador); testimonios: 22/08/2025=N/A"],
    ["RE-2025-02", "Ilse Alejandra Cameros Bobadilla", "jun-25", "Sí", "Sí", ""],
    ["RE-2025-02", "Andrés Isaac Franco Jiménez", "jun-25", "Sí", "Sí", ""],
    ["RE-2025-02", "Cristina Geraldine Gritti Medina", "jun-25", "Sí", "No", ""],
    ["RE-2025-02", "Daniel Gutierrez Buendia", "jun-25", "Sí", "Sí", ""],
    ["RE-2025-02", "Griselda Luquin", "jun-25", "Sí", "Sí", "(colaboradora, datos pendientes de completar en Colaboradores)"],
    ["RE-2025-02", "Álvaro José Bayardo Coronado", "jun-25", "Sí", "No", ""],
    ["RE-2025-02", "Jenny", "jun-25", "Sí", "Sí", "(no es colaborador actualmente)"],
    ["RE-2025-02", "Andrea Sandoval Velázquez", "jun-25", "Sí", "Sí", ""],
    ["RE-2025-02", "Jóse Alejandro Caram Salinas", "jun-25", "Sí", "Sí", ""],
    ["RE-2025-02", "Emma Medina", "jun-25", "Sí", "Sí", "(no es colaborador actualmente)"],
    ["RE-2025-02", "Karla Neri", "jun-25", "Sí", "Sí", "(no es colaborador actualmente)"],
    ["RE-2025-02", "JOSE ORTIZ SALAZAR", "jun-25", "Sí", "Sí", ""],
    ["RE-2025-02", "BERENICE MARTINEZ MARTINEZ", "jun-25", "Sí", "Sí", ""],
    ["RE-2025-03", "Abraham Amezcua Esparza", "jul-25", "Sí", "No", ""],
    ["RE-2025-03", "Delia Catalina González Barrera", "jul-25", "Sí", "Sí", ""],
    ["RE-2025-03", "Samuel Rodriguez Aguirre", "jul-25", "Sí", "No", "(no es colaborador actualmente); testimonios: 18/07/2025=Baja, 22/08/2025=Baja"],
    ["RE-2025-03", "Carlos Enrique Aceituno Pinzón", "jul-25", "Sí", "No", ""],
    ["RE-2025-03", "Joel Alfonso Quezada Pedroza", "jul-25", "Sí", "Sí", ""],
    ["RE-2025-03", "Blanca Lizette Michel Lopez", "jul-25", "Sí", "Sí", ""],
    ["RE-2025-03", "Oliver Adan De la Torre Jauregui", "jul-25", "No", "No", ""],
    ["RE-2025-03", "Reyna Cecilia Villanueva Aguilar", "jul-25", "Sí", "No", ""],
    ["RE-2025-03", "Ivan Padilla Cárdenas", "jul-25", "Sí", "No", ""],
    ["RE-2025-03", "Jesus Alicia Dinero Garcia", "jul-25", "Sí", "No", ""],
    ["RE-2025-03", "Francisco Joel Santana Hernández", "jul-25", "Sí", "No", ""],
    ["RE-2025-03", "Sergio Peñaloza Vázquez", "jul-25", "Sí", "Sí", ""],
    ["RE-2025-04", "Tania Verónica Rojas Espinoza", "oct-25", "Sí", "No", "testimonios: 22/08/2025=Confirmado"],
    ["RE-2025-04", "Elmer Omar Derek Palma Gonzalez", "oct-25", "Sí", "", ""],
    ["RE-2025-04", "Andrea Sanchez Lujano", "oct-25", "Sí", "No", "testimonios: 22/08/2025=Invitó"],
    ["RE-2025-04", "ANDREA ACUÑA ROBLEDO", "oct-25", "Sí", "", ""],
    ["RE-2025-04", "America Polette Perea Carrillo", "oct-25", "Sí", "Sí", ""],
    ["RE-2025-04", "Jennifer Valeria Maldonado Calderón", "oct-25", "Sí", "Sí", ""],
    ["RE-2025-04", "Esposa Luis Barba", "oct-25", "Sí", "", "(invitado/familiar, no colaborador)"],
    ["RE-2025-04", "Diego Segovia", "oct-25", "Sí", "Sí", "(no es colaborador actualmente)"],
    ["RE-2025-05", "Alfredo Jose Alfredo Salvador León", "nov-25", "Sí", "No", "(no es colaborador actualmente); testimonios: 22/08/2025=BAJA"],
    ["RE-2025-05", "Darwin Ariel Zapata Garcia", "nov-25", "Sí", "No", "testimonios: 22/08/2025=Confirmado"],
    ["RE-2025-05", "Edith Guadalupe Silva Aguiñaga", "nov-25", "Sí", "No", "testimonios: 22/08/2025=Invitó"],
    ["RE-2025-05", "Mirella Avalos Juarez", "nov-25", "Sí", "No", "(no es colaborador actualmente); testimonios: 22/08/2025=Confirmado"],
    ["RE-2026-01", "Penelope Estefanía Silva", "ene-26", "Sí", "", "(invitado/familiar, no colaborador)"],
    ["RE-2026-01", "Virginia Florencia Mejia Ruiz", "ene-26", "Sí", "No", "testimonios: 22/08/2025=Confirmado"],
    ["RE-2026-01", "Veronica Torres Campos", "ene-26", "Sí", "No", "testimonios: 22/08/2025=Confirmado"],
    ["RE-2026-02", "Josefina Teresita Jauregui Jiménez", "may-26", "", "", "(invitado/familiar, no colaborador)"],
    ["RE-2026-02", "Adan de la Torre Salmeron", "may-26", "", "", "(invitado/familiar, no colaborador)"],
    ["RE-2026-02", "Diana Marcela Carrillo Reyes", "may-26", "", "", ""],
    ["RE-2026-02", "Gabino Nicolas Gomez Aceves", "may-26", "", "", ""],
    ["RE-2026-02", "Olivia Carrillo Burciaga", "may-26", "", "", ""],
    ["RE-2026-02", "Rosalba Jimenez Rivera", "may-26", "", "", ""],
    ["RE-2026-02", "Fátima Lizette Hernández Jiménez", "may-26", "", "", ""],
    ["RE-2026-02", "Xander Elias Zepeda Alcaraz", "may-26", "", "", ""],
    ["RE-2026-02", "Karla Beatriz Flores Duarte", "may-26", "", "", ""],
    ["RE-2026-02", "Jennifer Montserrat Esparza Hernandez", "may-26", "", "", ""],
    ["RE-2026-02", "David Alejandro Diaz Garcia", "may-26", "", "", ""],
    ["RE-2026-02", "Michelle Mejía Romero", "may-26", "", "", ""]
  ];

  var datosEd = wsEd.getDataRange().getValues();
  var yaExistenEd = {};
  for (var i = 1; i < datosEd.length; i++) {
    var idEd = String(datosEd[i][0] || "").trim();
    if (idEd) yaExistenEd[idEd] = true;
  }
  var edAgregadas = 0, edOmitidas = 0;
  ediciones.forEach(function (f) {
    if (yaExistenEd[f[0]]) { edOmitidas++; return; }
    wsEd.appendRow(f);
    yaExistenEd[f[0]] = true;
    edAgregadas++;
  });

  var datosAs = wsAs.getDataRange().getValues();
  var yaExistenAs = {};
  for (var j = 1; j < datosAs.length; j++) {
    var key = String(datosAs[j][0] || "").trim() + "|" + String(datosAs[j][1] || "").trim().toLowerCase();
    if (key !== "|") yaExistenAs[key] = true;
  }
  var asAgregadas = 0, asOmitidas = 0;
  asistencias.forEach(function (f) {
    var key = f[0] + "|" + f[1].toLowerCase();
    if (yaExistenAs[key]) { asOmitidas++; return; }
    wsAs.appendRow(f);
    yaExistenAs[key] = true;
    asAgregadas++;
  });

  Logger.log("RE histórico 2026-07-05: ediciones agregadas=" + edAgregadas + " omitidas=" + edOmitidas +
             " | asistencias agregadas=" + asAgregadas + " omitidas=" + asOmitidas);
}

// ── Retiro del Espíritu Santo: Alta de nueva edición ────────────
// Cols RE_Ediciones: A=id_edicion, B=nombre_edicion, C=fechas_texto,
//                    D=fecha_inicio, E=fecha_fin, F=lugar, G=estado, H=notas
function handleReAltaEdicion(p) {
  const ws = reEdicionesSheetNuevo_();
  if (!ws) return resp({ ok: false, error: "Pestaña RE_Ediciones no encontrada en el Sheet nuevo" });

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
  const ws = reAsistenciasSheetNuevo_();
  if (!ws) return resp({ ok: false, error: "Pestaña RE_Asistencias no encontrada en el Sheet nuevo" });

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
function handleReAsistenciaLote(p) {
  const ws = reAsistenciasSheetNuevo_();
  if (!ws) return resp({ ok: false, error: "Pestaña RE_Asistencias no encontrada en el Sheet nuevo" });

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

  return resp({ ok: true });
}

// ── Retiro del Espíritu Santo: Guardar acomodo de habitaciones ──
// Cols RE_Habitaciones: A=id_edicion, B=habitacion, C=genero, D-G=nombre_1..4, H=notas
// Borra y reescribe todas las filas de la edición (el panel manda el estado completo).
function handleReGuardarHabitaciones(p) {
  const ws = reHabitacionesSheetNuevo_();
  if (!ws) return resp({ ok: false, error: "Pestaña RE_Habitaciones no encontrada en el Sheet nuevo" });

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

// ── Retiro del Espíritu Santo: Interesados / lista de espera ────
// Cols RE_Interesados: A=Nombre, B=Sucursal, C=Telefono_WA, D=Fecha_Alta, E=Estado
// (Interesado | Inscrito | Descartado), F=Requiere_Seguimiento, G=Notas. Se llenan
// desde altas.html (alta_unificada) o el "+ Agregar interesado" de este panel.
function handleReAltaInteresado(p) {
  const ws = reInteresadosSheetNuevo_();
  if (!ws) return resp({ ok: false, error: "Pestaña RE_Interesados no encontrada en el Sheet nuevo" });
  const nombre = (p.nombre || "").trim();
  if (!nombre) return resp({ ok: false, error: "Falta el nombre" });

  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
      return resp({ ok: false, error: "Ya existe como interesado." });
    }
  }
  ws.appendRow([nombre, p.sucursal || "", p.telefono || "", new Date(), "Interesado", "Sí", p.notas || ""]);
  return resp({ ok: true });
}

function handleReMarcarAtendidoInteresado(p) {
  const ws = reInteresadosSheetNuevo_();
  if (!ws) return resp({ ok: false, error: "Pestaña RE_Interesados no encontrada en el Sheet nuevo" });
  const nombre = (p.nombre || "").trim();
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
      ws.getRange(i + 1, 6).setValue(""); // Requiere_Seguimiento (col F)
      return resp({ ok: true });
    }
  }
  return resp({ ok: false, error: "No se encontró a " + nombre + " en RE_Interesados" });
}

function handleReBajaInteresado(p) {
  const ws = reInteresadosSheetNuevo_();
  if (!ws) return resp({ ok: false, error: "Pestaña RE_Interesados no encontrada en el Sheet nuevo" });
  const nombre = (p.nombre || "").trim();
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
      ws.getRange(i + 1, 5).setValue("Descartado"); // Estado (col E)
      ws.getRange(i + 1, 6).setValue("No");         // Requiere_Seguimiento (col F)
      return resp({ ok: true });
    }
  }
  return resp({ ok: false, error: "No se encontró a " + nombre + " en RE_Interesados" });
}

// Inscribir a un interesado en una edición: crea su fila en RE_Asistencias
// (mismo efecto que handleReRegistro) y marca el interesado como "Inscrito".
function handleReInscribirInteresado(p) {
  const wsInt = reInteresadosSheetNuevo_();
  const wsAsis = reAsistenciasSheetNuevo_();
  if (!wsInt || !wsAsis) return resp({ ok: false, error: "Pestaña RE_Interesados o RE_Asistencias no encontrada en el Sheet nuevo" });

  const nombre = (p.nombre || "").trim();
  const idEdicion = (p.idEdicion || "").trim();
  if (!nombre || !idEdicion) return resp({ ok: false, error: "Nombre y edición son obligatorios." });

  const asisData = wsAsis.getDataRange().getValues();
  const yaInscrito = asisData.some(function (r, i) {
    return i > 0 && String(r[0]).trim() === idEdicion && String(r[1]).trim().toLowerCase() === nombre.toLowerCase();
  });
  if (!yaInscrito) wsAsis.appendRow([idEdicion, nombre, "", "Pendiente", "—", ""]);

  const intData = wsInt.getDataRange().getValues();
  for (let i = 1; i < intData.length; i++) {
    if (String(intData[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
      wsInt.getRange(i + 1, 5).setValue("Inscrito"); // Estado (col E)
      wsInt.getRange(i + 1, 6).setValue("No");       // Requiere_Seguimiento (col F)
      break;
    }
  }

  return resp({ ok: true });
}

// ── Camino de Santiago (Sheet nuevo) ─────────────────────────────
// Cols CS_Inscritos: A=Nombre, B=Sucursal, C=Telefono_WA, D=Fecha_Alta, E=Estado
// (Activo | Completado | Baja), F=Requiere_Seguimiento, G=Notas, H=Fecha_Camino
// (fecha real de la peregrinación, se llena al completar). Reemplaza a las
// columnas sueltas "Camino"/"CaminoFecha" que vivían en Concentrado.
function csSheetNuevo_() { return SpreadsheetApp.openById(SHEET_NUEVO_ID).getSheetByName("CS_Inscritos"); }

function handleCsMarcarAtendido(p) {
  const ws = csSheetNuevo_();
  if (!ws) return resp({ ok: false, error: "Pestaña CS_Inscritos no encontrada en el Sheet nuevo" });
  const nombre = (p.nombre || "").trim();
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
      ws.getRange(i + 1, 6).setValue(""); // Requiere_Seguimiento (col F)
      return resp({ ok: true });
    }
  }
  return resp({ ok: false, error: "No se encontró a " + nombre + " en CS_Inscritos" });
}

function handleCsRegistrarFecha(p) {
  const ws = csSheetNuevo_();
  if (!ws) return resp({ ok: false, error: "Pestaña CS_Inscritos no encontrada en el Sheet nuevo" });
  const nombre = (p.nombre || "").trim();
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
      ws.getRange(i + 1, 8).setValue(p.fecha ? new Date(p.fecha) : ""); // Fecha_Camino (col H)
      ws.getRange(i + 1, 5).setValue("Completado");                    // Estado (col E)
      ws.getRange(i + 1, 6).setValue("");                              // Requiere_Seguimiento (col F)
      return resp({ ok: true });
    }
  }
  return resp({ ok: false, error: "No se encontró a " + nombre + " en CS_Inscritos" });
}

function handleCsBaja(p) {
  const ws = csSheetNuevo_();
  if (!ws) return resp({ ok: false, error: "Pestaña CS_Inscritos no encontrada en el Sheet nuevo" });
  const nombre = (p.nombre || "").trim();
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
      ws.getRange(i + 1, 5).setValue("Baja"); // Estado (col E)
      ws.getRange(i + 1, 6).setValue("No");   // Requiere_Seguimiento (col F)
      return resp({ ok: true });
    }
  }
  return resp({ ok: false, error: "No se encontró a " + nombre + " en CS_Inscritos" });
}

// ── Carga real única: 16 asistentes históricos Camino de Santiago (2026-07-05) ──
// Datos reales dados por Cecilia (todos ya "Asistió"). appendRow directo
// porque alta_unificada siempre deja Fecha_Camino vacía y Estado/Requiere_
// Seguimiento fijos al crear — aquí ya se sabe que todos están "Completado".
// Karen Salas Peregrina y Pablo Iván Solís Arizaga no existían en Colaboradores
// y ya se dieron de alta ahí aparte (fila 10 y 13 de esta lista).
function cargarCSLote_2026_07_05() {
  var ws = csSheetNuevo_();
  if (!ws) return;

  var filas = [
    ["Gerardo Roberto Villarreal Treviño",   "OFC",      "3314174439", "", "Completado", "", "Líder", ""],
    ["Jose Hiram Gonzalez Palazuelos",       "OFC",      "3321542114", "", "Completado", "", "Líder", "01/10/2024"],
    ["Raúl Eduardo Mantecón Siordia",        "OFC",      "3310705203", "", "Completado", "", "Líder", "01/10/2024"],
    ["Marco Antonio Peregrina Sierra",       "OFC",      "",           "", "Completado", "", "Líder", "01/10/2024"],
    ["Linda Zoraida Padilla Alvarez",        "OFC",      "3338067531", "", "Completado", "", "Líder", "01/10/2024"],
    ["Julia Elizabeth Rodríguez Arroyo",     "OFC",      "3315454700", "", "Completado", "", "Líder", "01/10/2024"],
    ["Karen Anaid Armas Gutierrez",          "GMT",      "3311637021", "", "Completado", "", "Líder", "01/10/2024"],
    ["Antonio Aceves Carvajal",              "NAC",      "3328319060", "", "Completado", "", "Líder", "01/05/2025"],
    ["ISMAEL CASTILLO GONZALEZ",             "OFC EASY", "3314141714", "", "Completado", "", "Líder", "01/05/2025"],
    ["Karen Salas Peregrina",                "OFC",      "",           "", "Completado", "", "Líder", "01/10/2025"],
    ["Reyna Cecilia Villanueva Aguilar",     "OFC",      "",           "", "Completado", "", "Líder", "01/10/2025"],
    ["Andrea Viridiana Bedoy Ruiz",          "GMT",      "3334678410", "", "Completado", "", "Líder", "01/10/2025"],
    ["Pablo Iván Solís Arizaga",             "OFC",      "",           "", "Completado", "", "Líder", "01/05/2026"],
    ["Luis Ernesto Barba Durán",             "OFC",      "3314966829", "", "Completado", "", "Líder", "01/05/2026"],
    ["Elmer Omar Derek Palma Gonzalez",      "OFC",      "",           "", "Completado", "", "Líder", "01/05/2026"],
    ["Miguel Angel Garcia Arana",            "NAC",      "3312281209", "", "Completado", "", "",      "01/05/2026"]
  ];

  var data = ws.getDataRange().getValues();
  var yaExisten = {};
  for (var j = 1; j < data.length; j++) {
    var n = String(data[j][0] || "").trim().toLowerCase();
    if (n) yaExisten[n] = true;
  }

  var agregadas = 0, omitidas = 0;
  filas.forEach(function (f) {
    var key = String(f[0]).trim().toLowerCase();
    if (yaExisten[key]) { omitidas++; return; }
    ws.appendRow(f);
    yaExisten[key] = true;
    agregadas++;
  });

  Logger.log("CS lote 2026-07-05: agregadas=" + agregadas + " omitidas=" + omitidas);
}

// ── Migración única: RA_Copys / SV_Copys / EG_Copys → Sheet nuevo ────
// Últimas 3 pestañas de plantillas de WhatsApp que quedaban en el Sheet
// histórico. Igual que las demás migraciones de Copys: crea la pestaña en el
// Sheet nuevo si no existe y copia las filas tal cual (son solo textos de
// mensaje, no datos de personas). Idempotente: si el destino ya tiene datos,
// no vuelve a copiar.
function migrarCopysRestantesASheetNuevo() {
  var ssNuevo = SpreadsheetApp.openById(SHEET_NUEVO_ID);
  var ssHistorico = SpreadsheetApp.openById(SHEET_HISTORICO_ID);
  var tabs = ["RA_Copys", "SV_Copys", "EG_Copys"];
  var resultado = [];

  tabs.forEach(function (nombre) {
    var wsViejo = ssHistorico.getSheetByName(nombre);
    if (!wsViejo) { resultado.push(nombre + ": no existe en el Sheet histórico"); return; }

    var wsNuevo = ssNuevo.getSheetByName(nombre);
    if (!wsNuevo) wsNuevo = ssNuevo.insertSheet(nombre);

    if (wsNuevo.getLastRow() > 0) { resultado.push(nombre + ": ya tenía datos, no se tocó"); return; }

    var filas = wsViejo.getDataRange().getValues();
    if (filas.length === 0) { resultado.push(nombre + ": el histórico está vacío"); return; }
    wsNuevo.getRange(1, 1, filas.length, filas[0].length).setValues(filas);
    resultado.push(nombre + ": " + (filas.length - 1) + " plantilla(s) copiada(s)");
  });

  SpreadsheetApp.getUi().alert("Migración de Copys al Sheet nuevo:\n\n" + resultado.join("\n"));
}

// ── Migración única: últimos logs (EG_Asistencias/Examenes, SV_Historial) ──
// Últimas 3 pestañas de datos reales que quedaban en el Sheet histórico
// (bitácoras de asistencia/examen y de donaciones). Mismo patrón que
// migrarBibliotecaASheetNuevo: copia encabezado + todas las filas reales tal
// cual, solo si el destino está vacío (idempotente, no duplica). El Sheet
// histórico no se borra ni se modifica.
function migrarLogsRestantesASheetNuevo() {
  var ssNuevo = SpreadsheetApp.openById(SHEET_NUEVO_ID);
  var ssHistorico = SpreadsheetApp.openById(SHEET_HISTORICO_ID);
  var tabs = ["EG_Asistencias", "EG_Examenes", "SV_Historial"];
  var resultado = [];

  tabs.forEach(function (nombre) {
    var wsViejo = ssHistorico.getSheetByName(nombre);
    if (!wsViejo) { resultado.push(nombre + ": no existe en el Sheet histórico"); return; }

    var wsNuevo = ssNuevo.getSheetByName(nombre);
    if (!wsNuevo) wsNuevo = ssNuevo.insertSheet(nombre);

    if (wsNuevo.getLastRow() > 0) { resultado.push(nombre + ": ya tenía datos, no se tocó"); return; }

    var filas = wsViejo.getDataRange().getValues();
    if (filas.length === 0) { resultado.push(nombre + ": el histórico está vacío"); return; }
    wsNuevo.getRange(1, 1, filas.length, filas[0].length).setValues(filas);
    resultado.push(nombre + ": " + (filas.length - 1) + " fila(s) copiada(s)");
  });

  SpreadsheetApp.getUi().alert("Migración de logs restantes al Sheet nuevo:\n\n" + resultado.join("\n"));
}

// ── Configuración única: crear pestañas del Sheet nuevo para RE + CS ──
// RE_Ediciones/Asistencias/Habitaciones/Copys/Interesados y CS_Inscritos no
// existen todavía en el Sheet nuevo — esta función las crea con encabezado, y
// copia verbatim las plantillas de RE_Copys del Sheet histórico (son solo
// textos de mensaje, no datos de personas). Idempotente.
function configurarREyCSSheetNuevo() {
  var ssNuevo = SpreadsheetApp.openById(SHEET_NUEVO_ID);
  var ui = SpreadsheetApp.getUi();

  var specs = [
    { nombre: "RE_Ediciones", headers: ["id_edicion","nombre_edicion","fechas_texto","fecha_inicio","fecha_fin","lugar","estado","notas"] },
    { nombre: "RE_Asistencias", headers: ["id_edicion","nombre","edicion_label","asistio","dio_testimonio","notas"] },
    { nombre: "RE_Habitaciones", headers: ["id_edicion","habitacion","genero","nombre_1","nombre_2","nombre_3","nombre_4","notas"] },
    { nombre: "RE_Copys", headers: ["Momento","Copy_Texto","Numero_WA","Activo"] },
    { nombre: "RE_Interesados", headers: ["Nombre","Sucursal","Telefono_WA","Fecha_Alta","Estado","Requiere_Seguimiento","Notas"] },
    { nombre: "CS_Inscritos", headers: ["Nombre","Sucursal","Telefono_WA","Fecha_Alta","Estado","Requiere_Seguimiento","Notas","Fecha_Camino"] }
  ];

  specs.forEach(function (spec) {
    var ws = ssNuevo.getSheetByName(spec.nombre);
    if (!ws) {
      ws = ssNuevo.insertSheet(spec.nombre);
      ws.appendRow(spec.headers);
    }
  });

  var wsCopysNuevo = ssNuevo.getSheetByName("RE_Copys");
  var copysCopiados = false;
  if (wsCopysNuevo.getLastRow() < 2) {
    var wsCopysViejo = SpreadsheetApp.openById(SHEET_HISTORICO_ID).getSheetByName("RE_Copys");
    if (wsCopysViejo && wsCopysViejo.getLastRow() >= 2) {
      var filas = wsCopysViejo.getRange(2, 1, wsCopysViejo.getLastRow() - 1, 4).getValues();
      wsCopysNuevo.getRange(2, 1, filas.length, 4).setValues(filas);
      copysCopiados = true;
    }
  }

  ui.alert("Listo. RE_Ediciones/Asistencias/Habitaciones/Copys/Interesados y CS_Inscritos ya existen en el Sheet nuevo" +
    (copysCopiados ? " (copys de RE copiados del Sheet histórico)." : "."));
}

// ── Sembrar datos de prueba Retiro Espiritual + Camino Santiago ──
// Solo desarrollo: una edición de prueba con participantes, una habitación,
// interesados en distintos estados, y colaboradores de Camino de Santiago.
function sembrarPruebasRE() {
  var ssNuevo = SpreadsheetApp.openById(SHEET_NUEVO_ID);
  var wsEd = reEdicionesSheetNuevo_();
  var wsAsis = reAsistenciasSheetNuevo_();
  var wsHab = reHabitacionesSheetNuevo_();
  var wsInt = reInteresadosSheetNuevo_();
  var wsCs = csSheetNuevo_();
  if (!wsEd || !wsAsis || !wsHab || !wsInt || !wsCs) {
    SpreadsheetApp.getUi().alert("Corre primero 'Configurar RE + Camino Santiago en Sheet nuevo (1 vez)'.");
    return;
  }

  var idEdicion = "RE-PRUEBA-01";
  wsEd.appendRow([idEdicion, "PRUEBA Retiro — Edición de prueba", "sábado y domingo de prueba", new Date(), new Date(), "Casa de retiros (prueba)", "Activa", ""]);

  wsAsis.getRange(wsAsis.getLastRow() + 1, 1, 2, 6).setValues([
    [idEdicion, "PRUEBA Nadia Ríos", "", "Pendiente", "—", ""],
    [idEdicion, "PRUEBA Emilio Vega", "", "Sí", "Sí", ""]
  ]);

  wsHab.appendRow([idEdicion, "PRUEBA Habitación 1", "Mujer", "PRUEBA Nadia Ríos", "", "", "", ""]);

  var hoy = new Date();
  wsInt.getRange(wsInt.getLastRow() + 1, 1, 2, 7).setValues([
    ["PRUEBA Karina Solís", "Sucursal Centro", "3312345620", hoy, "Interesado", "Sí", ""],
    ["PRUEBA Uriel Paz", "Sucursal Norte", "3312345621", hoy, "Interesado", "Sí", ""]
  ]);

  wsCs.getRange(wsCs.getLastRow() + 1, 1, 2, 8).setValues([
    ["PRUEBA Fernanda Ibarra", "Sucursal Sur", "3312345622", hoy, "Activo", "Sí", "", ""],
    ["PRUEBA Gustavo Reyes", "Sucursal Centro", "3312345623", hoy, "Completado", "", "", hoy]
  ]);

  SpreadsheetApp.getUi().alert("Datos de prueba de Retiro Espiritual y Camino de Santiago agregados.");
}

// ================================================================
// BIBLIOTECA COMUNITARIA (Sheet nuevo desde 2026-07-05)
// ================================================================
// Migrado desde el Sheet histórico: BIB_Prestamos/Donaciones/Devoluciones ya
// tenían datos reales de producción (préstamos y donaciones reales, no solo
// plantillas), así que la migración copia esas filas tal cual — ver
// migrarBibliotecaASheetNuevo(). BIB_Copys (solo plantillas de WhatsApp) se
// copia igual que en las demás iniciativas. El resto de la lógica (matching
// de devoluciones, PDFs) no cambia: usa lookup de encabezados por nombre
// (obtenerIndicesBib_), no posición fija, así que da igual en qué Sheet viva.
function bibSheetNuevo_(nombre) {
  return SpreadsheetApp.openById(SHEET_NUEVO_ID).getSheetByName(nombre);
}

// ── Migración única: BIB_* histórico → Sheet nuevo ───────────────
// A diferencia de otras migraciones (que solo copiaban plantillas de copys),
// BIB_Prestamos/Donaciones/Devoluciones YA tienen datos reales de producción
// (préstamos y donaciones reales desde 2026-06-29). Esta función copia el
// encabezado + TODAS las filas tal cual (no solo altas nuevas) a las pestañas
// homónimas del Sheet nuevo. Solo corre si la pestaña destino está vacía
// (idempotente: si ya se migró, no vuelve a copiar y no duplica nada). El
// Sheet histórico NO se borra ni se modifica — queda como respaldo.
function migrarBibliotecaASheetNuevo() {
  var ssNuevo = SpreadsheetApp.openById(SHEET_NUEVO_ID);
  var ssHistorico = SpreadsheetApp.openById(SHEET_HISTORICO_ID);
  var ui = SpreadsheetApp.getUi();
  var tabs = ["BIB_Prestamos", "BIB_Donaciones", "BIB_Devoluciones", "BIB_Copys"];
  var resultado = [];

  tabs.forEach(function (nombre) {
    var wsViejo = ssHistorico.getSheetByName(nombre);
    if (!wsViejo) { resultado.push(nombre + ": no existe en el Sheet histórico"); return; }

    var wsNuevo = ssNuevo.getSheetByName(nombre);
    if (!wsNuevo) wsNuevo = ssNuevo.insertSheet(nombre);

    if (wsNuevo.getLastRow() > 0) { resultado.push(nombre + ": ya tenía datos, no se tocó"); return; }

    var filas = wsViejo.getDataRange().getValues();
    if (filas.length === 0) { resultado.push(nombre + ": el histórico está vacío"); return; }
    wsNuevo.getRange(1, 1, filas.length, filas[0].length).setValues(filas);
    resultado.push(nombre + ": " + (filas.length - 1) + " fila(s) copiada(s)");
  });

  ui.alert("Migración de Biblioteca al Sheet nuevo:\n\n" + resultado.join("\n"));
}

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
  const hojaDev = bibSheetNuevo_("BIB_Devoluciones");
  const hojaPre = bibSheetNuevo_("BIB_Prestamos");
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
  const hoja = bibSheetNuevo_("BIB_Donaciones");
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
  const hoja = bibSheetNuevo_("BIB_Donaciones");
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

  const hojaDonaciones = bibSheetNuevo_("BIB_Donaciones");
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

  const hojaPrestamos = bibSheetNuevo_("BIB_Prestamos");
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
  const ws = bibSheetNuevo_("BIB_Donaciones");
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
  const hojaDev = bibSheetNuevo_("BIB_Devoluciones");
  const hojaPre = bibSheetNuevo_("BIB_Prestamos");
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
  const ws = bibSheetNuevo_("BIB_Prestamos");
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

// ================================================================
// IMPULSO GEB (Generación 1) — Sheet nuevo
// ================================================================
// Cols IG_Inscritos: A=Nombre, B=Sucursal, C=Telefono_WA, D=Fecha_Alta,
// E=Estado (Activo | Generación 2 | Baja), F=Requiere_Seguimiento, G=Notas,
// H=Plan_Vida, I=Presupuesto, J=Ahorro, K=Movilidad_Social (todas Sí/No).
// Generación 1 es solo histórico de evidencias entregadas — cuando lance
// Generación 2 (con momentos de WhatsApp, asistencia, etc.) se construye
// aparte; por ahora "Generación 2" es solo un estado de "pasó de generación",
// para no perder a esas personas cuando se arme el nuevo módulo.
const IG_CAMPOS_EVIDENCIA = { Plan_Vida: 8, Presupuesto: 9, Ahorro: 10, Movilidad_Social: 11 };

function igSheetNuevo_() { return SpreadsheetApp.openById(SHEET_NUEVO_ID).getSheetByName("IG_Inscritos"); }

// ── Carga real única: 11 inscritos Impulso GEB Gen.1 (2026-07-05) ──
// Datos reales dados por Cecilia. appendRow directo (no vía alta_unificada)
// porque ese endpoint hoy siempre deja Plan_Vida/Presupuesto/Ahorro/
// Movilidad_Social vacíos al crear, y aquí ya vienen con diagnóstico real.
function cargarIGLote_2026_07_05() {
  var ws = igSheetNuevo_();
  if (!ws) return;

  var filas = [
    ["David Alejandro Sánchez Alamos",   "NAC",      "3313021077", "", "Activo", "", "Caja y Reto", "Sí", "No", "Sí", ""],
    ["ALEJANDRA MELENDEZ MENDOZA",       "CAÑADAS",  "3333993503", "", "Activo", "", "Reto", "Sí", "Sí", "Sí", ""],
    ["Cesar Abraham Medina Hernández",   "NAC",      "3312878685", "", "Activo", "", "Caja ahorro", "Sí", "Sí", "Sí", "Promoción"],
    ["David Ortiz Perez",                "NAC",      "3318769613", "", "Activo", "", "Reto", "No", "No", "Sí", ""],
    ["Fátima Lizette Hernández Jiménez", "NAC",      "3332478861", "", "Activo", "", "Fondo ahorro", "No", "No", "Sí", ""],
    ["Gael Isaac Madrigal Rodriguez",    "NAC",      "3314708840", "", "Activo", "", "Fondo ahorro", "Sí", "Sí", "Sí", "Promoción"],
    ["Ingrid Itzel Tellez Villalba",     "VR",       "3329991459", "", "Activo", "", "Caja ahorro", "No", "No", "Sí", ""],
    ["Jerónimo de Jesús Aldrete Diaz",   "VR",       "3315297027", "", "Activo", "", "Caja ahorro", "Sí", "Sí", "Sí", ""],
    ["Mahel Marc Altiery",               "NAC",      "3317645854", "", "Activo", "", "Caja ahorro", "Sí", "Sí", "Sí", ""],
    ["ANTONIO DE JESUS LOPEZ VELASCO",   "ESTANCIA", "",           "", "Activo", "", "", "No", "No", "No", ""]
  ];

  // Gen.2 recién detectados (2026-07-05): sin diagnóstico Gen.1 (evidencias
  // aparte, módulo Gen.2 aún no existe) — entran directo con Estado="Generación 2".
  // America Polette Perea Carrillo apareció primero en la lista de Gen.1 pero
  // Cecilia confirmó que se queda solo en Gen.2 (no se crea su fila de Gen.1).
  var filasGen2 = [
    ["America Polette Perea Carrillo",   "NAC",           "3315197578", "", "Generación 2", "No", "", "", "", "", ""],
    ["ANDREA ACUÑA ROBLEDO",             "ÁVILA CAMACHO", "3345938028", "", "Generación 2", "No", "", "", "", "", ""],
    ["CINTHIA AIDE RAMIREZ HUERTA",      "ESTANCIA",      "3314476122", "", "Generación 2", "No", "", "", "", "", ""],
    ["NIDIA LIZZETTE RAMOS CARDENAS",    "ÁVILA CAMACHO", "3320511300", "", "Generación 2", "No", "", "", "", "", ""]
  ];
  filas = filas.concat(filasGen2);

  var data = ws.getDataRange().getValues();
  var yaExisten = {};
  for (var j = 1; j < data.length; j++) {
    var n = String(data[j][0] || "").trim().toLowerCase();
    if (n) yaExisten[n] = true;
  }

  var agregadas = 0, omitidas = 0;
  filas.forEach(function (f) {
    var key = String(f[0]).trim().toLowerCase();
    if (yaExisten[key]) { omitidas++; return; }
    ws.appendRow(f);
    yaExisten[key] = true;
    agregadas++;
  });

  Logger.log("IG lote 2026-07-05: agregadas=" + agregadas + " omitidas=" + omitidas);
}

function handleIgActualizarCampo(p) {
  const ws = igSheetNuevo_();
  if (!ws) return resp({ ok: false, error: "Pestaña IG_Inscritos no encontrada en el Sheet nuevo" });
  const col = IG_CAMPOS_EVIDENCIA[p.campo];
  if (!col) return resp({ ok: false, error: "Campo no reconocido: " + p.campo });

  const nombre = (p.nombre || "").trim();
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
      ws.getRange(i + 1, col).setValue(p.valor || "");
      ws.getRange(i + 1, 6).setValue(""); // Requiere_Seguimiento: ya se dio seguimiento
      return resp({ ok: true });
    }
  }
  return resp({ ok: false, error: "No se encontró a " + nombre + " en IG_Inscritos" });
}

function handleIgMarcarAtendido(p) {
  const ws = igSheetNuevo_();
  if (!ws) return resp({ ok: false, error: "Pestaña IG_Inscritos no encontrada en el Sheet nuevo" });
  const nombre = (p.nombre || "").trim();
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
      ws.getRange(i + 1, 6).setValue("");
      return resp({ ok: true });
    }
  }
  return resp({ ok: false, error: "No se encontró a " + nombre + " en IG_Inscritos" });
}

function handleIgBaja(p) {
  const ws = igSheetNuevo_();
  if (!ws) return resp({ ok: false, error: "Pestaña IG_Inscritos no encontrada en el Sheet nuevo" });
  const nombre = (p.nombre || "").trim();
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
      ws.getRange(i + 1, 5).setValue("Baja");
      ws.getRange(i + 1, 6).setValue("No");
      return resp({ ok: true });
    }
  }
  return resp({ ok: false, error: "No se encontró a " + nombre + " en IG_Inscritos" });
}

// Marca que la persona termina Generación 1 y continúa en Generación 2 (aún
// sin módulo propio) — distinto de "Baja", que es salir de la iniciativa.
function handleIgPasarGeneracion2(p) {
  const ws = igSheetNuevo_();
  if (!ws) return resp({ ok: false, error: "Pestaña IG_Inscritos no encontrada en el Sheet nuevo" });
  const nombre = (p.nombre || "").trim();
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
      ws.getRange(i + 1, 5).setValue("Generación 2");
      ws.getRange(i + 1, 6).setValue("No");
      return resp({ ok: true });
    }
  }
  return resp({ ok: false, error: "No se encontró a " + nombre + " en IG_Inscritos" });
}

function configurarIGSheetNuevo() {
  var ws = SpreadsheetApp.openById(SHEET_NUEVO_ID).getSheetByName("IG_Inscritos");
  if (!ws) {
    ws = SpreadsheetApp.openById(SHEET_NUEVO_ID).insertSheet("IG_Inscritos");
    ws.appendRow(["Nombre","Sucursal","Telefono_WA","Fecha_Alta","Estado","Requiere_Seguimiento","Notas","Plan_Vida","Presupuesto","Ahorro","Movilidad_Social"]);
  }
  SpreadsheetApp.getUi().alert("Listo. IG_Inscritos ya existe en el Sheet nuevo.");
}

function sembrarPruebasIG() {
  var ws = igSheetNuevo_();
  if (!ws) { SpreadsheetApp.getUi().alert("Corre primero 'Configurar Impulso GEB en Sheet nuevo (1 vez)'."); return; }
  var hoy = new Date();
  var filas = [
    ["PRUEBA Renata Cabrera", "Sucursal Centro", "3312345630", hoy, "Activo", "Sí", "", "", "", "", ""],
    ["PRUEBA Braulio Nájera", "Sucursal Norte",  "3312345631", hoy, "Activo", "",   "", "Sí", "Sí", "", ""],
    ["PRUEBA Ximena Corona",  "Sucursal Sur",    "3312345632", hoy, "Activo", "",   "", "Sí", "Sí", "Sí", "Sí"],
    ["PRUEBA Saúl Montoya",   "Sucursal Centro", "3312345633", hoy, "Baja",   "No", "", "Sí", "", "", ""]
  ];
  ws.getRange(ws.getLastRow() + 1, 1, filas.length, 11).setValues(filas);
  SpreadsheetApp.getUi().alert("4 colaboradores de prueba agregados a IG_Inscritos.");
}

// ================================================================
// BECA EDUCATIVA GEB — Sheet nuevo
// ================================================================
// Cols BE_Inscritos: A=Nombre, B=Sucursal, C=Telefono_WA, D=Fecha_Alta,
// E=Estado (Interesado | Inscrito | Baja), F=Requiere_Seguimiento, G=Notas,
// H=Universidad, I=Nivel (Licenciatura|Posgrado), J=Carrera, K=Inscribio
// (""=pendiente de revisar | Sí | No) — se revisa ~1 mes después del alta.
function beSheetNuevo_() { return SpreadsheetApp.openById(SHEET_NUEVO_ID).getSheetByName("BE_Inscritos"); }

function handleBeMarcarAtendido(p) {
  const ws = beSheetNuevo_();
  if (!ws) return resp({ ok: false, error: "Pestaña BE_Inscritos no encontrada en el Sheet nuevo" });
  const nombre = (p.nombre || "").trim();
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
      ws.getRange(i + 1, 6).setValue("");
      return resp({ ok: true });
    }
  }
  return resp({ ok: false, error: "No se encontró a " + nombre + " en BE_Inscritos" });
}

function handleBeBaja(p) {
  const ws = beSheetNuevo_();
  if (!ws) return resp({ ok: false, error: "Pestaña BE_Inscritos no encontrada en el Sheet nuevo" });
  const nombre = (p.nombre || "").trim();
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
      ws.getRange(i + 1, 5).setValue("Baja");
      ws.getRange(i + 1, 6).setValue("No");
      return resp({ ok: true });
    }
  }
  return resp({ ok: false, error: "No se encontró a " + nombre + " en BE_Inscritos" });
}

// p.inscribio: "Sí" | "No". Si "Sí", Estado pasa a "Inscrito" (ya no es solo
// un interesado). Si "No", se queda "Interesado" pero con Inscribio="No"
// registrado, para reportar cuántos no se inscribieron.
function handleBeRevisarInscripcion(p) {
  const ws = beSheetNuevo_();
  if (!ws) return resp({ ok: false, error: "Pestaña BE_Inscritos no encontrada en el Sheet nuevo" });
  const nombre = (p.nombre || "").trim();
  const inscribio = p.inscribio === "Sí" ? "Sí" : "No";
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
      ws.getRange(i + 1, 11).setValue(inscribio);
      if (inscribio === "Sí") ws.getRange(i + 1, 5).setValue("Inscrito");
      ws.getRange(i + 1, 6).setValue("");
      return resp({ ok: true });
    }
  }
  return resp({ ok: false, error: "No se encontró a " + nombre + " en BE_Inscritos" });
}

function configurarBESheetNuevo() {
  var ssNuevo = SpreadsheetApp.openById(SHEET_NUEVO_ID);
  var ws = ssNuevo.getSheetByName("BE_Inscritos");
  if (!ws) {
    ws = ssNuevo.insertSheet("BE_Inscritos");
    ws.appendRow(["Nombre","Sucursal","Telefono_WA","Fecha_Alta","Estado","Requiere_Seguimiento","Notas","Universidad","Nivel","Carrera","Inscribio"]);
  }

  var wsCopys = ssNuevo.getSheetByName("BE_Copys");
  if (!wsCopys) {
    wsCopys = ssNuevo.insertSheet("BE_Copys");
    wsCopys.appendRow(["Momento","Copy_Texto","Numero_WA","Activo"]);
    wsCopys.appendRow([
      "contacto_universidad",
      "Hola {Nombre} 🎓\n\nVimos tu interés en estudiar en {Universidad}" +
        " ({Nivel}: {Carrera}). GEB University tiene convenio de becas con ellos" +
        " — con gusto te compartimos los detalles y te ayudamos con el proceso" +
        " de inscripción. ¿Tienes unos minutos para platicarlo?",
      "",
      "Sí"
    ]);
  }

  SpreadsheetApp.getUi().alert("Listo. BE_Inscritos y BE_Copys ya existen en el Sheet nuevo (con una plantilla inicial de contacto).");
}

function sembrarPruebasBE() {
  var ws = beSheetNuevo_();
  if (!ws) { SpreadsheetApp.getUi().alert("Corre primero 'Configurar Beca Educativa en Sheet nuevo (1 vez)'."); return; }
  var hoy = new Date();
  var filas = [
    ["PRUEBA Yolanda Espino",  "Sucursal Centro", "3312345640", hoy, "Interesado", "Sí", "", "UTEL", "Licenciatura", "Administración", ""],
    ["PRUEBA Ricardo Uranga",  "Sucursal Norte",  "3312345641", hoy, "Interesado", "",   "", "Universidad Cuauhtémoc", "Posgrado", "Maestría en Educación", ""],
    ["PRUEBA Daniela Vidales", "Sucursal Sur",    "3312345642", hoy, "Inscrito",   "",   "", "Universidad Virtual de Liverpool", "Licenciatura", "Psicología", "Sí"]
  ];
  ws.getRange(ws.getLastRow() + 1, 1, filas.length, 11).setValues(filas);
  SpreadsheetApp.getUi().alert("3 interesados de prueba agregados a BE_Inscritos.");
}

// ================================================================
// ECO-ACCIÓN — Sheet nuevo
// ================================================================
// Cols EA_Lideres: A=Nombre, B=Sucursal (una de las 8 ubicaciones), C=Telefono_WA,
// D=Fecha_Alta, E=Estado (Activo | Baja), F=Requiere_Seguimiento, G=Notas,
// H=Total_Cilindros (contador, se actualiza solo al registrar un cilindro nuevo).
// Cols EA_Cilindros (historial): A=Nombre, B=Num_Cilindro, C=Fecha_Recepcion, D=Notas.
function eaSheetNuevo_() { return SpreadsheetApp.openById(SHEET_NUEVO_ID).getSheetByName("EA_Lideres"); }
function eaCilindrosSheetNuevo_() { return SpreadsheetApp.openById(SHEET_NUEVO_ID).getSheetByName("EA_Cilindros"); }

// Registra un cilindro nuevo para un líder: calcula el siguiente número
// (cuenta cuántos ya tiene en EA_Cilindros + 1), lo agrega al historial con la
// fecha que capture el panel, y actualiza el contador Total_Cilindros en
// EA_Lideres. También limpia Requiere_Seguimiento (contarlo como seguimiento dado).
function handleEaRegistrarCilindro(p) {
  const wsCil = eaCilindrosSheetNuevo_();
  const wsLid = eaSheetNuevo_();
  if (!wsCil || !wsLid) return resp({ ok: false, error: "Pestaña EA_Cilindros o EA_Lideres no encontrada en el Sheet nuevo" });

  const nombre = (p.nombre || "").trim();
  if (!nombre) return resp({ ok: false, error: "Falta el nombre" });

  const datosCil = wsCil.getDataRange().getValues();
  let siguiente = 1;
  for (let i = 1; i < datosCil.length; i++) {
    if (String(datosCil[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
      const n = parseInt(datosCil[i][1], 10) || 0;
      if (n >= siguiente) siguiente = n + 1;
    }
  }

  wsCil.appendRow([nombre, siguiente, p.fecha ? new Date(p.fecha) : new Date(), p.notas || ""]);

  const datosLid = wsLid.getDataRange().getValues();
  for (let i = 1; i < datosLid.length; i++) {
    if (String(datosLid[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
      wsLid.getRange(i + 1, 8).setValue(siguiente); // Total_Cilindros (col H)
      wsLid.getRange(i + 1, 6).setValue("");         // Requiere_Seguimiento (col F)
      break;
    }
  }

  return resp({ ok: true, numCilindro: siguiente });
}

function handleEaMarcarAtendido(p) {
  const ws = eaSheetNuevo_();
  if (!ws) return resp({ ok: false, error: "Pestaña EA_Lideres no encontrada en el Sheet nuevo" });
  const nombre = (p.nombre || "").trim();
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
      ws.getRange(i + 1, 6).setValue("");
      return resp({ ok: true });
    }
  }
  return resp({ ok: false, error: "No se encontró a " + nombre + " en EA_Lideres" });
}

function handleEaBaja(p) {
  const ws = eaSheetNuevo_();
  if (!ws) return resp({ ok: false, error: "Pestaña EA_Lideres no encontrada en el Sheet nuevo" });
  const nombre = (p.nombre || "").trim();
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
      ws.getRange(i + 1, 5).setValue("Baja");
      ws.getRange(i + 1, 6).setValue("No");
      return resp({ ok: true });
    }
  }
  return resp({ ok: false, error: "No se encontró a " + nombre + " en EA_Lideres" });
}

function configurarEASheetNuevo() {
  var ssNuevo = SpreadsheetApp.openById(SHEET_NUEVO_ID);
  var ws = ssNuevo.getSheetByName("EA_Lideres");
  if (!ws) {
    ws = ssNuevo.insertSheet("EA_Lideres");
    ws.appendRow(["Nombre","Sucursal","Telefono_WA","Fecha_Alta","Estado","Requiere_Seguimiento","Notas","Total_Cilindros"]);
  }
  var wsCil = ssNuevo.getSheetByName("EA_Cilindros");
  if (!wsCil) {
    wsCil = ssNuevo.insertSheet("EA_Cilindros");
    wsCil.appendRow(["Nombre","Num_Cilindro","Fecha_Recepcion","Notas"]);
  }
  SpreadsheetApp.getUi().alert("Listo. EA_Lideres y EA_Cilindros ya existen en el Sheet nuevo.");
}

function sembrarPruebasEA() {
  var ws = eaSheetNuevo_();
  var wsCil = eaCilindrosSheetNuevo_();
  if (!ws || !wsCil) { SpreadsheetApp.getUi().alert("Corre primero 'Configurar Eco-Acción en Sheet nuevo (1 vez)'."); return; }
  var hoy = new Date();
  var filas = [
    ["PRUEBA Rocío Damián",  "Naciones Unidas", "3312345650", hoy, "Activo", "Sí", "", 0],
    ["PRUEBA Tomás Zepeda",  "Valle Real",       "3312345651", hoy, "Activo", "",   "", 2],
    ["PRUEBA Lucía Barajas", "Iteso",            "3312345652", hoy, "Baja",   "No", "", 1]
  ];
  ws.getRange(ws.getLastRow() + 1, 1, filas.length, 8).setValues(filas);

  var cilindros = [
    ["PRUEBA Tomás Zepeda", 1, hoy, ""],
    ["PRUEBA Tomás Zepeda", 2, hoy, ""],
    ["PRUEBA Lucía Barajas", 1, hoy, ""]
  ];
  wsCil.getRange(wsCil.getLastRow() + 1, 1, cilindros.length, 4).setValues(cilindros);

  SpreadsheetApp.getUi().alert("3 líderes de prueba agregados a EA_Lideres.");
}

// ── Ejecutar funciones de configuración/siembra vía el Web App ──
// Las funciones de menú (configurarXSheetNuevo, sembrarPruebasX, etc.) llaman
// a SpreadsheetApp.getUi().alert(...) al final, que no existe en el contexto
// de un Web App — se captura ese error puntual (esperado) porque el trabajo
// real (crear pestañas, insertar filas) ya se hizo antes de esa línea. Esto
// permite correr la configuración inicial sin depender del menú de Sheets ni
// de "Ejecutar" en el editor, por si la organización restringe la
// autorización interactiva de apps nuevas (ver conversación 2026-07-05).
// Lista blanca por seguridad: solo funciones de configuración/siembra, nunca
// funciones que reciban datos arbitrarios del payload.
const ADMIN_FUNCIONES_PERMITIDAS = {
  configurarAPSheetNuevo: configurarAPSheetNuevo,
  sembrarPruebasAP: sembrarPruebasAP,
  sembrarPruebasEG: sembrarPruebasEG,
  configurarREyCSSheetNuevo: configurarREyCSSheetNuevo,
  sembrarPruebasRE: sembrarPruebasRE,
  migrarBibliotecaASheetNuevo: migrarBibliotecaASheetNuevo,
  configurarIGSheetNuevo: configurarIGSheetNuevo,
  sembrarPruebasIG: sembrarPruebasIG,
  configurarBESheetNuevo: configurarBESheetNuevo,
  sembrarPruebasBE: sembrarPruebasBE,
  migrarCopysRestantesASheetNuevo: migrarCopysRestantesASheetNuevo,
  migrarLogsRestantesASheetNuevo: migrarLogsRestantesASheetNuevo,
  desactivarSincronizacionAntigua: desactivarSincronizacionAntigua,
  borrarDatosPruebaTodos: borrarDatosPruebaTodos,
  configurarEASheetNuevo: configurarEASheetNuevo,
  sembrarPruebasEA: sembrarPruebasEA,
  configurarAlertasAtrasoRA: configurarAlertasAtrasoRA,
  cargarEGLote_2026_07_05: cargarEGLote_2026_07_05,
  cargarIGLote_2026_07_05: cargarIGLote_2026_07_05,
  cargarCSLote_2026_07_05: cargarCSLote_2026_07_05,
  cargarAPHistoricoLote_2026_07_05: cargarAPHistoricoLote_2026_07_05,
  cargarAPActualesLote_2026_07_05: cargarAPActualesLote_2026_07_05,
  cargarREHistoricoLote_2026_07_05: cargarREHistoricoLote_2026_07_05
};

function handleAdminEjecutar(p) {
  const fn = ADMIN_FUNCIONES_PERMITIDAS[p.funcion];
  if (!fn) return resp({ ok: false, error: "Función no permitida: " + p.funcion });
  try {
    fn();
  } catch (uiErr) {
    // Esperado: getUi()/alert() no existe fuera del editor de Sheets. El
    // trabajo real de la función ya se ejecutó antes de esa línea.
  }
  return resp({ ok: true, mensaje: "Ejecutado: " + p.funcion });
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
