// ================================================================
// GEB Iniciativas Humanas — Web App
// Despliega como: Ejecutar como "Yo", Acceso "Cualquier persona"
// ================================================================
// Desconectado por completo del Sheet histórico (2026-07-05): el Sheet nuevo
// es la única fuente de datos operativos para las 9 iniciativas. Todas las
// funciones de migración de una sola vez (Sheet histórico → Sheet nuevo) ya
// se usaron y se borraron de este archivo (2026-07-10) para no arriesgar
// volver a correrlas por error — ver historial de git si hace falta
// consultar cómo se hizo alguna migración en particular.

// Sheet nuevo unificado — única fuente de datos de las 9 iniciativas.
const SHEET_NUEVO_ID = "1oaWADtW9SmqxuOoOM1bf65Ht_NkpE4l_o_DMutN_IbI";
const SS = SpreadsheetApp.openById(SHEET_NUEVO_ID);

// Biblioteca Virtual externa (solo lectura/escritura de la pestaña "Fisicos")
const SHEET_BIBLIOTECA_ID = "1FDZB3aR-YAyVMsiAjo92PuUdH0iTfBmreCWv5DvCpdM";

// Hojas de las 7 iniciativas restringidas por baja de empresa (Impulso,
// Convenios/Beca Educativa, Eco-Acción, Atención Psicológica, Retiros +
// Camino de Santiago, Escuela, Reto Ahorro) — todas comparten el mismo layout
// base A=Nombre,B=Sucursal,C=Telefono_WA,D=Fecha_Alta,E=Estado,
// F=Requiere_Seguimiento,G=Notas. Salvando Vidas y Biblioteca NO están aquí:
// la persona puede seguir participando ahí aunque ya no sea colaborador.
const HOJAS_RESTRINGIDAS_BAJA_EMPRESA = [
  "AP_Inscritos", "IG_Inscritos", "BE_Inscritos", "EA_Lideres",
  "EG_Inscritos", "RA_Inscritos", "RE_Interesados", "CS_Inscritos"
];

// Busca `nombre` en la columna A de `sheetName` (Sheet nuevo) y, si la
// encuentra, pone Estado (col E) y Requiere_Seguimiento (col F). Reutilizado
// por los handlers "Reactivar" y por la cascada de baja de empresa — mismo
// patrón que ya usaban handleSvBaja/handleEgBaja/etc. antes de este cambio.
function _marcarEstadoPorNombre(sheetName, nombre, estado, seguimiento) {
  var ws = SS.getSheetByName(sheetName);
  if (!ws) return false;
  var data = ws.getDataRange().getValues();
  var nombreLower = nombre.trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombreLower) {
      ws.getRange(i + 1, 5).setValue(estado);
      ws.getRange(i + 1, 6).setValue(seguimiento);
      return true;
    }
  }
  return false;
}

const LOGO_BIBLIO_ID = "1NqoFmESlsTP4dpFscYglcs9o6THQP4TR";
const LOGO_UGEB_ID = "1mBIHoKyngoa7cBiSvHCVY0x_pIkFyARJ";
const COLOR_PRIMARIO = "#185FA5";

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("GEB CRM")
    .addItem("🔧 Configurar alertas de atraso Reto Ahorro (1 vez)", "configurarAlertasAtrasoRA")
    .addItem("🗑️ Borrar TODOS los datos de prueba (PRUEBA...)", "borrarDatosPruebaTodos")
    .addItem("🧹 Quitar emojis de todos los Copys (1 vez)", "limpiarEmojisCopys")
    .addToUi();
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

// ── Quitar emojis de todos los Copys (desarrollo) ────────────────
// Los emojis en Copy_Texto no siempre llegan bien por WhatsApp. Recorre
// todas las pestañas *_Copys del Sheet nuevo, busca la columna Copy_Texto
// por encabezado (no por posición fija, ya que varía entre pestañas) y
// quita cualquier emoji del texto. Solo reescribe las celdas que
// realmente cambiaron. Correr una sola vez; los copys nuevos que se
// capturen después (a mano en el Sheet, o desde el editor de Reto Ahorro
// en el portal) deben escribirse ya sin emojis.
function limpiarEmojisCopys() {
  var tabs = ["AP_Copys", "RA_Copys", "SV_Copys", "BE_Copys", "EG_Copys", "RE_Copys", "BIB_Copys"];
  var EMOJI_REGEX = /[\u{1F1E6}-\u{1FAFF}\u{2600}-\u{27BF}\u{2190}-\u{21FF}\u{2B00}-\u{2BFF}\u{FE0F}\u{200D}]/gu;
  var resultado = [];

  tabs.forEach(function (nombreHoja) {
    var ws = SS.getSheetByName(nombreHoja);
    if (!ws) { resultado.push(nombreHoja + ": pestaña no encontrada"); return; }

    var headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0];
    var colCopy = headers.indexOf("Copy_Texto");
    if (colCopy === -1) { resultado.push(nombreHoja + ": no tiene columna Copy_Texto"); return; }

    var data = ws.getDataRange().getValues();
    var limpiados = 0;
    for (var i = 1; i < data.length; i++) {
      var texto = data[i][colCopy];
      if (typeof texto !== "string" || !texto) continue;
      var limpio = texto
        .replace(EMOJI_REGEX, "")
        .split("\n").map(function (linea) { return linea.replace(/[ \t]{2,}/g, " ").trim(); }).join("\n")
        .replace(/\n{3,}/g, "\n\n")
        .trim();
      if (limpio !== texto) {
        ws.getRange(i + 1, colCopy + 1).setValue(limpio);
        limpiados++;
      }
    }
    resultado.push(nombreHoja + ": " + limpiados + " copy(s) limpiado(s)");
  });

  SpreadsheetApp.getUi().alert("Emojis eliminados de los copys:\n\n" + resultado.join("\n"));
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
    if (tipo === "ap_reactivar")      return handleApReactivar(payload);
    if (tipo === "ap_diagnostico")    return handleApDiagnostico(payload);
    if (tipo === "alta_unificada")    return handleAltaUnificada(payload);
    if (tipo === "baja_empresa")      return handleBajaEmpresa(payload);
    if (tipo === "sv_atendido")       return handleSvAtendido(payload);
    if (tipo === "sv_baja")           return handleSvBaja(payload);
    if (tipo === "sv_reactivar")      return handleSvReactivar(payload);
    if (tipo === "eg_marcarAtendido")   return handleEgMarcarAtendido(payload);
    if (tipo === "eg_baja")             return handleEgBaja(payload);
    if (tipo === "eg_reactivar")        return handleEgReactivar(payload);
    if (tipo === "eg_diagnostico")      return handleEgDiagnostico(payload);
    if (tipo === "ra_actualizarSemana") return handleRaActualizarSemana(payload);
    if (tipo === "ra_actualizarSemanaLote") return handleRaActualizarSemanaLote(payload);
    if (tipo === "ra_editarCopy")        return handleRaEditarCopy(payload);
    if (tipo === "ra_marcarAtendido")   return handleRaMarcarAtendido(payload);
    if (tipo === "ra_baja")             return handleRaBaja(payload);
    if (tipo === "ra_reactivar")        return handleRaReactivar(payload);
    if (tipo === "re_altaEdicion")           return handleReAltaEdicion(payload);
    if (tipo === "re_editarEdicion")         return handleReEditarEdicion(payload);
    if (tipo === "re_registro")             return handleReRegistro(payload);
    if (tipo === "re_asistenciaLote")        return handleReAsistenciaLote(payload);
    if (tipo === "re_guardarHabitaciones")   return handleReGuardarHabitaciones(payload);
    if (tipo === "re_altaInteresado")            return handleReAltaInteresado(payload);
    if (tipo === "re_marcarAtendidoInteresado")  return handleReMarcarAtendidoInteresado(payload);
    if (tipo === "re_bajaInteresado")            return handleReBajaInteresado(payload);
    if (tipo === "re_reactivarInteresado")       return handleReReactivarInteresado(payload);
    if (tipo === "re_inscribirInteresado")       return handleReInscribirInteresado(payload);
    if (tipo === "cs_marcarAtendido")            return handleCsMarcarAtendido(payload);
    if (tipo === "cs_registrarFecha")            return handleCsRegistrarFecha(payload);
    if (tipo === "cs_baja")                      return handleCsBaja(payload);
    if (tipo === "cs_reactivar")                 return handleCsReactivar(payload);
    if (tipo === "ig_actualizarCampo")   return handleIgActualizarCampo(payload);
    if (tipo === "ig_marcarAtendido")    return handleIgMarcarAtendido(payload);
    if (tipo === "ig_baja")              return handleIgBaja(payload);
    if (tipo === "ig_reactivar")         return handleIgReactivar(payload);
    if (tipo === "ig_pasarGeneracion2")  return handleIgPasarGeneracion2(payload);
    if (tipo === "be_marcarAtendido")    return handleBeMarcarAtendido(payload);
    if (tipo === "be_baja")              return handleBeBaja(payload);
    if (tipo === "be_reactivar")         return handleBeReactivar(payload);
    if (tipo === "be_revisarInscripcion") return handleBeRevisarInscripcion(payload);
    if (tipo === "ea_registrarCilindro")  return handleEaRegistrarCilindro(payload);
    if (tipo === "ea_marcarAtendido")     return handleEaMarcarAtendido(payload);
    if (tipo === "ea_baja")               return handleEaBaja(payload);
    if (tipo === "ea_reactivar")          return handleEaReactivar(payload);
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

// ── Salvando Vidas: Reactivar (deshace una baja individual) ──────
function handleSvReactivar(p) {
  var nombre = (p.nombre || "").trim();
  if (_marcarEstadoPorNombre("SV_Inscritos", nombre, "Activo", "")) return resp({ ok: true });
  return resp({ ok: false, error: "No se encontró a " + nombre + " en SV_Inscritos" });
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

function handleEgReactivar(p) {
  var nombre = (p.nombre || "").trim();
  if (_marcarEstadoPorNombre("EG_Inscritos", nombre, "Activo", "")) return resp({ ok: true });
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

function handleApReactivar(p) {
  var nombre = (p.nombre || "").trim();
  if (_marcarEstadoPorNombre("AP_Inscritos", nombre, "Activo", "")) return resp({ ok: true });
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

// ── Baja de empresa (colaborador deja de trabajar en GEB) ────────
// Marca Baja_Empresa="Sí" + Fecha_Baja_Empresa en Colaboradores (col K/L —
// hoy que se agregaron a las 10 columnas base: Nombre,Sucursal,Unidad,Correo,
// FechaNac,Genero,Telefono_WA,Turno,Rol,Fecha_Alta_Empresa). Es de una sola
// vía (no hay handler de "reactivar empresa" — si regresa a trabajar se
// maneja aparte). En cascada, marca Baja en las 7 iniciativas restringidas
// (HOJAS_RESTRINGIDAS_BAJA_EMPRESA) donde ya no puede participar sin ser
// colaborador; Salvando Vidas y Biblioteca no se tocan aquí porque ahí sí
// puede seguir participando.
function handleBajaEmpresa(p) {
  var nombre = (p.nombre || "").trim();
  if (!nombre) return resp({ ok: false, error: "Falta el nombre" });

  var wsColab = SS.getSheetByName("Colaboradores");
  if (!wsColab) return resp({ ok: false, error: "Pestaña Colaboradores no encontrada en el Sheet nuevo" });

  var data = wsColab.getDataRange().getValues();
  var nombreLower = nombre.toLowerCase();
  var fila = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombreLower) { fila = i; break; }
  }
  if (fila === -1) return resp({ ok: false, error: "No se encontró a " + nombre + " en Colaboradores" });

  wsColab.getRange(fila + 1, 11).setValue("Sí");       // Baja_Empresa (col K)
  wsColab.getRange(fila + 1, 12).setValue(new Date()); // Fecha_Baja_Empresa (col L)

  var bajasAplicadas = [];
  HOJAS_RESTRINGIDAS_BAJA_EMPRESA.forEach(function (sheetName) {
    if (_marcarEstadoPorNombre(sheetName, nombre, "Baja", "No")) bajasAplicadas.push(sheetName);
  });

  return resp({ ok: true, bajasAplicadas: bajasAplicadas });
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

// ── Reto Ahorro: actualizar semana en lote (selección múltiple en el panel) ──
// Mismo efecto que handleRaActualizarSemana, pero para varios nombres a la vez
// — evita tener que abrir el formulario individual uno por uno.
function handleRaActualizarSemanaLote(p) {
  var ws = raSheetNuevo_();
  if (!ws) return resp({ ok: false, error: "Pestaña RA_Inscritos no encontrada en el Sheet nuevo" });

  var nombres = (p.nombres || []).map(function (n) { return String(n).trim().toLowerCase(); });
  var semana  = parseInt(p.semana, 10) || 0;
  var data    = ws.getDataRange().getValues();
  var actualizados = [];

  for (var i = 1; i < data.length; i++) {
    var nombreFila = String(data[i][0]).trim().toLowerCase();
    if (nombres.indexOf(nombreFila) === -1) continue;
    ws.getRange(i + 1, 8).setValue(semana);       // Semana_Actual (col H)
    ws.getRange(i + 1, 6).setValue("");           // Requiere_Seguimiento (col F)
    ws.getRange(i + 1, 9).setValue(new Date());   // Fecha_Ultima_Semana (col I)
    actualizados.push(data[i][0]);
  }

  return resp({ ok: true, actualizados: actualizados });
}

// ── Reto Ahorro: editar el texto de un copy directo desde el panel ──────
// RA_Copys: A=Momento, B=Copy_Texto, C=Activo. Antes solo se podía editar
// entrando al Sheet directamente — esto permite hacerlo desde el portal.
function handleRaEditarCopy(p) {
  var ws = SS.getSheetByName("RA_Copys");
  if (!ws) return resp({ ok: false, error: "Pestaña RA_Copys no encontrada en el Sheet nuevo" });

  var momento = (p.momento || "").trim();
  var data = ws.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === momento) {
      ws.getRange(i + 1, 2).setValue(p.copyTexto || ""); // Copy_Texto (col B)
      return resp({ ok: true });
    }
  }

  return resp({ ok: false, error: "No se encontró el momento '" + momento + "' en RA_Copys" });
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

function handleRaReactivar(p) {
  var nombre = (p.nombre || "").trim();
  if (_marcarEstadoPorNombre("RA_Inscritos", nombre, "Activo", "")) return resp({ ok: true });
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

// ── Retiro del Espíritu Santo: Editar una edición existente ──────
// Corrige datos de una edición ya creada (nombre, fechas, lugar, estado,
// notas) sin tocar su id_edicion ni las asistencias/habitaciones ya
// ligadas a ella. Pensado para arreglar errores de captura, incluyendo el
// campo Estado (a mano hoy tiene valores como "Completada"/"Pendiente de
// registro" que no siempre coinciden con las opciones del alta original).
function handleReEditarEdicion(p) {
  const ws = reEdicionesSheetNuevo_();
  if (!ws) return resp({ ok: false, error: "Pestaña RE_Ediciones no encontrada en el Sheet nuevo" });

  const idEdicion = (p.idEdicion || "").trim();
  if (!idEdicion) return resp({ ok: false, error: "Falta el id de la edición" });

  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === idEdicion) {
      ws.getRange(i + 1, 2).setValue(p.nombreEdicion || "");
      ws.getRange(i + 1, 3).setValue(p.fechasTexto || "");
      ws.getRange(i + 1, 4).setValue(p.fechaInicio ? new Date(p.fechaInicio) : "");
      ws.getRange(i + 1, 5).setValue(p.fechaFin ? new Date(p.fechaFin) : "");
      ws.getRange(i + 1, 6).setValue(p.lugar || "");
      ws.getRange(i + 1, 7).setValue(p.estado || "");
      ws.getRange(i + 1, 8).setValue(p.notas || "");
      return resp({ ok: true });
    }
  }

  return resp({ ok: false, error: "No se encontró la edición " + idEdicion });
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

function handleReReactivarInteresado(p) {
  var nombre = (p.nombre || "").trim();
  if (_marcarEstadoPorNombre("RE_Interesados", nombre, "Interesado", "")) return resp({ ok: true });
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

function handleCsReactivar(p) {
  const nombre = (p.nombre || "").trim();
  if (_marcarEstadoPorNombre("CS_Inscritos", nombre, "Activo", "")) return resp({ ok: true });
  return resp({ ok: false, error: "No se encontró a " + nombre + " en CS_Inscritos" });
}

// ================================================================
// BIBLIOTECA COMUNITARIA (Sheet nuevo desde 2026-07-05)
// ================================================================
// Los 3 Forms de préstamos/donaciones/devoluciones ya apuntan directo a
// SHEET_NUEVO_ID (reconectados 2026-07-10) — el resto de la lógica (matching
// de devoluciones, PDFs) usa lookup de encabezados por nombre
// (obtenerIndicesBib_), no posición fija, así que da igual en qué Sheet viva.
function bibSheetNuevo_(nombre) {
  return SpreadsheetApp.openById(SHEET_NUEVO_ID).getSheetByName(nombre);
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

function handleIgReactivar(p) {
  const nombre = (p.nombre || "").trim();
  if (_marcarEstadoPorNombre("IG_Inscritos", nombre, "Activo", "")) return resp({ ok: true });
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

// Vuelve a "Interesado" (no "Inscrito" — si ya se había inscrito antes de la
// baja, Cecilia revisa/marca la inscripción de nuevo desde el panel).
function handleBeReactivar(p) {
  const nombre = (p.nombre || "").trim();
  if (_marcarEstadoPorNombre("BE_Inscritos", nombre, "Interesado", "")) return resp({ ok: true });
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

function handleEaReactivar(p) {
  const nombre = (p.nombre || "").trim();
  if (_marcarEstadoPorNombre("EA_Lideres", nombre, "Activo", "")) return resp({ ok: true });
  return resp({ ok: false, error: "No se encontró a " + nombre + " en EA_Lideres" });
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
  borrarDatosPruebaTodos: borrarDatosPruebaTodos,
  configurarAlertasAtrasoRA: configurarAlertasAtrasoRA,
  limpiarEmojisCopys: limpiarEmojisCopys
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
