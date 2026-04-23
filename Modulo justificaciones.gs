// ==========================================
// MODULO_JUSTIFICACIONES.GS — Procesamiento de justificaciones regionales
// ==========================================

// ==========================================
// SWITCH Y CONFIGURACIÓN
// ==========================================

/**
 * Obtiene el estado del switch de justificaciones con soporte multi-región
 */
function obtenerEstadoSwitchJustificaciones() {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get('justif_switch_state_v2');
    if (cached) {
      try { return JSON.parse(cached); } catch (e) {}
    }

    var ss = getSpreadsheet('JUSTIFICACIONES');
    var sheetConfig = ss.getSheetByName(CONFIG.HOJAS.CONFIG_JUSTIFICACIONES);

    // Crear hoja si no existe
    if (!sheetConfig) {
      sheetConfig = ss.insertSheet(CONFIG.HOJAS.CONFIG_JUSTIFICACIONES);
      sheetConfig.appendRow(["REGION","Habilitado","Fecha Limite","Fecha_Evento","Nombre_Actividad"]);
    } else {
      // Migrar formato antiguo si aplica
      var primeraFila = sheetConfig.getRange(1, 1, 1, 1).getValue();
      var primeraFilaStr = String(primeraFila).trim().toUpperCase();
      if (primeraFilaStr === "HABILITADO" || primeraFilaStr === "TRUE" || primeraFilaStr === "FALSE") {
        var datosAntiguos = sheetConfig.getDataRange().getValues();
        var filaAntiguaConfig = datosAntiguos.length > 1 ? datosAntiguos[1] : null;
        sheetConfig.clearContents();
        sheetConfig.appendRow(["REGION","Habilitado","Fecha Limite","Fecha_Evento","Nombre_Actividad"]);
        if (filaAntiguaConfig) {
          var habAnt = filaAntiguaConfig[0] === true || String(filaAntiguaConfig[0]).toLowerCase() === "true";
          if (habAnt) {
            sheetConfig.appendRow(["13. Región Metropolitana de Santiago", true, filaAntiguaConfig[1] || "", filaAntiguaConfig[2] || "", "Asamblea General"]);
          }
        }
        Logger.log("✅ CONFIG_JUSTIFICACIONES migrado a formato multi-región");
      }
    }

    var lastRow = sheetConfig.getLastRow();
    if (lastRow < 2) return { habilitado: false, fechaLimite: "", fechaEvento: null, configuraciones: [] };

    var data = sheetConfig.getRange(2, 1, lastRow - 1, 5).getValues();
    var configuraciones = [];
    var ahora = new Date();

    for (var i = 0; i < data.length; i++) {
      var region = String(data[i][0] || "").trim();
      var habilitado = (data[i][1] === true || String(data[i][1]).toLowerCase() === "true");
      var fechaLimiteRaw = data[i][2];
      var fechaEventoRaw = data[i][3];
      var nombreActividad = String(data[i][4] || "Asamblea").trim();

      if (!region) continue;

      var fechaLimiteValue = fechaLimiteRaw
        ? (fechaLimiteRaw instanceof Date ? fechaLimiteRaw.toISOString() : String(fechaLimiteRaw).trim())
        : "";
      var fechaEvento = (fechaEventoRaw && String(fechaEventoRaw).trim() !== "")
        ? (fechaEventoRaw instanceof Date
            ? Utilities.formatDate(fechaEventoRaw, Session.getScriptTimeZone(), "yyyy-MM-dd")
            : String(fechaEventoRaw).trim())
        : null;

      // Auto-deshabilitar si venció la fecha límite
      if (habilitado && fechaLimiteValue) {
        var limite = new Date(fechaLimiteValue);
        if (ahora > limite) {
          sheetConfig.getRange(i + 2, 2).setValue(false);
          habilitado = false;
          Logger.log("⏰ Región " + region + " deshabilitada automáticamente por vencimiento de plazo");
        }
      }

      configuraciones.push({ region: region, habilitado: habilitado, fechaLimite: fechaLimiteValue, fechaEvento: fechaEvento, nombreActividad: nombreActividad });
    }

    var algunaHabilitada = configuraciones.some(function(c) { return c.habilitado; });
    var resultado = { habilitado: algunaHabilitada, configuraciones: configuraciones };

    try { cache.put('justif_switch_state_v2', JSON.stringify(resultado), 120); } catch (e) {}
    return resultado;

  } catch (e) {
    Logger.log('Error en obtenerEstadoSwitchJustificaciones: ' + e.toString());
    return { habilitado: false, fechaLimite: "", fechaEvento: null, configuraciones: [] };
  }
}

/**
 * Actualiza el switch de justificaciones para una región específica
 */
function actualizarSwitchJustificaciones(nuevoEstado, fechaLimite, fechaEvento, region, nombreActividad) {
  var lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) {
    try {
      try { CacheService.getScriptCache().remove('justif_switch_state_v2'); } catch(e) {}

      var ss = getSpreadsheet('JUSTIFICACIONES');
      var sheetConfig = ss.getSheetByName(CONFIG.HOJAS.CONFIG_JUSTIFICACIONES);
      if (!sheetConfig) {
        sheetConfig = ss.insertSheet(CONFIG.HOJAS.CONFIG_JUSTIFICACIONES);
        sheetConfig.appendRow(["REGION","Habilitado","Fecha Limite","Fecha_Evento","Nombre_Actividad"]);
      }

      // Deshabilitar TODAS si no hay región específica
      if (!nuevoEstado && !region) {
        var lr = sheetConfig.getLastRow();
        if (lr >= 2) sheetConfig.getRange(2, 2, lr - 1, 1).setValue(false);
        return { success: true, message: "Todas las configuraciones deshabilitadas." };
      }

      var regionTarget = region ? String(region).trim() : "13. Región Metropolitana de Santiago";
      var nombreAct = (nombreActividad && String(nombreActividad).trim() !== "") ? String(nombreActividad).trim() : "Asamblea General";
      var valorEvento = (fechaEvento && String(fechaEvento).trim() !== "") ? String(fechaEvento).trim() : "";

      var lastRow = sheetConfig.getLastRow();
      var filaExistente = -1;
      if (lastRow >= 2) {
        var dataActual = sheetConfig.getRange(2, 1, lastRow - 1, 1).getValues();
        for (var i = 0; i < dataActual.length; i++) {
          if (String(dataActual[i][0]).trim() === regionTarget) { filaExistente = i + 2; break; }
        }
      }

      if (filaExistente > 0) {
        sheetConfig.getRange(filaExistente, 1, 1, 5).setValues([[regionTarget, nuevoEstado, fechaLimite || "", valorEvento, nombreAct]]);
      } else {
        sheetConfig.appendRow([regionTarget, nuevoEstado, fechaLimite || "", valorEvento, nombreAct]);
      }

      return { success: true, message: "Configuración actualizada para " + regionTarget };
    } catch (e) {
      return { success: false, message: "Error: " + e.toString() };
    } finally {
      lock.releaseLock();
    }
  } else {
    return { success: false, message: "Servidor ocupado." };
  }
}

/**
 * Elimina la configuración de una región específica
 */
function eliminarConfigRegionJustificaciones(region) {
  try {
    if (!region || String(region).trim() === "") return { success: false, message: "Debe indicar una región." };
    var regionTarget = String(region).trim();
    try { CacheService.getScriptCache().remove('justif_switch_state_v2'); } catch(e) {}
    var ss = getSpreadsheet('JUSTIFICACIONES');
    var sheetConfig = ss.getSheetByName(CONFIG.HOJAS.CONFIG_JUSTIFICACIONES);
    if (!sheetConfig) return { success: false, message: "Hoja de configuración no encontrada." };
    var lastRow = sheetConfig.getLastRow();
    if (lastRow < 2) return { success: false, message: "No hay configuraciones registradas." };
    var data = sheetConfig.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === regionTarget) {
        sheetConfig.deleteRow(i + 2);
        return { success: true, message: 'Configuración de "' + regionTarget + '" eliminada.' };
      }
    }
    return { success: false, message: "No se encontró configuración para esa región." };
  } catch (e) {
    return { success: false, message: "Error: " + e.toString() };
  }
}

// ==========================================
// VALIDACIONES DE DISPONIBILIDAD
// ==========================================

/**
 * Obtiene la info de actividad para la región del usuario
 */
function obtenerInfoActividadPorRegion(rut) {
  try {
    var usuario = obtenerUsuarioPorRut(rut);
    if (!usuario.encontrado) return { habilitado: false, sinConfiguracion: true, mensaje: "Usuario no encontrado." };

    var regionUsuario = String(usuario.region || "").trim();
    if (!regionUsuario) return { habilitado: false, sinRegion: true, mensaje: "No tienes una región asignada en el sistema. Por favor actualiza tu región en Mis Datos." };

    var estadoSwitch = obtenerEstadoSwitchJustificaciones();
    var configuraciones = estadoSwitch.configuraciones || [];

    if (configuraciones.length === 0) return { habilitado: false, sinConfiguracion: true, regionUsuario: regionUsuario, mensaje: "El módulo de justificaciones no está activo en este momento. Consulta con tu delegado sindical." };

    var configRegion = null;
    for (var i = 0; i < configuraciones.length; i++) {
      if (configuraciones[i].region === regionUsuario) { configRegion = configuraciones[i]; break; }
    }

    if (!configRegion) return { habilitado: false, regionNoConfigurada: true, regionUsuario: regionUsuario, mensaje: "No hay una actividad programada para tu región (" + regionUsuario + ") en este momento.\n\nSi crees que se trata de un error, comunícate con tu delegado sindical o verifica que tu región esté correctamente registrada en Mis Datos." };

    if (!configRegion.habilitado) return { habilitado: false, vencido: true, regionUsuario: regionUsuario, nombreActividad: configRegion.nombreActividad, fechaEvento: configRegion.fechaEvento, mensaje: 'El plazo para justificaciones de la actividad "' + configRegion.nombreActividad + '" ha vencido para tu región.' };

    return { habilitado: true, regionUsuario: regionUsuario, nombreActividad: configRegion.nombreActividad, fechaLimite: configRegion.fechaLimite, fechaEvento: configRegion.fechaEvento };

  } catch (e) {
    Logger.log("Error en obtenerInfoActividadPorRegion: " + e.toString());
    return { habilitado: false, sinConfiguracion: true, mensaje: "Error al verificar configuración: " + e.message };
  }
}

/**
 * Verifica si el usuario ya tiene justificación para la actividad activa de su región
 */
function verificarJustificacionActividad(rut) {
  try {
    var infoActividad = obtenerInfoActividadPorRegion(rut);
    if (!infoActividad.habilitado) return { tieneJustificacion: false, sinActividad: true };

    var codigoActividad = null;
    if (infoActividad.fechaEvento && infoActividad.nombreActividad) {
      codigoActividad = infoActividad.fechaEvento + "_" + infoActividad.nombreActividad;
    } else if (infoActividad.fechaEvento) {
      codigoActividad = generarCodigoAsambleaEvento(infoActividad.fechaEvento);
    }

    if (!codigoActividad) return { tieneJustificacion: false };

    SpreadsheetApp.flush();
    var sheet = getSheet('JUSTIFICACIONES', 'JUSTIFICACIONES');
    var data = sheet.getDataRange().getValues();
    var COL = CONFIG.COLUMNAS.JUSTIFICACIONES;
    var rutLimpio = cleanRut(rut);

    for (var i = 1; i < data.length; i++) {
      var filaRut = cleanRut(String(data[i][COL.RUT]));
      var filaAsamblea = String(data[i][COL.ASAMBLEA] || "").trim();
      var filaEstado = String(data[i][COL.ESTADO] || "").trim();
      if (filaRut === rutLimpio && filaAsamblea === codigoActividad && filaEstado !== "Rechazado") {
        return { tieneJustificacion: true, estado: filaEstado, codigoActividad: codigoActividad, nombreActividad: infoActividad.nombreActividad, idJustificacion: String(data[i][COL.ID]) };
      }
    }
    return { tieneJustificacion: false, codigoActividad: codigoActividad };
  } catch (e) {
    return { tieneJustificacion: false, error: e.toString() };
  }
}

/**
 * Verifica disponibilidad del módulo de justificaciones para un RUT dado
 */
function verificarDisponibilidadJustificaciones(rut) {
  if (rut) {
    var infoActividad = obtenerInfoActividadPorRegion(rut);
    return {
      habilitado: infoActividad.habilitado,
      mensaje: infoActividad.habilitado ? "" : (infoActividad.mensaje || "Módulo deshabilitado para tu región."),
      regionNoConfigurada: infoActividad.regionNoConfigurada || false,
      sinRegion: infoActividad.sinRegion || false,
      vencido: infoActividad.vencido || false
    };
  }
  var estadoSwitch = obtenerEstadoSwitchJustificaciones();
  if (!estadoSwitch.habilitado) return { habilitado: false, mensaje: "Módulo de justificaciones temporalmente deshabilitado.\nConsulte con la directiva." };
  return { habilitado: true };
}

/**
 * Valida si el usuario puede enviar una justificación (por evento o por mes)
 */
function validarJustificacionMesActual(rut) {
  try {
    SpreadsheetApp.flush();
    var sheet = getSheet('JUSTIFICACIONES', 'JUSTIFICACIONES');
    var data = sheet.getDataRange().getValues();
    var COL = CONFIG.COLUMNAS.JUSTIFICACIONES;
    var hoy = new Date();

    var infoActividad = obtenerInfoActividadPorRegion(rut);
    var fechaEvento = (infoActividad.habilitado && infoActividad.fechaEvento) ? infoActividad.fechaEvento : null;
    var codigoEventoActivo = null;
    if (fechaEvento && infoActividad.nombreActividad) {
      codigoEventoActivo = fechaEvento + "_" + infoActividad.nombreActividad;
    } else if (fechaEvento) {
      codigoEventoActivo = generarCodigoAsambleaEvento(fechaEvento);
    }

    Logger.log("Validando justificacion | Region: " + (infoActividad.regionUsuario || "?") + " | Evento: " + (codigoEventoActivo || "Sin evento (modo mes)"));

    // ── MODO A: Validación por evento específico ──
    if (codigoEventoActivo) {
      var justificacionesDelEvento = [];
      for (var i = 1; i < data.length; i++) {
        if (cleanRut(data[i][COL.RUT]) === cleanRut(rut) && String(data[i][COL.ASAMBLEA] || "").trim() === codigoEventoActivo) {
          justificacionesDelEvento.push({ id: data[i][COL.ID], estado: data[i][COL.ESTADO], tipo: data[i][COL.MOTIVO], fecha: data[i][COL.FECHA] ? Utilities.formatDate(new Date(data[i][COL.FECHA]), Session.getScriptTimeZone(), "dd/MM/yyyy") : "", asamblea: data[i][COL.ASAMBLEA] });
        }
      }
      if (justificacionesDelEvento.length === 0) return { permitido: true, mensaje: "Puede enviar la justificación", justificacionExistente: null };
      var todasRechazadasA = justificacionesDelEvento.every(function(j) { return j.estado === 'Rechazado'; });
      if (todasRechazadasA) return { permitido: true, mensaje: "Puede reintentar (anterior rechazada)", justificacionExistente: justificacionesDelEvento[0] };
      var justEnviada = justificacionesDelEvento.find(function(j) { return j.estado === 'Enviado'; });
      var justAceptada = justificacionesDelEvento.find(function(j) { return j.estado === 'Aceptado' || j.estado === 'Aceptado/Obs'; });
      if (justEnviada) return { permitido: false, mensaje: "Ya tienes una justificación pendiente para el evento " + codigoEventoActivo, justificacionExistente: justEnviada, tipoBloqueo: 'enviada', codigoEvento: codigoEventoActivo };
      if (justAceptada) return { permitido: false, mensaje: "Ya tienes una justificación aceptada para el evento " + codigoEventoActivo, justificacionExistente: justAceptada, tipoBloqueo: 'aceptada', codigoEvento: codigoEventoActivo };
      return { permitido: false, mensaje: "Ya existe una justificación para el evento " + codigoEventoActivo, justificacionExistente: justificacionesDelEvento[0], tipoBloqueo: 'enviada', codigoEvento: codigoEventoActivo };
    }

    // ── MODO B: Fallback — validación por mes calendario ──
    var mesActual = hoy.getMonth(), yearActual = hoy.getFullYear();
    var justificacionesDelMes = [];
    for (var j = 1; j < data.length; j++) {
      var filaFecha = new Date(data[j][COL.FECHA]);
      if (cleanRut(data[j][COL.RUT]) === cleanRut(rut) && filaFecha.getMonth() === mesActual && filaFecha.getFullYear() === yearActual) {
        justificacionesDelMes.push({ id: data[j][COL.ID], estado: data[j][COL.ESTADO], tipo: data[j][COL.MOTIVO], fecha: Utilities.formatDate(filaFecha, Session.getScriptTimeZone(), "dd/MM/yyyy"), asamblea: data[j][COL.ASAMBLEA] || "" });
      }
    }
    if (justificacionesDelMes.length === 0) return { permitido: true, mensaje: "Puede enviar la justificación", justificacionExistente: null };
    var nombreMes = hoy.toLocaleString('es-CL', { month: 'long', year: 'numeric' });
    var todasRechazadasB = justificacionesDelMes.every(function(j) { return j.estado === 'Rechazado'; });
    if (todasRechazadasB) return { permitido: true, mensaje: "Puede reintentar (anterior rechazada)", justificacionExistente: justificacionesDelMes[0] };
    var hayEnviada = justificacionesDelMes.some(function(j) { return j.estado === 'Enviado'; });
    var hayAceptada = justificacionesDelMes.some(function(j) { return j.estado === 'Aceptado' || j.estado === 'Aceptado/Obs'; });
    if (hayEnviada) return { permitido: false, mensaje: "Ya tienes una justificación pendiente para " + nombreMes, justificacionExistente: justificacionesDelMes.find(function(j) { return j.estado === 'Enviado'; }), tipoBloqueo: 'enviada' };
    if (hayAceptada) return { permitido: false, mensaje: "Ya tienes una justificación aceptada para " + nombreMes, justificacionExistente: justificacionesDelMes.find(function(j) { return j.estado === 'Aceptado' || j.estado === 'Aceptado/Obs'; }), tipoBloqueo: 'aceptada' };
    return { permitido: false, mensaje: "Límite de justificaciones alcanzado para " + nombreMes, justificacionExistente: justificacionesDelMes[0], tipoBloqueo: 'enviada' };

  } catch (error) {
    Logger.log('Error en validarJustificacionMesActual: ' + error.toString());
    return { permitido: false, mensaje: "Error al validar: " + error.message, justificacionExistente: null };
  }
}

// ==========================================
// CRUD JUSTIFICACIONES
// ==========================================

/**
 * Envía una justificación al sistema
 */
function enviarJustificacion(rutGestor, tipo, motivo, archivoData, rutBeneficiario) {
  var CARPETA_ID = CONFIG.CARPETAS.JUSTIFICACIONES;

  var rutParaVerif = rutBeneficiario ? rutBeneficiario : rutGestor;
  var disp = verificarDisponibilidadJustificaciones(rutParaVerif);
  if (!disp.habilitado) return { success: false, message: disp.mensaje || "Módulo de justificaciones no disponible para tu región." };

  var lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) {
    try {
      var sheetJustif = getSheet('JUSTIFICACIONES', 'JUSTIFICACIONES');
      var COL_JUST = CONFIG.COLUMNAS.JUSTIFICACIONES;

      var gestor = obtenerUsuarioPorRut(rutGestor);
      if (!gestor.encontrado) return { success: false, message: "Error de sesión." };

      var rutTarget = rutBeneficiario ? cleanRut(rutBeneficiario) : cleanRut(rutGestor);
      var esGestionDirigente = rutTarget !== cleanRut(rutGestor);
      var beneficiario;

      if (!esGestionDirigente) {
        beneficiario = gestor;
      } else {
        beneficiario = obtenerUsuarioPorRut(rutBeneficiario);
        if (!beneficiario.encontrado) return { success: false, message: "RUT del socio no encontrado." };
      }

      var validacion = validarJustificacionMesActual(beneficiario.rut);
      if (!validacion.permitido) {
        return { success: false, message: validacion.mensaje, tipoError: 'restriccion_mes', justificacionExistente: validacion.justificacionExistente, tipoBloqueo: validacion.tipoBloqueo, codigoEvento: validacion.codigoEvento || null };
      }

      var validacionCorreos = validarCorreosParaPermisos(
        { rut: beneficiario.rut, nombre: beneficiario.nombre, correo: beneficiario.correo },
        esGestionDirigente ? { rut: gestor.rut, nombre: gestor.nombre, correo: gestor.correo } : null,
        esGestionDirigente
      );

      var idUnico = Utilities.getUuid();
      var fileUrl = "Sin archivo";
      var alertaPermisos = null;

      if (archivoData && archivoData.base64) {
        var nombreArchivo = "JUSTIF-" + idUnico + "-" + cleanRut(beneficiario.rut);
        var resultadoSubida = subirArchivoConPermisos(archivoData, CARPETA_ID, nombreArchivo, validacionCorreos.correosParaPermisos, []);
        if (!resultadoSubida.success) return { success: false, message: resultadoSubida.mensajeError };
        fileUrl = resultadoSubida.url;
        alertaPermisos = generarAlertaPermisos(validacionCorreos, resultadoSubida);
      } else {
        alertaPermisos = generarAlertaPermisos(validacionCorreos, null);
      }

      var fechaHoy = new Date();
      var infoActividadParaRegistro = obtenerInfoActividadPorRegion(beneficiario.rut);
      var codigoAsamblea;
      if (infoActividadParaRegistro.habilitado && infoActividadParaRegistro.fechaEvento) {
        codigoAsamblea = infoActividadParaRegistro.fechaEvento + "_" + (infoActividadParaRegistro.nombreActividad || "Asamblea");
      } else {
        codigoAsamblea = generarCodigoAsamblea(fechaHoy);
      }

      var gestion = "Socio", nomDirigente = "", correoDirigente = "";
      if (esGestionDirigente) { gestion = "Dirigente"; nomDirigente = gestor.nombre; correoDirigente = gestor.correo; }

      var newRow = [];
      newRow[COL_JUST.ID]             = idUnico;
      newRow[COL_JUST.FECHA]          = fechaHoy;
      newRow[COL_JUST.RUT]            = beneficiario.rut;
      newRow[COL_JUST.NOMBRE]         = beneficiario.nombre;
      newRow[COL_JUST.REGION]         = beneficiario.region;
      newRow[COL_JUST.MOTIVO]         = tipo;
      newRow[COL_JUST.ARGUMENTO]      = motivo;
      newRow[COL_JUST.RESPALDO]       = fileUrl;
      newRow[COL_JUST.ESTADO]         = "Enviado";
      newRow[COL_JUST.OBSERVACION]    = "";
      newRow[COL_JUST.NOTIFICACION]   = "Enviado";
      newRow[COL_JUST.ASAMBLEA]       = codigoAsamblea;
      newRow[COL_JUST.GESTION]        = gestion;
      newRow[COL_JUST.DIRIGENTE]      = nomDirigente;
      newRow[COL_JUST.CORREO_DIRIGENTE] = correoDirigente;
      sheetJustif.appendRow(newRow);

      // Agregar validación de datos en celda ESTADO
      var lastRow = sheetJustif.getLastRow();
      var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Enviado','Aceptado','Aceptado/Obs','Rechazado'], true)
        .setAllowInvalid(false).build();
      sheetJustif.getRange(lastRow, COL_JUST.ESTADO + 1).setDataValidation(rule);

      // Correo al socio (si gestiona por sí mismo)
      if (!esGestionDirigente && esCorreoValido(beneficiario.correo)) {
        var respaldoDisplay = (fileUrl && fileUrl.includes("http"))
          ? '<a href="' + fileUrl + '" style="color:#ea580c;text-decoration:none;font-weight:bold;">Ver Documento Adjunto</a>'
          : "";
        enviarCorreoEstilizado(
          beneficiario.correo,
          "Justificación Ingresada - Sindicato SLIM n°3",
          "Comprobante de Justificación",
          "Hola <strong>" + beneficiario.nombre + "</strong>, tu justificación ha sido ingresada correctamente en el sistema. A continuación los detalles registrados:",
          { "FECHA": Utilities.formatDate(fechaHoy, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"), "RUT": formatRutServer(beneficiario.rut), "NOMBRE": beneficiario.nombre, "REGION": beneficiario.region, "MOTIVO": tipo, "ARGUMENTO": motivo, "RESPALDO": respaldoDisplay, "OBSERVACION": "", "ASAMBLEA": codigoAsamblea, "GESTION": gestion, "DIRIGENTE": nomDirigente },
          "#ea580c"
        );
      }

      // Correo de respaldo al dirigente
      if (esGestionDirigente && esCorreoValido(correoDirigente) && correoDirigente !== beneficiario.correo) {
        var respaldoDirigente = (fileUrl && fileUrl.includes("http"))
          ? '<a href="' + fileUrl + '" style="color:#475569;text-decoration:none;font-weight:bold;">Ver Documento Adjunto</a>'
          : "";
        enviarCorreoEstilizado(
          correoDirigente,
          "Respaldo Gestión Justificación - Sindicato SLIM n°3",
          "Gestión Realizada",
          "Has ingresado exitosamente una justificación para el socio <strong>" + beneficiario.nombre + "</strong>.",
          { "FECHA": Utilities.formatDate(fechaHoy, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"), "RUT": formatRutServer(beneficiario.rut), "NOMBRE": beneficiario.nombre, "REGION": beneficiario.region, "MOTIVO": tipo, "ARGUMENTO": motivo, "RESPALDO": respaldoDirigente, "OBSERVACION": "", "ASAMBLEA": codigoAsamblea, "GESTION": gestion, "DIRIGENTE": nomDirigente },
          "#475569"
        );
      }

      // Copia al socio cuando el dirigente gestiona en su nombre
      if (esGestionDirigente && esCorreoValido(beneficiario.correo)) {
        var respaldoSocio = (fileUrl && fileUrl.includes("http"))
          ? '<a href="' + fileUrl + '" style="color:#ea580c;text-decoration:none;font-weight:bold;">Ver Documento Adjunto</a>'
          : "";
        enviarCorreoEstilizado(
          beneficiario.correo,
          "Justificación Ingresada - Sindicato SLIM n°3",
          "Comprobante de Justificación",
          "Hola <strong>" + beneficiario.nombre + "</strong>, un dirigente ha ingresado una justificación a tu nombre. A continuación los detalles registrados:",
          { "FECHA": Utilities.formatDate(fechaHoy, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"), "RUT": formatRutServer(beneficiario.rut), "NOMBRE": beneficiario.nombre, "REGION": beneficiario.region, "MOTIVO": tipo, "ARGUMENTO": motivo, "RESPALDO": respaldoSocio, "OBSERVACION": "", "ASAMBLEA": codigoAsamblea, "GESTION": gestion, "DIRIGENTE": nomDirigente },
          "#ea580c"
        );
      }

      var respuesta = { success: true, message: "Justificación enviada exitosamente." };
      if (alertaPermisos && alertaPermisos.mostrarAlerta) {
        respuesta.mostrarAlerta = true;
        respuesta.tipoAlerta = alertaPermisos.tipoAlerta;
        respuesta.mensajeAlerta = alertaPermisos.mensajeAlerta;
      }
      return respuesta;

    } catch (e) {
      Logger.log("Error en enviarJustificacion: " + e.toString());
      return { success: false, message: "Error: " + e.toString() };
    } finally {
      lock.releaseLock();
    }
  } else {
    return { success: false, message: "Servidor ocupado." };
  }
}

/**
 * Elimina una justificación en estado "Enviado"
 */
function eliminarJustificacion(idJustif) {
  var lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) {
    try {
      var sheet = getSheet('JUSTIFICACIONES', 'JUSTIFICACIONES');
      var data = sheet.getDataRange().getValues();
      var COL = CONFIG.COLUMNAS.JUSTIFICACIONES;
      var estadoSwitch = obtenerEstadoSwitchJustificaciones();

      for (var i = 1; i < data.length; i++) {
        if (String(data[i][COL.ID]) === String(idJustif)) {
          var estado = String(data[i][COL.ESTADO]);
          if (!estadoSwitch.habilitado && estado === "Enviado") {
            return { success: false, message: "El plazo para agregar o modificar información ha vencido. Si al final del mes aparece con multa puede realizar la apelación." };
          }
          if (estado !== "Enviado") return { success: false, message: "No se puede eliminar." };
          sheet.deleteRow(i + 1);
          return { success: true, message: "Eliminado." };
        }
      }
      return { success: false, message: "No encontrado." };
    } catch (e) {
      return { success: false, message: "Error: " + e.toString() };
    } finally {
      lock.releaseLock();
    }
  } else {
    return { success: false, message: "Ocupado." };
  }
}

/**
 * Obtiene el historial de justificaciones de un usuario
 */
function obtenerHistorialJustificaciones(rutInput) {
  try {
    var sheet = getSheet('JUSTIFICACIONES', 'JUSTIFICACIONES');
    var COL = CONFIG.COLUMNAS.JUSTIFICACIONES;
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, registros: [] };
    var lastCol = sheet.getLastColumn();
    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
    var rutLimpio = cleanRut(rutInput);
    var registros = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (cleanRut(row[COL.RUT]) === rutLimpio) {
        registros.push({ id: row[COL.ID], fecha: formatearFechaConHora(row[COL.FECHA]), tipo: row[COL.MOTIVO], motivo: row[COL.ARGUMENTO], url: row[COL.RESPALDO], estado: row[COL.ESTADO], obs: row[COL.OBSERVACION], asamblea: row[COL.ASAMBLEA], gestion: row[COL.GESTION], nomDirigente: row[COL.DIRIGENTE] });
      }
    }
    registros.reverse();
    return { success: true, registros: registros };
  } catch (e) {
    Logger.log("❌ Error en obtenerHistorialJustificaciones: " + e.toString());
    return { success: false, message: "Error: " + e.toString() };
  }
}

// ==========================================
// TRIGGER — VERIFICAR CAMBIOS EN JUSTIFICACIONES
// ==========================================

/**
 * Trigger: cada 8 horas. Detecta cambios de estado y notifica al usuario.
 * NOTA: Esta es la versión correcta con getSheet(). La versión duplicada con
 * getActiveSpreadsheet() que estaba al final del Code.gs original fue eliminada.
 */
function verificarCambiosJustificaciones() {
  try {
    var sheet = getSheet('JUSTIFICACIONES', 'JUSTIFICACIONES');
    if (!sheet) { console.error("❌ No se pudo acceder a la hoja de justificaciones"); return; }

    var data = sheet.getDataRange().getValues();
    var COL = CONFIG.COLUMNAS.JUSTIFICACIONES;

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var idRegistro   = String(row[COL.ID]);
      var estadoActual = String(row[COL.ESTADO]);
      var estadoNotif  = String(row[COL.NOTIFICACION]);
      var nombre       = row[COL.NOMBRE];
      var tipo         = row[COL.MOTIVO];
      var obs          = row[COL.OBSERVACION];
      var asamblea     = row[COL.ASAMBLEA];
      var fechaSolicitud = row[COL.FECHA];
      var asambleaActual = row[COL.ASAMBLEA];

      if (fechaSolicitud && !asambleaActual) {
        var codigoAsamblea = generarCodigoAsamblea(new Date(fechaSolicitud));
        sheet.getRange(i + 1, COL.ASAMBLEA + 1).setValue(codigoAsamblea);
      }

      if (estadoActual !== estadoNotif) {
        var rutUsuario = row[COL.RUT];
        var sheetUsers = getSheet('USUARIOS', 'USUARIOS');
        if (!sheetUsers) { console.error("❌ No se pudo acceder a la hoja de usuarios"); continue; }

        var dataUsers = sheetUsers.getDataRange().getDisplayValues();
        var COL_USER = CONFIG.COLUMNAS.USUARIOS;
        var correoUsuario = "";

        for (var j = 1; j < dataUsers.length; j++) {
          if (cleanRut(dataUsers[j][COL_USER.RUT]) === cleanRut(rutUsuario)) {
            correoUsuario = dataUsers[j][COL_USER.CORREO];
            break;
          }
        }

        if (correoUsuario && correoUsuario.includes("@")) {
          var color = "#ea580c", titulo = "Actualización de Justificación";
          if (estadoActual.includes("Aceptado")) { color = "#15803d"; titulo = "Justificación Aceptada"; }
          else if (estadoActual.includes("Rechazado")) { color = "#b91c1c"; titulo = "Justificación Rechazada"; }

          enviarCorreoEstilizado(
            correoUsuario,
            titulo + " - Sindicato SLIM n°3",
            titulo,
            "Hola " + nombre + ", el estado de tu justificación ha cambiado.",
            { "ID": idRegistro, "Tipo": tipo, "Nuevo Estado": estadoActual, "Observación": obs || "Sin observaciones", "Asamblea": asamblea || "Pendiente asignación" },
            color
          );
        }
        sheet.getRange(i + 1, COL.NOTIFICACION + 1).setValue(estadoActual);
      }
    }
  } catch (e) {
    console.error("❌ Error verificando justificaciones: " + e.toString());
  }
}

/**
 * Corrige permisos de archivos de justificaciones existentes.
 * Ejecutar manualmente una sola vez.
 */
function corregirPermisosJustificacionesExistentes() {
  try {
    var sheet = getSheet('JUSTIFICACIONES', 'JUSTIFICACIONES');
    var data = sheet.getDataRange().getValues();
    var COL = CONFIG.COLUMNAS.JUSTIFICACIONES;
    var archivosCorregidos = 0, errores = 0;

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var gestion = String(row[COL.GESTION]);
      var urlArchivo = String(row[COL.RESPALDO]);
      var correoDirigente = String(row[COL.CORREO_DIRIGENTE]);
      var correoSocio = obtenerCorreoDeRut(row[COL.RUT]);

      if (gestion === "Dirigente" && urlArchivo.includes("drive.google.com") && correoDirigente && correoDirigente.includes("@")) {
        try {
          var fileId = "";
          if (urlArchivo.includes("/d/")) { fileId = urlArchivo.split("/d/")[1].split("/")[0]; }
          else if (urlArchivo.includes("id=")) { fileId = urlArchivo.split("id=")[1].split("&")[0]; }
          if (fileId) {
            var file = DriveApp.getFileById(fileId);
            var viewers = file.getViewers();
            var tieneDirigente = viewers.some(function(v) { return v.getEmail() === correoDirigente; });
            var tieneSocio = viewers.some(function(v) { return v.getEmail() === correoSocio; });
            var cambios = [];
            if (!tieneDirigente) { file.addViewer(correoDirigente); cambios.push("dirigente: " + correoDirigente); }
            if (correoSocio && correoSocio.includes("@") && !tieneSocio) { file.addViewer(correoSocio); cambios.push("socio: " + correoSocio); }
            if (cambios.length > 0) { archivosCorregidos++; Logger.log('✅ Fila ' + (i + 1) + ' - Permisos otorgados: ' + cambios.join(', ')); }
          }
        } catch (fileErr) { errores++; Logger.log('⚠️ Error en fila ' + (i + 1) + ': ' + fileErr.toString()); }
      }
    }

    Logger.log('📊 RESUMEN: ✅ Archivos corregidos: ' + archivosCorregidos + ' | ⚠️ Errores: ' + errores);
    return { success: true, archivosCorregidos: archivosCorregidos, errores: errores };
  } catch (error) {
    Logger.log('❌ Error: ' + error.message);
    return { success: false, message: error.message };
  }
}
