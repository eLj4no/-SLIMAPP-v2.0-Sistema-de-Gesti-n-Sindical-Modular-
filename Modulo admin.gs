// ==========================================
// MODULO_ADMIN.GS — Panel administrador, gestiones, triggers
// ==========================================

// ==========================================
// GESTIÓN DE SOCIOS (DIRIGENTE / ADMIN)
// ==========================================

/**
 * Obtiene TODAS las gestiones realizadas por dirigentes en todos los módulos.
 * Filtra registros donde gestion="Dirigente".
 */
function obtenerGestionesDirigente(rutDirigente) {
  try {
    var rutLimpio = cleanRut(rutDirigente);

    var resultado = { prestamos: [], justificaciones: [], apelaciones: [], permisosMedicos: [] };

    // PRÉSTAMOS
    var sheetPrestamos = getSheet('PRESTAMOS', 'PRESTAMOS');
    var dataPrestamos  = sheetPrestamos.getDataRange().getDisplayValues();
    var COL_PRES       = CONFIG.COLUMNAS.PRESTAMOS;

    for (var i = 1; i < dataPrestamos.length; i++) {
      var row          = dataPrestamos[i];
      var fechaTerminoStr = "S/D";
      var ftRaw        = row[COL_PRES.FECHA_TERMINO];
      if (ftRaw) {
        try {
          var d = new Date(ftRaw);
          fechaTerminoStr = !isNaN(d.getTime())
            ? Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy")
            : String(ftRaw).split(' ')[0];
        } catch(e) { fechaTerminoStr = String(ftRaw).split(' ')[0]; }
      }

      resultado.prestamos.push({
        id: row[COL_PRES.ID], fecha: row[COL_PRES.FECHA],
        rutSocio: row[COL_PRES.RUT], nombreSocio: row[COL_PRES.NOMBRE],
        tipo: row[COL_PRES.TIPO] || "Préstamo", monto: row[COL_PRES.MONTO] || "$0",
        cuotas: row[COL_PRES.CUOTAS] || "S/D", medio: row[COL_PRES.MEDIO_PAGO] || "S/D",
        estado: row[COL_PRES.ESTADO], observacion: row[COL_PRES.OBSERVACION] || "",
        fechaTermino: fechaTerminoStr
      });
    }

    // JUSTIFICACIONES
    var sheetJustif = getSheet('JUSTIFICACIONES', 'JUSTIFICACIONES');
    var dataJustif  = sheetJustif.getDataRange().getDisplayValues();
    var COL_JUST    = CONFIG.COLUMNAS.JUSTIFICACIONES;

    for (var j = 1; j < dataJustif.length; j++) {
      var rowJ = dataJustif[j];
      if (rowJ[COL_JUST.GESTION] === "Dirigente") {
        resultado.justificaciones.push({
          id: rowJ[COL_JUST.ID], fecha: rowJ[COL_JUST.FECHA],
          rutSocio: rowJ[COL_JUST.RUT], nombreSocio: rowJ[COL_JUST.NOMBRE],
          tipo: rowJ[COL_JUST.MOTIVO], motivo: rowJ[COL_JUST.ARGUMENTO],
          url: rowJ[COL_JUST.RESPALDO], estado: rowJ[COL_JUST.ESTADO],
          obs: rowJ[COL_JUST.OBSERVACION], asamblea: rowJ[COL_JUST.ASAMBLEA]
        });
      }
    }

    // APELACIONES
    var sheetApel = getSheet('APELACIONES', 'APELACIONES');
    var dataApel  = sheetApel.getDataRange().getDisplayValues();
    var COL_APEL  = CONFIG.COLUMNAS.APELACIONES;

    for (var k = 1; k < dataApel.length; k++) {
      var rowA = dataApel[k];
      if (rowA[COL_APEL.GESTION] === "Dirigente") {
        resultado.apelaciones.push({
          id: rowA[COL_APEL.ID], fecha: rowA[COL_APEL.FECHA_SOLICITUD],
          rutSocio: rowA[COL_APEL.RUT], nombreSocio: rowA[COL_APEL.NOMBRE],
          mesApelacion: rowA[COL_APEL.MES_APELACION], tipoMotivo: rowA[COL_APEL.TIPO_MOTIVO],
          detalleMotivo: rowA[COL_APEL.DETALLE_MOTIVO], urlComprobante: rowA[COL_APEL.URL_COMPROBANTE],
          urlLiquidacion: rowA[COL_APEL.URL_LIQUIDACION], estado: rowA[COL_APEL.ESTADO],
          obs: rowA[COL_APEL.OBSERVACION], urlComprobanteDevolucion: rowA[COL_APEL.URL_COMPROBANTE_DEVOLUCION] || ""
        });
      }
    }

    // PERMISOS MÉDICOS
    var sheetPermisos = getSheet('PERMISOS_MEDICOS', 'PERMISOS_MEDICOS');
    var dataPermisos  = sheetPermisos.getDataRange().getDisplayValues();
    var COL_PERM      = CONFIG.COLUMNAS.PERMISOS_MEDICOS;

    for (var m = 1; m < dataPermisos.length; m++) {
      var rowP = dataPermisos[m];
      if (rowP[COL_PERM.GESTION] === "Dirigente") {
        resultado.permisosMedicos.push({
          id: rowP[COL_PERM.ID], fecha: rowP[COL_PERM.FECHA_SOLICITUD],
          rutSocio: rowP[COL_PERM.RUT], nombreSocio: rowP[COL_PERM.NOMBRE],
          tipoPermiso: rowP[COL_PERM.TIPO_PERMISO], fechaInicio: rowP[COL_PERM.FECHA_INICIO],
          motivo: rowP[COL_PERM.MOTIVO_DETALLE], urlDocumento: rowP[COL_PERM.URL_DOCUMENTO],
          estado: rowP[COL_PERM.ESTADO]
        });
      }
    }

    return { success: true, datos: resultado };

  } catch (e) {
    return { success: false, message: "Error: " + e.toString() };
  }
}

// ==========================================
// INFORME ADMINISTRADOR
// ==========================================

/**
 * Genera informe Excel de préstamos "Solicitado", lo envía al correo ADMIN
 * y cambia el estado a "Enviado" en la BD.
 */
function generarInformeAdministrador() {
  try {
    var sheetPrestamos = getSheet('PRESTAMOS', 'PRESTAMOS');
    var data           = sheetPrestamos.getDataRange().getValues();
    var COL            = CONFIG.COLUMNAS.PRESTAMOS;

    var prestamosSolicitados = [], filasActualizar = [];

    for (var i = 1; i < data.length; i++) {
      var estado = String(data[i][COL.ESTADO]);
      if (estado !== "Solicitado") continue;

      prestamosSolicitados.push({
        rut:          data[i][COL.RUT],
        nombre:       data[i][COL.NOMBRE],
        tipoPrestamo: data[i][COL.TIPO],
        cuotas:       data[i][COL.CUOTAS],
        medioPago:    data[i][COL.MEDIO_PAGO],
        monto:        data[i][COL.MONTO]
      });
      filasActualizar.push(i + 1);
    }

    if (prestamosSolicitados.length === 0) {
      return { success: false, message: "No hay préstamos en estado 'Solicitado' para procesar." };
    }

    var ss = getSpreadsheet('PRESTAMOS');
    var sheetInforme = ss.getSheetByName("INFORME_PRESTAMOS_TEMP");
    if (sheetInforme) ss.deleteSheet(sheetInforme);
    sheetInforme = ss.insertSheet("INFORME_PRESTAMOS_TEMP");

    sheetInforme.appendRow(["RUT", "NOMBRE", "TIPO PRÉSTAMO", "CUOTAS", "MEDIO PAGO", "MONTO"]);
    prestamosSolicitados.forEach(function(p) {
      sheetInforme.appendRow([p.rut, p.nombre, p.tipoPrestamo, p.cuotas, p.medioPago, p.monto]);
    });

    var lastRow = sheetInforme.getLastRow();
    var lastCol = sheetInforme.getLastColumn();
    sheetInforme.getRange(1, 1, 1, lastCol).setFontWeight("bold").setBackground("#4c1d95").setFontColor("#ffffff");
    sheetInforme.setFrozenRows(1);
    sheetInforme.autoResizeColumns(1, lastCol);

    var url   = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export?format=xlsx&gid=" + sheetInforme.getSheetId();
    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token } });
    var blob  = response.getBlob();
    blob.setName("Informe_Prestamos_Solicitados_" + new Date().toLocaleDateString('es-CL').replace(/\//g, '-') + ".xlsx");

    var sheetUsers = getSheet('USUARIOS', 'USUARIOS');
    var dataUsers  = sheetUsers.getDataRange().getDisplayValues();
    var COL_USER   = CONFIG.COLUMNAS.USUARIOS;
    var correoAdmin = "admin@sindicato.com";
    for (var j = 1; j < dataUsers.length; j++) {
      if (String(dataUsers[j][COL_USER.ROL]).toUpperCase() === "ADMIN") {
        correoAdmin = dataUsers[j][COL_USER.CORREO];
        break;
      }
    }

    var htmlCorreo = '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">'
      + '<div style="background:linear-gradient(135deg,#4c1d95,#5b21b6);color:white;padding:30px;text-align:center;border-radius:10px 10px 0 0;">'
      + '<h1 style="margin:0;font-size:24px;">Informe de Préstamos Solicitados</h1></div>'
      + '<div style="background:#f8f9fa;padding:30px;border-radius:0 0 10px 10px;">'
      + '<p style="color:#1e293b;font-size:16px;">Se adjunta el informe de préstamos en estado <strong>"Solicitado"</strong> generado el <strong>'
      + new Date().toLocaleDateString('es-CL') + '</strong>.</p>'
      + '<p style="color:#64748b;font-size:14px;">Total de préstamos procesados: <strong>' + prestamosSolicitados.length + '</strong></p>'
      + '<p style="color:#dc2626;font-size:14px;font-weight:bold;">⚠️ Estos préstamos han sido cambiados automáticamente al estado "Enviado".</p>'
      + '<hr style="border:none;border-top:1px solid #e2e8f0;margin:20px 0;">'
      + '<p style="color:#94a3b8;font-size:12px;text-align:center;">Sindicato SLIM n°3 — Sistema de Gestión</p>'
      + '</div></div>';

    MailApp.sendEmail({
      to: correoAdmin,
      subject: "Informe de Préstamos Solicitados - Sindicato SLIM n°3",
      htmlBody: htmlCorreo,
      attachments: [blob]
    });

    filasActualizar.forEach(function(fila) {
      sheetPrestamos.getRange(fila, COL.ESTADO + 1).setValue("Enviado");
    });

    ss.deleteSheet(sheetInforme);

    return { success: true, message: "Informe generado y enviado. " + prestamosSolicitados.length + " préstamo(s) cambiado(s) a \"Enviado\"." };

  } catch (e) {
    return { success: false, message: "Error al generar informe: " + e.toString() };
  }
}

// ==========================================
// CAMBIO DE ROL (Panel Admin)
// ==========================================

/**
 * Busca usuarios por nombre con coincidencia parcial (fuzzy multi-word).
 * Retorna lista de candidatos con RUT, nombre, cargo, rol actual.
 */
function buscarUsuarioPorNombre(textoBusqueda) {
  try {
    if (!textoBusqueda || String(textoBusqueda).trim().length < 2) {
      return { success: false, message: "Ingresa al menos 2 caracteres para buscar." };
    }

    var sheet = getSheet('USUARIOS', 'USUARIOS');
    var data  = sheet.getDataRange().getDisplayValues();
    var COL   = CONFIG.COLUMNAS.USUARIOS;

    var palabras = String(textoBusqueda).trim().toUpperCase().split(/\s+/);
    var resultados = [];

    for (var i = 1; i < data.length; i++) {
      var nombreRow = String(data[i][COL.NOMBRE] || "").toUpperCase();
      var coincide  = palabras.every(function(p) { return nombreRow.indexOf(p) !== -1; });
      if (!coincide) continue;

      resultados.push({
        rut:     data[i][COL.RUT],
        nombre:  data[i][COL.NOMBRE],
        cargo:   data[i][COL.CARGO]  || "—",
        region:  data[i][COL.REGION] || "—",
        rolActual: String(data[i][COL.ROL] || "SOCIO").trim().toUpperCase()
      });

      if (resultados.length >= 10) break;
    }

    return { success: true, resultados: resultados, total: resultados.length };
  } catch (e) {
    Logger.log("❌ Error en buscarUsuarioPorNombre: " + e.toString());
    return { success: false, message: "Error: " + e.toString() };
  }
}

/**
 * Cambia el rol de un usuario y envía notificación por correo.
 * Solo ADMIN puede ejecutar esta función.
 */
function cambiarRolUsuario(rutAdmin, rutObjetivo, nuevoRol) {
  var ROLES_VALIDOS = ["SOCIO", "DIRIGENTE", "ADMIN", "TESTING"];

  var lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) {
    try {
      var validacion = verificarRolUsuario(rutAdmin, ['ADMIN']);
      if (!validacion.autorizado) {
        return { success: false, message: "No tienes permisos para cambiar roles." };
      }

      var nuevoRolNorm = String(nuevoRol || "").trim().toUpperCase();
      if (ROLES_VALIDOS.indexOf(nuevoRolNorm) === -1) {
        return { success: false, message: "Rol inválido. Roles permitidos: " + ROLES_VALIDOS.join(", ") };
      }

      var rutLimpio = cleanRut(rutObjetivo);
      var sheet     = getSheet('USUARIOS', 'USUARIOS');
      var data      = sheet.getDataRange().getValues();
      var COL       = CONFIG.COLUMNAS.USUARIOS;

      for (var i = 1; i < data.length; i++) {
        if (cleanRut(String(data[i][COL.RUT])) !== rutLimpio) continue;

        var rolAnterior = String(data[i][COL.ROL] || "SOCIO").trim().toUpperCase();
        var nombre      = data[i][COL.NOMBRE];
        var correo      = data[i][COL.CORREO];

        if (rolAnterior === nuevoRolNorm) {
          return { success: false, message: "El usuario ya tiene el rol " + nuevoRolNorm + ". No se realizaron cambios." };
        }

        sheet.getRange(i + 1, COL.ROL + 1).setValue(nuevoRolNorm);
        CacheService.getScriptCache().remove('user_' + rutLimpio);

        Logger.log("🔄 ROL CAMBIADO: " + nombre + " (" + rutLimpio + ") | " + rolAnterior + " → " + nuevoRolNorm + " | Por: " + rutAdmin);

        if (esCorreoValido(correo)) {
          var PERMISOS_ROL = {
            "SOCIO":     ["Módulos de socios: justificaciones, préstamos, apelaciones, permisos médicos", "Registro de asistencia", "SLIM Quest"],
            "DIRIGENTE": ["Todos los módulos de socio", "Consulta ID Credencial", "Gestión de socios: ingresar solicitudes en nombre de terceros", "Panel Dirigente: vista de gestiones realizadas"],
            "ADMIN":     ["Acceso completo al sistema", "Panel Administrador", "Gestión de switches", "Cambio de roles", "Generación de informes"],
            "TESTING":   ["Acceso de pruebas al sistema"]
          };
          var permisos = PERMISOS_ROL[nuevoRolNorm] || [];
          var permisosHtml = permisos.map(function(p) { return "<li style='margin-bottom:4px;'>" + p + "</li>"; }).join("");

          enviarCorreoEstilizado(
            correo,
            "Actualización de Rol - Sindicato SLIM n°3",
            "Tu rol ha sido actualizado",
            "Hola <strong>" + nombre + "</strong>, tu nivel de acceso en la plataforma del sindicato ha sido modificado por la administración.",
            {
              "ROL ANTERIOR": rolAnterior,
              "NUEVO ROL":    nuevoRolNorm,
              "ACCESOS":      "<ul style='margin:0;padding-left:18px;'>" + permisosHtml + "</ul>",
              "MODIFICADO POR": "Administración Sindicato SLIM N°3"
            },
            "#7c3aed"
          );
        }

        return { success: true, message: "Rol cambiado exitosamente de " + rolAnterior + " a " + nuevoRolNorm + ".", rolAnterior: rolAnterior, nuevoRol: nuevoRolNorm, nombre: nombre };
      }

      return { success: false, message: "Usuario no encontrado con RUT " + formatRutDisplay(rutObjetivo) };

    } catch (e) {
      Logger.log("❌ Error en cambiarRolUsuario: " + e.toString());
      return { success: false, message: "Error: " + e.toString() };
    } finally {
      lock.releaseLock();
    }
  } else {
    return { success: false, message: "Servidor ocupado." };
  }
}

// ==========================================
// CONFIGURAR TRIGGERS (ejecutar manualmente UNA VEZ)
// ==========================================

/**
 * Elimina todos los triggers existentes y los recrea.
 * PRECAUCIÓN: Confirmar todos los horarios antes de ejecutar.
 */
function configurarTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) { ScriptApp.deleteTrigger(trigger); });

  // Verificar cambios en justificaciones — cada 8 horas
  ScriptApp.newTrigger('verificarCambiosJustificaciones').timeBased().everyHours(8).create();

  // Verificar cambios en apelaciones — cada 8 horas
  ScriptApp.newTrigger('verificarCambiosApelaciones').timeBased().everyHours(8).create();

  // Procesar validación de préstamos — diario a las 8 AM
  ScriptApp.newTrigger('procesarValidacionPrestamos').timeBased().everyDays(1).atHour(8).create();

  // Procesar permisos de comprobantes de devolución — cada 1 hora
  ScriptApp.newTrigger('procesarPermisosComprobantesDevolucion').timeBased().everyHours(1).create();

  // Verificar cambios en préstamos — diario a las 8 AM
  ScriptApp.newTrigger('verificarCambiosPrestamos').timeBased().everyDays(1).atHour(8).create();

  // Verificar cambios en credenciales — diario a las 8 AM
  ScriptApp.newTrigger('verificarCambiosCredenciales').timeBased().everyDays(1).atHour(8).create();

  // Verificar notificaciones pendientes de asistencia — diario a las 20 hrs
  ScriptApp.newTrigger('verificarNotificacionesAsistencia').timeBased().everyDays(1).atHour(20).create();

  // Reintentar notificaciones socio permisos médicos — cada 30 minutos
  ScriptApp.newTrigger('reintentarNotificacionSocio').timeBased().everyMinutes(30).create();

  // Reintentar notificaciones representante legal — cada 30 minutos
  ScriptApp.newTrigger('reintentarNotificacionRepLegal').timeBased().everyMinutes(30).create();

  Logger.log("✅ Triggers configurados exitosamente");
  Logger.log("Total de triggers activos: " + ScriptApp.getProjectTriggers().length);

  return {
    success: true,
    message: "Triggers configurados correctamente",
    triggers: [
      "verificarCambiosJustificaciones (cada 8 horas)",
      "verificarCambiosApelaciones (cada 8 horas)",
      "procesarValidacionPrestamos (diario 8 AM)",
      "procesarPermisosComprobantesDevolucion (cada 1 hora)",
      "verificarCambiosPrestamos (diario 8 AM)",
      "verificarCambiosCredenciales (diario 8 AM)",
      "verificarNotificacionesAsistencia (diario 20:00)",
      "reintentarNotificacionSocio (cada 30 minutos)",
      "reintentarNotificacionRepLegal (cada 30 minutos)"
    ]
  };
}
