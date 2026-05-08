// ==========================================
// MODULO_DENUNCIAS.GS — Módulo Denuncias Internas v2.1
// Autor: eLj4n0 | Sindicato SLIM N°3
// ==========================================

// ==========================================
// DENUNCIA A JEFATURA — REGISTRO
// ==========================================

/**
 * Registra una denuncia a jefatura en BD_DENUNCIAS-INTERNAS.
 * Roles permitidos: SOCIO, DIRIGENTE, ADMIN.
 * Dirigentes y admins pueden operar a nombre de un socio (rutSocio distinto de rutOperador).
 *
 * @param {string} rutOperador     - RUT del usuario logueado
 * @param {string} rutSocio        - RUT del socio denunciante (puede diferir si es DIRIGENTE/ADMIN)
 * @param {Object} datosFormulario - Campos del formulario
 * @returns {Object} { success, message, idDenuncia }
 */
function registrarDenunciaJefatura(rutOperador, rutSocio, datosFormulario) {
  var lock = LockService.getUserLock();
  if (!lock.tryLock(15000)) {
    return { success: false, message: "El sistema está procesando otra solicitud. Intenta nuevamente." };
  }

  try {
    // 1. Validar rol del operador
    var validacion = verificarRolUsuario(rutOperador, ['SOCIO', 'DIRIGENTE', 'ADMIN']);
    if (!validacion.autorizado) {
      return { success: false, message: "No tienes permisos para realizar esta gestión." };
    }

    var esElevado = (validacion.rol === 'DIRIGENTE' || validacion.rol === 'ADMIN');

    // 2. Si es SOCIO, rutSocio debe coincidir con rutOperador
    var rutSocioLimpio = cleanRut(rutSocio);
    var rutOperLimpio  = cleanRut(rutOperador);

    if (!esElevado && rutSocioLimpio !== rutOperLimpio) {
      return { success: false, message: "No puedes realizar denuncias a nombre de otro socio." };
    }

    // 3. Obtener datos del socio denunciante
    var datosSocio = obtenerUsuarioPorRut(rutSocioLimpio);
    if (!datosSocio.encontrado) {
      return { success: false, message: "No se encontró al socio con RUT " + rutSocio + " en el sistema." };
    }

    // 4. Validar campos obligatorios
    var campos = datosFormulario;
    if (!campos.categoria)        return { success: false, message: "Debes seleccionar una categoría de denuncia." };
    if (!campos.subcategoria)     return { success: false, message: "Debes seleccionar una subcategoría." };
    if (!campos.tipoCargo)        return { success: false, message: "Debes indicar el tipo de cargo del denunciado." };
    if (!campos.nombreDenunciado) return { success: false, message: "Debes ingresar el nombre del denunciado." };
    if (!campos.lugarTrabajo)     return { success: false, message: "Debes indicar el lugar de trabajo." };
    if (!campos.descripcionHechos)return { success: false, message: "Debes describir los hechos." };

    // 5. Validar formato nombre denunciado (al menos 1 nombre + 1 apellido)
    var partesNombre = String(campos.nombreDenunciado).trim().split(/\s+/);
    if (partesNombre.length < 2) {
      return { success: false, message: "El nombre del denunciado debe incluir al menos un nombre y un apellido." };
    }
    if (partesNombre.length > 3) {
      return { success: false, message: "El nombre del denunciado puede tener máximo un nombre y dos apellidos." };
    }

    // 6. Validar restricciones de texto
    var lugarStr = String(campos.lugarTrabajo).trim();
    var palabrasLugar = lugarStr.split(/\s+/);
    if (palabrasLugar.length > 7 || lugarStr.length > 50) {
      return { success: false, message: "El lugar de trabajo no puede superar 7 palabras o 50 caracteres." };
    }

    var descripcionStr = String(campos.descripcionHechos).trim();
    if (descripcionStr.length > 500) {
      return { success: false, message: "La descripción no puede superar 500 caracteres." };
    }

    // 7. Generar ID único — prefijo DJ (Denuncia Jefatura) + timestamp + aleatorio
    var timestamp   = new Date().getTime().toString(36).toUpperCase();
    var aleatorio   = Math.random().toString(36).substring(2, 6).toUpperCase();
    var idDenuncia  = "DJ-" + timestamp + "-" + aleatorio;

    // 8. Marca temporal
    var ahora = new Date();
    var tz    = Session.getScriptTimeZone();
    var fechaFormateada = Utilities.formatDate(ahora, tz, "dd/MM/yyyy HH:mm:ss");

    // 9. Subir archivo adjunto a Drive (si viene)
    var urlAdjunto = "";
    if (campos.archivoBase64 && campos.archivoNombre && campos.archivoMimeType) {
      try {
        var blob = Utilities.newBlob(
          Utilities.base64Decode(campos.archivoBase64),
          campos.archivoMimeType,
          idDenuncia + "_" + campos.archivoNombre
        );
        var carpetaDrive = DriveApp.getFolderById(CONFIG.CARPETAS.DENUNCIAS_JEFATURAS);
        var archivo = carpetaDrive.createFile(blob);
        archivo.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
        urlAdjunto = archivo.getUrl();
      } catch (eDrive) {
        Logger.log("⚠️ Error al subir adjunto denuncia " + idDenuncia + ": " + eDrive.toString());
        // No bloquear el registro si falla el adjunto
        urlAdjunto = "ERROR_ADJUNTO";
      }
    }

    // 10. Escribir en la hoja
    var sheet = getSheet('DENUNCIAS', 'DENUNCIAS_JEFATURAS');
    var COL   = CONFIG.COLUMNAS.DENUNCIAS_JEFATURAS;

    var fila = new Array(12);
    fila[COL.ID_DENUNCIA]        = idDenuncia;
    fila[COL.FECHA_REGISTRO]     = fechaFormateada;
    fila[COL.RUT_DENUNCIANTE]    = datosSocio.rut;
    fila[COL.NOMBRE_DENUNCIANTE] = datosSocio.nombre;
    fila[COL.CATEGORIA]          = String(campos.categoria).trim();
    fila[COL.SUBCATEGORIA]       = String(campos.subcategoria).trim();
    fila[COL.TIPO_DE_CARGO]      = String(campos.tipoCargo).trim();
    fila[COL.NOMBRE_DENUNCIADO]  = String(campos.nombreDenunciado).trim().toUpperCase();
    fila[COL.LUGAR_TRABAJO]      = String(campos.lugarTrabajo).trim().toUpperCase();
    fila[COL.DESCRIPCION_HECHOS] = descripcionStr;
    fila[COL.URL_ARCHIVO_ADJUNTO]= urlAdjunto;
    fila[COL.ESTADO_SOCIO]       = "PENDIENTE";

    sheet.appendRow(fila);
    SpreadsheetApp.flush();

    // 11. Enviar comprobante por correo al socio
    var correoEnviado = false;
    if (datosSocio.correo && esCorreoValido(datosSocio.correo)) {
      try {
        _enviarComprobanteDenuncia(datosSocio, idDenuncia, fechaFormateada, campos, urlAdjunto);

        // Actualizar ESTADO_SOCIO a "Enviado"
        var lastRow = sheet.getLastRow();
        var dataIds = sheet.getRange(2, COL.ID_DENUNCIA + 1, lastRow - 1, 1).getValues();
        for (var r = 0; r < dataIds.length; r++) {
          if (dataIds[r][0] === idDenuncia) {
            sheet.getRange(r + 2, COL.ESTADO_SOCIO + 1).setValue("Enviado");
            break;
          }
        }
        correoEnviado = true;
      } catch (eCorreo) {
        Logger.log("⚠️ Error al enviar correo denuncia " + idDenuncia + ": " + eCorreo.toString());
      }
    }

    Logger.log("✅ DENUNCIA REGISTRADA: " + idDenuncia + " | Socio: " + datosSocio.nombre + " (" + rutSocioLimpio + ") | Operador: " + rutOperLimpio);

    return {
      success:       true,
      message:       "Denuncia registrada exitosamente.",
      idDenuncia:    idDenuncia,
      correoEnviado: correoEnviado,
      nombreSocio:   datosSocio.nombre
    };

  } catch (e) {
    Logger.log("❌ Error en registrarDenunciaJefatura: " + e.toString());
    return { success: false, message: "Error al registrar la denuncia: " + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// CORREO COMPROBANTE — DENUNCIA JEFATURA
// ==========================================

/**
 * Envía correo de comprobante al socio denunciante.
 * @private
 */
function _enviarComprobanteDenuncia(datosSocio, idDenuncia, fechaFormateada, campos, urlAdjunto) {
  var adjuntoInfo = urlAdjunto && urlAdjunto !== "" && urlAdjunto !== "ERROR_ADJUNTO"
    ? "<a href='" + urlAdjunto + "' style='color:#00e84a;'>Ver archivo adjunto</a>"
    : "Sin archivo adjunto";

  var htmlCorreo = "<div style='font-family:Arial,sans-serif;max-width:600px;margin:auto;'>" +
    "<div style='background:#0f172a;padding:24px;border-radius:12px 12px 0 0;text-align:center;'>" +
      "<h1 style='color:#00e84a;font-size:22px;margin:0;'>Sindicato SLIM N°3</h1>" +
      "<p style='color:#94a3b8;font-size:13px;margin:4px 0 0;'>Comprobante de Denuncia Interna</p>" +
    "</div>" +
    "<div style='background:#1e293b;padding:24px;'>" +
      "<p style='color:#e2e8f0;'>Estimado/a <strong>" + datosSocio.nombre + "</strong>,</p>" +
      "<p style='color:#94a3b8;font-size:14px;'>Tu denuncia a jefatura ha sido recibida y registrada por el sindicato. A continuación el detalle:</p>" +
      "<div style='background:#0f172a;border-radius:8px;padding:16px;margin:16px 0;'>" +
        "<table style='width:100%;font-size:13px;border-collapse:collapse;'>" +
          "<tr><td style='color:#64748b;padding:6px 0;'>ID Denuncia</td><td style='color:#00e84a;font-weight:bold;text-align:right;'>" + idDenuncia + "</td></tr>" +
          "<tr><td style='color:#64748b;padding:6px 0;'>Fecha de Registro</td><td style='color:#e2e8f0;text-align:right;'>" + fechaFormateada + "</td></tr>" +
          "<tr><td style='color:#64748b;padding:6px 0;'>Categoría</td><td style='color:#e2e8f0;text-align:right;'>" + campos.categoria + "</td></tr>" +
          "<tr><td style='color:#64748b;padding:6px 0;'>Subcategoría</td><td style='color:#e2e8f0;text-align:right;'>" + campos.subcategoria + "</td></tr>" +
          "<tr><td style='color:#64748b;padding:6px 0;'>Cargo Denunciado</td><td style='color:#e2e8f0;text-align:right;'>" + campos.tipoCargo + "</td></tr>" +
          "<tr><td style='color:#64748b;padding:6px 0;'>Denunciado/a</td><td style='color:#e2e8f0;text-align:right;'>" + String(campos.nombreDenunciado).toUpperCase() + "</td></tr>" +
          "<tr><td style='color:#64748b;padding:6px 0;'>Lugar de Trabajo</td><td style='color:#e2e8f0;text-align:right;'>" + String(campos.lugarTrabajo).toUpperCase() + "</td></tr>" +
          "<tr><td style='color:#64748b;padding:6px 0;'>Archivo Adjunto</td><td style='color:#e2e8f0;text-align:right;'>" + adjuntoInfo + "</td></tr>" +
        "</table>" +
      "</div>" +
      "<div style='background:#0f172a;border-radius:8px;padding:16px;margin:16px 0;'>" +
        "<p style='color:#64748b;font-size:12px;margin:0 0 6px;'>Descripción de los hechos:</p>" +
        "<p style='color:#e2e8f0;font-size:13px;margin:0;'>" + campos.descripcionHechos + "</p>" +
      "</div>" +
      "<p style='color:#64748b;font-size:12px;margin-top:16px;'>Esta denuncia será revisada por la directiva del sindicato. Guarda este correo como respaldo. El ID de denuncia es tu número de seguimiento.</p>" +
    "</div>" +
    "<div style='background:#0f172a;padding:16px;border-radius:0 0 12px 12px;text-align:center;'>" +
      "<p style='color:#475569;font-size:11px;margin:0;'>Sindicato SLIM N°3 — Sistema de Denuncias Internas v2.1</p>" +
    "</div>" +
  "</div>";

  GmailApp.sendEmail(
    datosSocio.correo,
    "Comprobante Denuncia Interna [" + idDenuncia + "] — Sindicato SLIM N°3",
    "Tu denuncia " + idDenuncia + " fue registrada el " + fechaFormateada + ". Revisa el correo en formato HTML para ver el detalle completo.",
    { htmlBody: htmlCorreo, name: "Sindicato SLIM N°3" }
  );
}
