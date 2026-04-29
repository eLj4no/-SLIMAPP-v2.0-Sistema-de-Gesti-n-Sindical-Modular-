// ==========================================
// MODULO_SOCIOS.GS — Autenticación, datos personales, bancarios, vestuario, credencial
// ==========================================

// ==========================================
// ACCESO Y AUTENTICACIÓN
// ==========================================

/**
 * Valida usuario (Login)
 */
function validarUsuario(rutInput, passwordInput) {
  try {
    var sheet = getSheet('USUARIOS', 'USUARIOS');
    var data = sheet.getDataRange().getDisplayValues();
    var rutLimpioInput = cleanRut(rutInput);
    var COL = CONFIG.COLUMNAS.USUARIOS;

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (cleanRut(row[COL.RUT]) === rutLimpioInput) {
        var passDb = String(row[COL.ID_CREDENCIAL]);
        var nombreUsuario = row[COL.NOMBRE];
        var rolUsuario = String(row[COL.ROL]).trim().toUpperCase();
        var estadoUsuario = String(row[COL.ESTADO]).toUpperCase();

        if (String(passDb).toUpperCase() === String(passwordInput).toUpperCase()) {
          return {
            success: true,
            message: "Login exitoso",
            user: nombreUsuario || "Socio",
            role: rolUsuario || "SOCIO",
            state: estadoUsuario || "ACTIVO",
            estadoNegColect: String(row[COL.ESTADO_NEG_COLECT] || "").trim()
          };
        } else {
          return { success: false, message: "Contraseña incorrecta", errorType: "password" };
        }
      }
    }
    return { success: false, message: "RUT no encontrado", errorType: "rut" };
  } catch (e) {
    Logger.log('ERROR en validarUsuario: ' + e.toString());
    return { success: false, message: "Error Servidor: " + e.toString() };
  }
}

/**
 * Obtener datos completos del usuario
 */
function obtenerDatosUsuario(rutInput) {
  try {
    var sheet = getSheet('USUARIOS', 'USUARIOS');
    var data = sheet.getDataRange().getDisplayValues();
    var rutLimpioInput = cleanRut(rutInput);
    var COL = CONFIG.COLUMNAS.USUARIOS;

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (cleanRut(row[COL.RUT]) === rutLimpioInput) {
        return {
          success: true,
          datos: {
            rut:                row[COL.RUT]             || "---",
            nombre:             row[COL.NOMBRE]          || "Sin Nombre",
            cargo:              row[COL.CARGO]           || "---",
            site:               row[COL.SITE]            || "---",
            region:             row[COL.REGION],
            estado:             String(row[COL.ESTADO]).toUpperCase(),
            correo:             row[COL.CORREO],
            contacto:           row[COL.CONTACTO],
            estadoNegColect:    row[COL.ESTADO_NEG_COLECT]    || "",
            banco:              row[COL.BANCO]                || "",
            tipoCuenta:         row[COL.TIPO_CUENTA]          || "",
            numeroCuenta:       row[COL.NUMERO_CUENTA]        || "",
            tallaPolera:        row[COL.TALLA_POLERA]         || "",
            tallaPolar:         row[COL.TALLA_POLAR]          || "",
            tallaPantalon:      row[COL.TALLA_PANTALON]       || "",
            tallaCalzado:       row[COL.TALLA_CALZADO]        || "",
            calzadoEspecial:    row[COL.CALZADO_ESPECIAL]     || "NO",
            urlCertPieDiabetico:row[COL.URL_CERT_PIE_DIABETICO] || "",
            estadoCredencial:   obtenerEstadoCredencialPorRut(row[COL.RUT])
          }
        };
      }
    }
    return { success: false, message: "Datos no encontrados." };
  } catch (e) {
    return { success: false, message: "Error Datos: " + e.toString() };
  }
}

/**
 * Obtener datos de usuario por RUT — Función auxiliar centralizada con caché
 */
function obtenerUsuarioPorRut(rutInput) {
  var cache = CacheService.getScriptCache();
  var rutLimpio = cleanRut(rutInput);
  var cacheKey = 'user_' + rutLimpio;

  var cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { Logger.log('Error parsing cache: ' + e); }
  }

  var sheet = getSheet('USUARIOS', 'USUARIOS');
  var COL = CONFIG.COLUMNAS.USUARIOS;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { encontrado: false };

  var data = sheet.getRange(2, 1, lastRow - 1, COL.ESTADO_NEG_COLECT + 1).getDisplayValues();

  for (var i = 0; i < data.length; i++) {
    if (cleanRut(data[i][COL.RUT]) === rutLimpio) {
      var usuario = {
        encontrado:     true,
        rut:            data[i][COL.RUT],
        nombre:         data[i][COL.NOMBRE],
        correo:         data[i][COL.CORREO],
        region:         data[i][COL.REGION],
        cargo:          data[i][COL.CARGO],
        site:           data[i][COL.SITE],
        estado:         data[i][COL.ESTADO],
        rol:            data[i][COL.ROL],
        contacto:       data[i][COL.CONTACTO],
        estadoNegColect:data[i][COL.ESTADO_NEG_COLECT] || "",
        banco:          data[i][COL.BANCO]          || "",
        tipoCuenta:     data[i][COL.TIPO_CUENTA]    || "",
        numeroCuenta:   data[i][COL.NUMERO_CUENTA]  || ""
      };
      try { cache.put(cacheKey, JSON.stringify(usuario), 600); } catch (e) {}
      return usuario;
    }
  }
  return { encontrado: false };
}

// ==========================================
// RECUPERACIÓN DE CONTRASEÑA
// ==========================================

function recuperarContrasena(rutInput) {
  try {
    var sheet = getSheet('USUARIOS', 'USUARIOS');
    var data = sheet.getDataRange().getDisplayValues();
    var rutLimpio = cleanRut(rutInput);
    var COL = CONFIG.COLUMNAS.USUARIOS;

    for (var i = 1; i < data.length; i++) {
      if (cleanRut(data[i][COL.RUT]) === rutLimpio) {
        return { success: true, correo: data[i][COL.CORREO] || "No registrado" };
      }
    }
    return { success: false, message: "Usuario no encontrado." };
  } catch (e) {
    return { success: false, message: "Error: " + e.toString() };
  }
}

function enviarContrasenaCorreo(rutInput) {
  try {
    var sheet = getSheet('USUARIOS', 'USUARIOS');
    var data = sheet.getDataRange().getDisplayValues();
    var rutLimpio = cleanRut(rutInput);
    var COL = CONFIG.COLUMNAS.USUARIOS;

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (cleanRut(row[COL.RUT]) === rutLimpio) {
        var nombre   = row[COL.NOMBRE];
        var correo   = row[COL.CORREO];
        var password = row[COL.ID_CREDENCIAL];

        if (!correo || !correo.includes("@")) {
          return { success: false, message: "No tienes un correo registrado. Contacta con la directiva." };
        }

        enviarCorreoEstilizado(
          correo,
          "Recuperación de Contraseña - Sindicato SLIM n°3",
          "Recuperación de Contraseña",
          "Hola " + nombre + ", has solicitado recuperar tu contraseña de acceso al portal.",
          { "Tu contraseña es": password, "RUT": row[COL.RUT] },
          "#3b82f6"
        );

        return { success: true, message: "Contraseña enviada exitosamente." };
      }
    }
    return { success: false, message: "Usuario no encontrado." };
  } catch (e) {
    return { success: false, message: "Error: " + e.toString() };
  }
}

// ==========================================
// ACTUALIZACIÓN DE DATOS PERSONALES
// ==========================================

/**
 * Actualiza un campo individual del usuario
 */
function actualizarDatoUsuario(rutInput, campo, valor) {
  var lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) {
    try {
      var sheet = getSheet('USUARIOS', 'USUARIOS');
      var data = sheet.getDataRange().getValues();
      var rutLimpioInput = cleanRut(rutInput);
      var COL = CONFIG.COLUMNAS.USUARIOS;

      var colIndex = -1;
      if (campo === 'region')       colIndex = COL.REGION;
      else if (campo === 'correo')  colIndex = COL.CORREO;
      else if (campo === 'contacto')colIndex = COL.CONTACTO;
      else if (campo === 'banco')   colIndex = COL.BANCO;
      else if (campo === 'tipoCuenta')  colIndex = COL.TIPO_CUENTA;
      else if (campo === 'numeroCuenta')colIndex = COL.NUMERO_CUENTA;

      if (colIndex === -1) return { success: false, message: "Campo inválido" };

      for (var i = 1; i < data.length; i++) {
        if (cleanRut(String(data[i][COL.RUT])) === rutLimpioInput) {
          sheet.getRange(i + 1, colIndex + 1).setValue(valor);
          CacheService.getScriptCache().remove('user_' + rutLimpioInput);
          return { success: true, message: "OK" };
        }
      }
      return { success: false, message: "Usuario no hallado para editar" };
    } catch (e) {
      return { success: false, message: "Error Update: " + e.toString() };
    } finally {
      lock.releaseLock();
    }
  } else {
    return { success: false, message: "Servidor ocupado." };
  }
}

/**
 * Actualiza los 3 campos bancarios para Cuenta RUT de Banco Estado
 */
function actualizarBancoEstado(rutInput) {
  var lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) {
    try {
      var sheet = getSheet('USUARIOS', 'USUARIOS');
      var data = sheet.getDataRange().getValues();
      var rutLimpioInput = cleanRut(rutInput);
      var COL = CONFIG.COLUMNAS.USUARIOS;
      var rutBody = rutLimpioInput.slice(0, -1);

      for (var i = 1; i < data.length; i++) {
        if (cleanRut(String(data[i][COL.RUT])) === rutLimpioInput) {
          sheet.getRange(i + 1, COL.BANCO + 1).setValue("BANCO ESTADO (Cuenta RUT)");
          sheet.getRange(i + 1, COL.TIPO_CUENTA + 1).setValue("CUENTA VISTA");
          sheet.getRange(i + 1, COL.NUMERO_CUENTA + 1).setValue(rutBody);
          CacheService.getScriptCache().remove('user_' + rutLimpioInput);
          return { success: true, numeroCuenta: rutBody, tipoCuenta: "CUENTA VISTA" };
        }
      }
      return { success: false, message: "Usuario no encontrado." };
    } catch (e) {
      return { success: false, message: "Error al actualizar: " + e.toString() };
    } finally {
      lock.releaseLock();
    }
  } else {
    return { success: false, message: "Servidor ocupado. Intente nuevamente." };
  }
}

/**
 * Guarda los 3 campos bancarios en una sola operación atómica
 */
function actualizarDatosBancarios(rutInput, banco, tipoCuenta, numeroCuenta) {
  var lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) {
    try {
      var sheet = getSheet('USUARIOS', 'USUARIOS');
      var data = sheet.getDataRange().getValues();
      var rutLimpioInput = cleanRut(rutInput);
      var COL = CONFIG.COLUMNAS.USUARIOS;

      for (var i = 1; i < data.length; i++) {
        if (cleanRut(String(data[i][COL.RUT])) === rutLimpioInput) {
          sheet.getRange(i + 1, COL.BANCO + 1).setValue(banco);
          sheet.getRange(i + 1, COL.TIPO_CUENTA + 1).setValue(tipoCuenta);
          sheet.getRange(i + 1, COL.NUMERO_CUENTA + 1).setValue(numeroCuenta);
          CacheService.getScriptCache().remove('user_' + rutLimpioInput);
          return { success: true };
        }
      }
      return { success: false, message: "Usuario no encontrado." };
    } catch (e) {
      return { success: false, message: "Error al actualizar: " + e.toString() };
    } finally {
      lock.releaseLock();
    }
  } else {
    return { success: false, message: "Servidor ocupado. Intente nuevamente." };
  }
}

/**
 * Guarda datos de vestuario del usuario (incluye upload de certificado pie diabético)
 */
function guardarDatosVestuario(rutInput, datosVestuario) {
  var lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) {
    try {
      var sheet = getSheet('USUARIOS', 'USUARIOS');
      var data = sheet.getDataRange().getValues();
      var rutLimpioInput = cleanRut(rutInput);
      var COL = CONFIG.COLUMNAS.USUARIOS;

      var tallasValidas  = ['XS','S','M','L','XL','XXL','3XL','4XL','5XL'];
      var numerosValidos = ['32','34','36','38','40','42','44','46','48','50','52','54','56','58','60','62','64','66','67'];

      if (datosVestuario.tallaPolera && !tallasValidas.includes(datosVestuario.tallaPolera)) {
        return { success: false, message: "Talla Polera/Camisa inválida." };
      }
      if (datosVestuario.tallaPolar && !tallasValidas.includes(datosVestuario.tallaPolar)) {
        return { success: false, message: "Talla Polar/Chaqueta inválida." };
      }
      if (datosVestuario.tallaPantalon && !numerosValidos.includes(String(datosVestuario.tallaPantalon))) {
        return { success: false, message: "Talla Pantalón inválida." };
      }
      if (datosVestuario.tallaCalzado && !['32','33','34','35','36','37','38','39','40','41','42','43','44','45','46','47','48','50','52'].includes(String(datosVestuario.tallaCalzado))) {
        return { success: false, message: "Talla Calzado inválida." };
      }

      var urlCert = "";
      var calzadoEsp = String(datosVestuario.calzadoEspecial || "").toUpperCase() === "SI";

      if (calzadoEsp && datosVestuario.archivo && datosVestuario.archivo.base64) {
        var archivo = datosVestuario.archivo;
        var tiposPermitidos = ['image/jpeg','image/png','image/gif','image/webp','application/pdf',
          'application/msword','application/vnd.openxmlformats-officedocument.wordprocessingml.document'];

        if (!tiposPermitidos.includes(archivo.mimeType)) {
          return { success: false, message: "Tipo de archivo no permitido. Solo se aceptan imágenes, PDF o documentos Word." };
        }
        var sizeInBytes = Math.ceil((archivo.base64.length * 3) / 4);
        if (sizeInBytes > 15 * 1024 * 1024) {
          return { success: false, message: "El archivo excede el tamaño máximo permitido de 15 MB." };
        }
        try {
          var folder = DriveApp.getFolderById(CONFIG.CARPETAS.VESTUARIO_DOCS);
          var blob = Utilities.newBlob(
            Utilities.base64Decode(archivo.base64),
            archivo.mimeType,
            'CertPieDiabetico_' + rutLimpioInput + '_' + Utilities.formatDate(new Date(), 'America/Santiago', 'yyyyMMdd')
          );
          var file = folder.createFile(blob);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          urlCert = file.getUrl();
        } catch (driveErr) {
          Logger.log('❌ Error subiendo certificado pie diabético: ' + driveErr.toString());
          return { success: false, message: "Error al subir el archivo. Intenta nuevamente." };
        }
      } else if (calzadoEsp && !datosVestuario.archivo) {
        urlCert = datosVestuario.urlActual || "";
      } else {
        urlCert = "";
      }

      for (var i = 1; i < data.length; i++) {
        if (cleanRut(String(data[i][COL.RUT])) === rutLimpioInput) {
          var fila = i + 1;
          sheet.getRange(fila, COL.TALLA_POLERA + 1).setValue(datosVestuario.tallaPolera || "");
          sheet.getRange(fila, COL.TALLA_POLAR + 1).setValue(datosVestuario.tallaPolar || "");
          sheet.getRange(fila, COL.TALLA_PANTALON + 1).setValue(datosVestuario.tallaPantalon || "");
          sheet.getRange(fila, COL.TALLA_CALZADO + 1).setValue(datosVestuario.tallaCalzado || "");
          sheet.getRange(fila, COL.CALZADO_ESPECIAL + 1).setValue(calzadoEsp ? "SI" : "NO");
          sheet.getRange(fila, COL.URL_CERT_PIE_DIABETICO + 1).setValue(urlCert);
          CacheService.getScriptCache().remove('user_' + rutLimpioInput);
          return { success: true, message: "Datos de vestuario guardados correctamente.", urlCert: urlCert };
        }
      }
      return { success: false, message: "Usuario no encontrado en el sistema." };

    } catch (e) {
      Logger.log('❌ Error en guardarDatosVestuario: ' + e.toString());
      return { success: false, message: "Error del servidor: " + e.toString() };
    } finally {
      lock.releaseLock();
    }
  } else {
    return { success: false, message: "Servidor ocupado. Intenta nuevamente en unos segundos." };
  }
}

// ==========================================
// OBTENER CORREO POR RUT (auxiliar compartida)
// ==========================================

function obtenerCorreoDeRut(rut) {
  try {
    var sheet = getSheet('USUARIOS', 'USUARIOS');
    var data = sheet.getDataRange().getDisplayValues();
    var rutLimpio = cleanRut(rut);
    var COL = CONFIG.COLUMNAS.USUARIOS;
    for (var i = 1; i < data.length; i++) {
      if (cleanRut(data[i][COL.RUT]) === rutLimpio) return data[i][COL.CORREO];
    }
    return "";
  } catch (e) {
    console.error("Error obteniendo correo: " + e);
    return "";
  }
}

// ==========================================
// VALIDACIÓN QR (para QR_Access y QR_Asistencia)
// ==========================================

function validarUsuarioQR(rutInput) {
  try {
    var sheet = getSheet('USUARIOS', 'USUARIOS');
    var data = sheet.getDataRange().getDisplayValues();
    var rutLimpio = cleanRut(rutInput);
    var COL = CONFIG.COLUMNAS.USUARIOS;

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (cleanRut(row[COL.RUT]) === rutLimpio) {
        var estadoUsuario = String(row[COL.ESTADO]).toUpperCase();
        if (estadoUsuario !== "ACTIVO" && estadoUsuario !== "SI" && estadoUsuario !== "TRUE") {
          return { success: false, error: "Usuario desvinculado. Contacta con la directiva." };
        }
        return { success: true, nombre: row[COL.NOMBRE] || "Socio", rut: row[COL.RUT] };
      }
    }
    return { success: false, error: "RUT no encontrado en el sistema." };
  } catch (e) {
    return { success: false, error: "Error del servidor: " + e.toString() };
  }
}

// ==========================================
// GENERADOR DE LINKS DE REGISTRO Y CÓDIGOS QR
// ==========================================

/**
 * Genera el Link Registro (col P) y el código QR (col Q) para todos los usuarios
 * que aún no tienen esos datos. Ejecutar manualmente.
 */
function generarLinksRegistroYQR() {
  try {
    Logger.log('🚀 Iniciando generarLinksRegistroYQR...');
    Logger.log('🔗 URL base: ' + WEBAPP_BASE_URL);

    var sheet = getSheet('USUARIOS', 'USUARIOS');
    if (!sheet) { Logger.log('❌ No se pudo acceder a la hoja de usuarios.'); return; }

    var COL      = CONFIG.COLUMNAS.USUARIOS;
    var lastRow  = sheet.getLastRow();
    if (lastRow < 2) { Logger.log('⚠️ No hay usuarios en la hoja.'); return; }

    var totalFilas     = lastRow - 1;
    var rangoRut       = sheet.getRange(2, COL.RUT + 1,           totalFilas, 1).getValues();
    var rangoLinkActual= sheet.getRange(2, COL.LINK_REGISTRO + 1, totalFilas, 1).getValues();
    var rangoQRFormulas= sheet.getRange(2, COL.QR_REGISTRO + 1,   totalFilas, 1).getFormulas();

    var generadosLink = 0, generadosQR = 0, omitidosLink = 0, omitidosQR = 0, sinRut = 0;

    for (var i = 0; i < totalFilas; i++) {
      var rutRaw     = String(rangoRut[i][0]).trim();
      var linkActual = String(rangoLinkActual[i][0]).trim();
      var qrFormula  = String(rangoQRFormulas[i][0]).trim();
      var filaSheet  = i + 2;

      if (!rutRaw || rutRaw === '' || rutRaw === '0' || rutRaw.toLowerCase() === 'false') { sinRut++; continue; }
      var rutLimpio = cleanRut(rutRaw);
      if (!rutLimpio || rutLimpio.length < 7) { Logger.log('⚠️ Fila ' + filaSheet + ': RUT inválido → ' + rutRaw); sinRut++; continue; }

      var tieneLink = linkActual !== '' && linkActual !== '0' && linkActual.toLowerCase() !== 'false';
      if (tieneLink) {
        omitidosLink++;
      } else {
        var linkRegistro = WEBAPP_BASE_URL + '?action=register&rut=' + rutLimpio;
        sheet.getRange(filaSheet, COL.LINK_REGISTRO + 1).setValue(linkRegistro);
        Logger.log('✅ Link generado | Fila ' + filaSheet + ' | RUT: ' + rutLimpio);
        generadosLink++;
      }

      var tieneQR = qrFormula.toUpperCase().includes('IMAGE');
      if (tieneQR) {
        omitidosQR++;
      } else {
        var linkParaQR = tieneLink ? linkActual : WEBAPP_BASE_URL + '?action=register&rut=' + rutLimpio;
        var linkEncoded = encodeURIComponent(linkParaQR);
        var formulaQR = '=IMAGE("https://quickchart.io/qr?size=300&text=' + linkEncoded + '")';
        sheet.getRange(filaSheet, COL.QR_REGISTRO + 1).setFormula(formulaQR);
        Logger.log('✅ QR generado | Fila ' + filaSheet + ' | RUT: ' + rutLimpio);
        generadosQR++;
      }

      if ((generadosLink + generadosQR) % 100 === 0 && (generadosLink + generadosQR) > 0) Utilities.sleep(300);
    }

    Logger.log('📊 RESUMEN — Links: ' + generadosLink + ' | QR: ' + generadosQR + ' | Existentes link: ' + omitidosLink + ' | Existentes QR: ' + omitidosQR + ' | Sin RUT: ' + sinRut);

  } catch (e) {
    Logger.log('❌ Error en generarLinksRegistroYQR: ' + e.toString());
    throw e;
  }
}

// ==========================================
// MÓDULO: CREDENCIAL SINDICAL
// ==========================================

/**
 * Obtiene el estado de credencial de un usuario desde BD_CREDENCIALES
 */
function obtenerEstadoCredencialPorRut(rutInput) {
  try {
    var rutLimpio = cleanRut(String(rutInput));
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEETS.CREDENCIALES);
    var sheet = ss.getSheetByName(CONFIG.HOJAS.CREDENCIALES);
    if (!sheet) { Logger.log('⚠️ Hoja CREDENCIALES no encontrada'); return "S/D"; }
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return "S/D";
    var data = sheet.getRange(2, 1, lastRow - 1, 7).getDisplayValues();
    for (var i = 0; i < data.length; i++) {
      if (cleanRut(String(data[i][0])) === rutLimpio) {
        return String(data[i][6] || "").trim().toUpperCase() || "S/D";
      }
    }
    return "S/D";
  } catch (e) {
    Logger.log('❌ Error obteniendo estado credencial: ' + e.toString());
    return "S/D";
  }
}

/**
 * Trigger diario a las 8 AM: detecta cambios de estado en credenciales y envía notificaciones
 */
function verificarCambiosCredenciales() {
  var ESTADOS_CON_NOTIFICACION = ["ENTREGADO","DISPONIBLE","SOLICITADO","NO VIGENTE","DATOS INCORRECTOS","REIMPRIMIR"];
  try {
    Logger.log('🔄 Iniciando verificación de cambios en credenciales...');
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEETS.CREDENCIALES);
    var sheetImpresion = ss.getSheetByName(CONFIG.HOJAS.CREDENCIALES);
    if (!sheetImpresion) { Logger.log('❌ No se encontró la hoja IMPRESION en BD_CREDENCIALES'); return; }

    var sheetHistorial = ss.getSheetByName(CONFIG.HOJAS.HISTORIAL_CREDENCIALES);
    if (!sheetHistorial) {
      sheetHistorial = ss.insertSheet(CONFIG.HOJAS.HISTORIAL_CREDENCIALES);
      sheetHistorial.appendRow(['FECHA','RUT','NOMBRE','CORREO','ESTADO ANTERIOR','ESTADO NUEVO','EMAIL ENVIADO']);
      sheetHistorial.getRange(1,1,1,7).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
    }

    var lastRow = sheetImpresion.getLastRow();
    if (lastRow < 2) { Logger.log('ℹ️ No hay datos en la hoja IMPRESION'); return; }

    var data = sheetImpresion.getRange(2, 1, lastRow - 1, 11).getDisplayValues();
    var enviados = 0, errores = 0, inicializados = 0, sinCambios = 0;

    for (var i = 0; i < data.length; i++) {
      var fila = i + 2;
      var rut = String(data[i][0] || "").trim();
      var nombre = String(data[i][3] || data[i][4] || "Socio").trim();
      var correo = String(data[i][5] || "").trim();
      var estadoActual = String(data[i][6] || "").trim().toUpperCase();
      var estadoAnterior = String(data[i][9] || "").trim().toUpperCase();

      if (!rut || !estadoActual) continue;

      if (!estadoAnterior) {
        sheetImpresion.getRange(fila, 10).setValue(estadoActual);
        inicializados++;
        Logger.log('ℹ️ Fila ' + fila + ' (' + rut + '): Inicializado con estado "' + estadoActual + '"');
        continue;
      }

      if (estadoActual === estadoAnterior) { sinCambios++; continue; }

      Logger.log('🔔 Fila ' + fila + ' (' + rut + '): Cambio detectado "' + estadoAnterior + '" → "' + estadoActual + '"');
      var emailEstado = "SIN CORREO";

      if (ESTADOS_CON_NOTIFICACION.indexOf(estadoActual) !== -1 && correo.includes('@')) {
        try {
          enviarNotificacionCredencial(correo, nombre, estadoActual, rut);
          emailEstado = "ENVIADO";
          enviados++;
        } catch (emailErr) {
          emailEstado = "ERROR: " + emailErr.toString().substring(0, 80);
          errores++;
          Logger.log('❌ Error enviando email a ' + correo + ': ' + emailErr.toString());
        }
      } else if (!correo.includes('@')) {
        emailEstado = "SIN CORREO VÁLIDO";
      } else {
        emailEstado = "ESTADO SIN NOTIF.";
      }

      var fechaAhora = Utilities.formatDate(new Date(), 'America/Santiago', 'dd/MM/yyyy HH:mm:ss');
      sheetHistorial.appendRow([fechaAhora, rut, nombre, correo, estadoAnterior, estadoActual, emailEstado]);
      sheetImpresion.getRange(fila, 10).setValue(estadoActual);
      sheetImpresion.getRange(fila, 11).setValue(emailEstado);
    }

    Logger.log('✅ Verificación completada: Inicializados=' + inicializados + ' | Sin cambios=' + sinCambios + ' | Emails enviados=' + enviados + ' | Errores=' + errores);
  } catch (e) {
    Logger.log('❌ Error crítico en verificarCambiosCredenciales: ' + e.toString());
  }
}

/**
 * Envía notificación HTML sobre el estado de credencial
 */
function enviarNotificacionCredencial(correo, nombre, estadoNuevo, rut) {
  var MENSAJES = {
    "ENTREGADO":         { titulo:"¡Tu credencial sindical ha sido entregada!",      icono:"🎉", color:"#059669", colorClaro:"#d1fae5", colorBorde:"#6ee7b7", mensaje:"Tu tarjeta ha sido entregada y se encuentra disponible para su uso. Te recordamos que la credencial sindical se utiliza para las asambleas presenciales, la cual debes mostrar para la correcta marcación de tu asistencia.", nota:"Si la credencial está desgastada o la extraviaste, debes solicitar una nueva a un dirigente de la organización." },
    "DISPONIBLE":        { titulo:"¡Tu credencial sindical está lista para retiro!", icono:"📦", color:"#0d9488", colorClaro:"#ccfbf1", colorBorde:"#5eead4", mensaje:"Tu credencial sindical ya está impresa y disponible para su retiro. Debes acercarte a la oficina sindical o retirarla en la próxima asamblea presencial.", nota:"Recuerda llevar tu RUT al momento de retirarla." },
    "SOLICITADO":        { titulo:"Solicitud de credencial recibida",                icono:"⏳", color:"#d97706", colorClaro:"#fef3c7", colorBorde:"#fcd34d", mensaje:"Tus datos han sido ingresados al sistema y se ha solicitado al departamento de comunicaciones la gestión para la impresión de tu credencial sindical.", nota:"Cuando esté disponible, recibirás un correo electrónico notificándote su disponibilidad de retiro." },
    "NO VIGENTE":        { titulo:"Credencial sindical deshabilitada",               icono:"⚠️", color:"#dc2626", colorClaro:"#fee2e2", colorBorde:"#fca5a5", mensaje:"De acuerdo a nuestros registros, tu credencial sindical ha sido deshabilitada, lo que indica que podrías ya no pertenecer a la empresa o a la organización sindical.", nota:"Si consideras que existe un error, comunícate con algún dirigente para regularizar tu situación." },
    "DATOS INCORRECTOS": { titulo:"Atención: datos incorrectos para tu credencial", icono:"❗", color:"#ea580c", colorClaro:"#ffedd5", colorBorde:"#fdba74", mensaje:"No se ha podido crear tu credencial sindical debido a que existen datos incorrectos para su fabricación.", nota:"Debes actualizar tus datos en el módulo \"Mis Datos\" de la aplicación sindical, o comunicarte con un dirigente." },
    "REIMPRIMIR":        { titulo:"Solicitud de reimpresión recibida",               icono:"🖨️", color:"#7c3aed", colorClaro:"#ede9fe", colorBorde:"#c4b5fd", mensaje:"Hemos recibido una solicitud de reimpresión de tu credencial sindical. Nuestro equipo está procesando esta solicitud.", nota:"Una vez que esté disponible, se te notificará a través del correo electrónico registrado." }
  };

  var info = MENSAJES[estadoNuevo] || { titulo:"Actualización de credencial sindical", icono:"📋", color:"#64748b", colorClaro:"#f1f5f9", colorBorde:"#cbd5e1", mensaje:"El estado de tu credencial sindical ha sido actualizado.", nota:"Puedes revisar el estado actual en el módulo 'Mis Datos' de la aplicación sindical." };
  var fechaActual = Utilities.formatDate(new Date(), 'America/Santiago', "dd 'de' MMMM 'de' yyyy");

  var htmlBody = '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
    '<body style="margin:0;padding:0;background-color:#0f172a;font-family:\'Helvetica Neue\',Arial,sans-serif;">' +
    '<div style="max-width:600px;margin:0 auto;padding:20px;">' +
    '<div style="background:linear-gradient(135deg,' + info.color + ',' + info.color + 'dd);border-radius:16px 16px 0 0;padding:32px 24px;text-align:center;">' +
    '<div style="font-size:48px;margin-bottom:12px;">' + info.icono + '</div>' +
    '<h1 style="color:#ffffff;font-size:20px;font-weight:700;margin:0 0 8px 0;line-height:1.3;">' + info.titulo + '</h1>' +
    '<p style="color:rgba(255,255,255,0.85);font-size:13px;margin:0;">Sindicato SLIM N°3 · Credencial Sindical</p></div>' +
    '<div style="background:#ffffff;padding:28px 24px;">' +
    '<p style="color:#374151;font-size:15px;margin:0 0 20px 0;">Estimado(a) <strong>' + nombre + '</strong>,</p>' +
    '<div style="background:' + info.colorClaro + ';border:1px solid ' + info.colorBorde + ';border-radius:12px;padding:16px;text-align:center;margin-bottom:20px;">' +
    '<p style="color:#6b7280;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:1px;margin:0 0 8px 0;">NUEVO ESTADO DE CREDENCIAL</p>' +
    '<span style="display:inline-block;background:' + info.color + ';color:#ffffff;font-size:14px;font-weight:800;text-transform:uppercase;letter-spacing:1px;padding:8px 20px;border-radius:50px;">' + estadoNuevo + '</span></div>' +
    '<div style="background:#f8fafc;border-left:4px solid ' + info.color + ';border-radius:0 8px 8px 0;padding:16px;margin-bottom:16px;">' +
    '<p style="color:#374151;font-size:14px;line-height:1.6;margin:0;">' + info.mensaje + '</p></div>' +
    '<div style="background:#fffbeb;border:1px solid #fde68a;border-radius:8px;padding:14px;margin-bottom:20px;">' +
    '<p style="color:#92400e;font-size:12px;line-height:1.6;margin:0;"><strong>📌 Nota:</strong> ' + info.nota + '</p></div>' +
    '<table style="width:100%;border-collapse:collapse;margin-bottom:20px;font-size:13px;">' +
    '<tr style="background:#f1f5f9;"><td style="padding:10px 12px;font-weight:700;color:#475569;border-bottom:1px solid #e2e8f0;width:40%;">Nombre</td><td style="padding:10px 12px;color:#1e293b;border-bottom:1px solid #e2e8f0;">' + nombre + '</td></tr>' +
    '<tr><td style="padding:10px 12px;font-weight:700;color:#475569;border-bottom:1px solid #e2e8f0;">Fecha</td><td style="padding:10px 12px;color:#1e293b;border-bottom:1px solid #e2e8f0;">' + fechaActual + '</td></tr>' +
    '</table></div>' +
    '<div style="background:#1e293b;border-radius:0 0 16px 16px;padding:20px 24px;text-align:center;">' +
    '<p style="color:#94a3b8;font-size:12px;margin:0 0 4px 0;">Este es un mensaje automático del sistema de gestión</p>' +
    '<p style="color:#64748b;font-size:11px;margin:0;">Sindicato SLIM N°3 · No responder a este correo</p></div>' +
    '</div></body></html>';

  MailApp.sendEmail({
    to: correo,
    subject: info.icono + ' Credencial Sindical: ' + estadoNuevo + ' - Sindicato SLIM N°3',
    htmlBody: htmlBody,
    name: "Sindicato SLIM N°3"
  });
}

// ==========================================
// CONSULTA ID CREDENCIAL (para DIRIGENTE/ADMIN)
// ==========================================

/**
 * Consulta el ID Credencial de un usuario por RUT
 * Solo accesible para roles DIRIGENTE y ADMIN
 */
function consultarIdCredencialBackend(rutConsultante, rutBuscado) {
  try {
    var validacion = verificarRolUsuario(rutConsultante, ['DIRIGENTE', 'ADMIN']);
    if (!validacion.autorizado) {
      return { success: false, message: 'No tienes permisos para realizar esta consulta.' };
    }

    var rutLimpio = cleanRut(rutBuscado);
    if (!rutLimpio || rutLimpio.length < 7) {
      return { success: false, message: 'RUT inválido o incompleto.' };
    }

    var sheet = getSheet('USUARIOS', 'USUARIOS');
    if (!sheet) return { success: false, message: 'Error al acceder a la base de datos.' };

    var COL = CONFIG.COLUMNAS.USUARIOS;
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: 'No hay usuarios registrados en el sistema.' };

    var data = sheet.getRange(2, 1, lastRow - 1, COL.ESTADO_NEG_COLECT + 1).getDisplayValues();
    var formulasQR = sheet.getRange(2, COL.QR_REGISTRO + 1, lastRow - 1, 1).getFormulas();

    for (var i = 0; i < data.length; i++) {
      if (cleanRut(data[i][COL.RUT]) === rutLimpio) {
        var rolUsuarioBuscado = String(data[i][COL.ROL] || 'SOCIO').trim().toUpperCase();

        if (validacion.rol === 'DIRIGENTE' && rolUsuarioBuscado !== 'SOCIO') {
          return { success: false, message: 'Acceso restringido: Solo puedes consultar información de usuarios con rol SOCIO.', restricted: true };
        }

        var urlQR = extraerUrlDeImagen(formulasQR[i][0]);
        return {
          success: true,
          rut:             data[i][COL.RUT],
          nombre:          data[i][COL.NOMBRE],
          cargo:           data[i][COL.CARGO],
          estado:          data[i][COL.ESTADO],
          rol:             rolUsuarioBuscado,
          idCredencial:    data[i][COL.ID_CREDENCIAL] || 'S/D',
          qrRegistro:      urlQR,
          estadoCredencial:obtenerEstadoCredencialPorRut(data[i][COL.RUT])
        };
      }
    }

    return { success: false, message: 'No se encontró ningún usuario con el RUT ' + formatRutDisplay(rutBuscado) + ' en el sistema.' };

  } catch (e) {
    Logger.log('❌ ERROR en consultarIdCredencialBackend: ' + e.toString());
    return { success: false, message: 'Error inesperado: ' + e.toString() };
  }
}

// ==========================================
// MIGRACIÓN: COMPLETAR CAMPOS BANCARIOS EN BLANCO
// (Ejecutar manualmente una sola vez)
// ==========================================

function completarCamposBancariosEnBlanco() {
  try {
    var sheet = getSheet('USUARIOS', 'USUARIOS');
    if (!sheet) { Logger.log('❌ No se pudo acceder a la hoja de usuarios.'); return; }

    var COL = CONFIG.COLUMNAS.USUARIOS;
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) { Logger.log('⚠️ No hay registros en la hoja.'); return; }

    var data = sheet.getRange(2, 1, lastRow - 1, COL.NUMERO_CUENTA + 1).getValues();
    var contCompletados = 0, contOK = 0, contAvisos = 0, contOtrosBancos = 0, contSinRut = 0;

    for (var i = 0; i < data.length; i++) {
      var fila = i + 2;
      var rutRaw     = String(data[i][COL.RUT]          || '').trim();
      var banco      = String(data[i][COL.BANCO]         || '').trim();
      var tipoCuenta = String(data[i][COL.TIPO_CUENTA]   || '').trim();
      var numCuenta  = String(data[i][COL.NUMERO_CUENTA] || '').trim();
      var nombre     = String(data[i][COL.NOMBRE]        || '').trim();
      var bancoUp    = banco.toUpperCase();

      var rutLimpio = cleanRut(rutRaw);
      if (!rutLimpio || rutLimpio.length < 7) { contSinRut++; continue; }
      var rutBody = rutLimpio.slice(0, -1);

      if (!banco) {
        sheet.getRange(fila, COL.BANCO + 1).setValue('BANCO ESTADO (Cuenta RUT)');
        sheet.getRange(fila, COL.TIPO_CUENTA + 1).setValue('CUENTA VISTA');
        sheet.getRange(fila, COL.NUMERO_CUENTA + 1).setValue(rutBody);
        try { CacheService.getScriptCache().remove('user_' + rutLimpio); } catch(e) {}
        Logger.log('✅ Fila ' + fila + ' | ' + nombre + ' | Completado → Cuenta RUT: ' + rutBody);
        contCompletados++;
        if (contCompletados % 30 === 0) Utilities.sleep(500);

      } else if (bancoUp === 'BANCO ESTADO (CUENTA RUT)') {
        var tipoCuentaUp = tipoCuenta.toUpperCase();
        if (tipoCuentaUp !== '' && tipoCuentaUp !== 'CUENTA VISTA') { contOtrosBancos++; continue; }
        var tipoOK = (tipoCuentaUp === 'CUENTA VISTA'), numeroOK = (numCuenta === rutBody);
        if (tipoOK && numeroOK) {
          contOK++;
        } else {
          var problemas = [];
          if (!tipoOK) { sheet.getRange(fila, COL.TIPO_CUENTA + 1).setValue('CUENTA VISTA'); problemas.push('TIPO_CUENTA corregido'); }
          if (!numeroOK) { sheet.getRange(fila, COL.NUMERO_CUENTA + 1).setValue(rutBody); problemas.push('NUMERO_CUENTA corregido'); }
          try { CacheService.getScriptCache().remove('user_' + rutLimpio); } catch(e) {}
          Logger.log('🔧 Fila ' + fila + ' | ' + nombre + ' | ' + problemas.join(' | '));
          contAvisos++;
          if (contAvisos % 30 === 0) Utilities.sleep(500);
        }
      } else {
        contOtrosBancos++;
      }
    }

    Logger.log('============= RESUMEN =============');
    Logger.log('✅ Completados: ' + contCompletados + ' | Verificados OK: ' + contOK + ' | Inconsistencias: ' + contAvisos + ' | Otro banco: ' + contOtrosBancos + ' | Sin RUT: ' + contSinRut);
  } catch (e) {
    Logger.log('❌ Error en completarCamposBancariosEnBlanco: ' + e.toString());
  }
}
