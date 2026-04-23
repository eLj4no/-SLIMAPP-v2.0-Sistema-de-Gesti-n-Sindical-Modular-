// ==========================================
// MODULO_GAMIFICACION.GS — SLIM Quest
// ==========================================

var GRADOS_SLIM = [
  { nombre: "Aspirante",  minXP: 0,     maxXP: 1500,   icono: "🌱" },
  { nombre: "Aprendiz",   minXP: 1501,  maxXP: 4500,   icono: "⚙️" },
  { nombre: "Trabajador", minXP: 4501,  maxXP: 10000,  icono: "🔩" },
  { nombre: "Defensor",   minXP: 10001, maxXP: 18000,  icono: "🛡️" },
  { nombre: "Negociador", minXP: 18001, maxXP: 30000,  icono: "⚖️" },
  { nombre: "Dirigente",  minXP: 30001, maxXP: 999999, icono: "🏆" }
];

function calcularGrado_(xp) {
  for (var i = GRADOS_SLIM.length - 1; i >= 0; i--) {
    if (xp >= GRADOS_SLIM[i].minXP) return GRADOS_SLIM[i];
  }
  return GRADOS_SLIM[0];
}

// ==========================================
// PROGRESO E INICIALIZACIÓN
// ==========================================

/**
 * Obtiene el progreso completo de un socio en SLIM Quest.
 * Si el socio no tiene registro aún, lo crea automáticamente.
 */
function getProgresoSocio(rutInput) {
  try {
    var rutLimpio = cleanRut(rutInput);
    if (!rutLimpio) return { success: false, message: "RUT inválido." };

    var sheet = getSheet('GAMIFICACION', 'GAMIFICACION');
    if (!sheet) return { success: false, message: "Módulo de gamificación no configurado." };

    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var data = sheet.getRange(2, 1, lastRow - 1, 11).getDisplayValues();
      var COL  = CONFIG.COLUMNAS.GAMIFICACION;

      for (var i = 0; i < data.length; i++) {
        if (cleanRut(data[i][COL.RUT]) === rutLimpio) {
          var estado = String(data[i][COL.ESTADO] || "ACTIVO").toUpperCase();
          if (estado === "DESVINCULADO") {
            return { success: false, desvinculado: true, message: "Tu participación en SLIM Quest está suspendida porque tu estado en el sindicato es DESVINCULADO. Tu historial de XP y logros queda guardado.", xp: parseInt(data[i][COL.XP_TOTAL]) || 0, grado: data[i][COL.GRADO] || "Aspirante" };
          }

          var xp   = parseInt(data[i][COL.XP_TOTAL]) || 0;
          var grado = calcularGrado_(xp);
          var gradoSiguiente = null;
          for (var g = 0; g < GRADOS_SLIM.length; g++) {
            if (GRADOS_SLIM[g].minXP > xp) { gradoSiguiente = GRADOS_SLIM[g]; break; }
          }
          var logros = [];
          try { logros = JSON.parse(data[i][COL.LOGROS] || "[]"); } catch(e) {}

          return {
            success: true,
            rut: data[i][COL.RUT],
            nombre: data[i][COL.NOMBRE],
            xp: xp,
            grado: grado,
            gradoSiguiente: gradoSiguiente,
            xpParaSiguiente: gradoSiguiente ? gradoSiguiente.minXP - xp : 0,
            racha: parseInt(data[i][COL.RACHA_ACTUAL]) || 0,
            rachaMax: parseInt(data[i][COL.RACHA_MAX]) || 0,
            logros: logros,
            quizzesCompletados: parseInt(data[i][COL.QUIZZES_COMPLETADOS]) || 0,
            quizHoy: data[i][COL.QUIZ_ULTIMO_DIA] || "",
            ultimaActividad: data[i][COL.ULTIMA_ACTIVIDAD] || "",
            estado: estado
          };
        }
      }
    }

    var usuarioData = obtenerUsuarioPorRut(rutInput);
    if (!usuarioData.encontrado) return { success: false, message: "RUT no encontrado en el sistema." };
    return inicializarSocioGamificacion_(rutLimpio, usuarioData.nombre || "Socio", usuarioData.estado || "ACTIVO");

  } catch (e) {
    Logger.log("❌ Error en getProgresoSocio: " + e.toString());
    return { success: false, message: "Error: " + e.toString() };
  }
}

function inicializarSocioGamificacion_(rutLimpio, nombreSocio, estadoSocio) {
  try {
    var sheet  = getSheet('GAMIFICACION', 'GAMIFICACION');
    var hoy    = Utilities.formatDate(new Date(), "America/Santiago", "dd/MM/yyyy HH:mm");
    var estado = estadoSocio || "ACTIVO";

    sheet.appendRow([rutLimpio, nombreSocio, 0, "Aspirante", "[]", 0, 0, hoy, "", 0, estado, 0]);
    Logger.log("✅ Socio inicializado en SLIM Quest: " + rutLimpio + " (" + nombreSocio + ") — " + estado);

    return {
      success: true,
      rut: rutLimpio,
      nombre: nombreSocio,
      xp: 0,
      grado: GRADOS_SLIM[0],
      gradoSiguiente: GRADOS_SLIM[1],
      xpParaSiguiente: GRADOS_SLIM[1].minXP,
      racha: 0,
      rachaMax: 0,
      logros: [],
      quizzesCompletados: 0,
      quizHoy: "",
      ultimaActividad: hoy,
      estado: estado,
      recienCreado: true
    };
  } catch (e) {
    Logger.log("❌ Error en inicializarSocioGamificacion_: " + e.toString());
    return { success: false, message: "Error al inicializar: " + e.toString() };
  }
}

// ==========================================
// XP Y LOGROS
// ==========================================

function guardarXP(rutInput, cantidad, motivo) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, message: "Servidor ocupado, intenta nuevamente." };
  try {
    var rutLimpio = cleanRut(rutInput);
    var sheet     = getSheet('GAMIFICACION', 'GAMIFICACION');
    var lastRow   = sheet.getLastRow();
    var COL       = CONFIG.COLUMNAS.GAMIFICACION;
    var hoy       = Utilities.formatDate(new Date(), "America/Santiago", "dd/MM/yyyy HH:mm");

    if (lastRow < 2) {
      lock.releaseLock();
      inicializarSocioGamificacion_(rutInput);
      return guardarXP(rutInput, cantidad, motivo);
    }

    var data = sheet.getRange(2, 1, lastRow - 1, 10).getDisplayValues();
    for (var i = 0; i < data.length; i++) {
      if (cleanRut(data[i][COL.RUT]) === rutLimpio) {
        var xpActual  = parseInt(data[i][COL.XP_TOTAL]) || 0;
        var xpNuevo   = xpActual + cantidad;
        var gradoAnterior = data[i][COL.GRADO];
        var gradoNuevo    = calcularGrado_(xpNuevo);
        var filaReal      = i + 2;

        sheet.getRange(filaReal, COL.XP_TOTAL + 1).setValue(xpNuevo);
        sheet.getRange(filaReal, COL.GRADO + 1).setValue(gradoNuevo.nombre);
        sheet.getRange(filaReal, COL.ULTIMA_ACTIVIDAD + 1).setValue(hoy);

        Logger.log("✅ XP [" + motivo + "]: " + rutLimpio + " +" + cantidad + "XP → Total: " + xpNuevo + " | Grado: " + gradoNuevo.nombre);
        return { success: true, xpSumado: cantidad, xpTotal: xpNuevo, grado: gradoNuevo, subioGrado: gradoAnterior !== gradoNuevo.nombre, gradoAnterior: gradoAnterior, motivo: motivo };
      }
    }

    lock.releaseLock();
    inicializarSocioGamificacion_(rutInput);
    return guardarXP(rutInput, cantidad, motivo);

  } catch (e) {
    Logger.log("❌ Error en guardarXP: " + e.toString());
    return { success: false, message: "Error: " + e.toString() };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function otorgarLogro(rutInput, codigoLogro, nombreLogro, iconoLogro) {
  try {
    var rutLimpio = cleanRut(rutInput);
    var sheet     = getSheet('GAMIFICACION', 'GAMIFICACION');
    var lastRow   = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: "Socio no registrado en gamificación." };

    var data = sheet.getRange(2, 1, lastRow - 1, 10).getDisplayValues();
    var COL  = CONFIG.COLUMNAS.GAMIFICACION;

    for (var i = 0; i < data.length; i++) {
      if (cleanRut(data[i][COL.RUT]) === rutLimpio) {
        var logros = [];
        try { logros = JSON.parse(data[i][COL.LOGROS] || "[]"); } catch(e) {}

        if (logros.some(function(l) { return l.codigo === codigoLogro; })) {
          return { success: true, nuevo: false };
        }

        var fecha = Utilities.formatDate(new Date(), "America/Santiago", "dd/MM/yyyy");
        logros.push({ codigo: codigoLogro, nombre: nombreLogro, icono: iconoLogro, fecha: fecha });
        sheet.getRange(i + 2, COL.LOGROS + 1).setValue(JSON.stringify(logros));

        Logger.log("🏅 Logro [" + codigoLogro + "] otorgado a " + rutLimpio);
        return { success: true, nuevo: true, logro: { codigo: codigoLogro, nombre: nombreLogro, icono: iconoLogro } };
      }
    }
    return { success: false, message: "Socio no encontrado en gamificación." };
  } catch (e) {
    Logger.log("❌ Error en otorgarLogro: " + e.toString());
    return { success: false, message: "Error: " + e.toString() };
  }
}

// ==========================================
// LEADERBOARD
// ==========================================

function getLeaderboard(rutInput) {
  try {
    var sheet = getSheet('GAMIFICACION', 'GAMIFICACION');
    if (!sheet) return { success: false, message: "Hoja no encontrada." };

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, top10: [], miPosicion: null };

    var COL       = CONFIG.COLUMNAS.GAMIFICACION;
    var data      = sheet.getRange(2, 1, lastRow - 1, 12).getDisplayValues();
    var rutLimpio = rutInput ? cleanRut(rutInput) : "";
    var lista     = [];

    for (var i = 0; i < data.length; i++) {
      var xp     = parseInt(data[i][COL.XP_TOTAL]) || 0;
      var estado = String(data[i][COL.ESTADO]).toUpperCase().trim();
      if (estado === "DESVINCULADO") continue;
      lista.push({ rut: cleanRut(data[i][COL.RUT]), nombre: data[i][COL.NOMBRE] || "Socio", xp: xp, grado: calcularGrado_(xp) });
    }

    lista.sort(function(a, b) { return b.xp - a.xp; });

    var top10 = lista.slice(0, 10).map(function(s, idx) {
      var partes  = s.nombre.trim().split(" ");
      var visible = partes[0] + (partes[1] ? " " + partes[1] : "") + (partes[2] ? " " + partes[2][0] + "." : "");
      return { posicion: idx + 1, nombre: visible, xp: s.xp, grado: s.grado, esMio: (s.rut === rutLimpio) };
    });

    var miPosicion = null;
    if (rutLimpio) {
      var miIdx = -1;
      for (var j = 0; j < lista.length; j++) {
        if (lista[j].rut === rutLimpio) { miIdx = j; break; }
      }
      if (miIdx >= 0) miPosicion = { posicion: miIdx + 1, xp: lista[miIdx].xp, grado: lista[miIdx].grado };
    }

    Logger.log("🏆 Leaderboard | top10: " + top10.length + " | RUT: " + rutLimpio);
    return { success: true, top10: top10, miPosicion: miPosicion };

  } catch (e) {
    Logger.log("❌ Error en getLeaderboard: " + e.toString());
    return { success: false, message: e.toString() };
  }
}

// ==========================================
// SINCRONIZACIÓN (trigger diario 1 AM)
// ==========================================

function sincronizarSociosGamificacion() {
  try {
    Logger.log("🔄 Iniciando sincronización de socios con SLIM Quest...");

    var sheetUsuarios = getSheet('USUARIOS', 'USUARIOS');
    var sheetGame     = getSheet('GAMIFICACION', 'GAMIFICACION');
    if (!sheetUsuarios || !sheetGame) { Logger.log("❌ No se pudieron obtener las hojas necesarias."); return; }

    var COL_U = CONFIG.COLUMNAS.USUARIOS;
    var COL_G = CONFIG.COLUMNAS.GAMIFICACION;
    var hoy   = Utilities.formatDate(new Date(), "America/Santiago", "dd/MM/yyyy HH:mm");

    var lastRowGame = sheetGame.getLastRow();
    var mapaGame    = {};
    if (lastRowGame >= 2) {
      var dataGame = sheetGame.getRange(2, 1, lastRowGame - 1, 11).getDisplayValues();
      for (var i = 0; i < dataGame.length; i++) {
        var rut = cleanRut(dataGame[i][COL_G.RUT]);
        if (rut) mapaGame[rut] = { fila: i + 2, nombre: dataGame[i][COL_G.NOMBRE], estado: dataGame[i][COL_G.ESTADO] };
      }
    }

    var lastRowU = sheetUsuarios.getLastRow();
    if (lastRowU < 2) { Logger.log("ℹ️ No hay socios en BD_SLIMAPP."); return; }

    var dataU       = sheetUsuarios.getRange(2, 1, lastRowU - 1, COL_U.ESTADO + 1).getDisplayValues();
    var creados     = 0, actualizados = 0, sinRut = 0;
    var nuevasFilas = [];

    for (var j = 0; j < dataU.length; j++) {
      var rutLimpio = cleanRut(dataU[j][COL_U.RUT]);
      if (!rutLimpio) { sinRut++; continue; }

      var nombre     = String(dataU[j][COL_U.NOMBRE] || "Socio").trim();
      var estadoRaw  = String(dataU[j][COL_U.ESTADO] || "ACTIVO").trim().toUpperCase();
      var estadoNorm = (estadoRaw === "ACTIVO" || estadoRaw === "SI" || estadoRaw === "TRUE") ? "ACTIVO" : "DESVINCULADO";

      if (mapaGame[rutLimpio]) {
        var reg          = mapaGame[rutLimpio];
        var nombreCambio = reg.nombre !== nombre;
        var estadoCambio = String(reg.estado || "").toUpperCase() !== estadoNorm;
        if (nombreCambio || estadoCambio) {
          sheetGame.getRange(reg.fila, COL_G.NOMBRE + 1).setValue(nombre);
          sheetGame.getRange(reg.fila, COL_G.ESTADO + 1).setValue(estadoNorm);
          actualizados++;
          Logger.log("🔄 Actualizado: " + rutLimpio + " | Estado: " + estadoNorm);
        }
      } else {
        nuevasFilas.push([rutLimpio, nombre, 0, "Aspirante", "[]", 0, 0, hoy, "", 0, estadoNorm, 0]);
        creados++;
      }
    }

    if (nuevasFilas.length > 0) {
      var primeraFilaLibre = sheetGame.getLastRow() + 1;
      sheetGame.getRange(primeraFilaLibre, 1, nuevasFilas.length, 12).setValues(nuevasFilas);
    }

    Logger.log("✅ Socios creados: " + creados + " | Actualizados: " + actualizados + " | Sin RUT: " + sinRut);

  } catch (e) {
    Logger.log("❌ Error en sincronizarSociosGamificacion: " + e.toString());
  }
}

function configurarTriggerGamificacion() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === "sincronizarSociosGamificacion") {
      ScriptApp.deleteTrigger(trigger);
      Logger.log("🗑️ Trigger anterior eliminado.");
    }
  });
  ScriptApp.newTrigger("sincronizarSociosGamificacion").timeBased().everyDays(1).atHour(1).create();
  Logger.log("✅ Trigger diario configurado: sincronizarSociosGamificacion todos los días a la 1am.");
}

// ==========================================
// QUIZ DIARIO
// ==========================================

function obtenerPreguntasQuiz(rutInput, cantidad) {
  try {
    cantidad = cantidad || 5;
    var sheet = getSheet('GAMIFICACION', 'BANCO_PREGUNTAS');
    if (!sheet) return { success: false, message: "Banco de preguntas no disponible." };

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: "No hay preguntas cargadas aún." };

    var data = sheet.getRange(2, 1, lastRow - 1, 13).getDisplayValues();
    var COL  = CONFIG.COLUMNAS.BANCO_PREGUNTAS;

    var gradoActual = "Aspirante";
    try {
      var progreso = getProgresoSocio(rutInput);
      if (progreso.success && progreso.grado) gradoActual = progreso.grado.nombre;
    } catch(e) {}

    var PESOS = {
      "Aspirante":  { BASICO: 4, INTERMEDIO: 1, AVANZADO: 0 },
      "Aprendiz":   { BASICO: 3, INTERMEDIO: 2, AVANZADO: 0 },
      "Trabajador": { BASICO: 2, INTERMEDIO: 2, AVANZADO: 1 },
      "Defensor":   { BASICO: 1, INTERMEDIO: 3, AVANZADO: 1 },
      "Negociador": { BASICO: 1, INTERMEDIO: 2, AVANZADO: 2 },
      "Dirigente":  { BASICO: 0, INTERMEDIO: 2, AVANZADO: 3 }
    };
    var pesoActual = PESOS[gradoActual] || PESOS["Aspirante"];
    var porNivel   = { BASICO: [], INTERMEDIO: [], AVANZADO: [] };

    for (var i = 0; i < data.length; i++) {
      var activa = String(data[i][COL.ACTIVA]).toUpperCase();
      if (activa !== "TRUE" && activa !== "VERDADERO" && activa !== "1") continue;
      var nivel = String(data[i][COL.NIVEL]).toUpperCase().trim();
      if (nivel === "DIRIGENTE") continue;
      var pregunta = {
        id: data[i][COL.ID], categoria: data[i][COL.CATEGORIA], nivel: nivel,
        pregunta: data[i][COL.PREGUNTA],
        opciones: { A: data[i][COL.OPCION_A], B: data[i][COL.OPCION_B], C: data[i][COL.OPCION_C], D: data[i][COL.OPCION_D] },
        respuesta: data[i][COL.RESPUESTA].toUpperCase().trim(), explicacion: data[i][COL.EXPLICACION],
        xp: parseInt(data[i][COL.XP]) || 20, fuente: data[i][COL.FUENTE]
      };
      if (porNivel[nivel] !== undefined) porNivel[nivel].push(pregunta);
    }

    function sacarAleatorio(arr, n) {
      var copia = arr.slice(), resultado = [];
      for (var k = 0; k < n && copia.length > 0; k++) {
        var idx = Math.floor(Math.random() * copia.length);
        resultado.push(copia.splice(idx, 1)[0]);
      }
      return resultado;
    }

    var seleccion = [];
    seleccion = seleccion.concat(sacarAleatorio(porNivel.BASICO,     pesoActual.BASICO));
    seleccion = seleccion.concat(sacarAleatorio(porNivel.INTERMEDIO, pesoActual.INTERMEDIO));
    seleccion = seleccion.concat(sacarAleatorio(porNivel.AVANZADO,   pesoActual.AVANZADO));

    if (seleccion.length < cantidad) {
      var usados    = {};
      seleccion.forEach(function(p) { usados[p.id] = true; });
      var restantes = [];
      for (var r = 0; r < data.length; r++) {
        var act = String(data[r][COL.ACTIVA]).toUpperCase();
        if ((act !== "TRUE" && act !== "VERDADERO" && act !== "1") || usados[data[r][COL.ID]]) continue;
        restantes.push({
          id: data[r][COL.ID], categoria: data[r][COL.CATEGORIA],
          nivel: String(data[r][COL.NIVEL]).toUpperCase().trim(), pregunta: data[r][COL.PREGUNTA],
          opciones: { A: data[r][COL.OPCION_A], B: data[r][COL.OPCION_B], C: data[r][COL.OPCION_C], D: data[r][COL.OPCION_D] },
          respuesta: String(data[r][COL.RESPUESTA]).toUpperCase().trim(),
          explicacion: data[r][COL.EXPLICACION], xp: parseInt(data[r][COL.XP]) || 20, fuente: data[r][COL.FUENTE]
        });
      }
      seleccion = seleccion.concat(sacarAleatorio(restantes, cantidad - seleccion.length));
    }

    for (var s = seleccion.length - 1; s > 0; s--) {
      var jj = Math.floor(Math.random() * (s + 1));
      var tmp = seleccion[s]; seleccion[s] = seleccion[jj]; seleccion[jj] = tmp;
    }

    Logger.log("✅ Quiz generado para " + rutInput + " | Grado: " + gradoActual + " | Preguntas: " + seleccion.length);
    return { success: true, preguntas: seleccion.slice(0, cantidad), gradoActual: gradoActual };

  } catch (e) {
    Logger.log("❌ Error en obtenerPreguntasQuiz: " + e.toString());
    return { success: false, message: "Error: " + e.toString() };
  }
}

function registrarResultadoQuiz(rutInput, correctas, xpGanado) {
  try {
    var rutLimpio = cleanRut(rutInput);
    if (!rutLimpio) return { success: false, message: "RUT inválido." };

    var sheet   = getSheet('GAMIFICACION', 'GAMIFICACION');
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: "Socio no inicializado." };

    var data  = sheet.getRange(2, 1, lastRow - 1, 12).getDisplayValues();
    var COL   = CONFIG.COLUMNAS.GAMIFICACION;
    var hoy   = Utilities.formatDate(new Date(), "America/Santiago", "dd/MM/yyyy");
    var ahora = Utilities.formatDate(new Date(), "America/Santiago", "dd/MM/yyyy HH:mm");

    for (var i = 0; i < data.length; i++) {
      if (cleanRut(data[i][COL.RUT]) !== rutLimpio) continue;

      var filaReal      = i + 2;
      var quizUltimoDia = String(data[i][COL.QUIZ_ULTIMO_DIA] || "").trim();
      if (quizUltimoDia === hoy) return { success: false, message: "Ya completaste el quiz de hoy. ¡Vuelve mañana!" };

      var rachaActual  = parseInt(data[i][COL.RACHA_ACTUAL]) || 0;
      var rachaMax     = parseInt(data[i][COL.RACHA_MAX])    || 0;
      var ayer         = new Date();
      ayer.setDate(ayer.getDate() - 1);
      var ayerStr      = Utilities.formatDate(ayer, "America/Santiago", "dd/MM/yyyy");
      var nuevaRacha   = (quizUltimoDia === ayerStr) ? rachaActual + 1 : 1;
      var nuevaRachaMax= Math.max(nuevaRacha, rachaMax);

      var BONOS_RACHA  = { 3: 20, 7: 50, 14: 80, 21: 100, 30: 160, 60: 280, 100: 500 };
      var xpFinal      = xpGanado;
      var xpBonoRacha  = 0;
      if (BONOS_RACHA[nuevaRacha] !== undefined) {
        xpBonoRacha = BONOS_RACHA[nuevaRacha];
        xpFinal    += xpBonoRacha;
      } else if (nuevaRacha > 100 && nuevaRacha % 7 === 0) {
        xpBonoRacha = 100;
        xpFinal    += xpBonoRacha;
      }

      var xpActual       = parseInt(data[i][COL.XP_TOTAL]) || 0;
      var xpNuevo        = xpActual + xpFinal;
      var gradoNuevo     = calcularGrado_(xpNuevo);
      var quizzesAnt     = parseInt(data[i][COL.QUIZZES_COMPLETADOS]) || 0;
      var quizzesPerfAnt = parseInt(data[i][COL.QUIZZES_PERFECTOS])   || 0;
      var qTotalNuevo    = quizzesAnt + 1;
      var qPerfNuevo     = (correctas === 5) ? quizzesPerfAnt + 1 : quizzesPerfAnt;

      sheet.getRange(filaReal, COL.XP_TOTAL + 1).setValue(xpNuevo);
      sheet.getRange(filaReal, COL.GRADO + 1).setValue(gradoNuevo.nombre);
      sheet.getRange(filaReal, COL.RACHA_ACTUAL + 1).setValue(nuevaRacha);
      sheet.getRange(filaReal, COL.RACHA_MAX + 1).setValue(nuevaRachaMax);
      sheet.getRange(filaReal, COL.QUIZ_ULTIMO_DIA + 1).setValue(hoy);
      sheet.getRange(filaReal, COL.QUIZZES_COMPLETADOS + 1).setValue(qTotalNuevo);
      sheet.getRange(filaReal, COL.ULTIMA_ACTIVIDAD + 1).setValue(ahora);
      sheet.getRange(filaReal, COL.QUIZZES_PERFECTOS + 1).setValue(qPerfNuevo);

      var logrosNuevos = [];
      function evalLogro(codigo, nombre, icono) {
        var r = otorgarLogro(rutInput, codigo, nombre, icono);
        if (r.success && r.nuevo) logrosNuevos.push(r.logro);
      }

      if (qTotalNuevo === 1)   evalLogro("PRIMER_QUIZ",   "Primer Quiz",           "🎮");
      if (qTotalNuevo === 10)  evalLogro("10_QUIZZES",    "10 Quizzes",            "⭐");
      if (qTotalNuevo === 25)  evalLogro("25_QUIZZES",    "Estudiante Sindical",   "🎓");
      if (qTotalNuevo === 50)  evalLogro("50_QUIZZES",    "Comprometido",          "📖");
      if (qTotalNuevo === 100) evalLogro("100_QUIZZES",   "Maestro del Sindicato", "🏛️");
      if (correctas === 5)     evalLogro("QUIZ_PERFECTO", "Quiz Perfecto",         "🎯");
      if (qPerfNuevo === 3)    evalLogro("3_PERFECTOS",   "Imparable",             "💯");
      if (qPerfNuevo === 10)   evalLogro("10_PERFECTOS",  "Sin Errores",           "🌟");
      if (nuevaRacha === 3)    evalLogro("RACHA_3",   "Primeros pasos",    "✨");
      if (nuevaRacha === 7)    evalLogro("RACHA_7",   "Racha de 7 días",   "🔥");
      if (nuevaRacha === 14)   evalLogro("RACHA_14",  "Racha de 2 semanas","🔥🔥");
      if (nuevaRacha === 30)   evalLogro("RACHA_30",  "Racha de 30 días",  "📅");
      if (nuevaRacha === 60)   evalLogro("RACHA_60",  "Racha de 60 días",  "🗓️");
      if (nuevaRacha === 100)  evalLogro("RACHA_100", "Centenario",        "💎");

      Logger.log("✅ Quiz OK: " + rutLimpio + " | " + correctas + "/5 | +" + xpFinal + " XP (bono: +" + xpBonoRacha + ") | Racha: " + nuevaRacha);

      if (data[i][COL.GRADO] !== gradoNuevo.nombre) {
        var correoSocio = obtenerUsuarioPorRut(rutInput).correo || '';
        enviarCorreoNivel(correoSocio, data[i][COL.NOMBRE], gradoNuevo.nombre, xpNuevo);
      }

      return {
        success: true,
        correctas: correctas,
        xpGanado: xpFinal,
        xpBase: xpGanado,
        xpBonoRacha: xpBonoRacha,
        nuevaRacha: nuevaRacha,
        xpTotal: xpNuevo,
        grado: gradoNuevo,
        subioGrado: data[i][COL.GRADO] !== gradoNuevo.nombre,
        gradoAnterior: data[i][COL.GRADO],
        logrosNuevos: logrosNuevos
      };
    }

    return { success: false, message: "Socio no encontrado en gamificación." };

  } catch (e) {
    Logger.log("❌ Error en registrarResultadoQuiz: " + e.toString());
    return { success: false, message: "Error: " + e.toString() };
  }
}

// ==========================================
// CORREO DE SUBIDA DE NIVEL
// ==========================================

function enviarCorreoNivel(correo, nombre, gradoNuevo, xpTotal) {
  if (!correo || !correo.includes('@')) return;

  var CONFIG_GRADO = {
    'Aspirante':  { headerBg:'#15803d', color:'#22c55e', badgeBg:'#dcfce7', badgeText:'#14532d', icono:'🌱', nivel:'1/6', quote:'Cada gran viaje comienza con el primer paso. ¡Has dado el tuyo!', nextNivel:'⚙️ Aprendiz', nextXp:'1.501 XP' },
    'Aprendiz':   { headerBg:'#1d4ed8', color:'#3b82f6', badgeBg:'#dbeafe', badgeText:'#1e3a8a', icono:'⚙️', nivel:'2/6', quote:'El conocimiento es la herramienta más poderosa del movimiento sindical.', nextNivel:'🔩 Trabajador', nextXp:'4.501 XP' },
    'Trabajador': { headerBg:'#c2410c', color:'#f97316', badgeBg:'#ffedd5', badgeText:'#7c2d12', icono:'🔩', nivel:'3/6', quote:'El trabajo organizado mueve montañas. Tú eres la fuerza del sindicato.', nextNivel:'🛡️ Defensor', nextXp:'10.001 XP' },
    'Defensor':   { headerBg:'#6d28d9', color:'#8b5cf6', badgeBg:'#ede9fe', badgeText:'#4c1d95', icono:'🛡️', nivel:'4/6', quote:'Defender los derechos colectivos es el corazón del sindicalismo.', nextNivel:'⚖️ Negociador', nextXp:'18.001 XP' },
    'Negociador': { headerBg:'#b45309', color:'#d97706', badgeBg:'#fef3c7', badgeText:'#78350f', icono:'⚖️', nivel:'5/6', quote:'La negociación efectiva nace del conocimiento profundo y la preparación incansable.', nextNivel:'🏆 Dirigente', nextXp:'30.001 XP' },
    'Dirigente':  { headerBg:'#92400e', color:'#f59e0b', badgeBg:'#fffbeb', badgeText:'#78350f', icono:'🏆', nivel:'6/6', quote:'El verdadero líder no es quien dirige, sino quien inspira y transforma.', nextNivel: null, nextXp: null }
  };

  var cfg = CONFIG_GRADO[gradoNuevo];
  if (!cfg) return;

  var xpFmt = Number(xpTotal).toLocaleString('es-CL');
  var nextSection = cfg.nextNivel
    ? '<div style="background:' + cfg.badgeBg + ';border-radius:12px;padding:14px 16px;margin-bottom:16px;border:1px solid ' + cfg.color + '33;"><p style="font-size:10px;font-weight:700;color:' + cfg.badgeText + ';text-transform:uppercase;letter-spacing:0.5px;margin:0 0 5px;">Próximo nivel</p><p style="font-size:13px;color:' + cfg.badgeText + ';margin:0;font-weight:600;">' + cfg.nextNivel + ' — desde ' + cfg.nextXp + '</p></div>'
    : '<div style="background:#fef3c7;border-radius:12px;padding:14px 16px;margin-bottom:16px;border:1px solid #fbbf2433;"><p style="font-size:10px;font-weight:700;color:#78350f;text-transform:uppercase;letter-spacing:0.5px;margin:0 0 5px;">Nivel máximo</p><p style="font-size:13px;color:#78350f;margin:0;font-weight:600;">🏅 Has completado todos los niveles de SLIM Quest</p></div>';

  var html = '<!DOCTYPE html><html><body style="margin:0;padding:20px;background:#f1f5f9;font-family:Arial,sans-serif;">'
    + '<div style="max-width:520px;margin:0 auto;background:#fff;border-radius:16px;overflow:hidden;border:1px solid #e2e8f0;">'
    + '<div style="background:' + cfg.headerBg + ';padding:32px 24px;text-align:center;">'
    + '<div style="font-size:56px;line-height:1;margin-bottom:10px;">' + cfg.icono + '</div>'
    + '<h1 style="margin:0 0 4px;font-size:22px;font-weight:600;color:#fff;">¡Subiste a ' + gradoNuevo + '!</h1>'
    + '<p style="margin:0;font-size:13px;color:rgba(255,255,255,.75);">Sindicato SLIM N°3 · SLIM Quest</p></div>'
    + '<div style="padding:24px;">'
    + '<p style="font-size:15px;font-weight:600;color:#1e293b;margin-bottom:12px;">Hola, ' + nombre + '</p>'
    + '<p style="font-size:13px;color:#64748b;line-height:1.7;margin-bottom:20px;">Tu constancia y dedicación en SLIM Quest te han llevado a un nuevo nivel.</p>'
    + '<div style="display:flex;gap:8px;margin-bottom:20px;">'
    + '<div style="flex:1;text-align:center;background:#f8fafc;border-radius:10px;padding:12px 6px;"><div style="font-size:18px;font-weight:700;color:' + cfg.color + ';">' + xpFmt + '</div><div style="font-size:10px;color:#94a3b8;margin-top:2px;">XP acumulados</div></div>'
    + '<div style="flex:1;text-align:center;background:#f8fafc;border-radius:10px;padding:12px 6px;"><div style="font-size:15px;font-weight:700;color:' + cfg.color + ';">Nivel ' + cfg.nivel + '</div><div style="font-size:10px;color:#94a3b8;margin-top:2px;">Tu posición</div></div>'
    + '</div>'
    + '<div style="border-left:3px solid ' + cfg.color + ';background:' + cfg.badgeBg + ';border-radius:0 10px 10px 0;padding:14px 16px;margin-bottom:16px;">'
    + '<p style="font-size:13px;font-style:italic;color:' + cfg.badgeText + ';line-height:1.6;margin:0;">"' + cfg.quote + '"</p></div>'
    + nextSection + '</div>'
    + '<div style="padding:16px 24px;border-top:1px solid #f1f5f9;text-align:center;">'
    + '<p style="font-size:11px;color:#94a3b8;margin:2px 0;">Mensaje automático de SLIMAPP</p>'
    + '<p style="font-size:11px;color:#cbd5e1;margin:2px 0;">Sindicato SLIM N°3 · No responder a este correo</p>'
    + '</div></div></body></html>';

  try {
    MailApp.sendEmail({ to: correo, subject: cfg.icono + ' ¡Subiste a ' + gradoNuevo + ' en SLIM Quest! — Sindicato SLIM N°3', htmlBody: html, name: 'Sindicato SLIM N°3' });
    Logger.log('📧 Correo de nivel enviado a ' + correo + ' — Grado: ' + gradoNuevo);
  } catch(e) {
    Logger.log('⚠️ Error enviando correo de nivel: ' + e.toString());
  }
}

// ==========================================
// NIVEL SECRETO (exclusivo grado Dirigente)
// ==========================================

function obtenerPreguntasSecreto(rutInput, cantidad) {
  try {
    cantidad = cantidad || 5;
    var progreso = getProgresoSocio(rutInput);
    if (!progreso.success || progreso.grado.nombre !== "Dirigente") {
      return { success: false, accesoDenegado: true, message: "El Nivel Secreto es exclusivo para socios con grado Dirigente. Sigue acumulando XP en el quiz diario para alcanzarlo." };
    }

    var sheet = getSheet('GAMIFICACION', 'BANCO_PREGUNTAS');
    if (!sheet) return { success: false, message: "Banco de preguntas no disponible." };

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: "No hay preguntas en el nivel secreto." };

    var data = sheet.getRange(2, 1, lastRow - 1, 13).getDisplayValues();
    var COL  = CONFIG.COLUMNAS.BANCO_PREGUNTAS;
    var pool = [];

    for (var i = 0; i < data.length; i++) {
      var activa = String(data[i][COL.ACTIVA]).toUpperCase();
      if (activa !== "TRUE" && activa !== "VERDADERO" && activa !== "1") continue;
      if (String(data[i][COL.NIVEL]).toUpperCase().trim() !== "DIRIGENTE") continue;
      pool.push({
        id: data[i][COL.ID], categoria: data[i][COL.CATEGORIA], nivel: "DIRIGENTE",
        pregunta: data[i][COL.PREGUNTA],
        opciones: { A: data[i][COL.OPCION_A], B: data[i][COL.OPCION_B], C: data[i][COL.OPCION_C], D: data[i][COL.OPCION_D] },
        respuesta: data[i][COL.RESPUESTA].toUpperCase().trim(),
        explicacion: data[i][COL.EXPLICACION], xp: parseInt(data[i][COL.XP]) || 50, fuente: data[i][COL.FUENTE]
      });
    }

    if (pool.length === 0) return { success: false, message: "No hay preguntas del nivel secreto cargadas aún." };

    for (var j = pool.length - 1; j > 0; j--) {
      var k = Math.floor(Math.random() * (j + 1));
      var tmp = pool[j]; pool[j] = pool[k]; pool[k] = tmp;
    }

    Logger.log("🔐 Quiz secreto generado para " + rutInput + " | Preguntas: " + Math.min(cantidad, pool.length));
    return { success: true, preguntas: pool.slice(0, cantidad), modoSecreto: true };

  } catch (e) {
    Logger.log("❌ Error en obtenerPreguntasSecreto: " + e.toString());
    return { success: false, message: "Error: " + e.toString() };
  }
}

// ==========================================
// MIGRACIÓN: ACTUALIZAR XP BANCO DE PREGUNTAS
// (Ejecutar manualmente una sola vez)
// ==========================================

function actualizarXpBancoPreguntas() {
  try {
    var sheet   = getSheet('GAMIFICACION', 'BANCO_PREGUNTAS');
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) { Logger.log("⚠️ Banco vacío."); return; }

    var COL          = CONFIG.COLUMNAS.BANCO_PREGUNTAS;
    var data         = sheet.getRange(2, 1, lastRow - 1, 13).getDisplayValues();
    var XP_MAP       = { BASICO: 15, INTERMEDIO: 25, AVANZADO: 40, DIRIGENTE: 50 };
    var actualizadas = 0, omitidas = 0;

    for (var i = 0; i < data.length; i++) {
      var nivel   = String(data[i][COL.NIVEL]).toUpperCase().trim();
      var xpNuevo = XP_MAP[nivel];
      if (!xpNuevo) { omitidas++; continue; }
      var xpActual = parseInt(data[i][COL.XP]) || 0;
      if (xpActual === xpNuevo) { omitidas++; continue; }
      sheet.getRange(i + 2, COL.XP + 1).setValue(xpNuevo);
      actualizadas++;
    }

    Logger.log("✅ actualizarXpBancoPreguntas — Actualizadas: " + actualizadas + " | Sin cambio: " + omitidas);
  } catch (e) {
    Logger.log("❌ Error: " + e.toString());
  }
}

// ==========================================
// SWITCHES MÓDULOS RELACIONADOS
// ==========================================

function obtenerEstadoSwitchSlimQuest() {
  try {
    var estado = PropertiesService.getScriptProperties().getProperty('slimquest_habilitado');
    return { success: true, habilitado: (estado === null || estado === 'true') };
  } catch (e) {
    return { success: true, habilitado: true };
  }
}

function toggleSwitchSlimQuest(estado) {
  try {
    PropertiesService.getScriptProperties().setProperty('slimquest_habilitado', estado ? 'true' : 'false');
    return { success: true };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

function obtenerEstadoSwitchCalculadora() {
  try {
    var estado = PropertiesService.getScriptProperties().getProperty('calculadora_habilitada');
    return { success: true, habilitado: (estado === null || estado === 'true') };
  } catch (e) {
    return { success: true, habilitado: true };
  }
}

function toggleSwitchCalculadora(estado) {
  try {
    PropertiesService.getScriptProperties().setProperty('calculadora_habilitada', estado ? 'true' : 'false');
    return { success: true };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

function obtenerEstadoSwitchContratoColectivo() {
  try {
    var estado = PropertiesService.getScriptProperties().getProperty('contrato_colectivo_habilitado');
    return { success: true, habilitado: (estado === null || estado === 'true') };
  } catch (e) {
    return { success: true, habilitado: true };
  }
}

function toggleSwitchContratoColectivo(estado) {
  try {
    PropertiesService.getScriptProperties().setProperty('contrato_colectivo_habilitado', estado ? 'true' : 'false');
    return { success: true };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}
