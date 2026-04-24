# 🏢 SLIMAPP — Sistema de Gestión Sindical
### Sindicato SLIM N°3

> Plataforma web de gestión integral para ~2.230 socios activos distribuidos en 16 regiones de Chile.

---

## 📋 Descripción General

**SLIMAPP** es una aplicación web desarrollada sobre **Google Apps Script** que permite a los socios del Sindicato SLIM N°3 realizar trámites sindicales en línea: justificar inasistencias, apelar multas, solicitar préstamos, gestionar permisos médicos, registrar asistencia a asambleas y más. La directiva puede gestionar y responder cada solicitud desde la misma plataforma.

La app se sirve como **Google Web App** (accesible desde cualquier navegador/dispositivo) y utiliza **Google Sheets** como base de datos y **Google Drive** para almacenamiento de archivos adjuntos.

---

## 🗂️ Estructura del Repositorio

```
SLIMAPP/
├── Code.gs                    # Router principal y funciones compartidas
├── Modulo_admin.gs            # Panel de administración y configuración de roles
├── Modulo_gamificacion.gs     # Sistema SLIM Quest (XP, grados, logros, quiz)
├── Modulo_prestamos.gs        # Lógica completa de préstamos sindicales
├── Index.html                 # Frontend principal (HTML + TailwindCSS + JS vanilla)
├── QR_Access.html             # Vinculación QR personal del dispositivo del socio
├── QR_Asistencia.html         # Registro de asistencia en punto de control QR
└── README.md                  # Este archivo
```

> Los demás módulos (Justificaciones, Apelaciones, Permisos Médicos, Asistencia, Credenciales) están implementados en `Code.gs` o en archivos `.gs` adicionales no listados aquí.

---

## 🛠️ Stack Tecnológico

| Capa | Tecnología |
|---|---|
| **Backend** | Google Apps Script (runtime V8) |
| **Base de datos** | Google Sheets (cada spreadsheet = módulo) |
| **Almacenamiento** | Google Drive (archivos adjuntos) |
| **Frontend** | HTML + TailwindCSS (CDN) + JavaScript vanilla |
| **Iconos** | Material Icons Round (CDN) |
| **Alertas** | SweetAlert2 |
| **QR** | QuickChart API (`https://quickchart.io/qr?size=300&text={url}`) |
| **Audio** | SLIMSound Engine (Web Audio API integrada) |
| **Comunicación** | `google.script.run` (frontend → backend asíncrono) |
| **Emails** | `GmailApp` / `MailApp` con plantillas HTML estilizadas |
| **Caché** | `CacheService` de Google Apps Script (TTL configurable) |
| **Locks** | `LockService` para operaciones concurrentes críticas |

---

## 🗄️ Arquitectura de Bases de Datos

El sistema usa **8 Google Spreadsheets independientes**, uno por módulo.

| Clave CONFIG | Hoja principal | Descripción |
|---|---|---|
| `USUARIOS` | `BD_SLIMAPP` | Registro maestro de ~2.230 socios |
| `JUSTIFICACIONES` | `BD_JUSTIFICACIONES` | Justificaciones de inasistencia a asambleas |
| `APELACIONES` | `BD_APELACIONES` | Apelaciones de multas sindicales |
| `PRESTAMOS` | `BD_PRESTAMOS` | Solicitudes de préstamos |
| `PERMISOS_MEDICOS` | `BD_Permisos medicos` | Permisos médicos laborales |
| `CREDENCIALES` | `IMPRESION` | Estado de credenciales sindicales |
| `ASISTENCIA` | `BD_ASISTENCIA` | Registro de asistencia a asambleas |
| `GAMIFICACION` | `BD_GAMIFICACION` | Progreso y logros SLIM Quest |

### Esquema tabla `BD_SLIMAPP` (Socios)

| Col | Campo | Descripción |
|---|---|---|
| A | `RUT` | RUT sin puntos ni guión |
| B | `RUT VALIDADO` | `RUT VÁLIDO` / `RUT INVÁLIDO` |
| C | `FECHA DE INGRESO` | Fecha de ingreso a la empresa |
| D | `NOMBRE SISTEMA` | Nombre completo en mayúsculas |
| E | `CARGO` | Cargo del trabajador |
| F | `CORREO` | Correo electrónico personal |
| G | `SITE` | Sitio/lugar de trabajo |
| H | `REGION` | Región (debe coincidir exactamente con lista oficial) |
| I | `ROL` | `SOCIO` / `DIRIGENTE` / `ADMIN` / `TESTING` |
| J | `ESTADO` | `ACTIVO` / `DESVINCULADO` |
| K | `TELEFONO` | Teléfono de contacto |
| L | `BANCO` | Banco para pago de préstamos |
| M | `TIPO_CUENTA` | Tipo de cuenta bancaria |
| N | `NUMERO_CUENTA` | Número de cuenta bancaria |
| O | `ID_CREDENCIAL` | Contraseña de acceso (= ID de credencial sindical) |
| P | `QR_TOKEN` | Token único para vinculación de dispositivo QR |

> Hojas auxiliares de `USUARIOS`: `IDENTIFICADOR`, `ENVIADOS`, `PENDIENTES`

---

## 📦 Módulos del Sistema

### Para todos los socios

| # | Módulo | Descripción |
|---|---|---|
| 1 | **👤 Mis Datos** | Datos personales, contacto, datos bancarios, estado credencial |
| 2 | **📜 Contrato Colectivo** | Acordeón de 7 capítulos del contrato colectivo vigente |
| 3 | **📝 Justificaciones** | Justificar inasistencia a asambleas (sistema multi-región) |
| 4 | **⚖️ Apelaciones** | Apelar multas por inasistencia, adjuntar comprobantes |
| 5 | **💰 Préstamos** | Solicitar préstamos sindicales según tipo y antigüedad |
| 6 | **🏥 Permiso Médico** | Registrar permisos médicos, adjuntar documentos |
| 7 | **🧮 Calculadora HE** | Calculadora de horas extraordinarias |
| 8 | **📋 Registro Asistencia** | Registro QR físico o virtual en asambleas |
| 9 | **🎮 SLIM Quest** | Sistema de gamificación sindical (XP, grados, quiz diario) |

### Solo para DIRIGENTE / ADMIN

| Módulo | Descripción |
|---|---|
| **👥 Gestión de Socios** | Ingresar trámites a nombre de otro socio |
| **🪪 Consulta ID / Credencial** | Consultar estado de credencial por RUT |
| **🔗 Dirigentes** | Links a secciones restringidas de Google Sites |

### Solo para ADMIN

| Módulo | Descripción |
|---|---|
| **⚙️ Panel Admin** | Switches de módulos, triggers, configuración regional de justificaciones, cambio de roles |

---

## 🔐 Seguridad y Roles

### Roles de usuario

| Rol | Acceso |
|---|---|
| `SOCIO` | Sus propios módulos personales |
| `DIRIGENTE` | Todo lo de socio + gestión en nombre de socios + consultas avanzadas |
| `ADMIN` | Acceso completo + Panel Admin + switches + cambio de roles |
| `TESTING` | Restringido: solo **Mis Datos** y **SLIM Quest** |

### Reglas de acceso

- Socios con `ESTADO = DESVINCULADO` solo pueden ver **Mis Datos** (excepto ADMIN).
- Todas las funciones sensibles validan el rol con `verificarRolUsuario(rut, rolesPermitidos)`.
- **Login:** RUT como usuario + ID Credencial como contraseña (validación Módulo 11).
- **QR:** requiere vinculación previa del dispositivo vía `QR_Access.html`.
- Archivos subidos a Drive: privados por defecto, permisos concedidos solo a involucrados.

---

## 🌍 Sistema de Justificaciones Multi-Región (v2)

A partir del 23-04-2026, el módulo de justificaciones opera con arquitectura **multi-región**: cada región puede tener su propia actividad configurada de forma independiente en la hoja `CONFIG_JUSTIFICACIONES`.

### Flujo del administrador

1. Activar switch global de justificaciones en Panel Admin → abre modal de configuración.
2. Seleccionar tipo de asamblea (Ordinaria / Extraordinaria) + región → el sistema genera el nombre de actividad automáticamente.
3. Completar: fecha del evento, fecha límite de envío y hora límite.
4. Cada región ocupa una fila independiente en `CONFIG_JUSTIFICACIONES`.
5. Se puede eliminar regiones individualmente o deshabilitar todas con un clic.

### Flujo del socio

1. Al abrir Justificaciones → pestaña Nueva: el sistema consulta la región del socio.
2. Sin actividad para su región → modal de bloqueo informativo.
3. Con actividad activa → banner naranja con nombre de actividad, fecha del evento y plazo.
4. Si ya tiene una justificación enviada/aceptada para esa actividad → modal redirige al historial.
5. Campo `ASAMBLEA` en BD: `YYYY-MM-DD_Nombre Actividad`  
   Ejemplo: `2026-04-26_Asamblea Ordinaria - 07. RM Region Metropolitana - Santiago.`

---

## 🎮 SLIM Quest — Sistema de Gamificación

### Grados (por XP acumulada)

| Grado | XP mínimo | XP máximo | Icono |
|---|---|---|---|
| Aspirante | 0 | 1.500 | 🌱 |
| Aprendiz | 1.501 | 4.500 | ⚙️ |
| Trabajador | 4.501 | 10.000 | 🔩 |
| Defensor | 10.001 | 18.000 | 🛡️ |
| Negociador | 18.001 | 30.000 | ⚖️ |
| Dirigente | 30.001 | ∞ | 🏆 |

### Mecánicas
- **Quiz diario:** preguntas ponderadas según el grado del socio.
- **Racha:** días consecutivos completando el quiz (con bono de XP).
- **Logros:** se desbloquean automáticamente por hitos de actividad sindical.
- **Leaderboard:** top 10 socios por XP total.
- Socios con `ESTADO = DESVINCULADO` tienen el progreso suspendido (historial conservado).

---

## ⚙️ Switches de Módulos

El ADMIN puede habilitar/deshabilitar cada módulo en tiempo real desde el Panel Admin. El estado se lee con `obtenerEstadosSwitchDashboard()` en una sola llamada al cargar.

| Módulo | Mecanismo de control |
|---|---|
| Justificaciones | Hoja `CONFIG_JUSTIFICACIONES` (multi-región) |
| Apelaciones | `PropertiesService` — clave `apelaciones_habilitado` |
| Préstamos | `PropertiesService` — clave `prestamos_habilitado` |
| Permisos Médicos | `PropertiesService` — clave `permisos_medicos_habilitado` |
| Contrato Colectivo | `PropertiesService` — clave `contrato_colectivo_habilitado` |
| Calculadora HE | `PropertiesService` — clave `calculadora_habilitada` |
| Registro Asistencia | `PropertiesService` — clave `asistencia_habilitada` |
| SLIM Quest | `PropertiesService` — clave `slimquest_habilitado` |

---

## 🔄 Triggers Automáticos

| Trigger | Frecuencia | Función | Descripción |
|---|---|---|---|
| Justificaciones | Cada 8 hrs | `verificarCambiosJustificaciones()` | Notifica al socio cambios de estado |
| Apelaciones | Cada 8 hrs | `verificarCambiosApelaciones()` | Notifica cambios de estado |
| Apelaciones — Drive | Cada 1 hr | `procesarPermisosComprobantesDevolucion()` | Otorga permisos Drive pendientes |
| Préstamos | Diario 8 AM | `procesarValidacionPrestamos()` | Procesa validaciones de la directiva |
| Préstamos | Diario 8 AM | `verificarCambiosPrestamos()` | Notifica cambios de estado |
| Asistencia | Diario 20:00 | `verificarNotificacionesAsistencia()` | Envía notificaciones pendientes |
| Credenciales | Diario 8 AM | `verificarCambiosCredenciales()` | Notifica cambios de credencial |
| SLIM Quest | Diario 1 AM | `sincronizarSociosGamificacion()` | Sincroniza socios desde `BD_SLIMAPP` |

---

## 📁 Carpetas Google Drive

| Módulo | Descripción |
|---|---|
| Justificaciones | Documentos de respaldo de justificaciones |
| Apelaciones — Comprobantes | Comprobantes de multa |
| Apelaciones — Liquidaciones | Liquidaciones de sueldo |
| Apelaciones — Devoluciones | Comprobantes de devolución |
| Permisos Médicos | Documentos médicos adjuntos |
| Vestuario — Docs | Documentos módulo vestuario |

---

## 🔑 Referencia Rápida de Funciones Backend

### Autenticación
| Función | Descripción |
|---|---|
| `doGet(e)` | Router principal de la Web App |
| `validarUsuario(rut, password)` | Login con RUT + ID Credencial |
| `obtenerDatosUsuario(rut)` | Datos completos del socio |
| `actualizarDatosContacto(rut, correo, telefono)` | Actualiza correo y teléfono |
| `actualizarDatosBancarios(rut, banco, tipoCuenta, numeroCuenta)` | Actualiza datos bancarios (atómico) |

### Justificaciones
| Función | Descripción |
|---|---|
| `obtenerEstadoSwitchJustificaciones()` | Lee config multi-región. Caché TTL 2 min. |
| `enviarJustificacion(rutGestor, tipo, motivo, archivoData, rutBeneficiario)` | Registra justificación |
| `obtenerHistorialJustificaciones(rut)` | Historial del socio |
| `eliminarJustificacion(id)` | Elimina si estado = `Enviado` |
| `gestionarJustificacion(id, estado, obs, rutGestor)` | Acepta/rechaza (DIRIGENTE/ADMIN) |

### Apelaciones
| Función | Descripción |
|---|---|
| `enviarApelacion(...)` | Registra apelación con archivos adjuntos |
| `obtenerHistorialApelaciones(rut)` | Historial del socio |
| `eliminarApelacion(id)` | Elimina si estado = `Enviado` o `Rechazado` |
| `gestionarApelacion(id, estado, obs, rutGestor)` | Cambia estado y notifica |
| `adjuntarComprobanteDevolucion(id, archivoData)` | Adjunta comprobante de devolución |

### Préstamos
| Función | Descripción |
|---|---|
| `obtenerOpcionesPrestamoSocio(rut)` | Calcula montos disponibles según antigüedad |
| `crearSolicitudPrestamo(rutGestor, tipo, cuotas, medioPago, rutBeneficiario)` | Registra solicitud |
| `obtenerHistorialPrestamos(rut)` | Historial del socio |
| `eliminarSolicitud(id)` | Elimina si estado = `Solicitado` |
| `modificarSolicitudPrestamo(id, cuotas, medio)` | Modifica cuotas y medio de pago |

### Permisos Médicos
| Función | Descripción |
|---|---|
| `solicitarPermisoMedico(datos)` | Registra nueva solicitud |
| `adjuntarDocumentoPermiso(id, archivoData)` | Adjunta documento posterior |
| `obtenerHistorialPermisosMedicos(rut)` | Historial del socio |
| `eliminarPermisoMedico(id)` | Anula si estado = `Solicitado` |
| `gestionarPermisoMedico(id, estado, obs, rutGestor)` | Gestiona (DIRIGENTE/ADMIN) |

### Asistencia
| Función | Descripción |
|---|---|
| `registrarAsistencia(rutIndex, nombreControl, codigoTemporal)` | Registro QR |
| `registrarAsistenciaVirtual(rut, nombreControl)` | Registro virtual (con lock) |
| `obtenerHistorialAsistencia(rut)` | Historial del socio |
| `obtenerAsambleaVirtualActiva()` | Puntos virtuales dentro de su ventana horaria |
| `crearPuntoControl(nombre, tipo)` | Crea punto de control QR o virtual |

### SLIM Quest
| Función | Descripción |
|---|---|
| `obtenerDatosGamificacion(rut)` | Progreso completo del socio |
| `completarQuiz(rut, xpGanado, correctas)` | Procesa resultado (racha, bonos, nivel) |
| `obtenerPreguntasQuiz(rut, cantidad)` | Preguntas ponderadas por grado |
| `getLeaderboard(rut)` | Top 10 socios por XP |
| `sincronizarSociosGamificacion()` | Sincroniza socios desde `BD_SLIMAPP` |
| `desbloquearLogro(rut, nombre, descripcion, icono)` | Desbloquea logro al socio |

### Administración
| Función | Descripción |
|---|---|
| `cambiarRolUsuario(rutAdmin, rutObjetivo, nuevoRol)` | Cambia rol de un socio |
| `obtenerEstadosSwitchDashboard()` | Lee todos los switches en una llamada |
| `verificarRolUsuario(rut, rolesPermitidos)` | Validación de permisos |

### Auxiliares
| Función | Descripción |
|---|---|
| `cleanRut(rut)` | Normaliza RUT (sin puntos, sin guión, minúsculas) |
| `esCorreoValido(correo)` | Valida formato de correo |
| `enviarCorreoEstilizado(...)` | Envía correo HTML con diseño corporativo |
| `getSheet(modulo, hoja)` | Obtiene una hoja de un spreadsheet por clave CONFIG |

---

## 🎨 Convenciones de UI

- **Tema:** Light con acentos de color por módulo (TailwindCSS).
- **Fondo de tarjetas:** clase `glass-card`.
- **Overlay de bloqueo global:** `global-processing-overlay` (z-index 90) — bloquea toda interacción durante operaciones largas.
- **Modal de advertencia de carga:** informa al usuario sobre tiempos de espera en subida de archivos.
- **Sonidos:** motor `SLIMSound Engine` integrado (Web Audio API, sin dependencias externas).
- **Sin frameworks JS** — todo vanilla JavaScript.

---

## 🌎 Regiones del Sistema

```
01. XV. Region de Arica y Parinacota - Arica
02. I Region de Tarapacá - Iquique
03. II. Region de Antofagasta - Antofagasta
04. III Region de Atacama - Copiapó
05. IV Region de Coquimbo - La Serena
06. V Region de Valparaíso - Valparaíso.
07. RM Region Metropolitana - Santiago.
08. VI Region del Lib. Gral. Bdo. O'Higgins - Rancagua.
09. VII Region del Maule - Talca.
10. XVI Region del Ñuble - Chillán
11. VIII Region del Biobío - Concepción.
12. IX Region de Araucanía - Temuco
13. XIV Region de Los Ríos - Valdivia
14. X Region de los Lagos - Puerto Montt.
15. XI. Region de Aysén del General Carlos Ibáñez del Campo - Coyhaique
16. XII Region de Magallanes y la Antártica Chilena - Punta Arenas.
```

> ⚠️ El valor de región en `BD_SLIMAPP` debe coincidir **exactamente** con los valores de esta lista.

---

## 🏛️ Directiva Sindicato SLIM N°3

| Cargo | Nombre |
|---|---|
| Presidente | Carlos Orellana G. |
| Tesorero | Franco Collao V. |
| Secretario | Carlos Pacheco M. |
| Directora | Felicita Anartes C. |
| Directora | Sofía Leonardini C. |

**Oficina:** Ahumada 312, Of. 323, Santiago  
**Correo Representante Legal:** `juancarlos.pacheco@cl.issworld.com`

---

## 👨‍💻 Desarrollo

| | |
|---|---|
| **Organización** | Sindicato SLIM N°3 |
| **Desarrollador** | Alejandro Peñailillo G. — DUOC UC, Técnico Analista Programador |
| **Repositorio** | `eLj4n0/Sistema-SLIMAPP-Backend` |
| **Rama principal** | `main` |
| **Plataforma** | Google Apps Script + Google Workspace |

---

*Proyecto privado — Todos los derechos reservados*
