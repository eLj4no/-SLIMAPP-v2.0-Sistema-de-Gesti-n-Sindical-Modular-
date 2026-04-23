# 🏢 SLIMAPP — Sistema de Gestión Sindical
### Sindicato SLIM N°3

---

## 📋 Descripción

**SLIMAPP** es una aplicación web desarrollada para la gestión integral del **Sindicato SLIM N°3**, construida sobre **Google Apps Script**, **Google Sheets** como base de datos y **Google Drive** para almacenamiento de archivos. Se sirve como **Google Web App** y es accesible desde cualquier dispositivo con navegador.

Actualmente gestiona aproximadamente **2.230 socios activos** distribuidos en distintas regiones de Chile.

---

## 📁 Estructura del Proyecto

```
SLIMAPP/
├── Code.gs              # Backend completo (Google Apps Script)
├── Index.html           # Frontend principal (HTML + CSS + JS en un solo archivo)
├── QR_Access.html       # Página de vinculación QR personal del socio
├── QR_Asistencia.html   # Página de registro QR en punto de control
└── README.md            # Este archivo
```

---

## 🗄️ Arquitectura de Base de Datos

El sistema utiliza **8 Google Spreadsheets independientes**, cada uno dedicado a un módulo.

| Clave CONFIG | ID Spreadsheet | Hoja principal | Descripción |
|---|---|---|---|
| `USUARIOS` | `1m7KLd3b3BzKOAI10I5E32MVf_L34XWAGFonhTg37TVM` | `BD_SLIMAPP` | Registro maestro de socios |
| `JUSTIFICACIONES` | `1Hwbly__MXjl9uwJb-spXdah-R3v9SAMOCFHem92uOUg` | `BD_JUSTIFICACIONES` | Justificaciones de inasistencia |
| `APELACIONES` | `11nrvVsf84THWQ7j6NfAr_unyIcBV7aykxACS8R27PwE` | `BD_APELACIONES` | Apelaciones de multas |
| `PRESTAMOS` | `1h-_sJD4rOCuMjlfSouP7a6gfoodHyzI4MOBRUyOW5XU` | `BD_PRESTAMOS` | Solicitudes de préstamos |
| `PERMISOS_MEDICOS` | `1VYfm7cOgL3mVfVoI8DubIm8WG2srzQw9a6DtIEs3UMM` | `BD_Permisos medicos` | Permisos médicos |
| `CREDENCIALES` | `1HVyPxdYKuvIybeOCAPwAJaVHwlxEuOik4YW0XOXBE5o` | `IMPRESION` | Estado de credenciales sindicales |
| `ASISTENCIA` | `1SRQ8Mlc6bBdb0mitAfn4I-EUAS4BOrZRbqS9YAmg3Sk` | `BD_ASISTENCIA` | Registro de asistencia a asambleas |
| `GAMIFICACION` | `1SHDIhGv6XOc30Epm4vdusp3QGVD-pWzhIwzeD6iqbXQ` | `BD_GAMIFICACION` | Progreso y logros SLIM Quest |

---

## 📊 Estructura Detallada de cada Base de Datos

### 🗂️ USUARIOS — `BD_SLIMAPP`

Hoja principal: `BD_SLIMAPP` · Hojas auxiliares: `IDENTIFICADOR`, `ENVIADOS`, `PENDIENTES`

| # | Col | Campo | Descripción |
|---|---|---|---|
| 0 | A | `RUT` | RUT del socio (sin puntos ni guión) |
| 1 | B | `RUT VALIDADO` | Resultado de validación: `RUT VÁLIDO` / `RUT INVÁLIDO` |
| 2 | C | `FECHA DE INGRESO` | Fecha de ingreso a la empresa |
| 3 | D | `NOMBRE SISTEMA` | Nombre completo en mayúsculas |
| 4 | E | `CARGO` | Cargo del trabajador (ej. `GUARDIA DE SEGURIDAD`) |
| 5 | F | `CORREO` | Correo electrónico personal |
| 6 | G | `SITE` | Sitio o lugar de trabajo |
| 7 | H | `REGION` | Región del país (ej. `07. RM Region Metropolitana - Santiago.`) |
| 8 | I | `SEXO` | Sexo del socio |
| 9 | J | `ESTADO` | `ACTIVO` / `DESVINCULADO` |
| 10 | K | `DETALLE DESVINCULACION` | Motivo de desvinculación si aplica |
| 11 | L | `ID CREDENCIAL` | Código credencial sindical (ej. `S03678`) — usado como contraseña de login |
| 12 | M | `CORREO REGISTRADO` | Correo con el que el socio se registró en la app |
| 13 | N | `CONTACTO` | Número de teléfono celular (ej. `+56996114947`) |
| 14 | O | `ROL` | Rol en el sistema: `SOCIO` / `DIRIGENTE` / `ADMIN` / `TESTING` |
| 15 | P | `Link Registro` | URL de registro QR individual del socio |
| 16 | Q | `QR Registro` | Fórmula `=IMAGE(...)` con imagen QR del link de registro |
| 17 | R | `BANCO` | Banco del socio (ej. `BANCO ESTADO (Cuenta RUT)`) |
| 18 | S | `TIPO_CUENTA` | Tipo de cuenta bancaria (ej. `CUENTA VISTA`) |
| 19 | T | `NUEMRO_CUENTA` | Número de cuenta bancaria |
| 20 | U | `ESTADO_NEG_COLECT_2026` | Estado en negociación colectiva 2026 |
| 21 | V | `TALLA_POLERA` | Talla de polera para vestuario |
| 22 | W | `TALLA_POLAR` | Talla de polar para vestuario |
| 23 | X | `TALLA_PANTALON` | Talla de pantalón para vestuario |
| 24 | Y | `TALLA_CALZADO` | Número de calzado |
| 25 | Z | `CALZADO_ESPECIAL` | Indica si requiere calzado especial: `SI` / `NO` |
| 26 | AA | `URL_CERT_PIE_DIABETICO` | URL del certificado médico de pie diabético si aplica |

**Hoja `IDENTIFICADOR`** — Espejo de sesión activa del socio (17 columnas: RUT, NOMBRE, CORREO, CONTACTO, ID, REGION, ACTIVO, CODIGO_TEMP, ASAMBLEA, BANCO, TIPO_CUENTA, NUEMRO_CUENTA, CARGO, Talla POLERA, Talla POLAR, Talla PANTALON, TALLA CALZADO)

**Hoja `ENVIADOS`** — Registro de socios que recibieron correo de bienvenida (RUT, NOMBRE, CORREO)

**Hoja `PENDIENTES`** — Socios pendientes de registro en la app (RUT, NOMBRE, CORREO, CELULAR, ID, REGION)

---

### 🗂️ JUSTIFICACIONES — `BD_JUSTIFICACIONES`

Hojas: `BD_JUSTIFICACIONES`, `CONFIG_JUSTIFICACIONES`, `Resp_Justificaciones`

| # | Col | Campo | Descripción |
|---|---|---|---|
| 0 | A | `ID` | UUID único de la solicitud |
| 1 | B | `FECHA` | Fecha y hora de envío de la solicitud |
| 2 | C | `RUT` | RUT del socio |
| 3 | D | `NOMBRE` | Nombre completo del socio |
| 4 | E | `REGION` | Región del socio al momento del envío |
| 5 | F | `MOTIVO` | Tipo de motivo (ej. `Turno Laboral en ISS`, `Licencia Médica`, `Fuerza Mayor`) |
| 6 | G | `ARGUMENTO` | Texto libre con el argumento del socio |
| 7 | H | `RESPALDO` | URL del documento de respaldo en Google Drive |
| 8 | I | `ESTADO` | Estado actual: `Enviado` / `Aceptado` / `Aceptado/Obs` / `Rechazado` |
| 9 | J | `OBSERVACION` | Observación del dirigente al gestionar |
| 10 | K | `NOTIFICACION` | Estado notificación al socio: `Enviado` / `SIN_CORREO` |
| 11 | L | `ASAMBLEA` | Código de actividad — formato nuevo: `YYYY-MM-DD_Nombre Actividad` (ej. `2026-04-26_Asamblea Ordinaria - 07. RM Region Metropolitana - Santiago.`) |
| 12 | M | `Gestion` | `Socio` / `Dirigente` |
| 13 | N | `Dirigente` | Nombre del dirigente gestor (si aplica) |
| 14 | O | `Correo Dirigente` | Correo del dirigente gestor (si aplica) |

**Hoja `CONFIG_JUSTIFICACIONES`** — Configuración multi-región del módulo (estructura actualizada v2):

| Col | Campo | Descripción |
|---|---|---|
| A | `REGION` | Nombre exacto de la región (debe coincidir con `BD_SLIMAPP.REGION`) |
| B | `Habilitado` | Boolean — si el módulo está activo para esa región |
| C | `Fecha Limite` | ISO 8601 — fecha y hora tope para enviar justificaciones |
| D | `Fecha_Evento` | Fecha de la asamblea/evento (`YYYY-MM-DD`) |
| E | `Nombre_Actividad` | Nombre descriptivo (ej. `Asamblea Ordinaria - 07. RM Region Metropolitana - Santiago.`) |

> ⚠️ **Nota de migración:** La primera ejecución de `obtenerEstadoSwitchJustificaciones()` detecta automáticamente si la hoja tiene la estructura antigua (3 columnas) y la migra al nuevo formato de 5 columnas preservando los datos existentes.

---

### 🗂️ APELACIONES — `BD_APELACIONES`

Hojas: `BD_APELACIONES`, `MASIVAS`, `Resp_Apelaciones`

| # | Col | Campo | Descripción |
|---|---|---|---|
| 0 | A | `ID` | UUID único de la apelación |
| 1 | B | `Fecha Solicitud` | Fecha y hora de envío |
| 2 | C | `Rut` | RUT del socio |
| 3 | D | `Nombre` | Nombre completo del socio |
| 4 | E | `Correo` | Correo del socio |
| 5 | F | `Mes Apelación` | Mes al que corresponde la multa apelada |
| 6 | G | `Tipo Motivo` | Tipo de motivo (ej. `Turno Laboral en ISS`, `Fuerza Mayor`) |
| 7 | H | `Detalle Motivo` | Descripción libre del motivo |
| 8 | I | `URL Comprobante` | URL del comprobante en Drive |
| 9 | J | `URL Liquidación` | URL de la liquidación de sueldo en Drive |
| 10 | K | `Estado` | Estado: `Solicitado` / `En Revisión` / `Aceptado` / `Rechazado` / `Aceptado-Obs` / `Pagado` |
| 11 | L | `Observación` | Observación pública del dirigente |
| 12 | M | `Notificado` | Estado de notificación al socio |
| 13 | N | `Gestión` | `Socio` / `Dirigente` |
| 14 | O | `Nombre Dirigente` | Nombre del dirigente gestor |
| 15 | P | `Correo Dirigente` | Correo del dirigente gestor |
| 16 | Q | `URL Comprobante Devolución` | URL del comprobante de devolución de multa |
| 17 | R | `PERMISO_DEVOLUCION` | Estado del permiso Drive: `OK` / `REINTENTO_N` |
| 18 | S | `LOG_PERMISOS` | Log de intentos de otorgamiento de permisos Drive |
| 19 | T | `Observacion Interna` | Nota interna visible solo para dirigentes/admin |
| 20 | U | `SEMANA` | Semana de gestión asignada |

---

### 🗂️ PRÉSTAMOS — `BD_PRESTAMOS`

Hojas: `BD_PRESTAMOS`, `Validación-Prestamos`, `Registros-eliminados`

| # | Col | Campo | Descripción |
|---|---|---|---|
| 0 | A | `ID` | UUID único del préstamo |
| 1 | B | `Fecha` | Fecha y hora de la solicitud |
| 2 | C | `RUT` | RUT del socio |
| 3 | D | `Nombre` | Nombre completo del socio |
| 4 | E | `Correo` | Correo del socio |
| 5 | F | `Tipo Préstamo` | Tipo: `Vacaciones` / `Emergencia` / `Escolaridad` / etc. |
| 6 | G | `Monto` | Monto solicitado (ej. `$150.000`) |
| 7 | H | `Cuotas` | Número de cuotas para devolución |
| 8 | I | `Medio Pago` | Medio de pago: `Cuenta RUT` / `Transferencia` / etc. |
| 9 | J | `Estado` | Estado: `Solicitado` / `Vigente` / `Rechazado` / `Finalizado` |
| 10 | K | `Fecha Termino` | Fecha estimada de término del préstamo |
| 11 | L | `Gestion` | `Socio` / `Dirigente` |
| 12 | M | `Nombre Dirigente` | Nombre del dirigente gestor |
| 13 | N | `Correo Dirigente` | Correo del dirigente gestor |
| 14 | O | `Informe` | Marca de procesamiento: `OK` cuando fue procesado |
| 15 | P | `Observacion` | Observación adicional del dirigente |

---

### 🗂️ PERMISOS MÉDICOS — `BD_Permisos medicos`

| # | Col | Campo | Descripción |
|---|---|---|---|
| 0 | A | `ID` | UUID único del permiso |
| 1 | B | `Fecha Solicitud` | Fecha y hora de la solicitud |
| 2 | C | `Rut` | RUT del socio |
| 3 | D | `Nombre` | Nombre completo del socio |
| 4 | E | `Correo` | Correo del socio |
| 5 | F | `Tipo Permiso` | Descripción completa del tipo de permiso |
| 6 | G | `Fecha Inicio Permiso` | Fecha de inicio del permiso |
| 7 | H | `Motivo/Detalle` | Descripción libre del motivo |
| 8 | I | `URL Documento Respaldo` | URL del documento en Drive (o `Sin documento`) |
| 9 | J | `Estado` | `Solicitado` / `Solicitado con Documento` / `Documento Adjuntado` / `Aceptado` / `Rechazado` |
| 10 | K | `Fecha Subida Documento` | Fecha en que se adjuntó el documento |
| 11 | L | `Notificado Rep. Legal` | Boolean — si se notificó al representante legal |
| 12 | M | `Gestión` | `Socio` / `Dirigente` |
| 13 | N | `Nombre Dirigente` | Nombre del dirigente gestor |
| 14 | O | `Correo Dirigente` | Correo del dirigente gestor |
| 15 | P | `NOTIFICADO_SOCIO` | Boolean — si se notificó al socio |

---

### 🗂️ ASISTENCIA — `BD_ASISTENCIA`

Hojas: `BD_ASISTENCIA`, `PUNTOS_CONTROL`

| # | Col | Campo | Descripción |
|---|---|---|---|
| 0 | A | `FECHA_HORA` | Fecha y hora del registro (`dd/MM/yyyy HH:mm:ss`) |
| 1 | B | `RUT` | RUT del socio |
| 2 | C | `NOMBRE` | Nombre del socio |
| 3 | D | `ASAMBLEA` | Nombre del punto de control |
| 4 | E | `TIPO_ASISTENCIA` | `QR` / `VIRTUAL` / `PRESENCIAL` |
| 5 | F | `GESTION` | `Socio` / `Dirigente` / `Sistema` |
| 6 | G | `CODIGO_TEMP` | Código temporal usado en el registro |
| 7 | H | `NOTIF_CORREO` | `ENVIADO` / `SIN CORREO` / vacío |

**Hoja `PUNTOS_CONTROL`** — Puntos de control QR activos (NOMBRE, LINK, QR, CODIGO_TEMP, HORA_APERTURA, HORA_CIERRE, TIPO)

---

### 🗂️ CREDENCIALES — `BD_CREDENCIALES` (hoja `IMPRESION`)

Las columnas A–K incluyen:
- `RUT` (A), `ID CREDENCIAL` (B), `ESTADO` (C), `NOMBRE BD` (D), `NOMBRE FORMULARIO` (E)
- `CORREO` (F), `ESTADO CREDENCIAL` (G), `REGION` (H), `OBSERVACION` (I)
- `ESTADO ANTERIOR` (J) — usado por `verificarCambiosCredenciales` para detectar cambios
- `EMAIL ESTADO` (K) — registro de último email enviado

**Estados de credencial:** `VIGENTE` / `VENCIDA` / `PENDIENTE` / `REVOCADA` / `ENTREGADO` / `DISPONIBLE` / `SOLICITADO` / `NO VIGENTE` / `DATOS INCORRECTOS` / `REIMPRIMIR`

**Hoja `HISTORIAL_CREDENCIALES`** — Log de cambios de estado (FECHA, RUT, NOMBRE, CORREO, ESTADO ANTERIOR, ESTADO NUEVO, EMAIL ENVIADO)

---

### 🗂️ GAMIFICACIÓN — `BD_GAMIFICACION`

Hojas: `BD_GAMIFICACION`, `BANCO_PREGUNTAS`

| # | Col | Campo | Descripción |
|---|---|---|---|
| 0 | A | `RUT` | RUT limpio del socio |
| 1 | B | `NOMBRE` | Nombre del socio |
| 2 | C | `XP_TOTAL` | Puntos de experiencia acumulados |
| 3 | D | `GRADO` | Grado actual (Aspirante → Dirigente) |
| 4 | E | `LOGROS` | Array JSON de logros desbloqueados |
| 5 | F | `QUIZ_ULTIMO_DIA` | Fecha del último quiz completado |
| 6 | G | `RACHA_ACTUAL` | Días consecutivos completando el quiz |
| 7 | H | `ULTIMA_ACTIVIDAD` | Última fecha de actividad |
| 8 | I | `QUIZ_ULTIMO_DIA` | Alias interno de control de racha |
| 9 | J | `RACHA_MAX` | Racha máxima histórica |
| 10 | K | `ESTADO` | `ACTIVO` / `DESVINCULADO` |
| 11 | L | `QUIZZES_COMPLETADOS` | Contador total de quizzes completados |

---

## 📁 Carpetas Google Drive

| Módulo | Clave CONFIG | Carpeta ID |
|---|---|---|
| Justificaciones | `JUSTIFICACIONES` | `1UD9hQz1FuacSb3QYrahRl7IfvlpKn8v6` |
| Apelaciones — Comprobantes | `APELACIONES_COMPROBANTES` | `15BmK5pf5Txrxdzdrny23S5q35NDxLy4P` |
| Apelaciones — Liquidaciones | `APELACIONES_LIQUIDACIONES` | `1dR7fM6TW99tunNaMZliyvXc-L23nHKVY` |
| Apelaciones — Devoluciones | `APELACIONES_DEVOLUCIONES` | `1LGLKA3fiCJXf2ouIqlxq3jk_ZSxI3IyM` |
| Permisos Médicos | `PERMISOS_MEDICOS` | `1nCYxD5sJLszBBA6s2DquGW8vlKGZp4ty` |
| Vestuario — Documentos | `VESTUARIO_DOCS` | `1A4PVsIn8ndNMXdqnO9GZCovjtNdfr0BI` |

---

## 🔐 Seguridad y Control de Acceso

- Validación de rol en todas las funciones sensibles mediante `verificarRolUsuario(rut, rolesPermitidos)`.
- Socios desvinculados (`ESTADO = DESVINCULADO`) solo pueden acceder a **Mis Datos** (excepto ADMIN).
- Perfil `TESTING`: acceso exclusivo a **Mis Datos** y **SLIM Quest**.
- El acceso QR requiere vinculación previa del dispositivo vía `QR_Access.html`.
- Los archivos subidos a Drive se configuran como **privados** con permisos silenciosos solo para los involucrados.
- Login: RUT como usuario + ID Credencial como contraseña (validación Módulo 11).

---

## 👥 Roles de Usuario

| Rol | Acceso |
|---|---|
| `SOCIO` | Acceso a sus propios módulos personales |
| `DIRIGENTE` | Puede gestionar trámites a nombre de socios + módulo Gestión |
| `ADMIN` | Acceso completo + Panel Admin + configuración de switches |
| `TESTING` | Acceso restringido: solo `Mis Datos` y `SLIM Quest` |

---

## 📦 Módulos del Sistema

Los módulos se cargan en el dashboard en este orden:

1. **👤 Mis Datos** — Datos personales, banco, credencial, región
2. **📜 Contrato Colectivo** — Acordeón de 7 capítulos
3. **📝 Justificaciones** — Justificaciones de inasistencia a asambleas por región
4. **⚖️ Apelaciones** — Apelaciones de multas por inasistencia
5. **💰 Préstamos** — Solicitud de préstamos sindicales
6. **🏥 Permiso Médico** — Permisos médicos laborales
7. **🧮 Calculadora HE** — Calculadora de horas extraordinarias
8. **📋 Registro Asistencia** — Registro QR y presencial/virtual
9. **🎮 SLIM Quest** — Gamificación sindical

### Módulos adicionales (sin switch en dashboard):
- **⚙️ Panel Admin** — Solo ADMIN: switches, triggers, configuración regional justificaciones
- **🪪 Consulta ID / Credencial** — ADMIN/DIRIGENTE: estado de credencial por RUT
- **👥 Gestión de Socios** — DIRIGENTE/ADMIN: trámites a nombre de socios
- **🔗 Dirigentes** — Links a secciones restringidas de Google Sites

---

## 🌍 Sistema de Justificaciones por Región (v2)

A partir del merge `Justificaciones-por-region` (23-04-2026), el módulo de justificaciones opera con una arquitectura **multi-región**. Cada región puede tener su propia actividad configurada de forma independiente.

### Flujo del administrador

1. El switch global de justificaciones en Panel Admin activa el modal de configuración.
2. El admin selecciona: **tipo de asamblea** (Ordinaria / Extraordinaria) + **región** → el sistema genera automáticamente el nombre de la actividad.
3. Completa: fecha del evento, fecha límite de envío y hora límite.
4. Cada región tiene su fila independiente en `CONFIG_JUSTIFICACIONES`.
5. El admin puede eliminar regiones individualmente o deshabilitar todas con un clic.
6. El switch visual refleja si hay al menos una región activa.

### Flujo del socio

1. Al ingresar a Justificaciones → pestaña Nueva, el sistema consulta la región del socio.
2. Si no hay actividad configurada para su región → modal de bloqueo informativo.
3. Si hay actividad activa → banner naranja muestra nombre de la actividad, fecha del evento y plazo.
4. Si ya tiene una justificación enviada/aceptada para esa actividad → modal redirige al historial.
5. Al enviar → modal de carga con aviso de conexión estable.
6. Campo `ASAMBLEA` en la BD: `YYYY-MM-DD_Nombre Actividad` (ej. `2026-04-26_Asamblea Ordinaria - 07. RM Region Metropolitana - Santiago.`).

### Regiones disponibles (listado oficial)

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

> ⚠️ El valor de región en `BD_SLIMAPP` debe coincidir **exactamente** con los valores de esta lista para que la validación funcione correctamente.

---

## ⚙️ Switches de Módulos

| Módulo | Clave PropertiesService | Notas |
|---|---|---|
| Justificaciones | Multi-región en `CONFIG_JUSTIFICACIONES` | Hoja Google Sheets — arquitectura v2 |
| Apelaciones | `apelaciones_habilitado` | PropertiesService |
| Préstamos | `prestamos_habilitado` | PropertiesService |
| Permisos Médicos | `permisos_medicos_habilitado` | PropertiesService |
| Contrato Colectivo | `contrato_colectivo_habilitado` | PropertiesService |
| Calculadora HE | `calculadora_habilitada` | PropertiesService |
| Registro Asistencia | `asistencia_habilitada` | PropertiesService |
| SLIM Quest | `slimquest_habilitado` | PropertiesService |

El estado de todos los switches se carga en una sola llamada mediante `obtenerEstadosSwitchDashboard()`.

---

## 🎨 Stack Frontend

| Tecnología | Uso |
|---|---|
| **TailwindCSS** (CDN) | Framework CSS para todo el diseño |
| **Material Icons Round** (CDN) | Iconografía del sistema |
| **SweetAlert2** | Modales de alerta (sin uso de `.then()` — modales HTML nativos) |
| **QuickChart API** | Generación de QR (`https://quickchart.io/qr?size=300&text={url}`) |
| **SLIMSound Engine** | Motor de sonidos (Web Audio API) integrado en el frontend |
| **`google.script.run`** | Comunicación asíncrona frontend → backend |

---

## 🔑 Referencia de Funciones Backend (Code.gs)

### Autenticación y Usuario

| Función | Descripción |
|---|---|
| `doGet(e)` | Router principal de la Web App |
| `validarUsuario(rut, password)` | Login con RUT + ID Credencial |
| `obtenerDatosUsuario(rut)` | Datos completos del socio desde `BD_SLIMAPP` |
| `actualizarDatosContacto(rut, correo, telefono)` | Actualiza correo y teléfono |
| `actualizarDatosBancarios(rut, banco, tipoCuenta, numeroCuenta)` | Actualiza datos bancarios (atómico) |

### Justificaciones

| Función | Descripción |
|---|---|
| `obtenerEstadoSwitchJustificaciones()` | Lee configuraciones multi-región desde `CONFIG_JUSTIFICACIONES`. Auto-migra formato antiguo. Caché `justif_switch_state_v2` (2 min TTL). |
| `actualizarSwitchJustificaciones(estado, fechaLimite, fechaEvento, region, nombreActividad)` | Crea o actualiza la configuración de una región específica. Sin región: deshabilita todas. |
| `obtenerInfoActividadPorRegion(rut)` | Retorna la actividad activa para la región del socio. Diferencia: `sinRegion`, `regionNoConfigurada`, `vencido`, `habilitado`. |
| `verificarJustificacionActividad(rut)` | Verifica si el socio ya tiene justificación enviada/aceptada para la actividad activa de su región. |
| `eliminarConfigRegionJustificaciones(region)` | Elimina la fila de configuración de una región específica. Solo ADMIN. |
| `verificarDisponibilidadJustificaciones(rut?)` | Valida disponibilidad por región del socio. Sin RUT: verificación global. |
| `validarJustificacionMesActual(rut)` | Valida duplicados usando código `YYYY-MM-DD_NombreActividad` para modo evento. |
| `enviarJustificacion(rutGestor, tipo, motivo, archivoData, rutBeneficiario)` | Registra nueva justificación. Guarda campo `ASAMBLEA` como `fechaEvento_nombreActividad`. |
| `obtenerHistorialJustificaciones(rut)` | Historial de justificaciones del socio. |
| `eliminarJustificacion(id)` | Elimina justificación en estado `Enviado`. |
| `gestionarJustificacion(id, estado, obs, rutGestor)` | Acepta o rechaza justificación (DIRIGENTE/ADMIN). |
| `verificarCambiosJustificaciones()` | Trigger (cada 8 hrs): detecta cambios de estado y notifica al socio. |

### Apelaciones

| Función | Descripción |
|---|---|
| `enviarApelacion(rutGestor, mes, tipo, detalle, comprobante, liquidacion, rutBeneficiario)` | Registra nueva apelación con archivos adjuntos |
| `obtenerHistorialApelaciones(rut)` | Historial de apelaciones del socio |
| `eliminarApelacion(id)` | Elimina apelación en estado `Enviado` o `Rechazado` |
| `gestionarApelacion(id, estado, obs, rutGestor)` | Cambia estado y notifica al socio |
| `adjuntarComprobanteDevolucion(id, archivoData)` | Adjunta comprobante de devolución de multa |
| `verificarCambiosApelaciones()` | Trigger (cada 8 hrs): notifica cambios de estado |
| `procesarPermisosComprobantesDevolucion()` | Trigger (cada 1 hr): otorga permisos Drive pendientes |

### Préstamos

| Función | Descripción |
|---|---|
| `obtenerOpcionesPrestamoSocio(rut)` | Calcula montos disponibles según antigüedad |
| `solicitarPrestamo(datos)` | Registra nueva solicitud de préstamo |
| `obtenerHistorialPrestamos(rut)` | Historial de préstamos del socio |
| `eliminarSolicitud(id)` | Elimina préstamo en estado `Solicitado` |
| `modificarSolicitudPrestamo(id, cuotas, medio)` | Modifica cuotas y medio de pago |
| `procesarValidacionPrestamos()` | Trigger (diario 8 AM): procesa validaciones de la directiva |
| `verificarCambiosPrestamos()` | Trigger (diario 8 AM): notifica cambios de estado |

### Permisos Médicos

| Función | Descripción |
|---|---|
| `solicitarPermisoMedico(datos)` | Registra nueva solicitud de permiso médico |
| `adjuntarDocumentoPermiso(id, archivoData)` | Adjunta documento posterior a la solicitud |
| `obtenerHistorialPermisosMedicos(rut)` | Historial de permisos del socio |
| `eliminarPermisoMedico(id)` | Anula permiso en estado `Solicitado` |
| `gestionarPermisoMedico(id, estado, obs, rutGestor)` | Gestiona el permiso (DIRIGENTE/ADMIN) |

### Asistencia

| Función | Descripción |
|---|---|
| `registrarAsistencia(rut, nombreControl, codigoTemporal)` | Registro de asistencia vía QR |
| `registrarAsistenciaVirtual(rut, nombreControl)` | Registro virtual (sin código temporal, con lock) |
| `obtenerHistorialAsistencia(rut)` | Historial de asistencias del socio |
| `obtenerAsambleaVirtualActiva()` | Retorna puntos virtuales dentro de su ventana horaria |
| `obtenerPuntosControl()` | Lista de puntos de control activos con estado |
| `crearPuntoControl(nombre, tipo)` | Crea nuevo punto de control QR o virtual |
| `eliminarPuntoControl(nombre)` | Elimina punto de control |
| `guardarVentanaPuntoControl(nombre, apertura, cierre)` | Configura ventana horaria de registro |
| `verificarNotificacionesAsistencia()` | Trigger (diario 20:00): envía notificaciones pendientes |

### Credenciales

| Función | Descripción |
|---|---|
| `obtenerEstadoCredencialPorRut(rut)` | Estado de credencial del socio |
| `verificarCambiosCredenciales()` | Trigger (diario 8 AM): detecta cambios y notifica |
| `consultarIdCredencialBackend(rutConsultante, rutBuscado)` | Consulta completa de credencial (DIRIGENTE/ADMIN) |

### SLIM Quest

| Función | Descripción |
|---|---|
| `obtenerDatosGamificacion(rut)` | Progreso completo del socio |
| `desbloquearLogro(rut, nombre, descripcion, icono)` | Desbloquea logro al socio |
| `completarQuiz(rut, xpGanado, correctas)` | Procesa resultado del quiz (racha, bonos, nivel) |
| `obtenerPreguntasQuiz(rut, cantidad)` | Preguntas ponderadas según grado |
| `getLeaderboard(rut)` | Top 10 socios por XP |
| `sincronizarSociosGamificacion()` | Sincroniza socios desde `BD_SLIMAPP` (trigger diario 1 AM) |

### Auxiliares y Configuración

| Función | Descripción |
|---|---|
| `cleanRut(rut)` | Limpia RUT (elimina puntos, guión, mayúsculas) |
| `esCorreoValido(correo)` | Valida formato de correo |
| `enviarCorreoEstilizado(...)` | Envía correo HTML con diseño corporativo |
| `subirArchivoConPermisos(...)` | Sube archivo a Drive (3 intentos con retry) |
| `validarCorreosParaPermisos(...)` | Prepara lista de correos para acceso Drive |
| `getSheet(ssKey, sheetKey)` | Obtiene hoja por clave de CONFIG |
| `getSpreadsheet(ssKey)` | Obtiene spreadsheet por clave de CONFIG |
| `obtenerUsuarioPorRut(rut)` | Obtiene datos de socio con CacheService (TTL 10 min) |
| `obtenerEstadosSwitchDashboard()` | Estado de todos los módulos en una sola llamada |
| `configurarTriggers()` | Configura todos los triggers automáticos (ejecutar manualmente UNA vez) |
| `generarLinksRegistroYQR()` | Genera links y QR de registro para todos los socios |
| `verificarRolUsuario(rut, rolesPermitidos)` | Valida permisos de rol en funciones sensibles |
| `generarCodigoAsamblea(fecha)` | Genera código `YYYY_MM` para modo mes (fallback) |
| `generarCodigoAsambleaEvento(fechaEvento)` | Genera código `YYYY_MM_DD` para compatibilidad legacy |

---

## 🔁 Triggers Automáticos

| Función | Frecuencia | Descripción |
|---|---|---|
| `verificarCambiosJustificaciones` | Cada 8 horas | Detecta cambios de estado y notifica socios |
| `verificarCambiosApelaciones` | Cada 8 horas | Detecta cambios de estado y notifica socios |
| `procesarValidacionPrestamos` | Diario 8 AM | Procesa validaciones de préstamos de la directiva |
| `procesarPermisosComprobantesDevolucion` | Cada 1 hora | Otorga permisos Drive a comprobantes de devolución |
| `verificarCambiosPrestamos` | Diario 8 AM | Notifica cambios de estado en préstamos |
| `verificarCambiosCredenciales` | Diario 8 AM | Detecta cambios de credencial y notifica socios |
| `verificarNotificacionesAsistencia` | Diario 20:00 | Envía correos de confirmación de asistencia |
| `sincronizarSociosGamificacion` | Diario 1 AM | Sincroniza nuevos socios a BD_GAMIFICACION |

> ⚠️ Ejecutar `configurarTriggers()` manualmente desde el editor de Apps Script para activar todos los triggers. Esta función elimina duplicados antes de recrearlos.

---

## 🔑 Notas Técnicas Importantes

### Manejo de fechas
- Siempre anclar strings de fecha al mediodía (`T12:00:00`) para evitar desfases UTC → Santiago (UTC-3/UTC-4).
- Usar `getValues()` para columnas de fecha (evita inversión día/mes en formato US de `getDisplayValues()`).
- Strings como `"2026-02-25"` se interpretan como medianoche UTC → usar `'T12:00:00'` al construir `Date`.
- El campo `ASAMBLEA` en `BD_JUSTIFICACIONES` usa el formato `YYYY-MM-DD_Nombre Actividad` desde v2.

### Concurrencia y performance
- Usar `getUserLock()` para operaciones iniciadas por el usuario (evita conflicto con triggers que usan `getScriptLock()`).
- `getScriptLock()` en triggers puede retener el lock hasta 60 s — no usar en funciones de usuario.
- `CacheService` TTL 10 min para búsquedas de usuario (`user_RUT`); TTL 2 min para estado del switch de justificaciones (`justif_switch_state_v2`).
- `SpreadsheetApp.flush()` antes de `getDataRange().getValues()` en funciones de validación para evitar race conditions.
- `appendRow` es atómico y siempre escribe tras la última fila con datos — seguro para alta concurrencia.

### Restricciones de Google Apps Script
- **Sin template literals** (backticks `` ` ``) en `Code.gs` — usar concatenación con `+`.
- `Swal.fire().then()` no confiable en iframes sandboxed — usar modales HTML nativos con funciones separadas.
- `innerHTML +=` en loops destruye referencias DOM — pasar `this` en handlers `onclick`.
- `confirm()` nativo del navegador muestra la URL del script — usar modales HTML propios.
- `ScriptApp.getService().getUrl()` no retorna el formato `/a/~/macros/s/` de forma confiable.

### Archivos y Drive
- Eliminar `capture="environment"` de inputs de archivo (fuerza cámara en móvil, impide seleccionar PDF).
- `Drive.Permissions.insert` falla intermitentemente — usar 3 intentos con `Utilities.sleep()`.
- Fórmulas `=IMAGE()` en Sheets: usar `getFormulas()` en lugar de `getValues()` o `getDisplayValues()`.

### Columnas y estructura
- Los índices de columna se referencian numéricamente en `CONFIG.COLUMNAS.*` — renombrar cabeceras no rompe el código, pero insertar/eliminar columnas sí.
- Sentinel `SIN_CORREO`: registros sin correo válido reciben este valor en columnas de notificación para evitar reintentos.
- La hoja `CONFIG_JUSTIFICACIONES` usa la clave `justif_switch_state_v2` en CacheService — al modificar la estructura de la hoja, invalidar manualmente la caché.

### Patrones de navegación y UI
- `navAction` usa cadenas `if/else if` (no `switch/case`) — importante al agregar nuevos módulos.
- Visibilidad de switches se establece dentro del branch `navAction`, no en login ni en `setupDashboard`.
- `add flex` junto a `remove hidden` en `applyRolePermissions` para centrado correcto de tarjetas.
- Overlay global `global-processing-overlay` (z-index 90) bloquea toda interacción durante operaciones de eliminar/guardar.
- Modal `justif-upload-info-modal` y `appeal-upload-info-modal` informan al usuario sobre tiempos de carga y conexión estable.

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

**Organización:** Sindicato SLIM N°3
**Desarrollador:** Alejandro Peñailillo G. — DUOC UC, Técnico Analista Programador
**Repositorio:** `eLj4n0/Sistema-SLIMAPP-Backend`
**Rama principal:** `main`
**Plataforma:** Google Apps Script + Google Workspace

---

*Proyecto privado — Todos los derechos reservados*
