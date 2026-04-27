/**
 * ==============================
 * CONTROL INTERNO - Apps Script
 * Fuente principal: CRONOGRAMA_ANUAL
 * Versión optimizada: carga lazy
 * ==============================
 */

const SHEETS = {
  maestro: "MAESTRO_ACTIVIDADES",
  cronograma: "CRONOGRAMA_ANUAL",
  registro: "REGISTRO_DIARIO",
  usuarios: "USUARIOS",
  configuracion: "CONFIGURACION"
};

const SHEET_ALIASES = {
  MAESTRO_ACTIVIDADES: ["MAESTRO_ACTIVIDADES", "MAESTRO", "MAESTRO ACTIVIDADES"],
  CRONOGRAMA_ANUAL: ["CRONOGRAMA_ANUAL", "CRONOGRAMA", "CRONOGRAMA ANUAL"],
  REGISTRO_DIARIO: ["REGISTRO_DIARIO", "REGISTRO", "REGISTRO DIARIO"],
  USUARIOS: ["USUARIOS", "USUARIO"],
  CONFIGURACION: ["CONFIGURACION", "CONFIGURACIÓN", "PARAMETROS", "PARÁMETROS"]
};

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("Index")
    .setTitle("Control Interno")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * ==============================
 * USUARIOS / LOGIN
 * ==============================
 */

function getUsuariosActivos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shUsuarios = getSheetByAliases_(ss, SHEET_ALIASES.USUARIOS);

  if (!shUsuarios) return [];

  const data = shUsuarios.getDataRange().getValues();
  if (data.length < 2) return [];

  const idx = indexarCabecera_(data[0]);
  const usuarios = [];

  for (let i = 1; i < data.length; i++) {
    if (!esSi_(valorPorCabecera_(data[i], idx, ["Activo", "ACTIVO"]))) continue;

    const nombre = valorPorCabecera_(data[i], idx, ["Usuario", "USUARIO", "Nombre"]);
    if (nombre) usuarios.push(String(nombre).trim());
  }

  return valoresUnicos_(usuarios);
}

function autenticarUsuario(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shUsuarios = getSheetByAliases_(ss, SHEET_ALIASES.USUARIOS);

  validarHojas_([{ nombre: SHEETS.usuarios, hoja: shUsuarios }]);

  const usuarioIngresado = String(payload.usuario || "").trim();
  const claveIngresada = String(payload.clave || "");

  if (!usuarioIngresado || !claveIngresada) {
    throw new Error("Debes ingresar usuario y contraseña.");
  }

  const data = shUsuarios.getDataRange().getValues();
  if (data.length < 2) throw new Error("No hay usuarios configurados.");

  const idx = indexarCabecera_(data[0]);

  for (let i = 1; i < data.length; i++) {
    const activo = esSi_(valorPorCabecera_(data[i], idx, ["Activo", "ACTIVO"]));
    if (!activo) continue;

    const usuario = String(valorPorCabecera_(data[i], idx, ["Usuario", "USUARIO", "Nombre"])).trim();

    const clave = String(
      valorPorCabecera_(data[i], idx, ["Contrasena", "Contraseña", "Clave", "Password"])
    );

    if (usuario === usuarioIngresado && clave === claveIngresada) {
      return {
        ok: true,
        usuario: {
          nombre: usuario,
          rol: valorPorCabecera_(data[i], idx, ["Rol", "ROL"]) || "Usuario",
          correo: valorPorCabecera_(data[i], idx, ["Correo", "CORREO"]) || ""
        }
      };
    }
  }

  throw new Error("Usuario o contraseña incorrectos.");
}

/**
 * ==============================
 * DASHBOARD — solo actividades operativas
 * Carga: vencidas + hoy + próximas 7 días + en ejecución
 * NO devuelve todo el cronograma anual.
 * ==============================
 */
function getDashboardData(usuarioSesion) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const shCronograma = getSheetByAliases_(ss, SHEET_ALIASES.CRONOGRAMA_ANUAL);
  const shUsuarios   = getSheetByAliases_(ss, SHEET_ALIASES.USUARIOS);

  validarHojas_([{ nombre: SHEETS.cronograma, hoja: shCronograma }]);
  asegurarColumnasCronograma_(shCronograma);

  const correo  = Session.getActiveUser().getEmail() || "";
  const usuario = resolverUsuarioSesion_(shUsuarios, correo, usuarioSesion);

  const data = shCronograma.getDataRange().getValues();

  if (data.length < 2) {
    return {
      usuario: { nombre: usuario.nombre, correo: usuario.correo || correo, rol: usuario.rol },
      actividades: [],
      areas: [],
      frecuencias: [],
      responsables: [],
      kpis: { pendientes: 0, vencidas: 0, hoy: 0, completadas: 0 }
    };
  }

  const idx = indexarCabecera_(data[0]);

  const hoy  = new Date(); hoy.setHours(0,0,0,0);
  const en7  = new Date(hoy); en7.setDate(en7.getDate() + 7);

  const actividades   = [];
  let   kPendientes   = 0;
  let   kVencidas     = 0;
  let   kHoy          = 0;
  let   kCompletadas  = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const idCronograma = valorPorCabecera_(row, idx,
      ["ID_CRONOGRAMA", "ID Cronograma", "ID_Cronograma"]);
    if (!idCronograma) continue;

    const responsable = valorPorCabecera_(row, idx, ["RESPONSABLE", "Responsable"]);
    if (!usuarioPuedeVerActividad_(usuario, responsable)) continue;

    const estado = String(valorPorCabecera_(row, idx, ["ESTADO", "Estado"]) || "Pendiente");

    const fechaProgramada = valorPorCabecera_(row, idx,
      ["FECHA_PROGRAMADA", "Fecha_Programada", "Fecha Programada"]);
    const fDate = parseFecha_(fechaProgramada);

    // ── Contadores KPI ──────────────────────────────────────────────────────
    if (estado === "Realizado") {
      kCompletadas++;
    } else if (fDate) {
      const fNorm = new Date(fDate); fNorm.setHours(0,0,0,0);
      if (fNorm < hoy)  kVencidas++;
      if (mismoDia_(fNorm, hoy)) kHoy++;
      if (fNorm <= hoy) kPendientes++;
    }

    // ── Filtro operativo: solo filas relevantes para la vista inicial ────────
    const esOperativa = esActividadOperativa_(fDate, estado, hoy, en7);
    if (!esOperativa) continue;

    const fechaEjecucion  = valorPorCabecera_(row, idx,
      ["FECHA_EJECUCION", "Fecha_Ejecucion", "Fecha Ejecucion"]);
    const evidenciaUrl    = valorPorCabecera_(row, idx,
      ["EVIDENCIA_DRIVE_URL", "Evidencia_Drive_URL", "Evidencia Drive URL"]);

    actividades.push(mapearFila_(row, idx, idCronograma, fechaProgramada, fechaEjecucion, evidenciaUrl, estado));
  }

  return {
    usuario: { nombre: usuario.nombre, correo: usuario.correo || correo, rol: usuario.rol },
    actividades,
    areas:        valoresUnicos_(actividades.map(a => a.Area)),
    frecuencias:  valoresUnicos_(actividades.map(a => a.Frecuencia)),
    responsables: valoresUnicos_(actividades.map(a => a.Responsable)),
    kpis: { pendientes: kPendientes, vencidas: kVencidas, hoy: kHoy, completadas: kCompletadas }
  };
}

/**
 * Actividad operativa = vencida, de hoy, o en los próximos 7 días,
 * o que esté En ejecución. Excluye las ya Realizadas.
 */
function esActividadOperativa_(fDate, estado, hoy, en7) {
  if (estado === "Realizado") return false;

  const ejecucion = normalizarTexto_(estado) === "en_ejecucion" ||
                    normalizarTexto_(estado) === "en ejecucion";
  if (ejecucion) return true;

  if (!fDate) return false;
  const fNorm = new Date(fDate); fNorm.setHours(0,0,0,0);
  return fNorm <= en7;  // vencidas + hoy + próximos 7 días
}

/**
 * ==============================
 * CRONOGRAMA FILTRADO (lazy load)
 * Llamado solo desde el módulo "Todas las Actividades"
 * Filtros: periodo, mes, area, responsable, frecuencia, estado, texto
 * Devuelve máximo 200 filas.
 * ==============================
 */
function getCronogramaFiltrado(filtros, usuarioSesion) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const shCronograma = getSheetByAliases_(ss, SHEET_ALIASES.CRONOGRAMA_ANUAL);
  const shUsuarios   = getSheetByAliases_(ss, SHEET_ALIASES.USUARIOS);

  validarHojas_([{ nombre: SHEETS.cronograma, hoja: shCronograma }]);
  asegurarColumnasCronograma_(shCronograma);

  const correo  = Session.getActiveUser().getEmail() || "";
  const usuario = resolverUsuarioSesion_(shUsuarios, correo, usuarioSesion);

  const data = shCronograma.getDataRange().getValues();
  if (data.length < 2) return { actividades: [], total: 0, truncado: false };

  const idx = indexarCabecera_(data[0]);

  const f = {
    periodo:     String(filtros.periodo     || "").trim(),
    mes:         String(filtros.mes         || "").trim(),
    area:        String(filtros.area        || "").trim(),
    responsable: String(filtros.responsable || "").trim(),
    frecuencia:  String(filtros.frecuencia  || "").trim(),
    estado:      String(filtros.estado      || "").trim(),
    texto:       normalizarTexto_(filtros.texto || "")
  };

  const MAX = 200;
  const actividades = [];
  let totalCoincidencias = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const idCronograma = valorPorCabecera_(row, idx,
      ["ID_CRONOGRAMA", "ID Cronograma", "ID_Cronograma"]);
    if (!idCronograma) continue;

    const responsable = valorPorCabecera_(row, idx, ["RESPONSABLE", "Responsable"]);
    if (!usuarioPuedeVerActividad_(usuario, responsable)) continue;

    const estado      = String(valorPorCabecera_(row, idx, ["ESTADO", "Estado"]) || "Pendiente");
    const area        = String(valorPorCabecera_(row, idx, ["AREA", "Área", "Area"]) || "");
    const frecuencia  = String(valorPorCabecera_(row, idx, ["FRECUENCIA", "Frecuencia"]) || "");
    const periodo     = String(valorPorCabecera_(row, idx, ["PERIODO", "Periodo"]) || "");
    const actividad   = String(valorPorCabecera_(row, idx,
      ["ACTIVIDAD", "Actividad", "Nombre_Actividad"]) || "");

    const fechaProgramada = valorPorCabecera_(row, idx,
      ["FECHA_PROGRAMADA", "Fecha_Programada", "Fecha Programada"]);

    // Filtro mes (extrae del campo FECHA_PROGRAMADA)
    let mesCoincide = true;
    if (f.mes) {
      const fDate = parseFecha_(fechaProgramada);
      if (fDate) {
        const mesFecha = String(fDate.getMonth() + 1).padStart(2, "0");
        mesCoincide = mesFecha === f.mes;
      } else {
        mesCoincide = false;
      }
    }

    const coincide =
      (!f.periodo     || normalizarTexto_(periodo)     === normalizarTexto_(f.periodo))     &&
      mesCoincide                                                                             &&
      (!f.area        || normalizarTexto_(area)         === normalizarTexto_(f.area))        &&
      (!f.responsable || normalizarTexto_(responsable)  === normalizarTexto_(f.responsable)) &&
      (!f.frecuencia  || normalizarTexto_(frecuencia)   === normalizarTexto_(f.frecuencia))  &&
      (!f.estado      || normalizarTexto_(estado)       === normalizarTexto_(f.estado))      &&
      (!f.texto       || normalizarTexto_(actividad).includes(f.texto));

    if (!coincide) continue;

    totalCoincidencias++;

    if (actividades.length < MAX) {
      const fechaEjecucion = valorPorCabecera_(row, idx,
        ["FECHA_EJECUCION", "Fecha_Ejecucion", "Fecha Ejecucion"]);
      const evidenciaUrl = valorPorCabecera_(row, idx,
        ["EVIDENCIA_DRIVE_URL", "Evidencia_Drive_URL", "Evidencia Drive URL"]);

      actividades.push(mapearFila_(row, idx, idCronograma, fechaProgramada, fechaEjecucion, evidenciaUrl, estado));
    }
  }

  return {
    actividades,
    total:    totalCoincidencias,
    truncado: totalCoincidencias > MAX
  };
}

/**
 * ==============================
 * ACTUALIZAR ACTIVIDAD EN CRONOGRAMA
 * ==============================
 */
function actualizarActividad(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shCronograma = getSheetByAliases_(ss, SHEET_ALIASES.CRONOGRAMA_ANUAL);

  validarHojas_([{ nombre: SHEETS.cronograma, hoja: shCronograma }]);
  asegurarColumnasCronograma_(shCronograma);

  const data = shCronograma.getDataRange().getValues();
  if (!data.length) throw new Error("No hay datos en CRONOGRAMA_ANUAL.");

  const idx = indexarCabecera_(data[0]);

  let fila = -1;

  for (let i = 1; i < data.length; i++) {
    const idCronograma = valorPorCabecera_(data[i], idx,
      ["ID_CRONOGRAMA", "ID Cronograma", "ID_Cronograma"]);
    if (String(idCronograma) === String(payload.idRegistro)) {
      fila = i + 1;
      break;
    }
  }

  if (fila === -1) throw new Error("No se encontró el registro en CRONOGRAMA_ANUAL.");

  const estado = String(payload.estado || "").trim();
  const avance = Number(payload.avance || 0);

  if (!estado) throw new Error("El estado es obligatorio.");
  if (Number.isNaN(avance) || avance < 0 || avance > 100)
    throw new Error("El porcentaje de avance debe estar entre 0 y 100.");

  let evidenciaUrl  = "";
  let nombreArchivo = "";
  let fechaCargue   = "";

  if (payload.archivoBase64) {
    const carpetaId = obtenerParametro_("Carpeta_Drive_Evidencias");
    const carpeta   = DriveApp.getFolderById(carpetaId);
    const bytes     = Utilities.base64Decode(payload.archivoBase64);
    const blob      = Utilities.newBlob(
      bytes,
      payload.archivoMime || "application/octet-stream",
      payload.archivoNombre || "evidencia"
    );
    const archivo   = carpeta.createFile(blob);
    evidenciaUrl    = archivo.getUrl();
    nombreArchivo   = archivo.getName();
    fechaCargue     = new Date();
  }

  setValorPorCabecera_(shCronograma, fila, idx, ["ESTADO",       "Estado"],       estado);
  setValorPorCabecera_(shCronograma, fila, idx, ["CUMPLIMIENTO",  "Cumplimiento"], estado === "Realizado" ? "SI" : "NO");
  setValorPorCabecera_(shCronograma, fila, idx, ["FECHA_EJECUCION","Fecha_Ejecucion","Fecha Ejecucion"], new Date());
  setValorPorCabecera_(shCronograma, fila, idx, ["PORCENTAJE_AVANCE","Porcentaje_Avance","Porcentaje Avance"], avance);
  setValorPorCabecera_(shCronograma, fila, idx, ["OBSERVACION",   "Observacion","Observación"], String(payload.observacion || "").trim());

  if (evidenciaUrl) {
    setValorPorCabecera_(shCronograma, fila, idx, ["EVIDENCIA_DRIVE_URL","Evidencia_Drive_URL","Evidencia Drive URL"], evidenciaUrl);
    setValorPorCabecera_(shCronograma, fila, idx, ["NOMBRE_ARCHIVO","Nombre_Archivo","Nombre Archivo"], nombreArchivo);
    setValorPorCabecera_(shCronograma, fila, idx, ["FECHA_CARGUE_EVIDENCIA","Fecha_Cargue_Evidencia","Fecha Cargue Evidencia"], fechaCargue);
  }

  return true;
}

function generarActividadesHoyManual() {
  return {
    ok: true,
    creadas: 0,
    mensaje: "El sistema ahora trabaja sobre CRONOGRAMA_ANUAL. No es necesario generar actividades diarias."
  };
}

/**
 * ==============================
 * COLUMNAS NECESARIAS
 * ==============================
 */
function asegurarColumnasCronograma_(shCronograma) {
  const columnasNecesarias = [
    "FECHA_EJECUCION","PORCENTAJE_AVANCE","OBSERVACION",
    "EVIDENCIA_DRIVE_URL","NOMBRE_ARCHIVO","FECHA_CARGUE_EVIDENCIA",
    "REQUIERE_EVIDENCIA","TIPO_EVIDENCIA","PRIORIDAD"
  ];

  const lastCol = shCronograma.getLastColumn();
  const headers = shCronograma.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(h => String(h).trim());
  const existentes = headers.map(h => normalizarClaveCabecera_(h));

  columnasNecesarias.forEach(col => {
    if (!existentes.includes(normalizarClaveCabecera_(col))) {
      shCronograma.getRange(1, shCronograma.getLastColumn() + 1).setValue(col);
    }
  });
}

/**
 * ==============================
 * CONFIGURACIÓN
 * ==============================
 */
function obtenerParametro_(parametro) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getSheetByAliases_(ss, SHEET_ALIASES.CONFIGURACION);

  validarHojas_([{ nombre: SHEETS.configuracion, hoja: sh }]);

  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(parametro).trim()) {
      return String(data[i][1]).trim();
    }
  }

  throw new Error("No se encontró el parámetro: " + parametro);
}

/**
 * ==============================
 * HELPERS DE MAPEO
 * ==============================
 */
function mapearFila_(row, idx, idCronograma, fechaProgramada, fechaEjecucion, evidenciaUrl, estado) {
  return {
    ID_Registro:           String(idCronograma),
    ID_CRONOGRAMA:         String(idCronograma),
    ID:                    String(valorPorCabecera_(row, idx, ["ID","ID Actividad","Id Actividad"]) || ""),
    Actividad:             String(valorPorCabecera_(row, idx, ["ACTIVIDAD","Actividad","Nombre_Actividad"]) || ""),
    Area:                  String(valorPorCabecera_(row, idx, ["AREA","Área","Area"]) || ""),
    Responsable:           String(valorPorCabecera_(row, idx, ["RESPONSABLE","Responsable"]) || ""),
    Frecuencia:            String(valorPorCabecera_(row, idx, ["FRECUENCIA","Frecuencia"]) || ""),
    Periodo:               String(valorPorCabecera_(row, idx, ["PERIODO","Periodo"]) || ""),
    Fecha_Programada:      formatearFechaISO_(fechaProgramada),
    Fecha_Programada_Texto:formatearFecha_(fechaProgramada),
    Fecha_Ejecucion:       formatearFechaISO_(fechaEjecucion),
    Fecha_Ejecucion_Texto: formatearFecha_(fechaEjecucion),
    Estado:                estado,
    Cumplimiento:          String(valorPorCabecera_(row, idx, ["CUMPLIMIENTO","Cumplimiento"]) || "NO"),
    Porcentaje_Avance:     Number(valorPorCabecera_(row, idx, ["PORCENTAJE_AVANCE","Porcentaje_Avance","Porcentaje Avance"]) || 0),
    Observacion:           String(valorPorCabecera_(row, idx, ["OBSERVACION","Observacion","Observación"]) || ""),
    Evidencia_Drive_URL:   String(evidenciaUrl || ""),
    Evidencia_URL:         String(evidenciaUrl || ""),
    Nombre_Archivo:        String(valorPorCabecera_(row, idx, ["NOMBRE_ARCHIVO","Nombre_Archivo","Nombre Archivo"]) || ""),
    Fecha_Cargue_Evidencia:formatearFechaISO_(valorPorCabecera_(row, idx, ["FECHA_CARGUE_EVIDENCIA","Fecha_Cargue_Evidencia","Fecha Cargue Evidencia"])),
    Requiere_Evidencia:    String(valorPorCabecera_(row, idx, ["REQUIERE_EVIDENCIA","Requiere_Evidencia","Requiere Evidencia"]) || "Sí"),
    Tipo_Evidencia:        String(valorPorCabecera_(row, idx, ["TIPO_EVIDENCIA","Tipo_Evidencia","Tipo Evidencia"]) || "Archivo Drive"),
    Prioridad:             String(valorPorCabecera_(row, idx, ["PRIORIDAD","Prioridad"]) || "Media")
  };
}

/**
 * ==============================
 * UTILIDADES USUARIO
 * ==============================
 */
function resolverUsuarioSesion_(shUsuarios, correo, usuarioSesion) {
  if (usuarioSesion && usuarioSesion.nombre) {
    if (!shUsuarios) return usuarioSesion;
    const usuarioDB = obtenerUsuarioPorNombre_(shUsuarios, usuarioSesion.nombre);
    if (!usuarioDB) return usuarioSesion;
    return usuarioDB;
  }
  if (shUsuarios) return obtenerUsuarioActual_(shUsuarios, correo);
  return { nombre: "Jefe Control Interno", rol: "Jefe", correo: correo || "" };
}

function obtenerUsuarioActual_(shUsuarios, correo) {
  const data = shUsuarios.getDataRange().getValues();
  if (data.length < 2) return { nombre: "Jefe Control Interno", rol: "Jefe", correo: correo || "" };

  const idx = indexarCabecera_(data[0]);

  for (let i = 1; i < data.length; i++) {
    const correoFila = valorPorCabecera_(data[i], idx, ["Correo", "CORREO"]);
    const activo     = valorPorCabecera_(data[i], idx, ["Activo", "ACTIVO"]);

    if (normalizarTexto_(correoFila) === normalizarTexto_(correo) && esSi_(activo)) {
      return {
        nombre: String(valorPorCabecera_(data[i], idx, ["Usuario","USUARIO","Nombre"])).trim(),
        rol:    String(valorPorCabecera_(data[i], idx, ["Rol","ROL"]) || "Usuario").trim(),
        correo: String(correoFila || "").trim()
      };
    }
  }

  return { nombre: "Jefe Control Interno", rol: "Jefe", correo: correo || "" };
}

function obtenerUsuarioPorNombre_(shUsuarios, nombre) {
  const data = shUsuarios.getDataRange().getValues();
  if (data.length < 2) return null;

  const idx = indexarCabecera_(data[0]);

  for (let i = 1; i < data.length; i++) {
    if (!esSi_(valorPorCabecera_(data[i], idx, ["Activo", "ACTIVO"]))) continue;

    const usuario = String(valorPorCabecera_(data[i], idx, ["Usuario","USUARIO","Nombre"])).trim();

    if (usuario === String(nombre || "").trim()) {
      return {
        nombre: usuario,
        rol:    String(valorPorCabecera_(data[i], idx, ["Rol","ROL"]) || "Usuario").trim(),
        correo: String(valorPorCabecera_(data[i], idx, ["Correo","CORREO"]) || "").trim()
      };
    }
  }

  return null;
}

function usuarioPuedeVerActividad_(usuario, responsable) {
  if (normalizarTexto_(usuario.rol) === "jefe") return true;
  const usuarioNom = normalizarTexto_(usuario.nombre);
  const respNom    = normalizarTexto_(responsable);
  if (!respNom) return true;
  return usuarioNom === respNom;
}

/**
 * ==============================
 * UTILIDADES GENERALES
 * ==============================
 */
function getSheetByAliases_(ss, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const sh = ss.getSheetByName(aliases[i]);
    if (sh) return sh;
  }
  return null;
}

function validarHojas_(hojas) {
  const faltantes = hojas.filter(x => !x.hoja).map(x => x.nombre);
  if (faltantes.length) throw new Error("Faltan hojas requeridas: " + faltantes.join(", "));
}

function indexarCabecera_(headers) {
  const idx = {};
  headers.forEach((h, i) => { idx[normalizarClaveCabecera_(h)] = i; });
  return idx;
}

function valorPorCabecera_(row, idx, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const key = normalizarClaveCabecera_(aliases[i]);
    if (Object.prototype.hasOwnProperty.call(idx, key)) return row[idx[key]];
  }
  return "";
}

function columnaPorCabecera_(idx, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const key = normalizarClaveCabecera_(aliases[i]);
    if (Object.prototype.hasOwnProperty.call(idx, key)) return idx[key];
  }
  throw new Error("No se encontró la columna requerida: " + aliases.join(" / "));
}

function setValorPorCabecera_(sheet, fila, idx, aliases, valor) {
  const col = columnaPorCabecera_(idx, aliases);
  sheet.getRange(fila, col + 1).setValue(valor);
}

function valoresUnicos_(arr) {
  return [...new Set(arr.filter(Boolean))].sort();
}

function mismoDia_(a, b) {
  return a.getFullYear() === b.getFullYear() &&
         a.getMonth()    === b.getMonth()    &&
         a.getDate()     === b.getDate();
}

function parseFecha_(valor) {
  if (!valor) return null;
  if (Object.prototype.toString.call(valor) === "[object Date]" && !isNaN(valor)) return valor;
  const texto = String(valor).trim();
  const match = texto.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (match) return new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1]));
  const fecha = new Date(texto);
  if (!isNaN(fecha)) return fecha;
  return null;
}

function formatearFecha_(fecha) {
  const d = parseFecha_(fecha);
  if (!d) return "";
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy");
}

function formatearFechaISO_(fecha) {
  const d = parseFecha_(fecha);
  if (!d) return "";
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function normalizarTexto_(valor) {
  return String(valor || "").toLowerCase().trim()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function normalizarClaveCabecera_(valor) {
  return normalizarTexto_(valor).replace(/\s+/g, "_");
}

function esSi_(valor) {
  return ["si","sí","yes","true","1"].includes(normalizarTexto_(valor));
}

/**
 * ==============================
 * PRUEBA
 * ==============================
 */
function pruebaDashboard() {
  const data = getDashboardData();
  Logger.log(JSON.stringify(data, null, 2));
}

function pruebaCronogramaFiltrado() {
  const data = getCronogramaFiltrado({ area: "Facturación", estado: "Pendiente" });
  Logger.log(JSON.stringify(data, null, 2));
}