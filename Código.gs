/**
 * ==============================
 *  CONTROL INTERNO - Apps Script
 * ==============================
 */

const SHEETS = {
  maestro: "MAESTRO_ACTIVIDADES",
  registro: "REGISTRO_DIARIO",
  usuarios: "USUARIOS",
  configuracion: "CONFIGURACION"
};

const SHEET_ALIASES = {
  MAESTRO_ACTIVIDADES: ["MAESTRO_ACTIVIDADES", "MAESTRO", "MAESTRO ACTIVIDADES"],
  REGISTRO_DIARIO: ["REGISTRO_DIARIO", "REGISTRO", "REGISTRO DIARIO"],
  USUARIOS: ["USUARIOS", "USUARIO"],
  CONFIGURACION: ["CONFIGURACION", "CONFIGURACIÓN", "PARAMETROS", "PARÁMETROS"]
};

function generarActividadesDiarias() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shMaestro = getSheetByAliases_(ss, SHEET_ALIASES[SHEETS.maestro]);
  const shRegistro = getSheetByAliases_(ss, SHEET_ALIASES[SHEETS.registro]);

  validarHojas_([
    { nombre: SHEETS.maestro, hoja: shMaestro },
    { nombre: SHEETS.registro, hoja: shRegistro }
  ]);

  const maestro = shMaestro.getDataRange().getValues();
  const registro = shRegistro.getDataRange().getValues();

  if (maestro.length < 2) return 0;
  if (!registro.length) return 0;

  const idxM = indexarCabecera_(maestro[0]);
  const idxR = indexarCabecera_(registro[0]);

  const hoy = truncarFecha_(new Date());
  const fechaTexto = formatearFechaISO_(hoy);

  const idsExistentes = obtenerClavesRegistro_(registro, idxR);
  const nuevasFilas = [];

  for (let i = 1; i < maestro.length; i++) {
    const fila = maestro[i];

    const id = valorPorCabecera_(fila, idxM, ["ID", "ID Actividad", "Id Actividad"]);
    const actividad = valorPorCabecera_(fila, idxM, ["Actividad", "Nombre_Actividad", "Actividad_Control"]);
    const frecuencia = valorPorCabecera_(fila, idxM, ["Frecuencia"]);
    const responsable = valorPorCabecera_(fila, idxM, ["Responsable", "Encargado"]);
    const regla = valorPorCabecera_(fila, idxM, ["Regla", "Regla_Ejecucion", "Regla Ejecucion"]);
    const area = valorPorCabecera_(fila, idxM, ["Área", "Area"]);
    const activa = valorPorCabecera_(fila, idxM, ["Activa", "Activo"]);
    const requiereEvidencia = valorPorCabecera_(fila, idxM, ["Requiere_Evidencia", "Requiere Evidencia"]);
    const tipoEvidencia = valorPorCabecera_(fila, idxM, ["Tipo_Evidencia", "Tipo Evidencia"]);
    const prioridad = valorPorCabecera_(fila, idxM, ["Prioridad"]);

    if (!id || !esSi_(activa)) continue;
    if (!evaluarReglaEjecucion(frecuencia, regla, hoy)) continue;

    const clave = `${id}_${fechaTexto}`;
    if (idsExistentes.has(clave)) continue;

    const idRegistro = `REG-${id}-${fechaTexto}`;

    nuevasFilas.push([
      idRegistro,
      id,
      actividad,
      hoy,
      "",
      responsable,
      "Pendiente",
      "NO",
      0,
      "",
      "",
      "",
      "",
      requiereEvidencia,
      tipoEvidencia,
      prioridad,
      area
    ]);
  }

  if (nuevasFilas.length) {
    shRegistro
      .getRange(shRegistro.getLastRow() + 1, 1, nuevasFilas.length, nuevasFilas[0].length)
      .setValues(nuevasFilas);
  }

  return nuevasFilas.length;
}

function evaluarReglaEjecucion(frecuencia, regla, fecha) {
  const diaSemana = fecha.getDay();
  const diaMes = fecha.getDate();
  const mes = fecha.getMonth() + 1;
  const ultimoDiaMes = new Date(fecha.getFullYear(), fecha.getMonth() + 1, 0).getDate();

  frecuencia = normalizarTexto_(frecuencia);
  regla = normalizarTexto_(regla);

  if (frecuencia === "diario") return true;

  if (frecuencia === "semanal") {
    const dias = {
      domingo: 0,
      lunes: 1,
      martes: 2,
      miercoles: 3,
      jueves: 4,
      viernes: 5,
      sabado: 6
    };
    return dias[regla] === diaSemana;
  }

  if (frecuencia === "quincenal") {
    return diaMes === 15 || diaMes === ultimoDiaMes;
  }

  if (frecuencia === "mensual") {
    if (regla.includes("dia")) {
      const numero = Number(regla.replace(/\D/g, ""));
      return numero > 0 && diaMes === numero;
    }

    if (regla.includes("ultimo")) return esUltimoDiaHabil(fecha);
    return diaMes === 1;
  }

  if (frecuencia === "bimestral") {
    return mes % 2 === 0 && diaMes === 1;
  }

  if (frecuencia === "trimestral") {
    return [3, 6, 9, 12].includes(mes) && diaMes === 1;
  }

  if (frecuencia === "semestral") {
    return [6, 12].includes(mes) && diaMes === 1;
  }

  if (frecuencia === "anual") {
    if (regla.includes("junio")) return mes === 6 && diaMes === 1;
    if (regla.includes("diciembre")) return mes === 12 && diaMes === 1;
    return mes === 12 && diaMes === 1;
  }

  if (frecuencia === "eventual") return false;

  return false;
}

function esUltimoDiaHabil(fecha) {
  const prueba = new Date(fecha);
  prueba.setDate(prueba.getDate() + 1);

  while (prueba.getDay() === 0 || prueba.getDay() === 6) {
    prueba.setDate(prueba.getDate() + 1);
  }

  return prueba.getMonth() !== fecha.getMonth();
}

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("Index")
    .setTitle("Control Interno")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getUsuariosActivos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shUsuarios = getSheetByAliases_(ss, SHEET_ALIASES[SHEETS.usuarios]);

  if (!shUsuarios) return [];

  const data = shUsuarios.getDataRange().getValues();
  if (data.length < 2) return [];

  const idx = indexarCabecera_(data[0]);
  const usuarios = [];

  for (let i = 1; i < data.length; i++) {
    if (!esSi_(valorPorCabecera_(data[i], idx, ["Activo"]))) continue;
    const nombre = valorPorCabecera_(data[i], idx, ["Usuario", "Nombre"]);
    if (nombre) usuarios.push(nombre);
  }

  return valoresUnicos_(usuarios);
}

function autenticarUsuario(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shUsuarios = getSheetByAliases_(ss, SHEET_ALIASES[SHEETS.usuarios]);
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
    const activo = esSi_(valorPorCabecera_(data[i], idx, ["Activo"]));
    if (!activo) continue;

    const usuario = String(valorPorCabecera_(data[i], idx, ["Usuario", "Nombre"])).trim();
    const clave = String(
      valorPorCabecera_(data[i], idx, ["Contrasena", "Contraseña", "Clave", "Password"])
    );

    if (usuario === usuarioIngresado && clave === claveIngresada) {
      return {
        ok: true,
        usuario: {
          nombre: usuario,
          rol: valorPorCabecera_(data[i], idx, ["Rol"]) || "Usuario",
          correo: valorPorCabecera_(data[i], idx, ["Correo"]) || ""
        }
      };
    }
  }

  throw new Error("Usuario o contraseña incorrectos.");
}

function getDashboardData(usuarioSesion) {
  try {
    // Garantiza que lo programado para hoy se genere antes de consultar el dashboard.
    generarActividadesDiarias();
  } catch (e) {
    // Si falla la generación automática, no bloquea la carga del panel.
    Logger.log("Aviso generarActividadesDiarias: " + e.message);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shRegistro = getSheetByAliases_(ss, SHEET_ALIASES[SHEETS.registro]);
  const shMaestro = getSheetByAliases_(ss, SHEET_ALIASES[SHEETS.maestro]);
  const shUsuarios = getSheetByAliases_(ss, SHEET_ALIASES[SHEETS.usuarios]);

  validarHojas_([
    { nombre: SHEETS.registro, hoja: shRegistro },
    { nombre: SHEETS.maestro, hoja: shMaestro }
  ]);

  const correo = Session.getActiveUser().getEmail() || "";
  const usuario = resolverUsuarioSesion_(shUsuarios, correo, usuarioSesion);

  const registro = shRegistro.getDataRange().getValues();
  const maestro = shMaestro.getDataRange().getValues();

  if (!registro.length || !maestro.length) {
    return {
      usuario: { nombre: usuario.nombre, correo: correo, rol: usuario.rol },
      actividades: [],
      areas: [],
      frecuencias: [],
      responsables: []
    };
  }

  const idx = indexarCabecera_(registro[0]);
  const idxM = indexarCabecera_(maestro[0]);

  const mapaFrecuencia = {};
  for (let i = 1; i < maestro.length; i++) {
    const id = valorPorCabecera_(maestro[i], idxM, ["ID", "ID Actividad", "Id Actividad"]);
    mapaFrecuencia[id] = valorPorCabecera_(maestro[i], idxM, ["Frecuencia"]) || "";
  }

  const actividades = [];
  for (let i = 1; i < registro.length; i++) {
    const row = registro[i];
    const idRegistro = valorPorCabecera_(row, idx, ["ID_Registro"]);
    if (!idRegistro) continue;

    const responsable = valorPorCabecera_(row, idx, ["Responsable", "Encargado"]);
    if (!usuarioPuedeVerActividad_(usuario, responsable)) continue;

    actividades.push({
      ID_Registro: idRegistro,
      ID: valorPorCabecera_(row, idx, ["ID", "ID Actividad", "Id Actividad"]),
      Actividad: valorPorCabecera_(row, idx, ["Actividad", "Nombre_Actividad"]),
      Fecha_Programada: valorPorCabecera_(row, idx, ["Fecha_Programada", "Fecha Programada"]),
      Fecha_Programada_Texto: formatearFecha_(
        valorPorCabecera_(row, idx, ["Fecha_Programada", "Fecha Programada"])
      ),
      Fecha_Ejecucion: valorPorCabecera_(row, idx, ["Fecha_Ejecucion", "Fecha Ejecucion"]),
      Responsable: responsable,
      Estado: valorPorCabecera_(row, idx, ["Estado"]),
      Cumplimiento: valorPorCabecera_(row, idx, ["Cumplimiento"]),
      Porcentaje_Avance: valorPorCabecera_(row, idx, ["Porcentaje_Avance", "Porcentaje Avance"]),
      Observacion: valorPorCabecera_(row, idx, ["Observacion", "Observación"]),
      Evidencia_Drive_URL: valorPorCabecera_(row, idx, ["Evidencia_Drive_URL", "Evidencia Drive URL"]),
      Nombre_Archivo: valorPorCabecera_(row, idx, ["Nombre_Archivo", "Nombre Archivo"]),
      Fecha_Cargue_Evidencia: valorPorCabecera_(row, idx, ["Fecha_Cargue_Evidencia", "Fecha Cargue Evidencia"]),
      Requiere_Evidencia: valorPorCabecera_(row, idx, ["Requiere_Evidencia", "Requiere Evidencia"]),
      Tipo_Evidencia: valorPorCabecera_(row, idx, ["Tipo_Evidencia", "Tipo Evidencia"]),
      Prioridad: valorPorCabecera_(row, idx, ["Prioridad"]),
      Area: valorPorCabecera_(row, idx, ["Área", "Area"]),
      Frecuencia: mapaFrecuencia[valorPorCabecera_(row, idx, ["ID", "ID Actividad", "Id Actividad"])] || ""
    });
  }

  // Fallback: si no hay registros aún para el usuario, mostrar actividades activas desde MAESTRO.
  if (!actividades.length) {
    const hoy = truncarFecha_(new Date());
    const hoyTxt = formatearFechaISO_(hoy);

    for (let i = 1; i < maestro.length; i++) {
      const rowM = maestro[i];
      const activa = valorPorCabecera_(rowM, idxM, ["Activa", "Activo"]);
      if (!esSi_(activa)) continue;

      const responsableM = valorPorCabecera_(rowM, idxM, ["Responsable", "Encargado"]);
      if (!usuarioPuedeVerActividad_(usuario, responsableM)) continue;

      const id = valorPorCabecera_(rowM, idxM, ["ID", "ID Actividad", "Id Actividad"]);
      const actividad = valorPorCabecera_(rowM, idxM, ["Actividad", "Nombre_Actividad", "Actividad_Control"]);
      if (!id || !actividad) continue;

      actividades.push({
        ID_Registro: `VIRT-${id}-${hoyTxt}`,
        ID: id,
        Actividad: actividad,
        Fecha_Programada: hoy,
        Fecha_Programada_Texto: formatearFecha_(hoy),
        Fecha_Ejecucion: "",
        Responsable: responsableM,
        Estado: "Pendiente",
        Cumplimiento: "NO",
        Porcentaje_Avance: 0,
        Observacion: "",
        Evidencia_Drive_URL: "",
        Nombre_Archivo: "",
        Fecha_Cargue_Evidencia: "",
        Requiere_Evidencia: valorPorCabecera_(rowM, idxM, ["Requiere_Evidencia", "Requiere Evidencia"]),
        Tipo_Evidencia: valorPorCabecera_(rowM, idxM, ["Tipo_Evidencia", "Tipo Evidencia"]),
        Prioridad: valorPorCabecera_(rowM, idxM, ["Prioridad"]),
        Area: valorPorCabecera_(rowM, idxM, ["Área", "Area"]),
        Frecuencia: valorPorCabecera_(rowM, idxM, ["Frecuencia"])
      });
    }
  }

  return {
    usuario: { nombre: usuario.nombre, correo: correo, rol: usuario.rol },
    actividades: actividades,
    areas: valoresUnicos_(actividades.map(a => a.Area)),
    frecuencias: valoresUnicos_(actividades.map(a => a.Frecuencia)),
    responsables: valoresUnicos_(actividades.map(a => a.Responsable))
  };
}

function generarActividadesHoyManual(usuarioSesion) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shUsuarios = getSheetByAliases_(ss, SHEET_ALIASES[SHEETS.usuarios]);
  const correo = Session.getActiveUser().getEmail() || "";
  const usuario = resolverUsuarioSesion_(shUsuarios, correo, usuarioSesion);

  if (String(usuario.rol || "").trim() !== "Jefe") {
    throw new Error("Solo el rol Jefe puede ejecutar la generación manual.");
  }

  const creadas = generarActividadesDiarias();
  return {
    ok: true,
    creadas: creadas,
    mensaje: creadas
      ? `Se generaron ${creadas} actividad(es) para hoy.`
      : "No hubo actividades nuevas por generar hoy."
  };
}

function actualizarActividad(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shRegistro = getSheetByAliases_(ss, SHEET_ALIASES[SHEETS.registro]);
  validarHojas_([{ nombre: SHEETS.registro, hoja: shRegistro }]);

  const data = shRegistro.getDataRange().getValues();
  if (!data.length) throw new Error("No hay datos en REGISTRO_DIARIO.");

  const idx = indexarCabecera_(data[0]);

  let fila = -1;
  for (let i = 1; i < data.length; i++) {
    if (valorPorCabecera_(data[i], idx, ["ID_Registro"]) === payload.idRegistro) {
      fila = i + 1;
      break;
    }
  }

  if (fila === -1) throw new Error("No se encontró el registro.");

  const estado = String(payload.estado || "").trim();
  const avance = Number(payload.avance);
  if (!estado) throw new Error("El estado es obligatorio.");
  if (Number.isNaN(avance) || avance < 0 || avance > 100) {
    throw new Error("El porcentaje de avance debe estar entre 0 y 100.");
  }

  let evidenciaUrl = "";
  let nombreArchivo = "";
  let fechaCargue = "";

  if (payload.archivoBase64) {
    const carpetaId = obtenerParametro_("Carpeta_Drive_Evidencias");
    const carpeta = DriveApp.getFolderById(carpetaId);

    const bytes = Utilities.base64Decode(payload.archivoBase64);
    const blob = Utilities.newBlob(bytes, payload.archivoMime, payload.archivoNombre);
    const archivo = carpeta.createFile(blob);

    evidenciaUrl = archivo.getUrl();
    nombreArchivo = archivo.getName();
    fechaCargue = new Date();
  }

  shRegistro.getRange(fila, columnaPorCabecera_(idx, ["Fecha_Ejecucion", "Fecha Ejecucion"]) + 1).setValue(new Date());
  shRegistro.getRange(fila, columnaPorCabecera_(idx, ["Estado"]) + 1).setValue(estado);
  shRegistro.getRange(fila, columnaPorCabecera_(idx, ["Cumplimiento"]) + 1).setValue(estado === "Realizado" ? "SI" : "NO");
  shRegistro.getRange(fila, columnaPorCabecera_(idx, ["Porcentaje_Avance", "Porcentaje Avance"]) + 1).setValue(avance);
  shRegistro.getRange(fila, columnaPorCabecera_(idx, ["Observacion", "Observación"]) + 1).setValue(String(payload.observacion || "").trim());

  if (evidenciaUrl) {
    shRegistro.getRange(fila, columnaPorCabecera_(idx, ["Evidencia_Drive_URL", "Evidencia Drive URL"]) + 1).setValue(evidenciaUrl);
    shRegistro.getRange(fila, columnaPorCabecera_(idx, ["Nombre_Archivo", "Nombre Archivo"]) + 1).setValue(nombreArchivo);
    shRegistro.getRange(fila, columnaPorCabecera_(idx, ["Fecha_Cargue_Evidencia", "Fecha Cargue Evidencia"]) + 1).setValue(fechaCargue);
  }

  return true;
}

function obtenerUsuarioActual_(shUsuarios, correo) {
  const data = shUsuarios.getDataRange().getValues();
  const headers = data[0] || [];

  const idxUsuario = headers.indexOf("Usuario");
  const idxCorreo = headers.indexOf("Correo");
  const idxRol = headers.indexOf("Rol");
  const idxActivo = headers.indexOf("Activo");

  for (let i = 1; i < data.length; i++) {
    if (
      normalizarTexto_(data[i][idxCorreo]) === normalizarTexto_(correo) &&
      esSi_(data[i][idxActivo])
    ) {
      return { nombre: data[i][idxUsuario], rol: data[i][idxRol] || "Usuario" };
    }
  }

  return { nombre: "Jefe Control Interno", rol: "Jefe" };
}

function obtenerUsuarioPorNombre_(shUsuarios, nombre) {
  const data = shUsuarios.getDataRange().getValues();
  const idx = indexarCabecera_(data[0] || []);

  for (let i = 1; i < data.length; i++) {
    if (!esSi_(valorPorCabecera_(data[i], idx, ["Activo"]))) continue;
    const usuario = String(valorPorCabecera_(data[i], idx, ["Usuario", "Nombre"])).trim();
    if (usuario === String(nombre || "").trim()) {
      return {
        nombre: usuario,
        rol: valorPorCabecera_(data[i], idx, ["Rol"]) || "Usuario",
        correo: valorPorCabecera_(data[i], idx, ["Correo"]) || ""
      };
    }
  }
  return null;
}

function obtenerParametro_(parametro) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getSheetByAliases_(ss, SHEET_ALIASES[SHEETS.configuracion]);
  validarHojas_([{ nombre: SHEETS.configuracion, hoja: sh }]);

  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(parametro).trim()) {
      return data[i][1];
    }
  }

  throw new Error("No se encontró el parámetro: " + parametro);
}

function valoresUnicos_(arr) {
  return [...new Set(arr.filter(Boolean))].sort();
}

function formatearFecha_(fecha) {
  if (!fecha) return "";
  return Utilities.formatDate(new Date(fecha), Session.getScriptTimeZone(), "dd/MM/yyyy");
}

function pruebaDashboard() {
  const data = getDashboardData();
  Logger.log(JSON.stringify(data, null, 2));
}

/**
 * ==============
 *   Utilitarios
 * ==============
 */

function indexarCabecera_(headers) {
  const idx = {};
  headers.forEach((h, i) => {
    idx[normalizarClaveCabecera_(h)] = i;
  });
  return idx;
}

function valorPorCabecera_(row, idx, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const key = normalizarClaveCabecera_(aliases[i]);
    if (Object.prototype.hasOwnProperty.call(idx, key)) {
      return row[idx[key]];
    }
  }
  return "";
}

function columnaPorCabecera_(idx, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const key = normalizarClaveCabecera_(aliases[i]);
    if (Object.prototype.hasOwnProperty.call(idx, key)) {
      return idx[key];
    }
  }
  throw new Error("No se encontró la columna requerida: " + aliases.join(" / "));
}

function obtenerClavesRegistro_(registro, idxR) {
  const idsExistentes = new Set();

  for (let i = 1; i < registro.length; i++) {
    const idActividad = valorPorCabecera_(registro[i], idxR, ["ID", "ID Actividad", "Id Actividad"]);
    const fechaProg = valorPorCabecera_(registro[i], idxR, ["Fecha_Programada", "Fecha Programada"]);

    if (idActividad && fechaProg) {
      const fecha = formatearFechaISO_(new Date(fechaProg));
      idsExistentes.add(`${idActividad}_${fecha}`);
    }
  }

  return idsExistentes;
}

function formatearFechaISO_(fecha) {
  return Utilities.formatDate(new Date(fecha), Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function truncarFecha_(fecha) {
  const d = new Date(fecha);
  d.setHours(0, 0, 0, 0);
  return d;
}

function normalizarTexto_(valor) {
  return String(valor || "")
    .toLowerCase()
    .trim()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function normalizarClaveCabecera_(valor) {
  return normalizarTexto_(valor).replace(/\s+/g, "_");
}

function esSi_(valor) {
  return ["si", "sí", "yes", "true", "1"].includes(normalizarTexto_(valor));
}

function validarHojas_(hojas) {
  const faltantes = hojas
    .filter(x => !x.hoja)
    .map(x => x.nombre);

  if (faltantes.length) {
    throw new Error("Faltan hojas requeridas: " + faltantes.join(", "));
  }
}

function resolverUsuarioSesion_(shUsuarios, correo, usuarioSesion) {
  if (usuarioSesion && usuarioSesion.nombre) {
    if (!shUsuarios) return usuarioSesion;

    const usuarioDB = obtenerUsuarioPorNombre_(shUsuarios, usuarioSesion.nombre);
    if (!usuarioDB) return usuarioSesion;
    return usuarioDB;
  }

  if (shUsuarios) return obtenerUsuarioActual_(shUsuarios, correo);
  return { nombre: "Jefe Control Interno", rol: "Jefe" };
}

function usuarioPuedeVerActividad_(usuario, responsable) {
  if (normalizarTexto_(usuario.rol) === "jefe") return true;

  const usuarioNom = normalizarTexto_(usuario.nombre);
  const respNom = normalizarTexto_(responsable);

  // Si la actividad no tiene responsable definido, se muestra para evitar panel vacío.
  if (!respNom) return true;

  return usuarioNom === respNom;
}

function getSheetByAliases_(ss, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const sh = ss.getSheetByName(aliases[i]);
    if (sh) return sh;
  }
  return null;
}
