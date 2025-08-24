// =================================================================
// CONFIGURACIÓN GLOBAL
// =================================================================
const MASTER_SHEET_ID = '11agC695WaNjL5Df5XalmrJu5_CvA3H_ZDSGWJHA74jA';

const CONFIG = {
  SHEETS: {
    DATA: 'Data',
    USERS: 'Usuarios',
    CONFIG: 'Config',
    WEEKS: 'Semanas'
  },
  ROLES: {
    ADMIN: 'administrador',
    SUPERVISOR: 'supervisor',
    USER: 'usuario'
  },
  CHANNELS: ['Presencial', 'Telefonico', 'Virtual'],
  CACHE_EXPIRATION_SECONDS: 60 // 1 minuto
};

const COLS = {
  DATA: {
    CANAL: 0, SUBCANAL: 1, FECHA: 2,
    CLIENTES_ATENDIDOS: 3, CLIENTES_PRIMER_CONTACTO: 4, CLIENTES_ANTES_15_MIN: 5,
    CASOS_ESCALADOS: 6, TMO: 7, TRANSACCIONES_REALIZADAS: 8, CLIENTES_NO_ATENDIDOS: 9,
    TPA: 10, TPE: 11,
    NIVEL_SERVICIO_PORCENTAJE: 12, ABANDONO_PORCENTAJE: 13,
    LLAMADAS_RECIBIDAS: 14, LLAMADAS_ATENDIDAS: 15, LLAMADAS_ATENDIDAS_20S: 16, LLAMADAS_ABANDONADAS: 17
  },
  USUARIOS: { USUARIO: 0, CONTRASEÑA: 1, ROL: 2, ACTIVO: 3, SAL: 4 },
  
  CONFIG: { CANAL: 0, SUBCANAL: 1 }
};

// =================================================================
// FUNCIONES PRINCIPALES Y DE SERVICIO
// =================================================================

function doGet(e) {
  ensureInitialSetup();
  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
      .setTitle('Sistema de Indicadores')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// =================================================================
// CONFIGURACIÓN INICIAL
// =================================================================

function ensureInitialSetup() {
  try {
    const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const sheetNames = ss.getSheets().map(s => s.getName());

    if (!sheetNames.includes(CONFIG.SHEETS.USERS)) {
      const sheet = ss.insertSheet(CONFIG.SHEETS.USERS);
      sheet.getRange(1, 1, 1, 5).setValues([['Usuario', 'Contraseña', 'Rol', 'Activo', 'Sal']]).setFontWeight('bold');
      const sal = Utilities.getUuid();
      const contraseñaEncriptada = encriptarContraseña('Admin123', sal);
      sheet.getRange(2, 1, 1, 5).setValues([['admin', contraseñaEncriptada, CONFIG.ROLES.ADMIN, true, sal]]);
      sheet.autoResizeColumns(1, 5);
      Logger.log('Hoja "Usuarios" y usuario admin por defecto han sido creados.');
    }
  } catch (e) {
    Logger.log('Error en ensureInitialSetup: ' + e.message);
  }
}

// =================================================================
// LÓGICA DE INDICADORES
// =================================================================

function parseDateString(fechaStr) {
  if (!fechaStr) return null;
  const parts = fechaStr.split('-');
  return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
}

function guardarIndicadoresPresencialVirtual(canal, proveedor, fecha, clientesAtendidos, clientesPrimerContacto, clientesAntes15, casosEscalados, transacciones, clientesNoAtendidos, tpa, tpe) {
  try {
    const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const dataSheet = ss.getSheetByName(CONFIG.SHEETS.DATA);
    if (!dataSheet) throw new Error('La hoja "Data" no fue encontrada.');

    const fechaObj = parseDateString(fecha);
    const ca = Number(clientesAtendidos);
    const ca15 = Number(clientesAntes15);
    const cna = Number(clientesNoAtendidos);
    const nivelServicio = ca > 0 ? (ca15 / ca) * 100 : 0;
    const abandono = ca > 0 ? (cna / ca) * 100 : 0;

    let tpaSegundos = 0, tpeSegundos = 0;
    if (tpa && tpa.includes(':')) {
        const [min, seg] = tpa.split(':').map(Number);
        if (!isNaN(min) && !isNaN(seg)) tpaSegundos = min * 60 + seg;
    }
    if (tpe && tpe.includes(':')) {
        const [min, seg] = tpe.split(':').map(Number);
        if (!isNaN(min) && !isNaN(seg)) tpeSegundos = min * 60 + seg;
    }

    const fila = new Array(Object.keys(COLS.DATA).length).fill('');
    fila[COLS.DATA.CANAL] = canal;
    fila[COLS.DATA.SUBCANAL] = proveedor;
    fila[COLS.DATA.FECHA] = fechaObj;
    fila[COLS.DATA.CLIENTES_ATENDIDOS] = ca;
    fila[COLS.DATA.CLIENTES_PRIMER_CONTACTO] = Number(clientesPrimerContacto);
    fila[COLS.DATA.CLIENTES_ANTES_15_MIN] = ca15;
    fila[COLS.DATA.CASOS_ESCALADOS] = Number(casosEscalados);
    fila[COLS.DATA.TRANSACCIONES_REALIZADAS] = Number(transacciones);
    fila[COLS.DATA.CLIENTES_NO_ATENDIDOS] = cna;
    fila[COLS.DATA.TPA] = tpaSegundos;
    fila[COLS.DATA.TPE] = tpeSegundos;
    fila[COLS.DATA.NIVEL_SERVICIO_PORCENTAJE] = nivelServicio;
    fila[COLS.DATA.ABANDONO_PORCENTAJE] = abandono;

    dataSheet.appendRow(fila);
    return `Guardado exitosamente (${canal})`;
  } catch (e) {
    Logger.log(`ERROR al guardar (${canal}): ` + e.toString());
    throw new Error('Ocurrió un error en el servidor: ' + e.message);
  }
}

function guardarIndicadoresTelefonico(proveedor, fecha, llamadasRecibidas, llamadasAtendidas, llamadasAtendidas20s, abandono, tmoSeg, clientesAtendidos, clientesPrimerContacto, casosEscalados, transacciones, clientesNoAtendidos) {
  try {
    const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const dataSheet = ss.getSheetByName(CONFIG.SHEETS.DATA);
    if (!dataSheet) throw new Error('La hoja "Data" no fue encontrada.');

    const fechaObj = parseDateString(fecha);
    const la = Number(llamadasAtendidas);
    const la20 = Number(llamadasAtendidas20s);
    const lab = Number(abandono);

    const nivelServicio = la > 0 ? (la20 / la) * 100 : 0;
    const porcAbandono = la > 0 ? (lab / la) * 100 : 0;

    const fila = new Array(Object.keys(COLS.DATA).length).fill('');
    fila[COLS.DATA.CANAL] = 'Telefonico';
    fila[COLS.DATA.SUBCANAL] = proveedor;
    fila[COLS.DATA.FECHA] = fechaObj;
    fila[COLS.DATA.CLIENTES_ATENDIDOS] = Number(clientesAtendidos);
    fila[COLS.DATA.CLIENTES_PRIMER_CONTACTO] = Number(clientesPrimerContacto);
    fila[COLS.DATA.CASOS_ESCALADOS] = Number(casosEscalados);
    fila[COLS.DATA.TMO] = Number(tmoSeg);
    fila[COLS.DATA.TRANSACCIONES_REALIZADAS] = Number(transacciones);
    fila[COLS.DATA.CLIENTES_NO_ATENDIDOS] = Number(clientesNoAtendidos);
    fila[COLS.DATA.NIVEL_SERVICIO_PORCENTAJE] = nivelServicio;
    fila[COLS.DATA.ABANDONO_PORCENTAJE] = porcAbandono;
    fila[COLS.DATA.LLAMADAS_RECIBIDAS] = Number(llamadasRecibidas);
    fila[COLS.DATA.LLAMADAS_ATENDIDAS] = la;
    fila[COLS.DATA.LLAMADAS_ATENDIDAS_20S] = la20;
    fila[COLS.DATA.LLAMADAS_ABANDONADAS] = lab;

    dataSheet.appendRow(fila);
    return 'Guardado exitosamente (Telefónico)';
  } catch (e) {
    Logger.log('ERROR al guardar (Telefónico): ' + e.toString());
    throw new Error('Ocurrió un error en el servidor: ' + e.message);
  }
}

function obtenerSumaPorCanalYFechas(canal, fechaInicio, fechaFin, subcanal) {
    const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const dataSheet = ss.getSheetByName(CONFIG.SHEETS.DATA);
    if (!dataSheet || dataSheet.getLastRow() < 2) return {};

    const allData = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, Object.keys(COLS.DATA).length).getValues();
    const start = parseDateString(fechaInicio);
    const end = parseDateString(fechaFin);

    const filteredData = allData.filter(row => {
        try {
            return row[COLS.DATA.CANAL] === canal &&
                   (!subcanal || row[COLS.DATA.SUBCANAL] === subcanal) &&
                   (new Date(row[COLS.DATA.FECHA]) >= start && new Date(row[COLS.DATA.FECHA]) <= end);
        } catch(e) { return false; }
    });

    const resumen = {
      'Clientes Atendidos': 0, 'Clientes Atendidos Primer Contacto': 0,
      'Clientes Atendidos Antes de 15 Min': 0, 'Casos Escalados': 0,
      'Transacciones Realizadas': 0, 'Clientes No Atendidos': 0,
      'Llamadas Recibidas': 0, 'Llamadas Atendidas': 0,
      'Llamadas Atendidas < 20s': 0, 'Llamadas Abandonadas': 0
    };
    let totalTMO = 0, countTMO = 0, totalTPA = 0, countTPA = 0, totalTPE = 0, countTPE = 0;
    filteredData.forEach(row => {
        resumen['Clientes Atendidos'] += row[COLS.DATA.CLIENTES_ATENDIDOS] || 0;
        resumen['Clientes Atendidos Primer Contacto'] += row[COLS.DATA.CLIENTES_PRIMER_CONTACTO] || 0;
        resumen['Casos Escalados'] += row[COLS.DATA.CASOS_ESCALADOS] || 0;
        resumen['Transacciones Realizadas'] += row[COLS.DATA.TRANSACCIONES_REALIZADAS] || 0;
        resumen['Clientes No Atendidos'] += row[COLS.DATA.CLIENTES_NO_ATENDIDOS] || 0;

        if (canal === 'Telefonico') {
            resumen['Llamadas Recibidas'] += row[COLS.DATA.LLAMADAS_RECIBIDAS] || 0;
            resumen['Llamadas Atendidas'] += row[COLS.DATA.LLAMADAS_ATENDIDAS] || 0;
            resumen['Llamadas Atendidas < 20s'] += row[COLS.DATA.LLAMADAS_ATENDIDAS_20S] || 0;
            resumen['Llamadas Abandonadas'] += row[COLS.DATA.LLAMADAS_ABANDONADAS] || 0;
            if (row[COLS.DATA.TMO] > 0) { totalTMO += row[COLS.DATA.TMO]; countTMO++; }
        } else { // Presencial y Virtual
            resumen['Clientes Atendidos Antes de 15 Min'] += row[COLS.DATA.CLIENTES_ANTES_15_MIN] || 0;
            if (row[COLS.DATA.TPA] > 0) { totalTPA += row[COLS.DATA.TPA]; countTPA++; }
            if (row[COLS.DATA.TPE] > 0) { totalTPE += row[COLS.DATA.TPE]; countTPE++; }
        }
    });
    resumen['TMO (seg)'] = countTMO > 0 ? Math.round(totalTMO / countTMO) : 0;
    resumen['TPA (seg)'] = countTPA > 0 ? Math.round(totalTPA / countTPA) : 0;
    resumen['TPE (seg)'] = countTPE > 0 ? Math.round(totalTPE / countTPE) : 0;

    return resumen;
}

function obtenerResumenYBarraTodosCanales(fechaInicio, fechaFin) {
  const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
  const dataSheet = ss.getSheetByName(CONFIG.SHEETS.DATA);
  if (!dataSheet || dataSheet.getLastRow() < 2) return [];

  const allData = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, Object.keys(COLS.DATA).length).getValues();
  const start = parseDateString(fechaInicio);
  const end = parseDateString(fechaFin);

  const todasLasSemanas = obtenerSemanas();
  const filteredData = allData.filter(row => {
    try {
      const rowDate = new Date(row[COLS.DATA.FECHA]);
      return !isNaN(rowDate.getTime()) && rowDate >= start && rowDate <= end;
    } catch(e) { return false; }
  });
  const reformatDateForLabel = (dateString) => {
      if (!dateString) return '';
      const parts = dateString.split('-');
      return `${parts[2]}/${parts[1]}`;
  }

  const resultados = [];
  CONFIG.CHANNELS.forEach(canal => {
    const canalData = filteredData.filter(row => row[COLS.DATA.CANAL] === canal);

    const resumen = {
        'Llamadas Atendidas': 0, 'Clientes Atendidos': 0, 'Nivel de Servicio': 0, 'Tasa de Abandono': 0,
        'TMO Promedio': 0, 'TPA Promedio': 0, 'TPE Promedio': 0
    };
    let atendidosNS = 0, entrantesNS = 0, abandonadas = 0, atendidasAb = 0;
    let totalTMO = 0, countTMO = 0, totalTPA = 0, countTPA = 0, totalTPE = 0, countTPE = 0;

    canalData.forEach(row => {
        if (canal === 'Telefonico') {
            resumen['Llamadas Atendidas'] += row[COLS.DATA.LLAMADAS_ATENDIDAS] || 0;
            atendidosNS += row[COLS.DATA.LLAMADAS_ATENDIDAS_20S] || 0;
            entrantesNS += row[COLS.DATA.LLAMADAS_ATENDIDAS] || 0;
            abandonadas += row[COLS.DATA.LLAMADAS_ABANDONADAS] || 0;
            atendidasAb += row[COLS.DATA.LLAMADAS_ATENDIDAS] || 0;
            if (row[COLS.DATA.TMO] > 0) { totalTMO += row[COLS.DATA.TMO]; countTMO++; }
        } else { // Presencial y Virtual
            resumen['Clientes Atendidos'] += row[COLS.DATA.CLIENTES_ATENDIDOS] || 0;
            atendidosNS += row[COLS.DATA.CLIENTES_ANTES_15_MIN] || 0;
            entrantesNS += row[COLS.DATA.CLIENTES_ATENDIDOS] || 0;
            abandonadas += row[COLS.DATA.CLIENTES_NO_ATENDIDOS] || 0;
            atendidasAb += row[COLS.DATA.CLIENTES_ATENDIDOS] || 0;
            if (row[COLS.DATA.TPA] > 0) { totalTPA += row[COLS.DATA.TPA]; countTPA++; }
            if (row[COLS.DATA.TPE] > 0) { totalTPE += row[COLS.DATA.TPE]; countTPE++; }
        }
    });
    resumen['Nivel de Servicio'] = entrantesNS > 0 ? (atendidosNS / entrantesNS) * 100 : 0;
    resumen['Tasa de Abandono'] = atendidasAb > 0 ? (abandonadas / atendidasAb) * 100 : 0;
    resumen['TMO Promedio'] = countTMO > 0 ? Math.round(totalTMO / countTMO) : 0;
    resumen['TPA Promedio'] = countTPA > 0 ? Math.round(totalTPA / countTPA) : 0;
    resumen['TPE Promedio'] = countTPE > 0 ? Math.round(totalTPE / countTPE) : 0;

    const mapEvolucionSemanal = new Map();
    const semanasEnRango = todasLasSemanas.filter(semana => {
        const semanaInicio = parseDateString(semana.inicio);
        const semanaFin = parseDateString(semana.fin);
        return semanaFin >= start && semanaInicio <= end;
    });
    semanasEnRango.forEach(semana => {
        const label = `${semana.nombre} (${reformatDateForLabel(semana.inicio)}-${reformatDateForLabel(semana.fin)})`;
        mapEvolucionSemanal.set(label, 0);
    });

    const metricColumn = canal === 'Telefonico' ? COLS.DATA.LLAMADAS_ATENDIDAS : COLS.DATA.CLIENTES_ATENDIDOS;
    canalData.forEach(row => {
        const rowDate = new Date(row[COLS.DATA.FECHA]);
        for (const semana of semanasEnRango) {
            const semanaInicio = parseDateString(semana.inicio);
            const semanaFin = parseDateString(semana.fin);
            if(rowDate >= semanaInicio && rowDate <= semanaFin) {
                const label = `${semana.nombre} (${reformatDateForLabel(semana.inicio)}-${reformatDateForLabel(semana.fin)})`;
                const totalActual = mapEvolucionSemanal.get(label) || 0;
                mapEvolucionSemanal.set(label, totalActual + (row[metricColumn] || 0));
                break;
            }
        }
    });

    const labelsConDatos = [];
    const dataConDatos = [];
    for (const [label, total] of mapEvolucionSemanal.entries()) {
      if (total > 0) {
        labelsConDatos.push(label);
        dataConDatos.push(total);
      }
    }

    const evolucion = {
        labels: labelsConDatos,
        data: dataConDatos
    };

    resultados.push({
      canal: canal,
      resumen: resumen,
      evolucion: evolucion
    });
  });

  return resultados;
}

function obtenerSubcanales(canal) {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = `subcanales_${canal}`;
    const cached = cache.get(cacheKey);
    if (cached != null) return JSON.parse(cached);

    const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const configSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG);
    if (!configSheet || configSheet.getLastRow() < 2) return [];

    const data = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 2).getValues();
    const resultado = data
      .filter(row => row[COLS.CONFIG.CANAL] === canal)
      .map(row => row[COLS.CONFIG.SUBCANAL]);
    cache.put(cacheKey, JSON.stringify(resultado), CONFIG.CACHE_EXPIRATION_SECONDS * 6);
    return resultado;
  } catch (e) {
    Logger.log('Error en obtenerSubcanales: ' + e.toString());
    throw new Error('No se pudo obtener la configuración de subcanales.');
  }
}

function obtenerSemanas() {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'lista_semanas';
    const cached = cache.get(cacheKey);
    if (cached != null) return JSON.parse(cached);

    const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.WEEKS);
    if (!sheet || sheet.getLastRow() < 2) return [];

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    const semanas = data.map(row => ({
        nombre: row[0],
        inicio: new Date(row[1]).toISOString().split('T')[0],
        fin: new Date(row[2]).toISOString().split('T')[0]
    })).filter(s => s.nombre);
    cache.put(cacheKey, JSON.stringify(semanas), CONFIG.CACHE_EXPIRATION_SECONDS * 6);
    return semanas;
  } catch (e) {
    Logger.log('Error en obtenerSemanas: ' + e.toString());
    throw new Error('No se pudo obtener la lista de semanas.');
  }
}

function obtenerRangoTotalFechas() {
  try {
    const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.DATA);
    if (!sheet || sheet.getLastRow() < 2) {
      const hoy = new Date();
      const haceUnMes = new Date();
      haceUnMes.setDate(hoy.getDate() - 30);
      return {
        fechaInicioTotal: haceUnMes.toISOString().split('T')[0],
        fechaFinTotal: hoy.toISOString().split('T')[0]
      };
    }

    const colFechas = sheet.getRange(2, COLS.DATA.FECHA + 1, sheet.getLastRow() - 1, 1).getValues();
    const fechasValidas = colFechas.flat().filter(fecha => fecha instanceof Date && !isNaN(fecha));
    if (fechasValidas.length === 0) {
        const hoy = new Date();
        const haceUnMes = new Date();
        haceUnMes.setDate(hoy.getDate() - 30);
        return {
            fechaInicioTotal: haceUnMes.toISOString().split('T')[0],
            fechaFinTotal: hoy.toISOString().split('T')[0]
        };
    }

    const fechaMin = new Date(Math.min.apply(null, fechasValidas));
    const fechaMax = new Date(Math.max.apply(null, fechasValidas));
    return {
      fechaInicioTotal: fechaMin.toISOString().split('T')[0],
      fechaFinTotal: fechaMax.toISOString().split('T')[0]
    };
  } catch (e) {
    Logger.log('Error en obtenerRangoTotalFechas: ' + e.toString());
    const hoy = new Date();
    return { fechaInicioTotal: hoy.toISOString().split('T')[0], fechaFinTotal: hoy.toISOString().split('T')[0] };
  }
}

// =================================================================
// SEGURIDAD Y GESTIÓN DE USUARIOS
// =================================================================

function encriptarContraseña(contraseña, sal) {
  const text = contraseña + sal;
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text)
    .map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function obtenerHojaUsuarios() {
  const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
  return ss.getSheetByName(CONFIG.SHEETS.USERS);
}

function verificarCredenciales(usuario, contraseña) {
  const sheet = obtenerHojaUsuarios();
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const sal = data[i][COLS.USUARIOS.SAL];
    const contraseñaEncriptada = encriptarContraseña(contraseña, sal);

    if (data[i][COLS.USUARIOS.USUARIO] === usuario && data[i][COLS.USUARIOS.CONTRASEÑA] === contraseñaEncriptada && data[i][COLS.USUARIOS.ACTIVO] === true) {
      return { usuario: data[i][COLS.USUARIOS.USUARIO], rol: data[i][COLS.USUARIOS.ROL] };
    }
  }
  return null;
}

function encontrarFilaUsuario(sheet, usuario) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][COLS.USUARIOS.USUARIO] === usuario) {
      return {
        numFila: i + 1,
        datos: {
          usuario: data[i][COLS.USUARIOS.USUARIO],
          contraseña: data[i][COLS.USUARIOS.CONTRASEÑA],
          rol: data[i][COLS.USUARIOS.ROL],
          activo: data[i][COLS.USUARIOS.ACTIVO],
          sal: data[i][COLS.USUARIOS.SAL]
        }
      };
    }
  }
  return null;
}

function verificarAdmin(usuarioVerificar) {
    const sheet = obtenerHojaUsuarios();
    if (!sheet) return false;
    const infoUsuario = encontrarFilaUsuario(sheet, usuarioVerificar);
    return infoUsuario && infoUsuario.datos.rol === CONFIG.ROLES.ADMIN && infoUsuario.datos.activo;
}

function crearUsuario(usuario, contraseña, rol, adminUser) {
  if (!verificarAdmin(adminUser)) throw new Error('Acción no autorizada.');
  const sheet = obtenerHojaUsuarios();
  if (!sheet) throw new Error('La hoja de Usuarios no existe.');
  const existe = encontrarFilaUsuario(sheet, usuario);
  if (existe) throw new Error('El usuario ya existe');

  const sal = Utilities.getUuid();
  const contraseñaEncriptada = encriptarContraseña(contraseña, sal);
  sheet.appendRow([usuario, contraseñaEncriptada, rol, true, sal]);
  return 'Usuario creado exitosamente';
}

function obtenerUsuarios(adminUser) {
  if (!verificarAdmin(adminUser)) throw new Error('Acción no autorizada.');
  const sheet = obtenerHojaUsuarios();
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  return data.map(row => ({
    usuario: row[COLS.USUARIOS.USUARIO],
    rol: row[COLS.USUARIOS.ROL],
    activo: row[COLS.USUARIOS.ACTIVO]
  }));
}

function cambiarContraseña(usuario, contraseñaActual, nuevaContraseña, invocador) {
  const sheet = obtenerHojaUsuarios();
  if (!sheet) throw new Error('Hoja de usuarios no encontrada');
  const infoUsuario = encontrarFilaUsuario(sheet, usuario);
  if (!infoUsuario) throw new Error('Usuario no encontrado.');

  const esAdmin = verificarAdmin(invocador);
  const contraseñaActualEncriptada = encriptarContraseña(contraseñaActual, infoUsuario.datos.sal);

  if (esAdmin || infoUsuario.datos.contraseña === contraseñaActualEncriptada) {
      const nuevaContraseñaEncriptada = encriptarContraseña(nuevaContraseña, infoUsuario.datos.sal);
      sheet.getRange(infoUsuario.numFila, COLS.USUARIOS.CONTRASEÑA + 1).setValue(nuevaContraseñaEncriptada);
      return 'Contraseña actualizada exitosamente';
  }

  throw new Error('La contraseña actual es incorrecta.');
}

function desactivarUsuario(usuario, adminUser) {
  if (!verificarAdmin(adminUser)) throw new Error('Acción no autorizada.');
  const sheet = obtenerHojaUsuarios();
  if (!sheet) throw new Error('Hoja de usuarios no encontrada');

  const infoUsuario = encontrarFilaUsuario(sheet, usuario);
  if (infoUsuario) {
      sheet.getRange(infoUsuario.numFila, COLS.USUARIOS.ACTIVO + 1).setValue(false);
      return 'Usuario desactivado exitosamente';
  }
  throw new Error('Usuario no encontrado');
}

function activarUsuario(usuario, adminUser) {
  if (!verificarAdmin(adminUser)) throw new Error('Acción no autorizada.');
  const sheet = obtenerHojaUsuarios();
  if (!sheet) throw new Error('Hoja de usuarios no encontrada');

  const infoUsuario = encontrarFilaUsuario(sheet, usuario);
  if (infoUsuario) {
      sheet.getRange(infoUsuario.numFila, COLS.USUARIOS.ACTIVO + 1).setValue(true);
      return 'Usuario activado exitosamente';
  }
  throw new Error('Usuario no encontrado');
}

function eliminarUsuario(usuario, adminUser) {
  if (!verificarAdmin(adminUser)) throw new Error('Acción no autorizada.');
  if (usuario === 'admin') {
      throw new Error('No se puede eliminar el usuario administrador principal.');
  }
  const sheet = obtenerHojaUsuarios();
  if (!sheet) throw new Error('Hoja de usuarios no encontrada');

  const infoUsuario = encontrarFilaUsuario(sheet, usuario);
  if (infoUsuario) {
      sheet.deleteRow(infoUsuario.numFila);
      return 'Usuario eliminado exitosamente';
  }
  throw new Error('Usuario no encontrado');
}

// =================================================================
// GESTIÓN MANUAL DE CACHÉ (DESDE LA WEB APP)
// =================================================================

/**
 * Endpoint seguro llamado desde la aplicación web para limpiar la caché.
 * Verifica que el usuario que invoca la función sea administrador.
 * @param {string} invocador El nombre de usuario que realiza la solicitud.
 * @returns {string} Un mensaje de éxito.
 */
function limpiarCacheManualmente(invocador) {
  if (!verificarAdmin(invocador)) {
    throw new Error('Acción no autorizada. Se requiere rol de administrador.');
  }
  clearSpecificCache();
  return '✅ Caché del sistema limpiada con éxito.';
}

/**
 * Limpia la caché para la lista de semanas y los subcanales.
 * Esto asegura que la próxima vez que se soliciten los datos,
 * se leerán directamente de la hoja de cálculo.
 */
function clearSpecificCache() {
  try {
    const cache = CacheService.getScriptCache();
    const weekCacheKey = 'lista_semanas';

    // 1. Eliminar la caché de la lista de semanas
    cache.remove(weekCacheKey);
    Logger.log(`Caché eliminada para la clave: ${weekCacheKey}`);

    // 2. Construir y eliminar las claves de caché para todos los subcanales
    const subchannelKeys = CONFIG.CHANNELS.map(canal => `subcanales_${canal}`);
    cache.removeAll(subchannelKeys);
    Logger.log(`Caché eliminada para las claves: ${subchannelKeys.join(', ')}`);

  } catch (e) {
    Logger.log('Error al limpiar la caché específica: ' + e.toString());
    throw new Error('Ocurrió un error en el servidor al limpiar la caché.');
  }
}
