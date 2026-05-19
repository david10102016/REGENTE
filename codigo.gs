// ============================================================
// SISTEMA DE GESTIÓN DE TARDANZAS — U.E. Buena Fe
// Google Apps Script — codigo.gs
// Pegar en: script.google.com → nuevo proyecto
// ============================================================

// ── CONFIGURACIÓN GENERAL ──────────────────────────────────
const CONFIG = {
  USUARIO:          'regente',
  PASSWORD:         'buenafe2026',
  UMBRAL_RECURRENTE: 3,          // tardanzas/mes para marcar recurrente
  HORA_MANIANA_VERANO:  7,       // 7:00 am
  HORA_MANIANA_INVIERNO: 8,      // 8:00 am
  HORA_TARDE:       13,          // 13:30 pm
  MIN_TARDE:        30,
  NOMBRE_UE:        'U.E. Buena Fe',
  ANIO:             '2026',
};

// ── NOMBRES DE HOJAS ───────────────────────────────────────
const HOJA_ALUMNOS     = 'Alumnos';
const HOJA_TARDANZAS   = 'Tardanzas';
const HOJA_CONFIG      = 'Configuracion';

// ============================================================
// INICIALIZACIÓN — Crea las 3 hojas si no existen
// Ejecutar manualmente una vez al configurar el proyecto
// ============================================================
function inicializarSistema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Hoja Alumnos ──
  let hAlumnos = ss.getSheetByName(HOJA_ALUMNOS);
  if (!hAlumnos) {
    hAlumnos = ss.insertSheet(HOJA_ALUMNOS);
    const encAlumnos = ['carnet', 'nombre_completo', 'nivel', 'grado', 'paralelo', 'turno'];
    hAlumnos.appendRow(encAlumnos);
    hAlumnos.getRange(1, 1, 1, encAlumnos.length)
      .setFontWeight('bold')
      .setBackground('#1a3a5c')
      .setFontColor('#ffffff');
    hAlumnos.setFrozenRows(1);
    hAlumnos.setColumnWidth(1, 120);
    hAlumnos.setColumnWidth(2, 280);
    hAlumnos.setColumnWidth(3, 110);
    hAlumnos.setColumnWidth(4, 80);
    hAlumnos.setColumnWidth(5, 90);
    hAlumnos.setColumnWidth(6, 90);
    Logger.log('✅ Hoja Alumnos creada');
  }

  // ── Hoja Tardanzas ──
  let hTardanzas = ss.getSheetByName(HOJA_TARDANZAS);
  if (!hTardanzas) {
    hTardanzas = ss.insertSheet(HOJA_TARDANZAS);
    const encTardanzas = ['id', 'fecha', 'hora', 'carnet', 'nombre_completo', 'nivel', 'grado', 'paralelo', 'turno', 'recurrente', 'observacion'];
    hTardanzas.appendRow(encTardanzas);
    hTardanzas.getRange(1, 1, 1, encTardanzas.length)
      .setFontWeight('bold')
      .setBackground('#1a3a5c')
      .setFontColor('#ffffff');
    hTardanzas.setFrozenRows(1);
    hTardanzas.setColumnWidth(1, 140);
    hTardanzas.setColumnWidth(2, 100);
    hTardanzas.setColumnWidth(3, 70);
    hTardanzas.setColumnWidth(4, 110);
    hTardanzas.setColumnWidth(5, 280);
    hTardanzas.setColumnWidth(6, 110);
    hTardanzas.setColumnWidth(7, 80);
    hTardanzas.setColumnWidth(8, 90);
    hTardanzas.setColumnWidth(9, 90);
    hTardanzas.setColumnWidth(10, 100);
    hTardanzas.setColumnWidth(11, 200);
    Logger.log('✅ Hoja Tardanzas creada');
  }

  // ── Hoja Configuracion ──
  let hConfig = ss.getSheetByName(HOJA_CONFIG);
  if (!hConfig) {
    hConfig = ss.insertSheet(HOJA_CONFIG);
    hConfig.appendRow(['clave', 'valor']);
    hConfig.getRange(1, 1, 1, 2)
      .setFontWeight('bold')
      .setBackground('#1a3a5c')
      .setFontColor('#ffffff');
    hConfig.setFrozenRows(1);

    // Valores por defecto
    const defaults = [
      ['invierno',          'false'],
      ['umbral_recurrente', CONFIG.UMBRAL_RECURRENTE.toString()],
      ['nombre_ue',         CONFIG.NOMBRE_UE],
      ['anio',              CONFIG.ANIO],
      ['nombre_regente',    ''],
    ];
    defaults.forEach(row => hConfig.appendRow(row));
    Logger.log('✅ Hoja Configuracion creada');
  }

  Logger.log('✅ Sistema inicializado correctamente');
}

// ============================================================
// HEADERS CORS — requeridos para evitar bloqueos desde Vercel
// ============================================================
function getCorsHeaders() {
  return {
    'Access-Control-Allow-Origin':  '*',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
  };
}

function respuestaJSON(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// doGet — Maneja todas las consultas GET
// ============================================================
function doGet(e) {
  try {
    const action = e.parameter.action || '';

    // ── Login ──
    if (action === 'login') {
      const u = e.parameter.usuario || '';
      const p = e.parameter.password || '';
      if (u === CONFIG.USUARIO && p === CONFIG.PASSWORD) {
        return respuestaJSON({ status: 'ok', mensaje: 'Acceso autorizado' });
      }
      return respuestaJSON({ status: 'error', mensaje: 'Credenciales incorrectas' });
    }

    // ── Obtener alumnos ──
    if (action === 'getAlumnos') {
      return respuestaJSON({ status: 'ok', data: obtenerAlumnos(e.parameter) });
    }

    // ── Obtener tardanzas ──
    if (action === 'getTardanzas') {
      return respuestaJSON({ status: 'ok', data: obtenerTardanzas(e.parameter) });
    }

    // ── Obtener configuración ──
    if (action === 'getConfig') {
      return respuestaJSON({ status: 'ok', data: obtenerConfig() });
    }

    // ── Reporte diario ──
    if (action === 'reporteDiario') {
      return respuestaJSON({ status: 'ok', data: generarReporteDiario(e.parameter.fecha) });
    }

    // ── Reporte mensual ──
    if (action === 'reporteMensual') {
      return respuestaJSON({ status: 'ok', data: generarReporteMensual(e.parameter.mes) });
    }

    // ── Reporte recurrentes ──
    if (action === 'reporteRecurrentes') {
      return respuestaJSON({ status: 'ok', data: generarReporteRecurrentes(e.parameter.mes) });
    }

    // ── Reporte por curso ──
    if (action === 'reporteCurso') {
      return respuestaJSON({ status: 'ok', data: generarReporteCurso(e.parameter.mes) });
    }

    // ── Stats dashboard ──
    if (action === 'getStats') {
      return respuestaJSON({ status: 'ok', data: obtenerStats() });
    }

    return respuestaJSON({ status: 'error', mensaje: 'Acción no reconocida' });

  } catch (err) {
    Logger.log('ERROR doGet: ' + err.toString());
    return respuestaJSON({ status: 'error', mensaje: err.toString() });
  }
}

// ============================================================
// doPost — Maneja escrituras
// ============================================================
function doPost(e) {
  try {
    const datos = JSON.parse(e.postData.contents);
    const action = datos.action || '';

    // ── Registrar tardanza ──
    if (action === 'registrarTardanza') {
      return respuestaJSON(registrarTardanza(datos));
    }

    // ── Importar alumnos (lote desde PDF o Excel) ──
    if (action === 'importarAlumnos') {
      return respuestaJSON(importarAlumnos(datos));
    }

    // ── Agregar alumno manual ──
    if (action === 'agregarAlumno') {
      return respuestaJSON(agregarAlumnoManual(datos));
    }

    // ── Eliminar tardanza ──
    if (action === 'eliminarTardanza') {
      return respuestaJSON(eliminarTardanza(datos.id));
    }

    // ── Guardar configuración ──
    if (action === 'guardarConfig') {
      return respuestaJSON(guardarConfig(datos));
    }

    // ── Guardar observación general ──
    if (action === 'guardarObservacion') {
      return respuestaJSON(guardarObservacion(datos));
    }

    return respuestaJSON({ status: 'error', mensaje: 'Acción no reconocida' });

  } catch (err) {
    Logger.log('ERROR doPost: ' + err.toString());
    return respuestaJSON({ status: 'error', mensaje: err.toString() });
  }
}

// ============================================================
// ALUMNOS
// ============================================================
function obtenerAlumnos(params) {
  const hoja  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_ALUMNOS);
  const filas = hoja.getDataRange().getValues();
  if (filas.length <= 1) return [];

  let alumnos = filas.slice(1).map(f => ({
    carnet:          String(f[0]).trim(),
    nombre_completo: String(f[1]).trim(),
    nivel:           String(f[2]).trim(),
    grado:           String(f[3]).trim(),
    paralelo:        String(f[4]).trim(),
    turno:           String(f[5]).trim(),
  })).filter(a => a.carnet && a.nombre_completo);

  // Filtros opcionales
  if (params && params.nivel)    alumnos = alumnos.filter(a => a.nivel    === params.nivel);
  if (params && params.grado)    alumnos = alumnos.filter(a => a.grado    === params.grado);
  if (params && params.paralelo) alumnos = alumnos.filter(a => a.paralelo === params.paralelo);
  if (params && params.turno)    alumnos = alumnos.filter(a => a.turno    === params.turno);

  return alumnos;
}

function importarAlumnos(datos) {
  const hoja    = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_ALUMNOS);
  const filas   = hoja.getDataRange().getValues();
  const alumnos = datos.alumnos || [];
  const nivel   = datos.nivel   || '';
  const grado   = datos.grado   || '';
  const paralelo= datos.paralelo|| '';
  const turno   = datos.turno   || '';

  if (!alumnos.length) return { status: 'error', mensaje: 'No se recibieron alumnos' };

  // Mapa de carnets existentes → número de fila (1-based, sin encabezado)
  const mapaExistentes = {};
  filas.slice(1).forEach((f, idx) => {
    const carnet = String(f[0]).trim();
    if (carnet) mapaExistentes[carnet] = idx + 2; // +2 por encabezado + 0-based
  });

  let nuevos = 0, actualizados = 0;

  alumnos.forEach(a => {
    const carnet = String(a.carnet || '').trim();
    const nombre = String(a.nombre_completo || '').trim().toUpperCase();
    if (!carnet || !nombre) return;

    const fila = [carnet, nombre, nivel || a.nivel, grado || a.grado, paralelo || a.paralelo, turno || a.turno];

    if (mapaExistentes[carnet]) {
      // Actualizar fila existente
      hoja.getRange(mapaExistentes[carnet], 1, 1, 6).setValues([fila]);
      actualizados++;
    } else {
      // Insertar nuevo
      hoja.appendRow(fila);
      nuevos++;
    }
  });

  return {
    status:  'ok',
    mensaje: `${nuevos} alumnos nuevos, ${actualizados} actualizados`,
    nuevos,
    actualizados,
  };
}

function agregarAlumnoManual(datos) {
  const hoja  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_ALUMNOS);
  const filas = hoja.getDataRange().getValues();
  const carnet = String(datos.carnet || '').trim();

  if (!carnet) return { status: 'error', mensaje: 'El carnet es obligatorio' };

  // Verificar duplicado
  const existe = filas.slice(1).some(f => String(f[0]).trim() === carnet);
  if (existe) return { status: 'error', mensaje: 'Ya existe un alumno con ese carnet' };

  hoja.appendRow([
    carnet,
    String(datos.nombre_completo || '').trim().toUpperCase(),
    datos.nivel    || '',
    datos.grado    || '',
    datos.paralelo || '',
    datos.turno    || '',
  ]);

  return { status: 'ok', mensaje: 'Alumno agregado correctamente' };
}

// ============================================================
// TARDANZAS
// ============================================================
function registrarTardanza(datos) {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const hTardanzas = ss.getSheetByName(HOJA_TARDANZAS);
  const hConfig    = ss.getSheetByName(HOJA_CONFIG);
  const carnet     = String(datos.carnet || '').trim();

  if (!carnet) return { status: 'error', mensaje: 'Carnet requerido' };

  // Obtener umbral desde configuración
  const cfg     = obtenerConfig();
  const umbral  = parseInt(cfg.umbral_recurrente) || CONFIG.UMBRAL_RECURRENTE;

  // Fecha y hora actuales en Bolivia (UTC-4)
  const ahora   = new Date();
  const bolivia  = new Date(ahora.getTime() - (4 * 60 * 60 * 1000));
  const fecha   = Utilities.formatDate(bolivia, 'America/La_Paz', 'yyyy-MM-dd');
  const hora    = Utilities.formatDate(bolivia, 'America/La_Paz', 'HH:mm');

  // Contar tardanzas del mes para determinar recurrente
  const mes       = fecha.slice(0, 7);
  const tardanzas = hTardanzas.getDataRange().getValues();
  const enMes     = tardanzas.slice(1).filter(f =>
    String(f[3]).trim() === carnet &&
    String(f[1]).startsWith(mes)
  ).length;

  const esRecurrente = (enMes + 1) >= umbral;
  const id           = `T${Date.now()}`;

  hTardanzas.appendRow([
    id,
    fecha,
    hora,
    carnet,
    String(datos.nombre_completo || '').trim().toUpperCase(),
    datos.nivel    || '',
    datos.grado    || '',
    datos.paralelo || '',
    datos.turno    || '',
    esRecurrente ? 'SI' : 'NO',
    datos.observacion || '',
  ]);

  return {
    status:     'ok',
    mensaje:    'Tardanza registrada',
    id,
    fecha,
    hora,
    recurrente: esRecurrente,
    totalMes:   enMes + 1,
  };
}

function obtenerTardanzas(params) {
  const hoja  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_TARDANZAS);
  const filas = hoja.getDataRange().getValues();
  if (filas.length <= 1) return [];

  let lista = filas.slice(1).map(f => ({
    id:              String(f[0]),
    fecha:           String(f[1]),
    hora:            String(f[2]),
    carnet:          String(f[3]),
    nombre_completo: String(f[4]),
    nivel:           String(f[5]),
    grado:           String(f[6]),
    paralelo:        String(f[7]),
    turno:           String(f[8]),
    recurrente:      String(f[9]) === 'SI',
    observacion:     String(f[10] || ''),
  })).filter(t => t.id && t.carnet);

  if (params && params.fecha) lista = lista.filter(t => t.fecha === params.fecha);
  if (params && params.mes)   lista = lista.filter(t => t.fecha.startsWith(params.mes));

  return lista.reverse(); // más recientes primero
}

function eliminarTardanza(id) {
  const hoja  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_TARDANZAS);
  const filas = hoja.getDataRange().getValues();

  for (let i = 1; i < filas.length; i++) {
    if (String(filas[i][0]) === String(id)) {
      hoja.deleteRow(i + 1);
      return { status: 'ok', mensaje: 'Tardanza eliminada' };
    }
  }
  return { status: 'error', mensaje: 'Tardanza no encontrada' };
}

// ============================================================
// REPORTES
// ============================================================
function generarReporteDiario(fecha) {
  const hoja  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_TARDANZAS);
  const filas = hoja.getDataRange().getValues();
  const hoy   = fecha || Utilities.formatDate(new Date(), 'America/La_Paz', 'yyyy-MM-dd');
  const cfg   = obtenerConfig();

  const lista = filas.slice(1)
    .filter(f => String(f[1]) === hoy && String(f[0]))
    .map((f, i) => ({
      nro:             i + 1,
      hora:            String(f[2]),
      nombre_completo: String(f[4]),
      nivel:           String(f[5]),
      grado:           String(f[6]),
      paralelo:        String(f[7]),
      turno:           String(f[8]),
      recurrente:      String(f[9]) === 'SI',
    }));

  // Observación del día
  const obs = obtenerObservacionDia(hoy);

  return {
    fecha,
    lista,
    total:       lista.length,
    recurrentes: lista.filter(t => t.recurrente).length,
    observacion: obs,
    config:      cfg,
  };
}

function generarReporteMensual(mes) {
  const hoja  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_TARDANZAS);
  const filas = hoja.getDataRange().getValues();
  const m     = mes || Utilities.formatDate(new Date(), 'America/La_Paz', 'yyyy-MM');

  const lista = filas.slice(1)
    .filter(f => String(f[1]).startsWith(m) && String(f[0]))
    .map((f, i) => ({
      nro:             i + 1,
      fecha:           String(f[1]),
      hora:            String(f[2]),
      nombre_completo: String(f[4]),
      nivel:           String(f[5]),
      grado:           String(f[6]),
      paralelo:        String(f[7]),
      turno:           String(f[8]),
      recurrente:      String(f[9]) === 'SI',
    }));

  return { mes: m, lista, total: lista.length };
}

function generarReporteRecurrentes(mes) {
  const hoja  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_TARDANZAS);
  const filas = hoja.getDataRange().getValues();
  const m     = mes || Utilities.formatDate(new Date(), 'America/La_Paz', 'yyyy-MM');
  const cfg   = obtenerConfig();
  const umbral = parseInt(cfg.umbral_recurrente) || CONFIG.UMBRAL_RECURRENTE;

  // Agrupar por carnet
  const conteo = {};
  filas.slice(1).filter(f => String(f[1]).startsWith(m) && String(f[0])).forEach(f => {
    const carnet = String(f[3]);
    if (!conteo[carnet]) {
      conteo[carnet] = {
        carnet,
        nombre_completo: String(f[4]),
        nivel:           String(f[5]),
        grado:           String(f[6]),
        paralelo:        String(f[7]),
        turno:           String(f[8]),
        total:           0,
      };
    }
    conteo[carnet].total++;
  });

  const lista = Object.values(conteo)
    .filter(a => a.total >= umbral)
    .sort((a, b) => b.total - a.total)
    .map((a, i) => ({ nro: i + 1, ...a }));

  return { mes: m, lista, total: lista.length, umbral };
}

function generarReporteCurso(mes) {
  const hoja  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_TARDANZAS);
  const filas = hoja.getDataRange().getValues();
  const m     = mes || Utilities.formatDate(new Date(), 'America/La_Paz', 'yyyy-MM');

  const grupos = {};
  filas.slice(1).filter(f => String(f[1]).startsWith(m) && String(f[0])).forEach(f => {
    const key = `${f[5]} ${f[6]} ${f[7]}`;
    grupos[key] = (grupos[key] || 0) + 1;
  });

  const lista = Object.entries(grupos)
    .sort((a, b) => b[1] - a[1])
    .map(([curso, total], i) => ({ nro: i + 1, curso, total }));

  return { mes: m, lista };
}

// ============================================================
// ESTADÍSTICAS DASHBOARD
// ============================================================
function obtenerStats() {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const hAlumnos   = ss.getSheetByName(HOJA_ALUMNOS);
  const hTardanzas = ss.getSheetByName(HOJA_TARDANZAS);
  const cfg        = obtenerConfig();
  const umbral     = parseInt(cfg.umbral_recurrente) || CONFIG.UMBRAL_RECURRENTE;

  const hoy = Utilities.formatDate(new Date(), 'America/La_Paz', 'yyyy-MM-dd');
  const mes  = hoy.slice(0, 7);

  const filasAlumnos   = hAlumnos.getDataRange().getValues();
  const filasTardanzas = hTardanzas.getDataRange().getValues();

  const totalAlumnos = filasAlumnos.length - 1;

  const tardanzasHoy = filasTardanzas.slice(1).filter(f =>
    String(f[1]) === hoy && String(f[0])
  );
  const tardanzasMes = filasTardanzas.slice(1).filter(f =>
    String(f[1]).startsWith(mes) && String(f[0])
  );

  // Recurrentes del mes
  const conteo = {};
  tardanzasMes.forEach(f => {
    const c = String(f[3]);
    conteo[c] = (conteo[c] || 0) + 1;
  });
  const recurrentes = Object.values(conteo).filter(v => v >= umbral).length;

  // Últimas 6 tardanzas del día
  const ultimas = tardanzasHoy.slice(-6).reverse().map(f => ({
    hora:            String(f[2]),
    nombre_completo: String(f[4]),
    nivel:           String(f[5]),
    grado:           String(f[6]),
    paralelo:        String(f[7]),
    turno:           String(f[8]),
    recurrente:      String(f[9]) === 'SI',
  }));

  // Top recurrentes del mes
  const topRec = Object.entries(conteo)
    .filter(([, v]) => v >= umbral)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5)
    .map(([carnet, total]) => {
      const fila = filasTardanzas.slice(1).find(f => String(f[3]) === carnet);
      return {
        carnet,
        nombre_completo: fila ? String(fila[4]) : carnet,
        nivel:           fila ? String(fila[5]) : '',
        grado:           fila ? String(fila[6]) : '',
        paralelo:        fila ? String(fila[7]) : '',
        total,
      };
    });

  return {
    totalAlumnos,
    tardanzasHoy:  tardanzasHoy.length,
    tardanzasMes:  tardanzasMes.length,
    recurrentes,
    ultimas,
    topRecurrentes: topRec,
    fecha: hoy,
    mes,
  };
}

// ============================================================
// CONFIGURACIÓN
// ============================================================
function obtenerConfig() {
  const hoja  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_CONFIG);
  const filas = hoja.getDataRange().getValues();
  const cfg   = {};
  filas.slice(1).forEach(f => {
    if (f[0]) cfg[String(f[0]).trim()] = String(f[1]).trim();
  });
  return cfg;
}

function guardarConfig(datos) {
  const hoja  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_CONFIG);
  const filas = hoja.getDataRange().getValues();
  const campos = ['invierno', 'umbral_recurrente', 'nombre_ue', 'anio', 'nombre_regente'];

  campos.forEach(campo => {
    if (datos[campo] === undefined) return;
    let encontrado = false;
    for (let i = 1; i < filas.length; i++) {
      if (String(filas[i][0]).trim() === campo) {
        hoja.getRange(i + 1, 2).setValue(String(datos[campo]));
        encontrado = true;
        break;
      }
    }
    if (!encontrado) hoja.appendRow([campo, String(datos[campo])]);
  });

  return { status: 'ok', mensaje: 'Configuración guardada' };
}

// ============================================================
// OBSERVACIÓN GENERAL DEL DÍA
// Se guarda en la hoja Configuracion con clave obs_YYYY-MM-DD
// ============================================================
function guardarObservacion(datos) {
  const hoja  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_CONFIG);
  const filas = hoja.getDataRange().getValues();
  const hoy   = Utilities.formatDate(new Date(), 'America/La_Paz', 'yyyy-MM-dd');
  const clave = `obs_${datos.fecha || hoy}`;
  const texto = String(datos.observacion || '').trim();

  let encontrado = false;
  for (let i = 1; i < filas.length; i++) {
    if (String(filas[i][0]).trim() === clave) {
      hoja.getRange(i + 1, 2).setValue(texto);
      encontrado = true;
      break;
    }
  }
  if (!encontrado) hoja.appendRow([clave, texto]);

  return { status: 'ok', mensaje: 'Observación guardada' };
}

function obtenerObservacionDia(fecha) {
  const cfg   = obtenerConfig();
  const clave = `obs_${fecha}`;
  return cfg[clave] || '';
}

// ============================================================
// EXPORTAR PDF DEL REPORTE DIARIO
// Genera un PDF desde una hoja temporal y devuelve el link
// ============================================================
function exportarPDFDiario(fecha) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const cfg  = obtenerConfig();
  const data = generarReporteDiario(fecha);
  const hoy  = fecha || Utilities.formatDate(new Date(), 'America/La_Paz', 'yyyy-MM-dd');

  // Crear hoja temporal de impresión
  let hTemp = ss.getSheetByName('_PDF_TEMP');
  if (hTemp) ss.deleteSheet(hTemp);
  hTemp = ss.insertSheet('_PDF_TEMP');

  // Título
  hTemp.getRange(1, 1).setValue(cfg.nombre_ue || CONFIG.NOMBRE_UE);
  hTemp.getRange(1, 1).setFontSize(14).setFontWeight('bold');
  hTemp.getRange(2, 1).setValue(`Reporte de Tardanzas — ${hoy}`);
  hTemp.getRange(3, 1).setValue(`Regente: ${cfg.nombre_regente || '___________________'}`);
  hTemp.getRange(4, 1).setValue(`Total tardanzas: ${data.total}   Recurrentes: ${data.recurrentes}`);

  // Encabezado tabla
  const enc = ['N°', 'Hora', 'Nombre completo', 'Nivel', 'Grado', 'Paralelo', 'Turno', 'Estado'];
  hTemp.getRange(6, 1, 1, enc.length).setValues([enc])
    .setFontWeight('bold')
    .setBackground('#1a3a5c')
    .setFontColor('#ffffff');

  // Datos
  if (data.lista.length) {
    const rows = data.lista.map(t => [
      t.nro, t.hora, t.nombre_completo, t.nivel, t.grado, t.paralelo, t.turno,
      t.recurrente ? 'RECURRENTE' : 'Primera vez'
    ]);
    hTemp.getRange(7, 1, rows.length, enc.length).setValues(rows);
  } else {
    hTemp.getRange(7, 1).setValue('Sin tardanzas registradas este día.');
  }

  // Observación
  const filaObs = 7 + data.lista.length + 1;
  hTemp.getRange(filaObs, 1).setValue(`Observación: ${data.observacion || '—'}`);
  hTemp.getRange(filaObs + 2, 1).setValue('Firma y sello: ___________________');
  hTemp.getRange(filaObs + 2, 4).setValue(`Fecha impresión: ${hoy}`);

  // Generar PDF
  const ssId  = ss.getId();
  const shId  = hTemp.getSheetId();
  const url   = `https://docs.google.com/spreadsheets/d/${ssId}/export?format=pdf&gid=${shId}&portrait=true&fitw=true&gridlines=false`;
  const token = ScriptApp.getOAuthToken();
  const blob  = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${token}` }
  }).getBlob().setName(`Reporte_Tardanzas_${hoy}.pdf`);

  // Guardar en Drive y obtener link
  const archivo = DriveApp.createFile(blob);
  archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const linkPDF = archivo.getDownloadUrl();

  // Limpiar hoja temporal
  ss.deleteSheet(hTemp);

  return { status: 'ok', url: linkPDF };
}

// ============================================================
// FUNCIÓN AUXILIAR — Exportar PDF desde doGet
// ============================================================
function doGetPDF(e) {
  if (e.parameter.action === 'exportarPDF') {
    const result = exportarPDFDiario(e.parameter.fecha);
    return respuestaJSON(result);
  }
}
