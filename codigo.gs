// ============================================================
// SISTEMA DE GESTIÓN DE TARDANZAS — U.E. Buena Fe
// Google Apps Script — codigo.gs  *** VERSIÓN CORREGIDA ***
//
// CAMBIO CLAVE: todas las escrituras también van por doGet
// porque POST desde navegador externo es bloqueado por CORS.
// El frontend envía TODO como GET con parámetros en la URL.
// ============================================================

const CONFIG = {
  USUARIO:           'regente',
  PASSWORD:          'buenafe2026',
  UMBRAL_RECURRENTE: 3,
  NOMBRE_UE:         'U.E. Buena Fe',
  ANIO:              '2026',
};

const HOJA_ALUMNOS   = 'Alumnos';
const HOJA_TARDANZAS = 'Tardanzas';
const HOJA_CONFIG    = 'Configuracion';

// ============================================================
// INICIALIZACIÓN — Ejecutar manualmente UNA VEZ
// ============================================================
function inicializarSistema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let hA = ss.getSheetByName(HOJA_ALUMNOS);
  if (!hA) {
    hA = ss.insertSheet(HOJA_ALUMNOS);
    const enc = ['carnet','nombre_completo','nivel','grado','paralelo','turno'];
    hA.appendRow(enc);
    hA.getRange(1,1,1,enc.length).setFontWeight('bold').setBackground('#1a3a5c').setFontColor('#ffffff');
    hA.setFrozenRows(1);
  }

  let hT = ss.getSheetByName(HOJA_TARDANZAS);
  if (!hT) {
    hT = ss.insertSheet(HOJA_TARDANZAS);
    const enc = ['id','fecha','hora','carnet','nombre_completo','nivel','grado','paralelo','turno','recurrente','observacion'];
    hT.appendRow(enc);
    hT.getRange(1,1,1,enc.length).setFontWeight('bold').setBackground('#1a3a5c').setFontColor('#ffffff');
    hT.setFrozenRows(1);
  }

  let hC = ss.getSheetByName(HOJA_CONFIG);
  if (!hC) {
    hC = ss.insertSheet(HOJA_CONFIG);
    hC.appendRow(['clave','valor']);
    hC.getRange(1,1,1,2).setFontWeight('bold').setBackground('#1a3a5c').setFontColor('#ffffff');
    hC.setFrozenRows(1);
    [['invierno','false'],['umbral_recurrente','3'],['nombre_ue',CONFIG.NOMBRE_UE],
     ['anio',CONFIG.ANIO],['nombre_regente','']].forEach(r=>hC.appendRow(r));
  }
  Logger.log('Sistema inicializado');
}

// ============================================================
// RESPUESTA JSON con headers CORS
// ============================================================
function respuestaJSON(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// doGet — Maneja TODAS las peticiones (lecturas Y escrituras)
// El frontend envía todo por GET para evitar el bloqueo CORS
// que Google impone sobre POST desde dominios externos.
// ============================================================
function doGet(e) {
  try {
    const p = e.parameter;
    const action = p.action || '';

    // ── LECTURAS ──────────────────────────────────────────
    if (action === 'getAlumnos')
      return respuestaJSON({ status:'ok', data: obtenerAlumnos(p) });

    if (action === 'getTardanzas')
      return respuestaJSON({ status:'ok', data: obtenerTardanzas(p) });

    if (action === 'getConfig')
      return respuestaJSON({ status:'ok', data: obtenerConfig() });

    if (action === 'getStats')
      return respuestaJSON({ status:'ok', data: obtenerStats(p.fecha) });

    if (action === 'reporteDiario')
      return respuestaJSON({ status:'ok', data: generarReporteDiario(p.fecha) });

    if (action === 'reporteMensual')
      return respuestaJSON({ status:'ok', data: generarReporteMensual(p.mes) });

    if (action === 'reporteRecurrentes')
      return respuestaJSON({ status:'ok', data: generarReporteRecurrentes(p.mes) });

    if (action === 'reporteCurso')
      return respuestaJSON({ status:'ok', data: generarReporteCurso(p.mes) });

    // ── ESCRITURAS (antes eran POST, ahora también GET) ───
    if (action === 'registrarTardanza')
      return respuestaJSON(registrarTardanza(p));

    if (action === 'eliminarTardanza')
      return respuestaJSON(eliminarTardanza(p.id));

    if (action === 'agregarAlumno')
      return respuestaJSON(agregarAlumnoManual(p));

    if (action === 'guardarConfig')
      return respuestaJSON(guardarConfig(p));

    if (action === 'guardarObservacion')
      return respuestaJSON(guardarObservacion(p));

    // ── IMPORTAR ALUMNOS (datos JSON en parámetro 'datos') ─
    if (action === 'importarAlumnos') {
      const datos = JSON.parse(decodeURIComponent(p.datos || '[]'));
      return respuestaJSON(importarAlumnos({
        alumnos: datos,
        nivel:    p.nivel   || '',
        grado:    p.grado   || '',
        paralelo: p.paralelo|| '',
        turno:    p.turno   || '',
      }));
    }

    return respuestaJSON({ status:'error', mensaje:'Acción no reconocida: '+action });

  } catch(err) {
    Logger.log('ERROR doGet: '+err.toString());
    return respuestaJSON({ status:'error', mensaje: err.toString() });
  }
}

// ============================================================
// doPost — Se mantiene por compatibilidad pero ya no es necesario
// ============================================================
function doPost(e) {
  try {
    const datos = JSON.parse(e.postData.contents);
    // Redirigir todo al mismo handler
    e.parameter = e.parameter || {};
    Object.assign(e.parameter, datos);
    return doGet(e);
  } catch(err) {
    return respuestaJSON({ status:'error', mensaje: err.toString() });
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
  const nivel   = datos.nivel    || '';
  const grado   = datos.grado    || '';
  const paralelo= datos.paralelo || '';
  const turno   = datos.turno    || '';

  if (!alumnos.length) return { status:'error', mensaje:'No se recibieron alumnos' };

  const mapaExistentes = {};
  filas.slice(1).forEach((f, idx) => {
    const carnet = String(f[0]).trim();
    if (carnet) mapaExistentes[carnet] = idx + 2;
  });

  let nuevos = 0, actualizados = 0;
  alumnos.forEach(a => {
    const carnet = String(a.carnet || '').trim();
    const nombre = String(a.nombre_completo || '').trim().toUpperCase();
    if (!carnet || !nombre) return;
    const fila = [carnet, nombre, nivel || a.nivel, grado || a.grado, paralelo || a.paralelo, turno || a.turno];
    if (mapaExistentes[carnet]) {
      hoja.getRange(mapaExistentes[carnet], 1, 1, 6).setValues([fila]);
      actualizados++;
    } else {
      hoja.appendRow(fila);
      nuevos++;
    }
  });

  return { status:'ok', mensaje:`${nuevos} nuevos, ${actualizados} actualizados`, nuevos, actualizados };
}

function agregarAlumnoManual(datos) {
  const hoja   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_ALUMNOS);
  const filas  = hoja.getDataRange().getValues();
  const carnet = String(datos.carnet || '').trim();
  if (!carnet) return { status:'error', mensaje:'El carnet es obligatorio' };
  const existe = filas.slice(1).some(f => String(f[0]).trim() === carnet);
  if (existe) return { status:'error', mensaje:'Ya existe un alumno con ese carnet' };
  hoja.appendRow([carnet, String(datos.nombre_completo||'').trim().toUpperCase(),
    datos.nivel||'', datos.grado||'', datos.paralelo||'', datos.turno||'']);
  return { status:'ok', mensaje:'Alumno agregado' };
}

// ============================================================
// TARDANZAS
// ============================================================
function registrarTardanza(datos) {
  const hTardanzas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_TARDANZAS);
  const carnet     = String(datos.carnet || '').trim();
  if (!carnet) return { status:'error', mensaje:'Carnet requerido' };

  const cfg    = obtenerConfig();
  const umbral = parseInt(cfg.umbral_recurrente) || CONFIG.UMBRAL_RECURRENTE;

  // Fecha y hora en Bolivia UTC-4
  const ahora  = new Date();
  const bolivia = new Date(ahora.getTime() - (4 * 60 * 60 * 1000));
  const fecha  = Utilities.formatDate(bolivia, 'GMT', 'yyyy-MM-dd');
  const hora   = Utilities.formatDate(bolivia, 'GMT', 'HH:mm');

  const mes        = fecha.slice(0, 7);
  const tardanzas  = hTardanzas.getDataRange().getValues();
  const enMes      = tardanzas.slice(1).filter(f => {
    const fCarnet = String(f[3]).trim();
    const fFecha  = f[1] instanceof Date ? Utilities.formatDate(f[1],'America/La_Paz','yyyy-MM-dd') : String(f[1]);
    return fCarnet === carnet && fFecha.startsWith(mes);
  }).length;

  const esRecurrente = (enMes + 1) >= umbral;
  const id           = 'T' + Date.now();

  hTardanzas.appendRow([
    id, fecha, hora, carnet,
    String(datos.nombre_completo || '').trim().toUpperCase(),
    datos.nivel || '', datos.grado || '', datos.paralelo || '',
    datos.turno || '', esRecurrente ? 'SI' : 'NO', datos.observacion || '',
  ]);

  return { status:'ok', id, fecha, hora, recurrente: esRecurrente, totalMes: enMes + 1 };
}

function obtenerTardanzas(params) {
  const hoja  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_TARDANZAS);
  const filas = hoja.getDataRange().getValues();
  if (filas.length <= 1) return [];

  let lista = filas.slice(1).map(f => ({
    id:              String(f[0]),
    fecha:           f[1] instanceof Date ? Utilities.formatDate(f[1],'America/La_Paz','yyyy-MM-dd') : String(f[1]),
    hora:            f[2] instanceof Date ? Utilities.formatDate(f[2],'America/La_Paz','HH:mm') : String(f[2]),
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

  return lista.reverse();
}

function eliminarTardanza(id) {
  const hoja  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_TARDANZAS);
  const filas = hoja.getDataRange().getValues();
  for (let i = 1; i < filas.length; i++) {
    if (String(filas[i][0]) === String(id)) {
      hoja.deleteRow(i + 1);
      return { status:'ok', mensaje:'Tardanza eliminada' };
    }
  }
  return { status:'error', mensaje:'No encontrada' };
}

// ============================================================
// REPORTES
// ============================================================
function generarReporteDiario(fecha) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_TARDANZAS);
  const filas = hoja.getDataRange().getValues();
  const hoy  = fecha || Utilities.formatDate(new Date(new Date().getTime()-4*3600000),'GMT','yyyy-MM-dd');
  const lista = filas.slice(1).filter(f => (f[1] instanceof Date ? Utilities.formatDate(f[1],'America/La_Paz','yyyy-MM-dd') : String(f[1])) === hoy && String(f[0]))
    .map((f,i) => ({ nro:i+1, hora: f[2] instanceof Date ? Utilities.formatDate(f[2],'America/La_Paz','HH:mm') : String(f[2]), nombre_completo:String(f[4]),
      nivel:String(f[5]), grado:String(f[6]), paralelo:String(f[7]),
      turno:String(f[8]), recurrente:String(f[9])==='SI' }));
  return { fecha:hoy, lista, total:lista.length,
    recurrentes:lista.filter(t=>t.recurrente).length,
    observacion:obtenerObservacionDia(hoy), config:obtenerConfig() };
}

function generarReporteMensual(mes) {
  const hoja  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_TARDANZAS);
  const filas = hoja.getDataRange().getValues();
  const m     = mes || Utilities.formatDate(new Date(new Date().getTime()-4*3600000),'GMT','yyyy-MM');
  const lista = filas.slice(1).filter(f => (f[1] instanceof Date ? Utilities.formatDate(f[1],'America/La_Paz','yyyy-MM-dd') : String(f[1])).startsWith(m) && String(f[0]))
    .map((f,i) => ({ nro:i+1, fecha: f[1] instanceof Date ? Utilities.formatDate(f[1],'America/La_Paz','yyyy-MM-dd') : String(f[1]), hora: f[2] instanceof Date ? Utilities.formatDate(f[2],'America/La_Paz','HH:mm') : String(f[2]),
      nombre_completo:String(f[4]), nivel:String(f[5]), grado:String(f[6]),
      paralelo:String(f[7]), turno:String(f[8]), recurrente:String(f[9])==='SI' }));
  return { mes:m, lista, total:lista.length };
}

function generarReporteRecurrentes(mes) {
  const hoja   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_TARDANZAS);
  const filas  = hoja.getDataRange().getValues();
  const m      = mes || Utilities.formatDate(new Date(new Date().getTime()-4*3600000),'GMT','yyyy-MM');
  const cfg    = obtenerConfig();
  const umbral = parseInt(cfg.umbral_recurrente) || 3;
  const conteo = {};
  filas.slice(1).filter(f => (f[1] instanceof Date ? Utilities.formatDate(f[1],'America/La_Paz','yyyy-MM-dd') : String(f[1])).startsWith(m) && String(f[0])).forEach(f => {
    const c = String(f[3]);
    if (!conteo[c]) conteo[c] = { carnet:c, nombre_completo:String(f[4]),
      nivel:String(f[5]), grado:String(f[6]), paralelo:String(f[7]), turno:String(f[8]), total:0 };
    conteo[c].total++;
  });
  const lista = Object.values(conteo).filter(a=>a.total>=umbral)
    .sort((a,b)=>b.total-a.total).map((a,i)=>({nro:i+1,...a}));
  return { mes:m, lista, total:lista.length, umbral };
}

function generarReporteCurso(mes) {
  const hoja  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_TARDANZAS);
  const filas = hoja.getDataRange().getValues();
  const m     = mes || Utilities.formatDate(new Date(new Date().getTime()-4*3600000),'GMT','yyyy-MM');
  const grupos = {};
  filas.slice(1).filter(f => (f[1] instanceof Date ? Utilities.formatDate(f[1],'America/La_Paz','yyyy-MM-dd') : String(f[1])).startsWith(m) && String(f[0])).forEach(f => {
    const key = `${f[5]} ${f[6]} ${f[7]}`;
    grupos[key] = (grupos[key]||0)+1;
  });
  const lista = Object.entries(grupos).sort((a,b)=>b[1]-a[1])
    .map(([curso,total],i)=>({nro:i+1,curso,total}));
  return { mes:m, lista };
}

// ============================================================
// STATS DASHBOARD
// ============================================================
function obtenerStats(fechaParam) {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const hAlumnos   = ss.getSheetByName(HOJA_ALUMNOS);
  const hTardanzas = ss.getSheetByName(HOJA_TARDANZAS);
  const cfg        = obtenerConfig();
  const umbral     = parseInt(cfg.umbral_recurrente) || 3;

  // Usar fecha del cliente si viene, sino calcular en servidor UTC-4
  const hoy = fechaParam || Utilities.formatDate(new Date(new Date().getTime()-4*3600000),'GMT','yyyy-MM-dd');
  const mes  = hoy.slice(0,7);

  const filasA = hAlumnos.getDataRange().getValues();
  const filasT = hTardanzas.getDataRange().getValues();

  const _fmt = (f1) => f1 instanceof Date ? Utilities.formatDate(f1,'America/La_Paz','yyyy-MM-dd') : String(f1);
  const tardHoy = filasT.slice(1).filter(f => _fmt(f[1])===hoy && String(f[0]));
  const tardMes = filasT.slice(1).filter(f => _fmt(f[1]).startsWith(mes) && String(f[0]));

  const conteo = {};
  tardMes.forEach(f => { const c=String(f[3]); conteo[c]=(conteo[c]||0)+1; });
  const recurrentes = Object.values(conteo).filter(v=>v>=umbral).length;

  const ultimas = tardHoy.slice(-6).reverse().map(f => ({
    hora: f[2] instanceof Date ? Utilities.formatDate(f[2],'America/La_Paz','HH:mm') : String(f[2]),
    nombre_completo:String(f[4]),
    nivel:String(f[5]), grado:String(f[6]), paralelo:String(f[7]),
    turno:String(f[8]), recurrente:String(f[9])==='SI',
  }));

  const topRec = Object.entries(conteo).filter(([,v])=>v>=umbral)
    .sort((a,b)=>b[1]-a[1]).slice(0,5).map(([carnet,total]) => {
      const fila = filasT.slice(1).find(f=>String(f[3])===carnet);
      return { carnet, nombre_completo:fila?String(fila[4]):carnet,
        nivel:fila?String(fila[5]):'', grado:fila?String(fila[6]):'',
        paralelo:fila?String(fila[7]):'', total };
    });

  return { totalAlumnos:filasA.length-1, tardanzasHoy:tardHoy.length,
    tardanzasMes:tardMes.length, recurrentes, ultimas, topRecurrentes:topRec, fecha:hoy, mes };
}

// ============================================================
// CONFIGURACIÓN
// ============================================================
function obtenerConfig() {
  const hoja  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_CONFIG);
  const filas = hoja.getDataRange().getValues();
  const cfg   = {};
  filas.slice(1).forEach(f => { if(f[0]) cfg[String(f[0]).trim()]=String(f[1]).trim(); });
  return cfg;
}

function guardarConfig(datos) {
  const hoja   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_CONFIG);
  const filas  = hoja.getDataRange().getValues();
  const campos = ['invierno','umbral_recurrente','nombre_ue','anio','nombre_regente'];
  campos.forEach(campo => {
    if (datos[campo]===undefined) return;
    let found = false;
    for (let i=1; i<filas.length; i++) {
      if (String(filas[i][0]).trim()===campo) {
        hoja.getRange(i+1,2).setValue(String(datos[campo]));
        found=true; break;
      }
    }
    if (!found) hoja.appendRow([campo, String(datos[campo])]);
  });
  return { status:'ok' };
}

function guardarObservacion(datos) {
  const hoja  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_CONFIG);
  const filas = hoja.getDataRange().getValues();
  const hoy   = Utilities.formatDate(new Date(new Date().getTime()-4*3600000),'GMT','yyyy-MM-dd');
  const clave = 'obs_'+(datos.fecha||hoy);
  const texto = String(datos.observacion||'').trim();
  let found   = false;
  for (let i=1; i<filas.length; i++) {
    if (String(filas[i][0]).trim()===clave) {
      hoja.getRange(i+1,2).setValue(texto); found=true; break;
    }
  }
  if (!found) hoja.appendRow([clave, texto]);
  return { status:'ok' };
}

function obtenerObservacionDia(fecha) {
  const cfg = obtenerConfig();
  return cfg['obs_'+fecha] || '';
}

// ============================================================
// EXPORTAR PDF DEL REPORTE DIARIO
// Genera un PDF desde una hoja temporal y devuelve el link
// ============================================================
function exportarPDFDiario(fecha) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const cfg  = obtenerConfig();
  const hoy  = fecha || Utilities.formatDate(new Date(new Date().getTime()-4*3600000),'GMT','yyyy-MM-dd');
  const data = generarReporteDiario(hoy);

  // Crear hoja temporal de impresión
  let hTemp = ss.getSheetByName('_PDF_TEMP');
  if (hTemp) ss.deleteSheet(hTemp);
  hTemp = ss.insertSheet('_PDF_TEMP');

  // Título
  hTemp.getRange(1,1).setValue(cfg.nombre_ue || CONFIG.NOMBRE_UE);
  hTemp.getRange(1,1).setFontSize(14).setFontWeight('bold');
  hTemp.getRange(2,1).setValue('Reporte de Tardanzas — ' + hoy);
  hTemp.getRange(3,1).setValue('Regente: ' + (cfg.nombre_regente || '___________________'));
  hTemp.getRange(4,1).setValue('Total tardanzas: ' + data.total + '   Recurrentes: ' + data.recurrentes);

  // Encabezado tabla
  const enc = ['N°','Hora','Nombre completo','Nivel','Grado','Paralelo','Turno','Estado'];
  hTemp.getRange(6,1,1,enc.length).setValues([enc])
    .setFontWeight('bold').setBackground('#1a3a5c').setFontColor('#ffffff');

  // Datos
  if (data.lista.length) {
    const rows = data.lista.map(t => [
      t.nro, t.hora, t.nombre_completo, t.nivel, t.grado, t.paralelo, t.turno,
      t.recurrente ? 'RECURRENTE' : 'Primera vez'
    ]);
    hTemp.getRange(7,1,rows.length,enc.length).setValues(rows);
  } else {
    hTemp.getRange(7,1).setValue('Sin tardanzas registradas este día.');
  }

  // Observación y firma
  const filaObs = 7 + data.lista.length + 1;
  hTemp.getRange(filaObs,1).setValue('Observación: ' + (data.observacion || '—'));
  hTemp.getRange(filaObs+2,1).setValue('Firma y sello: ___________________');
  hTemp.getRange(filaObs+2,4).setValue('Fecha impresión: ' + hoy);

  // Generar PDF
  const ssId  = ss.getId();
  const shId  = hTemp.getSheetId();
  const url   = 'https://docs.google.com/spreadsheets/d/' + ssId +
                '/export?format=pdf&gid=' + shId + '&portrait=true&fitw=true&gridlines=false';
  const token = ScriptApp.getOAuthToken();
  const blob  = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token }
  }).getBlob().setName('Reporte_Tardanzas_' + hoy + '.pdf');

  // Guardar en Drive y obtener link
  const archivo = DriveApp.createFile(blob);
  archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const linkPDF = archivo.getDownloadUrl();

  // Limpiar hoja temporal
  ss.deleteSheet(hTemp);

  return { status:'ok', url: linkPDF };
}

// ============================================================
// EXPORTAR PDF desde doGet
// Llamar con: ?action=exportarPDF&fecha=2026-05-19
// ============================================================
function doGetPDF(e) {
  if (e.parameter.action === 'exportarPDF') {
    const result = exportarPDFDiario(e.parameter.fecha);
    return respuestaJSON(result);
  }
}
