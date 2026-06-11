// ═══════════════════════════════════════════════════════════════════
// STATE.JS — Estado global, persistencia IndexedDB y sesión JSON
// ═══════════════════════════════════════════════════════════════════

const APP_VERSION = '1.0.0';

// ── Procesadoras habilitadas (se persiste en localStorage) ──────────
let PROCS_ENABLED = (function(){
  try {
    const saved = JSON.parse(localStorage.getItem('procs_enabled') || '{}');
    return { FISERV: true, GETPOS: true, GOCUOTAS: false, ...saved };
  } catch { return { FISERV: true, GETPOS: true, GOCUOTAS: false }; }
})();

function toggleProc(id) {
  PROCS_ENABLED[id] = !PROCS_ENABLED[id];
  localStorage.setItem('procs_enabled', JSON.stringify(PROCS_ENABLED));
}

// Auto-detección: habilitada si hay archivos cargados para esa procesadora.
// FILES, _GOC_PAGOS y _GOC_CELULAR se definen en conciliador.js (carga después).
function isProcEnabled(id) {
  if (id === 'FISERV') {
    return !!(typeof FILES !== 'undefined' && FILES.fis);
  }
  if (id === 'GETPOS') {
    return !!(typeof FILES !== 'undefined' && FILES.gp);
  }
  if (id === 'GOCUOTAS') {
    return !!(
      (typeof _GOC_PAGOS   !== 'undefined' && _GOC_PAGOS.length   > 0) ||
      (typeof _GOC_CELULAR !== 'undefined' && _GOC_CELULAR.length  > 0)
    );
  }
  return PROCS_ENABLED[id] !== false;
}
const DB_NAME     = 'ConciliadorDB';
const DB_VERSION  = 3;   // v3: added 'archivos' store (biblioteca de archivos)

// ── Estado de archivos cargados
const FILES = { sky: null, fis: null, gp: null, ter: null, enu: null, baj: null, liq: null };

// ── Resultados de conciliación
let RESULTADO   = [];   // filas conciliadas
let CORREGIDAS  = {};   // { idx: { cupon, proc, motivo, obs, usuario, ts } }
let LOG_AUDIT   = [];   // log de todas las correcciones

// ── Tablas maestras
let TM = {
  sucursales:  [],  // { id, nombre, estado }
  vendedores:  [],  // { id, nombre, sucursal, legajo }
  terminales:  [],  // { terminal, procesadora, sucursal, nroCom, vigDesde, vigHasta }
  comercios:   [],  // { nroCom, procesadora, acuerdo, vigDesde, vigHasta }
  tarjetas:    [],  // { tarjeta, equivSkylab, equivProc }
  planes:      [],  // { plan, cuotas, tarjeta, procesadora, codigos }
  tasas:       [],  // { acuerdo, procesadora, comercio, tarjeta, plan, cuotas, tasa, coef, vigDesde, vigHasta }
  motivos:     ['Error de digitación','Tarjeta incorrecta','Plan incorrecto','Cuotas incorrectas',
                'Terminal incorrecta','Procesadora incorrecta','Diferencia de monto','Sin correspondencia','Otro'],
  estados:     ['Pendiente','Corregido','Validado','Sin correspondencia','En revisión'],
  equivalencias: {  // (nroCom_sky|suc) → nroCom_fiserv
    '100501000|214': '380877387',
    '100501000|237': '380877387',
    '100501000|531': '380877387',
    '100501000|562': '380877387',
  },
};

// ── Sesión actual
let SESSION = {
  id:        null,
  nombre:    '',
  usuario:   'Usuario',
  creada:    null,
  modificada:null,
  periodoDesde: '',
  periodoHasta: '',
};

// ════════════════════════════════════════════════════════════════════
// INDEXEDDB
// ════════════════════════════════════════════════════════════════════
let _db = null;

async function dbOpen() {
  if (_db) return _db;
  return new Promise((res, rej) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = e => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains('sesiones'))
        db.createObjectStore('sesiones', { keyPath: 'id' });
      if (!db.objectStoreNames.contains('tablasMaestras'))
        db.createObjectStore('tablasMaestras', { keyPath: 'clave' });
      if (!db.objectStoreNames.contains('periodos'))
        db.createObjectStore('periodos', { keyPath: 'id' });
      if (!db.objectStoreNames.contains('archivos'))
        db.createObjectStore('archivos', { keyPath: 'id' });
    };
    req.onsuccess  = e => { _db = e.target.result; res(_db); };
    req.onerror    = e => rej(e.target.error);
  });
}

async function dbGet(store, key) {
  const db = await dbOpen();
  return new Promise((res, rej) => {
    const tx  = db.transaction(store, 'readonly');
    const req = tx.objectStore(store).get(key);
    req.onsuccess = e => res(e.target.result);
    req.onerror   = e => rej(e.target.error);
  });
}

async function dbPut(store, value) {
  const db = await dbOpen();
  return new Promise((res, rej) => {
    const tx  = db.transaction(store, 'readwrite');
    const req = tx.objectStore(store).put(value);
    req.onsuccess = () => res();
    req.onerror   = e => rej(e.target.error);
  });
}

async function dbGetAll(store) {
  const db = await dbOpen();
  return new Promise((res, rej) => {
    const tx  = db.transaction(store, 'readonly');
    const req = tx.objectStore(store).getAll();
    req.onsuccess = e => res(e.target.result);
    req.onerror   = e => rej(e.target.error);
  });
}

async function dbDelete(store, key) {
  const db = await dbOpen();
  return new Promise((res, rej) => {
    const tx  = db.transaction(store, 'readwrite');
    const req = tx.objectStore(store).delete(key);
    req.onsuccess = () => res();
    req.onerror   = e => rej(e.target.error);
  });
}

// ════════════════════════════════════════════════════════════════════
// AUTO-GUARDADO
// ════════════════════════════════════════════════════════════════════
let _autoSaveTimer = null;

// Hook para que biblioteca.js pueda guardar también en el slot del período activo
let _periodoActivoHook = null;

function scheduleAutoSave() {
  clearTimeout(_autoSaveTimer);
  _autoSaveTimer = setTimeout(autoSave, 3000); // 3s de debounce
}

async function autoSave() {
  if (!SESSION.id) SESSION.id = 'ses_' + Date.now();
  SESSION.modificada = new Date().toISOString();

  const snapshot = buildSnapshot();
  try {
    await dbPut('sesiones', { id: SESSION.id, ...snapshot });
    // Guardar también en el slot del período activo (si hay uno abierto)
    if (typeof _periodoActivoHook === 'function') {
      await _periodoActivoHook(snapshot).catch(e => console.warn('periodoHook error:', e));
    }
    updateSaveIndicator('saved');
  } catch(e) {
    console.warn('AutoSave error:', e);
    updateSaveIndicator('error');
  }
}

// ════════════════════════════════════════════════════════════════════
// CONCILIACIONES POR PERÍODO
// Guarda resultados + correcciones en un slot fijo por mes (YYYY-MM)
// Reutiliza el store 'sesiones' con id 'per_YYYY-MM'
// ════════════════════════════════════════════════════════════════════

async function guardarConciliacionPeriodo(mes) {
  if (!mes) return;
  const snap = buildSnapshot();
  await dbPut('sesiones', {
    id: `per_${mes}`,
    ...snap,
    tipo: 'periodo_conc',
    periodoMes: mes,
    ultimaMod: new Date().toISOString(),
  });
}

async function cargarConciliacionPeriodo(mes) {
  if (!mes) return false;
  const snap = await dbGet('sesiones', `per_${mes}`);
  if (!snap) return false;
  return restaurarSnapshot(snap);
}

async function listarConciliacionesPeriodo() {
  try {
    const todas = await dbGetAll('sesiones');
    return todas
      .filter(s => s.tipo === 'periodo_conc')
      .sort((a, b) => (b.periodoMes || '').localeCompare(a.periodoMes || ''));
  } catch { return []; }
}

async function eliminarConciliacionPeriodo(mes) {
  await dbDelete('sesiones', `per_${mes}`);
}

function updateSaveIndicator(state) {
  const el = document.getElementById('save-indicator');
  if (!el) return;
  if (state === 'saved') {
    el.textContent = '✓ Guardado';
    el.style.color = 'var(--grn)';
    setTimeout(() => { el.textContent = ''; }, 3000);
  } else if (state === 'saving') {
    el.textContent = '↻ Guardando...';
    el.style.color = 'var(--yel)';
  } else {
    el.textContent = '⚠ Error al guardar';
    el.style.color = 'var(--red)';
  }
}

// ════════════════════════════════════════════════════════════════════
// SNAPSHOT — serializar/deserializar estado completo
// ════════════════════════════════════════════════════════════════════
function buildSnapshot() {
  return {
    version:    APP_VERSION,
    session:    { ...SESSION },
    tm:         JSON.parse(JSON.stringify(TM)),
    resultado:  RESULTADO.map(serializarFila),
    corregidas:       { ...CORREGIDAS },
    comErradoMarcas:  { ...(typeof COM_ERRADO_MARCAS !== 'undefined' ? COM_ERRADO_MARCAS : {}) },
    logAudit:         [...LOG_AUDIT],
    ctrSeguimiento:   { ...(typeof CTR_SEGUIMIENTO !== 'undefined' ? CTR_SEGUIMIENTO : {}) },
  };
}

function serializarFila(r) {
  // El objeto proc puede ser grande — guardamos solo los campos necesarios
  const p = r.proc ? {
    lote:    r.proc.lote,    ticket:  r.proc.ticket,  aut:    r.proc.aut,
    cupon:   r.proc.cupon,                             // GETPOS: cupón para cruce COBROS
    monto:   r.proc.monto,   montoN:  r.proc.montoN,  fecha:  r.proc.fecha,
    suc:     r.proc.suc,     tarjeta: r.proc.tarjeta,  cuotas: r.proc.cuotas,
    comFis:  r.proc.comFis,  nombre:  r.proc.nombre,   marca:  r.proc.marca,
    plan:    r.proc.plan,    equipo:  r.proc.equipo,   pos:    r.proc.pos,
    tipo:    r.proc.tipo,    arancel: r.proc.arancel,  cfo:    r.proc.cfo,
  } : null;
  return {
    sky: { ...r.sky },
    proc: p,
    metodo: r.metodo, estado: r.estado,
    procEncontrada: r.procEncontrada, procEsperada: r.procEsperada,
    comOK: r.comOK, sucOK: r.sucOK, matchParcial: r.matchParcial,
    esDevolucion: r.esDevolucion, esAnulSinCobro: r.esAnulSinCobro,
    correccionManual: r.correccionManual,
    procMontoNorm: r.procMontoNorm,
    difCuotas: r.difCuotas, skyCuotas: r.skyCuotas, skyCuotasTM: r.skyCuotasTM, procCuotas: r.procCuotas,
    difProcesadora: r.difProcesadora,
    difTasa: r.difTasa, difMonto: r.difMonto,
    tasaCobrada: r.tasaCobrada, tasaAcordada: r.tasaAcordada,
    accionSugerida: r.accionSugerida,
  };
}

function restaurarSnapshot(snap) {
  if (!snap) return false;
  try {
    SESSION    = { ...snap.session };
    TM         = snap.tm ? { ...TM, ...snap.tm } : TM;
    RESULTADO  = snap.resultado || [];
    CORREGIDAS = snap.corregidas || {};
    if (typeof COM_ERRADO_MARCAS !== 'undefined') Object.assign(COM_ERRADO_MARCAS, snap.comErradoMarcas || {});
    LOG_AUDIT  = snap.logAudit   || [];
    if (typeof CTR_SEGUIMIENTO !== 'undefined') Object.assign(CTR_SEGUIMIENTO, snap.ctrSeguimiento || {});
    return true;
  } catch(e) {
    console.error('Error restaurando snapshot:', e);
    return false;
  }
}

// ════════════════════════════════════════════════════════════════════
// EXPORTAR / IMPORTAR SESIÓN JSON
// ════════════════════════════════════════════════════════════════════
function exportarSesion() {
  const snap = buildSnapshot();
  const json = JSON.stringify(snap, null, 2);
  const blob = new Blob([json], { type: 'application/json' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  const ts   = new Date().toISOString().slice(0,10);
  a.href     = url;
  a.download = `Sesion_Conciliacion_${ts}.json`;
  a.click();
  URL.revokeObjectURL(url);
}

function importarSesion(file) {
  return new Promise((res, rej) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const snap = JSON.parse(e.target.result);
        if (restaurarSnapshot(snap)) {
          autoSave();
          res(snap);
        } else {
          rej(new Error('Formato de sesión inválido'));
        }
      } catch(err) { rej(err); }
    };
    reader.onerror = () => rej(new Error('Error leyendo archivo'));
    reader.readAsText(file);
  });
}

// ════════════════════════════════════════════════════════════════════
// CARGAR SESIONES GUARDADAS
// ════════════════════════════════════════════════════════════════════
async function listarSesiones() {
  try {
    return await dbGetAll('sesiones');
  } catch { return []; }
}

async function cargarSesion(id) {
  const snap = await dbGet('sesiones', id);
  if (snap && restaurarSnapshot(snap)) return true;
  return false;
}

async function eliminarSesion(id) {
  await dbDelete('sesiones', id);
}

// ════════════════════════════════════════════════════════════════════
// TABLAS MAESTRAS — persistencia
// ════════════════════════════════════════════════════════════════════
async function guardarTM() {
  try {
    await dbPut('tablasMaestras', { clave: 'tm', data: JSON.parse(JSON.stringify(TM)) });
    return true;
  } catch { return false; }
}

async function cargarTM() {
  try {
    const rec = await dbGet('tablasMaestras', 'tm');
    if (rec?.data) {
      TM = { ...TM, ...rec.data };
      return true;
    }
    return false;
  } catch { return false; }
}

// ════════════════════════════════════════════════════════════════════
// LOG DE AUDITORÍA
// ════════════════════════════════════════════════════════════════════
function registrarCorreccion({ skyIdx, campo, valorAntes, valorDespues, motivo, obs }) {
  const fila = RESULTADO[skyIdx];
  const entry = {
    ts:           new Date().toISOString(),
    usuario:      SESSION.usuario || 'Usuario',
    asiento:      fila?.sky?.asiento ?? skyIdx,
    suc:          fila?.sky?.suc ?? '',
    tarjeta:      fila?.sky?.tarjeta ?? '',
    monto:        fila?.sky?.monto ?? '',
    campo,
    valorAntes,
    valorDespues,
    motivo:       motivo || '',
    obs:          obs    || '',
    estadoSistema:'Pendiente',
  };
  LOG_AUDIT.push(entry);
  scheduleAutoSave();
  return entry;
}

// ════════════════════════════════════════════════════════════════════
// PERÍODOS HISTÓRICOS — base acumulativa de cierres de período
// ════════════════════════════════════════════════════════════════════

async function guardarPeriodo(periodo) {
  await dbPut('periodos', periodo);
}

async function listarPeriodos() {
  try { return (await dbGetAll('periodos')).sort((a,b) => a.periodoDesde < b.periodoDesde ? 1 : -1); }
  catch { return []; }
}

async function eliminarPeriodo(id) {
  await dbDelete('periodos', id);
}

function exportarPeriodoJSON(periodo) {
  const json = JSON.stringify(periodo, null, 2);
  const blob = new Blob([json], { type: 'application/json' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = `Periodo_${periodo.periodoDesde}_${periodo.periodoHasta}.json`;
  a.click();
  URL.revokeObjectURL(url);
}

function importarPeriodoJSON(file) {
  return new Promise((res, rej) => {
    const reader = new FileReader();
    reader.onload = async e => {
      try {
        const p = JSON.parse(e.target.result);
        if (!p.id || !p.periodoDesde) { rej(new Error('Archivo de período inválido')); return; }
        await guardarPeriodo(p);
        res(p);
      } catch(err) { rej(err); }
    };
    reader.onerror = () => rej(new Error('Error leyendo archivo'));
    reader.readAsText(file);
  });
}

// Init: cargar TM al arrancar
cargarTM().then(ok => {
  if (ok) console.log('Tablas maestras restauradas desde IndexedDB');
});
