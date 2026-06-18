// ═══════════════════════════════════════════════════════════════════
// STATE.JS — Estado global, persistencia IndexedDB y sesión JSON
// ═══════════════════════════════════════════════════════════════════

const APP_VERSION = '1.0.0';

// ── Procesadoras habilitadas (se persiste en localStorage) ──────────
let PROCS_ENABLED = (function(){
  try {
    const saved = JSON.parse(localStorage.getItem('procs_enabled') || '{}');
    return { FISERV: true, GETPOS: true, GOCUOTAS: false, PRISMA: true, ...saved };
  } catch { return { FISERV: true, GETPOS: true, GOCUOTAS: false, PRISMA: true }; }
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
const FILES = { sky: null, fis: null, gp: null, ter: null, enu: null, baj: null, liq: null, pri: null };

// ── Resultados de conciliación
let RESULTADO   = [];   // filas conciliadas
let CORREGIDAS  = {};   // { idx: { cupon, proc, motivo, obs, usuario, ts } }
let LOG_AUDIT   = [];   // log de todas las correcciones

// ── Período / Lote activos (sistema multi-lote) ─────────────────────
let _PERIODO_ACTIVO_ID = null;   // 'per_...' — período de conciliación activo
let _LOTE_ACTIVO_ID    = null;   // 'lot_...' — lote activo (para guardar al conciliar)

// ── Arrastre del período anterior (auto-cargado desde IDB al startup) ─
let _arrastreGuardado = null;    // { pendientes[], correcciones[], periodoHasta, nombre, ts }

// ── Tablas maestras
let TM = {
  sucursales:  [],  // { id, nombre, estado }
  vendedores:  [],  // { id, nombre, sucursal, legajo }
  terminales:  [],  // { terminal, procesadora, sucursal, nroCom, vigDesde, vigHasta }
  comercios:   [],  // { nroCom, procesadora, acuerdo, vigDesde, vigHasta }
  tarjetas:    [],  // { tarjeta, equivSkylab, equivProc }
  planes:      [],  // { plan, cuotas, tarjeta, procesadora, codigos }
  tasas:       [],  // { acuerdo, procesadora, comercio, tarjeta, plan, cuotas, tasa, coef, vigDesde, vigHasta }
  plazos:      [],  // { procesadora, comercio, tarjeta, dias_habiles, vigDesde, vigHasta }
  feriados:    [    // { fecha YYYY-MM-DD, descripcion }
    // ── 2025 ──────────────────────────────────────────────────────
    {fecha:'2025-01-01',descripcion:'Año Nuevo'},
    {fecha:'2025-03-03',descripcion:'Carnaval'},
    {fecha:'2025-03-04',descripcion:'Carnaval'},
    {fecha:'2025-03-24',descripcion:'Día Nac. de la Memoria'},
    {fecha:'2025-04-02',descripcion:'Día del Veterano (Malvinas)'},
    {fecha:'2025-04-17',descripcion:'Jueves Santo'},
    {fecha:'2025-04-18',descripcion:'Viernes Santo'},
    {fecha:'2025-05-01',descripcion:'Día del Trabajador'},
    {fecha:'2025-05-25',descripcion:'Día de la Patria'},
    {fecha:'2025-06-16',descripcion:'Gral. Belgrano (traslado)'},
    {fecha:'2025-07-09',descripcion:'Día de la Independencia'},
    {fecha:'2025-08-15',descripcion:'Gral. San Martín (traslado)'},
    {fecha:'2025-10-13',descripcion:'Diversidad Cultural (traslado)'},
    {fecha:'2025-11-24',descripcion:'Soberanía Nacional (traslado)'},
    {fecha:'2025-12-08',descripcion:'Inmaculada Concepción'},
    {fecha:'2025-12-25',descripcion:'Navidad'},
    // ── 2026 ──────────────────────────────────────────────────────
    {fecha:'2026-01-01',descripcion:'Año Nuevo'},
    {fecha:'2026-02-16',descripcion:'Carnaval'},
    {fecha:'2026-02-17',descripcion:'Carnaval'},
    {fecha:'2026-03-24',descripcion:'Día Nac. de la Memoria'},
    {fecha:'2026-04-02',descripcion:'Día del Veterano (Malvinas) / Jueves Santo'},
    {fecha:'2026-04-03',descripcion:'Viernes Santo'},
    {fecha:'2026-05-01',descripcion:'Día del Trabajador'},
    {fecha:'2026-05-25',descripcion:'Día de la Patria'},
    {fecha:'2026-06-15',descripcion:'Gral. Belgrano (traslado)'},
    {fecha:'2026-07-09',descripcion:'Día de la Independencia'},
    {fecha:'2026-08-17',descripcion:'Gral. San Martín'},
    {fecha:'2026-10-12',descripcion:'Diversidad Cultural'},
    {fecha:'2026-11-23',descripcion:'Soberanía Nacional (traslado)'},
    {fecha:'2026-12-08',descripcion:'Inmaculada Concepción'},
    {fecha:'2026-12-25',descripcion:'Navidad'},
  ],
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
  // _loteId y _loteFechas se preservan si existen (modo multi-lote)
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
    // multi-lote: preservar tag si viene de cargarPeriodoCompleto
    _loteId:     r._loteId     || undefined,
    _loteFechas: r._loteFechas || undefined,
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

// ════════════════════════════════════════════════════════════════════
// PERÍODOS DE CONCILIACIÓN CON LOTES (sistema multi-lote)
// Los períodos históricos (cerrarPeriodo) siguen usando listarPeriodos()
// Estos períodos tienen tipo:'conciliacion' en el store 'periodos'
// ════════════════════════════════════════════════════════════════════

// Crea un nuevo período de conciliación
async function crearPeriodoConciliacion(nombre) {
  const id  = 'per_' + Date.now();
  const per = {
    id,
    tipo:      'conciliacion',
    nombre,
    lotes:     [],
    creadoEn:  new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };
  await dbPut('periodos', per);
  return per;
}

// Lista todos los períodos de conciliación, más recientes primero
async function listarPeriodosConciliacion() {
  try {
    const todos = await dbGetAll('periodos');
    return todos
      .filter(p => p.tipo === 'conciliacion')
      .sort((a, b) => (b.creadoEn || '').localeCompare(a.creadoEn || ''));
  } catch { return []; }
}

// Obtiene un período de conciliación por ID
async function obtenerPeriodoConciliacion(id) {
  return dbGet('periodos', id);
}

// Agrega un lote a un período
async function agregarLotePeriodo(periodoId, { fechaDesde, fechaHasta }) {
  const per = await obtenerPeriodoConciliacion(periodoId);
  if (!per) throw new Error('Período no encontrado: ' + periodoId);
  const loteId = 'lot_' + Date.now();
  const lote = {
    id:           loteId,
    fechaDesde,
    fechaHasta,
    estado:       'pendiente',
    nOps:         0,
    creadoEn:     new Date().toISOString(),
    conciliadoEn: null,
  };
  per.lotes = per.lotes || [];
  per.lotes.push(lote);
  per.updatedAt = new Date().toISOString();
  await dbPut('periodos', per);
  return lote;
}

// Actualiza campos de un lote
async function actualizarLotePeriodo(periodoId, loteId, cambios) {
  const per = await obtenerPeriodoConciliacion(periodoId);
  if (!per) return;
  const idx = (per.lotes || []).findIndex(l => l.id === loteId);
  if (idx < 0) return;
  per.lotes[idx] = { ...per.lotes[idx], ...cambios };
  per.updatedAt  = new Date().toISOString();
  await dbPut('periodos', per);
}

// Elimina un lote y sus archivos + resultado
async function eliminarLotePeriodo(periodoId, loteId) {
  const per = await obtenerPeriodoConciliacion(periodoId);
  if (per) {
    per.lotes     = (per.lotes || []).filter(l => l.id !== loteId);
    per.updatedAt = new Date().toISOString();
    await dbPut('periodos', per);
  }
  try {
    const archivos = await dbGetAll('archivos');
    for (const a of archivos.filter(a => a.loteId === loteId)) {
      await dbDelete('archivos', a.id);
    }
  } catch(e) { console.warn('Error borrando archivos del lote:', e); }
  await dbDelete('sesiones', 'res_' + loteId).catch(() => {});
}

// Elimina un período con todos sus lotes, archivos y resultados
async function eliminarPeriodoConciliacion(periodoId) {
  const per = await obtenerPeriodoConciliacion(periodoId);
  if (per) {
    for (const lote of (per.lotes || [])) {
      await eliminarLotePeriodo(periodoId, lote.id);
    }
  }
  await dbDelete('periodos', periodoId);
}

// ── Resultado por lote ───────────────────────────────────────────────

// Guarda el RESULTADO actual en el slot del lote activo
async function guardarResultadoLote(periodoId, loteId, resultado) {
  const serialized = resultado.map(r => serializarFila({ ...r, _loteId: loteId }));
  await dbPut('sesiones', {
    id:           'res_' + loteId,
    tipo:         'lote_resultado',
    loteId,
    periodoId,
    resultado:    serialized,
    conciliadoEn: new Date().toISOString(),
  });
  await actualizarLotePeriodo(periodoId, loteId, {
    estado:       'conciliado',
    nOps:         resultado.length,
    conciliadoEn: new Date().toISOString(),
  });
}

// Carga el resultado de un lote
async function cargarResultadoLote(loteId) {
  return dbGet('sesiones', 'res_' + loteId);
}

// Merge de todos los lotes conciliados de un período → asigna RESULTADO global
async function cargarPeriodoCompleto(periodoId) {
  const per = await obtenerPeriodoConciliacion(periodoId);
  if (!per) return false;
  const lotes = (per.lotes || []).filter(l => l.estado === 'conciliado');
  if (!lotes.length) return false;
  let merged = [];
  for (const lote of lotes) {
    const res = await cargarResultadoLote(lote.id);
    if (res?.resultado?.length) {
      const tagged = res.resultado.map(r => ({
        ...r,
        _loteId:     lote.id,
        _loteFechas: `${lote.fechaDesde} – ${lote.fechaHasta}`,
      }));
      merged = merged.concat(tagged);
    }
  }
  RESULTADO = merged;
  return merged.length > 0;
}

// Init: cargar TM y arrastre al arrancar
cargarTM().then(ok => {
  if (ok) console.log('Tablas maestras restauradas desde IndexedDB');
});

dbGet('sesiones', 'arrastre_activo').then(rec => {
  if (rec?.pendientes?.length || rec?.correcciones?.length) {
    _arrastreGuardado = rec;
    console.log(`[Arrastre] ${rec.pendientes?.length||0} pendientes · ${rec.correcciones?.length||0} correcciones del período ${rec.periodoHasta}`);
  }
}).catch(() => {});
