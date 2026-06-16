// ═══════════════════════════════════════════════════════════════════
// BIBLIOTECA.JS — Gestión de archivos por período y lote
// Sistema multi-lote: cada período tiene N lotes con su propio rango
// de fechas, archivos y resultado de conciliación independiente.
// ═══════════════════════════════════════════════════════════════════

const _BIB_TIPOS = [
  // ── OPERACIONES: cruce diario cobrado vs facturado en Skylab ──────
  { id:'SKYLAB',      label:'Skylab (principal)',               modulo:'OPERACIONES', accept:'.xlsx,.xls', badge:'SK',  bg:'#0f2544', color:'#38bdf8' },
  { id:'TERMINALES',  label:'Terminales',                       modulo:'OPERACIONES', accept:'.xlsx,.xls', badge:'TER', bg:'#1a1f2e', color:'#8ba3c4' },
  { id:'FISERV_LIQ',  label:'FISERV Operaciones',               modulo:'OPERACIONES', accept:'.xlsx,.xls', badge:'FIS', bg:'#0f1e44', color:'#4f8ef7' },
  { id:'FISERV_CTR',  label:'FISERV Contracargos',              modulo:'OPERACIONES', accept:'.xlsx,.xls', badge:'FC',  bg:'#0f1e44', color:'#f87171' },
  { id:'GETPOS',      label:'GETPOS Operaciones',               modulo:'OPERACIONES', accept:'.xlsx,.xls', badge:'GP',  bg:'#0f2d20', color:'#34d399' },
  { id:'GETPOS_CTR',  label:'GETPOS Contracargos',              modulo:'OPERACIONES', accept:'.xlsx,.xls', badge:'GC',  bg:'#0f2d20', color:'#f87171' },
  { id:'GOC_PAGOS',   label:'Go Cuotas CSV',                    modulo:'OPERACIONES', accept:'.csv,.txt',  badge:'GQ',  bg:'#0f2d20', color:'#4ade80' },
  { id:'GOC_CELULAR', label:'Go Celular CSV',                   modulo:'OPERACIONES', accept:'.csv,.txt',  badge:'GK',  bg:'#1a1040', color:'#a78bfa' },
  { id:'GOC_VENTAS',  label:'GoC Ventas (IMEI)',                modulo:'OPERACIONES', accept:'.xlsx,.xls', badge:'GV',  bg:'#1a1f10', color:'#fbbf24' },
  // ── LIQUIDACIONES: cruce ops confirmadas vs pagos de procesadora ─
  { id:'LIQUIDACION',     label:'Cupones Liquidados (FISERV+GETPOS)', modulo:'LIQUIDACIONES', accept:'.xlsx,.xls', badge:'LIQ', bg:'#291000', color:'#fb923c' },
  { id:'GOC_LIQ_PAGOS',  label:'Go Cuotas Liquidación CSV',           modulo:'LIQUIDACIONES', accept:'.csv,.txt',  badge:'GQL', bg:'#0a2018', color:'#4ade80' },
  { id:'GOC_LIQ_CELULAR',label:'Go Celular Liquidación CSV',           modulo:'LIQUIDACIONES', accept:'.csv,.txt',  badge:'GCL', bg:'#130a30', color:'#a78bfa' },
];

// ── Loader: convierte un registro de la biblioteca en función de carga ─
function _bibMakeLoader(rec) {
  const blob = new Blob([rec.bytes]);
  const file = new File([blob], rec.nombre, { type: 'application/octet-stream' });
  const fi   = { files: [file] };
  switch (rec.tipo) {
    case 'SKYLAB':      return () => loadFile(fi, 'sky');
    case 'TERMINALES':  return () => loadFile(fi, 'ter');
    case 'FISERV_LIQ':  return () => loadFile(fi, 'fis');
    case 'FISERV_CTR':  return () => loadContracargos(fi, 'fis');
    case 'GETPOS':      return () => loadFile(fi, 'gp');
    case 'GETPOS_CTR':  return () => loadContracargos(fi, 'gp');
    case 'GOC_PAGOS':      return () => loadGocPagos(fi, 'GOCUOTAS');
    case 'GOC_CELULAR':    return () => loadGocPagos(fi, 'GOCELULAR');
    case 'GOC_VENTAS':     return () => loadGocVentas(fi);
    case 'LIQUIDACION':    return () => { loadLiquidaciones(fi); liqCargarCupones(fi); };
    case 'GOC_LIQ_PAGOS':  return () => loadGocPagosLiq(fi, 'GOCUOTAS');
    case 'GOC_LIQ_CELULAR':return () => loadGocPagosLiq(fi, 'GOCELULAR');
    default:               return null;
  }
}

// ══ CRUD archivos ════════════════════════════════════════════════════

async function guardarArchivoBiblioteca({ tipo, periodoId, loteId, nombre, bytes }) {
  const id = `bib_${tipo}_${Date.now()}`;
  const kb = Math.round(bytes.byteLength / 1024);
  await dbPut('archivos', {
    id, tipo,
    periodoId: periodoId || null,
    loteId:    loteId    || null,
    nombre, bytes, kb,
    fechaCarga: new Date().toISOString(),
  });
  return id;
}

async function listarArchivosBiblioteca() {
  try {
    return (await dbGetAll('archivos'))
      .sort((a, b) => (b.fechaCarga || '').localeCompare(a.fechaCarga || ''));
  } catch { return []; }
}

async function listarArchivosDeLote(loteId) {
  const todos = await listarArchivosBiblioteca();
  return todos.filter(a => a.loteId === loteId);
}

async function eliminarArchivoBiblioteca(id) {
  await dbDelete('archivos', id);
}

// ══ HELPERS UI ═══════════════════════════════════════════════════════

function _bibTipoInfo(id) {
  return _BIB_TIPOS.find(t => t.id === id) || { badge:'?', bg:'#333', color:'#999', label: id };
}

function _bibAlert(msg) {
  const el = document.getElementById('bib-alert');
  if (el) {
    el.textContent = '⚠ ' + msg;
    el.style.display = 'block';
    clearTimeout(el._t);
    el._t = setTimeout(() => { el.style.display = 'none'; }, 5000);
  }
}

// DD/MM/YYYY desde YYYY-MM-DD
function _bibFmtDate(d) {
  if (!d) return '—';
  const [y, m, day] = d.split('-');
  return `${day}/${m}/${y}`;
}

// ══ ABRIR / CERRAR ════════════════════════════════════════════════════

function abrirBiblioteca() {
  const modal = document.getElementById('modal-biblioteca');
  if (!modal) return;
  modal.classList.add('open');
  _renderBiblioteca();
}

function cerrarBiblioteca() {
  document.getElementById('modal-biblioteca')?.classList.remove('open');
}

// ══ RENDER PRINCIPAL ═════════════════════════════════════════════════

async function _renderBiblioteca() {
  const contenedor = document.getElementById('bib-content');
  if (!contenedor) return;

  const periodos      = await listarPeriodosConciliacion();
  const todosArchivos = await listarArchivosBiblioteca();

  await _actualizarBadgeBiblioteca(todosArchivos);

  if (!periodos.length) {
    contenedor.innerHTML = `
      <div style="text-align:center;padding:60px 20px;color:var(--m2)">
        <div style="font-size:36px;opacity:.15">📅</div>
        <div style="font-size:11px;margin-top:14px;font-weight:700;color:var(--txt);opacity:.4">Sin períodos creados</div>
        <div style="font-size:9px;margin-top:6px;line-height:1.8">
          Hacé clic en <b>＋ Nuevo período</b> para crear tu primer período de conciliación.<br>
          Dentro de cada período podés agregar múltiples lotes con sus rangos de fechas.
        </div>
      </div>`;
    return;
  }

  let html = '';
  for (const per of periodos) {
    const lotes     = per.lotes || [];
    const totalOps  = lotes.reduce((s, l) => s + (l.nOps || 0), 0);
    const lotesOk   = lotes.filter(l => l.estado === 'conciliado').length;
    const isActive  = per.id === _PERIODO_ACTIVO_ID;
    const dotCls    = isActive ? 'active-dot' : (lotesOk > 0 ? 'concil' : '');
    const dot       = isActive ? '●' : (lotesOk > 0 ? '✓' : '○');
    const estadoLbl = isActive ? `<span style="color:var(--grn);font-size:8px;font-weight:700">● ACTIVO</span>` : '';

    html += `
    <div class="bib-per-card${isActive ? ' active' : ''}" id="bibcard-${per.id}">
      <div class="bib-per-card-hdr" onclick="_bibTogglePer('${per.id}')">
        <div style="display:flex;align-items:center;gap:10px">
          <span class="sb-per-dot ${dotCls}" style="font-size:12px">${dot}</span>
          <div>
            <div class="bib-per-nombre">${per.nombre} ${estadoLbl}</div>
            <div class="bib-per-meta">
              ${lotes.length} lote${lotes.length !== 1 ? 's' : ''}
              ${lotesOk > 0 ? ` · <span style="color:var(--grn)">${totalOps.toLocaleString('es-AR')} ops</span>` : ' · sin conciliar'}
            </div>
          </div>
        </div>
        <div class="bib-per-card-acts" onclick="event.stopPropagation()">
          ${lotesOk > 0 ? `<button class="bib-act-btn bib-act-view" onclick="bibVerPeriodo('${per.id}')" title="Ver todas las ops consolidadas">📊 Ver todo</button>` : ''}
          ${lotesOk > 1 ? `<button class="bib-act-btn" onclick="bibReprocesarTodo('${per.id}')" title="Re-conciliar todos los lotes">⟳ Reprocesar</button>` : ''}
          <button class="bib-act-btn bib-act-del" onclick="bibEliminarPeriodo('${per.id}')" title="Eliminar período">✕</button>
        </div>
      </div>

      <!-- Lista de lotes (expandible) -->
      <div class="bib-lotes-list" id="bib-lotes-${per.id}">
        ${lotes.length === 0
          ? `<div style="padding:12px 18px;font-size:9px;color:var(--m2);font-style:italic">Sin lotes. Agregá el primero.</div>`
          : lotes.map(lote => _renderLoteRow(per.id, lote, todosArchivos)).join('')
        }
        <div style="padding:8px 18px 12px">
          <button class="bib-add-lote-btn" onclick="bibAbrirFormLote('${per.id}')">＋ Agregar lote</button>
        </div>
      </div>
    </div>`;
  }

  contenedor.innerHTML = html;
}

function _bibTogglePer(periodoId) {
  const el = document.getElementById('bib-lotes-' + periodoId);
  if (!el) return;
  el.style.display = el.style.display === 'none' ? '' : 'none';
}

function _renderLoteRow(periodoId, lote, todosArchivos) {
  const archivosLote  = todosArchivos.filter(a => a.loteId === lote.id);
  const tiposPres     = [...new Set(archivosLote.map(a => a.tipo))];
  const isActiveLote  = lote.id === _LOTE_ACTIVO_ID;

  const estadoBadge = lote.estado === 'conciliado'
    ? `<span class="bib-lote-badge ok">✅ ${(lote.nOps||0).toLocaleString('es-AR')} ops</span>`
    : `<span class="bib-lote-badge pend">🟡 Pendiente</span>`;

  const archBadges = tiposPres.map(t => {
    const ti = _bibTipoInfo(t);
    return `<span style="display:inline-flex;align-items:center;font-size:7px;font-weight:700;
      color:${ti.color};background:${ti.bg};border:1px solid ${ti.color}33;
      border-radius:3px;padding:1px 5px;margin-right:3px">${ti.badge}</span>`;
  }).join('');

  return `
  <div class="bib-lote-row${isActiveLote ? ' active' : ''}" id="bibrow-${lote.id}">
    <div class="bib-lote-left">
      <div class="bib-lote-fechas">
        <span style="font-size:9px;color:var(--m2)">📅</span>
        <b>${_bibFmtDate(lote.fechaDesde)}</b>
        <span style="color:var(--m2)">→</span>
        <b>${_bibFmtDate(lote.fechaHasta)}</b>
      </div>
      <div style="display:flex;align-items:center;gap:6px;flex-wrap:wrap;margin-top:4px">
        ${estadoBadge}
        ${archBadges || `<span style="font-size:8px;color:var(--m2)">sin archivos</span>`}
      </div>
    </div>
    <div class="bib-lote-acts">
      <button class="bib-act-btn bib-act-load" onclick="bibCargarLote('${periodoId}','${lote.id}')"
        title="Cargar archivos de este lote y conciliar">▶ Cargar</button>
      <button class="bib-act-btn" onclick="bibAbrirUploadLote('${periodoId}','${lote.id}')"
        title="Agregar archivos a este lote">＋ Arch.</button>
      <button class="bib-act-btn bib-act-del" onclick="bibEliminarLote('${periodoId}','${lote.id}')"
        title="Eliminar lote">✕</button>
    </div>
  </div>`;
}

// ══ FORMULARIOS ══════════════════════════════════════════════════════

function _bibShowForm(html) {
  const panel = document.getElementById('bib-form-panel');
  if (!panel) return;
  panel.innerHTML = html;
  panel.style.display = 'block';
}

function bibCerrarForm() {
  const panel = document.getElementById('bib-form-panel');
  if (panel) { panel.innerHTML = ''; panel.style.display = 'none'; }
}

// ── Nuevo período ────────────────────────────────────────────────────
function bibAbrirNuevoPeriodo() {
  _bibShowForm(`
    <div class="bib-form-box">
      <div class="bib-form-title">➕ Nuevo período</div>
      <div class="bib-form-row">
        <div class="bib-field" style="flex:1">
          <span class="bib-field-lbl">Nombre del período</span>
          <input class="bib-inp" id="bib-new-nombre" placeholder="Ej: Junio 2026 · 1ª quincena" maxlength="80"
            onkeydown="if(event.key==='Enter')bibCrearPeriodo()">
        </div>
      </div>
      <div id="bib-alert" style="display:none;color:var(--yel);font-size:9px;padding:4px 0"></div>
      <div class="bib-form-acts">
        <button class="bib-save-btn" onclick="bibCrearPeriodo()">✓ Crear período</button>
        <button class="bib-cancel-btn" onclick="bibCerrarForm()">Cancelar</button>
      </div>
    </div>`);
  setTimeout(() => document.getElementById('bib-new-nombre')?.focus(), 60);
}

async function bibCrearPeriodo() {
  const nombre = document.getElementById('bib-new-nombre')?.value?.trim();
  if (!nombre) { _bibAlert('Ingresá un nombre para el período.'); return; }
  try {
    await crearPeriodoConciliacion(nombre);
    bibCerrarForm();
    await _renderBiblioteca();
    await _renderPeriodSwitcher();
    if (typeof _showToast === 'function') _showToast(`✓ Período "${nombre}" creado`);
  } catch(e) {
    _bibAlert('Error: ' + e.message);
  }
}

// ── Agregar lote ─────────────────────────────────────────────────────
function bibAbrirFormLote(periodoId) {
  const hoy    = new Date();
  const desde  = hoy.toISOString().slice(0,10);
  const hastaD = new Date(hoy.getTime() + 6 * 86400000);
  const hasta  = hastaD.toISOString().slice(0,10);

  _bibShowForm(`
    <div class="bib-form-box">
      <div class="bib-form-title">➕ Nuevo lote</div>
      <input type="hidden" id="bib-lote-periodo" value="${periodoId}">
      <div class="bib-form-row">
        <div class="bib-field">
          <span class="bib-field-lbl">Fecha desde</span>
          <input type="date" class="bib-inp bib-inp-date" id="bib-lote-desde" value="${desde}">
        </div>
        <div class="bib-field">
          <span class="bib-field-lbl">Fecha hasta</span>
          <input type="date" class="bib-inp bib-inp-date" id="bib-lote-hasta" value="${hasta}">
        </div>
      </div>
      <div id="bib-alert" style="display:none;color:var(--yel);font-size:9px;padding:4px 0"></div>
      <div class="bib-form-acts">
        <button class="bib-save-btn" onclick="bibCrearLote()">✓ Crear lote</button>
        <button class="bib-cancel-btn" onclick="bibCerrarForm()">Cancelar</button>
      </div>
    </div>`);
}

async function bibCrearLote() {
  const periodoId = document.getElementById('bib-lote-periodo')?.value;
  const desde     = document.getElementById('bib-lote-desde')?.value;
  const hasta     = document.getElementById('bib-lote-hasta')?.value;
  if (!desde || !hasta) { _bibAlert('Ingresá las fechas del lote.'); return; }
  if (desde > hasta)    { _bibAlert('La fecha de inicio debe ser anterior al fin.'); return; }
  try {
    const lote = await agregarLotePeriodo(periodoId, { fechaDesde: desde, fechaHasta: hasta });
    bibCerrarForm();
    // Auto-abrir formulario de archivos para el lote nuevo
    await _renderBiblioteca();
    bibAbrirUploadLote(periodoId, lote.id);
    if (typeof _showToast === 'function')
      _showToast(`✓ Lote ${_bibFmtDate(desde)} – ${_bibFmtDate(hasta)} creado · subí los archivos`);
  } catch(e) {
    _bibAlert('Error: ' + e.message);
  }
}

// ── Subir archivos a un lote ─────────────────────────────────────────
function bibAbrirUploadLote(periodoId, loteId) {
  _PERIODO_ACTIVO_ID = periodoId;
  _LOTE_ACTIVO_ID    = loteId;

  _bibShowForm(`
    <div class="bib-form-box">
      <div class="bib-form-title">📎 Subir archivos al lote</div>
      <input type="hidden" id="bib-upload-periodo" value="${periodoId}">
      <input type="hidden" id="bib-upload-lote"    value="${loteId}">
      <div class="bib-form-row" style="flex-wrap:wrap">
        <div class="bib-field">
          <span class="bib-field-lbl">Tipo de archivo</span>
          <select class="bib-sel" id="bib-tipo">
            <option value="">— seleccionar —</option>
            ${(() => {
              const grupos = [...new Set(_BIB_TIPOS.map(t => t.modulo))];
              return grupos.map(grp => {
                const items = _BIB_TIPOS.filter(t => t.modulo === grp);
                return `<optgroup label="── ${grp} ──">${
                  items.map(t => `<option value="${t.id}">${t.label}</option>`).join('')
                }</optgroup>`;
              }).join('');
            })()}
          </select>
        </div>
        <div class="bib-field" style="flex:2">
          <span class="bib-field-lbl">Archivo</span>
          <label class="bib-file-label">
            <span class="bib-file-icon">📎</span>
            <span id="bib-file-name">Sin archivo elegido</span>
            <input type="file" id="bib-file" style="display:none"
              accept=".xlsx,.xls,.csv,.txt"
              onchange="_bibFileChange(this)">
          </label>
        </div>
      </div>
      <div id="bib-alert" style="display:none;color:var(--yel);font-size:9px;padding:4px 0"></div>
      <div class="bib-form-acts">
        <button class="bib-save-btn" id="bib-save-btn" onclick="guardarEnBiblioteca()">💾 Guardar archivo</button>
        <button class="bib-cancel-btn" onclick="bibCerrarForm()">Cerrar</button>
      </div>
      <div style="font-size:8px;color:var(--m2);margin-top:6px">
        Podés guardar varios archivos uno a uno para el mismo lote.
      </div>
    </div>`);
}

// ── Guardar archivo vinculado al lote activo ─────────────────────────
async function guardarEnBiblioteca() {
  const tipo      = document.getElementById('bib-tipo')?.value;
  const periodoId = document.getElementById('bib-upload-periodo')?.value || _PERIODO_ACTIVO_ID;
  const loteId    = document.getElementById('bib-upload-lote')?.value    || _LOTE_ACTIVO_ID;
  const fileInp   = document.getElementById('bib-file');
  const file      = fileInp?.files?.[0];

  if (!tipo)    { _bibAlert('Seleccioná el tipo de archivo.'); return; }
  if (!loteId)  { _bibAlert('No hay lote activo. Creá o seleccioná un lote primero.'); return; }
  if (!file)    { _bibAlert('Elegí un archivo para guardar.'); return; }

  const btn = document.getElementById('bib-save-btn');
  if (btn) { btn.disabled = true; btn.textContent = 'Guardando...'; }

  try {
    const bytes = await file.arrayBuffer();
    await guardarArchivoBiblioteca({ tipo, periodoId, loteId, nombre: file.name, bytes });
    fileInp.value = '';
    const fnEl = document.getElementById('bib-file-name');
    if (fnEl) fnEl.textContent = 'Sin archivo elegido';
    if (typeof _showToast === 'function') _showToast(`✓ "${file.name}" guardado`);
    await _renderBiblioteca();
    await _actualizarBadgeBiblioteca();
  } catch(e) {
    _bibAlert('Error al guardar: ' + e.message);
    console.error(e);
  } finally {
    if (btn) { btn.disabled = false; btn.textContent = '💾 Guardar archivo'; }
  }
}

// ══ ACCIONES DE LOTE ═════════════════════════════════════════════════

// Cargar archivos de un lote en memoria + restaurar resultado previo
async function bibCargarLote(periodoId, loteId) {
  _PERIODO_ACTIVO_ID = periodoId;
  _LOTE_ACTIVO_ID    = loteId;

  // Limpiar estado
  if (typeof RESULTADO   !== 'undefined') RESULTADO   = [];
  if (typeof CORREGIDAS  !== 'undefined') CORREGIDAS  = {};

  // Buscar resultado previo del lote
  const resGuardado     = await cargarResultadoLote(loteId);
  const tieneResultado  = resGuardado?.resultado?.length > 0;

  if (tieneResultado) {
    RESULTADO = resGuardado.resultado;
    _bibRestaurarUI();
  } else {
    _bibLimpiarUI();
  }

  // Cargar archivos del lote en FILES (para poder re-conciliar)
  const archivos = await listarArchivosDeLote(loteId);
  let cargados = 0;
  for (const rec of archivos) {
    try {
      const loader = _bibMakeLoader(rec);
      if (loader) { loader(); cargados++; }
    } catch(e) { console.warn('Error cargando archivo del lote:', rec.nombre, e); }
  }

  // Actualizar SESSION.periodoDesde/Hasta (necesario para cerrarPeriodo())
  const _per  = await obtenerPeriodoConciliacion(periodoId);
  const _lote = (_per?.lotes || []).find(l => l.id === loteId);
  if (_lote && typeof SESSION !== 'undefined') {
    SESSION.periodoDesde = _lote.fechaDesde;
    SESSION.periodoHasta = _lote.fechaHasta;
    SESSION.periodoMes   = _lote.fechaDesde?.slice(0, 7) || '';
    const pDesde = document.getElementById('param-desde');
    const pHasta = document.getElementById('param-hasta');
    if (pDesde) pDesde.value = _lote.fechaDesde;
    if (pHasta) pHasta.value = _lote.fechaHasta;
  }

  // Actualizar UI
  await _renderPeriodSwitcher();
  await _renderBiblioteca();
  document.getElementById('run-btn')?.removeAttribute('disabled');

  cerrarBiblioteca();

  // Obtener fechas del lote para el toast
  const per  = _per;
  const lote = _lote;
  const fechasStr = lote ? `${_bibFmtDate(lote.fechaDesde)} – ${_bibFmtDate(lote.fechaHasta)}` : '';

  if (typeof _showToast === 'function') {
    if (tieneResultado) {
      _showToast(`📅 Lote ${fechasStr} · ${RESULTADO.length.toLocaleString('es-AR')} ops restauradas${cargados ? ` · ${cargados} archivos` : ''}`, 4500);
    } else {
      _showToast(cargados
        ? `✓ ${cargados} archivos cargados (${fechasStr}) · listo para conciliar`
        : `📅 Lote ${fechasStr} activado · subí archivos para conciliar`, 4000);
    }
  }

  // Cambiar label del botón Conciliar
  const btn = document.getElementById('run-btn');
  const lbl = document.getElementById('run-lbl');
  const ico = document.getElementById('run-icon');
  if (btn && lbl && ico) {
    if (tieneResultado) {
      lbl.textContent = 'Reprocesar lote';
      ico.textContent = '↺';
    } else {
      lbl.textContent = 'Conciliar';
      ico.textContent = '▶';
    }
  }
}

// Ver todas las ops del período (merge de lotes conciliados)
async function bibVerPeriodo(periodoId) {
  _PERIODO_ACTIVO_ID = periodoId;
  _LOTE_ACTIVO_ID    = null;   // modo vista consolidada

  const ok = await cargarPeriodoCompleto(periodoId);

  // Setear fechas de SESSION desde el rango del período completo (para cerrarPeriodo)
  const _perFull = await obtenerPeriodoConciliacion(periodoId);
  if (_perFull && typeof SESSION !== 'undefined') {
    const _lotes = (_perFull.lotes || []).filter(l => l.estado === 'conciliado');
    if (_lotes.length) {
      const _desde = _lotes.map(l => l.fechaDesde).sort()[0];
      const _hasta = _lotes.map(l => l.fechaHasta).sort().reverse()[0];
      SESSION.periodoDesde = _desde;
      SESSION.periodoHasta = _hasta;
      SESSION.periodoMes   = _desde?.slice(0, 7) || '';
      const pDesde = document.getElementById('param-desde');
      const pHasta = document.getElementById('param-hasta');
      if (pDesde) pDesde.value = _desde;
      if (pHasta) pHasta.value = _hasta;
    }
  }

  await _renderPeriodSwitcher();

  if (!ok) {
    if (typeof _showToast === 'function') _showToast('Sin lotes conciliados en este período todavía', 3000);
    return;
  }

  _bibRestaurarUI();
  cerrarBiblioteca();

  const per  = await obtenerPeriodoConciliacion(periodoId);
  const lotes = (per?.lotes || []).filter(l => l.estado === 'conciliado');

  // Deshabilitar Conciliar en modo vista consolidada
  const btn = document.getElementById('run-btn');
  const lbl = document.getElementById('run-lbl');
  const ico = document.getElementById('run-icon');
  if (btn) {
    btn.disabled = true;
    if (lbl) lbl.textContent = 'Vista consolidada';
    if (ico) ico.textContent = '📊';
  }

  if (typeof _showToast === 'function')
    _showToast(`📊 ${per?.nombre || 'Período'} · ${RESULTADO.length.toLocaleString('es-AR')} ops · ${lotes.length} lotes`, 4500);
}

// Reprocesar todos los lotes del período en secuencia
async function bibReprocesarTodo(periodoId) {
  const per = await obtenerPeriodoConciliacion(periodoId);
  if (!per) return;
  const lotes = per.lotes || [];
  if (!lotes.length) { if (typeof _showToast === 'function') _showToast('Sin lotes para reprocesar'); return; }

  const ok = confirm(
    `¿Reprocesar los ${lotes.length} lote(s) del período "${per.nombre}"?\n` +
    `Cada lote se conciliará con sus archivos guardados y se actualizarán los resultados.`
  );
  if (!ok) return;

  for (let i = 0; i < lotes.length; i++) {
    const lote = lotes[i];
    if (typeof _showToast === 'function')
      _showToast(`⟳ Procesando lote ${i + 1} de ${lotes.length}: ${_bibFmtDate(lote.fechaDesde)} – ${_bibFmtDate(lote.fechaHasta)}`, 60000);

    // Activar lote
    _PERIODO_ACTIVO_ID = periodoId;
    _LOTE_ACTIVO_ID    = lote.id;

    // Limpiar y cargar archivos del lote
    if (typeof RESULTADO  !== 'undefined') RESULTADO  = [];
    if (typeof CORREGIDAS !== 'undefined') CORREGIDAS = {};

    const archivos = await listarArchivosDeLote(lote.id);
    for (const rec of archivos) {
      try {
        const loader = _bibMakeLoader(rec);
        if (loader) loader();
      } catch(e) { console.warn('Error cargando', rec.nombre, e); }
    }

    // Pausa para que los archivos se carguen en FILES
    await new Promise(r => setTimeout(r, 300));

    // Conciliar
    try {
      await conciliar();
      await new Promise(r => setTimeout(r, 200));
    } catch(e) {
      console.error(`Error conciliando lote ${lote.id}:`, e);
    }
  }

  // Al terminar, cargar vista consolidada
  await bibVerPeriodo(periodoId);
  if (typeof _showToast === 'function')
    _showToast(`✅ ${lotes.length} lotes reprocesados · ${RESULTADO.length.toLocaleString('es-AR')} ops totales`, 5000);
}

// ══ ELIMINAR ══════════════════════════════════════════════════════════

async function bibEliminarPeriodo(periodoId) {
  const per = await obtenerPeriodoConciliacion(periodoId);
  if (!per) return;
  const nl = (per.lotes || []).length;
  const msg = nl
    ? `¿Eliminar el período "${per.nombre}" con ${nl} lote(s)?\nSe borrarán todos sus archivos y resultados.`
    : `¿Eliminar el período "${per.nombre}"?`;
  if (!confirm(msg)) return;

  await eliminarPeriodoConciliacion(periodoId);

  if (_PERIODO_ACTIVO_ID === periodoId) {
    _PERIODO_ACTIVO_ID = null;
    _LOTE_ACTIVO_ID    = null;
    _bibLimpiarUI();
  }

  await _renderBiblioteca();
  await _renderPeriodSwitcher();
  if (typeof _showToast === 'function') _showToast(`Período "${per.nombre}" eliminado`);
}

async function bibEliminarLote(periodoId, loteId) {
  if (!confirm('¿Eliminar este lote y todos sus archivos y resultados?')) return;
  await eliminarLotePeriodo(periodoId, loteId);
  if (_LOTE_ACTIVO_ID === loteId) {
    _LOTE_ACTIVO_ID = null;
    _bibLimpiarUI();
  }
  await _renderBiblioteca();
  await _renderPeriodSwitcher();
  if (typeof _showToast === 'function') _showToast('Lote eliminado');
}

// ══ FILE INPUT HELPERS ═════════════════════════════════════════════════

function _bibFileChange(inp) {
  const fnEl  = document.getElementById('bib-file-name');
  const nombre = inp.files[0]?.name || 'Sin archivo elegido';
  if (fnEl) fnEl.textContent = nombre;
  _bibAutoTipo(nombre);
}

function _bibAutoTipo(nombre) {
  const n   = nombre.toLowerCase();
  const sel = document.getElementById('bib-tipo');
  if (!sel || sel.value) return;
  const esCelular = n.includes('celular');
  const esCuota   = n.includes('cuota') || n.includes('gocuota');
  const esLiq     = n.includes('liquid') || n.includes('liq');
  if      (n.includes('skylab'))                               sel.value = 'SKYLAB';
  else if (n.includes('terminal'))                             sel.value = 'TERMINALES';
  else if (n.includes('fiserv') && n.includes('contra'))       sel.value = 'FISERV_CTR';
  else if (n.includes('getpos') && n.includes('contra'))       sel.value = 'GETPOS_CTR';
  else if (n.includes('getpos'))                               sel.value = 'GETPOS';
  // GoC Liquidaciones (tiene "cuota"/"celular" + "liquid") — antes de los de OPERACIONES
  else if (esCelular && esLiq)                                 sel.value = 'GOC_LIQ_CELULAR';
  else if (esCuota   && esLiq)                                 sel.value = 'GOC_LIQ_PAGOS';
  // GoC Operaciones
  else if (esCelular)                                          sel.value = 'GOC_CELULAR';
  else if (esCuota)                                            sel.value = 'GOC_PAGOS';
  else if (n.includes('venta'))                                sel.value = 'GOC_VENTAS';
  // FISERV / LIQUIDACIONES xlsx (sin "cuota"/"celular")
  else if (n.includes('fiserv') || (esLiq && (n.endsWith('.xlsx') || n.endsWith('.xls'))))
                                                               sel.value = 'FISERV_LIQ';
  else if (esLiq)                                              sel.value = 'LIQUIDACION';
}

// ══ BADGE Y PILL ══════════════════════════════════════════════════════

async function _actualizarBadgeBiblioteca(archivos) {
  try {
    const todos = archivos || (await listarArchivosBiblioteca());
    const el = document.getElementById('bib-badge');
    if (el) el.textContent = todos.length || '';
  } catch {}
}

// ══ PILL DEL PERÍODO ACTIVO EN TOPBAR ════════════════════════════════

function _actualizarPillPeriodo() {
  const pill = document.getElementById('sb-periodo-activo-pill');
  if (!pill) return;

  if (!_PERIODO_ACTIVO_ID) {
    pill.textContent = 'Sin período';
    pill.className   = 'tb-period-empty';
    return;
  }

  pill.textContent = '📅 Cargando...';
  pill.className   = 'tb-period-label';

  obtenerPeriodoConciliacion(_PERIODO_ACTIVO_ID).then(per => {
    if (!per) return;
    let label = per.nombre;
    if (_LOTE_ACTIVO_ID) {
      const lote = (per.lotes || []).find(l => l.id === _LOTE_ACTIVO_ID);
      if (lote) label += ` · ${_bibFmtDate(lote.fechaDesde)}–${_bibFmtDate(lote.fechaHasta)}`;
    }
    if (pill) {
      pill.textContent = label;
      pill.className   = 'tb-period-label';
    }
  }).catch(() => {});
}

// ══ PERIOD SWITCHER (dropdown del topbar) ════════════════════════════

async function _renderPeriodSwitcher() {
  const el = document.getElementById('sb-period-switcher');
  _actualizarPillPeriodo();
  if (!el) return;

  const periodos = await listarPeriodosConciliacion();

  if (!periodos.length) {
    el.innerHTML = `<div class="sb-per-empty">Sin períodos. Abrí 📁 Biblioteca para crear uno.</div>`;
    return;
  }

  el.innerHTML = periodos.map(per => {
    const lotes     = per.lotes || [];
    const totalOps  = lotes.reduce((s, l) => s + (l.nOps || 0), 0);
    const lotesOk   = lotes.filter(l => l.estado === 'conciliado').length;
    const isActive  = per.id === _PERIODO_ACTIVO_ID;
    const dotCls    = isActive ? 'active-dot' : (lotesOk > 0 ? 'concil' : '');
    const dot       = isActive ? '●' : (lotesOk > 0 ? '✓' : '○');
    const metaTxt   = lotesOk > 0
      ? `${totalOps.toLocaleString('es-AR')} ops`
      : `${lotes.length} lote${lotes.length !== 1 ? 's' : ''}`;

    return `<button class="sb-per-btn ${isActive ? 'active' : ''}"
        onclick="bibVerPeriodo('${per.id}');_closePeriodDrop()"
        title="${per.nombre} · ${lotes.length} lotes${lotesOk > 0 ? ` · ${totalOps} ops` : ''}">
      <span class="sb-per-dot ${dotCls}">${dot}</span>
      <span class="sb-per-lbl">${per.nombre}</span>
      <span class="sb-per-meta">${metaTxt}</span>
    </button>`;
  }).join('');
}

// ══ UI HELPERS ════════════════════════════════════════════════════════

function _bibRestaurarUI() {
  try {
    const strip = document.getElementById('tab-strip-cruce');
    if (strip) strip.style.display = 'flex';
    document.querySelectorAll('#mod-cruce .tab-body').forEach(t => t.classList.remove('active'));
    document.getElementById('t-empty')?.classList.remove('active');
    document.getElementById('t-all')?.classList.add('active');
    document.getElementById('dl-bar')?.classList.add('show');
    if (strip) {
      strip.querySelectorAll('.tb').forEach(b => b.classList.remove('active'));
      strip.querySelectorAll('.tb')[1]?.classList.add('active');
    }
    if (typeof renderTodo === 'function')      renderTodo();
    if (typeof setupDownloads === 'function')  setupDownloads();
  } catch(e) { console.warn('_bibRestaurarUI:', e); }
}

function _bibLimpiarUI() {
  try {
    const strip = document.getElementById('tab-strip-cruce');
    if (strip) strip.style.display = 'none';
    document.querySelectorAll('.tab-body').forEach(t => t.classList.remove('active'));
    document.getElementById('t-empty')?.classList.add('active');
    const dash = document.getElementById('dashboard');
    if (dash) dash.style.display = 'none';
    document.getElementById('dl-bar')?.classList.remove('show');
    if (typeof clearLog === 'function') clearLog();
  } catch(e) { console.warn('_bibLimpiarUI:', e); }
}

// ══ VACIAR ARCHIVOS HUÉRFANOS ═════════════════════════════════════════

async function bibLimpiarArchivos() {
  const db = await dbOpen();
  const count = await new Promise(res => {
    const tx = db.transaction('archivos', 'readonly');
    tx.objectStore('archivos').count().onsuccess = e => res(e.target.result);
  });
  if (count === 0) {
    _showToast('La biblioteca ya está vacía.');
    return;
  }
  if (!confirm(`Hay ${count} archivo(s) guardados en la biblioteca.\n¿Eliminar todos? Esta acción no se puede deshacer.`)) return;
  await new Promise((res, rej) => {
    const tx = db.transaction('archivos', 'readwrite');
    const req = tx.objectStore('archivos').clear();
    req.onsuccess = () => res();
    req.onerror   = e => rej(e.target.error);
  });
  await _actualizarBadgeBiblioteca();
  _renderBiblioteca();
  _showToast(`🗑 ${count} archivo(s) eliminados de la biblioteca.`);
}

// ══ INIT ══════════════════════════════════════════════════════════════

document.addEventListener('DOMContentLoaded', () => {
  _renderPeriodSwitcher();
  _actualizarBadgeBiblioteca();
});
