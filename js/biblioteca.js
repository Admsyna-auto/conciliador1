// ═══════════════════════════════════════════════════════════════════
// BIBLIOTECA.JS — Almacenamiento persistente de archivos por período
// Permite subir, clasificar y recargar archivos de cualquier mes
// ═══════════════════════════════════════════════════════════════════

const _BIB_TIPOS = [
  { id:'SKYLAB',      label:'Skylab (principal)',      accept:'.xlsx,.xls', badge:'SK', bg:'#0f2544', color:'#38bdf8' },
  { id:'TERMINALES',  label:'Terminales',              accept:'.xlsx,.xls', badge:'TER',bg:'#1a1f2e', color:'#8ba3c4' },
  { id:'FISERV_LIQ',  label:'FISERV Liquidación',      accept:'.xlsx,.xls', badge:'FIS',bg:'#0f1e44', color:'#4f8ef7' },
  { id:'FISERV_CTR',  label:'FISERV Contracargos',     accept:'.xlsx,.xls', badge:'FC', bg:'#0f1e44', color:'#f87171' },
  { id:'GETPOS',      label:'GETPOS',                  accept:'.xlsx,.xls', badge:'GP', bg:'#0f2d20', color:'#34d399' },
  { id:'GETPOS_CTR',  label:'GETPOS Contracargos',     accept:'.xlsx,.xls', badge:'GC', bg:'#0f2d20', color:'#f87171' },
  { id:'GOC_PAGOS',   label:'Go Cuotas CSV',           accept:'.csv,.txt',  badge:'GQ', bg:'#0f2d20', color:'#4ade80' },
  { id:'GOC_CELULAR', label:'Go Celular CSV',          accept:'.csv,.txt',  badge:'GK', bg:'#1a1040', color:'#a78bfa' },
  { id:'GOC_VENTAS',  label:'GoC Ventas (IMEI)',       accept:'.xlsx,.xls', badge:'GV', bg:'#1a1f10', color:'#fbbf24' },
  { id:'LIQUIDACION', label:'Liquidaciones (Cobros)',  accept:'.xlsx,.xls', badge:'LIQ',bg:'#291000', color:'#fb923c' },
];

// ── Mapeo de tipo → función de carga en sesión ───────────────────────
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
    case 'GOC_PAGOS':   return () => loadGocPagos(fi, 'GOCUOTAS');
    case 'GOC_CELULAR': return () => loadGocPagos(fi, 'GOCELULAR');
    case 'GOC_VENTAS':  return () => loadGocVentas(fi);
    case 'LIQUIDACION': return () => loadLiquidaciones(fi);
    default:            return null;
  }
}

// ══ CRUD IndexedDB (usa helpers de state.js) ══════════════════════════

async function guardarArchivoBiblioteca({ tipo, periodo, nombre, bytes }) {
  const id = `bib_${tipo}_${Date.now()}`;
  const kb = Math.round(bytes.byteLength / 1024);
  await dbPut('archivos', { id, tipo, periodo, nombre, bytes, kb, fechaCarga: new Date().toISOString() });
  return id;
}

async function listarArchivosBiblioteca() {
  try {
    return (await dbGetAll('archivos'))
      .sort((a, b) => b.periodo.localeCompare(a.periodo) || b.fechaCarga.localeCompare(a.fechaCarga));
  } catch { return []; }
}

async function obtenerArchivoBiblioteca(id) {
  return dbGet('archivos', id);
}

async function eliminarArchivoBiblioteca(id) {
  await dbDelete('archivos', id);
}

// ══ HELPERS UI ════════════════════════════════════════════════════════

function _bibFmtPeriodo(p) {
  if (!p) return '—';
  const [y, m] = p.split('-');
  if (!y || !m) return p;
  const M = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
  return `${M[parseInt(m, 10) - 1] || m} ${y}`;
}

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

// ══ RENDER TABLA ══════════════════════════════════════════════════════

async function _renderBiblioteca() {
  const tbody   = document.getElementById('bib-tbody');
  const cntEl   = document.getElementById('bib-count');
  const selPer  = document.getElementById('bib-flt-per');
  if (!tbody) return;

  tbody.innerHTML = `<tr><td colspan="7" style="text-align:center;color:var(--m2);padding:28px 0">
    Cargando...</td></tr>`;

  const todos    = await listarArchivosBiblioteca();
  const filtTipo = document.getElementById('bib-flt-tipo')?.value || '';
  const filtPer  = selPer?.value || '';

  // Poblar selector de períodos (dinámico según lo que hay guardado)
  if (selPer) {
    const periodos = [...new Set(todos.map(a => a.periodo))].sort().reverse();
    const cur = selPer.value;
    selPer.innerHTML = '<option value="">Todos los períodos</option>' +
      periodos.map(p =>
        `<option value="${p}" ${p === cur ? 'selected' : ''}>${_bibFmtPeriodo(p)}</option>`
      ).join('');
  }

  const filtrados = todos.filter(a =>
    (!filtTipo || a.tipo === filtTipo) &&
    (!filtPer  || a.periodo === filtPer)
  );

  if (cntEl) cntEl.textContent = `${filtrados.length} de ${todos.length} archivos`;

  if (!filtrados.length) {
    const msg = todos.length
      ? 'Sin resultados para los filtros aplicados'
      : '¡Todavía no hay archivos guardados! Usá el formulario de arriba para agregar el primero.';
    tbody.innerHTML = `<tr><td colspan="7" style="text-align:center;color:var(--m2);padding:48px 0;font-size:10px">
      📭 ${msg}</td></tr>`;
    return;
  }

  // Agrupar por período para mostrar filas de grupo
  const byPer = {};
  filtrados.forEach(a => {
    if (!byPer[a.periodo]) byPer[a.periodo] = [];
    byPer[a.periodo].push(a);
  });

  let html = '';
  for (const [per, arcs] of Object.entries(byPer)) {
    // Fila de cabecera del período
    html += `<tr class="bib-per-row">
      <td colspan="5" style="
        padding:8px 14px;background:var(--s3);
        color:var(--cyn);font-family:var(--head);font-size:11px;font-weight:700;
        border-top:1px solid var(--b2);">
        📅 ${_bibFmtPeriodo(per)}
        <span style="font-size:8px;font-weight:400;color:var(--m2);margin-left:8px">${arcs.length} archivos</span>
      </td>
      <td colspan="2" style="background:var(--s3);border-top:1px solid var(--b2);padding:8px 14px;text-align:right">
        <button class="bib-btn-load-all" onclick="bibCargarPeriodo('${per}')"
          title="Cargar todos los archivos de este período en la sesión">
          ▶▶ Cargar período completo
        </button>
      </td>
    </tr>`;

    // Filas de archivos del período
    for (const a of arcs) {
      const ti  = _bibTipoInfo(a.tipo);
      const fec = a.fechaCarga ? a.fechaCarga.slice(0, 16).replace('T', ' ') : '—';
      html += `<tr class="bib-file-row">
        <td>
          <div style="display:flex;align-items:center;gap:8px">
            <div style="
              width:28px;height:28px;border-radius:5px;
              background:${ti.bg};color:${ti.color};
              font-size:7.5px;font-weight:700;
              display:flex;align-items:center;justify-content:center;flex-shrink:0;
              border:1px solid ${ti.color}33">
              ${ti.badge}
            </div>
            <div>
              <div style="font-size:10px;font-weight:600;color:var(--txt)">${ti.label}</div>
            </div>
          </div>
        </td>
        <td style="font-weight:600;color:var(--cyn);font-size:10px">${_bibFmtPeriodo(a.periodo)}</td>
        <td style="max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;
          color:var(--m1);font-size:9px" title="${a.nombre}">${a.nombre}</td>
        <td style="color:var(--m2);font-size:9px;text-align:right">${a.kb ?? '?'} KB</td>
        <td style="color:var(--m2);font-size:9px">${fec}</td>
        <td>
          <button class="bib-btn-load" onclick="bibCargarEnSesion('${a.id}')">
            ▶ Cargar en sesión
          </button>
        </td>
        <td>
          <button class="bib-btn-del" onclick="bibEliminar('${a.id}',this)" title="Eliminar de biblioteca">✕</button>
        </td>
      </tr>`;
    }
  }
  tbody.innerHTML = html;
}

// ══ ACCIONES ══════════════════════════════════════════════════════════

async function guardarEnBiblioteca() {
  const tipo    = document.getElementById('bib-tipo')?.value;
  const periodo = document.getElementById('bib-periodo')?.value;
  const fileInp = document.getElementById('bib-file');
  const file    = fileInp?.files?.[0];

  if (!tipo)    { _bibAlert('Seleccioná el tipo de archivo.'); return; }
  if (!periodo) { _bibAlert('Seleccioná el período (mes/año).'); return; }
  if (!file)    { _bibAlert('Elegí un archivo para guardar.'); return; }

  const btn = document.getElementById('bib-save-btn');
  if (btn) { btn.disabled = true; btn.textContent = 'Guardando...'; }

  try {
    const bytes = await file.arrayBuffer();
    await guardarArchivoBiblioteca({ tipo, periodo, nombre: file.name, bytes });
    // Reset form
    fileInp.value = '';
    const fnEl = document.getElementById('bib-file-name');
    if (fnEl) fnEl.textContent = 'Sin archivo elegido';
    // No resetear tipo/periodo para facilitar carga de varios archivos del mismo mes
    if (typeof _showToast === 'function')
      _showToast(`✓ "${file.name}" guardado · ${_bibFmtPeriodo(periodo)}`);
    await _renderBiblioteca();
  } catch (e) {
    _bibAlert('Error al guardar: ' + e.message);
    console.error(e);
  } finally {
    if (btn) { btn.disabled = false; btn.textContent = '💾 Guardar'; }
  }
}

async function bibCargarEnSesion(id) {
  const rec = await obtenerArchivoBiblioteca(id);
  if (!rec) { _bibAlert('Archivo no encontrado en la biblioteca.'); return; }

  const loader = _bibMakeLoader(rec);
  if (!loader) { _bibAlert('No hay cargador disponible para: ' + rec.tipo); return; }

  try {
    loader();
    cerrarBiblioteca();
    if (typeof _showToast === 'function')
      _showToast(`✓ "${rec.nombre}" · ${_bibFmtPeriodo(rec.periodo)} → cargado en sesión`);
  } catch (e) {
    _bibAlert('Error al cargar: ' + e.message);
    console.error(e);
  }
}

async function bibCargarPeriodo(periodo) {
  const todos = await listarArchivosBiblioteca();
  const delPer = todos.filter(a => a.periodo === periodo);
  if (!delPer.length) return;

  let cargados = 0;
  const errores = [];
  for (const rec of delPer) {
    try {
      const loader = _bibMakeLoader(rec);
      if (loader) { loader(); cargados++; }
    } catch (e) { errores.push(rec.nombre); }
  }

  cerrarBiblioteca();
  if (typeof _showToast === 'function') {
    const msg = errores.length
      ? `✓ ${cargados} archivos cargados (${errores.length} errores)`
      : `✓ ${cargados} archivos cargados · ${_bibFmtPeriodo(periodo)}`;
    _showToast(msg, errores.length ? 5000 : 3500);
  }
}

async function bibEliminar(id, btnEl) {
  if (!confirm('¿Eliminar este archivo de la biblioteca?\n(No afecta la sesión actual)')) return;
  if (btnEl) btnEl.disabled = true;
  await eliminarArchivoBiblioteca(id);
  if (typeof _showToast === 'function') _showToast('Archivo eliminado de la biblioteca');
  await _renderBiblioteca();
}

// ══ FILE INPUT HELPERS ════════════════════════════════════════════════

function _bibFileChange(inp) {
  const fnEl = document.getElementById('bib-file-name');
  const nombre = inp.files[0]?.name || 'Sin archivo elegido';
  if (fnEl) fnEl.textContent = nombre;
  _bibAutoTipo(nombre);
}

function _bibAutoTipo(nombre) {
  const n = nombre.toLowerCase();
  const sel = document.getElementById('bib-tipo');
  if (!sel) return;
  // Auto-detect solo si no hay nada seleccionado
  if (sel.value) return;
  if      (n.includes('skylab'))                           sel.value = 'SKYLAB';
  else if (n.includes('terminal'))                         sel.value = 'TERMINALES';
  else if (n.includes('fiserv') && n.includes('contra'))   sel.value = 'FISERV_CTR';
  else if (n.includes('fiserv') || n.includes('liq'))      sel.value = 'FISERV_LIQ';
  else if (n.includes('getpos') && n.includes('contra'))   sel.value = 'GETPOS_CTR';
  else if (n.includes('getpos'))                           sel.value = 'GETPOS';
  else if (n.includes('celular'))                          sel.value = 'GOC_CELULAR';
  else if (n.includes('cuota') || n.includes('gocuota'))   sel.value = 'GOC_PAGOS';
  else if (n.includes('venta'))                            sel.value = 'GOC_VENTAS';
  else if (n.includes('liquid'))                           sel.value = 'LIQUIDACION';
}

// ══ BADGE EN SIDEBAR (cantidad de archivos guardados) ════════════════

async function _actualizarBadgeBiblioteca() {
  try {
    const todos = await listarArchivosBiblioteca();
    const el = document.getElementById('bib-badge');
    if (el) el.textContent = todos.length || '';
  } catch { /* ignore */ }
}
