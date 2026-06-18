// ═══════════════════════════════════════════════════════════════════
// PROCESADORAS.JS — Módulos dedicados por procesadora (FISERV / GETPOS)
// ═══════════════════════════════════════════════════════════════════

// Estado de filtros por procesadora
const _PROC_FILTROS = {
  FISERV: { suc:'', estado:'', search:'' },
  GETPOS: { suc:'', estado:'', search:'' },
  PRISMA: { suc:'', estado:'', search:'' },
};

// ── Filtro de filas por procesadora ──────────────────────────────────
function _procRows(procId) {
  if (typeof RESULTADO === 'undefined' || !RESULTADO.length) return [];
  if (procId === 'FISERV') return RESULTADO.filter(r => !r.sky?.esGETPos && !r.sky?.esGOCUOTAS && !r.sky?.esPRISMA);
  if (procId === 'GETPOS') return RESULTADO.filter(r =>  r.sky?.esGETPos && !r.sky?.esGOCUOTAS && !r.sky?.esPRISMA);
  if (procId === 'PRISMA') return RESULTADO.filter(r =>  r.sky?.esPRISMA);
  return [];
}

// ── Aplicar filtros activos ───────────────────────────────────────────
function _procFiltrar(procId, rows) {
  const f = _PROC_FILTROS[procId];
  if (!f) return rows;
  let out = rows;
  if (f.suc)    out = out.filter(r => r.sky?.suc === f.suc);
  if (f.estado) out = out.filter(r => r.estado === f.estado);
  if (f.search) {
    const s = f.search.toLowerCase();
    out = out.filter(r => {
      const sky = r.sky;
      return (sky?.cupon??'').toString().toLowerCase().includes(s)
          || (sky?.vendedor??'').toLowerCase().includes(s)
          || (sky?.asiento??'').toString().toLowerCase().includes(s)
          || fmtARS(sky?.monto??0).includes(s);
    });
  }
  return out;
}

// ── Actualizar filtro y re-renderizar tabla ───────────────────────────
function _procAplicarFiltro(procId, campo, valor) {
  if (_PROC_FILTROS[procId]) _PROC_FILTROS[procId][campo] = valor;
  _procRenderTabla(procId, _procTabActiva(procId));
}

function _procLimpiarFiltros(procId) {
  if (_PROC_FILTROS[procId]) Object.keys(_PROC_FILTROS[procId]).forEach(k => _PROC_FILTROS[procId][k] = '');
  const p = procId.toLowerCase();
  ['suc','estado','search'].forEach(k => {
    const el = document.getElementById(`pf-${p}-${k}`);
    if (el) el.value = '';
  });
  _procRenderTabla(procId, _procTabActiva(procId));
}

function _procTabActiva(procId) {
  const p = procId.toLowerCase();
  const strip = document.getElementById(`tab-strip-proc-${p}`);
  if (!strip) return 'todo';
  const active = strip.querySelector('.tb.active');
  return active?.dataset?.tab || 'todo';
}

// ── Render de la tabla para una sub-tab ──────────────────────────────
function _procRenderTabla(procId, tab) {
  const p = procId.toLowerCase();
  const allRows = _procRows(procId);

  let rows;
  switch (tab) {
    case 'sinmatch': rows = allRows.filter(r => r.estado === 'SIN MATCH'); break;
    case 'mal':      rows = allRows.filter(r => r.estado?.startsWith('MAL FACTURADO')); break;
    case 'com':      rows = allRows.filter(r => r.estado === 'COM. ERRADO'); break;
    case 'ok':       rows = allRows.filter(r => r.estado?.startsWith('OK')); break;
    default:         rows = allRows;
  }

  const filtradas = _procFiltrar(procId, rows);
  const stats = document.getElementById(`pf-${p}-stats`);
  if (stats) stats.textContent = `${filtradas.length} / ${rows.length}`;

  renderTable(`tbl-proc-${p}-${tab}`, filtradas);
}

// ── Render del módulo completo ────────────────────────────────────────
function renderModuloProc(procId) {
  const p   = procId.toLowerCase();
  const col  = procId === 'FISERV' ? 'var(--acc)' : procId === 'PRISMA' ? '#06b6d4' : 'var(--grn)';
  const icon = procId === 'FISERV' ? '🏦' : procId === 'PRISMA' ? '🌐' : '🏧';
  const panel = document.getElementById(`mod-${p}`);
  if (!panel) return;

  if (typeof RESULTADO === 'undefined' || !RESULTADO.length) {
    panel.innerHTML = `
      <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;
        height:100%;gap:14px;color:var(--m2);text-align:center;padding:40px">
        <div style="font-size:42px;opacity:.15">${icon}</div>
        <div style="font-size:16px;font-weight:700;color:var(--txt);opacity:.4">${procId}</div>
        <p style="font-size:10px;max-width:380px;line-height:1.8">
          Cargá los archivos de <b style="color:${col}">${procId}</b> y ejecutá el Cruce Automático.
        </p>
      </div>`;
    return;
  }

  const allRows = _procRows(procId);
  if (!allRows.length) {
    panel.innerHTML = `
      <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;
        height:100%;gap:14px;color:var(--m2);text-align:center;padding:40px">
        <div style="font-size:42px;opacity:.15">${icon}</div>
        <div style="font-size:14px;font-weight:600;color:var(--txt);opacity:.4">Sin operaciones ${procId}</div>
        <p style="font-size:10px;max-width:380px;line-height:1.8">
          No se encontraron operaciones de <b style="color:${col}">${procId}</b> en el resultado.
          Verificá que el archivo de la procesadora esté cargado.
        </p>
      </div>`;
    return;
  }

  // ── Calcular KPIs ─────────────────────────────────────────────────
  const total    = allRows.length;
  const ok       = allRows.filter(r => r.estado?.startsWith('OK')).length;
  const sinMatch = allRows.filter(r => r.estado === 'SIN MATCH').length;
  const mal      = allRows.filter(r => r.estado?.startsWith('MAL FACTURADO')).length;
  const com      = allRows.filter(r => r.estado === 'COM. ERRADO').length;
  const pct      = total ? (ok / total * 100).toFixed(1) : '0.0';
  const pctCls   = parseFloat(pct) >= 90 ? 'grn' : parseFloat(pct) >= 70 ? 'yel' : 'red';

  // Diferencia total: suma abs(difVal) de filas con proc y monto distinto
  let difTotal = 0;
  allRows.forEach(r => {
    const procMonto = r.procMontoNorm ?? r.proc?.monto ?? null;
    if (procMonto !== null && r.sky?.monto !== undefined) {
      const dif = r.sky.monto - procMonto;
      if (Math.abs(dif) > 0.01) difTotal += Math.abs(dif);
    }
  });

  // ── Sucursales para el filtro ─────────────────────────────────────
  const sucs = [...new Set(allRows.map(r => r.sky?.suc).filter(Boolean))].sort((a,b)=>+a-+b);

  panel.innerHTML = `
  <div style="display:flex;flex-direction:column;flex:1;min-height:0;overflow:hidden">

    <!-- KPIs -->
    <div style="display:grid;grid-template-columns:repeat(6,1fr);gap:8px;
      padding:12px 18px;background:var(--s1);border-bottom:1px solid var(--b1);flex-shrink:0">
      <div class="dif-kpi" style="border-color:${col}33">
        <div class="dif-kpi-lbl">${icon} Total ${procId}</div>
        <div class="dif-kpi-val cyn">${total.toLocaleString('es-AR')}</div>
        <div style="font-size:8px;color:var(--m2)">${fmtARS(allRows.reduce((s,r)=>s+Math.abs(r.sky?.monto||0),0))}</div>
      </div>
      <div class="dif-kpi" style="border-color:rgba(52,211,153,.25)">
        <div class="dif-kpi-lbl">✓ OK</div>
        <div class="dif-kpi-val grn">${ok.toLocaleString('es-AR')}</div>
        <div style="font-size:8px;color:var(--grn)">${fmtARS(allRows.filter(r=>r.estado?.startsWith('OK')).reduce((s,r)=>s+Math.abs(r.sky?.monto||0),0))}</div>
      </div>
      <div class="dif-kpi" style="border-color:rgba(248,113,113,.25)">
        <div class="dif-kpi-lbl">✗ Sin Match</div>
        <div class="dif-kpi-val red">${sinMatch.toLocaleString('es-AR')}</div>
        <div style="font-size:8px;color:var(--red)">${fmtARS(allRows.filter(r=>r.estado==='SIN MATCH').reduce((s,r)=>s+Math.abs(r.sky?.monto||0),0))}</div>
      </div>
      <div class="dif-kpi" style="border-color:rgba(251,146,60,.25)">
        <div class="dif-kpi-lbl">⚠ Mal Fact.</div>
        <div class="dif-kpi-val org">${mal.toLocaleString('es-AR')}</div>
        <div style="font-size:8px;color:var(--org)">${com > 0 ? `+ ${com} com. errado` : '—'}</div>
      </div>
      <div class="dif-kpi">
        <div class="dif-kpi-lbl">% OK</div>
        <div class="dif-kpi-val ${pctCls}">${pct}%</div>
        <div style="font-size:8px;color:var(--m2)">${ok} de ${total}</div>
      </div>
      <div class="dif-kpi" style="border-color:rgba(248,113,113,.2)">
        <div class="dif-kpi-lbl">Δ Diferencias $</div>
        <div class="dif-kpi-val ${difTotal > 0 ? 'red' : 'grn'}">${fmtARS(difTotal)}</div>
        <div style="font-size:8px;color:var(--m2)">suma absoluta</div>
      </div>
    </div>

    <!-- Filter bar -->
    <div class="filter-bar" style="flex-shrink:0;flex-wrap:wrap;align-items:center">
      <input class="filter-inp" id="pf-${p}-search"
        placeholder="🔍 Cupón, vendedor, asiento..."
        oninput="_procAplicarFiltro('${procId}','search',this.value)"
        style="flex:1;min-width:160px">
      <select class="filter-sel" id="pf-${p}-suc"
        onchange="_procAplicarFiltro('${procId}','suc',this.value)">
        <option value="">Todas las suc.</option>
        ${sucs.map(s=>`<option value="${s}">${s}</option>`).join('')}
      </select>
      <select class="filter-sel" id="pf-${p}-estado"
        onchange="_procAplicarFiltro('${procId}','estado',this.value)">
        <option value="">Todos los estados</option>
        <option value="OK">✓ OK</option>
        <option value="OK (equiv.)">✓ OK (equiv.)</option>
        <option value="SIN MATCH">✗ Sin Match</option>
        <option value="MAL FACTURADO">⚠ Mal Facturado</option>
        <option value="COM. ERRADO">💱 Com. Errado</option>
        <option value="ANULACION SIN COBRO">✗ Anul. S/Cobro</option>
        <option value="REVISION URGENTE">⚠ Rev. Urgente</option>
      </select>
      <button class="btn-clear" onclick="_procLimpiarFiltros('${procId}')">✕</button>
      <span class="filter-stats" id="pf-${p}-stats" style="margin-left:auto;font-size:9px"></span>
    </div>

    <!-- Tab strip -->
    <div class="tab-strip" id="tab-strip-proc-${p}">
      <button class="tb active" data-tab="todo"
        onclick="showTab('ptab-${p}-todo','tab-strip-proc-${p}',this);_procRenderTabla('${procId}','todo')">
        📋 Todo <span class="cnt">${total}</span>
      </button>
      <button class="tb" data-tab="ok"
        onclick="showTab('ptab-${p}-ok','tab-strip-proc-${p}',this);_procRenderTabla('${procId}','ok')">
        ✓ OK <span class="cnt" style="background:rgba(52,211,153,.15);color:var(--grn)">${ok}</span>
      </button>
      <button class="tb" data-tab="sinmatch"
        onclick="showTab('ptab-${p}-sinmatch','tab-strip-proc-${p}',this);_procRenderTabla('${procId}','sinmatch')" style="color:${sinMatch>0?'var(--red)':''}">
        ✗ Sin Match <span class="cnt" style="background:rgba(248,113,113,.15);color:var(--red)">${sinMatch}</span>
      </button>
      ${mal > 0 ? `<button class="tb" data-tab="mal"
        onclick="showTab('ptab-${p}-mal','tab-strip-proc-${p}',this);_procRenderTabla('${procId}','mal')" style="color:var(--org)">
        ⚠ Mal Fact. <span class="cnt" style="background:rgba(251,146,60,.15);color:var(--org)">${mal}</span>
      </button>` : ''}
      ${com > 0 ? `<button class="tb" data-tab="com"
        onclick="showTab('ptab-${p}-com','tab-strip-proc-${p}',this);_procRenderTabla('${procId}','com')">
        💱 Com. Errado <span class="cnt">${com}</span>
      </button>` : ''}
    </div>

    <!-- Tab bodies -->
    <div class="tab-body active" id="ptab-${p}-todo"
      style="flex-direction:column;flex:1;min-height:0">
      <div class="tbl-wrap"><table id="tbl-proc-${p}-todo"><thead></thead><tbody></tbody></table></div>
    </div>
    <div class="tab-body" id="ptab-${p}-ok"
      style="flex-direction:column;flex:1;min-height:0">
      <div class="tbl-wrap"><table id="tbl-proc-${p}-ok"><thead></thead><tbody></tbody></table></div>
    </div>
    <div class="tab-body" id="ptab-${p}-sinmatch"
      style="flex-direction:column;flex:1;min-height:0">
      <div class="tbl-wrap"><table id="tbl-proc-${p}-sinmatch"><thead></thead><tbody></tbody></table></div>
    </div>
    <div class="tab-body" id="ptab-${p}-mal"
      style="flex-direction:column;flex:1;min-height:0">
      <div class="tbl-wrap"><table id="tbl-proc-${p}-mal"><thead></thead><tbody></tbody></table></div>
    </div>
    <div class="tab-body" id="ptab-${p}-com"
      style="flex-direction:column;flex:1;min-height:0">
      <div class="tbl-wrap"><table id="tbl-proc-${p}-com"><thead></thead><tbody></tbody></table></div>
    </div>

  </div>`;

  // Renderizar la tabla del tab activo (todo por default)
  _procRenderTabla(procId, 'todo');
}

// ── Wrappers por procesadora ──────────────────────────────────────────
function renderModuloFiserv()  { renderModuloProc('FISERV'); }
function renderModuloGetpos()  { renderModuloProc('GETPOS'); }
function renderModuloPrisma()  { renderModuloProc('PRISMA'); }
