// ═══════════════════════════════════════════════════════════════════
// HISTORICO.JS — Módulo 8: Base acumulativa de períodos conciliados
// Cada "cierre de período" guarda un resumen compacto en IndexedDB
// y exporta un .json de backup para compartir entre equipos.
// ═══════════════════════════════════════════════════════════════════

// ── Construir resumen del período actual ────────────────────────────
function buildPeriodoActual() {
  if (!RESULTADO || !RESULTADO.length) return null;

  const desde = SESSION.periodoDesde || '';
  const hasta = SESSION.periodoHasta || '';
  const id    = `periodo_${desde}_${hasta}_${Date.now()}`;

  // ── Cruce
  const isOK  = r => r.estado?.startsWith('OK') || r.estado?.startsWith('CORREGIDO');
  const total = RESULTADO.length;
  const ok    = RESULTADO.filter(isOK).length;
  const sin   = RESULTADO.filter(r => r.estado === 'SIN MATCH').length;
  const mal   = RESULTADO.filter(r => r.estado?.startsWith('MAL FACTURADO')).length;
  const com   = RESULTADO.filter(r => r.estado === 'COM. ERRADO').length;
  const anu   = RESULTADO.filter(r => r.estado === 'ANULACION SIN COBRO').length;
  const dev   = RESULTADO.filter(r => r.sky.esNeg).length;
  const cor   = Object.keys(CORREGIDAS).length;
  const urg   = RESULTADO.filter(r => r.estado === 'REVISION URGENTE').length;
  const ref   = RESULTADO.filter(r => r.estado === 'REFACTURADO').length;
  const int   = RESULTADO.filter(r => r.sky.integrado).length;

  const sumM  = arr => arr.reduce((s,r) => s + Math.abs(r.sky?.monto||0), 0);
  const mTotal= sumM(RESULTADO.filter(r=>!r.sky.esNeg&&!r.sky.integrado));
  const mOK   = sumM(RESULTADO.filter(isOK));
  const mSin  = sumM(RESULTADO.filter(r=>r.estado==='SIN MATCH'));

  // ── Cobros (si disponible)
  let cobros = null;
  if (typeof COBROS_RESULT !== 'undefined' && COBROS_RESULT.length > 0) {
    const cob = COBROS_RESULT.filter(c => c.estado === 'COBRADO');
    const pen = COBROS_RESULT.filter(c => c.estado === 'PENDIENTE');
    const rec = COBROS_RESULT.filter(c => c.estado === 'RECHAZADO');
    const mCob = cob.reduce((s,c) => s + Math.abs(c.fila?.sky?.monto||0), 0);
    const mPen = pen.reduce((s,c) => s + Math.abs(c.fila?.sky?.monto||0), 0);
    const mRec = rec.reduce((s,c) => s + Math.abs(c.fila?.sky?.monto||0), 0);
    cobros = {
      totalLiq:  COBROS_RESULT.length,
      cobrado:   cob.length,
      pendiente: pen.length,
      rechazado: rec.length,
      montoCobrado:   mCob,
      montoPendiente: mPen,
      montoRechazado: mRec,
      pctCobradoOps:  COBROS_RESULT.length ? +(cob.length/COBROS_RESULT.length*100).toFixed(1) : 0,
      pctCobradoMonto:(mCob+mPen+mRec) ? +(mCob/(mCob+mPen+mRec)*100).toFixed(1) : 0,
    };
  }

  // ── Contracargos (si disponible)
  let contracargos = null;
  if (typeof _CTR_FIS !== 'undefined' || typeof _CTR_GP !== 'undefined') {
    const allCtr = [
      ...(typeof _CTR_FIS !== 'undefined' ? _CTR_FIS : []),
      ...(typeof _CTR_GP  !== 'undefined' ? _CTR_GP  : []),
    ];
    if (allCtr.length > 0) {
      const seg = typeof CTR_SEGUIMIENTO !== 'undefined' ? CTR_SEGUIMIENTO : {};
      const gan = allCtr.filter(c => (seg[c.id]?.estado || 'PENDIENTE') === 'GANADO');
      const per = allCtr.filter(c => (seg[c.id]?.estado || 'PENDIENTE') === 'PERDIDO');
      contracargos = {
        total:       allCtr.length,
        montoRiesgo: allCtr.reduce((s,c) => s+c.importe, 0),
        ganados:     gan.length,
        montoGanado: gan.reduce((s,c) => s + (seg[c.id]?.montoRecuperado || c.importe), 0),
        perdidos:    per.length,
        montoPerdido:per.reduce((s,c) => s+c.importe, 0),
      };
    }
  }

  // ── Top 10 sucursales por SIN MATCH
  const sucMap = {};
  RESULTADO.filter(r=>r.estado==='SIN MATCH').forEach(r => {
    const s = r.sky.suc || '?';
    if (!sucMap[s]) sucMap[s] = { suc:s, ops:0, monto:0 };
    sucMap[s].ops++;
    sucMap[s].monto += Math.abs(r.sky.monto||0);
  });
  const topSucursales = Object.values(sucMap)
    .sort((a,b) => b.monto - a.monto).slice(0, 10);

  // ── Diferencias
  let diferencias = null;
  if (typeof calcularTodasDiferencias === 'function') {
    calcularTodasDiferencias();
    const conDif = RESULTADO.filter(r => r.difMonto != null && Math.abs(r.difMonto) >= 100);
    const dFav   = conDif.filter(r => (r.difMonto||0) < 0).reduce((s,r)=>s+r.difMonto,0);
    const dCont  = conDif.filter(r => (r.difMonto||0) > 0).reduce((s,r)=>s+r.difMonto,0);
    diferencias  = {
      total:    conDif.length,
      totalDif: conDif.reduce((s,r)=>s+(r.difMonto||0),0),
      aFavor:   Math.abs(dFav),
      enContra: dCont,
    };
  }

  // Marcaciones manuales de diferencias de tasas (sobreviven al cerrar periodo)
  let tasasMarcaciones = {};
  try { tasasMarcaciones = JSON.parse(localStorage.getItem('tasas_marc') || '{}'); } catch {}

  // Operaciones sin liquidar al cierre — usadas como arrastre en el periodo siguiente
  const _serNoLiq = (arr, proc) => (arr || []).map(x => ({
    proc,
    fecha:    x.fila?.sky?.fecha    || '',
    suc:      x.fila?.sky?.suc      || '',
    vendedor: x.fila?.sky?.vendedor || '',
    tarjeta:  x.fila?.sky?.tarjeta  || '',
    cuotas:   x.fila?.sky?.cuotas   || 1,
    monto:    x.fila?.sky?.monto    || 0,
    plan:     x.fila?.sky?.plan     || '',
    estado:   x.fila?.estado        || '',
    lote:     x.lote  || x.fila?.sky?.lote  || '',
    cupon:    x.cupon || x.fila?.sky?.cupon || '',
    aut:      x.aut   || x.fila?.proc?.aut  || '',
    equipo:   x.fila?.proc?.equipo  || x.fila?.proc?.pos || '',
  }));
  const pendientesArrastre = [
    ..._serNoLiq(typeof _liqCache !== 'undefined' ? _liqCache.fiserv?.noLiquidadas : [], 'FISERV'),
    ..._serNoLiq(typeof _liqCache !== 'undefined' ? _liqCache.getpos?.noLiquidadas : [], 'GETPOS'),
    ..._serNoLiq(typeof _liqCache !== 'undefined' ? _liqCache.goc?.noLiquidadas    : [], 'GoC'),
  ];

  // Correcciones manuales serializadas con metadata SKY para re-matchear en el próximo período
  const correccionesArrastre = Object.entries(CORREGIDAS).map(([key, cor]) => {
    const fila = RESULTADO.find(r => _skyKey(r.sky) === key);
    return {
      key,
      cor: { ...cor },
      sky: fila ? {
        asiento:  fila.sky.asiento,
        suc:      fila.sky.suc,
        tarjeta:  fila.sky.tarjeta,
        monto:    fila.sky.monto,
        cupon:    fila.sky.cupon,
        plan:     fila.sky.plan,
        cuotas:   fila.sky.cuotas,
        vendedor: fila.sky.vendedor,
      } : null,
    };
  });

  return {
    id,
    nombre:       SESSION.nombre || `Período ${desde} – ${hasta}`,
    periodoDesde: desde,
    periodoHasta: hasta,
    fechaCierre:  new Date().toISOString(),
    usuario:      SESSION.usuario || 'Usuario',
    sesionId:     SESSION.id,

    cruce: { total, ok, sinMatch:sin, malFacturado:mal, comErrado:com,
             anulSC:anu, devoluciones:dev, correcciones:cor,
             revUrgente:urg, refacturado:ref, integradas:int,
             pctOK: total ? +(ok/total*100).toFixed(1) : 0 },

    montos: { total:mTotal, ok:mOK, sinMatch:mSin },
    cobros,
    contracargos,
    diferencias,
    topSucursales,
    tasasMarcaciones,
    pendientesArrastre,
    correccionesArrastre,
  };
}

// ════════════════════════════════════════════════════════════════════
// CERRAR PERÍODO
// ════════════════════════════════════════════════════════════════════
async function cerrarPeriodo() {
  if (!RESULTADO || !RESULTADO.length) {
    alert('No hay datos para cerrar. Ejecutá el cruce primero.');
    return;
  }
  // Si SESSION no tiene fechas, intentar inferirlas desde los lotes o desde RESULTADO
  if (!SESSION.periodoDesde || !SESSION.periodoHasta) {
    // Intentar desde el período activo de la biblioteca
    if (typeof _PERIODO_ACTIVO_ID !== 'undefined' && _PERIODO_ACTIVO_ID &&
        typeof obtenerPeriodoConciliacion === 'function') {
      const _perActivo = await obtenerPeriodoConciliacion(_PERIODO_ACTIVO_ID).catch(() => null);
      const _lotesOk   = (_perActivo?.lotes || []).filter(l => l.estado === 'conciliado');
      if (_lotesOk.length) {
        SESSION.periodoDesde = _lotesOk.map(l => l.fechaDesde).sort()[0];
        SESSION.periodoHasta = _lotesOk.map(l => l.fechaHasta).sort().reverse()[0];
      }
    }
    // Fallback: inferir desde RESULTADO
    if ((!SESSION.periodoDesde || !SESSION.periodoHasta) && RESULTADO.length) {
      const fechas = RESULTADO.map(r => r.sky?.fecha).filter(Boolean).sort();
      if (fechas.length) { SESSION.periodoDesde = fechas[0]; SESSION.periodoHasta = fechas[fechas.length-1]; }
    }
    // Si sigue sin fechas, error
    if (!SESSION.periodoDesde || !SESSION.periodoHasta) {
      alert('Completá las fechas de período (Desde / Hasta) antes de cerrar.');
      return;
    }
  }

  const p = buildPeriodoActual();
  if (!p) { alert('Error al construir el resumen del período.'); return; }

  const msg = `¿Cerrar el período ${p.periodoDesde} → ${p.periodoHasta}?

Resumen:
• ${p.cruce.total.toLocaleString('es-AR')} operaciones · ${p.cruce.pctOK}% OK
• ${p.cruce.sinMatch.toLocaleString('es-AR')} SIN MATCH · ${p.cruce.correcciones} correcciones
${p.cobros ? `• Cobros: ${p.cobros.pctCobradoOps}% cobrado` : ''}
${p.contracargos ? `• Contracargos: ${p.contracargos.total} · $${(p.contracargos.montoRiesgo/1e6).toFixed(1)}M en riesgo` : ''}

Se guardará en el historial y se descargará un backup .json.`;

  if (!confirm(msg)) return;

  // Guardar en IndexedDB
  await guardarPeriodo(p);

  // Guardar arrastre en IDB para auto-carga automática en el período siguiente
  const _arr = {
    id: 'arrastre_activo',
    pendientes:   p.pendientesArrastre   || [],
    correcciones: p.correccionesArrastre || [],
    periodoHasta: p.periodoHasta,
    nombre:       p.nombre,
    ts:           new Date().toISOString(),
  };
  await dbPut('sesiones', _arr);
  _arrastreGuardado = _arr;

  // Exportar backup JSON automáticamente
  exportarPeriodoJSON(p);

  typeof _showToast === 'function'
    ? _showToast(`✓ Período ${p.periodoDesde} cerrado y guardado`)
    : alert(`✓ Período cerrado. Backup exportado.`);

  // Actualizar badge si el módulo está abierto
  const badge = document.getElementById('mcnt-hist');
  if (badge) {
    const todos = await listarPeriodos();
    badge.textContent = todos.length;
  }
}

// ════════════════════════════════════════════════════════════════════
// RENDER MÓDULO HISTÓRICO
// ════════════════════════════════════════════════════════════════════
async function renderModuloHistorico() {
  const panel = document.getElementById('mod-hist');
  if (!panel) return;

  panel.innerHTML = `<div style="display:flex;align-items:center;justify-content:center;
    height:100%;color:var(--m2);font-size:10px">Cargando histórico...</div>`;

  const periodos = await listarPeriodos();

  if (!periodos.length) {
    panel.innerHTML = `
      <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;
        height:100%;gap:14px;color:var(--m2);text-align:center;padding:40px">
        <div style="font-size:42px;opacity:.15">📅</div>
        <div style="font-family:var(--head);font-size:16px;font-weight:700;color:var(--txt);opacity:.4">Sin períodos cerrados aún</div>
        <p style="font-size:10px;max-width:400px;line-height:1.8">
          Completá el cruce de un período, configurá las fechas en el panel izquierdo<br>
          y hacé clic en <b style="color:var(--grn)">Cerrar período ↓</b> para agregar el primero.<br>
          También podés importar períodos de otro equipo.
        </p>
        <label style="background:var(--s3);border:1px solid var(--b2);border-radius:5px;
          color:var(--m1);font-family:var(--mono);font-size:9px;padding:8px 18px;cursor:pointer;
          transition:all .15s" onmouseover="this.style.borderColor='var(--acc)';this.style.color='var(--acc)'"
          onmouseout="this.style.borderColor='var(--b2)';this.style.color='var(--m1)'">
          ⬆ Importar período
          <input type="file" accept=".json" style="display:none" onchange="importarPeriodoUI(this)">
        </label>
      </div>`;
    return;
  }

  // ── KPIs acumulados ─────────────────────────────────────────────
  const totalOps    = periodos.reduce((s,p) => s + p.cruce.total, 0);
  const totalOK     = periodos.reduce((s,p) => s + p.cruce.ok, 0);
  const totalSin    = periodos.reduce((s,p) => s + p.cruce.sinMatch, 0);
  const pctOKAcum   = totalOps ? (totalOK/totalOps*100).toFixed(1) : '—';
  const montoAcum   = periodos.reduce((s,p) => s + (p.montos?.total||0), 0);
  const montoSinAcum= periodos.reduce((s,p) => s + (p.montos?.sinMatch||0), 0);
  const cobPct      = periodos.filter(p=>p.cobros).map(p=>p.cobros.pctCobradoMonto);
  const cobPctProm  = cobPct.length ? (cobPct.reduce((s,v)=>s+v,0)/cobPct.length).toFixed(1)+'%' : '—';

  // ── Tab strip ───────────────────────────────────────────────────
  panel.innerHTML = `
  <div style="display:flex;flex-direction:column;flex:1;min-height:0;overflow:hidden">

    <!-- KPI strip -->
    <div style="display:grid;grid-template-columns:repeat(6,1fr);gap:8px;
      padding:12px 18px;background:var(--s1);border-bottom:1px solid var(--b1);flex-shrink:0">
      <div class="dif-kpi"><div class="dif-kpi-lbl">Períodos cerrados</div>
        <div class="dif-kpi-val cyn">${periodos.length}</div></div>
      <div class="dif-kpi"><div class="dif-kpi-lbl">Total ops acumuladas</div>
        <div class="dif-kpi-val cyn">${totalOps.toLocaleString('es-AR')}</div></div>
      <div class="dif-kpi"><div class="dif-kpi-lbl">% OK promedio</div>
        <div class="dif-kpi-val grn">${pctOKAcum}%</div></div>
      <div class="dif-kpi"><div class="dif-kpi-lbl">Sin Match acumulado</div>
        <div class="dif-kpi-val red">${totalSin.toLocaleString('es-AR')}</div></div>
      <div class="dif-kpi"><div class="dif-kpi-lbl">Monto total conciliado</div>
        <div class="dif-kpi-val" style="color:var(--acc);font-size:14px">${_hFmtM(montoAcum)}</div></div>
      <div class="dif-kpi"><div class="dif-kpi-lbl">% Cobrado promedio</div>
        <div class="dif-kpi-val grn">${cobPctProm}</div></div>
    </div>

    <!-- Tab strip -->
    <div class="tab-strip" id="tab-strip-hist">
      <button class="tb active" onclick="showTab('hist-graficos','tab-strip-hist',this)">📈 Tendencias</button>
      <button class="tb" onclick="showTab('hist-tabla','tab-strip-hist',this)">📋 Detalle por período</button>
    </div>

    <!-- Tab: Gráficos -->
    <div class="tab-body active" id="hist-graficos" style="flex-direction:column;flex:1;min-height:0">
      <div class="dash-scroll">
        <div class="dash-hdr" style="margin-bottom:0">
          <div style="font-size:10px;color:var(--m2)">
            Basado en <b style="color:var(--txt)">${periodos.length}</b> períodos ·
            <b style="color:var(--txt)">${periodos[periodos.length-1]?.periodoDesde}</b> →
            <b style="color:var(--txt)">${periodos[0]?.periodoHasta}</b>
          </div>
          <div style="display:flex;gap:6px">
            <label class="dl-btn" style="background:var(--s3);border:1px solid var(--b2);color:var(--m1);cursor:pointer">
              ⬆ Importar período
              <input type="file" accept=".json" style="display:none" onchange="importarPeriodoUI(this)">
            </label>
            <button class="dl-btn" style="background:#14532d;color:#86efac"
              onclick="exportarHistoricoXlsx()">⬇ Exportar Excel</button>
          </div>
        </div>
        <div class="dash-grid" style="margin-top:14px">

          <!-- Evolución ops -->
          <div class="dash-card">
            <div class="dash-card-title">Operaciones por período <span>(OK vs Sin Match)</span></div>
            <div class="dash-cw h220"><canvas id="ch-hist-ops"></canvas></div>
          </div>

          <!-- % Conciliación -->
          <div class="dash-card">
            <div class="dash-card-title">% Conciliación por período</div>
            <div class="dash-cw h220"><canvas id="ch-hist-pct"></canvas></div>
          </div>

          <!-- Montos -->
          <div class="dash-card dash-card-full">
            <div class="dash-card-title">Monto facturado y Sin Match por período <span>($)</span></div>
            <div class="dash-cw h200"><canvas id="ch-hist-montos"></canvas></div>
          </div>

          <!-- Cobros -->
          <div class="dash-card">
            <div class="dash-card-title">% Cobrado por período</div>
            <div class="dash-cw h220"><canvas id="ch-hist-cobpct"></canvas></div>
          </div>

          <!-- Monto pendiente cobro -->
          <div class="dash-card">
            <div class="dash-card-title">Monto pendiente de cobro por período <span>($)</span></div>
            <div class="dash-cw h220"><canvas id="ch-hist-cobpen"></canvas></div>
          </div>

          <!-- Contracargos -->
          <div class="dash-card dash-card-full">
            <div class="dash-card-title">Contracargos — Monto en riesgo / Ganado / Perdido por período</div>
            <div class="dash-cw h200"><canvas id="ch-hist-ctr"></canvas></div>
          </div>

        </div>
      </div>
    </div>

    <!-- Tab: Tabla detalle -->
    <div class="tab-body" id="hist-tabla" style="flex-direction:column;flex:1;min-height:0">
      <div class="cor-hdr-bar" style="border-left:3px solid var(--acc)">
        <span class="cor-hdr-title">Detalle por período</span>
        <span class="cor-stats">${periodos.length} períodos cerrados</span>
        <div style="margin-left:auto;display:flex;gap:6px">
          <label class="dl-btn" style="background:var(--s3);border:1px solid var(--b2);color:var(--m1);cursor:pointer">
            ⬆ Importar
            <input type="file" accept=".json" style="display:none" onchange="importarPeriodoUI(this)">
          </label>
          <button class="dl-btn" style="background:#14532d;color:#86efac"
            onclick="exportarHistoricoXlsx()">⬇ Excel</button>
          <button class="dl-btn" style="background:#450a0a;color:#fca5a5;border:1px solid rgba(248,113,113,.3)"
            onclick="borrarTodoElHistorial()"
            title="Eliminar todos los períodos del historial">🗑 Vaciar historial</button>
        </div>
      </div>
      <div class="tbl-wrap">
        <table id="tbl-hist"><thead></thead><tbody></tbody></table>
      </div>
    </div>

  </div>`;

  // Renderizar gráficos y tabla
  _renderHistGraficos(periodos);
  _renderHistTabla(periodos);

  // Actualizar badge
  const badge = document.getElementById('mcnt-hist');
  if (badge) badge.textContent = periodos.length;
}

// ── Helper formato compacto ─────────────────────────────────────────
function _hFmtM(v) {
  if (!v) return '$0';
  if (v >= 1e9) return '$' + (v/1e9).toFixed(2) + 'B';
  if (v >= 1e6) return '$' + (v/1e6).toFixed(1) + 'M';
  if (v >= 1e3) return '$' + (v/1e3).toFixed(0) + 'K';
  return '$' + v.toFixed(0);
}

// ── Gráficos de tendencia ───────────────────────────────────────────
function _renderHistGraficos(periodos) {
  const labels   = periodos.map(p => p.periodoDesde?.slice(0,7) || p.nombre).reverse();
  const pRev     = [...periodos].reverse();
  const C = { ok:'#34d399', sin:'#f87171', mal:'#fbbf24',
              acc:'#4f8ef7', cyn:'#38bdf8', vio:'#a78bfa', org:'#fb923c' };
  const g = (id) => document.getElementById(id);
  const gGrid = '#1a2235';
  const gTxt  = '#8ba3c4';
  const scalesBase = {
    x: { ticks:{ color:gTxt, font:{size:9} }, grid:{ color:gGrid } },
    y: { ticks:{ color:gTxt, font:{size:9} }, grid:{ color:gGrid } }
  };

  // 1. Ops OK vs Sin Match
  if (g('ch-hist-ops')) new Chart(g('ch-hist-ops'), { type:'bar', data:{
    labels,
    datasets:[
      { label:'OK',        data:pRev.map(p=>p.cruce.ok),      backgroundColor:C.ok+'99', borderRadius:3, borderWidth:0 },
      { label:'Sin Match', data:pRev.map(p=>p.cruce.sinMatch), backgroundColor:C.sin+'99', borderRadius:3, borderWidth:0 },
      { label:'Correcciones',data:pRev.map(p=>p.cruce.correcciones||0), backgroundColor:C.acc+'99', borderRadius:3, borderWidth:0 },
    ]},
    options:{ responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{ labels:{ color:gTxt, font:{size:9,family:'JetBrains Mono'}, boxWidth:10 } } },
      scales: scalesBase }
  });

  // 2. % Conciliación
  const pctData = pRev.map(p=>p.cruce.pctOK);
  if (g('ch-hist-pct')) new Chart(g('ch-hist-pct'), { type:'line', data:{
    labels,
    datasets:[{ label:'% OK',
      data: pctData,
      borderColor: C.ok, backgroundColor: C.ok+'22',
      borderWidth:2, fill:true, tension:0.3, pointRadius:4,
      pointBackgroundColor: pctData.map(v => v>=90?C.ok:v>=70?C.org:C.sin),
    }]},
    options:{ responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{ display:false },
        tooltip:{ callbacks:{ label: ctx => ` ${ctx.parsed.y}%` } } },
      scales:{ ...scalesBase,
        y:{ ...scalesBase.y, min:0, max:100,
            ticks:{ ...scalesBase.y.ticks, callback:v=>v+'%' } } } }
  });

  // 3. Montos totales vs sin match
  if (g('ch-hist-montos')) new Chart(g('ch-hist-montos'), { type:'bar', data:{
    labels,
    datasets:[
      { label:'Total facturado', data:pRev.map(p=>p.montos?.total||0), backgroundColor:C.acc+'55', borderColor:C.acc, borderWidth:1, borderRadius:3 },
      { label:'Monto OK',        data:pRev.map(p=>p.montos?.ok||0),    backgroundColor:C.ok+'99',  borderRadius:3, borderWidth:0 },
      { label:'Sin Match $',     data:pRev.map(p=>p.montos?.sinMatch||0), backgroundColor:C.sin+'99', borderRadius:3, borderWidth:0 },
    ]},
    options:{ responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{ labels:{ color:gTxt, font:{size:9,family:'JetBrains Mono'}, boxWidth:10 } } },
      scales:{ ...scalesBase, y:{ ...scalesBase.y, ticks:{ ...scalesBase.y.ticks, callback:v=>_hFmtM(v) } } } }
  });

  // 4. % Cobrado
  const cobPcts = pRev.map(p=>p.cobros?.pctCobradoMonto||null);
  if (g('ch-hist-cobpct')) new Chart(g('ch-hist-cobpct'), { type:'line', data:{
    labels,
    datasets:[{ label:'% Cobrado',
      data: cobPcts,
      borderColor:C.cyn, backgroundColor:C.cyn+'22',
      borderWidth:2, fill:true, tension:0.3, pointRadius:4,
      spanGaps:true,
    }]},
    options:{ responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{ display:false } },
      scales:{ ...scalesBase,
        y:{ ...scalesBase.y, min:0, max:100,
            ticks:{ ...scalesBase.y.ticks, callback:v=>v+'%' } } } }
  });

  // 5. Monto pendiente de cobro
  if (g('ch-hist-cobpen')) new Chart(g('ch-hist-cobpen'), { type:'bar', data:{
    labels,
    datasets:[
      { label:'Cobrado',   data:pRev.map(p=>p.cobros?.montoCobrado||0),   backgroundColor:C.ok+'99',  borderRadius:3, borderWidth:0 },
      { label:'Pendiente', data:pRev.map(p=>p.cobros?.montoPendiente||0), backgroundColor:C.sin+'99', borderRadius:3, borderWidth:0 },
    ]},
    options:{ responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{ labels:{ color:gTxt, font:{size:9,family:'JetBrains Mono'}, boxWidth:10 } } },
      scales:{ ...scalesBase, y:{ ...scalesBase.y, ticks:{ ...scalesBase.y.ticks, callback:v=>_hFmtM(v) } } } }
  });

  // 6. Contracargos
  if (g('ch-hist-ctr')) new Chart(g('ch-hist-ctr'), { type:'bar', data:{
    labels,
    datasets:[
      { label:'En riesgo', data:pRev.map(p=>p.contracargos?.montoRiesgo||0), backgroundColor:C.org+'99', borderRadius:3, borderWidth:0 },
      { label:'Ganado',    data:pRev.map(p=>p.contracargos?.montoGanado||0),  backgroundColor:C.ok+'99',  borderRadius:3, borderWidth:0 },
      { label:'Perdido',   data:pRev.map(p=>p.contracargos?.montoPerdido||0), backgroundColor:C.sin+'99', borderRadius:3, borderWidth:0 },
    ]},
    options:{ responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{ labels:{ color:gTxt, font:{size:9,family:'JetBrains Mono'}, boxWidth:10 } } },
      scales:{ ...scalesBase, y:{ ...scalesBase.y, ticks:{ ...scalesBase.y.ticks, callback:v=>_hFmtM(v) } } } }
  });
}

// ── Tabla detalle por período ────────────────────────────────────────
function _renderHistTabla(periodos) {
  const tbl = document.getElementById('tbl-hist'); if (!tbl) return;
  const HDR = ['Período','Desde','Hasta','Usuario','Total ops','OK','SIN MATCH',
               '% OK','Mal Fact.','Com. Errado','Correc.','Rev. Urg.',
               'Monto total','Monto OK','Monto sin match',
               'Liq. cobradas','% Cobrado','Monto cobrado','Monto pendiente',
               'Contracargos','Monto CTR riesgo','CTR ganados','CTR perdidos','Cierre',''];
  tbl.querySelector('thead').innerHTML = `<tr>${HDR.map(h=>`<th>${h}</th>`).join('')}</tr>`;
  tbl.querySelector('tbody').innerHTML = periodos.map(p => {
    const c = p.cruce, m = p.montos||{}, co = p.cobros, ct = p.contracargos;
    return `<tr>
      <td style="font-weight:600;color:var(--txt)">${p.nombre}</td>
      <td>${p.periodoDesde}</td>
      <td>${p.periodoHasta}</td>
      <td style="font-size:9px">${p.usuario}</td>
      <td class="num">${c.total.toLocaleString('es-AR')}</td>
      <td class="num" style="color:var(--grn)">${c.ok.toLocaleString('es-AR')}</td>
      <td class="num" style="color:var(--red)">${c.sinMatch.toLocaleString('es-AR')}</td>
      <td class="num" style="color:${c.pctOK>=90?'var(--grn)':c.pctOK>=70?'var(--yel)':'var(--red)'}">
        <b>${c.pctOK}%</b></td>
      <td class="num">${c.malFacturado}</td>
      <td class="num">${c.comErrado}</td>
      <td class="num">${c.correcciones}</td>
      <td class="num" style="color:${c.revUrgente>0?'var(--org)':'inherit'}">${c.revUrgente}</td>
      <td class="num" style="font-size:9px">${_hFmtM(m.total||0)}</td>
      <td class="num" style="font-size:9px;color:var(--grn)">${_hFmtM(m.ok||0)}</td>
      <td class="num" style="font-size:9px;color:var(--red)">${_hFmtM(m.sinMatch||0)}</td>
      <td class="num">${co ? co.cobrado.toLocaleString('es-AR')+' / '+co.totalLiq.toLocaleString('es-AR') : '—'}</td>
      <td class="num" style="color:${co?(co.pctCobradoMonto>=85?'var(--grn)':co.pctCobradoMonto>=60?'var(--yel)':'var(--red)'):'inherit'}">
        ${co ? co.pctCobradoMonto+'%' : '—'}</td>
      <td class="num" style="font-size:9px;color:var(--grn)">${co ? _hFmtM(co.montoCobrado) : '—'}</td>
      <td class="num" style="font-size:9px;color:var(--red)">${co ? _hFmtM(co.montoPendiente) : '—'}</td>
      <td class="num">${ct ? ct.total : '—'}</td>
      <td class="num" style="font-size:9px;color:var(--org)">${ct ? _hFmtM(ct.montoRiesgo) : '—'}</td>
      <td class="num" style="color:var(--grn)">${ct ? ct.ganados : '—'}</td>
      <td class="num" style="color:var(--red)">${ct ? ct.perdidos : '—'}</td>
      <td style="font-size:9px;color:var(--m2)">${p.fechaCierre?.slice(0,10)}</td>
      <td style="white-space:nowrap;display:flex;gap:3px;align-items:center">
        <button onclick="descargarPeriodo('${p.id}')" title="Descargar backup JSON"
          style="background:none;border:1px solid var(--b2);border-radius:3px;
            color:var(--m2);font-size:9px;cursor:pointer;padding:2px 7px"
          onmouseover="this.style.borderColor='var(--acc)';this.style.color='var(--acc)'"
          onmouseout="this.style.borderColor='var(--b2)';this.style.color='var(--m2)'">⬇</button>
        <button onclick="cargarCorreccionesDesde('${p.id}')"
          title="${(p.correccionesArrastre?.length||0)} corrección(es) guardadas — clic para importar al cruce actual"
          style="background:none;border:1px solid var(--b2);border-radius:3px;
            color:${p.correccionesArrastre?.length ? 'var(--acc)' : 'var(--m2)'};
            font-size:9px;cursor:pointer;padding:2px 7px;opacity:${p.correccionesArrastre?.length?1:.4}"
          onmouseover="this.style.borderColor='var(--acc)';this.style.color='var(--acc)';this.style.opacity=1"
          onmouseout="this.style.borderColor='var(--b2)';this.style.opacity='${p.correccionesArrastre?.length?1:.4}'">
          📋 ${p.correccionesArrastre?.length||0}</button>
        <button onclick="borrarPeriodo('${p.id}')" title="Eliminar período del historial"
          style="background:none;border:1px solid var(--b2);border-radius:3px;
            color:var(--m2);font-size:9px;cursor:pointer;padding:2px 7px"
          onmouseover="this.style.borderColor='var(--red)';this.style.color='var(--red)'"
          onmouseout="this.style.borderColor='var(--b2)';this.style.color='var(--m2)'">×</button>
      </td>
    </tr>`;
  }).join('');
}

// ════════════════════════════════════════════════════════════════════
// ARRASTRE DE CORRECCIONES MANUALES ENTRE PERÍODOS
// ════════════════════════════════════════════════════════════════════

function _importarCorreccionesArrastre(entries) {
  if (!entries?.length) {
    if (typeof _showToast === 'function') _showToast('El período no tiene correcciones guardadas');
    return;
  }
  if (!RESULTADO.length) {
    alert('Primero ejecutá el cruce del período actual para poder importar correcciones.');
    return;
  }

  let exactas = 0, rematch = 0, omitidas = 0;

  for (const { key, cor, sky } of entries) {
    // 1. Match exacto — mismo asiento número en el nuevo período
    if (RESULTADO.some(r => _skyKey(r.sky) === key)) {
      if (!CORREGIDAS[key]) { CORREGIDAS[key] = { ...cor, _arrastre: true }; exactas++; }
      continue;
    }

    // 2. Fuzzy — buscar fila con misma suc + tarjeta + cuotas + monto (±2%)
    if (!sky) { omitidas++; continue; }
    const tol = Math.abs(sky.monto || 0) * 0.02;
    const candidatos = RESULTADO.filter(r => {
      const s = r.sky;
      return !CORREGIDAS[_skyKey(s)] &&
        s.suc     === sky.suc &&
        s.tarjeta === sky.tarjeta &&
        s.cuotas  == sky.cuotas &&
        Math.abs(Math.abs(s.monto) - Math.abs(sky.monto)) <= (tol || 1);
    });

    if (candidatos.length === 1) {
      const newKey = _skyKey(candidatos[0].sky);
      CORREGIDAS[newKey] = { ...cor, _arrastre: true, _origKey: key };
      rematch++;
    } else {
      omitidas++;
    }
  }

  if (typeof aplicarCorreccionesManuales === 'function') aplicarCorreccionesManuales();
  scheduleAutoSave();

  const total = exactas + rematch;
  const msg = `✓ ${total} corrección${total !== 1 ? 'es' : ''} importada${total !== 1 ? 's' : ''}` +
    (rematch  ? ` (${rematch} re-matcheadas por suc+tarjeta+monto)` : '') +
    (omitidas ? ` · ${omitidas} sin coincidencia` : '');
  if (typeof _showToast === 'function') _showToast(msg);
  else alert(msg);

  if (typeof renderRevision === 'function') renderRevision();
  if (typeof renderModuloCobros === 'function') renderModuloCobros();
}

async function cargarCorreccionesDesde(periodoId) {
  const todos = await listarPeriodos();
  const p = todos.find(x => x.id === periodoId);
  if (!p) return;
  if (!p.correccionesArrastre?.length) {
    if (typeof _showToast === 'function')
      _showToast('Este período fue cerrado sin correcciones o con versión anterior de la app');
    return;
  }
  const n = p.correccionesArrastre.length;
  if (!confirm(`Importar ${n} corrección${n!==1?'es':''} del período ${p.periodoDesde} → ${p.periodoHasta}?\n\nSe agregarán al cruce actual. Las correcciones que no encuentren una operación idéntica se intentan reasignar por suc + tarjeta + monto.`)) return;
  _importarCorreccionesArrastre(p.correccionesArrastre);
}

function cargarCorreccionesDesdeJSON(input) {
  const file = input.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const data = JSON.parse(e.target.result);
      if (!data.correccionesArrastre?.length) {
        alert('El archivo no contiene correcciones exportadas.\n(Debe ser un backup generado con esta versión de la app)');
        return;
      }
      const n = data.correccionesArrastre.length;
      if (!confirm(`Importar ${n} corrección${n!==1?'es':''} del período ${data.periodoDesde||'?'} → ${data.periodoHasta||'?'}?`)) return;
      _importarCorreccionesArrastre(data.correccionesArrastre);
    } catch(err) { alert('Error al leer el archivo: ' + err.message); }
  };
  reader.readAsText(file, 'utf-8');
}

// ── Acciones desde tabla ─────────────────────────────────────────────
async function descargarPeriodo(id) {
  const todos = await listarPeriodos();
  const p = todos.find(x => x.id === id);
  if (p) exportarPeriodoJSON(p);
}

async function borrarPeriodo(id) {
  if (!confirm('¿Eliminar este período del historial? El backup JSON no se borra.')) return;
  await eliminarPeriodo(id);
  renderModuloHistorico();
}

async function borrarTodoElHistorial() {
  const todos = await listarPeriodos();
  if (!todos.length) { alert('El historial ya está vacío.'); return; }
  if (!confirm(`¿Eliminar los ${todos.length} períodos del historial?\n\nEsta acción no se puede deshacer. Los backups .json descargados no se borran.`)) return;
  for (const p of todos) await eliminarPeriodo(p.id);
  const badge = document.getElementById('mcnt-hist');
  if (badge) badge.textContent = '—';
  typeof _showToast === 'function'
    ? _showToast('🗑 Historial vaciado')
    : alert('Historial vaciado.');
  renderModuloHistorico();
}

// ── Importar período desde JSON ──────────────────────────────────────
async function importarPeriodoUI(input) {
  const file = input.files[0]; if (!file) return;
  try {
    const p = await importarPeriodoJSON(file);
    typeof _showToast === 'function'
      ? _showToast(`✓ Período ${p.periodoDesde} importado`)
      : alert(`✓ Período importado: ${p.nombre}`);
    renderModuloHistorico();
  } catch(e) {
    alert('Error al importar: ' + e.message);
  }
  input.value = '';
}

// ── Exportar histórico a Excel ───────────────────────────────────────
async function exportarHistoricoXlsx() {
  const periodos = await listarPeriodos();
  if (!periodos.length) { alert('No hay períodos en el historial.'); return; }

  const HDR = ['Período','Desde','Hasta','Usuario',
    'Total ops','OK','SIN MATCH','% OK','Mal Fact.','Com. Errado','Correcciones','Rev. Urgente',
    'Monto total','Monto OK','Monto sin match',
    'Liq. totales','Cobradas','Pendientes','% Cobrado ops','% Cobrado monto',
    'Monto cobrado','Monto pendiente',
    'Contracargos total','Monto CTR riesgo','CTR ganados','CTR perdidos',
    'Fecha cierre'];

  const data = periodos.map(p => {
    const c = p.cruce, m = p.montos||{}, co = p.cobros, ct = p.contracargos;
    return [
      p.nombre, p.periodoDesde, p.periodoHasta, p.usuario,
      c.total, c.ok, c.sinMatch, c.pctOK,
      c.malFacturado, c.comErrado, c.correcciones, c.revUrgente,
      m.total||0, m.ok||0, m.sinMatch||0,
      co?.totalLiq||0, co?.cobrado||0, co?.pendiente||0,
      co?.pctCobradoOps||0, co?.pctCobradoMonto||0,
      co?.montoCobrado||0, co?.montoPendiente||0,
      ct?.total||0, ct?.montoRiesgo||0, ct?.ganados||0, ct?.perdidos||0,
      p.fechaCierre?.slice(0,10)||'',
    ];
  });

  _exportXlsx([HDR, ...data], 'Histórico', `Historico_Conciliacion_${hoy()}.xlsx`);
}
