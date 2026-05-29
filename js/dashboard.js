// ═══════════════════════════════════════════════════════════════════
// DASHBOARD.JS — Gráficos y resúmenes de conciliación
// ═══════════════════════════════════════════════════════════════════

// ── Paleta alineada con el tema oscuro ──────────────────────────────
const DASH_CLR = {
  ok:      '#34d399',
  equiv:   '#2dd4bf',
  sm:      '#f87171',
  malFact: '#fbbf24',
  comErr:  '#fb923c',
  anulSC:  '#a78bfa',
  corr:    '#4f8ef7',
  refact:  '#6ee7b7',
  urg:     '#e879f9',
  dev:     '#38bdf8',
  grid:    '#1a2235',
  muted:   '#566882',
  txt:     '#8ba3c4',
};

const TARJ_PALETTE = [
  '#4f8ef7','#34d399','#fbbf24','#f87171','#a78bfa',
  '#38bdf8','#fb923c','#6ee7b7','#f472b6','#94a3b8',
  '#2dd4bf','#facc15','#818cf8','#f9a8d4','#67e8f9',
];

let _dashCharts = {};

// ── Helpers ─────────────────────────────────────────────────────────
function _dDestroyAll() {
  Object.values(_dashCharts).forEach(c => { try { c.destroy(); } catch(e){} });
  _dashCharts = {};
}

function _dEstColor(est) {
  if (!est) return DASH_CLR.muted;
  if (est.startsWith('OK (equiv'))       return DASH_CLR.equiv;
  if (est.startsWith('OK (int'))         return '#5eead4';
  if (est.startsWith('OK'))              return DASH_CLR.ok;
  if (est === 'SIN MATCH')               return DASH_CLR.sm;
  if (est.startsWith('MAL FACTURADO'))   return DASH_CLR.malFact;
  if (est === 'COM. ERRADO')             return DASH_CLR.comErr;
  if (est === 'ANULACION SIN COBRO')     return DASH_CLR.anulSC;
  if (est.startsWith('CORREGIDO'))       return DASH_CLR.corr;
  if (est === 'REVISION URGENTE')        return DASH_CLR.urg;
  if (est === 'REFACTURADO')             return DASH_CLR.refact;
  return DASH_CLR.muted;
}

function _dFmtM(v) {
  if (v >= 1e9) return '$' + (v/1e9).toFixed(2) + 'B';
  if (v >= 1e6) return '$' + (v/1e6).toFixed(2) + 'M';
  if (v >= 1e3) return '$' + (v/1e3).toFixed(1) + 'K';
  return '$' + v.toFixed(0);
}

function _dFmtPeso(v) {
  return '$ ' + v.toLocaleString('es-AR', {minimumFractionDigits:2, maximumFractionDigits:2});
}

function _dNormEst(est) {
  if (!est) return 'SIN ESTADO';
  if (est.startsWith('CORREGIDO')) return 'CORREGIDO MANUAL';
  if (est.startsWith('MAL FACTURADO')) return 'MAL FACTURADO';
  return est;
}

// ── Chart defaults factory ──────────────────────────────────────────
function _dScales(extra = {}) {
  return {
    x: { ticks: { color:DASH_CLR.txt, font:{family:'JetBrains Mono',size:9} }, grid: { color:DASH_CLR.grid }, ...extra.x },
    y: { ticks: { color:DASH_CLR.txt, font:{family:'JetBrains Mono',size:9} }, grid: { color:DASH_CLR.grid }, ...extra.y },
  };
}

function _dLegend(position='top') {
  return { position, labels: { color:DASH_CLR.txt, font:{family:'JetBrains Mono',size:10}, boxWidth:10, padding:8 } };
}

// ── Pestaña interna del dashboard ─────────────────────────────────
function showDashTab(tab, btn) {
  document.querySelectorAll('.dash-itab').forEach(b => b.classList.remove('active'));
  if (btn) btn.classList.add('active');
  const tx  = document.getElementById('dash-tab-tx');
  const pag = document.getElementById('dash-tab-pag');
  if (tx)  tx.style.display  = tab === 'tx'  ? '' : 'none';
  if (pag) pag.style.display = tab === 'pag' ? '' : 'none';
  setTimeout(() => Object.values(_dashCharts).forEach(c => { try{c.resize();}catch(e){} }), 60);
}

// ════════════════════════════════════════════════════════════════════
// RENDER PRINCIPAL
// ════════════════════════════════════════════════════════════════════
function renderDashboard() {
  const elEmpty   = document.getElementById('dash-empty');
  const elOffline = document.getElementById('dash-offline');
  const elContent = document.getElementById('dash-content');

  function _show(which) {
    [elEmpty, elOffline, elContent].forEach(el => { if(el) el.style.display='none'; });
    if(which) which.style.display = 'flex';
  }

  if (typeof Chart === 'undefined') { _show(elOffline); return; }
  if (!RESULTADO.length)             { _show(elEmpty);   return; }

  _show(elContent);
  _dDestroyAll();

  const rows  = RESULTADO;
  const total = rows.length;

  const isOK  = r => r.estado?.startsWith('OK') || r.estado?.startsWith('CORREGIDO');
  const isSM  = r => r.estado === 'SIN MATCH';
  const okRows = rows.filter(isOK);
  const smRows = rows.filter(isSM);
  const pctOK  = total ? ((okRows.length / total) * 100).toFixed(1) : '0.0';
  const montoOK  = okRows.reduce((s,r) => s + Math.abs(r.sky?.monto||0), 0);
  const montoSM  = smRows.reduce((s,r) => s + Math.abs(r.sky?.monto||0), 0);

  // ── KPIs ──────────────────────────────────────────────────────────
  function _kpi(id, v) { const e=document.getElementById(id); if(e) e.textContent=v; }
  _kpi('kpi-total', total.toLocaleString('es-AR'));
  _kpi('kpi-ok',    okRows.length.toLocaleString('es-AR'));
  _kpi('kpi-sm',    smRows.length.toLocaleString('es-AR'));
  _kpi('kpi-pct',   pctOK + '%');
  _kpi('kpi-mok',   _dFmtM(montoOK));
  _kpi('kpi-msm',   _dFmtM(montoSM));
  _kpi('dash-total-lbl', total.toLocaleString('es-AR'));

  // Color dinámico del KPI de % conciliación
  const kpiPctEl = document.getElementById('kpi-card-pct');
  if (kpiPctEl) {
    const p = parseFloat(pctOK);
    kpiPctEl.style.borderColor = p >= 90 ? 'rgba(52,211,153,.35)' : p >= 70 ? 'rgba(251,191,36,.35)' : 'rgba(248,113,113,.35)';
  }

  // ── 1. DONUT — Distribución de estados ────────────────────────────
  const estAgg = {};
  rows.forEach(r => { const e=_dNormEst(r.estado); estAgg[e]=(estAgg[e]||0)+1; });
  const estPairs = Object.entries(estAgg).sort((a,b)=>b[1]-a[1]);

  _dashCharts.estados = new Chart(document.getElementById('ch-estados'), {
    type: 'doughnut',
    data: {
      labels: estPairs.map(([k])=>k),
      datasets: [{
        data: estPairs.map(([,v])=>v),
        backgroundColor: estPairs.map(([k])=>_dEstColor(k)),
        borderWidth: 1, borderColor: '#111827', hoverOffset: 6,
      }]
    },
    options: {
      responsive:true, maintainAspectRatio:false, cutout:'62%',
      plugins: {
        legend: _dLegend('right'),
        tooltip: { callbacks: { label: ctx => ` ${ctx.label}: ${ctx.parsed.toLocaleString('es-AR')} (${((ctx.parsed/total)*100).toFixed(1)}%)` } }
      }
    }
  });

  // ── 2. BAR — Monto total por estado ───────────────────────────────
  const montoAgg = {};
  rows.forEach(r => { const e=_dNormEst(r.estado); montoAgg[e]=(montoAgg[e]||0)+Math.abs(r.sky?.monto||0); });
  const montoPairs = Object.entries(montoAgg).sort((a,b)=>b[1]-a[1]);

  _dashCharts.montoEst = new Chart(document.getElementById('ch-monto-estado'), {
    type: 'bar',
    data: {
      labels: montoPairs.map(([k])=>k),
      datasets: [{ label:'Monto ($)', data:montoPairs.map(([,v])=>v), backgroundColor:montoPairs.map(([k])=>_dEstColor(k)+'cc'), borderRadius:4, borderWidth:0 }]
    },
    options: {
      responsive:true, maintainAspectRatio:false,
      plugins: { legend:{display:false}, tooltip:{ callbacks:{ label: ctx => ' '+_dFmtPeso(ctx.parsed.y) } } },
      scales: {
        x: { ticks:{ color:DASH_CLR.txt, font:{size:9,family:'JetBrains Mono'}, maxRotation:30 }, grid:{ color:DASH_CLR.grid } },
        y: { ticks:{ color:DASH_CLR.txt, font:{size:9}, callback: v => _dFmtM(v) }, grid:{ color:DASH_CLR.grid } }
      }
    }
  });

  // ── 3. GROUPED BAR — Resultado por procesadora ────────────────────
  const procMap = {
    FISERV: { ok:0, sm:0, otro:0 },
    GETPOS:  { ok:0, sm:0, otro:0 },
    OTRO:   { ok:0, sm:0, otro:0 },
  };
  rows.forEach(r => {
    const p = r.procEncontrada==='FISERV' ? 'FISERV' : r.procEncontrada==='GETPOS' ? 'GETPOS' : 'OTRO';
    const b = isOK(r) ? 'ok' : isSM(r) ? 'sm' : 'otro';
    procMap[p][b]++;
  });

  _dashCharts.proc = new Chart(document.getElementById('ch-proc'), {
    type:'bar',
    data: {
      labels: ['FISERV', 'GETPOS', 'Sin asignar'],
      datasets: [
        { label:'OK / Corregido', data:[procMap.FISERV.ok,  procMap.GETPOS.ok,  procMap.OTRO.ok],  backgroundColor:DASH_CLR.ok+'cc',      borderRadius:3, borderWidth:0 },
        { label:'Sin Match',      data:[procMap.FISERV.sm,  procMap.GETPOS.sm,  procMap.OTRO.sm],  backgroundColor:DASH_CLR.sm+'cc',      borderRadius:3, borderWidth:0 },
        { label:'Otros estados',  data:[procMap.FISERV.otro,procMap.GETPOS.otro,procMap.OTRO.otro], backgroundColor:DASH_CLR.malFact+'cc', borderRadius:3, borderWidth:0 },
      ]
    },
    options: {
      responsive:true, maintainAspectRatio:false,
      plugins: { legend:{ labels:{ color:DASH_CLR.txt, font:{size:9,family:'JetBrains Mono'}, boxWidth:10 } } },
      scales: _dScales()
    }
  });

  // ── 4. HORIZONTAL BAR — Top 15 sucursales por Sin Match ───────────
  const sucSM = {};
  smRows.forEach(r => { const s=r.sky?.suc??'?'; sucSM[s]=(sucSM[s]||0)+1; });
  const sucSMPairs = Object.entries(sucSM).sort((a,b)=>b[1]-a[1]).slice(0,15);

  _dashCharts.sucSM = new Chart(document.getElementById('ch-suc-sm'), {
    type:'bar',
    data: {
      labels: sucSMPairs.map(([k])=>'Suc. '+k),
      datasets: [{ label:'Sin Match', data:sucSMPairs.map(([,v])=>v), backgroundColor:DASH_CLR.sm+'bb', borderRadius:3, borderWidth:0 }]
    },
    options: {
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins: { legend:{display:false} },
      scales: {
        x: { ticks:{ color:DASH_CLR.txt, font:{size:9} }, grid:{ color:DASH_CLR.grid } },
        y: { ticks:{ color:DASH_CLR.txt, font:{size:9,family:'JetBrains Mono'} }, grid:{ display:false } }
      }
    }
  });

  // ── 5. LINE — Evolución diaria ─────────────────────────────────────
  const byFecha = {};
  rows.forEach(r => {
    const f = r.sky?.fecha ?? '?';
    if (!byFecha[f]) byFecha[f] = { ok:0, sm:0, otros:0 };
    if (isOK(r)) byFecha[f].ok++;
    else if (isSM(r)) byFecha[f].sm++;
    else byFecha[f].otros++;
  });
  const fechas = Object.keys(byFecha).sort();

  _dashCharts.evol = new Chart(document.getElementById('ch-evolucion'), {
    type:'line',
    data: {
      labels: fechas,
      datasets: [
        { label:'OK conciliadas', data:fechas.map(f=>byFecha[f].ok),    borderColor:DASH_CLR.ok,      backgroundColor:DASH_CLR.ok+'18',      fill:true, tension:.35, pointRadius:3, borderWidth:2 },
        { label:'Sin Match',      data:fechas.map(f=>byFecha[f].sm),    borderColor:DASH_CLR.sm,      backgroundColor:DASH_CLR.sm+'18',      fill:true, tension:.35, pointRadius:3, borderWidth:2 },
        { label:'Otros estados',  data:fechas.map(f=>byFecha[f].otros), borderColor:DASH_CLR.malFact, backgroundColor:DASH_CLR.malFact+'18', fill:true, tension:.35, pointRadius:3, borderWidth:2 },
      ]
    },
    options: {
      responsive:true, maintainAspectRatio:false,
      plugins: { legend:{ labels:{ color:DASH_CLR.txt, font:{size:9,family:'JetBrains Mono'}, boxWidth:10 } } },
      scales: {
        x: { ticks:{ color:DASH_CLR.txt, font:{size:9}, maxRotation:40 }, grid:{ color:DASH_CLR.grid } },
        y: { ticks:{ color:DASH_CLR.txt, font:{size:9} }, grid:{ color:DASH_CLR.grid }, beginAtZero:true }
      }
    }
  });

  // ── 6. DONUT — Por tipo de tarjeta ─────────────────────────────────
  const tarjAgg = {};
  rows.forEach(r => { const t=r.sky?.tarjeta||'Sin dato'; tarjAgg[t]=(tarjAgg[t]||0)+1; });
  const tarjPairs = Object.entries(tarjAgg).sort((a,b)=>b[1]-a[1]);

  _dashCharts.tarjeta = new Chart(document.getElementById('ch-tarjeta'), {
    type:'doughnut',
    data: {
      labels: tarjPairs.map(([k])=>k),
      datasets: [{ data:tarjPairs.map(([,v])=>v), backgroundColor:TARJ_PALETTE.slice(0,tarjPairs.length), borderWidth:1, borderColor:'#111827', hoverOffset:5 }]
    },
    options: {
      responsive:true, maintainAspectRatio:false, cutout:'55%',
      plugins: {
        legend: _dLegend('right'),
        tooltip: { callbacks: { label: ctx => ` ${ctx.label}: ${ctx.parsed.toLocaleString('es-AR')} (${((ctx.parsed/total)*100).toFixed(1)}%)` } }
      }
    }
  });

  // ── 7. HORIZONTAL BAR — Diferencia $ acumulada por sucursal ───────
  const sucDifMap = {};
  rows.forEach(r => {
    const procM = r.procMontoNorm ?? r.proc?.monto ?? null;
    if (procM === null) return;
    const dif = Math.abs((r.sky?.monto||0) - procM);
    if (dif < 0.5) return;   // ignorar diferencias de centavos
    const suc = r.sky?.suc ?? '?';
    sucDifMap[suc] = (sucDifMap[suc]||0) + dif;
  });
  const sucDifPairs = Object.entries(sucDifMap).sort((a,b)=>b[1]-a[1]).slice(0,15);

  _dashCharts.sucDif = new Chart(document.getElementById('ch-suc-dif'), {
    type:'bar',
    data: {
      labels: sucDifPairs.length ? sucDifPairs.map(([k])=>'Suc. '+k) : ['Sin diferencias'],
      datasets: [{ label:'Diferencia $', data:sucDifPairs.length ? sucDifPairs.map(([,v])=>v) : [0], backgroundColor:DASH_CLR.malFact+'bb', borderRadius:3, borderWidth:0 }]
    },
    options: {
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins: { legend:{display:false}, tooltip:{ callbacks:{ label: ctx => ' '+_dFmtPeso(ctx.parsed.x) } } },
      scales: {
        x: { ticks:{ color:DASH_CLR.txt, font:{size:9}, callback: v => _dFmtM(v) }, grid:{ color:DASH_CLR.grid } },
        y: { ticks:{ color:DASH_CLR.txt, font:{size:9,family:'JetBrains Mono'} }, grid:{ display:false } }
      }
    }
  });

  // ── 8. STACKED BAR — Composición OK / SM / Otros por sucursal ─────
  const sucAll = {};
  rows.forEach(r => {
    const suc = r.sky?.suc ?? '?';
    if (!sucAll[suc]) sucAll[suc] = { ok:0, sm:0, otro:0 };
    if (isOK(r)) sucAll[suc].ok++;
    else if (isSM(r)) sucAll[suc].sm++;
    else sucAll[suc].otro++;
  });
  const sucAllPairs = Object.entries(sucAll)
    .sort((a,b) => (b[1].ok+b[1].sm+b[1].otro) - (a[1].ok+a[1].sm+a[1].otro))
    .slice(0, 25);

  _dashCharts.sucComp  = new Chart(document.getElementById('ch-suc-comp'), {
    type:'bar',
    data: {
      labels: sucAllPairs.map(([k])=>'Suc. '+k),
      datasets: [
        { label:'OK',       data:sucAllPairs.map(([,v])=>v.ok),   backgroundColor:DASH_CLR.ok+'cc',      borderWidth:0 },
        { label:'Sin Match',data:sucAllPairs.map(([,v])=>v.sm),   backgroundColor:DASH_CLR.sm+'cc',      borderWidth:0 },
        { label:'Otros',    data:sucAllPairs.map(([,v])=>v.otro), backgroundColor:DASH_CLR.malFact+'aa', borderWidth:0 },
      ]
    },
    options: {
      responsive:true, maintainAspectRatio:false,
      plugins: { legend:{ labels:{ color:DASH_CLR.txt, font:{size:9,family:'JetBrains Mono'}, boxWidth:10 } } },
      scales: {
        x: { stacked:true, ticks:{ color:DASH_CLR.txt, font:{size:9}, maxRotation:45 }, grid:{ color:DASH_CLR.grid } },
        y: { stacked:true, ticks:{ color:DASH_CLR.txt, font:{size:9} }, grid:{ color:DASH_CLR.grid } }
      }
    }
  });

  // ── Tarjeta: Resumen de impacto económico ─────────────────────────
  const reEl = document.getElementById('dash-resumen-economico');
  if (reEl) {
    const montoTot  = rows.reduce((s,r)=>s+Math.abs(r.sky?.monto||0),0);
    const montoOtros = rows.filter(r=>!isOK(r)&&!isSM(r)).reduce((s,r)=>s+Math.abs(r.sky?.monto||0),0);
    const difTotal  = rows.reduce((s,r)=>{
      const p=r.procMontoNorm??r.proc?.monto??null;
      return p!==null ? s+Math.abs((r.sky?.monto||0)-p) : s;
    },0);
    const pctMonto  = montoTot ? ((montoOK/montoTot)*100).toFixed(1) : '0.0';
    const line = (lbl,val,clr='var(--txt)') =>
      `<div style="display:flex;justify-content:space-between;align-items:center;
        padding:6px 10px;background:var(--s3);border-radius:var(--r4)">
        <span style="color:var(--m2)">${lbl}</span>
        <span style="font-weight:700;color:${clr}">${val}</span>
      </div>`;
    reEl.innerHTML =
      line('Monto total SKY',      _dFmtPeso(montoTot)) +
      line('Monto OK conciliado',  _dFmtPeso(montoOK),  DASH_CLR.ok) +
      line('Monto Sin Match',      _dFmtPeso(montoSM),  DASH_CLR.sm) +
      line('Monto en otros estados', _dFmtPeso(montoOtros), DASH_CLR.malFact) +
      line('Dif. total vs Procesadora', _dFmtPeso(difTotal), DASH_CLR.urg) +
      line('% Monto conciliado',   pctMonto+'%',
        parseFloat(pctMonto)>=90?DASH_CLR.ok:parseFloat(pctMonto)>=70?DASH_CLR.malFact:DASH_CLR.sm);
  }

  // Renderizar tab pagos
  _renderDashPagos();

  // ── 9. BAR — Monto sin match por sucursal (top 15) ─────────────────
  const sucSMmonto = {};
  smRows.forEach(r => { const s=r.sky?.suc??'?'; sucSMmonto[s]=(sucSMmonto[s]||0)+Math.abs(r.sky?.monto||0); });
  const sucSMMontoPairs = Object.entries(sucSMmonto).sort((a,b)=>b[1]-a[1]).slice(0,15);

  _dashCharts.sucSMmonto = new Chart(document.getElementById('ch-suc-sm-monto'), {
    type:'bar',
    data: {
      labels: sucSMMontoPairs.map(([k])=>'Suc. '+k),
      datasets: [{ label:'Monto Sin Match ($)', data:sucSMMontoPairs.map(([,v])=>v), backgroundColor:DASH_CLR.sm+'99', borderRadius:3, borderWidth:0 }]
    },
    options: {
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins: { legend:{display:false}, tooltip:{ callbacks:{ label: ctx => ' '+_dFmtPeso(ctx.parsed.x) } } },
      scales: {
        x: { ticks:{ color:DASH_CLR.txt, font:{size:9}, callback: v=>_dFmtM(v) }, grid:{ color:DASH_CLR.grid } },
        y: { ticks:{ color:DASH_CLR.txt, font:{size:9,family:'JetBrains Mono'} }, grid:{ display:false } }
      }
    }
  });
}

// ════════════════════════════════════════════════════════════════════
// RENDER PAGOS — Gráficos del módulo 4 · COBROS en el dashboard
// ════════════════════════════════════════════════════════════════════
function _renderDashPagos() {
  const hasCobros = typeof COBROS_RESULT !== 'undefined' && COBROS_RESULT.length > 0;

  const cobrados   = hasCobros ? COBROS_RESULT.filter(c => c.estado === 'COBRADO')   : [];
  const pendientes = hasCobros ? COBROS_RESULT.filter(c => c.estado === 'PENDIENTE') : [];
  const rechazados = hasCobros ? COBROS_RESULT.filter(c => c.estado === 'RECHAZADO') : [];
  const total      = hasCobros ? COBROS_RESULT.length : 0;

  const sumM = arr => arr.reduce((s,c) => s + Math.abs(c.fila?.sky?.monto||0), 0);
  const totCob = sumM(cobrados), totPen = sumM(pendientes), totRec = sumM(rechazados);
  const totTot = totCob + totPen + totRec;
  const pctCobOps = total   ? (cobrados.length / total   * 100) : 0;
  const pctCobMon = totTot  ? (totCob / totTot           * 100) : 0;

  // ── KPIs pagos — monto $ en grande, ops en chico ────────────────
  const _pk = (id,v) => { const e=document.getElementById(id); if(e) e.textContent=v; };
  _pk('pkpi-total',   _dFmtM(totTot));
  _pk('pkpi-mtotal',  total.toLocaleString('es-AR') + ' ops');
  _pk('pkpi-cob',     _dFmtM(totCob));
  _pk('pkpi-mcob',    cobrados.length.toLocaleString('es-AR') + ' ops');
  _pk('pkpi-pen',     _dFmtM(totPen));
  _pk('pkpi-mpen',    pendientes.length.toLocaleString('es-AR') + ' ops');
  _pk('pkpi-rec',     _dFmtM(totRec));
  _pk('pkpi-mrec',    rechazados.length.toLocaleString('es-AR') + ' ops');
  _pk('pkpi-pct',     pctCobMon.toFixed(1) + '%');
  _pk('pkpi-pctm',    pctCobOps.toFixed(1) + '% por ops');
  _pk('pkpi-mpen2',   _dFmtM(totPen));
  _pk('pkpi-pctpen',  totTot ? (totPen/totTot*100).toFixed(1)+'% del total' : '—');

  const PAG_IDS = ['ch-pag-donut-ops','ch-pag-donut-monto','ch-pag-proc','ch-pag-fecha-pago',
    'ch-pag-evol','ch-pag-suc-pend-ops','ch-pag-suc-pend-monto',
    'ch-pag-suc-comp','ch-pag-tarjeta'];

  // Mostrar/ocultar overlay "sin datos" SIN destruir los <canvas>
  // (si se reemplaza innerHTML del wrapper, el canvas queda eliminado del DOM
  //  y Chart.js no puede dibujarlo aunque haya datos luego)
  function _pagNoData(show) {
    PAG_IDS.forEach(id => {
      const canvas = document.getElementById(id);
      if (!canvas) return;
      canvas.style.display = show ? 'none' : '';
      const wrap = canvas.parentElement;
      if (!wrap) return;
      let overlay = wrap.querySelector('.pag-nodata');
      if (show) {
        if (!overlay) {
          overlay = document.createElement('div');
          overlay.className = 'pag-nodata';
          overlay.style.cssText = 'padding:30px;text-align:center;color:var(--m2);font-size:10px';
          overlay.textContent = 'Sin datos — cargá el archivo de liquidaciones';
          wrap.appendChild(overlay);
        }
        overlay.style.display = '';
      } else if (overlay) {
        overlay.style.display = 'none';
      }
    });
    const re = document.getElementById('dash-resumen-pagos');
    if (re && show) re.innerHTML = '<div style="padding:30px;text-align:center;color:var(--m2);font-size:10px">Sin datos — cargá el archivo de liquidaciones</div>';
  }

  if (!hasCobros) {
    _pagNoData(true);
    return;
  }

  // Hay datos — asegurarse de que los canvas estén visibles
  _pagNoData(false);

  const C_COB = '#34d399', C_PEN = '#f87171', C_REC = '#fbbf24';

  // ── 1. DONUT — Estado cobro (ops) ────────────────────────────────
  _dashCharts.pagDonutOps = new Chart(document.getElementById('ch-pag-donut-ops'), {
    type: 'doughnut',
    data: {
      labels: ['Cobrado','Pendiente','Rechazado'],
      datasets: [{ data:[cobrados.length, pendientes.length, rechazados.length],
        backgroundColor:[C_COB+'dd',C_PEN+'dd',C_REC+'dd'],
        borderWidth:1, borderColor:'#111827', hoverOffset:8 }]
    },
    options: {
      responsive:true, maintainAspectRatio:false, cutout:'60%',
      plugins: {
        legend: _dLegend('right'),
        tooltip: { callbacks: { label: ctx => ` ${ctx.label}: ${ctx.parsed.toLocaleString('es-AR')} (${total?(ctx.parsed/total*100).toFixed(1):0}%)` } }
      }
    }
  });

  // ── 2. DONUT — Estado cobro (monto $) ───────────────────────────
  _dashCharts.pagDonutMonto = new Chart(document.getElementById('ch-pag-donut-monto'), {
    type: 'doughnut',
    data: {
      labels: ['Cobrado','Pendiente','Rechazado'],
      datasets: [{ data:[totCob, totPen, totRec],
        backgroundColor:[C_COB+'dd',C_PEN+'dd',C_REC+'dd'],
        borderWidth:1, borderColor:'#111827', hoverOffset:8 }]
    },
    options: {
      responsive:true, maintainAspectRatio:false, cutout:'60%',
      plugins: {
        legend: _dLegend('right'),
        tooltip: { callbacks: { label: ctx => ` ${ctx.label}: ${_dFmtPeso(ctx.parsed)}` } }
      }
    }
  });

  // ── 3. GROUPED BAR — Por procesadora ────────────────────────────
  const procPag = { FISERV:{cob:0,pen:0,rec:0}, GETPOS:{cob:0,pen:0,rec:0} };
  COBROS_RESULT.forEach(c => {
    const p = c.fuenteCodigo === 'GETPOS' ? 'GETPOS' : 'FISERV';
    const k = c.estado==='COBRADO'?'cob':c.estado==='RECHAZADO'?'rec':'pen';
    procPag[p][k]++;
  });
  _dashCharts.pagProc = new Chart(document.getElementById('ch-pag-proc'), {
    type:'bar',
    data: {
      labels:['FISERV','GETPOS'],
      datasets:[
        { label:'Cobrado',   data:[procPag.FISERV.cob,procPag.GETPOS.cob], backgroundColor:C_COB+'cc', borderRadius:3, borderWidth:0 },
        { label:'Pendiente', data:[procPag.FISERV.pen,procPag.GETPOS.pen], backgroundColor:C_PEN+'cc', borderRadius:3, borderWidth:0 },
        { label:'Rechazado', data:[procPag.FISERV.rec,procPag.GETPOS.rec], backgroundColor:C_REC+'cc', borderRadius:3, borderWidth:0 },
      ]
    },
    options: {
      responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{ labels:{ color:DASH_CLR.txt, font:{size:9,family:'JetBrains Mono'}, boxWidth:10 } } },
      scales: _dScales()
    }
  });

  // ── 4. BAR — Calendario de cobros (fecha de pago) ───────────────
  const byFP = {};
  cobrados.forEach(c => {
    const f = c.liq?.fechaPago || c.liq?.fechaVenta || '?';
    if (f !== '?') byFP[f] = (byFP[f]||0) + 1;
  });
  const fpFechas = Object.keys(byFP).sort();
  _dashCharts.pagFechaPago = new Chart(document.getElementById('ch-pag-fecha-pago'), {
    type:'bar',
    data: {
      labels: fpFechas,
      datasets:[{ label:'Ops cobradas', data:fpFechas.map(f=>byFP[f]), backgroundColor:C_COB+'99', borderRadius:3, borderWidth:0 }]
    },
    options: {
      responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{display:false} },
      scales:{
        x:{ ticks:{ color:DASH_CLR.txt, font:{size:9}, maxRotation:40 }, grid:{ color:DASH_CLR.grid } },
        y:{ ticks:{ color:DASH_CLR.txt, font:{size:9} }, grid:{ color:DASH_CLR.grid }, beginAtZero:true }
      }
    }
  });

  // ── 5. LINE — Evolución por fecha de venta ──────────────────────
  const byFV = {};
  COBROS_RESULT.forEach(c => {
    const f = c.fila?.sky?.fecha || '?';
    if (!byFV[f]) byFV[f]={cob:0,pen:0,rec:0};
    const k = c.estado==='COBRADO'?'cob':c.estado==='RECHAZADO'?'rec':'pen';
    byFV[f][k]++;
  });
  const vFechas = Object.keys(byFV).filter(f=>f!=='?').sort();
  _dashCharts.pagEvol = new Chart(document.getElementById('ch-pag-evol'), {
    type:'line',
    data:{
      labels: vFechas,
      datasets:[
        { label:'Cobrado',   data:vFechas.map(f=>byFV[f].cob), borderColor:C_COB, backgroundColor:C_COB+'18', fill:true,  tension:.35, pointRadius:3, borderWidth:2 },
        { label:'Pendiente', data:vFechas.map(f=>byFV[f].pen), borderColor:C_PEN, backgroundColor:C_PEN+'18', fill:true,  tension:.35, pointRadius:3, borderWidth:2 },
        { label:'Rechazado', data:vFechas.map(f=>byFV[f].rec), borderColor:C_REC, backgroundColor:'transparent', fill:false, tension:.35, pointRadius:3, borderWidth:1, borderDash:[4,4] },
      ]
    },
    options:{
      responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{ labels:{ color:DASH_CLR.txt, font:{size:9,family:'JetBrains Mono'}, boxWidth:10 } } },
      scales:{
        x:{ ticks:{ color:DASH_CLR.txt, font:{size:9}, maxRotation:40 }, grid:{ color:DASH_CLR.grid } },
        y:{ ticks:{ color:DASH_CLR.txt, font:{size:9} }, grid:{ color:DASH_CLR.grid }, beginAtZero:true }
      }
    }
  });

  // ── 6. H-BAR — Top 15 suc por ops pendientes ────────────────────
  const sucPO = {};
  pendientes.forEach(c => { const s=c.fila?.sky?.suc??'?'; sucPO[s]=(sucPO[s]||0)+1; });
  const spoPairs = Object.entries(sucPO).sort((a,b)=>b[1]-a[1]).slice(0,15);
  _dashCharts.pagSucPendOps = new Chart(document.getElementById('ch-pag-suc-pend-ops'), {
    type:'bar',
    data:{ labels:spoPairs.map(([k])=>'Suc.'+k), datasets:[{ label:'Ops pendientes', data:spoPairs.map(([,v])=>v), backgroundColor:C_PEN+'bb', borderRadius:3, borderWidth:0 }] },
    options:{
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{display:false} },
      scales:{
        x:{ ticks:{ color:DASH_CLR.txt, font:{size:9} }, grid:{ color:DASH_CLR.grid } },
        y:{ ticks:{ color:DASH_CLR.txt, font:{size:9,family:'JetBrains Mono'} }, grid:{ display:false } }
      }
    }
  });

  // ── 7. H-BAR — Top 15 suc por monto pendiente ($) ───────────────
  const sucPM = {};
  pendientes.forEach(c => { const s=c.fila?.sky?.suc??'?'; sucPM[s]=(sucPM[s]||0)+Math.abs(c.fila?.sky?.monto||0); });
  const spmPairs = Object.entries(sucPM).sort((a,b)=>b[1]-a[1]).slice(0,15);
  _dashCharts.pagSucPendMonto = new Chart(document.getElementById('ch-pag-suc-pend-monto'), {
    type:'bar',
    data:{ labels:spmPairs.map(([k])=>'Suc.'+k), datasets:[{ label:'Monto pendiente', data:spmPairs.map(([,v])=>v), backgroundColor:C_PEN+'99', borderRadius:3, borderWidth:0 }] },
    options:{
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{display:false}, tooltip:{ callbacks:{ label: ctx => ' '+_dFmtPeso(ctx.parsed.x) } } },
      scales:{
        x:{ ticks:{ color:DASH_CLR.txt, font:{size:9}, callback:v=>_dFmtM(v) }, grid:{ color:DASH_CLR.grid } },
        y:{ ticks:{ color:DASH_CLR.txt, font:{size:9,family:'JetBrains Mono'} }, grid:{ display:false } }
      }
    }
  });

  // ── 8. STACKED BAR — Composición por sucursal (top 25) ──────────
  const sucComp2 = {};
  COBROS_RESULT.forEach(c => {
    const s = c.fila?.sky?.suc??'?';
    if (!sucComp2[s]) sucComp2[s]={cob:0,pen:0,rec:0};
    const k = c.estado==='COBRADO'?'cob':c.estado==='RECHAZADO'?'rec':'pen';
    sucComp2[s][k]++;
  });
  const scPairs = Object.entries(sucComp2)
    .sort((a,b)=>(b[1].cob+b[1].pen+b[1].rec)-(a[1].cob+a[1].pen+a[1].rec)).slice(0,25);
  _dashCharts.pagSucComp = new Chart(document.getElementById('ch-pag-suc-comp'), {
    type:'bar',
    data:{
      labels: scPairs.map(([k])=>'Suc.'+k),
      datasets:[
        { label:'Cobrado',   data:scPairs.map(([,v])=>v.cob), backgroundColor:C_COB+'cc', borderWidth:0 },
        { label:'Pendiente', data:scPairs.map(([,v])=>v.pen), backgroundColor:C_PEN+'cc', borderWidth:0 },
        { label:'Rechazado', data:scPairs.map(([,v])=>v.rec), backgroundColor:C_REC+'aa', borderWidth:0 },
      ]
    },
    options:{
      responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{ labels:{ color:DASH_CLR.txt, font:{size:9,family:'JetBrains Mono'}, boxWidth:10 } } },
      scales:{
        x:{ stacked:true, ticks:{ color:DASH_CLR.txt, font:{size:9}, maxRotation:45 }, grid:{ color:DASH_CLR.grid } },
        y:{ stacked:true, ticks:{ color:DASH_CLR.txt, font:{size:9} }, grid:{ color:DASH_CLR.grid } }
      }
    }
  });

  // ── 9. GROUPED BAR — Por tarjeta (cobrado/pendiente, monto) ─────
  const tarjPag = {};
  COBROS_RESULT.forEach(c => {
    const t = c.fila?.sky?.tarjeta || 'Sin dato';
    if (!tarjPag[t]) tarjPag[t]={cob:0,pen:0};
    if (c.estado==='COBRADO')   tarjPag[t].cob += Math.abs(c.fila?.sky?.monto||0);
    if (c.estado==='PENDIENTE') tarjPag[t].pen += Math.abs(c.fila?.sky?.monto||0);
  });
  const tpPairs = Object.entries(tarjPag)
    .sort((a,b)=>(b[1].cob+b[1].pen)-(a[1].cob+a[1].pen)).slice(0,12);
  _dashCharts.pagTarjeta = new Chart(document.getElementById('ch-pag-tarjeta'), {
    type:'bar',
    data:{
      labels: tpPairs.map(([k])=>k),
      datasets:[
        { label:'Cobrado',   data:tpPairs.map(([,v])=>v.cob), backgroundColor:C_COB+'cc', borderRadius:3, borderWidth:0 },
        { label:'Pendiente', data:tpPairs.map(([,v])=>v.pen), backgroundColor:C_PEN+'cc', borderRadius:3, borderWidth:0 },
      ]
    },
    options:{
      responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{ labels:{ color:DASH_CLR.txt, font:{size:9}, boxWidth:10 } } },
      scales:{
        x:{ ticks:{ color:DASH_CLR.txt, font:{size:9}, maxRotation:30 }, grid:{ color:DASH_CLR.grid } },
        y:{ ticks:{ color:DASH_CLR.txt, font:{size:9}, callback:v=>_dFmtM(v) }, grid:{ color:DASH_CLR.grid } }
      }
    }
  });

  // ── Resumen económico pagos ──────────────────────────────────────
  const reEl = document.getElementById('dash-resumen-pagos');
  if (reEl) {
    const line = (lbl,val,clr='var(--txt)') =>
      `<div style="display:flex;justify-content:space-between;align-items:center;
        padding:6px 10px;background:var(--s3);border-radius:var(--r4)">
        <span style="color:var(--m2)">${lbl}</span>
        <span style="font-weight:700;color:${clr}">${val}</span></div>`;
    reEl.innerHTML =
      line('Ops. analizadas',       total.toLocaleString('es-AR')) +
      line('Monto total analizado',  _dFmtPeso(totTot)) +
      line('Cobrado (ops)',          cobrados.length.toLocaleString('es-AR')+' ('+pctCobOps.toFixed(1)+'%)', C_COB) +
      line('Cobrado (monto)',        _dFmtPeso(totCob),    C_COB) +
      line('Pendiente (ops)',        pendientes.length.toLocaleString('es-AR'), C_PEN) +
      line('Pendiente (monto)',      _dFmtPeso(totPen),    C_PEN) +
      line('Rechazado (ops)',        rechazados.length.toLocaleString('es-AR'), C_REC) +
      line('Rechazado (monto)',      _dFmtPeso(totRec),    C_REC) +
      line('% Cobrado por monto',   pctCobMon.toFixed(1)+'%',
        pctCobMon>=80?C_COB:pctCobMon>=50?C_REC:C_PEN);
  }
}
