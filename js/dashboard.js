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

  _dashCharts.sucComp = new Chart(document.getElementById('ch-suc-comp'), {
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
