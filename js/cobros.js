// ═══════════════════════════════════════════════════════════════════
// COBROS.JS — Cruce cupones vs archivo de liquidaciones
// ═══════════════════════════════════════════════════════════════════
// Estrategia de matching por CÓDIGO ÚNICO:
//
//   Operaciones conciliadas (r.proc existe):
//     código = equipo + "_" + lote + "_" + ticket
//     • FISERV usa r.proc.ticket como cupón de la liquidación
//     • GETPOS  usa r.proc.cupon
//     • equipo  = r.proc.equipo (FISERV) | r.proc.pos (GETPOS)
//
//   Correcciones manuales (CORREGIDAS[idx] existe, sin proc):
//     código = lote_sky + "_" + cupon_corregido
//
//   SIN MATCH (sin proc ni corrección):
//     código = lote_sky + "_" + cupon_sky  (fallback)
//
//   Liquidaciones:
//     código = equipo + "_" + lote + "_" + cupon
//     (también indexado sin equipo para mayor tolerancia)
// ═══════════════════════════════════════════════════════════════════

let _LIQ_NORM = [];  // filas parseadas del archivo de liquidaciones
let _LIQ_IDX  = {};  // índice hash { codigo → [liqRow, ...] }

// ── Normalizar número: strip leading zeros ──────────────────────────
function normNum(v) {
  return String(v ?? '').trim().replace(/^0+/, '') || '0';
}

// ══════════════════════════════════════════════════════════════════
// PARSEO LIQUIDACIONES
// ══════════════════════════════════════════════════════════════════
function parseLiquidaciones(wb) {
  const ws   = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: false });

  _LIQ_NORM = rows
    .filter(r => r['Nro de Cupón'] != null && r['Nro de Lote'] != null)
    .map((r, i) => {
      const tarjeta = String(r['Tarjeta'] || '').trim();
      return {
        idx:          i,
        fechaVenta:   String(r['Fecha Venta']     || '').trim().slice(0, 10),
        fechaPago:    String(r['Fecha Pago']      || '').trim().slice(0, 10),
        fechaAdelanto:String(r['Fecha Adelanto']  || '').trim().slice(0, 10),
        nroLiq:       String(r['Nro Liquidación'] || '').trim(),
        equipo:       normNum(r['Nro Equipo']),
        nombreEquipo: String(r['Nombre de equipo'] || '').trim(),
        lote:         normNum(r['Nro de Lote']),
        cupon:        normNum(r['Nro de Cupón']),
        tarjeta,
        nroTarjeta:   String(r['Nro Tarjeta']         || '').trim(),
        aut:          normNum(r['Código Autorización']),
        cuotas:       parseInt(r['Cuotas']) || 1,
        importe:      parseFloat(String(r['Importe Venta'] || '').replace(/,/g, '')) || 0,
        nroCom:       normNum(r['Nro Comercio']),
        banco:        String(r['Banco Pagador']  || '').trim(),
        rechazo:      String(r['Rechazo']        || 'N').trim().toUpperCase(),
        arancel:      parseFloat(r['Arancel'])     || 0,
        ivaArancel:   parseFloat(r['IVA Arancel']) || 0,
        cfo:          parseFloat(r['CFO'])          || 0,
        ivaCfo:       parseFloat(r['Iva CFO'])      || 0,
        tipoOp:       String(r['Tipo operacion']  || '').trim(),
        bancoEmisor:  String(r['Banco Emisor']    || '').trim(),
        // Procesadora inferida del nombre de tarjeta
        proc: tarjeta.toUpperCase().includes('GETNET') ? 'GETPOS' : 'FISERV',
      };
    });

  _buildLiqIdx();
  return _LIQ_NORM;
}

// ── Código único de una fila de liquidación ──────────────────────────
function _codigoLiq(liq) {
  // Formato: equipo_lote_cupon
  return `${liq.equipo}_${liq.lote}_${liq.cupon}`;
}

// ── Construir índice hash ────────────────────────────────────────────
function _buildLiqIdx() {
  _LIQ_IDX = {};
  for (const liq of _LIQ_NORM) {
    const c = _codigoLiq(liq);       // equipo_lote_cupon
    const s = `${liq.lote}_${liq.cupon}`;  // sin equipo (tolerancia)

    (_LIQ_IDX[c] = _LIQ_IDX[c] || []).push(liq);
    (_LIQ_IDX[s] = _LIQ_IDX[s] || []).push(liq);

    // Índice por auth (fallback de último recurso)
    if (liq.aut && liq.aut !== '0') {
      const ka = `aut_${liq.aut}_${liq.equipo}`;
      (_LIQ_IDX[ka] = _LIQ_IDX[ka] || []).push(liq);
    }
  }
}

// ══════════════════════════════════════════════════════════════════
// CÓDIGO ÚNICO PARA UNA FILA DE RESULTADO
// ══════════════════════════════════════════════════════════════════
// Devuelve { full, short, autKey } con los distintos niveles de clave.
// full  = equipo_lote_ticket  (mayor especificidad, misma estructura que liquidación)
// short = lote_ticket         (sin equipo, para tolerancia)
// autKey = aut_auth_equipo    (fallback)
function _codigosDeRec(r) {
  const p   = r.proc;
  const cor = typeof CORREGIDAS !== 'undefined' ? CORREGIDAS[r.sky.idx] : null;

  if (p) {
    // ── Fila conciliada con procesadora ──────────────────────────
    // FISERV: cupón en la liquidación = p.ticket
    // GETPOS: cupón en la liquidación = p.cupon
    const equipo = normNum(p.equipo || p.pos || '');
    const lote   = normNum(p.lote   || r.sky.lote || '');
    const ticket = normNum(p.ticket || p.cupon    || '');   // ticket=FISERV, cupon=GETPOS
    const aut    = normNum(p.aut    || '');

    if (lote === '0' || ticket === '0') return null;

    return {
      full:   `${equipo}_${lote}_${ticket}`,
      short:  `${lote}_${ticket}`,
      autKey: (aut !== '0' && equipo !== '0') ? `aut_${aut}_${equipo}` : null,
      equipo, lote, ticket,
    };
  }

  if (cor) {
    // ── Corrección manual: recruzar no encontró proc ──────────────
    // Usar cupon corregido por el usuario + lote de SKY
    const lote  = normNum(r.sky.lote  || '');
    const cupon = normNum(cor.cupon   || r.sky.cupon || '');

    if (lote === '0' || cupon === '0') return null;
    return {
      full:   `0_${lote}_${cupon}`,  // sin equipo conocido
      short:  `${lote}_${cupon}`,
      autKey: null,
      equipo: '0', lote, ticket: cupon,
    };
  }

  // ── SIN MATCH sin corrección — fallback con campos SKY ───────────
  const lote  = normNum(r.sky.lote  || '');
  const cupon = normNum(r.sky.cupon || '');

  if (lote === '0' || cupon === '0') return null;
  return {
    full:   `0_${lote}_${cupon}`,
    short:  `${lote}_${cupon}`,
    autKey: null,
    equipo: '0', lote, ticket: cupon,
  };
}

// ── Buscar liquidación para una fila de RESULTADO ───────────────────
function _buscarEnLiq(r) {
  if (!_LIQ_NORM.length) return null;

  const codes = _codigosDeRec(r);
  if (!codes) return null;

  // 1. Código completo equipo+lote+ticket (más preciso)
  if (codes.full && codes.equipo !== '0') {
    const hit = _LIQ_IDX[codes.full];
    if (hit?.length) return hit[0];
  }

  // 2. Lote + ticket sin equipo (para SIN MATCH o equipo no disponible)
  const hit2 = _LIQ_IDX[codes.short];
  if (hit2?.length) return hit2[0];

  // 3. Auth + equipo (último recurso)
  if (codes.autKey) {
    const hit3 = _LIQ_IDX[codes.autKey];
    if (hit3?.length) return hit3[0];
  }

  return null;
}

// ══════════════════════════════════════════════════════════════════
// CRUCE
// ══════════════════════════════════════════════════════════════════
let COBROS_RESULT = [];  // [{ fila, liq, estado, codigoProc }]

function cruzarCobros() {
  COBROS_RESULT = [];
  if (!_LIQ_NORM.length || !RESULTADO.length) return;

  // Solo operaciones activas positivas (excluye integradas y devoluciones)
  const candidatos = RESULTADO.filter(r => !r.sky.integrado && !r.sky.esNeg);

  for (const r of candidatos) {
    const codes = _codigosDeRec(r);
    const liq   = _buscarEnLiq(r);

    let estado;
    if (!liq) {
      estado = 'PENDIENTE';
    } else if (liq.rechazo && liq.rechazo !== 'N' && liq.rechazo !== '') {
      estado = 'RECHAZADO';
    } else {
      estado = 'COBRADO';
    }

    COBROS_RESULT.push({
      fila:       r,
      liq,
      estado,
      codigoProc: codes?.full  || '—',   // para depuración / exportación
      codigoLiq:  liq ? _codigoLiq(liq) : '—',
      fuenteCodigo: r.proc
        ? (r.proc.ticket ? 'FISERV(ticket)' : 'GETPOS(cupon)')
        : (typeof CORREGIDAS !== 'undefined' && CORREGIDAS[r.sky.idx]
            ? 'MANUAL(corregido)' : 'SKY(fallback)'),
    });
  }
}

// ══════════════════════════════════════════════════════════════════
// RENDER — MÓDULO
// ══════════════════════════════════════════════════════════════════
let _cobrosTab = 'pendientes';
let _cobSuc    = '';
let _cobProc   = '';

function renderModuloCobros() {
  const cont = document.getElementById('mod-cobros');
  if (!cont) return;

  if (!RESULTADO.length) {
    cont.innerHTML = `<div class="cobros-empty">
      Primero realizá el <b style="color:var(--acc)">Cruce Automático</b> (módulo 1).</div>`;
    return;
  }
  if (!_LIQ_NORM.length) {
    cont.innerHTML = `<div class="cobros-empty">
      Cargá el archivo de <b style="color:var(--yel)">Liquidaciones</b> en el panel izquierdo
      para cruzar los cobros.</div>`;
    return;
  }

  cruzarCobros();

  const cobrados   = COBROS_RESULT.filter(c => c.estado === 'COBRADO');
  const pendientes = COBROS_RESULT.filter(c => c.estado === 'PENDIENTE');
  const rechazados = COBROS_RESULT.filter(c => c.estado === 'RECHAZADO');

  const sumM   = arr => arr.reduce((s, c) => s + Math.abs(c.fila.sky.monto || 0), 0);
  const totCob = sumM(cobrados), totPen = sumM(pendientes), totRec = sumM(rechazados);
  const totTot = totCob + totPen + totRec;
  const pctCob = totTot ? (totCob / totTot * 100) : 0;
  const pctPen = totTot ? (totPen / totTot * 100) : 0;

  const badge = document.getElementById('mcnt-cobros');
  if (badge) badge.textContent = pendientes.length || '0';

  cont.innerHTML = `
  <div class="cobros-wrap">

    <!-- KPIs -->
    <div class="cobros-kpis">
      <div class="cob-kpi" style="border-top:3px solid var(--grn)">
        <div class="cob-kpi-lbl">✓ Cobrados</div>
        <div class="cob-kpi-n" style="color:var(--grn)">${cobrados.length.toLocaleString()}</div>
        <div class="cob-kpi-m">${fmtARS(totCob)}</div>
        <div class="cob-kpi-pct">${pctCob.toFixed(1)}% del total</div>
      </div>
      <div class="cob-kpi" style="border-top:3px solid var(--red)">
        <div class="cob-kpi-lbl">⏳ Pendientes de cobro</div>
        <div class="cob-kpi-n" style="color:var(--red)">${pendientes.length.toLocaleString()}</div>
        <div class="cob-kpi-m" style="color:var(--red)">${fmtARS(totPen)}</div>
        <div class="cob-kpi-pct">${pctPen.toFixed(1)}% del total</div>
      </div>
      <div class="cob-kpi" style="border-top:3px solid var(--yel)">
        <div class="cob-kpi-lbl">✗ Rechazados</div>
        <div class="cob-kpi-n" style="color:var(--yel)">${rechazados.length.toLocaleString()}</div>
        <div class="cob-kpi-m" style="color:var(--yel)">${fmtARS(totRec)}</div>
        <div class="cob-kpi-pct">${totTot ? (totRec/totTot*100).toFixed(1) : 0}% del total</div>
      </div>
      <div class="cob-kpi" style="border-top:3px solid var(--acc)">
        <div class="cob-kpi-lbl">∑ Total analizado</div>
        <div class="cob-kpi-n" style="color:var(--acc)">${COBROS_RESULT.length.toLocaleString()}</div>
        <div class="cob-kpi-m">${fmtARS(totTot)}</div>
        <div class="cob-kpi-pct">ops. activas positivas</div>
      </div>
    </div>

    <!-- Barra de progreso -->
    <div style="padding:0 20px 14px">
      <div style="height:7px;background:var(--b1);border-radius:4px;overflow:hidden;display:flex">
        <div style="width:${pctCob.toFixed(1)}%;background:var(--grn);transition:width .5s"></div>
        <div style="width:${pctPen.toFixed(1)}%;background:var(--red);transition:width .5s"></div>
      </div>
      <div style="display:flex;justify-content:space-between;margin-top:5px;font-size:9px;color:var(--m2)">
        <span style="color:var(--grn)">▓ ${pctCob.toFixed(1)}% cobrado</span>
        <span style="color:var(--red)">▓ ${pctPen.toFixed(1)}% pendiente</span>
      </div>
    </div>

    <!-- Tab strip -->
    <div class="tab-strip" id="tstrip-cobros"
         style="padding:0 20px;border-bottom:1px solid var(--b1);flex-shrink:0">
      <button class="tb ${_cobrosTab==='cobrados'?'active':''}"
        onclick="showCobrosTab('cobrados',this)">
        ✓ Cobrados
        <span class="cnt" style="background:rgba(52,211,153,.15);color:var(--grn)">${cobrados.length}</span>
      </button>
      <button class="tb ${_cobrosTab==='pendientes'?'active':''}"
        onclick="showCobrosTab('pendientes',this)">
        ⏳ Pendientes
        <span class="cnt" style="background:rgba(248,113,113,.15);color:var(--red)">${pendientes.length}</span>
      </button>
      <button class="tb ${_cobrosTab==='rechazados'?'active':''}"
        onclick="showCobrosTab('rechazados',this)">
        ✗ Rechazados
        <span class="cnt" style="background:rgba(251,191,36,.15);color:var(--yel)">${rechazados.length}</span>
      </button>
    </div>

    <div id="cobros-tab-body"
         style="flex:1;overflow:hidden;display:flex;flex-direction:column;min-height:0"></div>
  </div>`;

  showCobrosTab(_cobrosTab);
}

// ── Cambiar tab activa ──────────────────────────────────────────────
function showCobrosTab(tab, btn) {
  _cobrosTab = tab;
  if (btn) {
    document.querySelectorAll('#tstrip-cobros .tb').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
  }
  const body = document.getElementById('cobros-tab-body');
  if (!body) return;

  const MAP = { cobrados:'COBRADO', pendientes:'PENDIENTE', rechazados:'RECHAZADO' };
  let rows = COBROS_RESULT.filter(c => c.estado === MAP[tab]);
  if (_cobSuc)  rows = rows.filter(c => c.fila.sky.suc === _cobSuc);
  if (_cobProc) rows = rows.filter(c => c.fuenteCodigo.startsWith(_cobProc));

  _renderTablaCobros(body, tab, rows);
}

// ── Tabla interna ───────────────────────────────────────────────────
function _renderTablaCobros(body, tipo, rows) {
  if (!rows.length) {
    body.innerHTML = `<div style="padding:40px;text-align:center;color:var(--m2)">
      No hay operaciones ${tipo}.</div>`;
    return;
  }

  // Sucursales únicas para filtro
  const MAP2 = { cobrados:'COBRADO', pendientes:'PENDIENTE', rechazados:'RECHAZADO' };
  const allRows = COBROS_RESULT.filter(c => c.estado === MAP2[tipo]);
  const sucs = [...new Set(allRows.map(c => c.fila.sky.suc))].sort((a,b)=>+a-+b);
  const totMonto = rows.reduce((s, c) => s + Math.abs(c.fila.sky.monto || 0), 0);

  let HDR, tplRow;

  if (tipo === 'cobrados') {
    HDR = ['Suc.','Fecha Vta.','Vendedor','Tarjeta SKY','Plan',
           'Cupón','Lote','Monto SKY',
           'Fecha Pago','Nro Liq.','Equipo','Tarjeta Liq.','Cod.Auth.',
           'Tipo Op.','Código Único'];
    tplRow = c => {
      const s = c.fila.sky, l = c.liq;
      const fuenteBadge = _fuenteBadge(c.fuenteCodigo);
      return `<tr>
        <td>${s.suc}</td>
        <td>${s.fecha}</td>
        <td class="td-trunc" title="${s.vendedor||''}">${s.vendedor||'—'}</td>
        <td>${s.tarjeta}</td>
        <td class="td-trunc" style="font-size:9px" title="${s.plan||''}">${s.plan||'—'}</td>
        <td class="num">${s.cupon}</td>
        <td class="num">${s.lote}</td>
        <td class="num">${fmtARS(s.monto)}</td>
        <td style="color:var(--grn);font-weight:700">${l.fechaPago||'—'}</td>
        <td class="num" style="color:var(--cyn)">${l.nroLiq}</td>
        <td class="num" style="font-size:9px">${l.equipo}</td>
        <td class="td-trunc" style="font-size:9px" title="${l.tarjeta}">${l.tarjeta}</td>
        <td class="num" style="font-size:9px">${l.aut||'—'}</td>
        <td style="font-size:9px">${l.tipoOp||'—'}</td>
        <td class="num" style="font-size:8px;color:var(--m2)">${c.codigoProc}${fuenteBadge}</td>
      </tr>`;
    };
  } else if (tipo === 'pendientes') {
    HDR = ['Estado Cruce','Suc.','Fecha Vta.','Vendedor','Tarjeta','Plan',
           'Cupón','Lote','Monto SKY','Proc. Esperada','Código Único','Fuente'];
    tplRow = c => {
      const s = c.fila.sky;
      const procEsp = c.fila.procEsperada;
      const fuenteBadge = _fuenteBadge(c.fuenteCodigo);
      return `<tr>
        <td>${estadoBadge(c.fila.estado)}</td>
        <td>${s.suc}</td>
        <td>${s.fecha}</td>
        <td class="td-trunc" title="${s.vendedor||''}">${s.vendedor||'—'}</td>
        <td>${s.tarjeta}</td>
        <td class="td-trunc" style="font-size:9px" title="${s.plan||''}">${s.plan||'—'}</td>
        <td class="num">${s.cupon}</td>
        <td class="num">${s.lote}</td>
        <td class="num" style="color:var(--red);font-weight:700">${fmtARS(s.monto)}</td>
        <td><span class="st ${procEsp==='FISERV'?'st-fis':procEsp==='GETPOS'?'st-gp':''}">${procEsp||'—'}</span></td>
        <td class="num" style="font-size:8px;color:var(--m2)">${c.codigoProc}</td>
        <td style="font-size:8px">${fuenteBadge}</td>
      </tr>`;
    };
  } else { // rechazados
    HDR = ['Suc.','Fecha Vta.','Vendedor','Tarjeta','Cupón','Lote','Monto SKY',
           'Fecha Vta. Liq.','Nro Liq.','Banco Pagador','Rechazo','Tarjeta Liq.','Código Único'];
    tplRow = c => {
      const s = c.fila.sky, l = c.liq;
      return `<tr>
        <td>${s.suc}</td>
        <td>${s.fecha}</td>
        <td class="td-trunc" title="${s.vendedor||''}">${s.vendedor||'—'}</td>
        <td>${s.tarjeta}</td>
        <td class="num">${s.cupon}</td>
        <td class="num">${s.lote}</td>
        <td class="num" style="color:var(--yel)">${fmtARS(s.monto)}</td>
        <td>${l?.fechaVenta||'—'}</td>
        <td class="num">${l?.nroLiq||'—'}</td>
        <td style="font-size:9px">${l?.banco||'—'}</td>
        <td style="color:var(--red);font-weight:700">${l?.rechazo||'—'}</td>
        <td class="td-trunc" style="font-size:9px" title="${l?.tarjeta||''}">${l?.tarjeta||'—'}</td>
        <td class="num" style="font-size:8px;color:var(--m2)">${c.codigoProc}</td>
      </tr>`;
    };
  }

  body.innerHTML = `
  <div class="cobros-toolbar">
    <button class="btn-exp" onclick="exportarCobros('${tipo}')">↓ Exportar Excel</button>
    <select class="filter-sel" onchange="_cobFiltroSuc(this.value)">
      <option value="">Todas las sucursales</option>
      ${sucs.map(s=>`<option value="${s}" ${_cobSuc===s?'selected':''}>${s}</option>`).join('')}
    </select>
    <select class="filter-sel" onchange="_cobFiltroProc(this.value)">
      <option value="">Ambas procesadoras</option>
      <option value="FISERV" ${_cobProc==='FISERV'?'selected':''}>FISERV</option>
      <option value="GETPOS" ${_cobProc==='GETPOS'?'selected':''}>GETPOS</option>
    </select>
    <span style="font-size:10px;color:var(--m2)">${rows.length.toLocaleString()} ops · ${fmtARS(totMonto)}</span>
  </div>
  <div class="tbl-wrap">
    <table class="res-tbl">
      <thead><tr>${HDR.map(h=>`<th>${h}</th>`).join('')}</tr></thead>
      <tbody>${rows.map(tplRow).join('')}</tbody>
    </table>
  </div>`;
}

// ── Badge de fuente del código ────────────────────────────────────────
function _fuenteBadge(fuente) {
  if (!fuente) return '';
  if (fuente.startsWith('FISERV'))  return `<span style="font-size:7px;margin-left:3px;background:rgba(79,142,247,.12);color:var(--acc);border:1px solid rgba(79,142,247,.25);border-radius:2px;padding:1px 3px">FIS</span>`;
  if (fuente.startsWith('GETPOS'))  return `<span style="font-size:7px;margin-left:3px;background:rgba(167,139,250,.12);color:var(--vio);border:1px solid rgba(167,139,250,.25);border-radius:2px;padding:1px 3px">GP</span>`;
  if (fuente.startsWith('MANUAL'))  return `<span style="font-size:7px;margin-left:3px;background:rgba(251,191,36,.12);color:var(--yel);border:1px solid rgba(251,191,36,.25);border-radius:2px;padding:1px 3px">MAN</span>`;
  return `<span style="font-size:7px;margin-left:3px;background:rgba(107,114,128,.1);color:var(--m2);border:1px solid rgba(107,114,128,.2);border-radius:2px;padding:1px 3px">SKY</span>`;
}

// ── Filtros ───────────────────────────────────────────────────────────
function _cobFiltroSuc(v)  { _cobSuc  = v; showCobrosTab(_cobrosTab); }
function _cobFiltroProc(v) { _cobProc = v; showCobrosTab(_cobrosTab); }

// ══════════════════════════════════════════════════════════════════
// EXPORTAR EXCEL
// ══════════════════════════════════════════════════════════════════
function exportarCobros(tipo) {
  const MAP = { cobrados:'COBRADO', pendientes:'PENDIENTE', rechazados:'RECHAZADO' };
  let rows = COBROS_RESULT.filter(c => c.estado === MAP[tipo]);
  if (_cobSuc)  rows = rows.filter(c => c.fila.sky.suc === _cobSuc);
  if (_cobProc) rows = rows.filter(c => c.fuenteCodigo.startsWith(_cobProc));
  if (!rows.length) { alert('No hay datos para exportar.'); return; }

  let HDR, dataFn;

  if (tipo === 'cobrados') {
    HDR = ['Suc.','Fecha Venta','Vendedor','Tarjeta SKY','Plan','Cupón SKY','Lote SKY','Monto SKY',
           'Fecha Pago','Nro Liquidación','Equipo','Tarjeta Liq.','Cód.Auth.',
           'Tipo Operación','Arancel','IVA Arancel','CFO','Banco Emisor',
           'Estado Cruce','Código Único Proc.','Código Único Liq.','Fuente Código'];
    dataFn = c => {
      const s = c.fila.sky, l = c.liq;
      return [s.suc, s.fecha, s.vendedor||'', s.tarjeta, s.plan||'',
        s.cupon, s.lote, s.monto,
        l.fechaPago, l.nroLiq, l.equipo, l.tarjeta, l.aut, l.tipoOp,
        l.arancel, l.ivaArancel, l.cfo, l.bancoEmisor,
        c.fila.estado, c.codigoProc, c.codigoLiq, c.fuenteCodigo];
    };
  } else if (tipo === 'pendientes') {
    HDR = ['Estado Cruce','Suc.','Fecha Venta','Vendedor','Tarjeta','Plan',
           'Cupón','Lote','Monto SKY','Proc. Esperada','Nro Comercio',
           'Código Único Proc.','Fuente Código'];
    dataFn = c => {
      const s = c.fila.sky;
      return [c.fila.estado, s.suc, s.fecha, s.vendedor||'', s.tarjeta, s.plan||'',
        s.cupon, s.lote, s.monto, c.fila.procEsperada||'', s.nroCom||'',
        c.codigoProc, c.fuenteCodigo];
    };
  } else {
    HDR = ['Suc.','Fecha Venta','Vendedor','Tarjeta','Cupón','Lote','Monto SKY',
           'Fecha Vta. Liq.','Nro Liquidación','Banco Pagador','Rechazo','Tarjeta Liq.',
           'Estado Cruce','Código Único Proc.','Código Único Liq.','Fuente Código'];
    dataFn = c => {
      const s = c.fila.sky, l = c.liq;
      return [s.suc, s.fecha, s.vendedor||'', s.tarjeta,
        s.cupon, s.lote, s.monto,
        l?.fechaVenta||'', l?.nroLiq||'', l?.banco||'', l?.rechazo||'', l?.tarjeta||'',
        c.fila.estado, c.codigoProc, c.codigoLiq||'—', c.fuenteCodigo];
    };
  }

  const ws  = XLSX.utils.aoa_to_sheet([HDR, ...rows.map(dataFn)]);
  const wb2 = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb2, ws, tipo[0].toUpperCase() + tipo.slice(1));
  const ts = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb2, `Cobros_${tipo}_${ts}.xlsx`);
}
