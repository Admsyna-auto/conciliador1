// ═══════════════════════════════════════════════════════════════════
// LIQUIDACIONES.JS — Cruce de operaciones confirmadas vs liquidaciones
// ═══════════════════════════════════════════════════════════════════
// Flujo:
//  1. OPERACIONES cruza cobrado (FISERV/GETPOS/GoC) vs facturado (Skylab)
//  2. LIQUIDACIONES cruza las operaciones confirmadas vs lo que la
//     procesadora declara haber pagado (archivos de liquidación)
//
// GoC ya sube las liquidaciones en OPERACIONES → no necesita archivo nuevo.
// FISERV y GETPOS necesitan archivos de liquidación separados.
// ═══════════════════════════════════════════════════════════════════

// ── Estado ───────────────────────────────────────────────────────────
let _LIQ_FIS = [];   // Filas parseadas del archivo de liquidación FISERV
let _LIQ_GP  = [];   // Filas parseadas del archivo de liquidación GETPOS

// ── Helpers locales ──────────────────────────────────────────────────
function _liqGCol(row, ...cands) {
  for (const c of cands) {
    const v = row[c];
    if (v !== undefined && v !== null && v !== '') return v;
    // Búsqueda case-insensitive
    const key = Object.keys(row).find(k => k.trim().toLowerCase() === c.trim().toLowerCase());
    if (key && row[key] !== null && row[key] !== undefined && row[key] !== '') return row[key];
  }
  return null;
}

function _liqNorm(v) {
  return String(v ?? '').trim().replace(/^0+/, '') || '0';
}

function _liqMonto(v) {
  if (v === null || v === undefined || v === '') return 0;
  const s = String(v).replace(/\./g,'').replace(',','.');
  return parseFloat(s) || 0;
}

function _liqFecha(v) {
  if (!v) return '';
  const s = String(v).trim();
  // DD/MM/YYYY → YYYY-MM-DD
  const m1 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m1) return `${m1[3]}-${m1[2].padStart(2,'0')}-${m1[1].padStart(2,'0')}`;
  // YYYY-MM-DD ya normalizado
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0,10);
  // DD-MM-YYYY
  const m2 = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})/);
  if (m2) return `${m2[3]}-${m2[2].padStart(2,'0')}-${m2[1].padStart(2,'0')}`;
  return s;
}

// ── PARSER: Liquidación FISERV ───────────────────────────────────────
// Detecta columnas automáticamente. Extrae: lote, cupon, aut, monto_bruto,
// monto_neto, fecha_pago, terminal.
// La clave de cruce con RESULTADO es el N° de Lote.
function parseLiqFiserv(wb) {
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: true });
  if (!rows.length) return [];

  const out = [];
  for (const r of rows) {
    const loteRaw  = _liqGCol(r,'Lote','N° Lote','Nro Lote','Numero Lote','Nro. Lote','LOTE','Nro.Lote','NroLote');
    const brutoRaw = _liqGCol(r,'Monto Bruto','Importe Bruto','Bruto','Monto Total','MONTO BRUTO','Importe');
    const netoRaw  = _liqGCol(r,'Monto Neto','Importe Neto','Neto','MONTO NETO','Neto Pagado');
    const fechaRaw = _liqGCol(r,'Fecha Pago','Fecha Deposito','Fecha Depósito','Fecha de Pago',
                               'FECHA PAGO','Fecha','FECHA');
    const termRaw  = _liqGCol(r,'Terminal','Equipo','Nro Terminal','N° Terminal','Codigo Terminal');
    const cupRaw   = _liqGCol(r,'Cupon','Cupón','Ticket','Nro Cupon','N° Cupon','Cupón');
    const autRaw   = _liqGCol(r,'Autorizacion','Autorización','Aut','Cod Aut','Cód. Aut.');

    const lote = typeof loteRaw === 'number'
      ? String(Math.round(loteRaw))
      : _liqNorm(loteRaw);

    if (!lote || lote === '0') continue;

    out.push({
      lote,
      cupon:     typeof cupRaw === 'number' ? String(Math.round(cupRaw)) : _liqNorm(cupRaw),
      aut:       typeof autRaw === 'number' ? String(Math.round(autRaw)) : _liqNorm(autRaw),
      monto_bruto: typeof brutoRaw === 'number' ? brutoRaw : _liqMonto(brutoRaw),
      monto_neto:  typeof netoRaw  === 'number' ? netoRaw  : _liqMonto(netoRaw),
      fecha:     _liqFecha(fechaRaw),
      terminal:  typeof termRaw === 'number' ? String(Math.round(termRaw)) : String(termRaw ?? '').trim(),
    });
  }
  console.log(`[LIQ-FIS] Parseados: ${out.length} registros`);
  return out;
}

// ── PARSER: Liquidación GETPOS ───────────────────────────────────────
// Clave de cruce: N° Autorización
function parseLiqGetpos(wb) {
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: true });
  if (!rows.length) return [];

  const out = [];
  for (const r of rows) {
    const autRaw   = _liqGCol(r,'Autorizacion','Autorización','Aut','Cod Aut','Cód. Aut.',
                                'Codigo de Autorizacion','AUTORIZACION','Aut.');
    const brutoRaw = _liqGCol(r,'Monto Bruto','Importe Bruto','Bruto','Monto Transaccion',
                                'Monto Transacción','MONTO BRUTO','Importe');
    const netoRaw  = _liqGCol(r,'Monto Neto','Importe Neto','Neto','MONTO NETO');
    const fechaRaw = _liqGCol(r,'Fecha Pago','Fecha Deposito','Fecha Depósito','Fecha',
                                'Fecha de Pago','FECHA');
    const cupRaw   = _liqGCol(r,'Cupon','Cupón','Ticket','Nro Cupon');

    const aut = typeof autRaw === 'number'
      ? String(Math.round(autRaw))
      : _liqNorm(autRaw);

    if (!aut || aut === '0') continue;

    out.push({
      aut,
      cupon:      typeof cupRaw   === 'number' ? String(Math.round(cupRaw))   : _liqNorm(cupRaw),
      monto_bruto: typeof brutoRaw === 'number' ? brutoRaw : _liqMonto(brutoRaw),
      monto_neto:  typeof netoRaw  === 'number' ? netoRaw  : _liqMonto(netoRaw),
      fecha:      _liqFecha(fechaRaw),
    });
  }
  console.log(`[LIQ-GP] Parseados: ${out.length} registros`);
  return out;
}

// ── CRUCE: FISERV ────────────────────────────────────────────────────
// Agrupa RESULTADO FISERV por lote y verifica contra _LIQ_FIS.
function _cruzarLiqFiserv() {
  if (typeof RESULTADO === 'undefined' || !RESULTADO.length) return null;

  // Filas FISERV en RESULTADO (no GETPOS, no GOC)
  const filasFis = RESULTADO.filter(r => !r.sky?.esGETPos && !r.sky?.esGOCUOTAS);

  // Índice de lotes en la liquidación
  // Un lote puede tener múltiples rows (por terminal); usamos Set de lotes
  const lotesLiq = new Set(_LIQ_FIS.map(r => r.lote));
  // Índice completo para "Extras": lotes que no matchean con nada en RESULTADO
  const liqByLote = {};
  for (const r of _LIQ_FIS) {
    if (!liqByLote[r.lote]) liqByLote[r.lote] = [];
    liqByLote[r.lote].push(r);
  }

  // Lotes usados por RESULTADO
  const lotesUsados = new Set();

  const liquidadas    = [];
  const noLiquidadas  = [];
  const sinConfirmar  = [];   // SIN MATCH en OPERACIONES → nunca llegan a la liq.

  for (const fila of filasFis) {
    const esOK = fila.estado?.startsWith('OK');
    const lote = fila.proc?.lote || fila.sky?.lote || '';
    const loteN = _liqNorm(lote);

    if (!esOK) {
      sinConfirmar.push(fila);
      continue;
    }

    if (!_LIQ_FIS.length) {
      // Sin archivo de liquidación: todas quedan en un limbo
      noLiquidadas.push({ fila, loteN, enLiq: null });
      continue;
    }

    if (loteN && loteN !== '0' && lotesLiq.has(loteN)) {
      lotesUsados.add(loteN);
      liquidadas.push({ fila, loteN, liqRows: liqByLote[loteN] || [] });
    } else {
      noLiquidadas.push({ fila, loteN, enLiq: false });
    }
  }

  // Extras: lotes en liquidación que no usó ninguna fila RESULTADO
  const extrasLotes = [...lotesLiq].filter(l => !lotesUsados.has(l));
  const extras = extrasLotes.flatMap(l => liqByLote[l]);

  // KPIs de monto
  const montoLiquidado   = liquidadas.reduce((s, x) => {
    // Si hay rows de liq para este lote, usar monto_neto de liq; si no, usar sky.monto
    const net = x.liqRows.reduce((a, r) => a + (r.monto_neto || r.monto_bruto || 0), 0);
    return s + (net || Math.abs(x.fila.sky?.monto || 0));
  }, 0);
  const montoNoLiq = noLiquidadas.reduce((s, x) => s + Math.abs(x.fila.sky?.monto || 0), 0);
  const montoExtras = extras.reduce((s, r) => s + (r.monto_neto || r.monto_bruto || 0), 0);

  return { liquidadas, noLiquidadas, sinConfirmar, extras,
           montoLiquidado, montoNoLiq, montoExtras,
           totalOK: liquidadas.length + noLiquidadas.length,
           tieneLiq: _LIQ_FIS.length > 0 };
}

// ── CRUCE: GETPOS ────────────────────────────────────────────────────
// Cruza por N° Autorización
function _cruzarLiqGetpos() {
  if (typeof RESULTADO === 'undefined' || !RESULTADO.length) return null;

  const filasGP = RESULTADO.filter(r => r.sky?.esGETPos && !r.sky?.esGOCUOTAS);

  // Índice autorización → fila liquidación
  const autLiq = {};
  for (const r of _LIQ_GP) {
    if (r.aut && r.aut !== '0') autLiq[r.aut] = r;
  }
  const autsUsados = new Set();

  const liquidadas   = [];
  const noLiquidadas = [];
  const sinConfirmar = [];

  for (const fila of filasGP) {
    const esOK = fila.estado?.startsWith('OK');
    const aut  = fila.proc?.aut || fila.sky?.aut || '';
    const autN = _liqNorm(aut);

    if (!esOK) {
      sinConfirmar.push(fila);
      continue;
    }

    if (!_LIQ_GP.length) {
      noLiquidadas.push({ fila, autN, enLiq: null });
      continue;
    }

    if (autN && autN !== '0' && autLiq[autN]) {
      autsUsados.add(autN);
      liquidadas.push({ fila, autN, liqRow: autLiq[autN] });
    } else {
      noLiquidadas.push({ fila, autN, enLiq: false });
    }
  }

  const extras = Object.entries(autLiq)
    .filter(([a]) => !autsUsados.has(a))
    .map(([, r]) => r);

  const montoLiquidado  = liquidadas.reduce((s, x) =>
    s + (x.liqRow?.monto_neto || x.liqRow?.monto_bruto || Math.abs(x.fila.sky?.monto || 0)), 0);
  const montoNoLiq  = noLiquidadas.reduce((s, x) => s + Math.abs(x.fila.sky?.monto || 0), 0);
  const montoExtras = extras.reduce((s, r) => s + (r.monto_neto || r.monto_bruto || 0), 0);

  return { liquidadas, noLiquidadas, sinConfirmar, extras,
           montoLiquidado, montoNoLiq, montoExtras,
           totalOK: liquidadas.length + noLiquidadas.length,
           tieneLiq: _LIQ_GP.length > 0 };
}

// ── CRUCE: Go Cuotas ─────────────────────────────────────────────────
// Los archivos GoC YA son las liquidaciones → no necesita archivo extra.
// Usa RESULTADO + _GOC_PAGOS + _GOC_CELULAR.
function _cruzarLiqGoC() {
  if (typeof RESULTADO === 'undefined' || !RESULTADO.length) return null;

  const filasGoC = RESULTADO.filter(r => r.sky?.esGOCUOTAS);
  const _gocAll  = [
    ...(typeof _GOC_PAGOS   !== 'undefined' ? _GOC_PAGOS   : []),
    ...(typeof _GOC_CELULAR !== 'undefined' ? _GOC_CELULAR : []),
  ];

  // Órdenes usadas en RESULTADO para detectar extras
  const ordenesUsadas = new Set();

  const liquidadas   = [];
  const noLiquidadas = [];

  for (const fila of filasGoC) {
    const esGoC = fila.estado?.includes('GoC') || fila.estado?.includes('GoCelular');
    if (esGoC) {
      const ord = _liqNorm(fila.proc?.cupon || fila.proc?.ticket || '');
      if (ord && ord !== '0') ordenesUsadas.add(ord);
      liquidadas.push({ fila, orden: ord });
    } else {
      noLiquidadas.push({ fila, orden: '' });
    }
  }

  // Extras: pagos GoC que no matchearon con ninguna fila Skylab
  const extras = _gocAll.filter(p => {
    const o = _liqNorm(p.orden || '');
    return o && o !== '0' && !ordenesUsadas.has(o);
  });

  const montoLiquidado  = liquidadas.reduce((s, x) => s + Math.abs(x.fila.sky?.monto || 0), 0);
  const montoNoLiq      = noLiquidadas.reduce((s, x) => s + Math.abs(x.fila.sky?.monto || 0), 0);
  const montoExtras     = extras.reduce((s, p) => s + (p.importe || 0), 0);

  return { liquidadas, noLiquidadas, sinConfirmar:[], extras,
           montoLiquidado, montoNoLiq, montoExtras,
           totalOK: liquidadas.length + noLiquidadas.length,
           tieneLiq: _gocAll.length > 0 };
}

// ── UPLOAD handlers ───────────────────────────────────────────────────
function liqCargarFiserv(input) {
  const file = input.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, { type:'array' });
      _LIQ_FIS = parseLiqFiserv(wb);
      if (!_LIQ_FIS.length) {
        typeof _showToast === 'function'
          ? _showToast('⚠ No se encontraron filas válidas en el archivo FISERV (¿columna Lote presente?)')
          : alert('No se encontraron filas con N° de Lote.');
      } else {
        typeof _showToast === 'function'
          ? _showToast(`✓ Liquidación FISERV cargada · ${_LIQ_FIS.length.toLocaleString('es-AR')} registros`)
          : null;
      }
      renderModuloLiqFiserv();
    } catch(err) {
      alert('Error leyendo archivo: ' + err.message);
      console.error(err);
    }
    input.value = '';
  };
  reader.readAsArrayBuffer(file);
}

function liqCargarGetpos(input) {
  const file = input.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, { type:'array' });
      _LIQ_GP = parseLiqGetpos(wb);
      if (!_LIQ_GP.length) {
        typeof _showToast === 'function'
          ? _showToast('⚠ No se encontraron filas válidas (¿columna Autorización presente?)')
          : alert('No se encontraron filas con N° Autorización.');
      } else {
        typeof _showToast === 'function'
          ? _showToast(`✓ Liquidación GETPOS cargada · ${_LIQ_GP.length.toLocaleString('es-AR')} registros`)
          : null;
      }
      renderModuloLiqGetpos();
    } catch(err) {
      alert('Error leyendo archivo: ' + err.message);
      console.error(err);
    }
    input.value = '';
  };
  reader.readAsArrayBuffer(file);
}

// ── RENDER helpers ────────────────────────────────────────────────────
function _liqFmtARS(v) {
  return typeof fmtARS === 'function' ? fmtARS(v) : `$${(+v||0).toFixed(2)}`;
}

function _liqPct(n, d) {
  if (!d) return '0.0';
  return ((n / d) * 100).toFixed(1);
}

// Barra de progreso interna
function _liqPBar(pct, color) {
  return `<div style="width:100%;height:3px;background:var(--b1);border-radius:2px;margin-top:4px">
    <div style="width:${Math.min(100,pct)}%;height:100%;background:${color};border-radius:2px"></div>
  </div>`;
}

// Genera el HTML de KPIs + tab strip para el panel
function _liqBuildPanel(opts) {
  const { id, kpis, tabs, activeTab, fileSection } = opts;
  return `
  <div style="display:flex;flex-direction:column;flex:1;min-height:0;overflow:hidden">

    ${fileSection || ''}

    <!-- KPIs -->
    <div style="display:grid;grid-template-columns:repeat(${kpis.length},1fr);gap:8px;
      padding:12px 18px;background:var(--s1);border-bottom:1px solid var(--b1);flex-shrink:0">
      ${kpis.map(k => `
        <div class="dif-kpi" style="border-color:${k.borderColor||'var(--b2)'}">
          <div class="dif-kpi-lbl">${k.label}</div>
          <div class="dif-kpi-val ${k.cls||''}">${k.val}</div>
          <div style="font-size:8px;color:var(--m2)">${k.sub||''}</div>
          ${k.pct !== undefined ? _liqPBar(k.pct, k.borderColor||'var(--acc)') : ''}
        </div>`).join('')}
    </div>

    <!-- Tab strip -->
    <div class="tab-strip" id="tab-strip-liq-${id}">
      ${tabs.map((t,i) => `
        <button class="tb${i===0?' active':''}" data-tab="${t.key}"
          onclick="showTab('liqtab-${id}-${t.key}','tab-strip-liq-${id}',this);_liqRenderTab('${id}','${t.key}')">
          ${t.label} <span class="cnt" style="${t.cntStyle||''}">${t.count}</span>
        </button>`).join('')}
    </div>

    <!-- Tab bodies -->
    ${tabs.map((t,i) => `
      <div class="tab-body${i===0?' active':''}" id="liqtab-${id}-${t.key}"
        style="flex-direction:column;flex:1;min-height:0">
        <div class="tbl-wrap">
          <table id="tbl-liq-${id}-${t.key}"><thead></thead><tbody></tbody></table>
        </div>
      </div>`).join('')}

  </div>`;
}

// Renderiza una tabla en un panel de liquidación
function _liqRenderTable(tableId, cols, rows) {
  const tbl = document.getElementById(tableId);
  if (!tbl) return;
  tbl.querySelector('thead').innerHTML =
    `<tr>${cols.map(c => `<th>${c.label}</th>`).join('')}</tr>`;
  tbl.querySelector('tbody').innerHTML = rows.length
    ? rows.map(row => `<tr>${cols.map(c => `<td${c.cls?' class="'+c.cls+'"':''}>${row[c.key] ?? '—'}</td>`).join('')}</tr>`).join('')
    : `<tr><td colspan="${cols.length}" style="text-align:center;padding:32px;color:var(--m2);font-style:italic">Sin registros</td></tr>`;
}

// ── RENDER: Sección de carga de archivo ───────────────────────────────
function _liqFileSection(proc, procLabel, icon, hasFile, count) {
  const p = proc.toLowerCase();
  return `
  <div style="display:flex;align-items:center;gap:12px;padding:10px 18px;
    background:var(--s2);border-bottom:1px solid var(--b1);flex-shrink:0">
    <span style="font-size:14px">${icon}</span>
    <div style="font-size:10px;font-weight:700;color:var(--txt)">${procLabel} · Archivo de Liquidación</div>
    ${hasFile
      ? `<div style="display:flex;align-items:center;gap:6px">
           <span style="background:rgba(52,211,153,.15);color:var(--grn);padding:2px 8px;
             border-radius:4px;font-size:9px;font-weight:600">✓ ${count.toLocaleString('es-AR')} registros</span>
           <label style="cursor:pointer;background:none;border:1px solid var(--b2);
             border-radius:4px;padding:3px 10px;font-size:9px;color:var(--m1);font-family:var(--sans)">
             ↺ Reemplazar
             <input type="file" accept=".xlsx,.xls,.csv" style="display:none"
               onchange="liqCargar${procLabel}(this)">
           </label>
         </div>`
      : `<label style="cursor:pointer;background:linear-gradient(135deg,var(--acc),var(--cyn));
           color:#fff;border:none;border-radius:4px;padding:4px 14px;font-size:9px;
           font-weight:600;font-family:var(--sans)">
           ＋ Cargar liquidación
           <input type="file" accept=".xlsx,.xls,.csv" style="display:none"
             onchange="liqCargar${procLabel}(this)">
         </label>`}
    <span style="margin-left:auto;font-size:9px;color:var(--m2)">
      Clave de cruce: ${proc === 'FISERV' ? 'N° Lote' : 'N° Autorización'}
    </span>
  </div>`;
}

// ── Estado de cruce activo (para re-render de tabs) ───────────────────
let _liqCruceCache = { fiserv: null, getpos: null, goc: null };

function _liqRenderTab(proc, tab) {
  const c = _liqCruceCache[proc];
  if (!c) return;
  _liqPopulateTab(proc, tab, c);
}

function _liqPopulateTab(proc, tab, cruce) {
  const fmtM = v => _liqFmtARS(Math.abs(v || 0));

  if (proc === 'fiserv') {
    if (tab === 'liq') {
      const rows = cruce.liquidadas.map(x => ({
        estado:    '✓ Liquidada',
        fecha:     x.fila.sky?.fecha || '',
        suc:       x.fila.sky?.suc || '',
        lote:      x.loteN,
        cupon:     x.fila.proc?.ticket || x.fila.proc?.cupon || x.fila.sky?.cupon || '',
        aut:       x.fila.proc?.aut || '',
        monto_sky: fmtM(x.fila.sky?.monto),
        monto_liq: x.liqRows.length
          ? _liqFmtARS(x.liqRows.reduce((s,r)=>s+(r.monto_neto||r.monto_bruto||0),0))
          : '—',
        fecha_liq: x.liqRows[0]?.fecha || '—',
      }));
      _liqRenderTable(`tbl-liq-fiserv-liq`, _liqColsFiserv(), rows);
    }
    if (tab === 'noliq') {
      const rows = cruce.noLiquidadas.map(x => ({
        estado:    x.enLiq === null ? '⚠ Sin archivo liq.' : '✗ No liquidada',
        fecha:     x.fila.sky?.fecha || '',
        suc:       x.fila.sky?.suc || '',
        lote:      x.loteN,
        cupon:     x.fila.proc?.ticket || x.fila.proc?.cupon || x.fila.sky?.cupon || '',
        aut:       x.fila.proc?.aut || '',
        monto_sky: fmtM(x.fila.sky?.monto),
        monto_liq: '—',
        fecha_liq: '—',
      }));
      _liqRenderTable(`tbl-liq-fiserv-noliq`, _liqColsFiserv(), rows);
    }
    if (tab === 'sinconf') {
      const rows = cruce.sinConfirmar.map(f => ({
        estado:    f.estado || 'SIN MATCH',
        fecha:     f.sky?.fecha || '',
        suc:       f.sky?.suc || '',
        lote:      f.sky?.lote || '',
        cupon:     f.sky?.cupon || '',
        aut:       '',
        monto_sky: fmtM(f.sky?.monto),
        monto_liq: '—',
        fecha_liq: '—',
      }));
      _liqRenderTable(`tbl-liq-fiserv-sinconf`, _liqColsFiserv(), rows);
    }
    if (tab === 'extras') {
      const rows = cruce.extras.map(r => ({
        estado:    '⚠ Extra',
        fecha:     '',
        suc:       '',
        lote:      r.lote,
        cupon:     r.cupon || '',
        aut:       r.aut || '',
        monto_sky: '—',
        monto_liq: _liqFmtARS(r.monto_neto || r.monto_bruto || 0),
        fecha_liq: r.fecha || '—',
      }));
      _liqRenderTable(`tbl-liq-fiserv-extras`, _liqColsFiserv(), rows);
    }
  }

  if (proc === 'getpos') {
    if (tab === 'liq') {
      const rows = cruce.liquidadas.map(x => ({
        estado:    '✓ Liquidada',
        fecha:     x.fila.sky?.fecha || '',
        suc:       x.fila.sky?.suc || '',
        aut:       x.autN,
        cupon:     x.fila.proc?.cupon || x.fila.sky?.cupon || '',
        monto_sky: fmtM(x.fila.sky?.monto),
        monto_liq: _liqFmtARS(x.liqRow?.monto_neto || x.liqRow?.monto_bruto || 0),
        fecha_liq: x.liqRow?.fecha || '—',
      }));
      _liqRenderTable(`tbl-liq-getpos-liq`, _liqColsGetpos(), rows);
    }
    if (tab === 'noliq') {
      const rows = cruce.noLiquidadas.map(x => ({
        estado:    x.enLiq === null ? '⚠ Sin archivo liq.' : '✗ No liquidada',
        fecha:     x.fila.sky?.fecha || '',
        suc:       x.fila.sky?.suc || '',
        aut:       x.autN,
        cupon:     x.fila.proc?.cupon || x.fila.sky?.cupon || '',
        monto_sky: fmtM(x.fila.sky?.monto),
        monto_liq: '—',
        fecha_liq: '—',
      }));
      _liqRenderTable(`tbl-liq-getpos-noliq`, _liqColsGetpos(), rows);
    }
    if (tab === 'sinconf') {
      const rows = cruce.sinConfirmar.map(f => ({
        estado:    f.estado || 'SIN MATCH',
        fecha:     f.sky?.fecha || '',
        suc:       f.sky?.suc || '',
        aut:       '',
        cupon:     f.sky?.cupon || '',
        monto_sky: fmtM(f.sky?.monto),
        monto_liq: '—',
        fecha_liq: '—',
      }));
      _liqRenderTable(`tbl-liq-getpos-sinconf`, _liqColsGetpos(), rows);
    }
    if (tab === 'extras') {
      const rows = cruce.extras.map(r => ({
        estado:    '⚠ Extra',
        fecha:     '',
        suc:       '',
        aut:       r.aut,
        cupon:     r.cupon || '',
        monto_sky: '—',
        monto_liq: _liqFmtARS(r.monto_neto || r.monto_bruto || 0),
        fecha_liq: r.fecha || '—',
      }));
      _liqRenderTable(`tbl-liq-getpos-extras`, _liqColsGetpos(), rows);
    }
  }

  if (proc === 'goc') {
    if (tab === 'liq') {
      const rows = cruce.liquidadas.map(x => ({
        estado:    x.fila.estado || '✓ Liquidada',
        fecha:     x.fila.sky?.fecha || '',
        suc:       x.fila.sky?.suc || '',
        orden:     x.orden,
        cupon:     x.fila.sky?.cupon || '',
        monto_sky: fmtM(x.fila.sky?.monto),
        metodo:    x.fila.metodo || '',
      }));
      _liqRenderTable(`tbl-liq-goc-liq`, _liqColsGoC(), rows);
    }
    if (tab === 'noliq') {
      const rows = cruce.noLiquidadas.map(x => ({
        estado:    x.fila.estado || 'SIN MATCH',
        fecha:     x.fila.sky?.fecha || '',
        suc:       x.fila.sky?.suc || '',
        orden:     '',
        cupon:     x.fila.sky?.cupon || '',
        monto_sky: fmtM(x.fila.sky?.monto),
        metodo:    '—',
      }));
      _liqRenderTable(`tbl-liq-goc-noliq`, _liqColsGoC(), rows);
    }
    if (tab === 'extras') {
      const rows = cruce.extras.map(p => ({
        estado:    '⚠ Extra GoC',
        fecha:     p.fechaOrigen || '',
        suc:       p.sucNombre || '',
        orden:     p.orden,
        cupon:     p.refExt || '',
        monto_sky: '—',
        metodo:    p.fuente || '',
      }));
      _liqRenderTable(`tbl-liq-goc-extras`, _liqColsGoC(), rows);
    }
  }
}

function _liqColsFiserv() {
  return [
    { key:'estado',    label:'Estado' },
    { key:'fecha',     label:'Fecha Skylab' },
    { key:'suc',       label:'Suc.' },
    { key:'lote',      label:'Lote' },
    { key:'cupon',     label:'Cupón' },
    { key:'aut',       label:'Autorización' },
    { key:'monto_sky', label:'Monto Skylab', cls:'num' },
    { key:'monto_liq', label:'Monto Liq.', cls:'num' },
    { key:'fecha_liq', label:'Fecha Liq.' },
  ];
}
function _liqColsGetpos() {
  return [
    { key:'estado',    label:'Estado' },
    { key:'fecha',     label:'Fecha Skylab' },
    { key:'suc',       label:'Suc.' },
    { key:'aut',       label:'Autorización' },
    { key:'cupon',     label:'Cupón' },
    { key:'monto_sky', label:'Monto Skylab', cls:'num' },
    { key:'monto_liq', label:'Monto Liq.', cls:'num' },
    { key:'fecha_liq', label:'Fecha Liq.' },
  ];
}
function _liqColsGoC() {
  return [
    { key:'estado',    label:'Estado' },
    { key:'fecha',     label:'Fecha' },
    { key:'suc',       label:'Suc.' },
    { key:'orden',     label:'N° Orden GoC' },
    { key:'cupon',     label:'Cupón Skylab' },
    { key:'monto_sky', label:'Monto', cls:'num' },
    { key:'metodo',    label:'Método' },
  ];
}

// ── RENDER PRINCIPAL ──────────────────────────────────────────────────
function renderModuloLiqFiserv() {
  const panel = document.getElementById('mod-liq-fiserv');
  if (!panel) return;

  if (typeof RESULTADO === 'undefined' || !RESULTADO.length) {
    panel.innerHTML = _liqEmptyState('FISERV', '🏦',
      'Ejecutá el <b>Cruce automático</b> en OPERACIONES antes de liquidar.');
    return;
  }

  const cruce = _cruzarLiqFiserv();
  _liqCruceCache.fiserv = cruce;

  const { liquidadas, noLiquidadas, sinConfirmar, extras,
          montoLiquidado, montoNoLiq, montoExtras, totalOK, tieneLiq } = cruce;
  const pct = totalOK ? +_liqPct(liquidadas.length, totalOK) : 0;

  const fileSection = _liqFileSection('FISERV', 'Fiserv', '🏦', tieneLiq, _LIQ_FIS.length);

  const kpis = [
    { label:'🏦 FISERV OK', val: totalOK.toLocaleString('es-AR'),
      borderColor:'var(--acc)', cls:'cyn',
      sub: _liqFmtARS(montoLiquidado + montoNoLiq) },
    { label:'✓ Liquidadas', val: liquidadas.length.toLocaleString('es-AR'),
      borderColor:'rgba(52,211,153,.3)', cls:'grn',
      sub: _liqFmtARS(montoLiquidado), pct, borderColorBar:'var(--grn)' },
    { label:'✗ No liquidadas', val: noLiquidadas.length.toLocaleString('es-AR'),
      borderColor:'rgba(248,113,113,.3)', cls:'red',
      sub: _liqFmtARS(montoNoLiq),
      pct: totalOK ? +_liqPct(noLiquidadas.length, totalOK) : 0 },
    { label:'? Sin confirmar (Ops.)', val: sinConfirmar.length.toLocaleString('es-AR'),
      borderColor:'rgba(251,191,36,.25)', cls:'yel',
      sub: 'SIN MATCH en OPERACIONES' },
    { label:'⚠ Extras en liq.', val: extras.length.toLocaleString('es-AR'),
      borderColor:'rgba(251,146,60,.25)', cls:'org',
      sub: _liqFmtARS(montoExtras) },
    { label:'% Liquidado', val: `${pct}%`,
      borderColor:'rgba(79,142,247,.3)',
      cls: pct >= 90 ? 'grn' : pct >= 70 ? 'yel' : 'red',
      sub: tieneLiq ? `${liquidadas.length} de ${totalOK}` : 'Cargá archivo de liq.', pct },
  ];

  const tabs = [
    { key:'liq',     label:'✓ Liquidadas',    count:liquidadas.length,
      cntStyle:'background:rgba(52,211,153,.15);color:var(--grn)' },
    { key:'noliq',   label:'✗ No liquidadas',  count:noLiquidadas.length,
      cntStyle:`background:rgba(248,113,113,.15);color:var(--red)${noLiquidadas.length?';font-weight:700':''}` },
    ...(sinConfirmar.length ? [{ key:'sinconf', label:'? Sin confirmar', count:sinConfirmar.length,
      cntStyle:'background:rgba(251,191,36,.1);color:#fbbf24' }] : []),
    ...(extras.length ? [{ key:'extras', label:'⚠ Extras liq.', count:extras.length,
      cntStyle:'background:rgba(251,146,60,.12);color:var(--org)' }] : []),
  ];

  panel.innerHTML = _liqBuildPanel({ id:'fiserv', kpis, tabs, fileSection });

  // Poblar tab activo (primero)
  _liqPopulateTab('fiserv', tabs[0].key, cruce);
}

function renderModuloLiqGetpos() {
  const panel = document.getElementById('mod-liq-getpos');
  if (!panel) return;

  if (typeof RESULTADO === 'undefined' || !RESULTADO.length) {
    panel.innerHTML = _liqEmptyState('GETPOS', '🏧',
      'Ejecutá el <b>Cruce automático</b> en OPERACIONES antes de liquidar.');
    return;
  }

  const cruce = _cruzarLiqGetpos();
  _liqCruceCache.getpos = cruce;

  const { liquidadas, noLiquidadas, sinConfirmar, extras,
          montoLiquidado, montoNoLiq, montoExtras, totalOK, tieneLiq } = cruce;
  const pct = totalOK ? +_liqPct(liquidadas.length, totalOK) : 0;

  const fileSection = _liqFileSection('GETPOS', 'Getpos', '🏧', tieneLiq, _LIQ_GP.length);

  const kpis = [
    { label:'🏧 GETPOS OK', val: totalOK.toLocaleString('es-AR'),
      borderColor:'var(--grn)', cls:'grn',
      sub: _liqFmtARS(montoLiquidado + montoNoLiq) },
    { label:'✓ Liquidadas', val: liquidadas.length.toLocaleString('es-AR'),
      borderColor:'rgba(52,211,153,.3)', cls:'grn',
      sub: _liqFmtARS(montoLiquidado), pct },
    { label:'✗ No liquidadas', val: noLiquidadas.length.toLocaleString('es-AR'),
      borderColor:'rgba(248,113,113,.3)', cls:'red',
      sub: _liqFmtARS(montoNoLiq),
      pct: totalOK ? +_liqPct(noLiquidadas.length, totalOK) : 0 },
    { label:'? Sin confirmar', val: sinConfirmar.length.toLocaleString('es-AR'),
      borderColor:'rgba(251,191,36,.25)', cls:'yel', sub:'SIN MATCH en OPERACIONES' },
    { label:'⚠ Extras en liq.', val: extras.length.toLocaleString('es-AR'),
      borderColor:'rgba(251,146,60,.25)', cls:'org', sub: _liqFmtARS(montoExtras) },
    { label:'% Liquidado', val: `${pct}%`,
      borderColor:'rgba(52,211,153,.3)',
      cls: pct >= 90 ? 'grn' : pct >= 70 ? 'yel' : 'red',
      sub: tieneLiq ? `${liquidadas.length} de ${totalOK}` : 'Cargá archivo de liq.', pct },
  ];

  const tabs = [
    { key:'liq',    label:'✓ Liquidadas',   count:liquidadas.length,
      cntStyle:'background:rgba(52,211,153,.15);color:var(--grn)' },
    { key:'noliq',  label:'✗ No liquidadas', count:noLiquidadas.length,
      cntStyle:`background:rgba(248,113,113,.15);color:var(--red)${noLiquidadas.length?';font-weight:700':''}` },
    ...(sinConfirmar.length ? [{ key:'sinconf', label:'? Sin confirmar', count:sinConfirmar.length,
      cntStyle:'background:rgba(251,191,36,.1);color:#fbbf24' }] : []),
    ...(extras.length ? [{ key:'extras', label:'⚠ Extras liq.', count:extras.length,
      cntStyle:'background:rgba(251,146,60,.12);color:var(--org)' }] : []),
  ];

  panel.innerHTML = _liqBuildPanel({ id:'getpos', kpis, tabs, fileSection });
  _liqPopulateTab('getpos', tabs[0].key, cruce);
}

function renderModuloLiqGoC() {
  const panel = document.getElementById('mod-liq-goc');
  if (!panel) return;

  const filasGoC = typeof RESULTADO !== 'undefined'
    ? RESULTADO.filter(r => r.sky?.esGOCUOTAS) : [];

  if (!filasGoC.length) {
    panel.innerHTML = _liqEmptyState('Go Cuotas', '💳',
      'Cargá el archivo de liquidación GoC en OPERACIONES y ejecutá el cruce.');
    return;
  }

  const cruce = _cruzarLiqGoC();
  _liqCruceCache.goc = cruce;

  const { liquidadas, noLiquidadas, extras,
          montoLiquidado, montoNoLiq, montoExtras, totalOK, tieneLiq } = cruce;
  const pct = totalOK ? +_liqPct(liquidadas.length, totalOK) : 0;

  // Info: GoC usa los archivos ya cargados en OPERACIONES
  const infoSection = `
  <div style="display:flex;align-items:center;gap:12px;padding:10px 18px;
    background:var(--s2);border-bottom:1px solid var(--b1);flex-shrink:0">
    <span style="font-size:14px">💳</span>
    <div style="font-size:10px;font-weight:700;color:var(--txt)">Go Cuotas · Liquidación</div>
    <span style="background:rgba(79,142,247,.12);color:var(--acc);padding:2px 8px;
      border-radius:4px;font-size:9px">ℹ Los archivos GoC cargados en OPERACIONES son las liquidaciones</span>
    <span style="margin-left:auto;font-size:9px;color:var(--m2)">
      Clave: N° Orden GoC
    </span>
  </div>`;

  const kpis = [
    { label:'💳 Operaciones GoC', val: filasGoC.length.toLocaleString('es-AR'),
      borderColor:'rgba(79,142,247,.4)', cls:'cyn',
      sub: _liqFmtARS(montoLiquidado + montoNoLiq) },
    { label:'✓ Liquidadas (con orden)', val: liquidadas.length.toLocaleString('es-AR'),
      borderColor:'rgba(52,211,153,.3)', cls:'grn',
      sub: _liqFmtARS(montoLiquidado), pct },
    { label:'✗ Sin orden GoC', val: noLiquidadas.length.toLocaleString('es-AR'),
      borderColor:'rgba(248,113,113,.3)', cls:'red',
      sub: _liqFmtARS(montoNoLiq),
      pct: totalOK ? +_liqPct(noLiquidadas.length, totalOK) : 0 },
    { label:'⚠ Pagos GoC sin Skylab', val: extras.length.toLocaleString('es-AR'),
      borderColor:'rgba(251,146,60,.25)', cls:'org', sub: _liqFmtARS(montoExtras) },
    { label:'% Cubierto', val: `${pct}%`,
      borderColor:'rgba(79,142,247,.3)',
      cls: pct >= 90 ? 'grn' : pct >= 70 ? 'yel' : 'red',
      sub: `${liquidadas.length} de ${totalOK}`, pct },
  ];

  const tabs = [
    { key:'liq',   label:'✓ Con orden GoC',    count:liquidadas.length,
      cntStyle:'background:rgba(52,211,153,.15);color:var(--grn)' },
    { key:'noliq', label:'✗ Sin orden GoC', count:noLiquidadas.length,
      cntStyle:`background:rgba(248,113,113,.15);color:var(--red)${noLiquidadas.length?';font-weight:700':''}` },
    ...(extras.length ? [{ key:'extras', label:'⚠ Extras GoC', count:extras.length,
      cntStyle:'background:rgba(251,146,60,.12);color:var(--org)' }] : []),
  ];

  panel.innerHTML = _liqBuildPanel({ id:'goc', kpis, tabs, fileSection: infoSection });
  _liqPopulateTab('goc', tabs[0].key, cruce);
}

function _liqEmptyState(nombre, icon, msg) {
  return `<div style="display:flex;flex-direction:column;align-items:center;justify-content:center;
    height:100%;gap:14px;color:var(--m2);text-align:center;padding:40px">
    <div style="font-size:42px;opacity:.15">${icon}</div>
    <div style="font-size:16px;font-weight:700;color:var(--txt);opacity:.4">${nombre}</div>
    <p style="font-size:10px;max-width:400px;line-height:1.8">${msg}</p>
  </div>`;
}

// ── Actualizar todos los paneles de liquidación ───────────────────────
function renderModuloLiquidaciones() {
  renderModuloLiqFiserv();
  renderModuloLiqGetpos();
  renderModuloLiqGoC();
}
