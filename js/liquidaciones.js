// ═══════════════════════════════════════════════════════════════════
// LIQUIDACIONES.JS — Cruce de operaciones confirmadas vs liquidaciones
// ═══════════════════════════════════════════════════════════════════
//
// Flujo:
//  OPERACIONES: cobrado (FISERV/GETPOS/GoC) vs facturado (Skylab) → RESULTADO
//  LIQUIDACIONES: RESULTADO OK vs lo que la procesadora declaró pagar
//
// Archivos reales:
//  • FISERV + GETPOS → LIQUIDACIONES.xlsx (formato FISERV clásico)
//      Columnas clave: Nro de Lote, Nro de Cupón, Código Autorización, Importe Venta
//      Clave FISERV = Lote + Cupón  |  Clave GETPOS = Código Autorización
//
//  • Go Cuotas → mismo CSV que sube el usuario en OPERACIONES (ya es la liq.)
//      Columnas: Número de Orden, Importe, Fecha Pago, Sucursal Nombre
//      Usa los datos ya cargados en _GOC_PAGOS / _GOC_CELULAR
// ═══════════════════════════════════════════════════════════════════

// ── Estado ──────────────────────────────────────────────────────────
// Un solo array cubre FISERV y GETPOS (mismo archivo de cupones liquidados)
let _LIQ_CUPONES = [];

// ── Helpers ─────────────────────────────────────────────────────────
function _liqGCol(row, ...cands) {
  for (const c of cands) {
    const v = row[c];
    if (v !== undefined && v !== null && v !== '') return v;
    // case-insensitive fallback
    const key = Object.keys(row).find(k =>
      k.trim().toLowerCase() === c.trim().toLowerCase());
    if (key !== undefined && row[key] !== null && row[key] !== undefined && row[key] !== '') return row[key];
  }
  return null;
}

function _liqNorm(v) {
  if (v === null || v === undefined) return '0';
  return String(v).trim().replace(/^0+/, '') || '0';
}

function _liqNormAut(v) {
  // Autorizaciones pueden tener espacios al final (p.ej. "756870  ")
  return _liqNorm(String(v ?? '').trim());
}

function _liqMonto(v) {
  if (v === null || v === undefined || v === '') return 0;
  // Formatos: "$ 55.275,31", "55275.31", "55.275,31"
  const s = String(v)
    .replace(/[^0-9,.-]/g, '')   // quitar $ espacios etc.
    .replace(/\.(?=\d{3}(?:[,\.]|$))/g, '') // quitar puntos miles
    .replace(',', '.');
  return parseFloat(s) || 0;
}

function _liqFecha(v) {
  if (!v) return '';
  const s = String(v).trim();
  // YYYY-MM-DD (ya normalizado)
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  // DD/MM/YYYY
  const m1 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m1) return `${m1[3]}-${m1[2].padStart(2,'0')}-${m1[1].padStart(2,'0')}`;
  // DD-MM-YYYY
  const m2 = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})/);
  if (m2) return `${m2[3]}-${m2[2].padStart(2,'0')}-${m2[1].padStart(2,'0')}`;
  return s;
}

// ── PARSER: Cupones Liquidados (LIQUIDACIONES.xlsx) ─────────────────
// Formato FISERV clásico — un cupón por fila con lote+cupón+autorización
// Este mismo archivo se usa para FISERV (lote+cupón) y GETPOS (autorización)
function parseLiqCupones(wb) {
  const ws = wb.Sheets[wb.SheetNames[0]];
  // raw:true → números como JS number (cupón, lote, importe llegan numéricos)
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: true });
  if (!rows.length) { console.warn('[LIQ] Archivo sin datos'); return []; }

  // Debug primeros headers
  console.log('[LIQ] Columnas:', Object.keys(rows[0]).slice(0, 20).join(' | '));

  const out = [];
  for (const r of rows) {
    const tipo = String(_liqGCol(r,'Tipo operacion','Tipo Operacion','Tipo','TIPO') || '').toUpperCase().trim();
    // Incluir solo ventas (skipear devoluciones/anulaciones para el cruce principal)
    const esDevolucion = ['DEVOLUCION','DEVOLUCIÓN','ANULACION','ANULACIÓN','REVERSO'].some(t => tipo.includes(t));
    if (esDevolucion) continue;

    // Lote
    const loteRaw = _liqGCol(r,'Nro de Lote','Nro Lote','N° de Lote','LOTE','Lote','nro_lote');
    const lote = typeof loteRaw === 'number' ? String(Math.round(loteRaw)) : _liqNorm(loteRaw);

    // Cupón (Nro de Cupón es numérico en el xlsx)
    const cupRaw = _liqGCol(r,'Nro de Cupón','Nro de Cupon','Cupon','Cupón','Ticket','ticket','Nro Cupon');
    const cupon = typeof cupRaw === 'number' ? String(Math.round(cupRaw)) : _liqNorm(cupRaw);

    // Autorización (puede tener espacios: "756870  ")
    const autRaw = _liqGCol(r,'Código Autorización','Codigo Autorizacion','Cód. Autorización',
                              'Autorización','Autorizacion','Aut','aut');
    const aut = _liqNormAut(autRaw);

    // Skipear filas sin datos útiles
    if (lote === '0' && cupon === '0' && aut === '0') continue;

    // Importe
    const montoRaw = _liqGCol(r,'Importe Venta','Importe','Monto','importe');
    const monto = typeof montoRaw === 'number' ? montoRaw : _liqMonto(montoRaw);

    out.push({
      lote,
      cupon,
      aut,
      monto,
      equipo:     String(_liqGCol(r,'Nro Equipo','Equipo','Terminal','terminal') || '').trim(),
      liq_id:     String(_liqGCol(r,'Nro Liquidación','Nro Liquidacion','Liquidacion') || '').trim(),
      fecha_venta:_liqFecha(_liqGCol(r,'Fecha Venta','Fecha','fecha')),
      fecha_pago: _liqFecha(_liqGCol(r,'Fecha Pago','FechaPago')),
      tarjeta:    String(_liqGCol(r,'Tarjeta','tarjeta') || '').trim(),
      tipo,
    });
  }

  console.log(`[LIQ] Parseados: ${out.length} cupones liquidados`);
  // Debug muestra (primeras 3 filas)
  if (out.length > 0) {
    console.debug('[LIQ] Muestra:', out.slice(0,3).map(r =>
      `lote=${r.lote} cupon=${r.cupon} aut=${r.aut} monto=${r.monto}`).join(' | '));
  }
  return out;
}

// ── UPLOAD handler (único — cubre FISERV y GETPOS) ───────────────────
function liqCargarCupones(input) {
  const file = input.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, { type:'array' });
      _LIQ_CUPONES = parseLiqCupones(wb);
      if (!_LIQ_CUPONES.length) {
        typeof _showToast === 'function'
          ? _showToast('⚠ No se encontraron cupones válidos — verificá las columnas del archivo')
          : alert('Sin cupones válidos');
      } else {
        typeof _showToast === 'function'
          ? _showToast(`✓ ${_LIQ_CUPONES.length.toLocaleString('es-AR')} cupones liquidados cargados`)
          : null;
      }
      // Refrescar ambos paneles
      renderModuloLiqFiserv();
      renderModuloLiqGetpos();
    } catch(err) {
      alert('Error leyendo archivo: ' + err.message);
      console.error('[LIQ]', err);
    }
    input.value = '';
  };
  reader.readAsArrayBuffer(file);
}

// Alias para compatibilidad (ambos paneles llaman al mismo handler)
function liqCargarFiserv(input)  { liqCargarCupones(input); }
function liqCargarGetpos(input) { liqCargarCupones(input); }

// ── ÍNDICES de liquidación (se recalculan en cada cruce) ─────────────
function _liqBuildIndexes() {
  const byLoteCupon = {};   // `${lote}-${cupon}` → fila
  const byAut       = {};   // `${aut}` → fila

  for (const r of _LIQ_CUPONES) {
    // Índice lote+cupón (para FISERV)
    if (r.lote !== '0' && r.cupon !== '0') {
      const k = `${r.lote}-${r.cupon}`;
      byLoteCupon[k] = r;
    }
    // Índice autorización (para GETPOS y fallback FISERV)
    if (r.aut && r.aut !== '0') {
      byAut[r.aut] = r;
    }
  }
  return { byLoteCupon, byAut };
}

// ── CRUCE: FISERV ────────────────────────────────────────────────────
// Clave: Lote + Cupón (con fallback por autorización)
function _cruzarLiqFiserv() {
  if (typeof RESULTADO === 'undefined' || !RESULTADO.length) return null;

  const filasFis = RESULTADO.filter(r => !r.sky?.esGETPos && !r.sky?.esGOCUOTAS);
  const { byLoteCupon, byAut } = _liqBuildIndexes();
  const autsUsados = new Set();  // para calcular extras

  const liquidadas    = [];
  const noLiquidadas  = [];
  const sinConfirmar  = [];

  for (const fila of filasFis) {
    const esOK = fila.estado?.startsWith('OK');
    if (!esOK) { sinConfirmar.push(fila); continue; }

    const lote  = _liqNorm(fila.proc?.lote   || fila.sky?.lote   || '');
    const cupon = _liqNorm(fila.proc?.ticket  || fila.proc?.cupon || fila.sky?.cupon || '');
    const aut   = _liqNormAut(fila.proc?.aut  || '');

    if (!_LIQ_CUPONES.length) {
      noLiquidadas.push({ fila, lote, cupon, aut, enLiq: null, liqRow: null });
      continue;
    }

    // Buscar en liquidación: primero por lote+cupón, luego por autorización
    const liqRow = byLoteCupon[`${lote}-${cupon}`] || byAut[aut];

    if (liqRow) {
      if (liqRow.aut && liqRow.aut !== '0') autsUsados.add(liqRow.aut);
      liquidadas.push({ fila, lote, cupon, aut, liqRow });
    } else {
      noLiquidadas.push({ fila, lote, cupon, aut, enLiq: false, liqRow: null });
    }
  }

  // Extras: cupones de la liq que ninguna fila de RESULTADO (FISERV) reclamó
  // Solo contar extras de FISERV (filas con lote+cupón)
  const extras = _LIQ_CUPONES.filter(r => {
    if (r.lote === '0' || r.cupon === '0') return false;
    const k = `${r.lote}-${r.cupon}`;
    return !byLoteCupon[k] || !liquidadas.some(x =>
      `${x.lote}-${x.cupon}` === k);
  });

  const montoLiquidado = liquidadas.reduce((s,x) => s + (x.liqRow?.monto || Math.abs(x.fila.sky?.monto||0)), 0);
  const montoNoLiq     = noLiquidadas.reduce((s,x) => s + Math.abs(x.fila.sky?.monto||0), 0);
  const montoExtras    = extras.reduce((s,r) => s + r.monto, 0);

  return {
    liquidadas, noLiquidadas, sinConfirmar, extras,
    montoLiquidado, montoNoLiq, montoExtras,
    totalOK: liquidadas.length + noLiquidadas.length,
    tieneLiq: _LIQ_CUPONES.length > 0,
  };
}

// ── CRUCE: GETPOS ────────────────────────────────────────────────────
// Clave: Código Autorización
function _cruzarLiqGetpos() {
  if (typeof RESULTADO === 'undefined' || !RESULTADO.length) return null;

  const filasGP = RESULTADO.filter(r => r.sky?.esGETPos && !r.sky?.esGOCUOTAS);
  const { byAut } = _liqBuildIndexes();

  const liquidadas   = [];
  const noLiquidadas = [];
  const sinConfirmar = [];

  for (const fila of filasGP) {
    const esOK = fila.estado?.startsWith('OK');
    if (!esOK) { sinConfirmar.push(fila); continue; }

    const aut = _liqNormAut(fila.proc?.aut || '');

    if (!_LIQ_CUPONES.length) {
      noLiquidadas.push({ fila, aut, enLiq: null, liqRow: null });
      continue;
    }

    const liqRow = byAut[aut];
    if (liqRow) {
      liquidadas.push({ fila, aut, liqRow });
    } else {
      noLiquidadas.push({ fila, aut, enLiq: false, liqRow: null });
    }
  }

  // Extras GETPOS: autorizaciones en la liq sin lote (probablemente GETPOS)
  // que no matchearon con ningún RESULTADO GETPOS
  const autsUsadosGP = new Set(liquidadas.map(x => x.aut));
  const extras = _LIQ_CUPONES.filter(r => {
    // Filas sin lote son candidatas GETPOS
    const sinLote = r.lote === '0' || !r.lote;
    return sinLote && r.aut !== '0' && !autsUsadosGP.has(r.aut);
  });

  const montoLiquidado = liquidadas.reduce((s,x) => s + (x.liqRow?.monto || Math.abs(x.fila.sky?.monto||0)), 0);
  const montoNoLiq     = noLiquidadas.reduce((s,x) => s + Math.abs(x.fila.sky?.monto||0), 0);
  const montoExtras    = extras.reduce((s,r) => s + r.monto, 0);

  return {
    liquidadas, noLiquidadas, sinConfirmar, extras,
    montoLiquidado, montoNoLiq, montoExtras,
    totalOK: liquidadas.length + noLiquidadas.length,
    tieneLiq: _LIQ_CUPONES.length > 0,
  };
}

// ── CRUCE: Go Cuotas ─────────────────────────────────────────────────
// Usa datos ya cargados en _GOC_PAGOS / _GOC_CELULAR (archivos de OPERACIONES).
// El CSV de GoC que se sube en OPERACIONES ES la liquidación.
function _cruzarLiqGoC() {
  if (typeof RESULTADO === 'undefined' || !RESULTADO.length) return null;

  const filasGoC = RESULTADO.filter(r => r.sky?.esGOCUOTAS);
  const _gocAll  = [
    ...(typeof _GOC_PAGOS   !== 'undefined' ? _GOC_PAGOS   : []),
    ...(typeof _GOC_CELULAR !== 'undefined' ? _GOC_CELULAR : []),
  ];

  const ordenesUsadas = new Set();
  const liquidadas    = [];
  const noLiquidadas  = [];

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

  // Extras: órdenes GoC sin match en Skylab
  const extras = _gocAll.filter(p => {
    const o = _liqNorm(p.orden || '');
    return o && o !== '0' && !ordenesUsadas.has(o);
  });

  const montoLiquidado = liquidadas.reduce((s,x) => s + Math.abs(x.fila.sky?.monto||0), 0);
  const montoNoLiq     = noLiquidadas.reduce((s,x) => s + Math.abs(x.fila.sky?.monto||0), 0);
  const montoExtras    = extras.reduce((s,p) => s + (p.importe||0), 0);

  return {
    liquidadas, noLiquidadas, sinConfirmar:[], extras,
    montoLiquidado, montoNoLiq, montoExtras,
    totalOK: liquidadas.length + noLiquidadas.length,
    tieneLiq: _gocAll.length > 0,
  };
}

// ── RENDER helpers ────────────────────────────────────────────────────
function _liqFmtARS(v) {
  return typeof fmtARS === 'function' ? fmtARS(Math.abs(v || 0)) : `$${(+v||0).toFixed(2)}`;
}

function _liqPct(n, d) {
  if (!d) return '0.0';
  return ((n / d) * 100).toFixed(1);
}

function _liqPBar(pct, color) {
  return `<div style="width:100%;height:3px;background:var(--b1);border-radius:2px;margin-top:4px">
    <div style="width:${Math.min(100, +pct)}%;height:100%;background:${color};border-radius:2px"></div>
  </div>`;
}

// ── Sección upload de archivo ─────────────────────────────────────────
function _liqFileSection(proc, procLabel, icon, hasFile, count) {
  const subtxt = proc === 'FISERV'
    ? 'Clave: Lote + Cupón · también cubre GETPOS (por Autorización)'
    : 'Clave: Código Autorización · archivo compartido con FISERV';

  return `
  <div style="display:flex;align-items:center;gap:12px;padding:10px 18px;
    background:var(--s2);border-bottom:1px solid var(--b1);flex-shrink:0;flex-wrap:wrap">
    <span style="font-size:14px">${icon}</span>
    <div style="font-size:10px;font-weight:700;color:var(--txt)">
      ${procLabel} · Cupones Liquidados
    </div>
    ${hasFile
      ? `<span style="background:rgba(52,211,153,.15);color:var(--grn);padding:2px 8px;
           border-radius:4px;font-size:9px;font-weight:600">
           ✓ ${count.toLocaleString('es-AR')} cupones cargados
         </span>
         <label style="cursor:pointer;background:none;border:1px solid var(--b2);
           border-radius:4px;padding:3px 10px;font-size:9px;color:var(--m1);font-family:var(--sans)">
           ↺ Reemplazar
           <input type="file" accept=".xlsx,.xls" style="display:none"
             onchange="liqCargarCupones(this)">
         </label>`
      : `<label style="cursor:pointer;background:linear-gradient(135deg,var(--acc),var(--cyn));
           color:#fff;border:none;border-radius:4px;padding:4px 14px;font-size:9px;
           font-weight:600;font-family:var(--sans)">
           ＋ Cargar LIQUIDACIONES.xlsx
           <input type="file" accept=".xlsx,.xls" style="display:none"
             onchange="liqCargarCupones(this)">
         </label>`}
    <span style="font-size:9px;color:var(--m2);margin-left:auto">${subtxt}</span>
  </div>`;
}

// ── Panel genérico (KPIs + tabs + tablas) ────────────────────────────
function _liqBuildPanel(opts) {
  const { id, kpis, tabs, fileSection } = opts;
  return `
  <div style="display:flex;flex-direction:column;flex:1;min-height:0;overflow:hidden">
    ${fileSection || ''}

    <!-- KPIs -->
    <div style="display:grid;grid-template-columns:repeat(${kpis.length},1fr);gap:8px;
      padding:12px 18px;background:var(--s1);border-bottom:1px solid var(--b1);flex-shrink:0">
      ${kpis.map(k=>`
        <div class="dif-kpi" style="border-color:${k.bc||'var(--b2)'}">
          <div class="dif-kpi-lbl">${k.label}</div>
          <div class="dif-kpi-val ${k.cls||''}">${k.val}</div>
          <div style="font-size:8px;color:var(--m2)">${k.sub||''}</div>
          ${k.pct!==undefined ? _liqPBar(k.pct, k.bc||'var(--acc)') : ''}
        </div>`).join('')}
    </div>

    <!-- Tab strip -->
    <div class="tab-strip" id="tab-strip-liq-${id}">
      ${tabs.map((t,i)=>`
        <button class="tb${i===0?' active':''}" data-tab="${t.key}"
          onclick="showTab('liqtab-${id}-${t.key}','tab-strip-liq-${id}',this);_liqRenderTab('${id}','${t.key}')">
          ${t.label} <span class="cnt" style="${t.cs||''}">${t.n}</span>
        </button>`).join('')}
    </div>

    <!-- Tab bodies -->
    ${tabs.map((t,i)=>`
      <div class="tab-body${i===0?' active':''}" id="liqtab-${id}-${t.key}"
        style="flex-direction:column;flex:1;min-height:0">
        <div class="tbl-wrap">
          <table id="tbl-liq-${id}-${t.key}"><thead></thead><tbody></tbody></table>
        </div>
      </div>`).join('')}
  </div>`;
}

// Pobla una tabla con cols + rows
function _liqRenderTable(tableId, cols, rows) {
  const tbl = document.getElementById(tableId);
  if (!tbl) return;
  tbl.querySelector('thead').innerHTML =
    `<tr>${cols.map(c=>`<th>${c.label}</th>`).join('')}</tr>`;
  tbl.querySelector('tbody').innerHTML = rows.length
    ? rows.map(row=>`<tr style="${row._tr_style||''}">${cols.map(c=>`<td${c.cls?` class="${c.cls}"`:''}>${row[c.key]??'—'}</td>`).join('')}</tr>`).join('')
    : `<tr><td colspan="${cols.length}" style="text-align:center;padding:32px;color:var(--m2);font-style:italic">Sin registros</td></tr>`;
}

// ── Cache de cruces para re-render de tabs ───────────────────────────
const _liqCache = { fiserv:null, getpos:null, goc:null };

function _liqRenderTab(proc, tab) {
  const c = _liqCache[proc];
  if (!c) return;
  _liqPopTab(proc, tab, c);
}

// ── Formatea cada tab ────────────────────────────────────────────────
function _liqPopTab(proc, tab, cruce) {
  const fmtM = v => _liqFmtARS(Math.abs(v||0));

  /* ── FISERV ──────────────────────────────────────────────────── */
  if (proc === 'fiserv') {
    const colsFis = [
      { key:'_estado',   label:'Estado' },
      { key:'fecha',     label:'Fecha Skylab' },
      { key:'suc',       label:'Suc.' },
      { key:'lote',      label:'Lote' },
      { key:'cupon',     label:'Cupón' },
      { key:'aut',       label:'Autorización' },
      { key:'monto_sky', label:'Monto Skylab',  cls:'num' },
      { key:'monto_liq', label:'Monto Liq.',    cls:'num' },
      { key:'fecha_pago',label:'Fecha Pago Liq.' },
      { key:'liq_id',    label:'N° Liq.' },
    ];

    if (tab === 'liq') {
      _liqRenderTable(`tbl-liq-fiserv-liq`, colsFis,
        cruce.liquidadas.map(x => ({
          _estado:   '✓ Liquidada',
          fecha:     x.fila.sky?.fecha || '',
          suc:       x.fila.sky?.suc  || '',
          lote:      x.lote,
          cupon:     x.cupon,
          aut:       x.aut,
          monto_sky: fmtM(x.fila.sky?.monto),
          monto_liq: fmtM(x.liqRow?.monto),
          fecha_pago:x.liqRow?.fecha_pago || '',
          liq_id:    x.liqRow?.liq_id || '',
        })));
    }
    if (tab === 'noliq') {
      _liqRenderTable(`tbl-liq-fiserv-noliq`, colsFis,
        cruce.noLiquidadas.map(x => ({
          _estado:   x.enLiq === null ? '⚠ Sin archivo' : '✗ No liquidada',
          fecha:     x.fila.sky?.fecha || '',
          suc:       x.fila.sky?.suc  || '',
          lote:      x.lote,
          cupon:     x.cupon,
          aut:       x.aut,
          monto_sky: fmtM(x.fila.sky?.monto),
          monto_liq: '—',
          fecha_pago:'—',
          liq_id:    '—',
        })));
    }
    if (tab === 'sinconf') {
      _liqRenderTable(`tbl-liq-fiserv-sinconf`, colsFis,
        cruce.sinConfirmar.map(f => ({
          _estado:   f.estado || 'SIN MATCH',
          fecha:     f.sky?.fecha || '',
          suc:       f.sky?.suc  || '',
          lote:      f.sky?.lote || '',
          cupon:     f.sky?.cupon|| '',
          aut:       '',
          monto_sky: fmtM(f.sky?.monto),
          monto_liq: '—',
          fecha_pago:'—',
          liq_id:    '—',
        })));
    }
    if (tab === 'extras') {
      _liqRenderTable(`tbl-liq-fiserv-extras`, colsFis,
        cruce.extras.map(r => ({
          _estado:   '⚠ Extra',
          fecha:     r.fecha_venta || '',
          suc:       '',
          lote:      r.lote,
          cupon:     r.cupon,
          aut:       r.aut,
          monto_sky: '—',
          monto_liq: fmtM(r.monto),
          fecha_pago:r.fecha_pago || '—',
          liq_id:    r.liq_id || '—',
        })));
    }
  }

  /* ── GETPOS ─────────────────────────────────────────────────── */
  if (proc === 'getpos') {
    const colsGP = [
      { key:'_estado',   label:'Estado' },
      { key:'fecha',     label:'Fecha Skylab' },
      { key:'suc',       label:'Suc.' },
      { key:'aut',       label:'Autorización' },
      { key:'cupon',     label:'Cupón' },
      { key:'monto_sky', label:'Monto Skylab', cls:'num' },
      { key:'monto_liq', label:'Monto Liq.',   cls:'num' },
      { key:'fecha_pago',label:'Fecha Pago' },
    ];

    if (tab === 'liq') {
      _liqRenderTable(`tbl-liq-getpos-liq`, colsGP,
        cruce.liquidadas.map(x => ({
          _estado:   '✓ Liquidada',
          fecha:     x.fila.sky?.fecha || '',
          suc:       x.fila.sky?.suc  || '',
          aut:       x.aut,
          cupon:     x.fila.proc?.cupon || x.fila.sky?.cupon || '',
          monto_sky: fmtM(x.fila.sky?.monto),
          monto_liq: fmtM(x.liqRow?.monto),
          fecha_pago:x.liqRow?.fecha_pago || '',
        })));
    }
    if (tab === 'noliq') {
      _liqRenderTable(`tbl-liq-getpos-noliq`, colsGP,
        cruce.noLiquidadas.map(x => ({
          _estado:   x.enLiq === null ? '⚠ Sin archivo' : '✗ No liquidada',
          fecha:     x.fila.sky?.fecha || '',
          suc:       x.fila.sky?.suc  || '',
          aut:       x.aut,
          cupon:     x.fila.proc?.cupon || x.fila.sky?.cupon || '',
          monto_sky: fmtM(x.fila.sky?.monto),
          monto_liq: '—',
          fecha_pago:'—',
        })));
    }
    if (tab === 'sinconf') {
      _liqRenderTable(`tbl-liq-getpos-sinconf`, colsGP,
        cruce.sinConfirmar.map(f => ({
          _estado:   f.estado || 'SIN MATCH',
          fecha:     f.sky?.fecha || '',
          suc:       f.sky?.suc  || '',
          aut:       '',
          cupon:     f.sky?.cupon|| '',
          monto_sky: fmtM(f.sky?.monto),
          monto_liq: '—',
          fecha_pago:'—',
        })));
    }
    if (tab === 'extras') {
      _liqRenderTable(`tbl-liq-getpos-extras`, colsGP,
        cruce.extras.map(r => ({
          _estado:   '⚠ Extra',
          fecha:     r.fecha_venta || '',
          suc:       '',
          aut:       r.aut,
          cupon:     r.cupon,
          monto_sky: '—',
          monto_liq: fmtM(r.monto),
          fecha_pago:r.fecha_pago || '',
        })));
    }
  }

  /* ── GoC ──────────────────────────────────────────────────────── */
  if (proc === 'goc') {
    const colsGoC = [
      { key:'_estado',   label:'Estado' },
      { key:'fecha',     label:'Fecha Skylab' },
      { key:'suc',       label:'Sucursal' },
      { key:'orden',     label:'N° Orden GoC' },
      { key:'cupon_sky', label:'Cupón Skylab' },
      { key:'monto_sky', label:'Monto Skylab', cls:'num' },
      { key:'fecha_pago',label:'Fecha Pago GoC' },
      { key:'metodo',    label:'Método' },
    ];

    if (tab === 'liq') {
      _liqRenderTable(`tbl-liq-goc-liq`, colsGoC,
        cruce.liquidadas.map(x => ({
          _estado:   x.fila.estado || '✓ Con orden',
          fecha:     x.fila.sky?.fecha || '',
          suc:       x.fila.sky?.suc  || '',
          orden:     x.orden,
          cupon_sky: x.fila.sky?.cupon || '',
          monto_sky: fmtM(x.fila.sky?.monto),
          fecha_pago:'—',
          metodo:    x.fila.metodo || '',
        })));
    }
    if (tab === 'noliq') {
      _liqRenderTable(`tbl-liq-goc-noliq`, colsGoC,
        cruce.noLiquidadas.map(x => ({
          _estado:   x.fila.estado || 'SIN MATCH',
          fecha:     x.fila.sky?.fecha || '',
          suc:       x.fila.sky?.suc  || '',
          orden:     '—',
          cupon_sky: x.fila.sky?.cupon || '',
          monto_sky: fmtM(x.fila.sky?.monto),
          fecha_pago:'—',
          metodo:    '—',
        })));
    }
    if (tab === 'extras') {
      _liqRenderTable(`tbl-liq-goc-extras`, colsGoC,
        cruce.extras.map(p => ({
          _estado:   '⚠ Pago sin Skylab',
          fecha:     p.fechaOrigen || '',
          suc:       p.sucNombre || '',
          orden:     p.orden,
          cupon_sky: p.refExt || '—',
          monto_sky: fmtM(p.importe),
          fecha_pago:p.fechaPago || '',
          metodo:    p.fuente || '',
        })));
    }
  }
}

// ── RENDER MÓDULOS PRINCIPALES ────────────────────────────────────────
function renderModuloLiqFiserv() {
  const panel = document.getElementById('mod-liq-fiserv');
  if (!panel) return;

  if (typeof RESULTADO === 'undefined' || !RESULTADO.length) {
    panel.innerHTML = _liqEmpty('FISERV', '🏦',
      'Ejecutá el <b>Cruce automático</b> en OPERACIONES primero.');
    return;
  }

  const cruce = _cruzarLiqFiserv();
  _liqCache.fiserv = cruce;
  const { liquidadas, noLiquidadas, sinConfirmar, extras,
          montoLiquidado, montoNoLiq, montoExtras, totalOK, tieneLiq } = cruce;
  const pct = totalOK ? +_liqPct(liquidadas.length, totalOK) : 0;

  const kpis = [
    { label:'🏦 FISERV OK',       val:totalOK.toLocaleString('es-AR'),
      bc:'var(--acc)', cls:'cyn', sub:_liqFmtARS(montoLiquidado+montoNoLiq) },
    { label:'✓ Liquidadas',       val:liquidadas.length.toLocaleString('es-AR'),
      bc:'rgba(52,211,153,.3)', cls:'grn', sub:_liqFmtARS(montoLiquidado), pct },
    { label:'✗ No liquidadas',    val:noLiquidadas.length.toLocaleString('es-AR'),
      bc:'rgba(248,113,113,.3)', cls:'red', sub:_liqFmtARS(montoNoLiq),
      pct:totalOK ? +_liqPct(noLiquidadas.length,totalOK) : 0 },
    { label:'? Sin confirmar',    val:sinConfirmar.length.toLocaleString('es-AR'),
      bc:'rgba(251,191,36,.25)', cls:'yel', sub:'SIN MATCH en OPERACIONES' },
    { label:'⚠ Extras en liq.',  val:extras.length.toLocaleString('es-AR'),
      bc:'rgba(251,146,60,.25)', cls:'org', sub:_liqFmtARS(montoExtras) },
    { label:'% Liquidado',        val:`${pct}%`,
      bc:'rgba(79,142,247,.3)',
      cls: pct>=90?'grn':pct>=70?'yel':'red',
      sub: tieneLiq ? `${liquidadas.length} de ${totalOK}` : 'Cargá LIQUIDACIONES.xlsx', pct },
  ];

  const tabs = [
    { key:'liq',     label:'✓ Liquidadas',   n:liquidadas.length,
      cs:'background:rgba(52,211,153,.15);color:var(--grn)' },
    { key:'noliq',   label:'✗ No liquidadas', n:noLiquidadas.length,
      cs:`background:rgba(248,113,113,.15);color:var(--red)${noLiquidadas.length?';font-weight:700':''}` },
    ...(sinConfirmar.length ? [{ key:'sinconf', label:'? Sin confirmar', n:sinConfirmar.length,
      cs:'background:rgba(251,191,36,.1);color:#fbbf24' }] : []),
    ...(extras.length ? [{ key:'extras', label:'⚠ Extras', n:extras.length,
      cs:'background:rgba(251,146,60,.12);color:var(--org)' }] : []),
  ];

  panel.innerHTML = _liqBuildPanel({
    id:'fiserv', kpis, tabs,
    fileSection: _liqFileSection('FISERV','Fiserv','🏦', tieneLiq, _LIQ_CUPONES.length),
  });
  _liqPopTab('fiserv', tabs[0].key, cruce);
}

function renderModuloLiqGetpos() {
  const panel = document.getElementById('mod-liq-getpos');
  if (!panel) return;

  if (typeof RESULTADO === 'undefined' || !RESULTADO.length) {
    panel.innerHTML = _liqEmpty('GETPOS', '🏧',
      'Ejecutá el <b>Cruce automático</b> en OPERACIONES primero.');
    return;
  }

  const cruce = _cruzarLiqGetpos();
  _liqCache.getpos = cruce;
  const { liquidadas, noLiquidadas, sinConfirmar, extras,
          montoLiquidado, montoNoLiq, montoExtras, totalOK, tieneLiq } = cruce;
  const pct = totalOK ? +_liqPct(liquidadas.length, totalOK) : 0;

  const kpis = [
    { label:'🏧 GETPOS OK',       val:totalOK.toLocaleString('es-AR'),
      bc:'var(--grn)', cls:'grn', sub:_liqFmtARS(montoLiquidado+montoNoLiq) },
    { label:'✓ Liquidadas',       val:liquidadas.length.toLocaleString('es-AR'),
      bc:'rgba(52,211,153,.3)', cls:'grn', sub:_liqFmtARS(montoLiquidado), pct },
    { label:'✗ No liquidadas',    val:noLiquidadas.length.toLocaleString('es-AR'),
      bc:'rgba(248,113,113,.3)', cls:'red', sub:_liqFmtARS(montoNoLiq),
      pct:totalOK ? +_liqPct(noLiquidadas.length,totalOK) : 0 },
    { label:'? Sin confirmar',    val:sinConfirmar.length.toLocaleString('es-AR'),
      bc:'rgba(251,191,36,.25)', cls:'yel', sub:'SIN MATCH en OPERACIONES' },
    { label:'⚠ Extras en liq.',  val:extras.length.toLocaleString('es-AR'),
      bc:'rgba(251,146,60,.25)', cls:'org', sub:_liqFmtARS(montoExtras) },
    { label:'% Liquidado',        val:`${pct}%`,
      bc:'rgba(52,211,153,.3)',
      cls: pct>=90?'grn':pct>=70?'yel':'red',
      sub: tieneLiq ? `${liquidadas.length} de ${totalOK}` : 'Cargá LIQUIDACIONES.xlsx', pct },
  ];

  const tabs = [
    { key:'liq',   label:'✓ Liquidadas',   n:liquidadas.length,
      cs:'background:rgba(52,211,153,.15);color:var(--grn)' },
    { key:'noliq', label:'✗ No liquidadas', n:noLiquidadas.length,
      cs:`background:rgba(248,113,113,.15);color:var(--red)${noLiquidadas.length?';font-weight:700':''}` },
    ...(sinConfirmar.length ? [{ key:'sinconf', label:'? Sin confirmar', n:sinConfirmar.length,
      cs:'background:rgba(251,191,36,.1);color:#fbbf24' }] : []),
    ...(extras.length ? [{ key:'extras', label:'⚠ Extras', n:extras.length,
      cs:'background:rgba(251,146,60,.12);color:var(--org)' }] : []),
  ];

  panel.innerHTML = _liqBuildPanel({
    id:'getpos', kpis, tabs,
    fileSection: _liqFileSection('GETPOS','Getpos','🏧', tieneLiq, _LIQ_CUPONES.length),
  });
  _liqPopTab('getpos', tabs[0].key, cruce);
}

function renderModuloLiqGoC() {
  const panel = document.getElementById('mod-liq-goc');
  if (!panel) return;

  const filasGoC = typeof RESULTADO !== 'undefined'
    ? RESULTADO.filter(r => r.sky?.esGOCUOTAS) : [];

  if (!filasGoC.length) {
    panel.innerHTML = _liqEmpty('Go Cuotas', '💳',
      'Cargá el archivo de Go Cuotas en OPERACIONES y ejecutá el cruce.');
    return;
  }

  const cruce = _cruzarLiqGoC();
  _liqCache.goc = cruce;
  const { liquidadas, noLiquidadas, extras,
          montoLiquidado, montoNoLiq, montoExtras, totalOK, tieneLiq } = cruce;
  const pct = totalOK ? +_liqPct(liquidadas.length, totalOK) : 0;

  const infoBar = `
  <div style="display:flex;align-items:center;gap:10px;padding:10px 18px;
    background:var(--s2);border-bottom:1px solid var(--b1);flex-shrink:0">
    <span style="font-size:14px">💳</span>
    <div style="font-size:10px;font-weight:700;color:var(--txt)">Go Cuotas · Liquidación</div>
    <span style="background:rgba(79,142,247,.12);color:var(--acc);padding:2px 8px;
      border-radius:4px;font-size:9px">
      ℹ El CSV de GoC cargado en OPERACIONES ya es la liquidación — no requiere archivo adicional
    </span>
    ${tieneLiq
      ? `<span style="background:rgba(52,211,153,.12);color:var(--grn);padding:2px 8px;
           border-radius:4px;font-size:9px;font-weight:600">
           ✓ Datos GoC disponibles
         </span>`
      : `<span style="background:rgba(251,191,36,.1);color:#fbbf24;padding:2px 8px;
           border-radius:4px;font-size:9px">
           ⚠ Cargá el CSV de GoC en OPERACIONES primero
         </span>`}
    <span style="margin-left:auto;font-size:9px;color:var(--m2)">Clave: Número de Orden</span>
  </div>`;

  const kpis = [
    { label:'💳 Ops GoC Skylab',  val:filasGoC.length.toLocaleString('es-AR'),
      bc:'rgba(79,142,247,.4)', cls:'cyn', sub:_liqFmtARS(montoLiquidado+montoNoLiq) },
    { label:'✓ Con orden GoC',    val:liquidadas.length.toLocaleString('es-AR'),
      bc:'rgba(52,211,153,.3)', cls:'grn', sub:_liqFmtARS(montoLiquidado), pct },
    { label:'✗ Sin orden GoC',   val:noLiquidadas.length.toLocaleString('es-AR'),
      bc:'rgba(248,113,113,.3)', cls:'red', sub:_liqFmtARS(montoNoLiq),
      pct:totalOK ? +_liqPct(noLiquidadas.length,totalOK) : 0 },
    { label:'⚠ Pagos s/Skylab',  val:extras.length.toLocaleString('es-AR'),
      bc:'rgba(251,146,60,.25)', cls:'org', sub:_liqFmtARS(montoExtras) },
    { label:'% Cubierto',         val:`${pct}%`,
      bc:'rgba(79,142,247,.3)',
      cls: pct>=90?'grn':pct>=70?'yel':'red',
      sub:`${liquidadas.length} de ${totalOK}`, pct },
  ];

  const tabs = [
    { key:'liq',   label:'✓ Con orden GoC',  n:liquidadas.length,
      cs:'background:rgba(52,211,153,.15);color:var(--grn)' },
    { key:'noliq', label:'✗ Sin orden GoC', n:noLiquidadas.length,
      cs:`background:rgba(248,113,113,.15);color:var(--red)${noLiquidadas.length?';font-weight:700':''}` },
    ...(extras.length ? [{ key:'extras', label:'⚠ Pagos extra GoC', n:extras.length,
      cs:'background:rgba(251,146,60,.12);color:var(--org)' }] : []),
  ];

  panel.innerHTML = _liqBuildPanel({ id:'goc', kpis, tabs, fileSection: infoBar });
  _liqPopTab('goc', tabs[0].key, cruce);
}

// ── Empty state ──────────────────────────────────────────────────────
function _liqEmpty(nombre, icon, msg) {
  return `<div style="display:flex;flex-direction:column;align-items:center;justify-content:center;
    height:100%;gap:14px;color:var(--m2);text-align:center;padding:40px">
    <div style="font-size:42px;opacity:.15">${icon}</div>
    <div style="font-size:16px;font-weight:700;color:var(--txt);opacity:.4">${nombre}</div>
    <p style="font-size:10px;max-width:400px;line-height:1.8">${msg}</p>
  </div>`;
}

// ── Refrescar todos ──────────────────────────────────────────────────
function renderModuloLiquidaciones() {
  renderModuloLiqFiserv();
  renderModuloLiqGetpos();
  renderModuloLiqGoC();
}
