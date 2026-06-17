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

// ── Días hábiles entre dos fechas ISO (excluye sáb, dom y TM.feriados) ──
function _diasHabilesEntre(iso1, iso2) {
  if (!iso1 || !iso2) return null;
  const feriados = new Set((TM?.feriados || []).map(f => f.fecha));
  const d1 = new Date(iso1 + 'T00:00:00');
  const d2 = new Date(iso2 + 'T00:00:00');
  if (isNaN(d1) || isNaN(d2) || d1 >= d2) return 0;
  let dias = 0;
  const cur = new Date(d1);
  cur.setDate(cur.getDate() + 1);
  while (cur <= d2) {
    const dow = cur.getDay();
    const iso = cur.toISOString().slice(0, 10);
    if (dow !== 0 && dow !== 6 && !feriados.has(iso)) dias++;
    cur.setDate(cur.getDate() + 1);
  }
  return dias;
}

// ── Buscar plazo acordado en TM.plazos ──────────────────────────────
function _buscarPlazoEnTM(procesadora, comercio, tarjeta, fecha) {
  if (!TM?.plazos?.length) return null;
  const candidatos = TM.plazos.filter(p => {
    const mP = !p.procesadora || p.procesadora.toUpperCase() === (procesadora||'').toUpperCase();
    const mC = !p.comercio    || p.comercio === String(comercio||'');
    const mT = !p.tarjeta     || p.tarjeta.toUpperCase() === (tarjeta||'').toUpperCase();
    const vigente = (!p.vigDesde || p.vigDesde <= fecha) && (!p.vigHasta || p.vigHasta >= fecha);
    return mP && mC && mT && vigente;
  });
  if (!candidatos.length) return null;
  candidatos.sort((a, b) =>
    [b.procesadora,b.comercio,b.tarjeta].filter(Boolean).length -
    [a.procesadora,a.comercio,a.tarjeta].filter(Boolean).length);
  return candidatos[0];
}

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
  // Strip leading apostrophes (Excel text prefix) y espacios, luego quita ceros iniciales
  return _liqNorm(String(v ?? '').replace(/^['"]+/, '').trim());
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

    const cuotasRaw = _liqGCol(r,'Cuotas','cuotas','CUOTAS','Nro Cuotas');
    const cfoRaw    = _liqGCol(r,'CFO','cfo','Costo Financiero');
    const arancRaw  = _liqGCol(r,'Arancel','arancel','ARANCEL');
    const tnaRaw    = _liqGCol(r,'TNA','tna','Tasa Nominal');

    out.push({
      lote,
      cupon,
      aut,
      monto,
      equipo:        String(_liqGCol(r,'Nro Equipo','Nro. Equipo','Equipo','Terminal','terminal') || '').trim(),
      nombre_equipo: String(_liqGCol(r,'Nombre de equipo','Nombre Equipo','NombreEquipo') || '').trim(),
      nro_comercio:  String(_liqGCol(r,'Nro Comercio','Nro. Comercio','Nro.Comercio','NroComercio') || '').trim(),
      liq_id:        String(_liqGCol(r,'Nro Liquidación','Nro Liquidacion','Liquidacion') || '').trim(),
      fecha_venta:   _liqFecha(_liqGCol(r,'Fecha Venta','Fecha','fecha')),
      fecha_pago:    _liqFecha(_liqGCol(r,'Fecha Pago','FechaPago')),
      tarjeta:       String(_liqGCol(r,'Tarjeta','tarjeta') || '').trim(),
      nro_tarjeta:   String(_liqGCol(r,'Nro Tarjeta','NroTarjeta','Nro. Tarjeta','Numero Tarjeta','NumeroTarjeta') || '').trim(),
      banco_pagador: String(_liqGCol(r,'Banco Pagador','BancoPagador','Banco','banco_pagador') || '').trim(),
      tipo,
      cuotas:  typeof cuotasRaw === 'number' ? Math.round(cuotasRaw) : (parseInt(cuotasRaw) || 1),
      cfo:     typeof cfoRaw    === 'number' ? cfoRaw    : _liqMonto(cfoRaw),
      arancel: typeof arancRaw  === 'number' ? arancRaw  : _liqMonto(arancRaw),
      tna:     typeof tnaRaw    === 'number' ? tnaRaw    : (parseFloat(String(tnaRaw || '0')) || 0),
    });
  }

  console.log(`[LIQ] Parseados: ${out.length} cupones liquidados`);
  if (out.length > 0) {
    console.log('[LIQ] Muestra cfo/arancel:', out.slice(0,5).map(r =>
      `cuotas=${r.cuotas} monto=${r.monto} cfo=${r.cfo} arancel=${r.arancel}`).join(' | '));
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

// ── CRUCE: FISERV ─────────────────────────────────────────────────────
// Dirección: PROC → LIQ
// Para cada proc FISERV confirmado en RESULTADO, buscar su fila en el liq.
// Clave: Cupón | Lote | Importe | Equipo | Cuota  (tolerancia ±1 peso en importe)
function _cruzarLiqFiserv() {
  if (typeof RESULTADO === 'undefined' || !RESULTADO.length) return null;

  const filasFis     = RESULTADO.filter(r => !r.sky?.esGETPos && !r.sky?.esGOCUOTAS);
  const sinConfirmar = [];
  const confirmadas  = [];
  for (const fila of filasFis) {
    const esConfirmada = fila.estado?.startsWith('OK') ||
      fila.estado === 'DIF. CUOTAS' ||
      fila.estado?.startsWith('COM. ERRADO') ||
      fila.estado?.startsWith('MAL FACTURADO') ||
      (fila.estado?.startsWith('CORREGIDO MANUAL') && !!fila.proc);
    if (!esConfirmada || fila.esDevolucion === 'SI' || fila.sky?.esNeg) sinConfirmar.push(fila);
    else confirmadas.push(fila);
  }

  if (!_LIQ_CUPONES.length) {
    return {
      liquidadas: [], sinConfirmar, extras: [], fueraPlazo: [],
      noLiquidadas: confirmadas.map(fila => ({
        fila,
        lote:  _liqNorm(fila.proc?.lote || ''),
        cupon: _liqNorm(fila.proc?.ticket || fila.proc?.cupon || ''),
        aut:   _liqNormAut(fila.proc?.aut || ''),
        enLiq: null, liqRow: null,
      })),
      montoLiquidado:0, montoNoLiq: confirmadas.reduce((s,f)=>s+Math.abs(f.sky?.monto||0),0),
      montoExtras:0, montoFueraPlazo:0,
      totalOK: confirmadas.length, tieneLiq:false, tienePlazos:!!(TM?.plazos?.length),
    };
  }

  // Índice de liqRows FISERV: "cupon|lote|monto|equipo|cuotas" → [liqRows]
  // Se indexa con monto, monto+1 y monto-1 para tolerar redondeo de centavos
  const liqIdx = {};
  for (const r of _LIQ_CUPONES) {
    if (!r.lote || r.lote === '0' || !r.cupon || r.cupon === '0') continue;
    const m  = Math.round(Math.abs(r.monto || 0));
    const eq = _liqNorm(r.equipo || '');
    const cu = r.cuotas || 0;
    for (const mv of [m, m + 1, Math.max(0, m - 1)]) {
      const k = `${r.cupon}|${r.lote}|${mv}|${eq}|${cu}`;
      (liqIdx[k] = liqIdx[k] || []).push(r);
    }
  }

  const usadosLiq    = new Set();  // referencias a liqRows ya asignadas
  const liquidadas   = [];
  const noLiquidadas = [];

  for (const fila of confirmadas) {
    const pLote   = _liqNorm(fila.proc?.lote || '');
    const pCupon  = _liqNorm(fila.proc?.ticket || fila.proc?.cupon || '');
    const pMonto  = parseInt(fila.proc?.montoN) || Math.round(Math.abs(fila.sky?.monto || 0));
    const pEquipo = _liqNorm(String(fila.proc?.equipo || ''));
    const pCuotas = parseInt(fila.proc?.cuotas || fila.sky?.cuotas) || 0;
    const pAut    = _liqNormAut(fila.proc?.aut || '');

    const k      = `${pCupon}|${pLote}|${pMonto}|${pEquipo}|${pCuotas}`;
    const liqRow = (liqIdx[k] || []).find(r => !usadosLiq.has(r)) || null;

    if (liqRow) {
      usadosLiq.add(liqRow);
      const fechaVenta = liqRow.fecha_venta || fila.sky?.fecha || '';
      const fechaPago  = liqRow.fecha_pago  || '';
      const plazoTM    = _buscarPlazoEnTM('FISERV', liqRow.nro_comercio || '', liqRow.tarjeta || '', fechaVenta);
      let diasHabiles = null, diasEsperados = null, diasExtra = null;
      if (plazoTM && fechaVenta && fechaPago) {
        diasHabiles   = _diasHabilesEntre(fechaVenta, fechaPago);
        diasEsperados = parseInt(plazoTM.dias_habiles) || 0;
        diasExtra     = diasHabiles - diasEsperados;
      }
      liquidadas.push({ fila, lote: pLote, cupon: pCupon, aut: pAut,
        liqRow, diasHabiles, diasEsperados, diasExtra, plazoTM });
    } else {
      noLiquidadas.push({ fila, lote: pLote, cupon: pCupon, aut: pAut, enLiq: false, liqRow: null });
    }
  }

  // Extras: filas del liq con lote+cupon que ningún proc reclamó
  const extras = _LIQ_CUPONES.filter(r =>
    r.lote && r.lote !== '0' && r.cupon && r.cupon !== '0' && !usadosLiq.has(r)
  );

  const fueraPlazo      = liquidadas.filter(x => x.diasExtra !== null && x.diasExtra > 0);
  const montoLiquidado  = liquidadas.reduce((s,x) => s + (x.liqRow?.monto || Math.abs(x.fila.sky?.monto||0)), 0);
  const montoNoLiq      = noLiquidadas.reduce((s,x) => s + Math.abs(x.fila.sky?.monto||0), 0);
  const montoExtras     = extras.reduce((s,r) => s + (r.monto||0), 0);
  const montoFueraPlazo = fueraPlazo.reduce((s,x) => s + (x.liqRow?.monto||0), 0);

  return {
    liquidadas, noLiquidadas, sinConfirmar, extras, fueraPlazo,
    montoLiquidado, montoNoLiq, montoExtras, montoFueraPlazo,
    totalOK: liquidadas.length + noLiquidadas.length,
    tieneLiq: true, tienePlazos: !!(TM?.plazos?.length),
  };
}

// ── CRUCE: GETPOS ─────────────────────────────────────────────────────
// Dirección: PROC → LIQ
// Para cada proc GETPOS confirmado en RESULTADO, buscar su fila en el liq.
// Clave: Cód. Aut. | Importe | Equipo | Cuota  (tolerancia ±1 peso en importe)
// proc.pos (Código del POS) ↔ liqRow.equipo (Nro Equipo)
function _cruzarLiqGetpos() {
  if (typeof RESULTADO === 'undefined' || !RESULTADO.length) return null;

  const filasGP      = RESULTADO.filter(r => r.sky?.esGETPos && !r.sky?.esGOCUOTAS);
  const sinConfirmar = [];
  const confirmadas  = [];
  for (const fila of filasGP) {
    const esConfirmada = fila.estado?.startsWith('OK') ||
      fila.estado === 'DIF. CUOTAS' ||
      fila.estado?.startsWith('COM. ERRADO') ||
      fila.estado?.startsWith('MAL FACTURADO') ||
      (fila.estado?.startsWith('CORREGIDO MANUAL') && !!fila.proc);
    if (!esConfirmada || fila.esDevolucion === 'SI' || fila.sky?.esNeg) sinConfirmar.push(fila);
    else confirmadas.push(fila);
  }

  if (!_LIQ_CUPONES.length) {
    return {
      liquidadas: [], sinConfirmar, extras: [],
      noLiquidadas: confirmadas.map(fila => ({
        fila, aut: _liqNormAut(fila.proc?.aut || ''), enLiq: null, liqRow: null,
      })),
      montoLiquidado:0, montoNoLiq: confirmadas.reduce((s,f)=>s+Math.abs(f.sky?.monto||0),0),
      montoExtras:0, totalOK: confirmadas.length, tieneLiq:false,
    };
  }

  // Índice de liqRows con dos claves por entrada:
  //   FULL:  "aut|monto|equipo|cuotas"  (con equipo)
  //   SHORT: "aut|monto||cuotas"        (sin equipo, fallback si proc.pos está vacío)
  // monto se indexa también con ±1 para tolerar redondeo de centavos
  const liqIdx = {};
  for (const r of _LIQ_CUPONES) {
    const aut = _liqNormAut(r.aut || '');
    if (!aut || aut === '0') continue;
    const m  = Math.round(Math.abs(r.monto || 0));
    const eq = _liqNorm(r.equipo || '');
    const cu = r.cuotas || 0;
    for (const mv of [m, m + 1, Math.max(0, m - 1)]) {
      const kFull  = `${aut}|${mv}|${eq}|${cu}`;
      const kShort = `${aut}|${mv}||${cu}`;
      (liqIdx[kFull]  = liqIdx[kFull]  || []).push(r);
      if (eq !== '0') (liqIdx[kShort] = liqIdx[kShort] || []).push(r);
    }
  }

  const usadosLiq    = new Set();
  const liquidadas   = [];
  const noLiquidadas = [];

  for (const fila of confirmadas) {
    const pAut    = _liqNormAut(fila.proc?.aut || '');
    const pMonto  = parseInt(fila.proc?.montoN) || Math.round(Math.abs(fila.sky?.monto || 0));
    const pEquipo = _liqNorm(String(fila.proc?.pos || fila.proc?.equipo || ''));
    const pCuotas = parseInt(fila.proc?.cuotas || fila.sky?.cuotas) || 0;

    // 1°: clave completa aut+monto+equipo+cuotas
    const kFull  = `${pAut}|${pMonto}|${pEquipo}|${pCuotas}`;
    // 2°: fallback sin equipo (cuando proc.pos está vacío o no coincide)
    const kShort = `${pAut}|${pMonto}||${pCuotas}`;

    const liqRow =
      (liqIdx[kFull]  || []).find(r => !usadosLiq.has(r)) ||
      (pEquipo === '0' ? (liqIdx[kShort] || []).find(r => !usadosLiq.has(r)) : null) ||
      null;

    if (liqRow) {
      usadosLiq.add(liqRow);
      liquidadas.push({ fila, aut: pAut, liqRow });
    } else {
      noLiquidadas.push({ fila, aut: pAut, enLiq: false, liqRow: null });
    }
  }

  // Extras: liqRows con aut válido que ningún proc reclamó
  const extras = _LIQ_CUPONES.filter(r => {
    const aut = _liqNormAut(r.aut || '');
    return aut && aut !== '0' && !usadosLiq.has(r);
  });

  const montoLiquidado = liquidadas.reduce((s,x) => s + (x.liqRow?.monto || Math.abs(x.fila.sky?.monto||0)), 0);
  const montoNoLiq     = noLiquidadas.reduce((s,x) => s + Math.abs(x.fila.sky?.monto||0), 0);
  const montoExtras    = extras.reduce((s,r) => s + (r.monto||0), 0);

  return {
    liquidadas, noLiquidadas, sinConfirmar, extras,
    montoLiquidado, montoNoLiq, montoExtras,
    totalOK: liquidadas.length + noLiquidadas.length,
    tieneLiq: true,
  };
}

// ── CRUCE: Go Cuotas ─────────────────────────────────────────────────
// Prioridad: usa _GOC_LIQ_PAGOS/_GOC_LIQ_CELULAR si están cargados
// (archivos de LIQUIDACIONES); si no, cae en _GOC_PAGOS/_GOC_CELULAR
// (archivos de OPERACIONES, que en GoC suelen ser el mismo informe).
function _cruzarLiqGoC() {
  if (typeof RESULTADO === 'undefined' || !RESULTADO.length) return null;

  const filasGoC = RESULTADO.filter(r => r.sky?.esGOCUOTAS);

  const liqPagos   = (typeof _GOC_LIQ_PAGOS   !== 'undefined' && _GOC_LIQ_PAGOS.length)
                       ? _GOC_LIQ_PAGOS   : (typeof _GOC_PAGOS   !== 'undefined' ? _GOC_PAGOS   : []);
  const liqCelular = (typeof _GOC_LIQ_CELULAR !== 'undefined' && _GOC_LIQ_CELULAR.length)
                       ? _GOC_LIQ_CELULAR : (typeof _GOC_CELULAR !== 'undefined' ? _GOC_CELULAR : []);
  const _gocAll    = [...liqPagos, ...liqCelular];

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

// ── Exportar todas las pestañas de un módulo a Excel multi-hoja ──────
function _liqExportarTodo(proc) {
  const cruce = _liqCache[proc];
  if (!cruce) { _showToast('Primero ejecutá el cruce de Liquidaciones.'); return; }

  const wb  = XLSX.utils.book_new();
  const hoy = new Date().toISOString().slice(0, 10);
  const num = v => (typeof v === 'number' ? v : parseFloat(String(v).replace(/[^0-9.,-]/g,'').replace(',','.')) || 0);

  const addSheet = (name, headers, rows) => {
    if (!rows.length) return;
    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
  };

  /* ── FISERV ── */
  if (proc === 'fiserv') {
    const H = [
      'Estado',
      // SKY
      'Fecha SKY','Suc.','Tarjeta SKY','Plan SKY','Cuotas SKY','Lote SKY','Cupón SKY','Nro. Asiento','Monto SKY',
      // LIQ
      'Fecha Venta','Fecha Pago','Nro Equipo','Nombre Equipo',
      'Nro Lote','Nro Cupón','Tarjeta Liq.','Nro Tarjeta','Cód. Autorización',
      'Cuotas Liq.','Importe Venta','Nro Comercio','Banco Pagador','N° Liq.',
    ];
    const _rFis = (sky, r, lote, cupon, aut) => [
      sky?.fecha||'', sky?.suc||'', sky?.tarjeta||'', sky?.plan||'',
      sky?.cuotas||'', sky?.lote||'', sky?.cupon||'', sky?.asiento||'', num(sky?.monto),
      r?.fecha_venta||'', r?.fecha_pago||'', r?.equipo||'', r?.nombre_equipo||'',
      lote||'', cupon||'', r?.tarjeta||'', r?.nro_tarjeta||'', aut||'',
      r?.cuotas||'', num(r?.monto), r?.nro_comercio||'', r?.banco_pagador||'', r?.liq_id||'',
    ];
    addSheet('Liquidadas', H, cruce.liquidadas.map(x => [
      'Liquidada', ..._rFis(x.fila.sky, x.liqRow, x.lote, x.cupon, x.aut)]));
    addSheet('No liquidadas', H, cruce.noLiquidadas.map(x => [
      x.enLiq===null ? 'Sin archivo' : 'No liquidada',
      ..._rFis(x.fila.sky, null, x.lote, x.cupon, x.aut)]));
    addSheet('Sin confirmar', H, cruce.sinConfirmar.map(f => [
      f.estado||'SIN MATCH',
      ..._rFis(f.sky, null, f.sky?.lote, f.sky?.cupon, '')]));
    addSheet('Extras en Liq.', H, cruce.extras.map(r => [
      'Extra en Liq.', ..._rFis(null, r, r.lote, r.cupon, r.aut)]));
    if (cruce.fueraPlazo?.length) {
      addSheet('Fuera de plazo',
        [...H, 'Días Esperados','Días Reales','Días Extra'],
        cruce.fueraPlazo.map(x => [
          `+${x.diasExtra}d`,
          ..._rFis(x.fila.sky, x.liqRow, x.lote, x.cupon, x.aut),
          x.diasEsperados, x.diasHabiles, x.diasExtra]));
    }
  }

  /* ── GETPOS ── */
  if (proc === 'getpos') {
    const H = [
      'Estado',
      // SKY
      'Fecha SKY','Suc.','Tarjeta SKY','Plan SKY','Cuotas SKY','Lote SKY','Cupón SKY','Nro. Asiento','Monto SKY',
      // LIQ
      'Fecha Venta','Fecha Pago','Nro Equipo','Nombre Equipo',
      'Nro Cupón','Tarjeta Liq.','Nro Tarjeta','Cód. Autorización',
      'Cuotas Liq.','Importe Venta','Nro Comercio','Banco Pagador',
    ];
    const _rGP = (sky, r, aut, cupon) => [
      sky?.fecha||'', sky?.suc||'', sky?.tarjeta||'', sky?.plan||'',
      sky?.cuotas||'', sky?.lote||'', sky?.cupon||'', sky?.asiento||'', num(sky?.monto),
      r?.fecha_venta||'', r?.fecha_pago||'', r?.equipo||'', r?.nombre_equipo||'',
      cupon||r?.cupon||'', r?.tarjeta||'', r?.nro_tarjeta||'', aut||r?.aut||'',
      r?.cuotas||'', num(r?.monto), r?.nro_comercio||'', r?.banco_pagador||'',
    ];
    addSheet('Liquidadas', H, cruce.liquidadas.map(x => [
      'Liquidada', ..._rGP(x.fila.sky, x.liqRow, x.aut, x.fila.proc?.cupon||x.fila.sky?.cupon)]));
    addSheet('No liquidadas', H, cruce.noLiquidadas.map(x => [
      x.enLiq===null ? 'Sin archivo' : 'No liquidada',
      ..._rGP(x.fila.sky, null, x.aut, x.fila.proc?.cupon||x.fila.sky?.cupon)]));
    addSheet('Sin confirmar', H, cruce.sinConfirmar.map(f => [
      f.estado||'SIN MATCH', ..._rGP(f.sky, null, '', f.sky?.cupon)]));
    addSheet('Extras en Liq.', H, cruce.extras.map(r => [
      'Extra en Liq.', ..._rGP(null, r, r.aut, r.cupon)]));
  }

  /* ── GoC ── */
  if (proc === 'goc') {
    const H = ['Estado','Fecha Skylab','Sucursal','N° Orden GoC',
               'Cupón Skylab','Monto Skylab','Fecha Pago GoC','Método'];
    addSheet('Liquidadas', H, cruce.liquidadas.map(x => [
      x.fila.estado||'Con orden', x.fila.sky?.fecha||'', x.fila.sky?.suc||'',
      x.orden, x.fila.sky?.cupon||'', num(x.fila.sky?.monto), '', x.fila.metodo||'']));
    addSheet('No liquidadas', H, cruce.noLiquidadas.map(x => [
      x.fila.estado||'SIN MATCH', x.fila.sky?.fecha||'', x.fila.sky?.suc||'',
      '', x.fila.sky?.cupon||'', num(x.fila.sky?.monto), '', '']));
    addSheet('Extras', H, cruce.extras.map(p => [
      'Pago sin Skylab', p.fechaOrigen||'', p.sucNombre||'',
      p.orden, p.refExt||'', num(p.importe), p.fechaPago||'', p.fuente||'']));
  }

  if (!wb.SheetNames.length) { _showToast('Sin datos para exportar.'); return; }
  const label = proc.toUpperCase();
  XLSX.writeFile(wb, `liquidaciones_${label}_${hoy}.xlsx`);
  _showToast(`✓ liquidaciones_${label}_${hoy}.xlsx`);
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
    <div style="display:flex;align-items:center;gap:6px">
      <div class="tab-strip" id="tab-strip-liq-${id}" style="display:flex;align-items:center;flex-wrap:wrap;flex:1;margin-bottom:0">
        ${tabs.map((t,i)=>`
          <button class="tb${i===0?' active':''}" data-tab="${t.key}"
            onclick="showTab('liqtab-${id}-${t.key}','tab-strip-liq-${id}',this);_liqRenderTab('${id}','${t.key}')">
            ${t.label} <span class="cnt" style="${t.cs||''}">${t.n}</span>
          </button>`).join('')}
      </div>
      <button onclick="_liqExportarTodo('${id}')"
        style="background:none;border:1px solid var(--b2);color:var(--m1);
          border-radius:4px;padding:3px 10px;font-size:9px;cursor:pointer;
          font-family:var(--sans);white-space:nowrap;flex-shrink:0"
        title="Exportar todas las pestañas a Excel">
        ↓ Exportar todo
      </button>
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
    // Columnas SKY + todas las columnas del archivo de liquidaciones
    const colsFis = [
      { key:'_estado',       label:'Estado' },
      // — Skylab —
      { key:'fecha',         label:'Fecha SKY' },
      { key:'suc',           label:'Suc.' },
      { key:'tarjeta_sky',   label:'Tarjeta SKY' },
      { key:'plan_sky',      label:'Plan SKY' },
      { key:'cuotas_sky',    label:'Cuotas SKY', cls:'num' },
      { key:'lote_sky',      label:'Lote SKY' },
      { key:'cupon_sky',     label:'Cupón SKY' },
      { key:'asiento_sky',   label:'Nro. Asiento' },
      { key:'monto_sky',     label:'Monto SKY', cls:'num' },
      // — Liquidación —
      { key:'fecha_venta',   label:'Fecha Venta' },
      { key:'fecha_pago',    label:'Fecha Pago' },
      { key:'nro_equipo',    label:'Nro Equipo' },
      { key:'nombre_equipo', label:'Nombre Equipo' },
      { key:'lote',          label:'Nro Lote' },
      { key:'cupon',         label:'Nro Cupón' },
      { key:'tarjeta_liq',   label:'Tarjeta Liq.' },
      { key:'nro_tarjeta',   label:'Nro Tarjeta' },
      { key:'aut',           label:'Cód. Autorización' },
      { key:'cuotas_liq',    label:'Cuotas Liq.', cls:'num' },
      { key:'monto_liq',     label:'Importe Venta', cls:'num' },
      { key:'nro_comercio',  label:'Nro Comercio' },
      { key:'banco_pagador', label:'Banco Pagador' },
      { key:'liq_id',        label:'N° Liq.' },
    ];

    const _skyFis = (sky) => ({
      fecha:         sky?.fecha       || '',
      suc:           sky?.suc         || '',
      tarjeta_sky:   sky?.tarjeta     || '',
      plan_sky:      sky?.plan        || '',
      cuotas_sky:    sky?.cuotas      || '',
      lote_sky:      sky?.lote        || '',
      cupon_sky:     sky?.cupon       || '',
      asiento_sky:   sky?.asiento     || '',
      monto_sky:     fmtM(sky?.monto),
    });
    const _liqFis = (r) => ({
      fecha_venta:   r?.fecha_venta   || '—',
      fecha_pago:    r?.fecha_pago    || '—',
      nro_equipo:    r?.equipo        || '—',
      nombre_equipo: r?.nombre_equipo || '—',
      lote:          r?.lote          || '—',
      cupon:         r?.cupon         || '—',
      tarjeta_liq:   r?.tarjeta       || '—',
      nro_tarjeta:   r?.nro_tarjeta   || '—',
      aut:           r?.aut           || '—',
      cuotas_liq:    r?.cuotas        || '—',
      monto_liq:     r ? fmtM(r.monto) : '—',
      nro_comercio:  r?.nro_comercio  || '—',
      banco_pagador: r?.banco_pagador || '—',
      liq_id:        r?.liq_id        || '—',
    });
    const _noLiq = () => ({
      fecha_venta:'—', fecha_pago:'—', nro_equipo:'—', nombre_equipo:'—',
      lote:'—', cupon:'—', tarjeta_liq:'—', nro_tarjeta:'—', aut:'—',
      cuotas_liq:'—', monto_liq:'—', nro_comercio:'—', banco_pagador:'—', liq_id:'—',
    });


    if (tab === 'liq') {
      _liqRenderTable(`tbl-liq-fiserv-liq`, colsFis,
        cruce.liquidadas.map(x => ({
          _estado: '✓ Liquidada',
          ..._skyFis(x.fila.sky),
          ..._liqFis(x.liqRow),
          lote:  x.lote  || x.liqRow?.lote  || '—',
          cupon: x.cupon || x.liqRow?.cupon || '—',
          aut:   x.aut   || x.liqRow?.aut   || '—',
        })));
    }
    if (tab === 'noliq') {
      _liqRenderTable(`tbl-liq-fiserv-noliq`, colsFis,
        cruce.noLiquidadas.map(x => ({
          _estado: x.enLiq === null ? '⚠ Sin archivo' : '✗ No liquidada',
          ..._skyFis(x.fila.sky),
          ..._noLiq(),
          lote:  x.lote  || '—',
          cupon: x.cupon || '—',
          aut:   x.aut   || '—',
        })));
    }
    if (tab === 'sinconf') {
      _liqRenderTable(`tbl-liq-fiserv-sinconf`, colsFis,
        cruce.sinConfirmar.map(f => ({
          _estado: f.estado || 'SIN MATCH',
          ..._skyFis(f.sky),
          ..._noLiq(),
          lote:  f.sky?.lote  || '—',
          cupon: f.sky?.cupon || '—',
          aut:   '—',
        })));
    }
    if (tab === 'extras') {
      _liqRenderTable(`tbl-liq-fiserv-extras`, colsFis,
        cruce.extras.map(r => ({
          _estado: '⚠ Extra en Liq.',
          fecha: r.fecha_venta || '', suc: '', tarjeta_sky: '—', plan_sky: '—',
          cuotas_sky: '—', lote_sky: '—', cupon_sky: '—', asiento_sky: '—', monto_sky: '—',
          ..._liqFis(r),
        })));
    }
    if (tab === 'fueraplazo') {
      const colsFP = [
        { key:'_estado',       label:'Estado' },
        { key:'fecha',         label:'Fecha SKY' },
        { key:'suc',           label:'Suc.' },
        { key:'tarjeta_sky',   label:'Tarjeta SKY' },
        { key:'plan_sky',      label:'Plan SKY' },
        { key:'cuotas_sky',    label:'Cuotas SKY', cls:'num' },
        { key:'lote_sky',      label:'Lote SKY' },
        { key:'cupon_sky',     label:'Cupón SKY' },
        { key:'asiento_sky',   label:'Nro. Asiento' },
        { key:'monto_sky',     label:'Monto SKY', cls:'num' },
        { key:'fecha_venta',   label:'Fecha Venta' },
        { key:'fecha_pago',    label:'Fecha Pago' },
        { key:'nro_equipo',    label:'Nro Equipo' },
        { key:'nombre_equipo', label:'Nombre Equipo' },
        { key:'lote',          label:'Nro Lote' },
        { key:'cupon',         label:'Nro Cupón' },
        { key:'tarjeta_liq',   label:'Tarjeta Liq.' },
        { key:'nro_tarjeta',   label:'Nro Tarjeta' },
        { key:'aut',           label:'Cód. Autorización' },
        { key:'cuotas_liq',    label:'Cuotas Liq.', cls:'num' },
        { key:'monto_liq',     label:'Importe Venta', cls:'num' },
        { key:'nro_comercio',  label:'Nro Comercio' },
        { key:'banco_pagador', label:'Banco Pagador' },
        { key:'liq_id',        label:'N° Liq.' },
        { key:'dias_esperados',label:'Días Esperados', cls:'num' },
        { key:'dias_reales',   label:'Días Reales', cls:'num' },
        { key:'dias_extra',    label:'Días Extra', cls:'num' },
      ];
      _liqRenderTable(`tbl-liq-fiserv-fueraplazo`, colsFP,
        cruce.fueraPlazo.map(x => ({
          _estado:       `⏱ +${x.diasExtra}d`,
          ..._skyFis(x.fila.sky),
          ..._liqFis(x.liqRow),
          lote:          x.lote  || x.liqRow?.lote  || '—',
          cupon:         x.cupon || x.liqRow?.cupon || '—',
          aut:           x.aut   || x.liqRow?.aut   || '—',
          dias_esperados:x.diasEsperados,
          dias_reales:   x.diasHabiles,
          dias_extra:    x.diasExtra,
          _tr_style:     x.diasExtra > 5 ? 'background:rgba(248,113,113,.08)' : 'background:rgba(167,139,250,.06)',
        })));
    }
  }

  /* ── GETPOS ─────────────────────────────────────────────────── */
  if (proc === 'getpos') {
    const colsGP = [
      { key:'_estado',       label:'Estado' },
      // — Skylab —
      { key:'fecha',         label:'Fecha SKY' },
      { key:'suc',           label:'Suc.' },
      { key:'tarjeta_sky',   label:'Tarjeta SKY' },
      { key:'plan_sky',      label:'Plan SKY' },
      { key:'cuotas_sky',    label:'Cuotas SKY', cls:'num' },
      { key:'lote_sky',      label:'Lote SKY' },
      { key:'cupon_sky',     label:'Cupón SKY' },
      { key:'asiento_sky',   label:'Nro. Asiento' },
      { key:'monto_sky',     label:'Monto SKY', cls:'num' },
      // — Liquidación —
      { key:'fecha_venta',   label:'Fecha Venta' },
      { key:'fecha_pago',    label:'Fecha Pago' },
      { key:'nro_equipo',    label:'Nro Equipo' },
      { key:'nombre_equipo', label:'Nombre Equipo' },
      { key:'cupon',         label:'Nro Cupón' },
      { key:'tarjeta_liq',   label:'Tarjeta Liq.' },
      { key:'nro_tarjeta',   label:'Nro Tarjeta' },
      { key:'aut',           label:'Cód. Autorización' },
      { key:'cuotas_liq',    label:'Cuotas Liq.', cls:'num' },
      { key:'monto_liq',     label:'Importe Venta', cls:'num' },
      { key:'nro_comercio',  label:'Nro Comercio' },
      { key:'banco_pagador', label:'Banco Pagador' },
    ];

    const _skyGP = (sky) => ({
      fecha:       sky?.fecha   || '',
      suc:         sky?.suc     || '',
      tarjeta_sky: sky?.tarjeta || '',
      plan_sky:    sky?.plan    || '',
      cuotas_sky:  sky?.cuotas  || '',
      lote_sky:    sky?.lote    || '',
      cupon_sky:   sky?.cupon   || '',
      asiento_sky: sky?.asiento || '',
      monto_sky:   fmtM(sky?.monto),
    });
    const _liqGP = (r, aut, cupon) => ({
      fecha_venta:   r?.fecha_venta   || '—',
      fecha_pago:    r?.fecha_pago    || '—',
      nro_equipo:    r?.equipo        || '—',
      nombre_equipo: r?.nombre_equipo || '—',
      cupon:         cupon            || r?.cupon || '—',
      tarjeta_liq:   r?.tarjeta       || '—',
      nro_tarjeta:   r?.nro_tarjeta   || '—',
      aut:           aut              || r?.aut   || '—',
      cuotas_liq:    r?.cuotas        || '—',
      monto_liq:     r ? fmtM(r.monto) : '—',
      nro_comercio:  r?.nro_comercio  || '—',
      banco_pagador: r?.banco_pagador || '—',
    });
    const _noLiqGP = (aut, cupon) => ({
      fecha_venta:'—', fecha_pago:'—', nro_equipo:'—', nombre_equipo:'—',
      cupon: cupon || '—', tarjeta_liq:'—', nro_tarjeta:'—',
      aut: aut || '—', cuotas_liq:'—', monto_liq:'—', nro_comercio:'—', banco_pagador:'—',
    });

    if (tab === 'liq') {
      _liqRenderTable(`tbl-liq-getpos-liq`, colsGP,
        cruce.liquidadas.map(x => ({
          _estado: '✓ Liquidada',
          ..._skyGP(x.fila.sky),
          ..._liqGP(x.liqRow, x.aut, x.fila.proc?.cupon || x.fila.sky?.cupon),
        })));
    }
    if (tab === 'noliq') {
      _liqRenderTable(`tbl-liq-getpos-noliq`, colsGP,
        cruce.noLiquidadas.map(x => ({
          _estado: x.enLiq === null ? '⚠ Sin archivo' : '✗ No liquidada',
          ..._skyGP(x.fila.sky),
          ..._noLiqGP(x.aut, x.fila.proc?.cupon || x.fila.sky?.cupon),
        })));
    }
    if (tab === 'sinconf') {
      _liqRenderTable(`tbl-liq-getpos-sinconf`, colsGP,
        cruce.sinConfirmar.map(f => ({
          _estado: f.estado || 'SIN MATCH',
          ..._skyGP(f.sky),
          ..._noLiqGP('', f.sky?.cupon),
        })));
    }
    if (tab === 'extras') {
      _liqRenderTable(`tbl-liq-getpos-extras`, colsGP,
        cruce.extras.map(r => ({
          _estado: '⚠ Extra en Liq.',
          fecha: r.fecha_venta || '', suc: '', tarjeta_sky: '—', plan_sky: '—',
          cuotas_sky: '—', lote_sky: '—', cupon_sky: '—', asiento_sky: '—', monto_sky: '—',
          ..._liqGP(r, r.aut, r.cupon),
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
  const { liquidadas, noLiquidadas, sinConfirmar, extras, fueraPlazo,
          montoLiquidado, montoNoLiq, montoExtras, montoFueraPlazo,
          totalOK, tieneLiq, tienePlazos } = cruce;
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
    ...(tienePlazos ? [{ label:'⏱ Fuera de plazo', val:fueraPlazo.length.toLocaleString('es-AR'),
      bc:'rgba(167,139,250,.3)', cls: fueraPlazo.length ? 'red' : 'grn',
      sub: fueraPlazo.length ? _liqFmtARS(montoFueraPlazo) : 'Todos en plazo' }] : []),
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
    ...(tienePlazos && fueraPlazo.length ? [{ key:'fueraplazo', label:'⏱ Fuera de plazo', n:fueraPlazo.length,
      cs:`background:rgba(167,139,250,.12);color:#a78bfa;font-weight:700` }] : []),
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

  const _hasLiqFile = (typeof _GOC_LIQ_PAGOS   !== 'undefined' && _GOC_LIQ_PAGOS.length > 0)
                   || (typeof _GOC_LIQ_CELULAR !== 'undefined' && _GOC_LIQ_CELULAR.length > 0);
  const _hasOpsFile = (typeof _GOC_PAGOS   !== 'undefined' && _GOC_PAGOS.length > 0)
                   || (typeof _GOC_CELULAR !== 'undefined' && _GOC_CELULAR.length > 0);
  const _liqCount   = (typeof _GOC_LIQ_PAGOS   !== 'undefined' ? _GOC_LIQ_PAGOS.length   : 0)
                    + (typeof _GOC_LIQ_CELULAR !== 'undefined' ? _GOC_LIQ_CELULAR.length : 0);
  const _opsCount   = (typeof _GOC_PAGOS   !== 'undefined' ? _GOC_PAGOS.length   : 0)
                    + (typeof _GOC_CELULAR !== 'undefined' ? _GOC_CELULAR.length : 0);

  const infoBar = `
  <div style="display:flex;align-items:center;gap:10px;padding:10px 18px;
    background:var(--s2);border-bottom:1px solid var(--b1);flex-shrink:0;flex-wrap:wrap">
    <span style="font-size:14px">💳</span>
    <div style="font-size:10px;font-weight:700;color:var(--txt)">Go Cuotas · Liquidación</div>

    ${_hasLiqFile
      ? `<span style="background:rgba(52,211,153,.15);color:var(--grn);padding:2px 8px;
           border-radius:4px;font-size:9px;font-weight:600">
           ✓ Archivo de Liquidación cargado · ${_liqCount.toLocaleString('es-AR')} registros
         </span>`
      : _hasOpsFile
        ? `<span style="background:rgba(79,142,247,.12);color:var(--acc);padding:2px 8px;
             border-radius:4px;font-size:9px">
             ℹ Usando CSV de OPERACIONES como liquidación · ${_opsCount.toLocaleString('es-AR')} registros
           </span>
           <label style="cursor:pointer;background:none;border:1px solid var(--b2);
             border-radius:4px;padding:3px 10px;font-size:9px;color:var(--m1);font-family:var(--sans)">
             ＋ Cargar CSV de Liquidación
             <input type="file" accept=".csv,.txt" style="display:none"
               onchange="loadGocPagosLiq(this,'GOCUOTAS')">
           </label>`
        : `<span style="background:rgba(251,191,36,.1);color:#fbbf24;padding:2px 8px;
             border-radius:4px;font-size:9px">
             ⚠ No hay datos GoC — cargá el CSV de Liquidación o el de OPERACIONES
           </span>
           <label style="cursor:pointer;background:linear-gradient(135deg,var(--acc),var(--cyn));
             color:#fff;border:none;border-radius:4px;padding:4px 14px;font-size:9px;
             font-weight:600;font-family:var(--sans)">
             ＋ Cargar Go Cuotas Liquidación CSV
             <input type="file" accept=".csv,.txt" style="display:none"
               onchange="loadGocPagosLiq(this,'GOCUOTAS')">
           </label>`}

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

// ══════════════════════════════════════════════════════════════════════
// MÓDULO TASAS — Diferencias de tasa entre lo cobrado y lo acordado
// Fuentes: TM.tasas (acordado) vs CFO/Importe Venta (cobrado real)
// ══════════════════════════════════════════════════════════════════════

// Tasa directa como fracción decimal: CFO / Importe Venta
function _liqTD(r) {
  return (r?.monto > 0) ? (r.cfo || 0) / r.monto : 0;
}

// Lookup en TM.tasas usando la fecha exacta de la operación (no "hoy")
function _buscarTasaParaFecha(tarjeta, cuotas, procesadora, fecha) {
  if (!TM?.tasas?.length) return null;
  const d   = (fecha || '').slice(0, 10) || new Date().toISOString().slice(0, 10);
  const up  = s => String(s || '').toUpperCase().trim();
  const c   = parseInt(cuotas) || 1;

  const candidatos = TM.tasas.filter(t => {
    if (t.tarjeta    && up(t.tarjeta)    !== up(tarjeta))    return false;
    if (t.cuotas     && parseInt(t.cuotas) !== c)            return false;
    if (t.procesadora && up(t.procesadora) !== up(procesadora)) return false;
    if (t.vigDesde   && t.vigDesde > d)                     return false;
    if (t.vigHasta   && t.vigHasta < d)                     return false;
    return true;
  });
  if (!candidatos.length) return null;

  // Más específico primero (más campos completos = mayor score)
  candidatos.sort((a, b) =>
    [b.tarjeta, b.cuotas, b.procesadora].filter(Boolean).length -
    [a.tarjeta, a.cuotas, a.procesadora].filter(Boolean).length
  );
  return candidatos[0];
}

// Cruce principal: compara TD cobrada con tasas acordadas de TM
function _cruzarTasas() {
  if (typeof RESULTADO === 'undefined' || !RESULTADO.length) return null;
  if (!_LIQ_CUPONES.length) return null;

  // Asegurar que los cruces FISERV/GETPOS estén computados (usan la dirección correcta PROC→LIQ)
  if (!_liqCache.fiserv) _liqCache.fiserv = _cruzarLiqFiserv();
  if (!_liqCache.getpos) _liqCache.getpos = _cruzarLiqGetpos();

  // Usar el liqRow ya encontrado correctamente por los cruces PROC→LIQ
  const filaLiqMap = new Map();
  for (const x of (_liqCache.fiserv?.liquidadas  || [])) filaLiqMap.set(x.fila, x.liqRow);
  for (const x of (_liqCache.fiserv?.fueraPlazo  || [])) filaLiqMap.set(x.fila, x.liqRow);
  for (const x of (_liqCache.getpos?.liquidadas  || [])) filaLiqMap.set(x.fila, x.liqRow);

  const pasarDescuento = [];   // error vendedor: tasa cobrada ≠ tasa del plan Skylab
  const reclamarProc   = [];   // error procesadora: tasa cobrada > tasa acordada para SU plan
  const sinTasa        = [];   // no hay config en TM.tasas para esta operación
  let   _okCount       = 0;   // con liqRow+CFO, tasa configurada y sin diferencia

  const filas = RESULTADO.filter(r => !r.sky?.esGOCUOTAS);

  for (const fila of filas) {
    const liqRow = filaLiqMap.get(fila) || null;
    if (!liqRow || !liqRow.monto) continue;

    const td_cobrada = _liqTD(liqRow);
    if (td_cobrada === 0) continue;   // débito / sin CFO

    const fecha      = liqRow.fecha_venta || fila.sky?.fecha || '';
    const skyTarjeta = fila.sky?.tarjeta  || '';
    const skyCuotas  = fila.sky?.cuotas   || 1;
    const liqTarjeta = liqRow.tarjeta     || '';
    const liqCuotas  = liqRow.cuotas      || 1;

    // Procesadora: FISERV o GETPOS según la fuente de la fila
    const procNom = (fila.procEncontrada || '')
      .toUpperCase().includes('GET') ? 'GETPOS' : 'FISERV';

    // Tasa acordada para el plan facturado (Skylab)
    const tmSky = _buscarTasaParaFecha(skyTarjeta, skyCuotas, procNom, fecha);
    // Tasa acordada para el plan cobrado (procesadora)
    const tmLiq = _buscarTasaParaFecha(liqTarjeta, liqCuotas, procNom, fecha);

    if (!tmSky && !tmLiq) {
      sinTasa.push({ fila, liqRow, td_cobrada, liqTarjeta, liqCuotas, skyTarjeta, skyCuotas, procNom });
      continue;
    }

    // Case 1 — PASAR A DESCUENTO: tasa cobrada ≠ tasa del plan Skylab
    let hayDifSky = false;
    if (tmSky) {
      const td_fact = parseFloat(tmSky.tasa || 0) / 100;
      const difPct  = td_cobrada - td_fact;
      if (Math.abs(difPct) >= 0.0005) {
        pasarDescuento.push({ fila, liqRow, td_cobrada, td_fact, tmRow: tmSky, difPct,
          difMonto: difPct * liqRow.monto, procNom });
        hayDifSky = true;
      }
    }

    // Case 2 — RECLAMAR A PROCESADORA: tasa cobrada > tasa acordada para SU propio plan
    let hayDifProc = false;
    if (tmLiq) {
      const td_ac  = parseFloat(tmLiq.tasa || 0) / 100;
      const difPct = td_cobrada - td_ac;
      if (difPct > 0.0005) {
        reclamarProc.push({ fila, liqRow, td_cobrada, td_ac, tmRow: tmLiq, difPct,
          difMonto: difPct * liqRow.monto, procNom });
        hayDifProc = true;
      }
    }

    if (!hayDifSky && !hayDifProc) _okCount++;
  }

  return {
    pasarDescuento, reclamarProc, sinTasa, _okCount,
    montoPD: pasarDescuento.reduce((s, x) => s + Math.abs(x.difMonto), 0),
    montoRP: reclamarProc.reduce((s, x) => s + Math.abs(x.difMonto), 0),
    tieneTasas: !!(TM?.tasas?.length),
  };
}

// ── Render tabla de diferencias ────────────────────────────────────────
function _tasasRenderTabla(filas, modo) {
  if (!filas.length) {
    return `<div style="display:flex;align-items:center;justify-content:center;
      height:80px;color:var(--grn);font-size:10px">
      ✓ Sin diferencias encontradas
    </div>`;
  }

  const pctFmt  = v => `${(v * 100).toFixed(4).replace(/\.?0+$/, '')}%`;
  const montoFmt = v => typeof _liqFmtARS === 'function' ? _liqFmtARS(v) : `$${Math.abs(v).toFixed(2)}`;

  const rows = filas.map(x => {
    const { fila, liqRow, td_cobrada, difPct, difMonto } = x;
    const td_ref = modo === 'pd' ? x.td_fact : x.td_ac;
    const sky    = fila.sky || {};
    const dif$_cls = difMonto > 0 ? 'color:var(--red);font-weight:700'
                                  : 'color:var(--grn);font-weight:700';
    return `<tr style="border-bottom:1px solid var(--b1)">
      <td>${liqRow.fecha_venta || '—'}</td>
      <td>${liqRow.equipo      || '—'}</td>
      <td>${sky.suc            || '—'}</td>
      <td>${liqRow.lote        || '—'}</td>
      <td>${liqRow.cupon       || '—'}</td>
      <td>${liqRow.tarjeta     || '—'}</td>
      <td style="text-align:center">${liqRow.cuotas || 1}</td>
      <td style="text-align:right">${montoFmt(liqRow.monto)}</td>
      <td>${liqRow.nro_comercio || '—'}</td>
      <td style="text-align:right;font-weight:700;color:var(--yel)">${pctFmt(td_cobrada)}</td>
      <td>${sky.nroCom         || '—'}</td>
      <td>${sky.tarjeta        || '—'}</td>
      <td>${sky.plan           || (sky.cuotas ? sky.cuotas + ' cuotas' : '—')}</td>
      <td style="text-align:right;color:var(--acc)">${pctFmt(td_ref)}</td>
      <td style="text-align:right;${dif$_cls.replace('font-weight:700','')}">${pctFmt(difPct)}</td>
      <td style="text-align:right;${dif$_cls}">${montoFmt(difMonto)}</td>
      <td>${sky.vendedor       || '—'}</td>
      <td>${sky.opId           || '—'}</td>
      <td>${sky.opNum          || '—'}</td>
      <td style="font-size:9px;color:${sky.integrado ? 'var(--grn)' : 'var(--m2)'}">
        ${sky.integrado ? 'INTEGRADO' : 'DESINTEGRADO'}
      </td>
    </tr>`;
  }).join('');

  const th = s => `<th style="padding:5px 8px;text-align:left;color:var(--m2);
    font-weight:600;white-space:nowrap;border-bottom:1px solid var(--b1)">${s}</th>`;
  const thr = s => `<th style="padding:5px 8px;text-align:right;color:var(--m2);
    font-weight:600;white-space:nowrap;border-bottom:1px solid var(--b1)">${s}</th>`;
  const thc = s => `<th style="padding:5px 8px;text-align:center;color:var(--m2);
    font-weight:600;white-space:nowrap;border-bottom:1px solid var(--b1)">${s}</th>`;

  return `<div style="overflow:auto;flex:1">
    <table style="width:100%;border-collapse:collapse;font-size:9px">
      <thead><tr style="background:var(--s2)">
        ${th('Fecha Venta')}${th('Nro Equipo')}${thc('SUC')}
        ${th('Lote')}${th('Cupón')}${th('Tarjeta Cobrada')}${thc('Plan<br>Cobrado')}
        ${thr('Importe Venta')}${th('Nro Com.<br>Cobrado')}${thr('Tasa Cobrada')}
        ${th('Nro Com.<br>Facturado')}${th('Tarjeta Facturada')}${th('Plan Facturado')}
        ${thr(modo === 'pd' ? 'TASA Facturada' : 'TASA Acordada')}
        ${thr('DIF %')}${thr('DIF $')}
        ${th('Vendedor')}${th('ID')}${th('N° FC')}${thc('Integrado')}
      </tr></thead>
      <tbody style="font-family:var(--mono)">${rows}</tbody>
    </table>
  </div>`;
}

// ── Exportar a Excel ───────────────────────────────────────────────────
function _tasasExportarSinTasa(filas) {
  if (!filas || !filas.length) { _showToast && _showToast('Sin filas para exportar'); return; }
  const pctFmt = v => `${(v * 100).toFixed(4).replace(/\.?0+$/, '')}%`;
  const data = filas.map(x => {
    const { fila, liqRow, td_cobrada, liqTarjeta, liqCuotas, skyTarjeta, skyCuotas } = x;
    const sky = fila?.sky || {};
    return {
      'Fecha Venta':            liqRow?.fecha_venta || '',
      'Nro Equipo':             liqRow?.equipo || '',
      'SUC':                    sky.suc || '',
      'Nro de Lote':            liqRow?.lote || '',
      'Nro de Cupón':           liqRow?.cupon || '',
      'Tarjeta Cobrada':        liqTarjeta || '',
      'Plan Cobrado (cuotas)':  liqCuotas || 1,
      'Importe Venta':          liqRow?.monto || 0,
      'Tasa Cobrada':           pctFmt(td_cobrada),
      'Tarjeta Facturada':      skyTarjeta || '',
      'Plan Facturado (cuotas)':skyCuotas || 1,
      'VENDEDOR':               sky.vendedor || '',
      'ID':                     sky.opId || '',
      'N° FC':                  sky.opNum || '',
      'MOTIVO':                 'Sin tasa acordada en TM',
    };
  });
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, 'Sin tasa config');
  XLSX.writeFile(wb, `sin_tasa_configurada_${new Date().toISOString().slice(0,10)}.xlsx`);
}

function _tasasExportar(filas, modo) {
  if (!filas.length) { _showToast && _showToast('Sin filas para exportar'); return; }
  const pctFmt = v => `${(v * 100).toFixed(4).replace(/\.?0+$/, '')}%`;
  const data = filas.map(x => {
    const { fila, liqRow, td_cobrada, difPct, difMonto } = x;
    const td_ref = modo === 'pd' ? x.td_fact : x.td_ac;
    const sky    = fila.sky || {};
    return {
      'Fecha Venta':         liqRow.fecha_venta || '',
      'Nro Equipo':          liqRow.equipo || '',
      'SUC':                 sky.suc || '',
      'Nro de Lote':         liqRow.lote || '',
      'Nro de Cupón':        liqRow.cupon || '',
      'Tarjeta Cobrada':     liqRow.tarjeta || '',
      'Plan Cobrado':        liqRow.cuotas || 1,
      'Importe Venta':       liqRow.monto || 0,
      'Nro Comercio Cobrado':liqRow.nro_comercio || '',
      'Tasa Cobrada':        td_cobrada,
      'Nro Comercio Facturado': sky.nroCom || '',
      'Tarjeta Facturada':   sky.tarjeta || '',
      'Plan Facturado':      sky.plan || '',
      [modo === 'pd' ? 'TASA Facturada' : 'TASA Acordada']: td_ref,
      'DIF %':               difPct,
      'DIF $':               difMonto,
      'VENDEDOR':            sky.vendedor || '',
      'ID':                  sky.opId || '',
      'N° FC':               sky.opNum || '',
      'INTEGRADO':           sky.integrado ? 'INTEGRADO' : 'DESINTEGRADO',
    };
  });
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, modo === 'pd' ? 'Pasar Descuento' : 'Reclamar Proc');
  const fname = `diferencias_tasa_${modo === 'pd' ? 'vendedor' : 'procesadora'}_${new Date().toISOString().slice(0,10)}.xlsx`;
  XLSX.writeFile(wb, fname);
}

// ── Render panel TASAS ────────────────────────────────────────────────
function renderModuloLiqTasas() {
  const panel = document.getElementById('mod-liq-tasas');
  if (!panel) return;

  if (!_LIQ_CUPONES.length) {
    panel.innerHTML = _liqEmpty('Tasas', '📊',
      'Cargá el archivo de Liquidaciones (FISERV/GETPOS) para analizar diferencias de tasa.');
    return;
  }

  if (!TM?.tasas?.length) {
    panel.innerHTML = `<div style="display:flex;flex-direction:column;align-items:center;
      justify-content:center;height:100%;gap:16px;color:var(--m2);text-align:center;padding:40px">
      <div style="font-size:36px;opacity:.15">⚙</div>
      <div style="font-size:14px;font-weight:700;color:var(--txt);opacity:.5">
        Sin tasas configuradas
      </div>
      <p style="font-size:10px;max-width:420px;line-height:1.8">
        Configurá las tasas acordadas en<br>
        <b style="color:var(--acc)">Configuración → Tablas Maestras → Tasas / Acuerdos</b><br>
        con los campos: Procesadora · Tarjeta · Cuotas · Tasa% · Desde · Hasta
      </p>
      <button onclick="showMegaTab('configuracion',document.getElementById('mega-configuracion'));
        setTimeout(()=>showTM('tasas'),200)"
        style="background:linear-gradient(135deg,var(--acc),var(--cyn));color:#fff;border:none;
          border-radius:6px;padding:8px 20px;font-size:10px;font-weight:700;cursor:pointer;
          font-family:var(--sans)">
        Ir a Tasas / Acuerdos →
      </button>
    </div>`;
    return;
  }

  if (typeof RESULTADO === 'undefined' || !RESULTADO.length) {
    panel.innerHTML = _liqEmpty('Tasas', '📊',
      'Ejecutá el cruce de OPERACIONES primero para analizar diferencias de tasa.');
    return;
  }

  const cruce = _cruzarTasas();
  if (!cruce) {
    panel.innerHTML = _liqEmpty('Tasas', '📊', 'No se encontraron operaciones para analizar.');
    return;
  }

  const { pasarDescuento, reclamarProc, sinTasa } = cruce;

  // ── Estado activo ────────────────────────────────────────────────────
  if (!window._liqTasasProc) window._liqTasasProc = 'fiserv';
  if (!window._liqTasasTab)  window._liqTasasTab  = 'pd';

  // ── Filtrar por procesadora ──────────────────────────────────────────
  const byProc = (arr, p) => p === 'getpos'
    ? arr.filter(x => x.procNom === 'GETPOS')
    : arr.filter(x => x.procNom !== 'GETPOS');

  const pd_fis = byProc(pasarDescuento, 'fiserv');
  const rp_fis = byProc(reclamarProc,   'fiserv');
  const st_fis = byProc(sinTasa,        'fiserv');
  const pd_gp  = byProc(pasarDescuento, 'getpos');
  const rp_gp  = byProc(reclamarProc,   'getpos');
  const st_gp  = byProc(sinTasa,        'getpos');

  const hasGoc = (RESULTADO?.filter(r => r.sky?.esGOCUOTAS).length || 0) > 0;

  const _kpis = (pd, rp, st, liqTotal) => {
    const mPD     = pd.reduce((s, x) => s + Math.abs(x.difMonto), 0);
    const mRP     = rp.reduce((s, x) => s + Math.abs(x.difMonto), 0);
    const analizadas = pd.length + rp.length + st.length; // excluye _okCount (sin dif ni sin tasa)
    const conDif  = pd.length + rp.length;
    const pctCob  = liqTotal ? Math.round(analizadas / liqTotal * 100) : null;
    const pctDif  = liqTotal ? Math.round(conDif / liqTotal * 100) : null;
    const kpiCard = (label, val, sub, bc, cls, extra='') => `
      <div style="min-width:130px;flex:1;background:var(--s2);border:1px solid ${bc};
        border-radius:6px;padding:10px 14px">
        <div style="font-size:9px;color:var(--m2);margin-bottom:4px">${label}</div>
        <div style="font-size:18px;font-weight:800;color:${cls?`var(--${cls})`:'var(--txt)'};
          font-family:var(--mono)">${val}</div>
        <div style="font-size:9px;color:var(--m2);margin-top:2px">${sub}</div>
        ${extra}
      </div>`;
    return [
      // KPI cobertura: analizadas vs liquidadas
      kpiCard(
        '📋 Liquidadas analizadas',
        liqTotal ? `${analizadas.toLocaleString('es-AR')} / ${liqTotal.toLocaleString('es-AR')}` : analizadas.toLocaleString('es-AR'),
        liqTotal
          ? `${pctCob}% tienen CFO · ${conDif.toLocaleString('es-AR')} con dif. (${pctDif}%)`
          : `${conDif.toLocaleString('es-AR')} con diferencia de tasa`,
        pctDif > 10 ? 'rgba(248,113,113,.25)' : 'rgba(79,142,247,.3)',
        'cyn'
      ),
      kpiCard('⚠ Error vendedor',   pd.length.toLocaleString('es-AR'), _liqFmtARS(mPD),       'rgba(251,191,36,.3)',  'yel'),
      kpiCard('🔴 Reclamar proc.',  rp.length.toLocaleString('es-AR'), _liqFmtARS(mRP),       'rgba(248,113,113,.3)', 'red'),
      kpiCard('? Sin tasa config.', st.length.toLocaleString('es-AR'), 'Sin match en TM',     'rgba(107,114,128,.2)', ''),
      kpiCard('💰 Total dif. $',    _liqFmtARS(mPD + mRP),            `${conDif.toLocaleString('es-AR')} ops con diferencia`, 'rgba(251,146,60,.25)', 'org'),
    ].join('');
  };

  // ── Render sub-tab (pd/rp) ───────────────────────────────────────────
  const renderSubTab = (tab, pd, rp, st) => {
    window._liqTasasTab = tab;
    const contenido = document.getElementById('liq-tasas-body');
    if (!contenido) return;
    const filas   = tab === 'pd' ? pd : rp;
    const expKey  = tab === 'pd' ? 'window._cruzeTasasPD' : 'window._cruzeTasasRP';
    const pdActive = tab === 'pd';
    contenido.innerHTML = `
      <!-- sub-tabs -->
      <div style="display:flex;gap:6px;padding:8px 14px;flex-shrink:0;border-bottom:1px solid var(--b1)">
        <button id="liq-tasas-tab-pd" onclick="renderModuloLiqTasas._renderSubTab('pd')"
          style="background:none;border:1px solid ${pdActive?'var(--yel)':'var(--b2)'};
            color:${pdActive?'var(--yel)':'var(--m1)'};border-radius:4px;padding:4px 14px;
            font-size:9px;font-weight:${pdActive?'700':'400'};cursor:pointer;font-family:var(--sans)">
          ⚠ Pasar a descuento (${pd.length})
        </button>
        <button id="liq-tasas-tab-rp" onclick="renderModuloLiqTasas._renderSubTab('rp')"
          style="background:none;border:1px solid ${!pdActive?'var(--red)':'var(--b2)'};
            color:${!pdActive?'var(--red)':'var(--m1)'};border-radius:4px;padding:4px 14px;
            font-size:9px;font-weight:${!pdActive?'700':'400'};cursor:pointer;font-family:var(--sans)">
          🔴 Reclamar a procesadora (${rp.length})
        </button>
        ${st.length ? `<div style="margin-left:auto;display:flex;align-items:center;gap:6px">
          <span style="font-size:9px;color:var(--m2)">⚙ ${st.length.toLocaleString('es-AR')} sin tasa config.</span>
          <button onclick="_tasasExportarSinTasa(window._cruzeTasasST)"
            style="background:none;border:1px solid var(--b2);color:var(--m1);border-radius:4px;
              padding:3px 10px;font-size:9px;cursor:pointer;font-family:var(--sans)">
            ↓ Exportar Excel
          </button>
        </div>` : ''}
      </div>
      <!-- tabla -->
      <div style="display:flex;align-items:center;justify-content:flex-end;
        padding:5px 12px;gap:8px;flex-shrink:0;border-bottom:1px solid var(--b0)">
        <span style="font-size:9px;color:var(--m2)">${filas.length} operaciones</span>
        <button onclick="_tasasExportar(${expKey},'${tab}')"
          style="background:none;border:1px solid var(--b2);color:var(--m1);border-radius:4px;
            padding:3px 10px;font-size:9px;cursor:pointer;font-family:var(--sans)">
          ↓ Exportar Excel
        </button>
      </div>
      <div style="flex:1;min-height:0;overflow:hidden">
        ${_tasasRenderTabla(filas, tab)}
      </div>`;
  };

  // ── Render proc tab ──────────────────────────────────────────────────
  const renderProcTab = proc => {
    window._liqTasasProc = proc;
    const area = document.getElementById('liq-tasas-proc-area');
    if (!area) return;

    // Actualizar estilos de botones del selector
    ['fiserv','getpos','goc'].forEach(p => {
      const btn = document.getElementById(`liq-tasas-proc-${p}`);
      if (!btn) return;
      const active = p === proc;
      btn.style.borderColor = active ? 'var(--acc)' : 'var(--b2)';
      btn.style.color       = active ? 'var(--acc)' : 'var(--m2)';
      btn.style.fontWeight  = active ? '700' : '400';
    });

    if (proc === 'goc') {
      area.innerHTML = `
        <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;
          height:100%;gap:12px;color:var(--m2);text-align:center;padding:40px">
          <div style="font-size:32px;opacity:.15">💳</div>
          <div style="font-size:12px;font-weight:700;color:var(--txt);opacity:.4">
            Análisis de tasas Go Cuotas
          </div>
          <p style="font-size:10px;max-width:380px;line-height:1.8;color:var(--m2)">
            El módulo de tasas para Go Cuotas está en desarrollo.<br>
            Las operaciones GoC usan una estructura de aranceles distinta<br>
            a la de FISERV / GETPOS.
          </p>
        </div>`;
      return;
    }

    const pd = proc === 'getpos' ? pd_gp : pd_fis;
    const rp = proc === 'getpos' ? rp_gp : rp_fis;
    const st = proc === 'getpos' ? st_gp : st_fis;
    const liqTotal = proc === 'getpos'
      ? (_liqCache.getpos?.liquidadas?.length || 0)
      : (_liqCache.fiserv?.liquidadas?.length || 0);

    // Guardar refs para exportar (filtradas por proc)
    window._cruzeTasasPD = pd;
    window._cruzeTasasRP = rp;
    window._cruzeTasasST = st;

    area.innerHTML = `
      <!-- KPIs del proc -->
      <div style="display:flex;gap:10px;padding:12px 16px;flex-shrink:0;flex-wrap:wrap;
        border-bottom:1px solid var(--b1)">
        ${_kpis(pd, rp, st, liqTotal)}
      </div>
      <!-- Sub-tabs body -->
      <div id="liq-tasas-body" style="display:flex;flex-direction:column;flex:1;min-height:0;overflow:hidden"></div>`;

    renderModuloLiqTasas._renderSubTab = (tab) => renderSubTab(tab, pd, rp, st);
    renderSubTab(window._liqTasasTab || 'pd', pd, rp, st);
  };

  // ── HTML principal ───────────────────────────────────────────────────
  const fisFiserv = RESULTADO.filter(r => !r.sky?.esGETPos && !r.sky?.esGOCUOTAS).length;
  const fisGetpos = RESULTADO.filter(r => r.sky?.esGETPos && !r.sky?.esGOCUOTAS).length;
  const fisGoc    = RESULTADO.filter(r => r.sky?.esGOCUOTAS).length;

  panel.innerHTML = `
  <div style="display:flex;flex-direction:column;height:100%;overflow:hidden">

    <!-- Selector de procesadora -->
    <div style="display:flex;align-items:center;gap:8px;padding:8px 14px;flex-shrink:0;
      border-bottom:2px solid var(--b1);background:var(--s1)">
      <span style="font-size:9px;color:var(--m2);margin-right:4px">Procesadora:</span>
      <button id="liq-tasas-proc-fiserv" onclick="renderModuloLiqTasas._renderProcTab('fiserv')"
        style="background:none;border:1px solid var(--acc);color:var(--acc);border-radius:4px;
          padding:4px 16px;font-size:9px;font-weight:700;cursor:pointer;font-family:var(--sans)">
        FISERV${pd_fis.length+rp_fis.length ? ` · ${(pd_fis.length+rp_fis.length).toLocaleString('es-AR')} dif.` : ''}
      </button>
      <button id="liq-tasas-proc-getpos" onclick="renderModuloLiqTasas._renderProcTab('getpos')"
        style="background:none;border:1px solid var(--b2);color:var(--m2);border-radius:4px;
          padding:4px 16px;font-size:9px;font-weight:400;cursor:pointer;font-family:var(--sans)">
        GETPOS${pd_gp.length+rp_gp.length ? ` · ${(pd_gp.length+rp_gp.length).toLocaleString('es-AR')} dif.` : ''}
      </button>
      ${hasGoc ? `<button id="liq-tasas-proc-goc" onclick="renderModuloLiqTasas._renderProcTab('goc')"
        style="background:none;border:1px solid var(--b2);color:var(--m2);border-radius:4px;
          padding:4px 16px;font-size:9px;font-weight:400;cursor:pointer;font-family:var(--sans)">
        Go Cuotas
      </button>` : ''}
    </div>

    <!-- Área dinámica por procesadora -->
    <div id="liq-tasas-proc-area" style="display:flex;flex-direction:column;flex:1;min-height:0;overflow:hidden"></div>
  </div>`;

  // Exponer funciones
  renderModuloLiqTasas._renderProcTab = renderProcTab;
  renderModuloLiqTasas._renderSubTab  = () => {};

  // Restaurar estado
  const procActivo = window._liqTasasProc || 'fiserv';
  renderProcTab(procActivo);
}
