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
// ── Búsqueda fuzzy de columna (tolera variaciones de nombre, tildes, case) ──
function _liqCol(r, ...names) {
  for (const n of names) {
    if (r[n] !== undefined && r[n] !== null) return r[n];
  }
  const rkeys = Object.keys(r);
  const norm = s => s.toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g,'').replace(/[^a-z0-9]/g,' ').trim();
  for (const n of names) {
    const nn = norm(n);
    const found = rkeys.find(k => norm(k) === nn);
    if (found !== undefined && r[found] !== undefined && r[found] !== null) return r[found];
  }
  return null;
}

function parseLiquidaciones(wb) {
  const ws   = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: false });

  // Log primeras columnas para diagnóstico
  if (rows.length > 0) {
    console.log('[LIQ] Columnas detectadas:', Object.keys(rows[0]).slice(0, 20));
    console.log('[LIQ] Primera fila muestra:', rows[0]);
  }

  _LIQ_NORM = rows
    .filter(r => {
      // Aceptar cualquier fila con al menos un campo identificador de operación
      // (equipo, cupón, auth o lote). Excluye filas completamente vacías.
      const eq  = _liqCol(r, 'Nro Equipo', 'Equipo', 'Terminal', 'Nro de Equipo');
      const cup = _liqCol(r, 'Nro de Cupón', 'Nro Cupon', 'Nro de Cupon', 'Cupon', 'Cupón');
      const aut = _liqCol(r, 'Código Autorización', 'Codigo Autorizacion', 'Cod Autorizacion',
                              'Cod. Autorizacion', 'Código de Autorización', 'Auth');
      const lot = _liqCol(r, 'Nro de Lote', 'Nro Lote', 'Lote');
      return (eq != null) || (cup != null && cup !== '') || (aut != null && aut !== '') || (lot != null);
    })
    .map((r, i) => {
      const tarjeta = String(_liqCol(r, 'Tarjeta') || '').trim();
      const impRaw  = _liqCol(r, 'Importe Venta', 'Importe de Venta', 'Importe', 'Monto', 'Monto Venta');
      return {
        idx:          i,
        fechaVenta:   String(_liqCol(r, 'Fecha Venta', 'Fecha de Venta', 'Fecha Operacion') || '').trim().slice(0, 10),
        fechaPago:    String(_liqCol(r, 'Fecha Pago', 'Fecha de Pago')   || '').trim().slice(0, 10),
        fechaAdelanto:String(_liqCol(r, 'Fecha Adelanto')                || '').trim().slice(0, 10),
        nroLiq:       String(_liqCol(r, 'Nro Liquidación', 'Nro Liquidacion', 'Nro de Liquidacion', 'Liquidacion') || '').trim(),
        equipo:       normNum(_liqCol(r, 'Nro Equipo', 'Equipo', 'Terminal', 'Nro de Equipo')),
        nombreEquipo: String(_liqCol(r, 'Nombre de equipo', 'Nombre Equipo', 'Nombre del Equipo') || '').trim(),
        lote:         normNum(_liqCol(r, 'Nro de Lote', 'Nro Lote', 'Lote')),
        cupon:        normNum(_liqCol(r, 'Nro de Cupón', 'Nro Cupon', 'Nro de Cupon', 'Cupon', 'Cupón')),
        tarjeta,
        nroTarjeta:   String(_liqCol(r, 'Nro Tarjeta', 'Número de Tarjeta', 'Numero Tarjeta') || '').trim(),
        aut:          normNum(_liqCol(r, 'Código Autorización', 'Codigo Autorizacion',
                                        'Cod. Autorizacion', 'Código de Autorización', 'Auth', 'Autorizacion')),
        cuotas:       parseInt(_liqCol(r, 'Cuotas', 'Nro Cuotas')) || 1,
        importe:      (() => {
          const s = String(impRaw || '').replace(/\./g, '').replace(',', '.');
          return Math.abs(parseFloat(s) || 0);
        })(),
        nroCom:       normNum(_liqCol(r, 'Nro Comercio', 'Nro de Comercio', 'Numero Comercio')),
        banco:        String(_liqCol(r, 'Banco Pagador', 'Banco')        || '').trim(),
        rechazo:      String(_liqCol(r, 'Rechazo')                       || 'N').trim().toUpperCase(),
        arancel:      parseFloat(_liqCol(r, 'Arancel'))                  || 0,
        ivaArancel:   parseFloat(_liqCol(r, 'IVA Arancel', 'Iva Arancel')) || 0,
        cfo:          parseFloat(_liqCol(r, 'CFO'))                      || 0,
        ivaCfo:       parseFloat(_liqCol(r, 'Iva CFO', 'IVA CFO'))       || 0,
        tipoOp:       String(_liqCol(r, 'Tipo operacion', 'Tipo Operacion', 'Tipo de Operacion') || '').trim(),
        bancoEmisor:  String(_liqCol(r, 'Banco Emisor')                  || '').trim(),
        // Procesadora inferida del nombre de tarjeta
        proc: tarjeta.toUpperCase().includes('GETNET') ? 'GETPOS' : 'FISERV',
      };
    });

  _buildLiqIdx();
  return _LIQ_NORM;
}

// ── Normalizar monto para clave hash (sin decimales, valor absoluto) ─
function _normM(v) { return String(Math.round(Math.abs(parseFloat(v) || 0))); }

// ── Descripción del código único para mostrar / exportar ─────────────
function _codigoLiq(liq) {
  return liq.proc === 'GETPOS'
    ? `GP_${liq.equipo}_${liq.aut}`
    : `FIS_${liq.equipo}_${liq.lote}_${liq.cupon}`;
}

// ══════════════════════════════════════════════════════════════════
// ÍNDICE MULTI-CLAVE (FISERV y GETPOS tienen estructuras distintas)
// ══════════════════════════════════════════════════════════════════
// GETPOS:
//   • Nro Equipo (liquidación) = Código del POS (procesadora) → coincide ✓
//   • Nro de Lote (liquidación) ≠ Lote (procesadora)          → NO coincide ✗
//   • Código Autorización      = auth de la procesadora       → coincide ✓
//   Clave primaria: GP_equipo_auth
//
// FISERV:
//   • Nro Equipo (liquidación) ≠ equipo de la procesadora     → NO coincide ✗
//     (son dos sistemas de ID distintos: 521171 vs 19383953)
//   • Nro de Lote  (liquidación) ≠ Nro de Lote (procesadora)  → NO coincide ✗
//   • Nro de Cupón (liquidación) = Nro de Cupón (procesadora) → coincide ✓
//   • Código Autorización       = auth de la procesadora      → coincide ✓
//   Clave primaria: FIS_A_auth_monto  (auth es asignado por Visa/MC, globalmente único)
//   Clave secundaria: FIS_C_cupon_monto  (cupon = Nro de Cupón)
// ══════════════════════════════════════════════════════════════════
function _buildLiqIdx() {
  _LIQ_IDX = {};

  const add = (key, liq) => {
    if (!key) return;
    // Descartar claves trivialmente degeneradas (todo ceros)
    const parts = key.split('_');
    if (parts.slice(1).every(p => p === '0')) return;
    (_LIQ_IDX[key] = _LIQ_IDX[key] || []).push(liq);
  };

  for (const liq of _LIQ_NORM) {
    const eq  = liq.equipo;
    const lot = liq.lote;
    const cup = liq.cupon;
    const aut = liq.aut;
    const mon = _normM(liq.importe);

    if (liq.proc === 'GETPOS') {
      // ── Claves GETPOS ───────────────────────────────────────────
      // Primaria: equipo + auth (el auth del GETPOS = Código Autorización de liq)
      if (aut && aut !== '0') {
        add(`GP_${eq}_${aut}`,             liq);
        add(`GP_${eq}_${aut}_${mon}`,      liq);
        add(`GP_${aut}_${mon}`,            liq);
        add(`GP_${aut}`,                   liq);
      }
      // Fallback: equipo + cupon (cuando cupon ≠ auth)
      if (cup && cup !== '0') {
        add(`GP_${eq}_CUP_${cup}`,         liq);
        add(`GP_${eq}_CUP_${cup}_${mon}`,  liq);
      }
    } else {
      // ── Claves FISERV ───────────────────────────────────────────
      // El equipo y lote de la liquidación ≠ equipo/lote de la procesadora.
      // La clave más confiable es el Código de Autorización (red Visa/MC).
      // También el Nro de Cupón puede coincidir con el ticket de la procesadora.

      // L1 – Auth-based (más confiable: asignado por la red de tarjetas)
      if (aut && aut !== '0') {
        add(`FIS_A_${aut}_${mon}`,         liq);  // auth + monto (primaria)
        add(`FIS_A_${aut}`,                liq);  // auth solo    (tolerancia)
      }
      // L2 – Cupón + monto (Nro de Cupón coincide entre procesadora y liquidación)
      if (cup && cup !== '0') {
        add(`FIS_C_${cup}_${mon}`,         liq);
        add(`FIS_C_${cup}`,                liq);
      }
      // L3 – Claves tradicionales (por si coinciden equipo/lote en algún caso)
      if (cup && cup !== '0' && lot && lot !== '0') {
        add(`FIS_${eq}_${lot}_${cup}`,     liq);
        add(`FIS_${lot}_${cup}`,           liq);
      }
      if (aut && aut !== '0') {
        add(`FIS_${eq}_${aut}`,            liq);
        add(`FIS_${eq}_${aut}_${mon}`,     liq);
      }
    }
  }
}

// ══════════════════════════════════════════════════════════════════
// CLAVES DE BÚSQUEDA PARA UNA FILA DE RESULTADO
// ══════════════════════════════════════════════════════════════════
function _clavesDeRec(r) {
  const p   = r.proc;
  const cor = typeof CORREGIDAS !== 'undefined' ? CORREGIDAS[r.sky.idx] : null;

  if (p) {
    const equipo = normNum(p.equipo || p.pos || '');
    const lote   = normNum(p.lote   || r.sky.lote || '');
    const ticket = normNum(p.ticket || '');   // solo FISERV tiene ticket
    const cupon  = normNum(p.cupon  || '');   // GETPOS usa cupon
    const aut    = normNum(p.aut    || '');
    const mon    = _normM(r.sky.monto);
    const esFIS  = r.procEncontrada === 'FISERV' || (!r.procEncontrada && !!p.ticket);

    if (esFIS) {
      // ── Cascada FISERV ──────────────────────────────────────────
      // El equipo y lote del archivo procesadora ≠ equipo/lote de la liquidación.
      // Prioridad:
      //   L1 – auth + monto  (Código Autorización: globalmente único, mismo en ambos archivos)
      //   L2 – auth solo     (tolerancia si monto difiere por centavos)
      //   L3 – ticket + monto (Nro de Cupón coincide entre procesadora y liquidación)
      //   L4 – ticket solo
      //   L5-L8 – claves tradicionales equipo/lote (por si acaso coinciden)
      return {
        proc:  'FISERV',
        keys: [
          // L1-L2: auth-based (más confiable)
          aut    !== '0'                   ? `FIS_A_${aut}_${mon}`            : null,
          aut    !== '0'                   ? `FIS_A_${aut}`                   : null,
          // L3-L4: ticket/cupon + monto
          ticket !== '0'                   ? `FIS_C_${ticket}_${mon}`         : null,
          ticket !== '0'                   ? `FIS_C_${ticket}`                : null,
          // L5-L8: claves tradicionales (equipo+lote+ticket, por si coinciden)
          ticket !== '0' && equipo !== '0' ? `FIS_${equipo}_${lote}_${ticket}` : null,
          ticket !== '0'                   ? `FIS_${lote}_${ticket}`           : null,
          aut    !== '0' && equipo !== '0' ? `FIS_${equipo}_${aut}`            : null,
          aut    !== '0' && equipo !== '0' ? `FIS_${equipo}_${aut}_${mon}`     : null,
        ].filter(Boolean),
        label: aut !== '0'
          ? `FIS_A_${aut}_${mon}`
          : `FIS_${lote}_${ticket}`,
      };
    } else {
      // ── Cascada GETPOS ──────────────────────────────────────────
      // Para GETPOS el auth del procesador = Código Autorización de la liquidación
      // El "cupon" GETPOS también puede ser el mismo valor que aut
      // El equipo (pos) del GETPOS = Nro Equipo de la liquidación
      // El lote GETPOS ≠ Nro de Lote de la liquidación (numeración diferente)
      const autGP  = aut !== '0' ? aut : cupon;   // aut o cupon como auth
      const cuponGP = cupon !== '0' ? cupon : aut;
      return {
        proc:  'GETPOS',
        keys: [
          autGP  !== '0' && equipo !== '0' ? `GP_${equipo}_${autGP}_${mon}`       : null,
          autGP  !== '0' && equipo !== '0' ? `GP_${equipo}_${autGP}`              : null,
          autGP  !== '0'                   ? `GP_${autGP}_${mon}`                 : null,
          autGP  !== '0'                   ? `GP_${autGP}`                        : null,
          // Intento con cupon como campo cupón (por si ≠ auth en liq)
          cuponGP !== '0' && cuponGP !== autGP && equipo !== '0'
            ? `GP_${equipo}_CUP_${cuponGP}`     : null,
          cuponGP !== '0' && cuponGP !== autGP && equipo !== '0'
            ? `GP_${equipo}_CUP_${cuponGP}_${mon}` : null,
        ].filter(Boolean),
        label: equipo !== '0' && autGP !== '0'
          ? `GP_${equipo}_${autGP}`
          : `GP_${autGP || cuponGP}`,
      };
    }
  }

  // ── Sin proc (corrección manual o SIN MATCH) ─────────────────────
  const lote  = normNum(r.sky.lote  || '');
  const cupon = normNum(cor?.cupon  || r.sky.cupon || '');
  const mon   = _normM(r.sky.monto);
  const esFallGP = r.sky.esGETPos;

  const prefix = esFallGP ? 'GP' : 'FIS';
  return {
    proc:  esFallGP ? 'GETPOS' : 'FISERV',
    keys: [
      `${prefix}_${lote}_${cupon}`,
      `${prefix}_${cupon}_${mon}`,
      `${prefix}_${cupon}`,
    ].filter(Boolean),
    label: `${prefix}_SKY_${lote}_${cupon}`,
  };
}

// ── Buscar liquidación para una fila de RESULTADO ───────────────────
function _buscarEnLiq(r) {
  if (!_LIQ_NORM.length) return null;

  const info = _clavesDeRec(r);
  if (!info) return null;

  for (const key of info.keys) {
    const hits = _LIQ_IDX[key];
    if (hits?.length) return hits[0];
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
    const info = _clavesDeRec(r);
    const liq  = _buscarEnLiq(r);

    let estado;
    if (!liq) {
      estado = 'PENDIENTE';
    } else if (liq.rechazo && liq.rechazo !== 'N' && liq.rechazo !== '') {
      estado = 'RECHAZADO';
    } else {
      estado = 'COBRADO';
    }

    // Fuente del código para display/filtros
    let fuenteCodigo;
    if (r.proc) {
      fuenteCodigo = info?.proc === 'GETPOS' ? 'GETPOS' : 'FISERV';
    } else if (typeof CORREGIDAS !== 'undefined' && CORREGIDAS[r.sky.idx]) {
      fuenteCodigo = 'MANUAL';
    } else {
      fuenteCodigo = 'SKY';
    }

    COBROS_RESULT.push({
      fila:        r,
      liq,
      estado,
      codigoProc:  info?.label  || '—',   // para depuración / exportación
      codigoLiq:   liq ? _codigoLiq(liq) : '—',
      fuenteCodigo,
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
