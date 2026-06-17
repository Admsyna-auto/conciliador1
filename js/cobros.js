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
// ── Normalizar nombre de columna para comparación fuzzy ──────────────
function _normColName(s) {
  const TILDES = { 'á':'a','é':'e','í':'i','ó':'o','ú':'u','ü':'u','ñ':'n',
                   'Á':'a','É':'e','Í':'i','Ó':'o','Ú':'u','Ü':'u','Ñ':'n' };
  return String(s)
    .split('').map(c => TILDES[c] || c).join('')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

// ── Resolver qué clave real del objeto corresponde a un campo ──────────
function _resolveKey(keys, ...candidates) {
  // 1. Exacto
  for (const c of candidates) {
    if (keys.includes(c)) return c;
  }
  // 2. Fuzzy (sin tildes, sin puntuación)
  const normKeys = keys.map(k => ({ k, n: _normColName(k) }));
  for (const c of candidates) {
    const nc = _normColName(c);
    const hit = normKeys.find(x => x.n === nc);
    if (hit) return hit.k;
  }
  return null;
}

function parseLiquidaciones(wb) {
  const ws   = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: false });

  if (!rows.length) {
    console.warn('[LIQ] El archivo no tiene filas.');
    _LIQ_NORM = []; _buildLiqIdx(); return _LIQ_NORM;
  }

  // ── Detectar nombres de columna UNA VEZ desde la primera fila ──────
  const allKeys = Object.keys(rows[0]);
  console.log('[LIQ] Columnas del archivo:', allKeys);

  const K = {
    equipo:    _resolveKey(allKeys, 'Nro Equipo', 'Equipo', 'Terminal', 'Nro de Equipo'),
    lote:      _resolveKey(allKeys, 'Nro de Lote', 'Nro Lote', 'Lote'),
    cupon:     _resolveKey(allKeys, 'Nro de Cupón', 'Nro de Cupon', 'Nro Cupon', 'Cupon', 'Cupón'),
    aut:       _resolveKey(allKeys, 'Código Autorización', 'Codigo Autorizacion',
                                    'Código Autorizacion', 'Cod Autorizacion',
                                    'Cod. Autorizacion', 'Autorizacion', 'Auth'),
    importe:   _resolveKey(allKeys, 'Importe Venta', 'Importe de Venta', 'Importe', 'Monto', 'Monto Venta'),
    tarjeta:   _resolveKey(allKeys, 'Tarjeta'),
    fechaVenta:_resolveKey(allKeys, 'Fecha Venta', 'Fecha de Venta', 'Fecha Operacion', 'Fecha'),
    fechaPago: _resolveKey(allKeys, 'Fecha Pago', 'Fecha de Pago'),
    fechaAdel: _resolveKey(allKeys, 'Fecha Adelanto'),
    nroLiq:    _resolveKey(allKeys, 'Nro Liquidación', 'Nro Liquidacion', 'Nro de Liquidacion', 'Liquidacion'),
    nombreEq:  _resolveKey(allKeys, 'Nombre de equipo', 'Nombre Equipo', 'Nombre del Equipo', 'Nombre'),
    nroTarj:   _resolveKey(allKeys, 'Nro Tarjeta', 'Número de Tarjeta', 'Numero Tarjeta', 'Nro. Tarjeta'),
    cuotas:    _resolveKey(allKeys, 'Cuotas', 'Nro Cuotas'),
    nroCom:    _resolveKey(allKeys, 'Nro Comercio', 'Nro de Comercio', 'Numero Comercio'),
    banco:     _resolveKey(allKeys, 'Banco Pagador', 'Banco'),
    rechazo:   _resolveKey(allKeys, 'Rechazo'),
    arancel:   _resolveKey(allKeys, 'Arancel'),
    ivaArancel:_resolveKey(allKeys, 'IVA Arancel', 'Iva Arancel', 'IVA del Arancel'),
    cfo:       _resolveKey(allKeys, 'CFO'),
    ivaCfo:    _resolveKey(allKeys, 'Iva CFO', 'IVA CFO'),
    tipoOp:    _resolveKey(allKeys, 'Tipo operacion', 'Tipo Operacion', 'Tipo de Operacion'),
    bancoEm:   _resolveKey(allKeys, 'Banco Emisor'),
  };
  console.log('[LIQ] Mapeo de columnas resuelto:', K);

  // ── Getter seguro ────────────────────────────────────────────────────
  const g = (r, k) => (k && r[k] !== undefined) ? r[k] : null;

  // ── Mapear TODAS las filas, filtrar al final ─────────────────────────
  _LIQ_NORM = rows
    .map((r, i) => {
      const tarjeta = String(g(r, K.tarjeta) || '').trim();
      const impRaw  = String(g(r, K.importe) || '0')
        .replace(/\./g, '').replace(',', '.');
      return {
        idx:          i,
        fechaVenta:   String(g(r, K.fechaVenta)  || '').trim().slice(0, 10),
        fechaPago:    String(g(r, K.fechaPago)    || '').trim().slice(0, 10),
        fechaAdelanto:String(g(r, K.fechaAdel)    || '').trim().slice(0, 10),
        nroLiq:       String(g(r, K.nroLiq)       || '').trim(),
        equipo:       normNum(g(r, K.equipo)),
        nombreEquipo: String(g(r, K.nombreEq)     || '').trim(),
        lote:         normNum(g(r, K.lote)),
        cupon:        normNum(g(r, K.cupon)),
        tarjeta,
        nroTarjeta:   String(g(r, K.nroTarj)      || '').trim(),
        aut:          normNum(g(r, K.aut)),
        cuotas:       parseInt(g(r, K.cuotas))     || 1,
        importe:      Math.abs(parseFloat(impRaw)  || 0),
        nroCom:       normNum(g(r, K.nroCom)),
        banco:        String(g(r, K.banco)         || '').trim(),
        rechazo:      String(g(r, K.rechazo)       || 'N').trim().toUpperCase(),
        arancel:      parseFloat(g(r, K.arancel))  || 0,
        ivaArancel:   parseFloat(g(r, K.ivaArancel))|| 0,
        cfo:          parseFloat(g(r, K.cfo))      || 0,
        ivaCfo:       parseFloat(g(r, K.ivaCfo))   || 0,
        tipoOp:       String(g(r, K.tipoOp)        || '').trim(),
        bancoEmisor:  String(g(r, K.bancoEm)       || '').trim(),
        proc: tarjeta.toUpperCase().includes('GETNET') ? 'GETPOS' : 'FISERV',
      };
    })
    // Excluir filas sin ningún dato útil (filas vacías, totales, etc.)
    .filter(liq => liq.equipo !== '0' || liq.aut !== '0' || liq.cupon !== '0' || liq.importe > 0);

  console.log('[LIQ] Filas parseadas:', _LIQ_NORM.length);
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

  // Si hay corrección manual el usuario pudo haber ingresado un auth code
  // o un cupón → intentamos todos los formatos reales del índice de liquidaciones
  // para maximizar las chances de cruzar contra la liq.
  const keys = [];
  if (!esFallGP) {
    // FISERV: el código ingresado puede ser auth (FIS_A_) o cupón (FIS_C_)
    if (cupon !== '0') {
      keys.push(`FIS_A_${cupon}_${mon}`);   // como si fuera auth + monto
      keys.push(`FIS_A_${cupon}`);          // como si fuera auth solo
      keys.push(`FIS_C_${cupon}_${mon}`);   // como si fuera cupón + monto
      keys.push(`FIS_C_${cupon}`);          // como si fuera cupón solo
      keys.push(`FIS_${lote}_${cupon}`);    // tradicional lote + cupón
    }
  } else {
    // GETPOS: el código ingresado suele ser el auth (= Código Autorización de la liq)
    if (cupon !== '0') {
      keys.push(`GP_${cupon}_${mon}`);
      keys.push(`GP_${cupon}`);
      keys.push(`GP_${lote}_${cupon}`);
    }
  }
  // Fallback genérico por si no hubo match arriba
  keys.push(`${prefix}_${cupon}_${mon}`);
  keys.push(`${prefix}_${cupon}`);

  return {
    proc:  esFallGP ? 'GETPOS' : 'FISERV',
    keys:  keys.filter(Boolean),
    label: `${prefix}_MANUAL_${cupon}`,
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
// RENDER — MÓDULO PENDIENTES / EXTRAS
// Lee directamente de _liqCache (liquidaciones.js).
// Sin liquidar = ops Skylab no encontradas en el archivo liq.
// Extras       = filas del archivo liq sin op Skylab correspondiente.
// ══════════════════════════════════════════════════════════════════
let _cobrosTab  = 'sinliq';
let _cobProc    = '';   // '' | 'FISERV' | 'GETPOS' | 'GoC'
let _cobArrastre = [];  // pendientes cargados desde backup período anterior

// ══════════════════════════════════════════════════════════════════
// OPTION C — PLAZO DE ACREDITACIÓN
// ══════════════════════════════════════════════════════════════════
function _cobFechaAcreditacion(fechaVenta, proc, tarjeta) {
  if (!fechaVenta || !TM?.plazos?.length) return null;
  const up = s => String(s || '').toUpperCase().trim();
  const procUp = up(proc);
  const tarjUp = up(tarjeta);

  const reglas = TM.plazos.filter(r =>
    (!r.procesadora || up(r.procesadora) === procUp) &&
    (!r.tarjeta     || up(r.tarjeta)     === tarjUp)
  );
  if (!reglas.length) return null;

  // Más específica primero (proc + tarjeta > proc solo > genérica)
  const score = r => (r.procesadora ? 2 : 0) + (r.tarjeta ? 1 : 0);
  const rule = reglas.sort((a, b) => score(b) - score(a))[0];

  const diasHab = parseInt(rule.dias_habiles) || 0;
  if (!diasHab) return null;

  const feriados = new Set((TM.feriados || []).map(f => f.fecha).filter(Boolean));
  const esDiaHabil = d => {
    if (d.getDay() === 0 || d.getDay() === 6) return false;
    return !feriados.has(d.toISOString().slice(0, 10));
  };

  // Parsear fecha: YYYY-MM-DD o DD/MM/YYYY o DD-MM-YYYY
  const s = String(fechaVenta);
  const parts = s.split(/[-\/]/);
  let base;
  if (parts[0].length === 4) {
    base = new Date(+parts[0], +parts[1] - 1, +parts[2]);
  } else {
    base = new Date(+parts[2], +parts[1] - 1, +parts[0]);
  }
  if (isNaN(base.getTime())) return null;

  let cur = new Date(base);
  let added = 0;
  while (added < diasHab) {
    cur.setDate(cur.getDate() + 1);
    if (esDiaHabil(cur)) added++;
  }
  return { fechaAcred: cur.toISOString().slice(0, 10), diasHab };
}

function _cobPlazoCell(op) {
  const s = op.fila?.sky || op;  // soporta tanto sinLiq como arrastre
  const res = _cobFechaAcreditacion(s.fecha, op.proc, s.tarjeta);
  if (!res) return `<td style="color:var(--m2);font-size:9px">—</td>`;
  const hoy = new Date().toISOString().slice(0, 10);
  const dias = Math.round(Math.abs(new Date(hoy) - new Date(res.fechaAcred)) / 86400000);
  if (res.fechaAcred < hoy) {
    return `<td style="color:var(--red);font-size:9px;font-weight:600" title="Esperado: ${res.fechaAcred}">${res.fechaAcred}<br><span style="font-size:8px">+${dias}d vencido</span></td>`;
  }
  return `<td style="color:var(--grn);font-size:9px" title="${res.diasHab} días hábiles">${res.fechaAcred}<br><span style="font-size:8px">en ${dias}d</span></td>`;
}

// ══════════════════════════════════════════════════════════════════
// OPTION A — ARRASTRE DESDE BACKUP JSON
// ══════════════════════════════════════════════════════════════════
function _cobCargarArrastre(input) {
  const file = input.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const data = JSON.parse(e.target.result);
      const pend = data.pendientesArrastre;
      if (!Array.isArray(pend) || !pend.length) {
        alert('El archivo no contiene pendientes de arrastre.\n(Debe ser un backup de período generado con esta versión de la app)');
        return;
      }
      _cobArrastre = pend;
      if (typeof _showToast === 'function') _showToast(`✓ ${pend.length} pendientes cargados del período anterior`);
      renderModuloCobros();
    } catch(err) {
      alert('Error al leer el archivo: ' + err.message);
    }
  };
  reader.readAsText(file, 'utf-8');
}

function _cobRematcharArrastre() {
  const liqCups = (typeof _LIQ_CUPONES !== 'undefined') ? _LIQ_CUPONES : [];
  const norm  = v => String(v ?? '').replace(/^0+/, '') || '0';
  const normA = v => String(v ?? '').replace(/^['"]+/, '').trim().replace(/^0+/, '') || '0';

  return _cobArrastre.map(op => {
    const opAut   = normA(op.aut);
    const opCupon = norm(op.cupon);
    const opLote  = norm(op.lote);
    const opMonto = Math.round(Math.abs(parseFloat(op.monto) || 0));

    let liqRow = null;
    for (const r of liqCups) {
      const rAut   = normA(r.aut);
      const rCupon = norm(r.cupon);
      const rLote  = norm(r.lote);
      const rMonto = Math.round(Math.abs(parseFloat(r.monto) || 0));

      if (opAut !== '0' && rAut !== '0' && opAut === rAut && opMonto === rMonto) { liqRow = r; break; }
      if (opAut !== '0' && rAut !== '0' && opAut === rAut)                       { liqRow = r; break; }
      if (opCupon !== '0' && rCupon !== '0' && opCupon === rCupon && opLote === rLote) { liqRow = r; break; }
      if (opCupon !== '0' && rCupon !== '0' && opCupon === rCupon && opMonto === rMonto) { liqRow = r; break; }
    }
    return { op, liqRow, acreditoEsteMes: !!liqRow };
  });
}

function renderModuloCobros() {
  const cont = document.getElementById('mod-cobros');
  if (!cont) return;

  if (!RESULTADO.length) {
    cont.innerHTML = `<div class="cobros-empty">
      Primero realizá el <b style="color:var(--acc)">Cruce Automático</b>.</div>`;
    return;
  }

  // Auto-computar caches si hay archivo liq cargado y aún no se corrió el cruce
  if (typeof _LIQ_CUPONES !== 'undefined' && _LIQ_CUPONES.length) {
    if (!_liqCache.fiserv) _liqCache.fiserv = _cruzarLiqFiserv();
    if (!_liqCache.getpos) _liqCache.getpos = _cruzarLiqGetpos();
    if (!_liqCache.goc)    _liqCache.goc    = _cruzarLiqGoC();
  }

  const hayCache = _liqCache.fiserv || _liqCache.getpos || _liqCache.goc;
  if (!hayCache) {
    cont.innerHTML = `<div class="cobros-empty">
      Cargá el archivo de <b style="color:var(--yel)">Liquidaciones</b> y ejecutá el cruce
      en las pestañas de LIQUIDACIONES para ver pendientes y extras.</div>`;
    return;
  }

  // Consolidar sin liquidar y extras de todos los procesadores
  const sinLiq = [
    ...(_liqCache.fiserv?.noLiquidadas || []).map(x => ({ ...x, proc: 'FISERV' })),
    ...(_liqCache.getpos?.noLiquidadas || []).map(x => ({ ...x, proc: 'GETPOS' })),
    ...(_liqCache.goc?.noLiquidadas    || []).map(x => ({ ...x, proc: 'GoC'    })),
  ];
  const extras = [
    ...(_liqCache.fiserv?.extras || []).map(r => ({ liqRow: r, proc: 'FISERV' })),
    ...(_liqCache.getpos?.extras || []).map(r => ({ liqRow: r, proc: 'GETPOS' })),
    ...(_liqCache.goc?.extras    || []).map(r => ({ liqRow: r, proc: 'GoC'    })),
  ];

  const montoSinLiq = sinLiq.reduce((s, x) => s + Math.abs(x.fila?.sky?.monto || 0), 0);
  const montoExtras = extras.reduce((s, x) => s + Math.abs(x.liqRow?.monto || 0), 0);

  const badge = document.getElementById('mcnt-cobros');
  if (badge) badge.textContent = sinLiq.length || '0';

  cont.innerHTML = `
  <div class="cobros-wrap">
    <!-- KPIs -->
    <div class="cobros-kpis">
      <div class="cob-kpi" style="border-top:3px solid var(--red)">
        <div class="cob-kpi-lbl">Sin liquidar</div>
        <div class="cob-kpi-n" style="color:var(--red)">${sinLiq.length.toLocaleString('es-AR')}</div>
        <div class="cob-kpi-m" style="color:var(--red)">${fmtARS(montoSinLiq)}</div>
        <div class="cob-kpi-pct">ops sin acreditar en liq</div>
      </div>
      <div class="cob-kpi" style="border-top:3px solid var(--yel)">
        <div class="cob-kpi-lbl">Extras procesadora</div>
        <div class="cob-kpi-n" style="color:var(--yel)">${extras.length.toLocaleString('es-AR')}</div>
        <div class="cob-kpi-m" style="color:var(--yel)">${fmtARS(montoExtras)}</div>
        <div class="cob-kpi-pct">filas liq sin op en Skylab</div>
      </div>
      <div class="cob-kpi" style="border-top:3px solid var(--acc)">
        <div class="cob-kpi-lbl">FISERV sin liquidar</div>
        <div class="cob-kpi-n">${(_liqCache.fiserv?.noLiquidadas?.length ?? '—').toLocaleString?.() ?? '—'}</div>
        <div class="cob-kpi-m">${_liqCache.fiserv ? fmtARS(_liqCache.fiserv.montoNoLiq||0) : '—'}</div>
        <div class="cob-kpi-pct">${_liqCache.fiserv?.tieneLiq ? 'archivo cargado' : 'sin archivo liq'}</div>
      </div>
      <div class="cob-kpi" style="border-top:3px solid var(--vio)">
        <div class="cob-kpi-lbl">GETPOS sin liquidar</div>
        <div class="cob-kpi-n">${(_liqCache.getpos?.noLiquidadas?.length ?? '—').toLocaleString?.() ?? '—'}</div>
        <div class="cob-kpi-m">${_liqCache.getpos ? fmtARS(_liqCache.getpos.montoNoLiq||0) : '—'}</div>
        <div class="cob-kpi-pct">${_liqCache.getpos?.tieneLiq ? 'archivo cargado' : 'sin archivo liq'}</div>
      </div>
    </div>

    <!-- Tabs + botón arrastre -->
    <div class="tab-strip" id="tstrip-cobros"
         style="padding:0 20px;border-bottom:1px solid var(--b1);flex-shrink:0;display:flex;align-items:center;gap:4px">
      <button class="tb ${_cobrosTab==='sinliq'?'active':''}"
        onclick="showCobrosTab('sinliq',this)">
        Sin liquidar
        <span class="cnt" style="background:rgba(248,113,113,.15);color:var(--red)">${sinLiq.length.toLocaleString('es-AR')}</span>
      </button>
      <button class="tb ${_cobrosTab==='extras'?'active':''}"
        onclick="showCobrosTab('extras',this)">
        Extras procesadora
        <span class="cnt" style="background:rgba(251,191,36,.15);color:var(--yel)">${extras.length.toLocaleString('es-AR')}</span>
      </button>
      <button class="tb ${_cobrosTab==='arrastre'?'active':''}"
        onclick="showCobrosTab('arrastre',this)">
        Arrastre mes ant.
        <span class="cnt" style="background:rgba(139,92,246,.15);color:var(--vio)">${_cobArrastre.length.toLocaleString('es-AR')}</span>
      </button>
      <label class="btn-exp" style="margin-left:auto;cursor:pointer;font-size:10px;padding:4px 10px"
             title="Cargar backup JSON del período anterior para re-matchear pendientes">
        📂 Cargar período ant.
        <input type="file" accept=".json" style="display:none" onchange="_cobCargarArrastre(this)">
      </label>
    </div>

    <div id="cobros-tab-body"
         style="flex:1;overflow:hidden;display:flex;flex-direction:column;min-height:0"></div>
  </div>`;

  // Guardar datos para tabs y filtros
  window._cobSinLiq = sinLiq;
  window._cobExtras = extras;
  showCobrosTab(_cobrosTab);
}

function showCobrosTab(tab, btn) {
  _cobrosTab = tab;
  if (btn) {
    document.querySelectorAll('#tstrip-cobros .tb').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
  }
  const body = document.getElementById('cobros-tab-body');
  if (!body) return;

  const sinLiq = window._cobSinLiq || [];
  const extras = window._cobExtras || [];

  if (tab === 'sinliq') {
    const rows = _cobProc ? sinLiq.filter(x => x.proc === _cobProc) : sinLiq;
    _renderCobSinLiq(body, rows);
  } else if (tab === 'arrastre') {
    _renderCobArrastre(body);
  } else {
    const rows = _cobProc ? extras.filter(x => x.proc === _cobProc) : extras;
    _renderCobExtras(body, rows);
  }
}

function _cobProcBadge(proc) {
  if (proc === 'FISERV') return `<span class="st st-fis">FISERV</span>`;
  if (proc === 'GETPOS') return `<span class="st st-gp">GETPOS</span>`;
  return `<span class="st" style="background:rgba(52,211,153,.1);color:var(--grn);border-color:rgba(52,211,153,.3)">GoC</span>`;
}

function _cobToolbar(tab, rows, monto) {
  return `
  <div class="cobros-toolbar">
    <button class="btn-exp" onclick="exportarCobros('${tab}')">↓ Exportar Excel</button>
    <select class="filter-sel" onchange="_cobFiltroProc(this.value)">
      <option value="">Todas las procesadoras</option>
      <option value="FISERV" ${_cobProc==='FISERV'?'selected':''}>FISERV</option>
      <option value="GETPOS" ${_cobProc==='GETPOS'?'selected':''}>GETPOS</option>
      <option value="GoC"    ${_cobProc==='GoC'   ?'selected':''}>Go Cuotas</option>
    </select>
    <span style="font-size:10px;color:var(--m2)">${rows.length.toLocaleString('es-AR')} registros · ${fmtARS(monto)}</span>
  </div>`;
}

function _renderCobSinLiq(body, rows) {
  const monto = rows.reduce((s, x) => s + Math.abs(x.fila?.sky?.monto || 0), 0);
  if (!rows.length) {
    body.innerHTML = _cobToolbar('sinliq', rows, 0) +
      `<div style="padding:40px;text-align:center;color:var(--m2)">No hay operaciones sin liquidar.</div>`;
    return;
  }
  const HDR = ['Proc.','Fecha','Suc.','Vendedor','Tarjeta','Plan','Cuotas','Monto SKY','Lote','Cupón','Estado cruce','Acreditación'];
  const tplRow = x => {
    const s = x.fila?.sky || {};
    return `<tr>
      <td>${_cobProcBadge(x.proc)}</td>
      <td>${s.fecha||'—'}</td>
      <td>${s.suc||'—'}</td>
      <td class="td-trunc" title="${s.vendedor||''}">${s.vendedor||'—'}</td>
      <td>${s.tarjeta||'—'}</td>
      <td class="td-trunc" style="font-size:9px" title="${s.plan||''}">${s.plan||'—'}</td>
      <td class="num">${s.cuotas||1}</td>
      <td class="num" style="color:var(--red);font-weight:700">${fmtARS(s.monto)}</td>
      <td class="num">${x.lote||s.lote||'—'}</td>
      <td class="num">${x.cupon||s.cupon||'—'}</td>
      <td>${estadoBadge(x.fila?.estado)}</td>
      ${_cobPlazoCell(x)}
    </tr>`;
  };
  body.innerHTML = _cobToolbar('sinliq', rows, monto) + `
  <div class="tbl-wrap">
    <table class="res-tbl">
      <thead><tr>${HDR.map(h=>`<th>${h}</th>`).join('')}</tr></thead>
      <tbody>${rows.map(tplRow).join('')}</tbody>
    </table>
  </div>`;
}

function _renderCobExtras(body, rows) {
  const monto = rows.reduce((s, x) => s + Math.abs(x.liqRow?.monto || 0), 0);
  if (!rows.length) {
    body.innerHTML = _cobToolbar('extras', rows, 0) +
      `<div style="padding:40px;text-align:center;color:var(--m2)">No hay extras de procesadora.</div>`;
    return;
  }
  const HDR = ['Proc.','Fecha Venta','Equipo','Lote','Cupón','Tarjeta','Cuotas','Monto','Cód. Auth.','Nro Comercio'];
  const tplRow = x => {
    const r = x.liqRow;
    return `<tr>
      <td>${_cobProcBadge(x.proc)}</td>
      <td>${r.fecha_venta||r.fecha||'—'}</td>
      <td class="num">${r.equipo||'—'}</td>
      <td class="num">${r.lote||'—'}</td>
      <td class="num">${r.cupon||'—'}</td>
      <td>${r.tarjeta||'—'}</td>
      <td class="num">${r.cuotas||1}</td>
      <td class="num" style="color:var(--yel);font-weight:700">${fmtARS(r.monto)}</td>
      <td class="num" style="font-size:9px">${r.aut||'—'}</td>
      <td class="num" style="font-size:9px">${r.nro_comercio||'—'}</td>
    </tr>`;
  };
  body.innerHTML = _cobToolbar('extras', rows, monto) + `
  <div class="tbl-wrap">
    <table class="res-tbl">
      <thead><tr>${HDR.map(h=>`<th>${h}</th>`).join('')}</tr></thead>
      <tbody>${rows.map(tplRow).join('')}</tbody>
    </table>
  </div>`;
}

function _renderCobArrastre(body) {
  if (!_cobArrastre.length) {
    body.innerHTML = `
    <div style="padding:40px;display:flex;flex-direction:column;align-items:center;gap:14px">
      <div style="color:var(--m2);font-size:13px">No hay pendientes del período anterior cargados.</div>
      <label class="btn-exp" style="cursor:pointer">
        📂 Cargar backup período anterior
        <input type="file" accept=".json" style="display:none" onchange="_cobCargarArrastre(this)">
      </label>
      <div style="font-size:11px;color:var(--m2);max-width:400px;text-align:center">
        Seleccioná el archivo JSON de cierre del período anterior para ver qué operaciones
        ya acreditaron este mes y cuáles siguen pendientes.
      </div>
    </div>`;
    return;
  }

  const liqCups = (typeof _LIQ_CUPONES !== 'undefined') ? _LIQ_CUPONES : [];
  const matched = _cobRematcharArrastre();
  const display = _cobProc ? matched.filter(x => x.op.proc === _cobProc) : matched;

  const acreditados = matched.filter(x =>  x.acreditoEsteMes);
  const pendientes  = matched.filter(x => !x.acreditoEsteMes);
  const montoAcred  = acreditados.reduce((s, x) => s + Math.abs(parseFloat(x.op.monto) || 0), 0);
  const montoPend   = pendientes.reduce( (s, x) => s + Math.abs(parseFloat(x.op.monto) || 0), 0);

  const HDR = ['Proc.','Estado ant.','Fecha','Suc.','Tarjeta','Cuotas','Monto','Lote','Cupón','Acreditó este mes','Plazo acred.'];

  const tplRow = x => {
    const op  = x.op;
    const res = _cobFechaAcreditacion(op.fecha, op.proc, op.tarjeta);
    let plazoCell = '<td style="color:var(--m2);font-size:9px">—</td>';
    if (res) {
      const hoy  = new Date().toISOString().slice(0, 10);
      const dias = Math.round(Math.abs(new Date(hoy) - new Date(res.fechaAcred)) / 86400000);
      plazoCell  = res.fechaAcred < hoy
        ? `<td style="color:var(--red);font-size:9px">${res.fechaAcred}<br><span style="font-size:8px">+${dias}d vencido</span></td>`
        : `<td style="color:var(--grn);font-size:9px">${res.fechaAcred}<br><span style="font-size:8px">en ${dias}d</span></td>`;
    }
    const estadoCell = x.acreditoEsteMes
      ? `<td><span style="background:rgba(52,211,153,.15);color:var(--grn);border:1px solid rgba(52,211,153,.3);border-radius:4px;padding:1px 6px;font-size:9px">✓ Acreditó</span></td>`
      : `<td><span style="background:rgba(248,113,113,.15);color:var(--red);border:1px solid rgba(248,113,113,.3);border-radius:4px;padding:1px 6px;font-size:9px">Pendiente</span></td>`;
    return `<tr>
      <td>${_cobProcBadge(op.proc)}</td>
      <td style="font-size:9px;color:var(--m2)">${op.estado||'—'}</td>
      <td>${op.fecha||'—'}</td>
      <td>${op.suc||'—'}</td>
      <td>${op.tarjeta||'—'}</td>
      <td class="num">${op.cuotas||1}</td>
      <td class="num" style="font-weight:700">${fmtARS(op.monto)}</td>
      <td class="num">${op.lote||'—'}</td>
      <td class="num">${op.cupon||'—'}</td>
      ${estadoCell}
      ${plazoCell}
    </tr>`;
  };

  body.innerHTML = `
  <div class="cobros-toolbar">
    <button class="btn-exp" onclick="exportarCobros('arrastre')">↓ Exportar Excel</button>
    <select class="filter-sel" onchange="_cobFiltroProc(this.value)">
      <option value="">Todas las procesadoras</option>
      <option value="FISERV" ${_cobProc==='FISERV'?'selected':''}>FISERV</option>
      <option value="GETPOS" ${_cobProc==='GETPOS'?'selected':''}>GETPOS</option>
      <option value="GoC"    ${_cobProc==='GoC'   ?'selected':''}>Go Cuotas</option>
    </select>
    <span style="font-size:10px;color:var(--m2)">${display.length} de ${matched.length} registros</span>
    <span style="font-size:10px;color:var(--grn)">✓ ${acreditados.length} acreditaron · ${fmtARS(montoAcred)}</span>
    <span style="font-size:10px;color:var(--red)">✗ ${pendientes.length} pendientes · ${fmtARS(montoPend)}</span>
    ${!liqCups.length ? '<span style="font-size:10px;color:var(--yel)">⚠ Sin archivo liq cargado</span>' : ''}
    <label class="btn-exp" style="cursor:pointer;font-size:10px;padding:3px 8px;margin-left:auto">
      📂 Cambiar período
      <input type="file" accept=".json" style="display:none" onchange="_cobCargarArrastre(this)">
    </label>
  </div>
  <div class="tbl-wrap">
    <table class="res-tbl">
      <thead><tr>${HDR.map(h=>`<th>${h}</th>`).join('')}</tr></thead>
      <tbody>${display.map(tplRow).join('')}</tbody>
    </table>
  </div>`;
}

function _cobFiltroProc(v) { _cobProc = v; showCobrosTab(_cobrosTab); }

// ══════════════════════════════════════════════════════════════════
// EXPORTAR EXCEL
// ══════════════════════════════════════════════════════════════════
function exportarCobros(tab) {
  const sinLiq = window._cobSinLiq || [];
  const extras = window._cobExtras || [];
  let HDR, dataFn, sheetName, rows;

  if (tab === 'sinliq') {
    rows = _cobProc ? sinLiq.filter(x => x.proc === _cobProc) : sinLiq;
    HDR  = ['Procesadora','Fecha','Suc.','Vendedor','Tarjeta','Plan','Cuotas','Monto SKY',
            'Lote','Cupón','Estado cruce','Fecha acred. esperada'];
    dataFn = x => {
      const s   = x.fila?.sky || {};
      const res = _cobFechaAcreditacion(s.fecha, x.proc, s.tarjeta);
      return [x.proc, s.fecha||'', s.suc||'', s.vendedor||'', s.tarjeta||'', s.plan||'',
              s.cuotas||1, s.monto||0, x.lote||s.lote||'', x.cupon||s.cupon||'',
              x.fila?.estado||'', res?.fechaAcred||''];
    };
    sheetName = 'Sin Liquidar';
  } else if (tab === 'arrastre') {
    const matched = _cobRematcharArrastre();
    rows = _cobProc ? matched.filter(x => x.op.proc === _cobProc) : matched;
    HDR  = ['Procesadora','Estado ant.','Fecha','Suc.','Tarjeta','Cuotas','Monto',
            'Lote','Cupón','Acreditó este mes','Fecha acred. esperada'];
    dataFn = x => {
      const op  = x.op;
      const res = _cobFechaAcreditacion(op.fecha, op.proc, op.tarjeta);
      return [op.proc, op.estado||'', op.fecha||'', op.suc||'', op.tarjeta||'',
              op.cuotas||1, op.monto||0, op.lote||'', op.cupon||'',
              x.acreditoEsteMes ? 'Sí' : 'No', res?.fechaAcred||''];
    };
    sheetName = 'Arrastre';
  } else {
    rows = _cobProc ? extras.filter(x => x.proc === _cobProc) : extras;
    HDR  = ['Procesadora','Fecha Venta','Equipo','Lote','Cupón','Tarjeta','Cuotas','Monto','Cód. Auth.','Nro Comercio'];
    dataFn = x => {
      const r = x.liqRow;
      return [x.proc, r.fecha_venta||r.fecha||'', r.equipo||'', r.lote||'', r.cupon||'',
              r.tarjeta||'', r.cuotas||1, r.monto||0, r.aut||'', r.nro_comercio||''];
    };
    sheetName = 'Extras';
  }

  if (!rows.length) { alert('No hay datos para exportar.'); return; }
  const ws  = XLSX.utils.aoa_to_sheet([HDR, ...rows.map(dataFn)]);
  const wb2 = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb2, ws, sheetName);
  const ts = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb2, `Cobros_${tab}_${ts}.xlsx`);
}
