// ═══════════════════════════════════════════════════════════════════
// CONCILIADOR.JS — Motor completo de conciliación
// ═══════════════════════════════════════════════════════════════════

// Estado de filtros
const FILTROS = { suc:'', tar:'', proc:'', fecha:'', vend:'', search:'' };
const HDR_BASE = ['Estado','Método','Proc.Esp.','Proc.Real','Integrado',
  'Suc. SKY','Tarjeta SKY','Plan SKY','Fecha SKY','Monto SKY','Cupón SKY','Lote SKY',
  'Tarjeta Proc.','Cuotas Proc.','Com.FIS','Cód.Auth. Proc.','Lote Proc.','Ticket Proc.','Suc. Proc.',
  'Dif. Monto','Com.OK','Match Parcial'];

// ── Normalización ────────────────────────────────────────
function norm(v) { const s=String(v??'').trim().replace(/^0+/,''); return s||'0'; }

// Parsea un monto desde string con cualquier formato a float
function parseMontoFloat(v) {
  if (!v && v !== 0) return 0;
  if (typeof v === 'number') return Math.abs(v);
  let s = String(v).trim()
    .replace(/^[\s$€£¥\u20AC]+/, '')
    .replace(/[\s$€£¥\u20AC]/g, '')
    .trim();
  // Formato AR: "1.234,56" → punto miles, coma decimal
  if (/^-?[\d.]+,\d{1,2}$/.test(s)) {
    s = s.replace(/\./g,'').replace(',','.');
  } else if (/^-?[\d,]+\.\d{1,2}$/.test(s)) {
    s = s.replace(/,/g,'');
  } else {
    if (/^-?\d{1,3}(\.\d{3})+$/.test(s)) s = s.replace(/\./g,'');
    else s = s.replace(/,/g,'');
  }
  const n = parseFloat(s);
  return isNaN(n) ? 0 : Math.abs(n);
}

function normMonto(v) {
  if (!v && v !== 0) return '0';
  // Quitar símbolo de moneda, espacios y caracteres no numéricos iniciales
  // Ej: "$ 9.999,00" → "9.999,00" | "$ 1.715.998,68" → "1.715.998,68"
  let s = String(v).trim()
    .replace(/^[\s$\u20AC\u00A3\uFFE5\u00A5]+/, '')  // quitar $ y otros símbolos al inicio
    .replace(/[\s$\u20AC]+/g, '')                       // quitar $ internos
    .trim();
  // Detectar separador decimal:
  // Formato AR: "1.234.567,89" → coma es decimal, punto es miles
  // Formato US: "1,234,567.89" → punto es decimal, coma es miles
  // Si termina en ",NN" (1-2 dígitos) con puntos antes → AR
  if (/^-?[\d.]+,\d{1,2}$/.test(s)) {
    s = s.replace(/\./g,'').replace(',','.');
  } else if (/^-?[\d,]+\.\d{1,2}$/.test(s)) {
    // Formato US con punto decimal
    s = s.replace(/,/g,'');
  } else {
    // Sin separador decimal claro → quitar comas y puntos de miles
    // Si hay exactamente un punto con 2 decimales al final, es decimal
    if (/^-?\d{1,3}(\.\d{3})+$/.test(s)) {
      // "1.234.567" → miles con punto → quitar puntos
      s = s.replace(/\./g,'');
    } else {
      s = s.replace(/,/g,'');
    }
  }
  const n = parseFloat(s);
  if (isNaN(n)) return '0';
  return Math.round(Math.abs(n)).toString();
}
function normFecha(v) {
  if (!v || v === 'null' || v === 'undefined') return '';
  // Date object (SheetJS con cellDates:true) → más confiable
  if (v instanceof Date && !isNaN(v)) return v.toISOString().slice(0,10);
  const s = String(v).trim();
  if (!s || s === 'Invalid Date') return '';
  // ISO: 2026-05-02 o 2026-05-02T00:00:00...
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0,10);
  // DD/MM/YYYY HH:MM (FISERV, GETPOS) — 1 o 2 dígitos
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return `${m[3]}-${String(m[2]).padStart(2,'0')}-${String(m[1]).padStart(2,'0')}`;
  // MM/DD/YYYY (SheetJS locale US)
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    // Si el primer número > 12 es día → DD/MM
    const a=parseInt(m[1]), b=parseInt(m[2]);
    if (a > 12) return `${m[3]}-${String(b).padStart(2,'0')}-${String(a).padStart(2,'0')}`;
    return `${m[3]}-${String(a).padStart(2,'0')}-${String(b).padStart(2,'0')}`;
  }
  // Excel serial number como string "46143"
  if (/^\d{5}$/.test(s)) {
    const d = new Date(Math.round((parseInt(s)-25569)*86400000));
    return d.toISOString().slice(0,10);
  }
  return s.slice(0,10);
}
function limpiarNombre(n) {
  // Igual que simp() en Python: minúsculas, quitar prefijo (NNN), solo alfanumérico
  let s = String(n||'').toLowerCase().trim();
  s = s.replace(/^\(\d+\)\s*/,'');           // quitar prefijo (NNN)
  s = s.replace(/[^a-z0-9 ]/g,' ');            // solo letras/números/espacio
  s = s.replace(/\s+/g,' ').trim();
  return s;
}
function fmtARS(v) {
  const n=parseFloat(v)||0;
  const parts=Math.abs(n).toFixed(2).split('.');
  parts[0]=parts[0].replace(/\B(?=(\d{3})+(?!\d))/g,'.');
  return `$ ${parts[0]},${parts[1]}`;
}
function fmtFecha(v) { return v ? String(v).slice(0,10) : ''; }
const delay = ms => new Promise(r=>setTimeout(r,ms));

// ── PARSEO ARCHIVOS ──────────────────────────────────────
function parseSkylab(wb) {
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, {defval:null, raw:false});
  return rows.map((r,i) => {
    const plan = String(r['Plan']??'').toUpperCase();
    let cuotas = 1;
    const mc = plan.match(/(\d+)\s*CUOTA/); if (mc) cuotas=parseInt(mc[1]);
    else { const mx = plan.match(/(\d+)X/); if (mx) cuotas=parseInt(mx[1]); }
    if (plan.includes('DEBITO')) cuotas=1;
    // Con raw:false SheetJS puede dar "True", "1", "1.0" o un número
    const _pi = r['P.Integrado'];
    const integrado = _pi != null &&
                      String(_pi).trim() !== '' &&
                      String(_pi).trim() !== '0' &&
                      String(_pi).trim().toLowerCase() !== 'false' &&
                      String(_pi).trim().toLowerCase() !== 'nan' &&
                      String(_pi).trim().toLowerCase() !== 'none';
    return {
      idx:      i,
      asiento:  r['Nro.Asiento'],
      suc:      String(r['Sucursal']??'').trim().replace(/^0+/,'') || '0',
      vendedor: r['Vendedor'],
      tarjeta:  String(r['Tarjeta']??'').trim(),
      tarjetaU: String(r['Tarjeta']??'').trim().toUpperCase(),
      plan:     r['Plan'],
      cuotas,
      fecha:    normFecha(r['Fec.de Vta.']),
      fecPago:  normFecha(r['Fec. de Pago'] || r['Fec.de Pago'] || r['Fecha de Pago'] || r['Fec de Pago'] || ''),
      lote:     norm(r['Lote']),
      cupon:    norm(r['Cupon']),
      nroCom:   String(r['Nro.Comercio']??'').replace('.0','').trim().replace(/^0+/,'') || '0',
      monto:    (() => {
        const raw = String(r['Venta Bruta']??'').trim();
        const neg = raw.startsWith('-');
        const n = parseMontoFloat(raw);
        return neg ? -n : n;
      })(),
      montoN:   normMonto(r['Venta Bruta']),
      neto:     parseFloat(r['Neto a Cobrar'])||0,
      esGETPos:    String(r['Tarjeta']??'').trim().toUpperCase()==='GETPOS',
      esGOCUOTAS:  String(r['Tarjeta']??'').trim().toUpperCase()==='GO CUOTAS',
      esNeg:    (parseFloat(String(r['Venta Bruta']||'').replace(/,/g,''))||0) < 0,
      integrado,
    };
  });
}
// DEBUG: exportar para diagnóstico (se llama desde conciliar())
window._debugSkyRows = null;
window._debugFisRows = null;
window._debugGpRows  = null;
// Índices globales para re-cruce manual
let _FIS_NORM = [], _FIS_REV = [], _GP_NORM = [], _GP_REV = [];
// Marcas de COM.ERRADO: {skyIdx: 'SIN_DIF' | 'CON_DIF'}
let COM_ERRADO_MARCAS = {};

function parseTerminales(wb) {
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, {defval:null, raw:false});
  const term2suc={}, nombre2suc={};
  for (const r of rows) {
    const nombre=String(r['Nombre']||'');
    const term=String(r['Terminal']||'').trim().replace(/^0+/,'');
    const m=nombre.match(/\((\d{3,})\)/);
    if (m) {
      const suc=m[1];
      term2suc[term]=suc;
      const clean=limpiarNombre(nombre);
      if (clean) nombre2suc[clean]=suc;
    }
  }
  return { term2suc, nombre2suc };
}


// Busca una columna en un row por nombre, insensible a tildes/case
function getCol(r, ...names) {
  for (const name of names) {
    if (r[name] !== undefined && r[name] !== null) return r[name];
  }
  // Búsqueda fuzzy: normalizar y comparar
  const rkeys = Object.keys(r);
  for (const name of names) {
    const norm_name = name.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9 ]/g,' ').trim();
    const found = rkeys.find(k => {
      const norm_k = k.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9 ]/g,' ').trim();
      return norm_k === norm_name;
    });
    if (found !== undefined && r[found] !== undefined && r[found] !== null) return r[found];
  }
  return null;
}

function parseFiserv(wb, term2suc) {
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, {defval:null, raw:false});
  const esNuevo = rows.length>0 && 'Ticket' in rows[0];
  if (rows.length > 0) window._debugFisRawRow = rows[0]; // DEBUG
  // Log keys del primer row para diagnóstico
  if (rows.length > 0) {
    const keys = Object.keys(rows[0]);
    const montoKey = keys.find(k => k.toLowerCase().replace(/[^a-z ]/g,'').trim() === 'monto total');
    if (!montoKey) console.warn('[FISERV] No se encontró columna Monto total. Keys:', keys.slice(0,20));
    else console.log('[FISERV] Columna monto encontrada:', montoKey, '=', rows[0][montoKey]);
  }
  const norm_r=[], rev_r=[];
  for (const r of rows) {
    let obj;
    if (esNuevo) {
      const tipo=String(r['Tipo de transacción']||'');
      if (!['Compra','Devolución','Anulación','Reverso'].includes(tipo)) continue;
      if (String(r['Estado']||'')!=='Completa') continue;
      const equipo=String(r['Terminal']||'').trim().replace(/^0+/,'');
      // Buscar monto por múltiples nombres posibles de columna
      const montoRaw = getCol(r,'Monto total','Monto Total','monto total');
      const comRaw   = getCol(r,'Cód. comercio','Cod. comercio','Cód comercio','cod comercio');
      const autRaw   = getCol(r,'Autorización','Autorizacion','autorización','autorizacion');
      const fechaRaw = getCol(r,'Fecha','fecha');
      const loteRaw  = getCol(r,'Lote','lote');
      const ticketRaw= getCol(r,'Ticket','ticket');
      obj={lote:norm(loteRaw),ticket:norm(ticketRaw),aut:norm(autRaw),
        monto:parseMontoFloat(montoRaw),montoN:normMonto(montoRaw),
        fecha:normFecha(fechaRaw),equipo,suc:term2suc[equipo]||'',
        comFis:String(comRaw||'').replace('.0','').trim().replace(/^0+/,'') || '',
        tarjeta:String(getCol(r,'Producto','producto')||''),
        cuotas:parseInt(getCol(r,'Cuotas','cuotas'))||1,tipo,
        arancel:null,cfo:null};
    } else {
      const tipo=String(r['Tipo operacion']||'').toUpperCase();
      const equipo=String(r['Nro Equipo']||'').trim();
      obj={lote:norm(r['Nro de Lote']),ticket:norm(r['Nro de Cupón']),aut:norm(r['Código Autorización']),
        monto:parseMontoFloat(r['Importe Venta']),montoN:normMonto(r['Importe Venta']),
        fecha:normFecha(r['Fecha Venta']),equipo,suc:term2suc[equipo]||'',
        comFis:String(r['Nro Comercio']||'').replace('.0','').trim().replace(/^0+/,'') || '',
        tarjeta:String(r['Tarjeta']||''),cuotas:parseInt(r['Cuotas'])||1,tipo,
        arancel:parseFloat(r['Arancel'])||0,cfo:parseFloat(r['CFO'])||0};
    }
    const esRev=['Devolución','Anulación','Reverso','DEVOLUCION','ANULACION','REVERSO']
      .includes(obj.tipo?.toUpperCase?.()?.trim?.()??'');
    (esRev?rev_r:norm_r).push(obj);
  }
  return { norm:norm_r, rev:rev_r };
}

function parseGetpos(wb, nombre2suc) {
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, {defval:null, raw:false});
  const norm_r=[], rev_r=[];
  for (const r of rows) {
    const tipo=String(getCol(r,'Tipo de Transacción','Tipo de Transaccion','tipo de transaccion')||'');
    const nombre=limpiarNombre(r['Nombre Establecimiento']||'');
    // bm(): exact match primero, luego substring (igual que Python)
    let gpSuc = nombre2suc[nombre] || '';
    if (!gpSuc) {
      for (const [nm, sv] of Object.entries(nombre2suc)) {
        if (nombre.length>=3 && nm.length>=3 && (nombre.includes(nm)||nm.includes(nombre))) {
          gpSuc=sv; break;
        }
      }
    }
    const esRev=['Anulación','Devolución'].includes(tipo);
    const montoRawGp=parseMontoFloat(getCol(r,'Monto Bruto Transacción','Monto Bruto Transaccion','monto bruto transaccion'));
    const obj={aut:norm(getCol(r,'Cód. Aut.','Cod. Aut.','cód aut','cod aut')),
      cupon:norm(getCol(r,'Nro de Cupón','Nro de Cupon','nro de cupon')),
      monto: esRev ? -montoRawGp : montoRawGp,
      montoN:normMonto(montoRawGp),
      fecha:normFecha(getCol(r,'Fecha de Operación','Fecha de Operacion','fecha de operacion')),
      nombre:String(r['Nombre Establecimiento']||''),
      suc:gpSuc,pos:String(r['Código del POS']||''),
      marca:String(r['Marca']||''),plan:String(r['Plan Cuotas']||''),
      cuotas:parseInt(String(r['Plan Cuotas']||'').match(/\d+/)?.[0])||1,tipo,
      tarjeta:r['Tarjeta'],cupon_raw:r['Nro de Cupón']};
    (esRev?rev_r:norm_r).push(obj);
  }
  return { norm:norm_r, rev:rev_r };
}

// ── ÍNDICES ──────────────────────────────────────────────

// Busca en un índice tolerando diferencias de ±1 en montoN (centavos de redondeo)
// La clave tiene formato "METODO|campo1|...|montoN" donde montoN es el ÚLTIMO campo
function idxGetTol(idx, key) {
  const hit = idx[key];
  if (hit && hit.length) return hit;
  // Intentar con monto ± 1 (diferencias de centavos por redondeo)
  const parts = key.split('|');
  const last  = parseInt(parts[parts.length - 1]);
  if (isNaN(last)) return null;
  for (const delta of [-1, 1]) {
    const altKey = [...parts.slice(0,-1), String(last+delta)].join('|');
    const altHit = idx[altKey];
    if (altHit && altHit.length) return altHit;
  }
  return null;
}

function buildIndexFis(rows) {
  const idx={};
  for (const r of rows) {
    const mn = r.montoN;
    const mnP = String(parseInt(mn)+1);
    const mnM = String(Math.max(0,parseInt(mn)-1));
    for (const m of [mn, mnP, mnM]) {
      const keys=[
        `M1|${r.lote}|${r.ticket}|${r.suc}|${r.fecha}|${m}`,
        `M2|${r.lote}|${r.ticket}||${r.fecha}|${m}`,
        `M3|${r.lote}|${r.ticket}|${r.suc}||${m}`,
        `M11|${r.aut}|${r.suc}|${r.fecha}|${m}`,
        `M13|${r.lote}|${r.ticket}1|${r.suc}|${r.fecha}|${m}`,
        `MN1|${r.suc}|${r.fecha}|${m}`,
      ];
      for (const k of keys) (idx[k]=idx[k]||[]).push(r);
    }
  }
  return idx;
}

function buildIndexGp(rows) {
  const idx={};
  for (const r of rows) {
    // Indexar con montoN Y con montoN±1 para tolerar diferencias de redondeo de centavos
    const mn = r.montoN;
    const mnP = String(parseInt(mn)+1); // ceil alternativo
    const mnM = String(Math.max(0,parseInt(mn)-1)); // floor alternativo
    for (const m of [mn, mnP, mnM]) {
      const keys=[
        `M4|${r.aut}|${m}|${r.fecha}|${r.suc}`,
        `M5|${r.aut}|${m}|${r.fecha}|`,
        `M6|${r.aut}|${m}||${r.suc}`,
        `M7|${r.cupon}|${m}|${r.fecha}|${r.suc}`,
        `M8|${r.cupon}|${m}||${r.suc}`,
        `MN2|${m}|${r.fecha}|${r.suc}`,
        `BKEY|${m}|${r.fecha}|${r.suc}`,
      ];
      for (const k of keys) (idx[k]=idx[k]||[]).push(r);
    }
  }
  return idx;
}

// ── VALIDAR COMERCIO ─────────────────────────────────────
function normCom(v) {
  // Normaliza nro de comercio: quita ceros iniciales igual que norm()
  const s = String(v||'').trim().replace(/^0+/,'');
  return s || '0';
}
function validarComercio(skyRow, procRow) {
  const comSky = normCom(skyRow.nroCom);
  const comFis = normCom(procRow.comFis);

  // Primero: si son iguales directamente → OK sin importar equivalencias
  if (comFis === comSky) return 'OK';

  // Buscar equiv: (nroCom_sky, suc) → nroCom_fis_alternativo
  const equiv = TM.equivalencias[`${comSky}|${skyRow.suc}`]
             || TM.equivalencias[`${skyRow.nroCom}|${skyRow.suc}`];
  if (equiv) {
    const equivN = normCom(equiv);
    // FISERV puede usar el comercio equivalente O el mismo que SKY → ambos son OK
    if (comFis === equivN || comFis === comSky) return 'OK (equiv.)';
    return 'DIFERENTE';
  }

  return 'DIFERENTE';
}

// ── MOTOR PRINCIPAL ──────────────────────────────────────
function conciliarRows(skyRows, fisNorm, fisRev, gpNorm, gpRev) {
  const idxFN=buildIndexFis(fisNorm), idxFR=buildIndexFis(fisRev);
  const idxGN=buildIndexGp(gpNorm),   idxGR=buildIndexGp(gpRev);
  const used=new Set();

  // Índice Go Cuotas — dos claves de búsqueda:
  //   idxGoC[orden]  → para la mayoría de sucursales (Nro de Orden = Cupon SKY)
  //   idxGoCRef[ref] → para suc. 251 y 097 (Referencia Externa = Cupon SKY)
  const idxGoC    = {};
  const idxGoCRef = {};
  if (typeof _GOC_PAGOS !== 'undefined' && _GOC_PAGOS.length) {
    _GOC_PAGOS.forEach(p => {
      if (p.orden) idxGoC[p.orden] = p;
      const ref = String(p.refExt || '').trim();
      if (ref && ref !== '-' && ref !== '0') idxGoCRef[ref] = p;
    });
  }
  const gocEnabled = typeof PROCS_ENABLED !== 'undefined'
    ? (PROCS_ENABLED.GOCUOTAS !== false)
    : true;

  return skyRows.map(s => {
    const idxFis = s.esNeg ? idxFR : idxFN;
    const idxGp  = s.esNeg ? idxGR : idxGN;
    const procEsp = s.esGETPos    ? 'GETPOS'    :
                    s.esGOCUOTAS  ? 'GOCUOTAS'  : 'FISERV';
    const {lote:ln,cupon:cn,suc:sn,fecha:fn,montoN:mn} = s;

    // ── Devoluciones: MN1 y MN2
    if (s.esNeg) {
      const mf=(idxGetTol(idxFis,`MN1|${sn}|${fn}|${mn}`)||[]).find(r=>!used.has(r));
      if (mf) { used.add(mf); return armarFila(s, mf,'FISERV','MN1',procEsp); }
      const mg=(idxGetTol(idxGp,`MN2|${mn}|${fn}|${sn}`)||[]).find(r=>!used.has(r));
      if (mg) { used.add(mg); return armarFila(s, mg,'GETPOS','MN2',procEsp); }
      return { sky:s, proc:null, metodo:'SIN MATCH', estado:'SIN MATCH',
        procEncontrada:'', procEsperada:procEsp, comOK:'', sucOK:'',
        matchParcial:'', esDevolucion:'SI', esAnulSinCobro:'NO' };
    }

    // ── Métodos en cascada
    const intentosFis=[
      [`M1|${ln}|${cn}|${sn}|${fn}|${mn}`,   idxFis,'FISERV','M1'],
      [`M2|${ln}|${cn}||${fn}|${mn}`,         idxFis,'FISERV','M2'],
      [`M3|${ln}|${cn}|${sn}||${mn}`,         idxFis,'FISERV','M3'],
      [`M11|${cn}|${sn}|${fn}|${mn}`,         idxFis,'FISERV','M11'],
      [`M13|${ln}|${cn}|${sn}|${fn}|${mn}`,   idxFis,'FISERV','M13'],
    ];
    const intentosGp=[
      [`M4|${cn}|${mn}|${fn}|${sn}`, idxGp,'GETPOS','M4'],
      [`M5|${cn}|${mn}|${fn}|`,      idxGp,'GETPOS','M5'],
      [`M6|${cn}|${mn}||${sn}`,      idxGp,'GETPOS','M6'],
      [`M7|${cn}|${mn}|${fn}|${sn}`, idxGp,'GETPOS','M7'],
      [`M8|${cn}|${mn}||${sn}`,      idxGp,'GETPOS','M8'],
    ];
    const intentos = s.esGETPos ? intentosGp : intentosFis;

    for (const [key, idx, procReal, met] of intentos) {
      const match=(idxGetTol(idx,key)||[]).find(r=>!used.has(r));
      if (match) { used.add(match); return armarFila(s, match, procReal, met, procEsp); }
    }

    // M9/M10 — match parcial GETPOS
    if (s.esGETPos) {
      const cands=(idxGp[`BKEY|${mn}|${fn}|${sn}`]||[]).filter(r=>!used.has(r));
      for (const r of cands) {
        if (cn.length>=3 && r.aut.length>=3 && (r.aut.includes(cn)||cn.includes(r.aut))) {
          used.add(r);
          const f=armarFila(s, r,'GETPOS','M9',procEsp);
          f.matchParcial=`SKY=${cn} ↔ GP.Aut=${r.aut}`;
          return f;
        }
        if (cn.length>=3 && r.cupon.length>=3 && (r.cupon.includes(cn)||cn.includes(r.cupon))) {
          used.add(r);
          const f=armarFila(s, r,'GETPOS','M10',procEsp);
          f.matchParcial=`SKY=${cn} ↔ GP.Cup=${r.cupon}`;
          return f;
        }
      }
    }

    // M12 — match parcial FISERV
    if (!s.esGETPos) {
      const cands=fisNorm.filter(r=>!used.has(r)&&r.montoN===mn&&r.fecha===fn&&r.suc===sn);
      for (const r of cands) {
        if (cn.length>=3&&r.aut.length>=3&&(r.aut.includes(cn)||cn.includes(r.aut))) {
          used.add(r);
          const f=armarFila(s, r,'FISERV','M12',procEsp);
          f.matchParcial=`SKY=${cn} ↔ FIS.Aut=${r.aut}`;
          return f;
        }
      }
    }

    // ── Go Cuotas: match por Número de Orden O por Referencia Externa
    // Suc 251 y 097 usan Referencia Externa como cupón en Skylab.
    // Siempre validamos que el monto no difiera más del 50% para evitar
    // matches incorrectos (mismo nro. de orden en distintos períodos, etc.)
    if (s.esGOCUOTAS && gocEnabled && (Object.keys(idxGoC).length || Object.keys(idxGoCRef).length)) {
      const cup      = norm(s.cupon);
      const skyMonto = Math.abs(s.monto);

      const _validarMonto = (hit) => {
        if (!hit) return false;
        const gocMonto = Math.abs(hit.importe);
        if (skyMonto === 0 || gocMonto === 0) return true;
        const ratio = Math.max(skyMonto, gocMonto) / Math.min(skyMonto, gocMonto);
        return ratio <= 1.5; // tolerancia del 50%
      };

      // Primero buscar por Número de Orden (mayoría de sucursales)
      let hit = idxGoC[cup];
      let met = 'GoC:Orden';

      if (hit && !_validarMonto(hit)) {
        // Match por orden encontrado pero monto muy diferente → buscar por RefExt
        hit = null;
      }

      // Si no encontró por orden (o monto no coincide), buscar por Referencia Externa
      if (!hit && idxGoCRef[cup]) {
        const hitRef = idxGoCRef[cup];
        if (_validarMonto(hitRef)) {
          hit = hitRef;
          met = 'GoC:RefExt';
        }
      }

      if (hit) return armarFilaGoC(s, hit, procEsp, met);
    }

    return { sky:s, proc:null, metodo:'SIN MATCH', estado:'SIN MATCH',
      procEncontrada:'', procEsperada:procEsp, comOK:'', sucOK:'',
      matchParcial:'', esDevolucion:'NO', esAnulSinCobro:'NO' };
  });
}

// ── Armar fila de resultado para Go Cuotas ──────────────────────────
function armarFilaGoC(s, pago, procEsp, met) {
  met = met || 'GoC:Orden';
  const procObj = {
    ticket:  pago.orden,
    aut:     pago.orden,
    cupon:   pago.orden,
    monto:   pago.importe,
    montoN:  normMonto(pago.importe),
    fecha:   pago.fechaOrigen,
    suc:     pago.sucNombre || '',
    tarjeta: 'Go Cuotas',
    cuotas:  pago.cuotas || 1,
    comFis:  '',
    nombre:  pago.nombre || '',
    marca:   'GOCUOTAS',
    plan:    pago.cuotas ? `${pago.cuotas} cuotas` : '',
    equipo:  '',
    pos:     '',
    tipo:    'Venta',
    arancel: 0, cfo: 0,
  };
  const difMonto = Math.abs(Math.abs(s.monto) - Math.abs(pago.importe));
  return {
    sky: s, proc: procObj,
    metodo:        met,
    estado:        'OK (GoC)',
    procEncontrada:'GOCUOTAS',
    procEsperada:  procEsp,
    comOK:         'OK',
    sucOK:         'OK',
    matchParcial:  '',
    esDevolucion:  'NO',
    esAnulSinCobro:'NO',
    difMonto:      difMonto > 0.5 ? +difMonto.toFixed(2) : 0,
    procMontoNorm: normMonto(pago.importe),
  };
}

function armarFila(s, proc, procReal, metodo, procEsp) {
  let estado, comOK='', sucOK='';
  if (procReal==='FISERV' && proc.comFis!==undefined) {
    comOK=validarComercio(s,proc);
    sucOK=proc.suc?(proc.suc===s.suc?'OK':'DIFERENTE'):'SIN DATO';
  }
  if (procEsp==='GETPOS'&&procReal==='FISERV') estado='MAL FACTURADO: debe ser GETPOS';
  else if (procEsp==='FISERV'&&procReal==='GETPOS') estado='MAL FACTURADO: debe ser FISERV';
  else if (comOK==='DIFERENTE') estado='COM. ERRADO';
  else if (comOK==='OK (equiv.)') estado='OK (equiv.)';
  else estado='OK';

  // Marcar integradas visualmente sin cambiar el estado
  if (s.integrado && estado==='OK') estado='OK (integrado)';

  // ── Diferencia de cuotas ────────────────────────────────────────
  // 1. Buscar cuotas en TM.planes (fuente autoritativa).
  //    Si el plan SKY no tiene cuotas en el texto (p.ej. "TARJETA NARANJA CLASICA")
  //    el valor extraído por parseSkylab es 1, pero TM.planes puede tener el dato correcto.
  const tmCuotas   = (typeof buscarCuotasEnTM === 'function')
    ? buscarCuotasEnTM(s.plan, s.tarjeta, procReal)
    : null;
  const skyCuotas  = tmCuotas !== null ? tmCuotas : Math.max(1, parseInt(s.cuotas) || 1);
  const skyCuotasTM = tmCuotas !== null; // true = el dato viene de TM.planes
  const procCuotas = Math.max(1, parseInt(proc?.cuotas) || 1);
  const difCuotas  = skyCuotas !== procCuotas;

  // Solo elevar a DIF. CUOTAS si el cruce de comercio y procesadora está OK;
  // si ya hay MAL FACTURADO o COM. ERRADO ese problema es más grave.
  if (difCuotas && (estado==='OK' || estado==='OK (equiv.)' || estado==='OK (integrado)')) {
    estado = 'DIF. CUOTAS';
  }

  // ── Diferencia de procesadora ────────────────────────────────────
  // Flag explícito (el estado ya queda como 'MAL FACTURADO: ...' en esos casos)
  const difProcesadora = estado.startsWith('MAL FACTURADO');

  // Normalizar signo del monto de procesadora para devoluciones
  const procMontoNorm = proc
    ? (s.esNeg && proc.monto > 0 ? -proc.monto : proc.monto)
    : null;

  // Calcular diferencia de monto real usando valores absolutos
  const difMonto = proc && s.monto && proc.monto
    ? Math.abs(Math.abs(s.monto) - Math.abs(proc.monto))
    : null;

  return { sky:s, proc, procMontoNorm, metodo, estado, procEncontrada:procReal, procEsperada:procEsp,
    comOK, sucOK, matchParcial:'', esDevolucion:s.esNeg?'SI':'NO', esAnulSinCobro:'NO',
    difCuotas, skyCuotas, skyCuotasTM, procCuotas, difProcesadora,
    difMonto: difMonto !== null ? +difMonto.toFixed(2) : null };
}

// ── ANULACIONES SIN COBRO ────────────────────────────────
function detectarAnulaciones(resultado) {
  const grupos={};
  for (const row of resultado) {
    const s=row.sky;
    const k=`${s.lote}|${s.cupon}|${s.nroCom}|${s.suc}|${s.fecha}|${s.montoN}`;
    (grupos[k]=grupos[k]||[]).push(row);
  }
  for (const filas of Object.values(grupos)) {
    const ps=filas.filter(f=>!f.sky.esNeg&&f.estado==='SIN MATCH');
    const ns=filas.filter(f=>f.sky.esNeg&&f.estado==='SIN MATCH');
    for (let i=0;i<Math.min(ps.length,ns.length);i++) {
      ps[i].estado='ANULACION SIN COBRO'; ps[i].esAnulSinCobro='SI';
      ns[i].estado='ANULACION SIN COBRO'; ns[i].esAnulSinCobro='SI';
    }
  }
}

// ── CLAVE ESTABLE POR ASIENTO (sobrevive re-ejecuciones del cruce) ──
// Usa Nro.Asiento si existe; si no, compone clave por fecha+cupon+monto.
function _skyKey(sky) {
  const a = String(sky?.asiento ?? '').trim();
  if (a && a !== 'null' && a !== '0' && a !== 'undefined') return `A:${a}`;
  return `K:${sky?.fecha}_${sky?.cupon}_${Math.abs(sky?.monto ?? 0).toFixed(2)}`;
}

// ── CORRECCIONES MANUALES ────────────────────────────────
function aplicarCorreccionesManuales() {
  for (const [key, cor] of Object.entries(CORREGIDAS)) {
    const idx = RESULTADO.findIndex(r => _skyKey(r.sky) === key);
    const fila = idx >= 0 ? RESULTADO[idx] : null;
    if (!fila) continue;
    fila.correccionManual=cor;
    fila.estado=`CORREGIDO MANUAL (${cor.proc})`;
    fila.procEncontrada=cor.proc; fila.metodo='MANUAL';
    if (!fila.proc && cor.cupon && cor.proc && (_FIS_NORM.length || _GP_NORM.length)) {
      const res = recruzarFila(idx, cor.cupon, cor.proc, cor.metodo || '', cor.montoReal ?? null);
      if (res?.ok) {
        fila.proc   = res.match;
        fila.comOK  = res.comOK;
        fila.sucOK  = res.sucOK;
        fila.estado = res.estado;
      }
    }
  }
}

// ── PROCESO PRINCIPAL ────────────────────────────────────
async function conciliar() {
  const btn=document.getElementById('run-btn');
  btn.disabled=true; btn.classList.add('running');
  document.getElementById('run-lbl').textContent='Procesando...';
  document.getElementById('prog').style.display='block';

  document.getElementById('t-empty')?.classList.remove('active');
  document.getElementById('t-log')?.classList.add('active');
  document.getElementById('tab-strip-cruce').style.display='flex';
  document.getElementById('log-dot').className='dot run';
  document.getElementById('log-title').textContent='Ejecutando conciliación...';
  clearLog();

  const setP=p=>document.getElementById('prog-bar').style.width=p+'%';

  try {
    log('Parseando Terminales...'); setP(5); await delay(20);
    const {term2suc,nombre2suc}=parseTerminales(FILES.ter.wb);
    log(`${Object.keys(term2suc).length} terminales mapeadas`,'ok');

    log('Parseando Skylab...'); setP(15); await delay(20);
    const skyRows=parseSkylab(FILES.sky.wb);
    const integrados=skyRows.filter(r=>r.integrado).length;
    const activos=skyRows.filter(r=>!r.integrado&&!r.esNeg).length;
    log(`${skyRows.length.toLocaleString()} ops · ${integrados} integradas (cruzadas) · ${activos} a conciliar`,'ok');
    // DEBUG: mostrar primeras 3 filas para diagnóstico
    if (skyRows.length > 0) {
      const s0 = skyRows[0];
      log(`  [DEBUG] SKY[0]: suc=${s0.suc} lote=${s0.lote} cupon=${s0.cupon} monto=${s0.monto} montoN=${s0.montoN} fecha=${s0.fecha} tarjeta=${s0.tarjeta}`);
    }
    window._debugSkyRows = skyRows; // debug

    // ── FISERV (si habilitado y archivo cargado)
    let fisNorm=[], fisRev=[];
    if (isProcEnabled('FISERV') && FILES.fis) {
      log('Parseando FISERV...'); setP(28); await delay(20);
      const r=parseFiserv(FILES.fis.wb,term2suc); fisNorm=r.norm; fisRev=r.rev;
      log(`${fisNorm.length.toLocaleString()} compras · ${fisRev.length} reversos`,'ok');
      window._debugFisRows = fisNorm;
    } else {
      log(isProcEnabled('FISERV') ? 'FISERV: sin archivo (saltando)' : 'FISERV: deshabilitado','warn');
    }
    _FIS_NORM = fisNorm; _FIS_REV = fisRev;

    // ── GETPOS (si habilitado y archivo cargado)
    let gpNorm=[], gpRev=[];
    if (isProcEnabled('GETPOS') && FILES.gp) {
      log('Parseando GETPOS...'); setP(40); await delay(20);
      const r=parseGetpos(FILES.gp.wb,nombre2suc); gpNorm=r.norm; gpRev=r.rev;
      log(`${gpNorm.length.toLocaleString()} ventas · ${gpRev.length} devoluciones`,'ok');
      window._debugGpRows = gpNorm;
    } else {
      log(isProcEnabled('GETPOS') ? 'GETPOS: sin archivo (saltando)' : 'GETPOS: deshabilitado','warn');
    }
    _GP_NORM = gpNorm; _GP_REV = gpRev;

    // ── Go Cuotas (si habilitado y datos cargados)
    const gocCount = (typeof _GOC_PAGOS !== 'undefined') ? _GOC_PAGOS.length : 0;
    if (isProcEnabled('GOCUOTAS') && gocCount > 0) {
      log(`Go Cuotas: ${gocCount} órdenes cargadas`,'ok');
    } else if (isProcEnabled('GOCUOTAS')) {
      log('Go Cuotas: sin archivo CSV (las ops quedarán Sin Match)','warn');
    }

    log(`Aplicando cascada de cruce (FIS:${fisNorm.length} GP:${gpNorm.length} GoC:${gocCount})...`);
    setP(58); await delay(20);
    RESULTADO=conciliarRows(skyRows,fisNorm,fisRev,gpNorm,gpRev);
    // Guardar operaciones de procesadora sin cruce con SKY
    const _usedFis = new Set(RESULTADO.filter(r=>r.proc&&r.procEncontrada==='FISERV').map(r=>r.proc));
    const _usedGp  = new Set(RESULTADO.filter(r=>r.proc&&r.procEncontrada==='GETPOS').map(r=>r.proc));
    window._FIS_NO_CRUZADAS = fisNorm.filter(r => !_usedFis.has(r));
    window._GP_NO_CRUZADAS  = gpNorm.filter(r  => !_usedGp.has(r));
    setP(74); await delay(20);

    log('Detectando anulaciones sin cobro...'); setP(82); await delay(20);
    detectarAnulaciones(RESULTADO);

    aplicarCorreccionesManuales();

    setP(92); await delay(20);
    renderTodo();
    setP(100);

    document.getElementById('log-dot').className='dot ok';
    document.getElementById('log-title').textContent='Completado';

    const st=contarEstados();
    // Desglose por método
    const metCount = {};
    for (const r of RESULTADO) { metCount[r.metodo||'?'] = (metCount[r.metodo||'?']||0)+1; }
    const metStr = Object.entries(metCount).sort((a,b)=>b[1]-a[1])
      .map(([m,n])=>`${m.split(':')[0]}:${n}`).join(' · ');
    log(`OK: ${st.ok} · Sin Match: ${st.sin} · Mal Fact.: ${st.mal} · Com.Errado: ${st.com} · Integradas: ${st.int}` + (st.urg?` · Rev.Urgente: ${st.urg}`:'') + (st.ref?` · Refacturado: ${st.ref}`:''),'ok');
    log(`  Métodos: ${metStr}`);

    SESSION.id=SESSION.id||('ses_'+Date.now());
    scheduleAutoSave();
    document.getElementById('dl-bar').classList.add('show');
    document.getElementById('dl-resumen').innerHTML =
      `<b>${RESULTADO.length.toLocaleString()} ops conciliadas</b> · ${st.sin} sin match · ${integrados} integradas`;
    setupDownloads();

  } catch(e) {
    document.getElementById('log-dot').className='dot err';
    log(`ERROR: ${e.message}`,'err');
    console.error(e);
  }
  btn.disabled=false; btn.classList.remove('running');
  document.getElementById('run-lbl').textContent='Reprocesar';
  document.getElementById('run-icon').textContent='↺';
}

function contarEstados() {
  return {
    ok:   RESULTADO.filter(r=>r.estado==='OK'||r.estado==='OK (equiv.)'||r.estado==='OK (integrado)').length,
    sin:  RESULTADO.filter(r=>r.estado==='SIN MATCH').length,
    mal:  RESULTADO.filter(r=>r.estado?.startsWith('MAL FACTURADO')).length,
    com:  RESULTADO.filter(r=>r.estado==='COM. ERRADO').length,
    int:  RESULTADO.filter(r=>r.estado==='INTEGRADO').length,
    urg:  RESULTADO.filter(r=>r.estado==='REVISION URGENTE').length,
    ref:  RESULTADO.filter(r=>r.estado==='REFACTURADO').length,
    anul: RESULTADO.filter(r=>r.estado==='ANULACION SIN COBRO').length,
    dev:  RESULTADO.filter(r=>r.sky.esNeg).length,
    dif:  RESULTADO.filter(r=>r.estado==='DIF. CUOTAS').length,
  };
}

// ── RENDER GENERAL ───────────────────────────────────────
function renderTodo() {
  renderTablas();
  updateCounts();
  renderTablaCorrecciones();
  renderTablaUrgente();
  renderTablaRefacturado();
  document.getElementById('dashboard').style.display='grid';
}

function renderTablas() {
  renderTable('tbl-all',  RESULTADO);
  renderTable('tbl-sin',  RESULTADO.filter(r=>r.estado==='SIN MATCH'));
  renderTable('tbl-mal',  RESULTADO.filter(r=>r.estado?.startsWith('MAL FACTURADO')));
  renderTablaCom(RESULTADO.filter(r=>r.estado==='COM. ERRADO'));
  renderTable('tbl-anul', RESULTADO.filter(r=>r.estado==='ANULACION SIN COBRO'));
  renderTable('tbl-dev',  RESULTADO.filter(r=>r.sky.esNeg));
  document.getElementById('mcnt-revision').textContent = RESULTADO.filter(r=>r.estado==='SIN MATCH').length;
  document.getElementById('mcnt-diferencias').textContent = RESULTADO.filter(r=>r.proc&&!r.sky.integrado).length;
}

function updateCounts() {
  const st=contarEstados();
  const total=RESULTADO.length;
  document.getElementById('k-total').textContent = total.toLocaleString();
  document.getElementById('k-ok').textContent    = st.ok.toLocaleString();
  document.getElementById('k-ok2').textContent   = RESULTADO.filter(r=>r.estado==='OK (equiv.)').length.toLocaleString();
  document.getElementById('k-sin').textContent   = st.sin.toLocaleString();
  document.getElementById('k-mal').textContent   = st.mal.toLocaleString();
  document.getElementById('k-com').textContent   = st.com.toLocaleString();
  document.getElementById('k-int').textContent   = st.int.toLocaleString();
  document.getElementById('cnt-all').textContent  = total.toLocaleString();
  document.getElementById('cnt-sin').textContent  = (st.sin+st.urg+st.ref).toLocaleString();
  document.getElementById('cnt-mal').textContent  = st.mal.toLocaleString();
  document.getElementById('cnt-com').textContent  = st.com.toLocaleString();
  document.getElementById('cnt-anul').textContent = st.anul.toLocaleString();
  document.getElementById('cnt-dev').textContent  = st.dev.toLocaleString();
  document.getElementById('mcnt-cruce').textContent = total.toLocaleString();
  // Nuevos estados especiales
  const urgCount = RESULTADO.filter(r=>r.estado==='REVISION URGENTE').length;
  const refCount = RESULTADO.filter(r=>r.estado==='REFACTURADO').length;
  const kUrg = document.getElementById('k-urg');
  const kRef = document.getElementById('k-ref');
  if (kUrg) kUrg.textContent = urgCount;
  if (kRef) kRef.textContent = refCount;
  // Actualizar counters de los tabs
  const cntUrgTab = document.getElementById('cnt-urg');
  const cntRefTab = document.getElementById('cnt-ref');
  if (cntUrgTab) cntUrgTab.textContent = urgCount;
  if (cntRefTab) cntRefTab.textContent = refCount;
  const corCount = Object.keys(CORREGIDAS).length;
  const corCnt = document.getElementById('cnt-cor');
  if (corCnt) corCnt.textContent = corCount.toLocaleString();

  // Contadores de las nuevas tabs de diferencias
  const difCuotasCount = st.dif;
  const difProcCount   = st.mal;
  const elDifCuotas = document.getElementById('cnt-dif-cuotas');
  const elDifProc   = document.getElementById('cnt-dif-proc');
  if (elDifCuotas) elDifCuotas.textContent = difCuotasCount;
  if (elDifProc)   elDifProc.textContent   = difProcCount;

  // Badge del módulo 3 muestra la suma de diferencias accionables
  const mcntDif = document.getElementById('mcnt-diferencias');
  if (mcntDif) mcntDif.textContent = (difCuotasCount + difProcCount).toLocaleString();
}

// ── RENDER TABLA ─────────────────────────────────────────
function estadoBadge(est) {
  const m={
    'OK':               'st-ok',
    'OK (integrado)':   'st-int',
    'OK (equiv.)':      'st-ok2',
    'SIN MATCH':        'st-sin',
    'COM. ERRADO':      'st-com',
    'ANULACION SIN COBRO':'st-anul',
    'INTEGRADO':        'st-int',
    'REVISION URGENTE': 'st-urgente',
    'REFACTURADO':      'st-refact',
    'DIF. CUOTAS':      'st-dif-cuotas',
  };
  if (!est) return '';
  const cls=m[est]||(est.includes('GETPOS')?'st-gp':est.includes('FISERV')?'st-fis':'st-mal');
  const label=est.replace('MAL FACTURADO: debe ser ','→').replace('CORREGIDO MANUAL ','✓');
  return `<span class="st ${cls}">${label}</span>`;
}

function rowClass(est) {
  if (!est) return '';
  if (est==='OK'||est==='OK (equiv.)'||est==='OK (integrado)') return 'row-ok';
  if (est==='INTEGRADO') return 'row-int';
  if (est==='SIN MATCH') return 'row-sin';
  if (est==='ANULACION SIN COBRO') return 'row-anul';
  if (est?.startsWith('MAL')) return 'row-mal';
  if (est==='COM. ERRADO') return 'row-com';
  if (est==='REVISION URGENTE') return 'row-urgente';
  if (est==='REFACTURADO') return 'row-refact';
  if (est==='DIF. CUOTAS') return 'row-dif';
  return '';
}

function toRow(r) {
  const s=r.sky, p=r.proc;
  const monto=`<span class="num ${s.monto<0?'num-neg':''}">${fmtARS(s.monto)}</span>`;
  const procMonto = r.procMontoNorm ?? p?.monto ?? null;
  const difVal = (procMonto !== null && s.monto !== undefined) ? s.monto - procMonto : null;
  const dif = difVal !== null
    ? `<span class="num ${difVal<0?'num-neg':''}">${fmtARS(difVal)}</span>`
    : '';
  return [
    estadoBadge(r.estado),
    `<span class="met">${r.metodo??''}</span>`,
    r.procEsperada??'', r.procEncontrada??'',
    s.integrado?'<span class="st st-int">SI</span>':'',
    s.suc, s.tarjeta, s.plan, s.fecha, monto, s.cupon, s.lote,
    p?.tarjeta??'', p?.cuotas??'', p?.comFis??'',
    p?.aut??'', p?.lote??'', p?.ticket??'', p?.suc??'',
    dif, r.comOK??'', r.matchParcial??'',
  ];
}

function renderTable(tblId, filas) {
  const t=document.getElementById(tblId); if (!t) return;
  t.querySelector('thead').innerHTML=`<tr>${HDR_BASE.map(h=>`<th>${h}</th>`).join('')}</tr>`;
  t.querySelector('tbody').innerHTML=filas.slice(0,3000).map(r=>{
    const cells=toRow(r);
    return `<tr class="${rowClass(r.estado)}">${cells.map(c=>`<td>${c}</td>`).join('')}</tr>`;
  }).join('');
}

// ── CORRECCIÓN MANUAL ────────────────────────────────────
const FILTROS_FIX = { suc:'', tar:'', proc:'', fecha:'', vend:'', search:'' };

function poblarFiltros() {
  const sinMatch=RESULTADO.filter(r=>r.estado==='SIN MATCH');
  const selSuc=document.getElementById('flt-suc'), curSuc=selSuc?.value||'';
  const selTar=document.getElementById('flt-tar'), curTar=selTar?.value||'';
  const selFec=document.getElementById('flt-fecha'), curFec=selFec?.value||'';
  const selVend=document.getElementById('flt-vend'), curVend=selVend?.value||'';

  if (selSuc) {
    const sucs=[...new Set(sinMatch.map(r=>r.sky.suc))].sort((a,b)=>+a-+b);
    selSuc.innerHTML='<option value="">Todas las suc.</option>'+
      sucs.map(s=>`<option value="${s}" ${s===curSuc?'selected':''}>${s}</option>`).join('');
  }
  if (selTar) {
    const tars=[...new Set(sinMatch.map(r=>r.sky.tarjeta))].sort();
    selTar.innerHTML='<option value="">Todas las tarjetas</option>'+
      tars.map(t=>`<option value="${t}" ${t===curTar?'selected':''}>${t}</option>`).join('');
  }
  if (selFec) {
    const fechas=[...new Set(sinMatch.map(r=>r.sky.fecha))].sort().reverse();
    selFec.innerHTML='<option value="">Todas las fechas</option>'+
      fechas.map(f=>`<option value="${f}" ${f===curFec?'selected':''}>${f}</option>`).join('');
  }
  if (selVend) {
    const vends=[...new Set(sinMatch.map(r=>r.sky.vendedor||'').filter(Boolean))].sort();
    selVend.innerHTML='<option value="">Todos los vendedores</option>'+
      vends.map(v=>`<option value="${v}" ${v===curVend?'selected':''}>${v}</option>`).join('');
  }
}

function aplicarFiltros() {
  FILTROS_FIX.suc    = document.getElementById('flt-suc')?.value    || '';
  FILTROS_FIX.tar    = document.getElementById('flt-tar')?.value    || '';
  FILTROS_FIX.proc   = document.getElementById('flt-proc')?.value   || '';
  FILTROS_FIX.fecha  = document.getElementById('flt-fecha')?.value  || '';
  FILTROS_FIX.vend   = document.getElementById('flt-vend')?.value   || '';
  FILTROS_FIX.search = (document.getElementById('flt-search')?.value||'').trim().toLowerCase();
  renderFilas();
}

function limpiarFiltros() {
  ['flt-suc','flt-tar','flt-proc','flt-fecha','flt-vend','flt-search'].forEach(id=>{
    const el=document.getElementById(id); if(el) el.value='';
  });
  Object.keys(FILTROS_FIX).forEach(k=>FILTROS_FIX[k]='');
  renderFilas();
}


// Genera options para el selector de método según procesadora
function buildMetodosOptions(proc, metodoActual) {
  const lista = proc==='GETPOS' ? METODOS_GETPOS : METODOS_FISERV;
  return lista.map(m=>
    `<option value="${m.val}" ${m.val===metodoActual?'selected':''}>${m.label}</option>`
  ).join('');
}

// Actualiza el selector de método cuando cambia la procesadora
function actualizarMetodos(idx) {
  const proc = document.getElementById(`fi-proc-${idx}`)?.value || '';
  const sel  = document.getElementById(`fi-met-${idx}`);
  if (sel) sel.innerHTML = buildMetodosOptions(proc, '');
}

// Marcar una operación sin match con estado especial
function marcarRevision(idx, estado) {
  idx = parseInt(idx);
  const fila = RESULTADO[idx]; if (!fila) return;
  fila.estado         = estado;
  fila.metodo         = estado;
  fila.procEncontrada = estado === 'REFACTURADO' ? 'N/A' : fila.procEsperada;
  fila.matchParcial   = estado === 'REFACTURADO'
    ? 'Marcado manualmente como refacturado — no coincide por refacturación'
    : estado === 'ANULACION SIN COBRO'
    ? 'Marcado manualmente como anulación sin cobro'
    : 'Marcado manualmente como revisión urgente — venta sin pago encontrado';
  // Flags de tipo
  if (estado === 'ANULACION SIN COBRO') fila.esAnulSinCobro = 'SI';
  // Guardar en CORREGIDAS con estado especial
  const cor = CORREGIDAS[_skyKey(fila.sky)] || {};
  CORREGIDAS[_skyKey(fila.sky)] = { ...cor, cupon: cor.cupon||'—', proc: cor.proc||fila.procEsperada,
    resultado: estado, metodo: estado };
  renderTablas(); updateCounts(); renderFilas(); renderTablaCorrecciones();
  renderTablaUrgente(); renderTablaRefacturado();
  scheduleAutoSave();
}

function filtrarSinMatch() {
  // Incluir positivos Y negativos sin match (devoluciones no encontradas)
  let rows=RESULTADO.filter(r=>['SIN MATCH','REVISION URGENTE','REFACTURADO'].includes(r.estado));
  if (FILTROS_FIX.suc)   rows=rows.filter(r=>r.sky.suc===FILTROS_FIX.suc);
  if (FILTROS_FIX.tar)   rows=rows.filter(r=>r.sky.tarjeta===FILTROS_FIX.tar);
  if (FILTROS_FIX.proc)  rows=rows.filter(r=>r.sky.esGETPos===FILTROS_FIX.proc==='GETPOS'||(!r.sky.esGETPos)===FILTROS_FIX.proc==='FISERV');
  if (FILTROS_FIX.fecha) rows=rows.filter(r=>r.sky.fecha===FILTROS_FIX.fecha);
  if (FILTROS_FIX.vend)  rows=rows.filter(r=>r.sky.vendedor===FILTROS_FIX.vend);
  if (FILTROS_FIX.search) {
    const q=FILTROS_FIX.search;
    rows=rows.filter(r=>{
      const s=r.sky;
      return s.cupon.includes(q)||s.montoN.includes(q)||
        (s.vendedor||'').toLowerCase().includes(q)||
        String(s.asiento||'').includes(q)||s.lote.includes(q)||s.fecha.includes(q);
    });
  }
  return rows;
}

// Tipo de operación (Venta / Devolución / Anulación)
// Nota: esDevolucion y esAnulSinCobro pueden ser 'SI'/'NO' (string) o true/false
function _tipoOp(r) {
  const isAnul = r.esAnulSinCobro === 'SI' || r.esAnulSinCobro === true;
  const isDev  = r.esDevolucion   === 'SI' || r.esDevolucion   === true;
  if (isAnul) return { label:'ANULACIÒN S/C', color:'#fb923c' };
  if (isDev)  return { label:'DEVOLUCION',    color:'#38bdf8' };
  if (r.sky?.esNeg) return { label:'NEGATIVO', color:'#f87171' };
  return              { label:'VENTA',         color:'#34d399' };
}

// Resetear cualquier fila a SIN MATCH (sin importar si está en CORREGIDAS)
function resetearFila(idx) {
  idx = parseInt(idx);
  const fila = RESULTADO[idx]; if (!fila) return;
  const key = _skyKey(fila.sky);
  if (CORREGIDAS[key]) delete CORREGIDAS[key];
  fila.proc = null; fila.metodo = 'SIN MATCH'; fila.estado = 'SIN MATCH';
  fila.procEncontrada = ''; fila.comOK = ''; fila.sucOK = '';
  fila.matchParcial = ''; fila.correccionManual = null;
  renderTablas(); updateCounts(); renderFilas();
  renderTablaCorrecciones(); renderTablaUrgente(); renderTablaRefacturado();
  scheduleAutoSave();
}

function renderFilas() {
  const filtrados=filtrarSinMatch();
  const total=RESULTADO.filter(r=>r.estado==='SIN MATCH').length;
  const corr=filtrados.filter(r=>!!CORREGIDAS[_skyKey(r.sky)]).length;
  const pend=filtrados.length-corr;
  const stats=document.getElementById('fix-stats');
  if (stats) stats.innerHTML=`Mostrando <b>${filtrados.length}</b> / <b>${total}</b> · `+
    `<b style="color:var(--grn)">${corr} ok</b> · <b style="color:var(--yel)">${pend} pendientes</b>`;
  const cont=document.getElementById('fix-content'); if (!cont) return;
  if (!filtrados.length) {
    cont.innerHTML=`<div class="fix-empty">${total===0?'No hay Sin Match.':'Sin resultados para los filtros.'}</div>`;
    return;
  }
  const hdr=`<div class="fix-hdr">
    <div>Suc · Asiento · Cupón</div><div>Plan · Lote · Vendedor</div>
    <div>Monto facturado / Fecha</div><div>Monto real cobrado</div>
    <div>Cupón procesadora</div>
    <div>Procesadora</div><div>Método de cruce</div><div></div></div>`;
  cont.innerHTML=hdr+filtrados.map(r=>{
    const s=r.sky, cor=CORREGIDAS[_skyKey(s)]||{};
    const applied=!!cor.cupon;
    const procDef=s.esGETPos?'GETPOS':'FISERV';
    const tipo=_tipoOp(r);
    const montoColor = s.monto < 0 ? 'var(--red)' : 'var(--grn)';
    const montoFmt   = s.monto < 0
      ? `−${fmtARS(Math.abs(s.monto))}`
      : fmtARS(s.monto);
    return `<div class="fix-row" id="fr-${s.idx}">
      <div><div class="fix-cell-lbl">Suc · Asiento · Cupón</div>
        <div class="fix-cell-val"><b style="color:var(--cyn)">${s.suc}</b> · ${s.asiento??'—'} · <b>${s.cupon}</b></div>
        <div class="fix-cell-sub" style="display:flex;align-items:center;gap:5px;flex-wrap:wrap">
          <span>${s.tarjeta}</span>
          <span style="font-size:7px;padding:1px 5px;border-radius:3px;
            color:${tipo.color};border:1px solid ${tipo.color}55;background:${tipo.color}18">${tipo.label}</span>
        </div></div>
      <div><div class="fix-cell-lbl">Plan · Lote</div>
        <div class="fix-cell-val">${s.plan} · Lote ${s.lote}</div>
        <div class="fix-cell-sub">${s.vendedor??''}</div></div>
      <div><div class="fix-cell-lbl">Monto facturado · Fecha</div>
        <div class="fix-cell-val fix-monto" style="color:${montoColor}">${montoFmt}</div>
        <div class="fix-cell-sub">${s.fecha}</div></div>
      <div>
        <div class="fix-cell-lbl">Monto real cobrado</div>
        <input class="fix-inp" id="fi-mon-${s.idx}" type="number" step="0.01"
          placeholder="Igual al facturado"
          title="Dejá vacío si el monto cobrado es igual al facturado"
          value="${cor.montoReal != null ? cor.montoReal : ''}"
          style="font-family:var(--mono);text-align:right">
        ${cor.montoReal != null ? `<div class="fix-cell-sub" style="color:${cor.montoReal != Math.abs(s.monto) ? 'var(--yel)' : 'var(--grn)'}">
          ${cor.montoReal != Math.abs(s.monto) ? '⚠ Difiere del facturado' : '✓ Igual al facturado'}</div>` : ''}
      </div>
      <div><input class="fix-inp" id="fi-cup-${s.idx}" placeholder="Cupón / Autorización" value="${cor.cupon||''}"></div>
      <div>
        <select class="fix-sel" id="fi-proc-${s.idx}" onchange="actualizarMetodos(${s.idx})">
          <option value="">Seleccionar...</option>
          <option value="FISERV" ${(cor.proc||procDef)==='FISERV'?'selected':''}>FISERV</option>
          <option value="GETPOS" ${(cor.proc||procDef)==='GETPOS'?'selected':''}>GETPOS</option>
        </select>
      </div>
      <div>
        <select class="fix-sel" id="fi-met-${s.idx}" style="font-size:8px">
          ${buildMetodosOptions(cor.proc||procDef, cor.metodo||'')}
        </select>
      </div>
      <div style="display:flex;flex-direction:column;gap:4px">
        <button class="fix-apply ${applied?'applied':''}" id="fi-btn-${s.idx}"
          onclick="aplicarCorreccion(${s.idx})">
          ${applied?'✓ Ok':'Aplicar'}</button>
        <button class="fix-apply" style="background:var(--org);font-size:7px"
          onclick="marcarRevision('${s.idx}','REVISION URGENTE')" title="Venta en Skylab pero no se encontró el pago">
          ⚠ Rev. Urgente</button>
        <button class="fix-apply" style="background:#7c3aed;font-size:7px"
          onclick="marcarRevision('${s.idx}','REFACTURADO')" title="La operación fue refacturada — no coincide por refacturación">
          ↺ Refacturado</button>
        <button class="fix-apply" style="background:#0369a1;font-size:7px"
          onclick="marcarRevision('${s.idx}','ANULACION SIN COBRO')" title="Venta anulada — no se cobró y no debe buscarse en procesadora">
          ✗ Anu. S/Cobro</button>
      </div>
    </div>`;
  }).join('');
}

function renderFixPanel() { poblarFiltros(); renderFilas(); }

function aplicarCorreccion(idx) {
  const cupon=document.getElementById(`fi-cup-${idx}`)?.value?.trim();
  const proc =document.getElementById(`fi-proc-${idx}`)?.value;
  const metodo=document.getElementById(`fi-met-${idx}`)?.value||'';
  if (!cupon||!proc) { alert('Completá el cupón y seleccioná la procesadora.'); return; }

  const fila=RESULTADO[idx];
  const key=_skyKey(fila.sky);
  const antes=CORREGIDAS[key]||{};

  // Monto real cobrado (vacío = igual al facturado)
  const montoRealRaw = document.getElementById(`fi-mon-${idx}`)?.value?.trim();
  const montoReal = montoRealRaw !== '' && montoRealRaw != null
    ? parseFloat(montoRealRaw.replace(',','.')) || null
    : null;

  // Re-cruzar usando montoReal si fue ingresado, sino monto SKY
  const res = (_FIS_NORM.length || _GP_NORM.length)
    ? recruzarFila(idx, cupon, proc, metodo, montoReal)
    : null;

  CORREGIDAS[key]={cupon, proc, metodo,
    montoReal,                              // null = igual al facturado
    difMonto: res?.difMonto ?? null,
    resultado: res ? (res.ok ? 'CRUZADO' : 'NO CRUZADO') : 'PENDIENTE',
    metodo: res?.met || '',
    motivo: res?.motivo || ''
  };

  registrarCorreccion({
    skyIdx:idx, campo:'Procesadora/Cupón',
    valorAntes:antes.proc?`${antes.proc} - ${antes.cupon}`:'Sin match',
    valorDespues:`${proc} - ${cupon}`,
    motivo:'Corrección manual',
    obs: res ? (res.ok ? `Cruzado: ${res.met}` : res.motivo) : '',
  });

  if (fila) {
    if (res?.ok) {
      fila.proc           = res.match;
      fila.metodo         = res.met;
      fila.estado         = res.estado;
      fila.procEncontrada = proc;
      fila.comOK          = res.comOK;
      fila.sucOK          = res.sucOK;
      fila.matchParcial   = `Manual: ${cupon}`;
      fila.correccionManual = CORREGIDAS[key];
    } else {
      fila.estado         = 'REVISION URGENTE';
      fila.procEncontrada = proc;
      fila.metodo         = 'REVISION URGENTE';
      fila.matchParcial   = res?.motivo || `Cupón manual: ${cupon}`;
      fila.correccionManual = CORREGIDAS[key];
    }
  }

  const btn=document.getElementById(`fi-btn-${idx}`);
  if (btn) {
    const ok = res?.ok;
    btn.textContent = ok ? '✓ Cruzado' : (res ? '✗ No encontrado' : '✓ Guardado');
    btn.classList.remove('applied');
    btn.classList.add(ok ? 'applied' : 'applied-fail');
    // Mostrar feedback en la fila
    const fr = document.getElementById(`fr-${idx}`);
    if (fr) {
      const existing = fr.querySelector('.cor-result');
      if (existing) existing.remove();
      const div = document.createElement('div');
      div.className = 'cor-result';
      div.style.cssText = `grid-column:1/-1;font-size:9px;padding:4px 0;
        color:${ok?'var(--grn)':'var(--red)'}`;
      div.textContent = ok
        ? `✓ Cruzado por ${res.met} · Estado: ${res.estado}`
        : `✗ ${res?.motivo||'No encontrado'}`;
      fr.appendChild(div);
    }
  }

  renderTablas(); updateCounts();
  // Re-renderizar pestaña de correcciones
  renderTablaCorrecciones();
  renderFilas();
  scheduleAutoSave();
}

// ── FILE LOADING ─────────────────────────────────────────
function loadFile(input, key) {
  const file=input.files[0]; if (!file) return;
  const r=new FileReader();
  r.onload=e=>{
    try {
      FILES[key]={wb:XLSX.read(e.target.result,{type:'array',cellDates:true}),name:file.name};
      document.getElementById(`fc-${key}`).classList.add('ok');
      const st=document.getElementById(`st-${key}`);
      st.textContent=`✓ ${file.name}`; st.className='fc-st ok';
    } catch { document.getElementById(`st-${key}`).textContent='✗ Error'; }
    checkReady();
  };
  r.readAsArrayBuffer(file);
}

// ── CARGA ESPECIAL: Liquidaciones (módulo COBROS) ─────────
function loadLiquidaciones(input) {
  const file = input.files[0]; if (!file) return;
  const fc = document.getElementById('fc-liq');
  const st = document.getElementById('st-liq');
  st.textContent = '↻ Cargando...'; st.className = 'fc-st';
  const r = new FileReader();
  r.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, { type:'array', cellDates:true });
      FILES.liq = { wb, name: file.name };
      _LIQ_NORM = parseLiquidaciones(wb);
      fc.classList.add('ok');
      st.textContent = `✓ ${file.name} (${_LIQ_NORM.length.toLocaleString()} liq.)`;
      st.className = 'fc-st ok';
      // Si el módulo cobros está activo, re-renderizar
      if (document.getElementById('mod-cobros')?.classList.contains('active')) {
        renderModuloCobros();
      }
      // Actualizar badge
      const badge = document.getElementById('mcnt-cobros');
      if (badge && COBROS_RESULT.length === 0 && RESULTADO.length > 0) {
        // Ejecutar cruce silencioso para actualizar el badge
        cruzarCobros();
        const pend = COBROS_RESULT.filter(c => c.estado === 'PENDIENTE').length;
        badge.textContent = pend || '—';
      }
    } catch(err) {
      st.textContent = '✗ Error al leer'; st.className = 'fc-st';
      console.error('Error cargando liquidaciones:', err);
    }
  };
  r.readAsArrayBuffer(file);
}

// ── CARGA GO CUOTAS ──────────────────────────────────────────────
function loadGocSkylab(input) {
  const file = input.files[0]; if (!file) return;
  const st = document.getElementById('st-goc-sky');
  if (st) { st.textContent = '↻ Cargando...'; st.className = 'fc-st'; }
  const r = new FileReader();
  r.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, { type:'array', cellDates:true });
      parseGocSkylab(wb);
      const n = _GOC_SKY.length;
      document.getElementById('fc-goc-sky')?.classList.add('ok');
      if (st) { st.textContent = `✓ ${file.name} (${n} ops)`; st.className = 'fc-st ok'; }
      const badge = document.getElementById('mcnt-goc');
      if (badge) badge.textContent = n || '—';
      if (document.getElementById('mod-goc')?.closest('.mod-panel.active')) renderModuloGoCuotas();
    } catch(err) { if (st) { st.textContent = '✗ Error'; st.className = 'fc-st'; } console.error(err); }
  };
  r.readAsArrayBuffer(file);
}

async function loadGocPagos(input) {
  const file = input.files[0]; if (!file) return;
  const st = document.getElementById('st-goc-pag');
  if (st) { st.textContent = '↻ Cargando...'; st.className = 'fc-st'; }
  try {
    await parseGocPagos(file);
    const n = _GOC_PAGOS.length;
    document.getElementById('fc-goc-pag')?.classList.add('ok');
    if (st) { st.textContent = `✓ ${file.name} (${n} pagos)`; st.className = 'fc-st ok'; }
    if (document.getElementById('mod-goc')?.closest('.mod-panel.active')) renderModuloGoCuotas();
  } catch(err) { if (st) { st.textContent = '✗ Error'; st.className = 'fc-st'; } console.error(err); }
}

function loadGocVentas(input) {
  const file = input.files[0]; if (!file) return;
  const st = document.getElementById('st-goc-ven');
  if (st) { st.textContent = '↻ Cargando...'; st.className = 'fc-st'; }
  const r = new FileReader();
  r.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, { type:'array', cellDates:true });
      parseGocVentas(wb);
      const n = _GOC_VENTAS.length;
      document.getElementById('fc-goc-ven')?.classList.add('ok');
      if (st) { st.textContent = `✓ ${file.name} (${n} artículos)`; st.className = 'fc-st ok'; }
      if (document.getElementById('mod-goc')?.closest('.mod-panel.active')) renderModuloGoCuotas();
    } catch(err) { if (st) { st.textContent = '✗ Error'; st.className = 'fc-st'; } console.error(err); }
  };
  r.readAsArrayBuffer(file);
}

// ── CARGA CONTRACARGOS ────────────────────────────────────────────
function loadContracargos(input, tipo) {
  const file = input.files[0]; if (!file) return;
  const fc = document.getElementById(`fc-ctr-${tipo}`);
  const st = document.getElementById(`st-ctr-${tipo}`);
  if (st) { st.textContent = '↻ Cargando...'; st.className = 'fc-st'; }
  const r = new FileReader();
  r.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, { type:'array', cellDates:true });
      if (tipo === 'fis') {
        parseContrarcargosFiserv(wb);
        const n = _CTR_FIS.length;
        if (fc) fc.classList.add('ok');
        if (st) { st.textContent = `✓ ${file.name} (${n} registros)`; st.className = 'fc-st ok'; }
      } else {
        parseContracargosGetpos(wb);
        const n = _CTR_GP.length;
        if (fc) fc.classList.add('ok');
        if (st) { st.textContent = `✓ ${file.name} (${n} disputas)`; st.className = 'fc-st ok'; }
      }
      if (document.getElementById('mod-ctr')?.closest('.mod-panel.active')) {
        renderModuloContracargos();
      }
    } catch(err) {
      if (st) { st.textContent = '✗ Error al leer'; st.className = 'fc-st'; }
      console.error('[CTR] Error:', err);
    }
  };
  r.readAsArrayBuffer(file);
}

function checkReady() {
  // Solo Skylab + Terminales son requeridos; las procesadoras son opcionales
  // (sin archivos de proc, todas las ops quedan SIN MATCH — es un estado válido)
  const base = FILES.sky && FILES.ter;
  const ok   = !!base;

  const fisFok = isProcEnabled('FISERV')   && !!FILES.fis;
  const gpOk   = isProcEnabled('GETPOS')   && !!FILES.gp;
  const gocOk  = isProcEnabled('GOCUOTAS') &&
    typeof _GOC_PAGOS !== 'undefined' && _GOC_PAGOS.length > 0;

  document.getElementById('run-btn').disabled = !ok;
  document.getElementById('hdr-chip-files')?.remove();

  // Badge de procesadoras activas bajo el botón
  const procsDiv = document.getElementById('run-procs');
  if (procsDiv) {
    const badges = [];
    const C = { FISERV:'var(--acc)', GETPOS:'var(--grn)', GOCUOTAS:'var(--yel)' };
    const L = { FISERV:'FISERV', GETPOS:'GETPOS', GOCUOTAS:'GoC' };
    ['FISERV','GETPOS','GOCUOTAS'].forEach(id => {
      if (!isProcEnabled(id)) return;
      const color = C[id];
      const ready = (id==='FISERV'&&FILES.fis)||(id==='GETPOS'&&FILES.gp)||
                    (id==='GOCUOTAS'&&typeof _GOC_PAGOS!=='undefined'&&_GOC_PAGOS.length>0);
      badges.push(`<span style="padding:1px 6px;border-radius:3px;font-size:7px;
        color:${color};border:1px solid ${color}55;background:${color}18;
        opacity:${ready?'1':'0.4'}">${L[id]}${ready?'':' ○'}</span>`);
    });
    procsDiv.innerHTML = badges.join('');
  }

  // Mensaje de lo que falta
  if (!ok) {
    const miss = [];
    if (!FILES.sky) miss.push('Skylab');
    if (!FILES.ter) miss.push('Terminales');
    document.getElementById('save-indicator').textContent =
      `⚠ Falta: ${miss.join(', ')}`;
  } else {
    document.getElementById('save-indicator').textContent = '';
  }
}

// ── LOG ──────────────────────────────────────────────────
function clearLog() { document.getElementById('log-wrap').innerHTML=''; }
function log(msg, type='info') {
  const w=document.getElementById('log-wrap'); if (!w) return;
  const d=document.createElement('div');
  d.className=`log-line ${type}`;
  const ico=type==='ok'?'✓ ':type==='warn'?'⚠ ':type==='err'?'✗ ':'› ';
  d.textContent=ico+msg;
  w.appendChild(d); w.scrollTop=w.scrollHeight;
}

// ── DOWNLOADS ────────────────────────────────────────────
function setupDownloads() {
  document.getElementById('dl-fis').onclick   = ()=>descargar('FISERV');
  document.getElementById('dl-gp').onclick    = ()=>descargar('GETPOS');
  document.getElementById('dl-full').onclick  = ()=>descargarCompleto();
  document.getElementById('dl-log').onclick   = ()=>exportarLogCorrecciones();
  document.getElementById('dl-sin-btn').onclick = ()=>exportarNoConciliadas();
  document.getElementById('dl-res').onclick   = ()=>exportarResumenes();
}

function descargar(proc) {
  const filas=RESULTADO.filter(r=>
    r.procEncontrada===proc &&
    (r.estado==='COM. ERRADO'||r.estado?.startsWith('MAL FACTURADO')||r.estado?.startsWith('CORREGIDO'))
  );
  const HDR = proc==='FISERV'
    ? ['Fecha Venta','Nro Equipo','SUC','Nro de Lote','Nro de Cupón','Tarjeta','Cuotas',
       'Importe Venta','Nro Comercio Cobrado','Tasa Cobrada','Nro Comercio Facturado',
       'Tarjeta Facturada','Plan Facturado','TASA Facturada','DIF %','DIF $',
       'VENDEDOR','ID','N° FC','INTEGRADO']
    : ['Fecha','Suc','Nombre establecimiento','Marca','Tarjeta','POS','Cupon','Aut',
       'Plan Cobrado','Monto','TD Cobrada','Comercio Facturado','Tarjeta Facturada',
       'Plan Facturado','Monto Facturado','TD Facturada','Dif %','Dif $','Vendedor','ID','NRO FC'];
  const rows=filas.map(r=>{
    const s=r.sky, p=r.proc;
    if (proc==='FISERV') return [
      fmtFecha(s.fecha),p?.equipo??'',s.suc,p?.lote??'',p?.ticket??s.cupon,
      p?.tarjeta??s.tarjeta,p?.cuotas??'',fmtARS(s.monto),
      p?.comFis??'',r.tasaCobrada??'',r.comOK==='DIFERENTE'?p?.comFis??'':'',
      s.tarjeta,s.plan,r.tasaAcordada??'',r.difTasa??'',r.difMonto??'',
      s.vendedor??'',s.asiento??'',s.cupon,s.integrado?'INTEGRADO':'DESINTEGRADO',
    ];
    return [
      fmtFecha(s.fecha),s.suc,p?.nombre??'',p?.marca??'',p?.tarjeta??'',
      p?.pos??'',s.cupon,p?.aut??'',p?.plan??s.plan,fmtARS(s.monto),
      r.tasaCobrada??'',r.comOK==='DIFERENTE'?p?.comFis??'':'',
      s.tarjeta,s.plan,s.monto,r.tasaAcordada??'',r.difTasa??'',r.difMonto??'',
      s.vendedor??'',s.asiento??'',s.cupon,
    ];
  });
  _exportXlsx([HDR,...rows],`Descuentos ${proc}`,`Descuentos_${proc}_${hoy()}.xlsx`);
}

function descargarCompleto() {
  const HDR=['Estado','Método','Is Dev.','Anul.S/Cobro','Integrado',
    'Proc.Esp.','Proc.Real','Com.OK','Suc.OK','Match Parcial',
    'Fecha SKY','Suc. SKY','Tarjeta SKY','Plan SKY','Cupón SKY','Lote SKY','Com. SKY','Monto SKY','Neto SKY','Vendedor SKY','Asiento SKY',
    'Tarjeta Proc.','Cuotas Proc.','Monto Proc.','Dif. Monto','Cód.Auth. Proc.','Lote Proc.','Ticket Proc.','Suc. Proc.','Com.FIS Proc.','Terminal Proc.'];
  const rows=RESULTADO.map(r=>{
    const s=r.sky, p=r.proc;
    const procMonto = r.procMontoNorm ?? p?.monto ?? null;
    const difMonto  = procMonto !== null ? +(s.monto - procMonto).toFixed(2) : '';
    return [r.estado,r.metodo??'',r.esDevolucion??'NO',r.esAnulSinCobro??'NO',s.integrado?'SI':'NO',
      r.procEsperada??'',r.procEncontrada??'',r.comOK??'',r.sucOK??'',r.matchParcial??'',
      fmtFecha(s.fecha),s.suc,s.tarjeta,s.plan,s.cupon,s.lote,s.nroCom,s.monto,s.neto,
      s.vendedor??'',s.asiento??'',
      p?.tarjeta??'',p?.cuotas??'',procMonto??'',difMonto,p?.aut??'',p?.lote??'',p?.ticket??'',p?.suc??'',p?.comFis??'',p?.equipo??''];
  });
  _exportXlsx([HDR,...rows],'Conciliación completa',`Conciliacion_${hoy()}.xlsx`);
}

function _exportXlsx(data, sheetName, filename) {
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(data), sheetName);
  XLSX.writeFile(wb, filename);
}

function hoy() { return new Date().toISOString().slice(0,10); }

// ══════════════════════════════════════════════════════════════════
// TABLA Y RE-PROCESO DE CORRECCIONES MANUALES
// ══════════════════════════════════════════════════════════════════
const HDR_COR = ['Estado cruce','Método','Fecha','Suc','Vendedor','Tarjeta','Plan',
  'Monto SKY','Monto real','Monto Proc.','DIF $','Cupón SKY','Proc.','Cupón ingresado','Tarjeta proc.','Com.FIS','Com.OK','Detalle',''];


// Eliminar una corrección y devolver la fila a SIN MATCH
function eliminarCorreccion(idx) {
  idx = parseInt(idx);
  const fila = RESULTADO[idx];
  if (!fila) return;
  const key = _skyKey(fila.sky);
  if (!CORREGIDAS[key]) return;
  delete CORREGIDAS[key];
  // Restaurar la fila a SIN MATCH
  if (fila) {
    fila.proc           = null;
    fila.metodo         = 'SIN MATCH';
    fila.estado         = 'SIN MATCH';
    fila.procEncontrada = '';
    fila.comOK          = '';
    fila.sucOK          = '';
    fila.matchParcial   = '';
    fila.correccionManual = null;
  }
  renderTablas(); updateCounts(); renderFilas(); renderTablaCorrecciones();
  renderTablaUrgente(); renderTablaRefacturado();
  scheduleAutoSave();
}

function reprocesarCorrecciones() {
  let ok=0, fail=0;
  for (const [key, cor] of Object.entries(CORREGIDAS)) {
    const idx = RESULTADO.findIndex(r => _skyKey(r.sky) === key);
    if (idx < 0) continue;
    const fila = RESULTADO[idx];
    const res=recruzarFila(idx, cor.cupon, cor.proc, cor.metodo||'', cor.montoReal ?? null);
    CORREGIDAS[key] = {...cor, difMonto: res.ok ? (res.difMonto ?? null) : null,
      resultado: res.ok ? 'CRUZADO' : 'NO CRUZADO',
      metodo: res.met||'', motivo: res.motivo||''};
    if (res.ok) {
      fila.proc=res.match; fila.metodo=res.met; fila.estado=res.estado;
      fila.procEncontrada=cor.proc; fila.comOK=res.comOK; fila.sucOK=res.sucOK;
      fila.matchParcial=`Manual: ${cor.cupon}`; fila.correccionManual=CORREGIDAS[key];
      ok++;
    } else {
      fila.metodo='REVISION URGENTE'; fila.estado='REVISION URGENTE';
      fila.matchParcial=res.motivo; fila.correccionManual=CORREGIDAS[key]; fail++;
    }
  }
  return {ok, fail};
}

function limpiarFiltrosCor() {
  ['cor-flt-suc','cor-flt-tar','cor-flt-proc','cor-flt-estado'].forEach(id => {
    const el = document.getElementById(id); if (el) el.value = '';
  });
  const s = document.getElementById('cor-flt-search'); if (s) s.value = '';
  renderTablaCorrecciones();
}

function renderTablaCorrecciones() {
  // ── Construir lista completa de filas para los filtros ──────────
  const allEntries = Object.entries(CORREGIDAS).map(([key, cor]) => {
    const idx = RESULTADO.findIndex(r => _skyKey(r.sky) === key);
    return { key, cor, idx, fila: idx >= 0 ? RESULTADO[idx] : null };
  }).filter(e => e.fila);

  // ── Poblar selects de filtro (suc y tarjeta) ────────────────────
  const sucSel   = document.getElementById('cor-flt-suc');
  const tarSel   = document.getElementById('cor-flt-tar');
  const curSuc   = sucSel?.value || '';
  const curTar   = tarSel?.value || '';
  if (sucSel) {
    const sucs = [...new Set(allEntries.map(e => e.fila.sky.suc))].sort((a,b)=>+a-+b);
    sucSel.innerHTML = `<option value="">Todas las suc.</option>` +
      sucs.map(s=>`<option value="${s}" ${s===curSuc?'selected':''}>${s}</option>`).join('');
  }
  if (tarSel) {
    const tars = [...new Set(allEntries.map(e => e.fila.sky.tarjeta).filter(Boolean))].sort();
    tarSel.innerHTML = `<option value="">Todas las tarjetas</option>` +
      tars.map(t=>`<option value="${t}" ${t===curTar?'selected':''}>${t}</option>`).join('');
  }

  // ── Leer filtros activos ────────────────────────────────────────
  const fSuc    = document.getElementById('cor-flt-suc')?.value    || '';
  const fTar    = document.getElementById('cor-flt-tar')?.value    || '';
  const fProc   = document.getElementById('cor-flt-proc')?.value   || '';
  const fEstado = document.getElementById('cor-flt-estado')?.value || '';
  const fSearch = (document.getElementById('cor-flt-search')?.value || '').toLowerCase().trim();

  // ── Aplicar filtros ─────────────────────────────────────────────
  const filtered = allEntries.filter(({ fila, cor }) => {
    const s = fila.sky;
    if (fSuc    && s.suc !== fSuc)              return false;
    if (fTar    && s.tarjeta !== fTar)          return false;
    if (fProc   && cor.proc !== fProc)          return false;
    if (fEstado && cor.resultado !== fEstado)   return false;
    if (fSearch) {
      const hay = [s.vendedor||'', s.cupon, String(s.asiento||''), cor.cupon||'', s.tarjeta]
        .join(' ').toLowerCase();
      if (!hay.includes(fSearch)) return false;
    }
    return true;
  });

  const entries = filtered.map(e => [e.key, e.cor]);   // formato [key, cor][]

  const tbl=document.getElementById('tbl-cor');
  const cnt=document.getElementById('cnt-cor');
  const stats=document.getElementById('cor-stats');
  const btnR=document.getElementById('btn-reproc');
  const fStats = document.getElementById('cor-filter-stats');
  if (cnt) cnt.textContent=allEntries.length.toLocaleString();
  if (fStats) fStats.textContent = filtered.length < allEntries.length
    ? `Mostrando ${filtered.length} de ${allEntries.length}` : '';
  if (btnR) btnR.disabled=entries.length===0||(!_FIS_NORM.length&&!_GP_NORM.length);
  if (!tbl) return;
  const cruzados=entries.filter(([,c])=>c.resultado==='CRUZADO').length;
  const noCruz=entries.filter(([,c])=>c.resultado==='NO CRUZADO').length;
  const pend=entries.filter(([,c])=>!c.resultado||c.resultado==='PENDIENTE').length;
  if (stats) stats.innerHTML=`<b>${entries.length}</b> correcciones · `+
    `<b style="color:var(--grn)">${cruzados} cruzadas</b> · `+
    `<b style="color:var(--red)">${noCruz} sin match</b>`+
    (pend?` · <b style="color:var(--yel)">${pend} pendientes</b>`:'');
  if (!entries.length) {
    tbl.querySelector('thead').innerHTML='';
    tbl.querySelector('tbody').innerHTML=
      `<tr><td colspan="${HDR_COR.length}" style="padding:30px;text-align:center;color:var(--m2);font-size:10px">`+
      `Usá la pestaña "Revisión manual" para corregir operaciones Sin Match.</td></tr>`;
    return;
  }
  tbl.querySelector('thead').innerHTML=`<tr>${HDR_COR.map(h=>`<th>${h}</th>`).join('')}</tr>`;
  tbl.querySelector('tbody').innerHTML=entries.map(([key,cor])=>{
    const idx=RESULTADO.findIndex(r=>_skyKey(r.sky)===key);
    const fila=idx>=0?RESULTADO[idx]:null; if (!fila) return '';
    const s=fila.sky, p=fila.proc;
    const res=cor.resultado;
    const bc=res==='CRUZADO'?'st st-ok':res==='NO CRUZADO'?'st st-mal':'st st-ver';
    const bl=res==='CRUZADO'?'✓ CRUZADO':res==='NO CRUZADO'?'✗ SIN MATCH':'⟳ PENDIENTE';
    const rc=res==='CRUZADO'?'row-ok':res==='NO CRUZADO'?'row-mal':'row-ver';
    const det=res==='CRUZADO'?(cor.metodo||fila.metodo||''):(cor.motivo||fila.matchParcial||'');
    return `<tr class="${rc}">
      <td><span class="${bc}">${bl}</span></td>
      <td><span class="met">${cor.metodo||fila.metodo||'—'}</span></td>
      <td>${s.fecha}</td><td>${s.suc}</td>
      <td style="max-width:130px;overflow:hidden;text-overflow:ellipsis">${s.vendedor||'—'}</td>
      <td>${s.tarjeta}</td><td>${s.plan}</td>
      <td class="num">$${Math.abs(s.monto).toLocaleString('es-AR',{minimumFractionDigits:2})}</td>
      <td class="num" style="color:${cor.montoReal != null && cor.montoReal !== Math.abs(s.monto) ? 'var(--yel)' : 'var(--grn)'}">
        ${cor.montoReal != null
          ? `$${parseFloat(cor.montoReal).toLocaleString('es-AR',{minimumFractionDigits:2})}${cor.montoReal !== Math.abs(s.monto) ? ' ⚠' : ' ✓'}`
          : '<span style="color:var(--m2)">= facturado</span>'}</td>
      <td class="num" style="color:var(--m1)">${p ? '$'+Math.abs(p.monto||0).toLocaleString('es-AR',{minimumFractionDigits:2}) : '—'}</td>
      <td class="num ${cor.difMonto>500?'num-pos':cor.difMonto===0?'':'num-pos'}">${cor.difMonto!=null?(cor.difMonto===0?'<span style="color:var(--grn)">✓ Igual</span>':'$'+cor.difMonto.toLocaleString('es-AR',{maximumFractionDigits:0})):'—'}</td>
      <td class="num" style="font-family:var(--mono)">${s.cupon}</td>
      <td><span class="st ${cor.proc==='FISERV'?'st-fis':'st-gp'}">${cor.proc}</span></td>
      <td class="num" style="font-family:var(--mono);color:var(--cyn)">${cor.cupon}</td>
      <td>${p?.tarjeta||'—'}</td>
      <td style="font-size:9px">${p?.comFis||'—'}</td>
      <td>${fila.comOK||'—'}</td>
      <td style="font-size:9px;color:var(--m1);max-width:180px;overflow:hidden;text-overflow:ellipsis" title="${det}">${det}</td>
      <td><button onclick="eliminarCorreccion(${idx})" title="Eliminar corrección y volver a Sin Match"
        style="background:none;border:1px solid var(--b2);border-radius:3px;color:var(--m2);
          font-size:10px;cursor:pointer;padding:2px 7px;transition:all .15s"
        onmouseover="this.style.borderColor='var(--red)';this.style.color='var(--red)'"
        onmouseout="this.style.borderColor='var(--b2)';this.style.color='var(--m2)'">×</button></td>
    </tr>`;
  }).join('');
}

async function reprocesarTodasCorrecciones() {
  const btn=document.getElementById('btn-reproc');
  if (btn) {btn.disabled=true; btn.textContent='↻ Procesando...';}
  await new Promise(r=>setTimeout(r,30));
  const {ok,fail}=reprocesarCorrecciones();
  renderTablas(); updateCounts(); renderTablaCorrecciones(); renderFilas(); scheduleAutoSave();
  if (btn) {
    btn.textContent=`✓ ${ok} cruzadas · ${fail} sin match`;
    setTimeout(()=>{btn.textContent='↺ Re-procesar todas'; btn.disabled=false;}, 3000);
  }
}

// ══════════════════════════════════════════════════════════════════
// RE-CRUCE MANUAL — busca cupón ingresado en FISERV o GETPOS
// ══════════════════════════════════════════════════════════════════
// Opciones de método disponibles por procesadora
const METODOS_FISERV = [
  { val:'F1', label:'Ticket + Monto (recomendado)' },
  { val:'F2', label:'Auth + Monto' },
  { val:'F3', label:'Ticket + Fecha + Monto' },
  { val:'F4', label:'Auth + Fecha + Monto' },
  { val:'F5', label:'Solo Ticket (sin monto)' },
  { val:'F6', label:'Solo Auth (sin monto)' },
  { val:'F7', label:'Ticket parcial + Monto' },
  { val:'F8', label:'Auth parcial + Monto' },
];
const METODOS_GETPOS = [
  { val:'G1', label:'CodAut + Monto (recomendado)' },
  { val:'G2', label:'NroCupón + Monto' },
  { val:'G3', label:'CodAut + Fecha + Monto' },
  { val:'G4', label:'Solo CodAut (sin monto)' },
  { val:'G5', label:'CodAut parcial + Monto' },
];

function recruzarFila(skyIdx, cuponManual, proc, metodo, montoRealOverride) {
  const fila=RESULTADO[skyIdx]; if (!fila) return {ok:false, motivo:'Fila no encontrada'};
  const s=fila.sky, sn=s.suc, fn=s.fecha;
  // Usar montoReal si fue corregido manualmente, sino usar el monto de Skylab
  const mn = montoRealOverride != null ? normMonto(montoRealOverride) : s.montoN;
  const cuNorm=norm(cuponManual);


  const armarResultado = (match, met, pool) => {
    if (!match) return null;
    const difMonto = Math.abs(Math.abs(s.monto) - Math.abs(match.monto || 0));
    if (proc==='FISERV') {
      const comOK=validarComercio(s,match);
      const sucOK=match.suc?(match.suc===sn?'OK':'DIFERENTE'):'SIN DATO';
      let estado;
      if (fila.procEsperada==='GETPOS') estado='MAL FACTURADO: debe ser GETPOS';
      else if (comOK==='DIFERENTE') estado='COM. ERRADO (manual)';
      else if (comOK==='OK (equiv.)') estado='OK (equiv.) (manual)';
      else estado='OK (manual)';
      return {ok:true, match, met, comOK, sucOK, estado, difMonto};
    } else {
      const sucOK=match.suc?(match.suc===sn?'OK':'DIFERENTE'):'SIN DATO';
      const estado=fila.procEsperada==='FISERV'?'MAL FACTURADO: debe ser FISERV':'OK (manual)';
      return {ok:true, match, met, comOK:'—', sucOK, estado, difMonto};
    }
  };

  if (proc==='FISERV') {
    const pool=_FIS_NORM;
    const byMonto=pool.filter(r=>r.montoN===mn);
    let match, met, res;

    // Aplicar solo el método seleccionado
    switch(metodo||'F1') {
      case 'F1': match=byMonto.find(r=>r.ticket===cuNorm); met='Ticket+Monto'; break;
      case 'F2': match=byMonto.find(r=>r.aut===cuNorm);    met='Auth+Monto';   break;
      case 'F3': match=pool.find(r=>r.ticket===cuNorm&&r.fecha===fn&&r.montoN===mn); met='Ticket+Fecha+Monto'; break;
      case 'F4': match=pool.find(r=>r.aut===cuNorm&&r.fecha===fn&&r.montoN===mn);    met='Auth+Fecha+Monto';   break;
      case 'F5': match=pool.find(r=>r.ticket===cuNorm); met='Solo Ticket'; break;
      case 'F6': match=pool.find(r=>r.aut===cuNorm);    met='Solo Auth';   break;
      case 'F7': match=byMonto.find(r=>cuNorm.length>=3&&(r.ticket.includes(cuNorm)||cuNorm.includes(r.ticket))); met='Ticket parcial+Monto'; break;
      case 'F8': match=byMonto.find(r=>cuNorm.length>=3&&(r.aut.includes(cuNorm)||cuNorm.includes(r.aut)));       met='Auth parcial+Monto';   break;
      default:   match=byMonto.find(r=>r.ticket===cuNorm); met='Ticket+Monto';
    }

    res = armarResultado(match, `MC: ${met}`, pool);
    if (res) return res;
    return {ok:false, motivo:`No encontrado con método "${met}" — cupón="${cuponManual}" · monto SKY=$${Math.abs(s.monto).toLocaleString('es-AR')}`};

  } else {
    const pool=_GP_NORM;
    const byMonto=pool.filter(r=>r.montoN===mn);
    let match, met;

    switch(metodo||'G1') {
      case 'G1': match=byMonto.find(r=>r.aut===cuNorm);   met='CodAut+Monto';       break;
      case 'G2': match=byMonto.find(r=>r.cupon===cuNorm); met='NroCupón+Monto';     break;
      case 'G3': match=pool.find(r=>r.aut===cuNorm&&r.fecha===fn&&r.montoN===mn);   met='CodAut+Fecha+Monto'; break;
      case 'G4': match=pool.find(r=>r.aut===cuNorm);      met='Solo CodAut';        break;
      case 'G5': match=byMonto.find(r=>cuNorm.length>=3&&(r.aut.includes(cuNorm)||cuNorm.includes(r.aut))); met='CodAut parcial+Monto'; break;
      default:   match=byMonto.find(r=>r.aut===cuNorm);   met='CodAut+Monto';
    }

    const res = armarResultado(match, `MC: ${met}`, pool);
    if (res) return res;
    return {ok:false, motivo:`No encontrado con método "${met}" — cupón="${cuponManual}" · monto SKY=$${Math.abs(s.monto).toLocaleString('es-AR')}`};
  }
}

// ══════════════════════════════════════════════════════════════════
// TABLA ESPECIALIZADA COM.ERRADO — muestra Método y Nro.Comercio SKY
// ══════════════════════════════════════════════════════════════════
const HDR_COM = [
  'Estado','Diferencia tasa','Método cruce','Proc.Esp.','Proc.Real',
  'Suc.Sky','Tarjeta Sky','Plan','Fecha','Monto',
  'Cupón','Com.SKY','Com.FIS','¿Coincide?',
  'Tarjeta Proc.','Cuotas','Cód.Auth.','Suc.Proc.','Acciones'
];


// Marcar operación COM.ERRADO como sin dif o con dif en tasa
function marcarComErrado(skyIdx, marca) {
  skyIdx = parseInt(skyIdx);
  if (COM_ERRADO_MARCAS[skyIdx] === marca) {
    delete COM_ERRADO_MARCAS[skyIdx]; // toggle off
  } else {
    COM_ERRADO_MARCAS[skyIdx] = marca;
  }
  // Re-renderizar solo la tabla COM.ERRADO
  renderTablaCom(RESULTADO.filter(r => r.estado === 'COM. ERRADO'));
  // Recalcular diferencias siempre (para que esté listo cuando el usuario cambia de módulo)
  if (typeof renderModuloDiferencias === 'function') {
    renderModuloDiferencias();
  }
  scheduleAutoSave();
}

function renderTablaCom(filas) {
  const t = document.getElementById('tbl-com'); if (!t) return;
  t.querySelector('thead').innerHTML =
    `<tr>${HDR_COM.map(h=>`<th>${h}</th>`).join('')}</tr>`;
  t.querySelector('tbody').innerHTML = filas.slice(0,3000).map(r => {
    const s=r.sky, p=r.proc;
    const monto=`<span class="num ${s.monto<0?'num-neg':''}">${fmtARS(s.monto)}</span>`;
    const comSky  = s.nroCom || '—';
    const comFis  = p?.comFis || '—';
    // Resaltar: COM SKY en cian, COM FIS en rojo si difiere
    const comSkyBadge = `<span style="color:var(--cyn);font-weight:600">${comSky}</span>`;
    const comFisBadge = r.comOK==='DIFERENTE'
      ? `<span style="color:var(--red);font-weight:600">${comFis}</span>`
      : `<span style="color:var(--grn)">${comFis}</span>`;
    const marca = COM_ERRADO_MARCAS[s.idx] || '';
    const badgeDif = marca === 'CON_DIF'
      ? `<span class="st" style="background:rgba(248,113,113,.15);color:var(--red);border:1px solid rgba(248,113,113,.35);font-weight:700">CON DIF.</span>`
      : marca === 'SIN_DIF'
      ? `<span class="st" style="background:rgba(52,211,153,.1);color:var(--grn);border:1px solid rgba(52,211,153,.2)">SIN DIF.</span>`
      : `<span style="font-size:8px;color:var(--m2)">Sin marcar</span>`;
    const btnsMarca = `
      <button onclick="marcarComErrado(${s.idx},'SIN_DIF')" title="Sin diferencia de tasa — solo error de código comercio"
        style="padding:3px 7px;border-radius:4px;font-size:8px;font-weight:600;cursor:pointer;
          border:1px solid ${marca==='SIN_DIF'?'var(--grn)':'var(--b2)'};
          background:${marca==='SIN_DIF'?'rgba(52,211,153,.15)':'transparent'};
          color:${marca==='SIN_DIF'?'var(--grn)':'var(--m2)'}">
        ✓ Sin dif.
      </button>
      <button onclick="marcarComErrado(${s.idx},'CON_DIF')" title="Con diferencia de tasa — pasa al módulo de Diferencias"
        style="padding:3px 7px;border-radius:4px;font-size:8px;font-weight:600;cursor:pointer;margin-left:4px;
          border:1px solid ${marca==='CON_DIF'?'var(--red)':'var(--b2)'};
          background:${marca==='CON_DIF'?'rgba(248,113,113,.15)':'transparent'};
          color:${marca==='CON_DIF'?'var(--red)':'var(--m2)'}">
        ✗ Con dif.
      </button>`;
    const cells = [
      estadoBadge(r.estado),
      badgeDif,
      `<span class="met" style="color:var(--acc)">${r.metodo??''}</span>`,
      r.procEsperada??'', r.procEncontrada??'',
      s.suc, s.tarjeta, s.plan, s.fecha, monto,
      s.cupon, comSkyBadge, comFisBadge,
      r.comOK??'', p?.tarjeta??'', p?.cuotas??'', p?.aut??'', p?.suc??'',
      btnsMarca,
    ];
    const rowBg = marca==='CON_DIF' ? 'row-mal' : marca==='SIN_DIF' ? 'row-ok' : 'row-com';
    return `<tr class="${rowBg}">${cells.map(c=>`<td>${c}</td>`).join('')}</tr>`;
  }).join('');
}

// ══════════════════════════════════════════════════════════════════
// TABLAS REVISIÓN URGENTE Y REFACTURADO
// ══════════════════════════════════════════════════════════════════
const HDR_URGENTE = [
  'Fecha','Suc','Vendedor','Tarjeta','Tipo','Plan','Monto',
  'Cupón SKY','Lote','Nro.Comercio','Proc. esperada','Detalle','Acciones'
];

function renderTablaUrgente() {
  const filas = RESULTADO.filter(r => r.estado === 'REVISION URGENTE');
  const tbl   = document.getElementById('tbl-urg');
  const stats = document.getElementById('urg-stats');
  const cnt   = document.getElementById('cnt-urg');
  if (cnt)   cnt.textContent   = filas.length;
  if (stats) stats.innerHTML   =
    `<b>${filas.length}</b> operaciones · venta en Skylab <b style="color:var(--red)">sin pago encontrado</b>`;
  if (!tbl) return;
  if (!filas.length) {
    tbl.querySelector('thead').innerHTML = '';
    tbl.querySelector('tbody').innerHTML =
      `<tr><td colspan="${HDR_URGENTE.length}" style="padding:30px;text-align:center;color:var(--m2);font-size:10px">
        No hay operaciones marcadas como Revisión Urgente.</td></tr>`;
    return;
  }
  tbl.querySelector('thead').innerHTML =
    `<tr>${HDR_URGENTE.map(h=>`<th>${h}</th>`).join('')}</tr>`;
  tbl.querySelector('tbody').innerHTML = filas.map(r => {
    const s = r.sky;
    const cor  = CORREGIDAS[_skyKey(s)] || {};
    const tipo = _tipoOp(r);
    const montoColor = s.monto < 0 ? 'var(--red)' : 'var(--org)';
    const montoFmt   = s.monto < 0 ? `−${fmtARS(Math.abs(s.monto))}` : fmtARS(s.monto);
    const detalle    = r.matchParcial || (cor.cupon ? `Cupón probado: ${cor.cupon}` : '—');
    return `<tr class="row-urgente">
      <td>${s.fecha}</td>
      <td>${s.suc}</td>
      <td style="max-width:140px;overflow:hidden;text-overflow:ellipsis">${s.vendedor||'—'}</td>
      <td>${s.tarjeta}</td>
      <td><span style="font-size:8px;padding:2px 7px;border-radius:3px;
        color:${tipo.color};border:1px solid ${tipo.color}55;background:${tipo.color}18">${tipo.label}</span></td>
      <td>${s.plan}</td>
      <td class="num" style="color:${montoColor};font-weight:600">${montoFmt}</td>
      <td class="num" style="font-family:var(--mono)">${s.cupon}</td>
      <td>${s.lote}</td>
      <td style="font-size:9px">${s.nroCom||'—'}</td>
      <td><span class="st ${s.esGETPos?'st-gp':'st-fis'}">${r.procEsperada||'—'}</span></td>
      <td style="font-size:9px;color:var(--m2);max-width:200px;overflow:hidden;text-overflow:ellipsis"
        title="${detalle}">${detalle}</td>
      <td style="white-space:nowrap;display:flex;gap:4px;align-items:center">
        <button onclick="resetearFila(${s.idx})" title="Quitar y devolver a Sin Match"
          style="background:none;border:1px solid var(--b2);border-radius:3px;color:var(--m2);
            font-size:11px;cursor:pointer;width:22px;height:22px;display:flex;align-items:center;justify-content:center"
          onmouseover="this.style.borderColor='var(--red)';this.style.color='var(--red)'"
          onmouseout="this.style.borderColor='var(--b2)';this.style.color='var(--m2)'">×</button>
        <button onclick="showMod('revision',document.getElementById('mbtn-revision'))"
          title="Ir a revisión manual para corregir"
          style="background:none;border:1px solid var(--b2);border-radius:3px;color:var(--m2);
            font-size:9px;cursor:pointer;padding:2px 7px"
          onmouseover="this.style.borderColor='var(--acc)';this.style.color='var(--acc)'"
          onmouseout="this.style.borderColor='var(--b2)';this.style.color='var(--m2)'">Reintentar</button>
      </td>
    </tr>`;
  }).join('');
}

const HDR_REFACT = [
  'Fecha','Suc','Vendedor','Tarjeta','Tipo','Plan','Monto',
  'Cupón SKY','Lote','Nro.Comercio','Proc. esperada','Observación','Acciones'
];

function renderTablaRefacturado() {
  const filas = RESULTADO.filter(r => r.estado === 'REFACTURADO');
  const tbl   = document.getElementById('tbl-ref');
  const stats = document.getElementById('ref-stats');
  const cnt   = document.getElementById('cnt-ref');
  if (cnt)   cnt.textContent   = filas.length;
  if (stats) stats.innerHTML   =
    `<b>${filas.length}</b> operaciones marcadas como <b style="color:var(--vio)">refacturadas</b>`;
  if (!tbl) return;
  if (!filas.length) {
    tbl.querySelector('thead').innerHTML = '';
    tbl.querySelector('tbody').innerHTML =
      `<tr><td colspan="${HDR_REFACT.length}" style="padding:30px;text-align:center;color:var(--m2);font-size:10px">
        No hay operaciones marcadas como Refacturado.</td></tr>`;
    return;
  }
  tbl.querySelector('thead').innerHTML =
    `<tr>${HDR_REFACT.map(h=>`<th>${h}</th>`).join('')}</tr>`;
  tbl.querySelector('tbody').innerHTML = filas.map(r => {
    const s = r.sky;
    const tipo = _tipoOp(r);
    const montoColor = s.monto < 0 ? 'var(--red)' : 'var(--vio)';
    const montoFmt   = s.monto < 0 ? `−${fmtARS(Math.abs(s.monto))}` : fmtARS(s.monto);
    return `<tr class="row-refact">
      <td>${s.fecha}</td>
      <td>${s.suc}</td>
      <td style="max-width:140px;overflow:hidden;text-overflow:ellipsis">${s.vendedor||'—'}</td>
      <td>${s.tarjeta}</td>
      <td><span style="font-size:8px;padding:2px 7px;border-radius:3px;
        color:${tipo.color};border:1px solid ${tipo.color}55;background:${tipo.color}18">${tipo.label}</span></td>
      <td>${s.plan}</td>
      <td class="num" style="color:${montoColor};font-weight:600">${montoFmt}</td>
      <td class="num" style="font-family:var(--mono)">${s.cupon}</td>
      <td>${s.lote}</td>
      <td style="font-size:9px">${s.nroCom||'—'}</td>
      <td><span class="st ${s.esGETPos?'st-gp':'st-fis'}">${r.procEsperada||'—'}</span></td>
      <td style="font-size:9px;color:var(--m2)">${r.matchParcial||'Marcado manualmente como refacturado'}</td>
      <td style="white-space:nowrap;display:flex;gap:4px;align-items:center">
        <button onclick="resetearFila(${s.idx})" title="Quitar y devolver a Sin Match"
          style="background:none;border:1px solid var(--b2);border-radius:3px;color:var(--m2);
            font-size:11px;cursor:pointer;width:22px;height:22px;display:flex;align-items:center;justify-content:center"
          onmouseover="this.style.borderColor='var(--red)';this.style.color='var(--red)'"
          onmouseout="this.style.borderColor='var(--b2)';this.style.color='var(--m2)'">×</button>
        <button onclick="showMod('revision',document.getElementById('mbtn-revision'))"
          title="Ir a revisión manual"
          style="background:none;border:1px solid var(--b2);border-radius:3px;color:var(--m2);
            font-size:9px;cursor:pointer;padding:2px 7px"
          onmouseover="this.style.borderColor='var(--acc)';this.style.color='var(--acc)'"
          onmouseout="this.style.borderColor='var(--b2)';this.style.color='var(--m2)'">Reintentar</button>
      </td>
    </tr>`;
  }).join('');
}

function exportarUrgente() {
  const filas = RESULTADO.filter(r => r.estado === 'REVISION URGENTE');
  const HDR = ['Nro.Asiento','Fecha','Sucursal','Vendedor','Tarjeta','Plan','Monto',
               'Cupón SKY','Lote','Nro.Comercio','Proc.Esperada','Detalle'];
  const data = filas.map(r => {
    const s = r.sky, cor = CORREGIDAS[_skyKey(s)] || {};
    return [s.asiento||'', s.fecha, s.suc, s.vendedor||'', s.tarjeta, s.plan,
      Math.abs(s.monto), s.cupon, s.lote, s.nroCom||'',
      r.procEsperada||'', cor.cupon?`Cupón probado: ${cor.cupon}`:r.matchParcial||''];
  });
  _exportXlsx([HDR,...data], 'Rev. Urgente', `RevisionUrgente_${hoy()}.xlsx`);
}

function exportarRefacturado() {
  const filas = RESULTADO.filter(r => r.estado === 'REFACTURADO');
  const HDR = ['Nro.Asiento','Fecha','Sucursal','Vendedor','Tarjeta','Plan','Monto',
               'Cupón SKY','Lote','Nro.Comercio','Proc.Esperada'];
  const data = filas.map(r => {
    const s = r.sky;
    return [s.asiento||'', s.fecha, s.suc, s.vendedor||'', s.tarjeta, s.plan,
      Math.abs(s.monto), s.cupon, s.lote, s.nroCom||'', r.procEsperada||''];
  });
  _exportXlsx([HDR,...data], 'Refacturado', `Refacturado_${hoy()}.xlsx`);
}

// ── Marcar fila por estado SIN re-renderizar (para imports en batch) ──
function _marcarRevisionSilent(idx, estado) {
  idx = parseInt(idx);
  const fila = RESULTADO[idx]; if (!fila) return false;
  fila.estado         = estado;
  fila.metodo         = estado;
  fila.procEncontrada = estado === 'REFACTURADO' ? 'N/A' : fila.procEsperada;
  fila.matchParcial   = estado === 'REFACTURADO'
    ? 'Importado como refacturado'
    : estado === 'ANULACION SIN COBRO'
    ? 'Importado como anulación sin cobro'
    : 'Importado como revisión urgente';
  if (estado === 'ANULACION SIN COBRO') fila.esAnulSinCobro = 'SI';
  const key = _skyKey(fila.sky);
  const cor = CORREGIDAS[key] || {};
  CORREGIDAS[key] = { ...cor, cupon: cor.cupon||'—', proc: cor.proc||fila.procEsperada,
    resultado: estado, metodo: estado };
  return true;
}

// ── Import genérico por estado (lee Nro.Asiento del Excel exportado) ─
function _importarPorEstado(input, estado) {
  const file = input.files[0]; if (!file) return;
  const r = new FileReader();
  r.onload = e => {
    try {
      const wb   = XLSX.read(e.target.result, { type:'array', cellDates:true });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: false });
      let ok = 0, nf = 0;
      rows.forEach(row => {
        const asiento = String(row['Nro.Asiento'] ?? row['Nro. Asiento'] ?? row['Asiento'] ?? '').trim();
        if (!asiento || asiento === 'null') return;
        const key = `A:${asiento}`;
        const idx = RESULTADO.findIndex(r => _skyKey(r.sky) === key);
        if (idx < 0) { nf++; return; }
        if (_marcarRevisionSilent(idx, estado)) ok++;
      });
      // Render único al final
      renderTablas(); updateCounts(); renderFilas();
      renderTablaCorrecciones(); renderTablaUrgente(); renderTablaRefacturado();
      scheduleAutoSave();
      const label = estado === 'REFACTURADO' ? 'refacturadas' : 'marcadas como urgente';
      typeof _showToast === 'function'
        ? _showToast(`✓ ${ok} ${label}${nf ? ` · ${nf} no encontradas` : ''}`)
        : alert(`✓ ${ok} ${label}` + (nf ? `\n⚠ ${nf} no encontradas` : ''));
    } catch(err) {
      alert('Error al leer el archivo: ' + err.message);
      console.error(err);
    }
  };
  r.readAsArrayBuffer(file);
  input.value = '';
}

function importarUrgente(input)     { _importarPorEstado(input, 'REVISION URGENTE'); }
function importarRefacturado(input) { _importarPorEstado(input, 'REFACTURADO'); }

// ══════════════════════════════════════════════════════════════════
// FISERV Y GETPOS SIN CRUCE
// ══════════════════════════════════════════════════════════════════
function renderNoCruzadasFis() {
  // Recalcular dinámicamente para reflejar correcciones manuales aplicadas
  const usedFis = new Set(RESULTADO.filter(r => r.proc && r.procEncontrada === 'FISERV').map(r => r.proc));
  const rows = _FIS_NORM.length
    ? _FIS_NORM.filter(r => !usedFis.has(r))
    : (window._FIS_NO_CRUZADAS || []);
  const tbl  = document.getElementById('tbl-dif-fis');
  const stats= document.getElementById('dif-fis-stats');
  const cnt  = document.getElementById('cnt-dif-fis');
  if (cnt)   cnt.textContent   = rows.length;
  if (stats) stats.innerHTML   =
    `<b>${rows.length}</b> operaciones FISERV <b style="color:var(--red)">sin match en Skylab</b>`;
  if (!tbl) return;
  const strip = document.getElementById('tab-strip-dif');
  if (strip) strip.style.display = 'flex';
  if (!rows.length) {
    tbl.querySelector('thead').innerHTML = '';
    tbl.querySelector('tbody').innerHTML =
      `<tr><td colspan="${HDR_FIS_NC.length}" style="padding:30px;text-align:center;color:var(--m2);font-size:10px">
        Todas las operaciones FISERV tienen correspondencia en Skylab.</td></tr>`;
    return;
  }
  tbl.querySelector('thead').innerHTML = `<tr>${HDR_FIS_NC.map(h=>`<th>${h}</th>`).join('')}</tr>`;
  tbl.querySelector('tbody').innerHTML = rows.map(r => `<tr>
    <td>${r.fecha}</td><td style="font-size:9px">${r.equipo||'—'}</td><td>${r.suc||'—'}</td>
    <td>${r.tarjeta||'—'}</td><td class="num">${r.cuotas||'—'}</td>
    <td class="num" style="color:var(--org)">${fmtARS(r.monto)}</td>
    <td>${r.lote||'—'}</td><td>${r.ticket||'—'}</td><td style="font-size:9px">${r.aut||'—'}</td>
    <td style="font-size:9px">${r.comFis||'—'}</td>
  </tr>`).join('');
}

function renderNoCruzadasGp() {
  // Recalcular dinámicamente para reflejar correcciones manuales aplicadas
  const usedGp = new Set(RESULTADO.filter(r => r.proc && r.procEncontrada === 'GETPOS').map(r => r.proc));
  const rows = _GP_NORM.length
    ? _GP_NORM.filter(r => !usedGp.has(r))
    : (window._GP_NO_CRUZADAS || []);
  const tbl  = document.getElementById('tbl-dif-gp');
  const stats= document.getElementById('dif-gp-stats');
  const cnt  = document.getElementById('cnt-dif-gp');
  if (cnt)   cnt.textContent   = rows.length;
  if (stats) stats.innerHTML   =
    `<b>${rows.length}</b> operaciones GETPOS <b style="color:var(--red)">sin match en Skylab</b>`;
  if (!tbl) return;
  if (!rows.length) {
    tbl.querySelector('thead').innerHTML = '';
    tbl.querySelector('tbody').innerHTML =
      `<tr><td colspan="${HDR_GP_NC.length}" style="padding:30px;text-align:center;color:var(--m2);font-size:10px">
        Todas las operaciones GETPOS tienen correspondencia en Skylab.</td></tr>`;
    return;
  }
  tbl.querySelector('thead').innerHTML = `<tr>${HDR_GP_NC.map(h=>`<th>${h}</th>`).join('')}</tr>`;
  tbl.querySelector('tbody').innerHTML = rows.map(r => `<tr>
    <td>${r.fecha}</td>
    <td style="max-width:160px;overflow:hidden;text-overflow:ellipsis">${r.nombre||'—'}</td>
    <td>${r.suc||'—'}</td><td>${r.marca||'—'}</td><td>${r.plan||'—'}</td>
    <td class="num" style="color:var(--org)">${fmtARS(r.monto)}</td>
    <td>${r.aut||'—'}</td><td>${r.cupon||'—'}</td>
  </tr>`).join('');
}

function exportarCorrecciones() {
  const entries = Object.entries(CORREGIDAS);
  if (!entries.length) { alert('No hay correcciones manuales para exportar.'); return; }

  // NOTA: la primera columna "Nro.Asiento" es la clave de reimportación.
  // Al importar este archivo, el sistema re-aplica las correcciones por asiento.
  const HDR = [
    'Nro.Asiento',   // ← clave de reimportación (no eliminar)
    // Skylab (referencia)
    'Fecha SKY','Sucursal','Vendedor','Tarjeta SKY','Plan','Cuotas SKY',
    'Monto SKY','Cupón SKY','Lote','Nro.Comercio SKY',
    // Corrección (las columnas que se reimportan)
    'Procesadora','Cupón ingresado','Método cruce','Estado cruce',
    // Procesadora
    'Monto Proc.','DIF $','Tarjeta Proc.','Cuotas Proc.',
    'Lote Proc.','Ticket Proc.','Com.FIS','Cód.Auth.','Suc.Proc.','Com.OK'
  ];

  const data = entries.map(([key, cor]) => {
    const idx  = RESULTADO.findIndex(r => _skyKey(r.sky) === key);
    const fila = idx >= 0 ? RESULTADO[idx] : null;
    if (!fila) return null;
    const s = fila.sky;
    const p = fila.proc;
    // Extraer asiento del key (A:12345 → 12345)
    const asiento = key.startsWith('A:') ? key.slice(2) : (s.asiento ?? key);
    return [
      asiento,
      // Skylab
      s.fecha, s.suc, s.vendedor ?? '', s.tarjeta, s.plan, s.cuotas ?? '',
      Math.abs(s.monto), s.cupon, s.lote, s.nroCom ?? '',
      // Corrección
      cor.proc ?? '', cor.cupon ?? '', cor.metodo ?? '', fila.estado ?? '',
      // Procesadora
      p ? Math.abs(p.monto || 0) : '',
      cor.difMonto != null ? cor.difMonto : '',
      p?.tarjeta ?? '', p?.cuotas ?? '',
      p?.lote ?? '', p?.ticket ?? '',
      p?.comFis ?? '', p?.aut ?? '', p?.suc ?? '', fila.comOK ?? '',
    ];
  }).filter(Boolean);

  _exportXlsx([HDR, ...data], 'Correcciones', `Correcciones_Manuales_${hoy()}.xlsx`);
}

// ── IMPORTAR correcciones desde Excel exportado previamente ─────────
function importarCorrecciones(input) {
  const file = input.files[0]; if (!file) return;
  const r = new FileReader();
  r.onload = e => {
    try {
      const wb   = XLSX.read(e.target.result, { type:'array', cellDates:true });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: false });
      if (!rows.length) { alert('El archivo no tiene datos.'); return; }

      let importadas = 0, noEncontradas = 0, sinAsiento = 0;
      rows.forEach(row => {
        const asientoRaw = String(row['Nro.Asiento'] ?? row['Nro. Asiento'] ?? '').trim();
        const cupon      = String(row['Cupón ingresado'] ?? row['Cupon ingresado'] ?? '').trim();
        const proc       = String(row['Procesadora'] ?? '').trim();
        const metodo     = String(row['Método cruce'] ?? row['Metodo cruce'] ?? '').trim();

        if (!cupon || !proc) return;
        if (!asientoRaw || asientoRaw === 'null') { sinAsiento++; return; }

        const key = `A:${asientoRaw}`;
        // Verificar si existe en el RESULTADO actual
        const existe = RESULTADO.some(r => _skyKey(r.sky) === key);
        if (!existe) { noEncontradas++; }

        // Guardar/sobreescribir la corrección (se aplicará al re-procesar)
        CORREGIDAS[key] = {
          cupon, proc, metodo,
          resultado: 'PENDIENTE',
          motivo:    'Importado desde Excel'
        };
        importadas++;
      });

      // Re-aplicar todas las correcciones con los datos actuales
      if (importadas > 0 && (_FIS_NORM.length || _GP_NORM.length)) {
        reprocesarCorrecciones();
        renderTodo();
        updateCounts();
      }
      scheduleAutoSave();

      const msg = [
        `✓ ${importadas} correcciones importadas`,
        noEncontradas ? `⚠ ${noEncontradas} sin fila en el RESULTADO actual (se guardan igual)` : '',
        sinAsiento    ? `ℹ ${sinAsiento} filas sin Nro.Asiento ignoradas` : '',
      ].filter(Boolean).join('\n');
      typeof _showToast === 'function'
        ? _showToast(`✓ ${importadas} correcciones importadas`)
        : alert(msg);
      if (noEncontradas || sinAsiento) console.warn('[IMPORT COR]', msg);

    } catch(err) {
      alert('Error leyendo el archivo: ' + err.message);
      console.error(err);
    }
  };
  r.readAsArrayBuffer(file);
  input.value = '';
}

function exportarNoCruzadasFis() {
  // Cruce inicial: estado al momento del cruce automático (sin correcciones manuales)
  const rows = window._FIS_NO_CRUZADAS || [];
  if (!rows.length) { alert('No hay datos del cruce inicial. Ejecutá el cruce primero.'); return; }
  const HDR = ['Fecha','Terminal','Sucursal','Tarjeta','Cuotas','Monto','Lote','Ticket','Autorización','Cód.Comercio'];
  const data = rows.map(r => [r.fecha, r.equipo||'', r.suc||'', r.tarjeta||'', r.cuotas||'',
    r.monto, r.lote||'', r.ticket||'', r.aut||'', r.comFis||'']);
  _exportXlsx([HDR,...data], 'FISERV Sin Cruce Inicial', `FISERV_SinCruce_Inicial_${hoy()}.xlsx`);
}

function exportarNoCruzadasFisConciliado() {
  // Cruce conciliado: descuenta las correcciones manuales aplicadas
  const usedFis = new Set(RESULTADO.filter(r => r.proc && r.procEncontrada === 'FISERV').map(r => r.proc));
  const rows = _FIS_NORM.length
    ? _FIS_NORM.filter(r => !usedFis.has(r))
    : (window._FIS_NO_CRUZADAS || []);
  if (!rows.length) { alert('No hay operaciones FISERV sin cruce.'); return; }
  const HDR = ['Fecha','Terminal','Sucursal','Tarjeta','Cuotas','Monto','Lote','Ticket','Autorización','Cód.Comercio'];
  const data = rows.map(r => [r.fecha, r.equipo||'', r.suc||'', r.tarjeta||'', r.cuotas||'',
    r.monto, r.lote||'', r.ticket||'', r.aut||'', r.comFis||'']);
  _exportXlsx([HDR,...data], 'FISERV Sin Cruce Conciliado', `FISERV_SinCruce_Conciliado_${hoy()}.xlsx`);
}

function exportarNoCruzadasGp() {
  // Cruce inicial: estado al momento del cruce automático (sin correcciones manuales)
  const rows = window._GP_NO_CRUZADAS || [];
  if (!rows.length) { alert('No hay datos del cruce inicial. Ejecutá el cruce primero.'); return; }
  const HDR = ['Fecha','Establecimiento','Sucursal','Marca','Plan','Monto','Cód.Aut.','Nro.Cupón'];
  const data = rows.map(r => [r.fecha, r.nombre||'', r.suc||'', r.marca||'', r.plan||'',
    r.monto, r.aut||'', r.cupon||'']);
  _exportXlsx([HDR,...data], 'GETPOS Sin Cruce Inicial', `GETPOS_SinCruce_Inicial_${hoy()}.xlsx`);
}

function exportarNoCruzadasGpConciliado() {
  // Cruce conciliado: descuenta las correcciones manuales aplicadas
  const usedGp = new Set(RESULTADO.filter(r => r.proc && r.procEncontrada === 'GETPOS').map(r => r.proc));
  const rows = _GP_NORM.length
    ? _GP_NORM.filter(r => !usedGp.has(r))
    : (window._GP_NO_CRUZADAS || []);
  if (!rows.length) { alert('No hay operaciones GETPOS sin cruce.'); return; }
  const HDR = ['Fecha','Establecimiento','Sucursal','Marca','Plan','Monto','Cód.Aut.','Nro.Cupón'];
  const data = rows.map(r => [r.fecha, r.nombre||'', r.suc||'', r.marca||'', r.plan||'',
    r.monto, r.aut||'', r.cupon||'']);
  _exportXlsx([HDR,...data], 'GETPOS Sin Cruce Conciliado', `GETPOS_SinCruce_Conciliado_${hoy()}.xlsx`);
}
