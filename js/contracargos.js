// ═══════════════════════════════════════════════════════════════════
// CONTRACARGOS.JS — Módulo 7: Análisis y seguimiento de contracargos
// FISERV: contracargos ya debitados (importes negativos)
// GETPOS: disputas en curso con estado y fecha de vencimiento
// ═══════════════════════════════════════════════════════════════════

let _CTR_FIS = [];   // contracargos FISERV parseados
let _CTR_GP  = [];   // contracargos GETPOS parseados

// Seguimiento manual: { [id]: { estado, notas, montoRecuperado, fechaRespuesta } }
// Se serializa en SESSION vía state.js
if (typeof CTR_SEGUIMIENTO === 'undefined') var CTR_SEGUIMIENTO = {};

const CTR_EST_LIST = ['PENDIENTE','EN GESTIÓN','RESPONDIDO','GANADO','PERDIDO','ACEPTADO'];

// ── Colores por estado ──────────────────────────────────────────────
function _ctrColor(estado) {
  switch (estado) {
    case 'PENDIENTE':   return { bg:'rgba(107,114,128,.12)', txt:'#9ca3af', brd:'rgba(107,114,128,.3)' };
    case 'EN GESTIÓN':  return { bg:'rgba(79,142,247,.12)',  txt:'#4f8ef7', brd:'rgba(79,142,247,.3)' };
    case 'RESPONDIDO':  return { bg:'rgba(56,189,248,.12)',  txt:'#38bdf8', brd:'rgba(56,189,248,.3)' };
    case 'GANADO':      return { bg:'rgba(52,211,153,.12)',  txt:'#34d399', brd:'rgba(52,211,153,.3)' };
    case 'PERDIDO':     return { bg:'rgba(248,113,113,.12)', txt:'#f87171', brd:'rgba(248,113,113,.3)' };
    case 'ACEPTADO':    return { bg:'rgba(251,146,60,.12)',  txt:'#fb923c', brd:'rgba(251,146,60,.3)' };
    default:            return { bg:'rgba(107,114,128,.1)',  txt:'#9ca3af', brd:'rgba(107,114,128,.2)' };
  }
}

function _ctrBadge(estado) {
  const c = _ctrColor(estado || 'PENDIENTE');
  return `<span style="font-size:8px;padding:2px 8px;border-radius:3px;font-weight:600;
    background:${c.bg};color:${c.txt};border:1px solid ${c.brd}">${estado || 'PENDIENTE'}</span>`;
}

// ── Días hasta vencimiento ──────────────────────────────────────────
function _diasHasta(fechaStr) {
  if (!fechaStr) return null;
  const hoy  = new Date(); hoy.setHours(0,0,0,0);
  const dest = new Date(fechaStr + 'T00:00:00');
  return Math.round((dest - hoy) / 86400000);
}

function _vencBadge(fechaStr) {
  if (!fechaStr) return '—';
  const d = _diasHasta(fechaStr);
  if (d === null) return fechaStr;
  let color, label;
  if (d < 0)      { color='#f87171'; label=`VENCIDO (${Math.abs(d)}d)`; }
  else if (d <= 3) { color='#f87171'; label=`⚠ ${d}d`; }
  else if (d <= 7) { color='#fb923c'; label=`⚠ ${d}d`; }
  else if (d <= 14){ color='#fbbf24'; label=`${d}d`; }
  else             { color='#34d399'; label=`${d}d`; }
  return `<span style="font-size:8px;padding:2px 6px;border-radius:3px;color:${color};
    border:1px solid ${color}55;background:${color}18" title="${fechaStr}">${label}</span>`;
}

// ── Limpiar texto con _x000D_ y "undefined - " ─────────────────────
function _cleanCtrText(s) {
  return String(s || '').replace(/_x000D_/g,'').replace(/^undefined\s*-\s*/i,'').trim();
}

// ── Extraer código de motivo (ej: "90 DESCONOCIMIENTO...") ──────────
function _parseMotivoFis(detalle) {
  const s = _cleanCtrText(detalle);
  const m = s.match(/^(\d+)\s+(.*)/s);
  return m ? { codigo: m[1], desc: m[2].trim() } : { codigo: '', desc: s };
}

// ── Fecha GP: "27/01/2026 12:00" → "2026-01-27" ────────────────────
function _parseFechaGP(s) {
  if (!s) return '';
  const m = String(s).match(/(\d{2})\/(\d{2})\/(\d{4})/);
  if (m) return `${m[3]}-${m[2]}-${m[1]}`;
  return String(s).slice(0,10);
}

// ════════════════════════════════════════════════════════════════════
// PARSERS
// ════════════════════════════════════════════════════════════════════
function parseContrarcargosFiserv(wb) {
  const ws   = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: false });
  if (!rows.length) { _CTR_FIS = []; return []; }

  const allKeys = Object.keys(rows[0]);
  console.log('[CTR-FIS] Columnas:', allKeys);

  const K = {
    fechaDeb:  _resolveKey(allKeys, 'Fecha Debitado', 'Fecha debitado'),
    fechaTx:   _resolveKey(allKeys, 'Fecha Operación', 'Fecha Operacion', 'Fecha Operacion'),
    fechaPres: _resolveKey(allKeys, 'Fecha Present', 'Fecha Presentación', 'Fecha Presentacion'),
    nroLiq:    _resolveKey(allKeys, 'Nro Liquidación', 'Nro Liquidacion', 'Nro. Liquidacion'),
    lote:      _resolveKey(allKeys, 'Nro Lote', 'Lote'),
    cupon:     _resolveKey(allKeys, 'Nrto Cupón', 'Nro Cupón', 'Nro Cupon', 'Cupon', 'Nrto Cupon'),
    tarjeta:   _resolveKey(allKeys, 'Tarjeta'),
    terminal:  _resolveKey(allKeys, 'Nro Terminal', 'Terminal'),
    importe:   _resolveKey(allKeys, 'Importe Bruto', 'Importe'),
    cuotas:    _resolveKey(allKeys, 'Cuotas'),
    nroTarj:   _resolveKey(allKeys, 'Nro Tarjeta'),
    aut:       _resolveKey(allKeys, 'Nro Autoriz', 'Nro Autorización', 'Nro Autorizacion', 'Autorizacion'),
    nroCom:    _resolveKey(allKeys, 'Nro Comercio'),
    motivo:    _resolveKey(allKeys, 'Motivo'),
    detalle:   _resolveKey(allKeys, 'Detalle'),
    cuit:      _resolveKey(allKeys, 'Cuit', 'CUIT'),
    arancel:   _resolveKey(allKeys, 'Arancel'),
    ivaArancel:_resolveKey(allKeys, 'Iva Arancel', 'IVA Arancel'),
  };
  console.log('[CTR-FIS] Mapeo:', K);

  const g = (r, k) => (k && r[k] !== undefined) ? r[k] : null;

  _CTR_FIS = rows.map((r, i) => {
    const impRaw   = String(g(r, K.importe) || '0').replace(/\./g,'').replace(',','.');
    const importe  = Math.abs(parseFloat(impRaw) || 0);
    const lot      = norm(g(r, K.lote));
    const cup      = norm(g(r, K.cupon));
    const motDet   = _parseMotivoFis(g(r, K.detalle));

    return {
      id:           `FIS_${lot}_${cup}_${i}`,
      proc:         'FISERV',
      fechaDebitado:normFecha(g(r, K.fechaDeb)),
      fechaTx:      normFecha(g(r, K.fechaTx)),
      fechaPresent: normFecha(g(r, K.fechaPres)),
      fechaVenc:    null,
      nroLiq:       String(g(r, K.nroLiq) || '').trim(),
      lote:         lot,
      cupon:        cup,
      tarjeta:      String(g(r, K.tarjeta) || '').trim(),
      terminal:     String(g(r, K.terminal) || '').trim(),
      importe,
      cuotas:       parseInt(g(r, K.cuotas)) || 1,
      nroTarjeta:   String(g(r, K.nroTarj) || '').trim(),
      aut:          norm(g(r, K.aut)),
      nroCom:       String(g(r, K.nroCom) || '').trim(),
      motivo:       String(g(r, K.motivo) || '').trim(),
      motivoCodigo: motDet.codigo,
      motivoDesc:   motDet.desc,
      cuit:         String(g(r, K.cuit) || '').trim(),
      arancel:      Math.abs(parseFloat(g(r, K.arancel)) || 0),
      ivaArancel:   Math.abs(parseFloat(g(r, K.ivaArancel)) || 0),
      matchIdx:     null,
    };
  }).filter(r => r.importe > 0);

  _cruzarContracargos();
  console.log('[CTR-FIS] Parseados:', _CTR_FIS.length);
  return _CTR_FIS;
}

function parseContracargosGetpos(wb) {
  const ws   = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: false });
  if (!rows.length) { _CTR_GP = []; return []; }

  const allKeys = Object.keys(rows[0]);
  console.log('[CTR-GP] Columnas:', allKeys);

  const K = {
    fechaTx:    _resolveKey(allKeys, 'Fecha de la transacción', 'Fecha de la transaccion', 'Fecha'),
    codigoDisp: _resolveKey(allKeys, 'Código de la disputa', 'Codigo de la disputa', 'Codigo disputa', 'ID disputa'),
    afiliacion: _resolveKey(allKeys, 'Afiliación', 'Afiliacion'),
    comercio:   _resolveKey(allKeys, 'Comercio'),
    marca:      _resolveKey(allKeys, 'Marca'),
    arn:        _resolveKey(allKeys, 'ARN'),
    aut:        _resolveKey(allKeys, 'Código de autorización', 'Codigo de autorizacion', 'Cod. Autorización', 'Cod Autorización', 'Autorización'),
    monto:      _resolveKey(allKeys, 'Monto de la disputa', 'Monto disputa', 'Monto'),
    moneda:     _resolveKey(allKeys, 'Moneda'),
    estatus:    _resolveKey(allKeys, 'Estatus', 'Estado', 'Status'),
    motivo:     _resolveKey(allKeys, 'Motivo'),
    fechaVenc:  _resolveKey(allKeys, 'Fecha de vencimiento del comercio', 'Fecha vencimiento', 'Vencimiento'),
    emisor:     _resolveKey(allKeys, 'Emisor'),
    procesador: _resolveKey(allKeys, 'Procesador'),
    cuotaCuota: _resolveKey(allKeys, 'Cuota a Cuota'),
    retiro:     _resolveKey(allKeys, 'Retiro de efectivo'),
  };
  console.log('[CTR-GP] Mapeo:', K);

  const g = (r, k) => (k && r[k] !== undefined) ? r[k] : null;

  _CTR_GP = rows.map((r, i) => {
    const montoRaw = String(g(r, K.monto) || '0').replace(/\./g,'').replace(',','.');
    const monto    = Math.abs(parseFloat(montoRaw) || 0);

    return {
      id:          String(g(r, K.codigoDisp) || `GP_${i}`).trim(),
      proc:        'GETPOS',
      fechaTx:     _parseFechaGP(g(r, K.fechaTx)),
      fechaDebitado: _parseFechaGP(g(r, K.fechaTx)),
      fechaVenc:   _parseFechaGP(g(r, K.fechaVenc)),
      arn:         String(g(r, K.arn) || '').trim(),
      aut:         norm(g(r, K.aut)),
      tarjeta:     String(g(r, K.marca) || '').trim(),
      comercio:    String(g(r, K.comercio) || '').trim(),
      afiliacion:  String(g(r, K.afiliacion) || '').trim(),
      importe:     monto,
      moneda:      String(g(r, K.moneda) || 'ARS').trim(),
      estatusProc: String(g(r, K.estatus) || '').trim(),
      motivo:      _cleanCtrText(g(r, K.motivo)),
      emisor:      String(g(r, K.emisor) || '').trim(),
      procesador:  String(g(r, K.procesador) || '').trim(),
      cuotaCuota:  String(g(r, K.cuotaCuota) || '').toUpperCase().includes('S'),
      retiro:      String(g(r, K.retiro) || '').toUpperCase().includes('S'),
      matchIdx:    null,
      lote:        '',
      cupon:       '',
      motivoCodigo:'',
      motivoDesc:  _cleanCtrText(g(r, K.motivo)),
    };
  }).filter(r => r.importe > 0 || r.id);

  _cruzarContracargos();
  console.log('[CTR-GP] Parseados:', _CTR_GP.length);
  return _CTR_GP;
}

// ── Cruzar con RESULTADO (si está disponible) ───────────────────────
function _cruzarContracargos() {
  if (!RESULTADO || !RESULTADO.length) return;

  // FISERV: cruzar por cupón (ticket) + lote
  _CTR_FIS.forEach(ctr => {
    if (ctr.matchIdx !== null) return;
    const idx = RESULTADO.findIndex(r =>
      r.proc && r.procEncontrada === 'FISERV' &&
      (norm(r.proc.ticket) === ctr.cupon || (ctr.aut && norm(r.proc.aut) === ctr.aut))
    );
    ctr.matchIdx = idx >= 0 ? idx : null;
  });

  // GETPOS: cruzar por auth
  _CTR_GP.forEach(ctr => {
    if (ctr.matchIdx !== null) return;
    const idx = RESULTADO.findIndex(r =>
      r.proc && r.procEncontrada === 'GETPOS' &&
      ctr.aut && norm(r.proc.aut) === ctr.aut
    );
    ctr.matchIdx = idx >= 0 ? idx : null;
  });
}

// ── Lista unificada de todos los contracargos ───────────────────────
function _allCtr() {
  return [..._CTR_FIS, ..._CTR_GP];
}

// ── Obtener/setear seguimiento de un contracargo ────────────────────
function _getCtrSeg(id) {
  return CTR_SEGUIMIENTO[id] || { estado:'PENDIENTE', notas:'', montoRecuperado:0, fechaRespuesta:'' };
}

function guardarSeguimientoCtr(id, campo, valor) {
  if (!CTR_SEGUIMIENTO[id]) CTR_SEGUIMIENTO[id] = { estado:'PENDIENTE', notas:'', montoRecuperado:0, fechaRespuesta:'' };
  CTR_SEGUIMIENTO[id][campo] = valor;
  scheduleAutoSave();
}

// ════════════════════════════════════════════════════════════════════
// RENDER MÓDULO
// ════════════════════════════════════════════════════════════════════
function renderModuloContracargos() {
  const hasFis = _CTR_FIS.length > 0;
  const hasGP  = _CTR_GP.length > 0;
  const all    = _allCtr();

  if (!hasFis && !hasGP) {
    document.getElementById('mod-ctr').innerHTML =
      `<div style="display:flex;flex-direction:column;align-items:center;justify-content:center;
        height:100%;gap:14px;color:var(--m2);text-align:center;padding:40px">
        <div style="font-size:42px;opacity:.15">🛡</div>
        <div style="font-family:var(--head);font-size:16px;font-weight:700;color:var(--txt);opacity:.4">Módulo de Contracargos</div>
        <p style="font-size:10px;max-width:400px;line-height:1.8">
          Cargá los archivos de contracargos de <b style="color:var(--acc)">FISERV</b> y/o <b style="color:var(--grn)">GETPOS</b>
          desde el panel izquierdo.
        </p>
      </div>`;
    return;
  }

  _cruzarContracargos();

  // KPIs
  const totalMonto  = all.reduce((s, c) => s + c.importe, 0);
  const gps         = all.filter(c => c.proc === 'GETPOS');
  const urgentes    = gps.filter(c => { const d = _diasHasta(c.fechaVenc); return d !== null && d <= 7 && d >= 0; });
  const vencidos    = gps.filter(c => { const d = _diasHasta(c.fechaVenc); return d !== null && d < 0; });
  const ganados     = all.filter(c => _getCtrSeg(c.id).estado === 'GANADO');
  const perdidos    = all.filter(c => _getCtrSeg(c.id).estado === 'PERDIDO');
  const montoRec    = ganados.reduce((s, c) => s + (_getCtrSeg(c.id).montoRecuperado || c.importe), 0);

  const panel = document.getElementById('mod-ctr');
  panel.innerHTML = `
    <div style="display:flex;flex-direction:column;flex:1;min-height:0;overflow:hidden">

      <!-- KPIs -->
      <div style="display:grid;grid-template-columns:repeat(6,1fr);gap:8px;
        padding:12px 18px;background:var(--s1);border-bottom:1px solid var(--b1);flex-shrink:0">
        <div class="dif-kpi">
          <div class="dif-kpi-lbl">Total contracargos</div>
          <div class="dif-kpi-val cyn">${all.length}</div>
          <div style="font-size:8px;color:var(--m2);margin-top:2px">${hasFis?_CTR_FIS.length+' FIS':''} ${hasGP?_CTR_GP.length+' GP':''}</div>
        </div>
        <div class="dif-kpi">
          <div class="dif-kpi-lbl">Monto en riesgo</div>
          <div class="dif-kpi-val red">${fmtARS(totalMonto)}</div>
        </div>
        <div class="dif-kpi" style="border-color:rgba(248,113,113,.3)">
          <div class="dif-kpi-lbl">⚠ Vencen ≤7 días</div>
          <div class="dif-kpi-val" style="color:${urgentes.length>0?'var(--red)':'var(--grn)'}">${urgentes.length}</div>
          <div style="font-size:8px;color:var(--m2);margin-top:2px">${vencidos.length} ya vencidos</div>
        </div>
        <div class="dif-kpi">
          <div class="dif-kpi-lbl">Cruzadas con SKY</div>
          <div class="dif-kpi-val cyn">${all.filter(c=>c.matchIdx!==null).length}</div>
          <div style="font-size:8px;color:var(--m2);margin-top:2px">de ${all.length}</div>
        </div>
        <div class="dif-kpi" style="border-color:rgba(52,211,153,.25)">
          <div class="dif-kpi-lbl">Ganadas</div>
          <div class="dif-kpi-val grn">${ganados.length}</div>
          <div style="font-size:8px;color:var(--grn);margin-top:2px">${fmtARS(montoRec)}</div>
        </div>
        <div class="dif-kpi" style="border-color:rgba(248,113,113,.25)">
          <div class="dif-kpi-lbl">Perdidas</div>
          <div class="dif-kpi-val red">${perdidos.length}</div>
          <div style="font-size:8px;color:var(--red);margin-top:2px">${fmtARS(perdidos.reduce((s,c)=>s+c.importe,0))}</div>
        </div>
      </div>

      <!-- Tab strip -->
      <div class="tab-strip" id="tab-strip-ctr">
        <button class="tb active" onclick="showTab('ctr-seguimiento','tab-strip-ctr',this)">
          📋 Seguimiento <span class="cnt" id="cnt-ctr-all">${all.length}</span>
        </button>
        ${hasFis ? `<button class="tb" onclick="showTab('ctr-fis','tab-strip-ctr',this)">
          FISERV <span class="cnt" id="cnt-ctr-fis">${_CTR_FIS.length}</span>
        </button>` : ''}
        ${hasGP ? `<button class="tb" onclick="showTab('ctr-gp','tab-strip-ctr',this)">
          GETPOS <span class="cnt" id="cnt-ctr-gp">${_CTR_GP.length}</span>
        </button>` : ''}
        ${urgentes.length+vencidos.length > 0 ? `<button class="tb" style="color:var(--red)"
          onclick="showTab('ctr-alertas','tab-strip-ctr',this)">
          ⚠ Alertas <span class="cnt" style="background:rgba(248,113,113,.15);color:var(--red)"
            id="cnt-ctr-urg">${urgentes.length+vencidos.length}</span>
        </button>` : ''}
      </div>

      <!-- Tab: Seguimiento unificado -->
      <div class="tab-body active" id="ctr-seguimiento" style="flex-direction:column;flex:1;min-height:0">
        <div class="cor-hdr-bar" style="border-left:3px solid var(--acc)">
          <span class="cor-hdr-title">Seguimiento de contracargos</span>
          <span class="cor-stats">${all.length} total · ${fmtARS(totalMonto)} en riesgo</span>
          <div style="margin-left:auto;display:flex;gap:6px">
            <select class="filter-sel" id="ctr-flt-est" onchange="renderTablaCtrSeguimiento()" style="font-size:8px">
              <option value="">Todos los estados</option>
              ${CTR_EST_LIST.map(e=>`<option value="${e}">${e}</option>`).join('')}
            </select>
            <select class="filter-sel" id="ctr-flt-proc" onchange="renderTablaCtrSeguimiento()" style="font-size:8px">
              <option value="">Ambas proc.</option>
              <option value="FISERV">FISERV</option>
              <option value="GETPOS">GETPOS</option>
            </select>
            <button class="dl-btn" style="background:#14532d;color:#86efac"
              onclick="exportarContracargos()">⬇ Exportar</button>
          </div>
        </div>
        <div class="tbl-wrap" id="ctr-seg-wrap">
          <table id="tbl-ctr-seg"><thead></thead><tbody></tbody></table>
        </div>
      </div>

      <!-- Tab: FISERV -->
      <div class="tab-body" id="ctr-fis" style="flex-direction:column;flex:1;min-height:0">
        <div class="cor-hdr-bar" style="border-left:3px solid var(--acc)">
          <span class="cor-hdr-title" style="color:var(--acc)">Contracargos FISERV</span>
          <span class="cor-stats">${_CTR_FIS.length} registros · ${fmtARS(_CTR_FIS.reduce((s,c)=>s+c.importe,0))}</span>
        </div>
        <div class="tbl-wrap"><table id="tbl-ctr-fis"><thead></thead><tbody></tbody></table></div>
      </div>

      <!-- Tab: GETPOS -->
      <div class="tab-body" id="ctr-gp" style="flex-direction:column;flex:1;min-height:0">
        <div class="cor-hdr-bar" style="border-left:3px solid var(--grn)">
          <span class="cor-hdr-title" style="color:var(--grn)">Contracargos GETPOS</span>
          <span class="cor-stats">${_CTR_GP.length} disputas · ${fmtARS(_CTR_GP.reduce((s,c)=>s+c.importe,0))}</span>
        </div>
        <div class="tbl-wrap"><table id="tbl-ctr-gp"><thead></thead><tbody></tbody></table></div>
      </div>

      <!-- Tab: Alertas -->
      <div class="tab-body" id="ctr-alertas" style="flex-direction:column;flex:1;min-height:0">
        <div class="cor-hdr-bar" style="border-left:3px solid var(--red)">
          <span class="cor-hdr-title" style="color:var(--red)">⚠ Alertas de vencimiento</span>
          <span class="cor-stats">${urgentes.length} próximos · ${vencidos.length} vencidos</span>
        </div>
        <div class="tbl-wrap"><table id="tbl-ctr-alertas"><thead></thead><tbody></tbody></table></div>
      </div>

    </div>`;

  // Renderizar las tablas
  renderTablaCtrSeguimiento();
  if (hasFis) _renderTablaCtrFis();
  if (hasGP)  _renderTablaCtrGp();
  if (urgentes.length + vencidos.length > 0) _renderTablaCtrAlertas(urgentes, vencidos);
}

// ── Tabla de seguimiento unificada ──────────────────────────────────
function renderTablaCtrSeguimiento() {
  const tbl    = document.getElementById('tbl-ctr-seg'); if (!tbl) return;
  const fltEst  = document.getElementById('ctr-flt-est')?.value  || '';
  const fltProc = document.getElementById('ctr-flt-proc')?.value || '';

  let filas = _allCtr();
  if (fltEst)  filas = filas.filter(c => _getCtrSeg(c.id).estado === fltEst);
  if (fltProc) filas = filas.filter(c => c.proc === fltProc);

  const HDR = ['Estado','Proc.','Fecha Tx','Vencimiento','Tarjeta','Motivo','Monto',
               'Match SKY','Monto recuperado','Notas','Fecha respuesta',''];
  tbl.querySelector('thead').innerHTML = `<tr>${HDR.map(h=>`<th>${h}</th>`).join('')}</tr>`;

  if (!filas.length) {
    tbl.querySelector('tbody').innerHTML =
      `<tr><td colspan="${HDR.length}" style="padding:20px;text-align:center;color:var(--m2);font-size:10px">
        Sin contracargos para los filtros seleccionados.</td></tr>`;
    return;
  }

  tbl.querySelector('tbody').innerHTML = filas.map(c => {
    const seg   = _getCtrSeg(c.id);
    const match = c.matchIdx !== null ? RESULTADO[c.matchIdx] : null;
    const matchCell = match
      ? `<span style="font-size:8px;color:var(--grn)">✓ ${match.sky.fecha} · ${match.sky.suc} · ${fmtARS(Math.abs(match.sky.monto))}</span>`
      : `<span style="font-size:8px;color:var(--m2)">—</span>`;

    return `<tr>
      <td>
        ${_ctrBadge(seg.estado)}
        <select onchange="guardarSeguimientoCtr('${c.id}','estado',this.value);renderModuloContracargos()"
          style="margin-top:4px;display:block;width:100%;background:var(--s3);border:1px solid var(--b2);
            border-radius:3px;color:var(--txt);font-family:var(--mono);font-size:8px;padding:2px 4px">
          ${CTR_EST_LIST.map(e=>`<option value="${e}" ${seg.estado===e?'selected':''}>${e}</option>`).join('')}
        </select>
      </td>
      <td><span class="st ${c.proc==='FISERV'?'st-fis':'st-gp'}">${c.proc}</span></td>
      <td style="font-size:9px">${c.fechaTx}</td>
      <td>${c.proc==='GETPOS' ? _vencBadge(c.fechaVenc) : '<span style="color:var(--m2);font-size:9px">N/A</span>'}</td>
      <td style="font-size:9px">${c.tarjeta}</td>
      <td style="font-size:9px;max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap"
        title="${c.motivoDesc}">${c.motivoCodigo ? `<b style="color:var(--yel)">${c.motivoCodigo}</b> ` : ''}${c.motivoDesc || c.motivo}</td>
      <td class="num" style="color:var(--red);font-weight:600">${fmtARS(c.importe)}</td>
      <td>${matchCell}</td>
      <td>
        <input type="number" step="0.01" placeholder="0"
          value="${seg.montoRecuperado || ''}"
          onchange="guardarSeguimientoCtr('${c.id}','montoRecuperado',parseFloat(this.value)||0)"
          style="width:90px;background:var(--s3);border:1px solid var(--b2);border-radius:3px;
            color:var(--grn);font-family:var(--mono);font-size:9px;padding:3px 6px;text-align:right">
      </td>
      <td>
        <input type="text" placeholder="Notas..."
          value="${(seg.notas||'').replace(/"/g,'&quot;')}"
          onchange="guardarSeguimientoCtr('${c.id}','notas',this.value)"
          style="width:140px;background:var(--s3);border:1px solid var(--b2);border-radius:3px;
            color:var(--txt);font-family:var(--mono);font-size:9px;padding:3px 6px">
      </td>
      <td>
        <input type="date"
          value="${seg.fechaRespuesta||''}"
          onchange="guardarSeguimientoCtr('${c.id}','fechaRespuesta',this.value)"
          style="background:var(--s3);border:1px solid var(--b2);border-radius:3px;
            color:var(--txt);font-family:var(--mono);font-size:9px;padding:3px 6px">
      </td>
      <td>
        <button onclick="resetearCtr('${c.id}')" title="Resetear seguimiento"
          style="background:none;border:1px solid var(--b2);border-radius:3px;color:var(--m2);
            font-size:9px;cursor:pointer;padding:2px 7px"
          onmouseover="this.style.borderColor='var(--red)';this.style.color='var(--red)'"
          onmouseout="this.style.borderColor='var(--b2)';this.style.color='var(--m2)'">↺</button>
      </td>
    </tr>`;
  }).join('');
}

function resetearCtr(id) {
  delete CTR_SEGUIMIENTO[id];
  renderTablaCtrSeguimiento();
  scheduleAutoSave();
}

// ── Tabla FISERV ─────────────────────────────────────────────────────
function _renderTablaCtrFis() {
  const tbl = document.getElementById('tbl-ctr-fis'); if (!tbl) return;
  const HDR = ['Fecha Tx','Fecha Debitado','Terminal','Tarjeta','Lote','Cupón','Aut.',
               'Cuotas','Importe','Motivo cód.','Detalle motivo','Arancel','Match SKY','Estado'];
  tbl.querySelector('thead').innerHTML = `<tr>${HDR.map(h=>`<th>${h}</th>`).join('')}</tr>`;
  tbl.querySelector('tbody').innerHTML = _CTR_FIS.map(c => {
    const match = c.matchIdx !== null ? RESULTADO[c.matchIdx] : null;
    const seg   = _getCtrSeg(c.id);
    return `<tr>
      <td>${c.fechaTx}</td>
      <td>${c.fechaDebitado}</td>
      <td style="font-size:9px">${c.terminal}</td>
      <td>${c.tarjeta}</td>
      <td class="num">${c.lote}</td>
      <td class="num">${c.cupon}</td>
      <td class="num" style="font-size:9px">${c.aut || '—'}</td>
      <td class="num">${c.cuotas}</td>
      <td class="num" style="color:var(--red);font-weight:600">${fmtARS(c.importe)}</td>
      <td><span style="font-size:8px;padding:1px 5px;border-radius:3px;background:rgba(251,191,36,.15);
        color:var(--yel);border:1px solid rgba(251,191,36,.3)">${c.motivoCodigo || '—'}</span></td>
      <td style="font-size:9px;max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap"
        title="${c.motivoDesc}">${c.motivoDesc}</td>
      <td class="num" style="font-size:9px;color:var(--m2)">${c.arancel ? fmtARS(c.arancel) : '—'}</td>
      <td style="font-size:9px">${match
        ? `<span style="color:var(--grn)">✓ Suc.${match.sky.suc} ${match.sky.fecha}</span>`
        : '<span style="color:var(--m2)">—</span>'}</td>
      <td>${_ctrBadge(seg.estado)}</td>
    </tr>`;
  }).join('');
}

// ── Tabla GETPOS ─────────────────────────────────────────────────────
function _renderTablaCtrGp() {
  const tbl = document.getElementById('tbl-ctr-gp'); if (!tbl) return;
  const HDR = ['Fecha Tx','Vencimiento','Código disputa','Marca','Auth.','ARN',
               'Monto','Estatus proc.','Motivo','Emisor','Match SKY','Estado'];
  tbl.querySelector('thead').innerHTML = `<tr>${HDR.map(h=>`<th>${h}</th>`).join('')}</tr>`;
  tbl.querySelector('tbody').innerHTML = _CTR_GP.map(c => {
    const match = c.matchIdx !== null ? RESULTADO[c.matchIdx] : null;
    const seg   = _getCtrSeg(c.id);
    const d     = _diasHasta(c.fechaVenc);
    const rowCls= d !== null && d < 0 ? 'row-mal' : d !== null && d <= 7 ? 'row-com' : '';
    return `<tr class="${rowCls}">
      <td>${c.fechaTx}</td>
      <td>${_vencBadge(c.fechaVenc)}</td>
      <td style="font-size:9px;font-family:var(--mono)">${c.id}</td>
      <td>${c.tarjeta}</td>
      <td class="num" style="font-family:var(--mono)">${c.aut || '—'}</td>
      <td style="font-size:8px;max-width:160px;overflow:hidden;text-overflow:ellipsis" title="${c.arn}">${c.arn}</td>
      <td class="num" style="color:var(--red);font-weight:600">${fmtARS(c.importe)}</td>
      <td style="font-size:9px">${c.estatusProc}</td>
      <td style="font-size:9px;max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap"
        title="${c.motivo}">${c.motivo}</td>
      <td style="font-size:9px;max-width:140px;overflow:hidden;text-overflow:ellipsis">${c.emisor}</td>
      <td style="font-size:9px">${match
        ? `<span style="color:var(--grn)">✓ Suc.${match.sky.suc} ${match.sky.fecha}</span>`
        : '<span style="color:var(--m2)">—</span>'}</td>
      <td>${_ctrBadge(seg.estado)}</td>
    </tr>`;
  }).join('');
}

// ── Tabla Alertas ─────────────────────────────────────────────────────
function _renderTablaCtrAlertas(urgentes, vencidos) {
  const tbl  = document.getElementById('tbl-ctr-alertas'); if (!tbl) return;
  const all  = [...vencidos, ...urgentes].sort((a,b) => {
    const da = _diasHasta(a.fechaVenc) ?? 999;
    const db = _diasHasta(b.fechaVenc) ?? 999;
    return da - db;
  });
  const HDR  = ['Urgencia','Código disputa','Fecha Tx','Vencimiento','Marca','Auth.','Monto','Motivo','Emisor','Estado seguimiento'];
  tbl.querySelector('thead').innerHTML = `<tr>${HDR.map(h=>`<th>${h}</th>`).join('')}</tr>`;
  tbl.querySelector('tbody').innerHTML = all.map(c => {
    const d   = _diasHasta(c.fechaVenc);
    const seg = _getCtrSeg(c.id);
    const urgLabel = d < 0
      ? `<span style="color:var(--red);font-weight:700">VENCIDO ${Math.abs(d)}d</span>`
      : d === 0
      ? `<span style="color:var(--red);font-weight:700">HOY</span>`
      : `<span style="color:${d<=3?'var(--red)':'var(--org)'};font-weight:700">QUEDAN ${d}d</span>`;
    return `<tr class="${d < 0 ? 'row-mal' : 'row-com'}">
      <td>${urgLabel}</td>
      <td style="font-size:9px;font-family:var(--mono)">${c.id}</td>
      <td>${c.fechaTx}</td>
      <td>${_vencBadge(c.fechaVenc)}</td>
      <td>${c.tarjeta}</td>
      <td class="num" style="font-family:var(--mono)">${c.aut || '—'}</td>
      <td class="num" style="color:var(--red);font-weight:600">${fmtARS(c.importe)}</td>
      <td style="font-size:9px;max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${c.motivo}</td>
      <td style="font-size:9px;max-width:140px;overflow:hidden;text-overflow:ellipsis">${c.emisor}</td>
      <td>${_ctrBadge(seg.estado)}</td>
    </tr>`;
  }).join('');
}

// ════════════════════════════════════════════════════════════════════
// EXPORT
// ════════════════════════════════════════════════════════════════════
function exportarContracargos() {
  const all = _allCtr();
  if (!all.length) { alert('No hay contracargos cargados.'); return; }

  const HDR = ['Procesadora','Fecha Tx','Fecha Debitado','Vencimiento (GP)',
               'Código/ID','ARN','Auth','Tarjeta','Cupón','Lote','Terminal',
               'Importe','Moneda','Motivo Código','Motivo Descripción',
               'Emisor','Estatus Procesadora',
               'Estado seguimiento','Monto recuperado','Notas','Fecha respuesta',
               'Match SKY — Fecha','Match SKY — Suc','Match SKY — Monto'];

  const data = all.map(c => {
    const seg   = _getCtrSeg(c.id);
    const match = c.matchIdx !== null ? RESULTADO[c.matchIdx] : null;
    return [
      c.proc,
      c.fechaTx, c.fechaDebitado, c.fechaVenc || '',
      c.id, c.arn || '', c.aut || '', c.tarjeta, c.cupon || '', c.lote || '', c.terminal || '',
      c.importe, c.moneda || 'ARS',
      c.motivoCodigo || '', c.motivoDesc || c.motivo || '',
      c.emisor || '', c.estatusProc || '',
      seg.estado, seg.montoRecuperado || 0, seg.notas || '', seg.fechaRespuesta || '',
      match?.sky?.fecha || '', match?.sky?.suc || '',
      match ? Math.abs(match.sky.monto) : '',
    ];
  });

  _exportXlsx([HDR, ...data], 'Contracargos', `Contracargos_${hoy()}.xlsx`);
}
