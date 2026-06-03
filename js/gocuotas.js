// ═══════════════════════════════════════════════════════════════════
// GOCUOTAS.JS — Módulo 10: Conciliación de Go Cuotas
//
// Flujo de cruce:
//   SkyGC.Cupon  ↔  GoCuotas.Número de Orden   → ¿cobrado?
//   SkyGC.Id     ↔  Ventas.Comprobante (n° final) → artículo + IMEI
//   IMEI requerido cuando plan contiene "CELULAR" o artículo es teléfono
// ═══════════════════════════════════════════════════════════════════

let _GOC_SKY     = [];   // Skylab Go Cuotas parseado
let _GOC_PAGOS   = [];   // Go Cuotas CSV parseado
let _GOC_VENTAS  = [];   // Ventas XLSX parseado
let _GOC_RESULT  = [];   // resultado del cruce

// ── Helpers de formato ──────────────────────────────────────────────
function _gFmt(v)  { return typeof fmtARS === 'function' ? fmtARS(v) : '$'+v; }
function _gNorm(v) { return String(v||'').trim().replace(/^0+/, '') || '0'; }
function _gParseMonto(s) {
  const clean = String(s||'0').replace(/\$/g,'').replace(/\s/g,'')
    .replace(/\./g,'').replace(',','.').trim();
  return Math.abs(parseFloat(clean)||0);
}

// ── Detectar si el artículo/plan requiere IMEI ───────────────────────
function _gRequiereImei(plan, articulo) {
  const p = String(plan||'').toUpperCase();
  const a = String(articulo||'').toUpperCase();
  if (p.includes('CELULAR') || p.includes('GOCELU') || p.includes('GO CEL')) return true;
  if (a.includes('IPHONE') || a.includes('SAMSUNG') || a.includes('MOTOROLA') ||
      a.includes('XIAOMI') || a.includes('LG ') || a.includes('HONOR ') ||
      a.includes('NOKIA')  || a.includes('HUAWEI') || a.includes(' A0') ||
      a.includes(' S2') || a.includes('REDMI') || a.includes('REALME') ||
      a.includes('OPPO') || a.includes('VIVO ')) return true;
  return false;
}

// ── Detectar si un string es un IMEI válido (15 dígitos) ────────────
function _gEsImei(s) {
  return /^\d{15}$/.test(String(s||'').trim().replace(/[^0-9]/g,''));
}

// ════════════════════════════════════════════════════════════════════
// PARSERS
// ════════════════════════════════════════════════════════════════════

// ── Parser: Skylab Go Cuotas (XLSX — misma estructura que Skylab) ────
function parseGocSkylab(wb) {
  const ws   = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: false });
  if (!rows.length) { _GOC_SKY = []; return []; }

  const allKeys = Object.keys(rows[0]);
  console.log('[GOC-SKY] Columnas:', allKeys);

  // La columna Cupon del Skylab Go Cuotas contiene el Número de Orden de GoC
  const K = {
    asiento:  _resolveKey(allKeys, 'Nro.Asiento', 'Asiento'),
    tarjeta:  _resolveKey(allKeys, 'Tarjeta'),
    plan:     _resolveKey(allKeys, 'Plan'),
    id:       _resolveKey(allKeys, 'Id'),          // n° de comprobante → link con Ventas
    numero:   _resolveKey(allKeys, 'Numero'),      // fecha como serial Excel
    fecha:    _resolveKey(allKeys, 'Fec.de Vta.', 'Fec. de Vta.', 'Fecha'),
    lote:     _resolveKey(allKeys, 'Lote'),
    cupon:    _resolveKey(allKeys, 'Cupon', 'Cupón', 'Cupon Nro', 'Nro de Cupon'),
    importe:  _resolveKey(allKeys, 'Venta Bruta', 'Importe', 'Monto'),
    neto:     _resolveKey(allKeys, 'Neto a Cobrar', 'Neto'),
    sucursal: _resolveKey(allKeys, 'Sucursal'),
    vendedor: _resolveKey(allKeys, 'Vendedor'),
    fecPago:  _resolveKey(allKeys, 'Fec.de Pago', 'Fec. de Pago', 'Fecha de Pago'),
    opFact:   _resolveKey(allKeys, 'Op.Fact.', 'Op. Fact.'),
  };
  console.log('[GOC-SKY] Mapeo:', K);

  const g = (r,k) => (k && r[k] !== undefined) ? r[k] : null;

  _GOC_SKY = rows.map((r,i) => {
    const cuponRaw = String(g(r, K.cupon) || g(r, K.lote) || '').trim();
    const idRaw    = String(g(r, K.id)    || '').trim();
    const impRaw   = g(r, K.importe);
    // Fecha: puede venir como serial Excel (número) o string dd/mm/yyyy
    const fechaRaw = g(r, K.fecha) || g(r, K.numero);
    let fecha = '';
    if (fechaRaw) {
      const n = parseFloat(String(fechaRaw));
      if (!isNaN(n) && n > 40000 && n < 60000) {
        // Excel date serial → fecha real
        const d = new Date((n - 25569) * 86400000);
        fecha = d.toISOString().slice(0,10);
      } else {
        fecha = normFecha(fechaRaw);
      }
    }

    return {
      idx:      i,
      asiento:  String(g(r, K.asiento) || '').trim(),
      plan:     String(g(r, K.plan)    || '').trim(),
      id:       idRaw,                          // → link con Ventas.Comprobante
      cupon:    _gNorm(cuponRaw),               // → link con GoCuotas.NúmeroOrden
      fecha,
      importe:  _gParseMonto(impRaw),
      neto:     _gParseMonto(g(r, K.neto)),
      sucursal: String(g(r, K.sucursal) || '').trim(),
      vendedor: String(g(r, K.vendedor) || '').trim(),
      fecPago:  normFecha(g(r, K.fecPago)),
      opFact:   String(g(r, K.opFact)  || '').trim(),
      // campos que se rellenan al cruzar
      _pagos:   null,   // { orden, importe, fechaPago }
      _venta:   null,   // { descripcion, imei }
      _estado:  'PENDIENTE',
      _imeiOk:  null,
    };
  }).filter(r => r.importe > 0 || r.cupon !== '0');

  console.log('[GOC-SKY] Skylab Go Cuotas parseados:', _GOC_SKY.length);
  return _GOC_SKY;
}

// ── Parser: Go Cuotas CSV (separado por ;) ───────────────────────────
function parseGocPagos(file) {
  return new Promise((res, rej) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const text  = e.target.result;
        const lines = text.split(/\r?\n/).filter(l => l.trim());
        if (lines.length < 2) { _GOC_PAGOS = []; res([]); return; }

        const sep    = ';';
        const hdrRaw = lines[0].split(sep).map(h => h.trim());
        console.log('[GOC-PAGOS] Columnas CSV:', hdrRaw);

        // Mapeo flexible de columnas
        const fi = (names) => {
          for (const n of names) {
            const idx = hdrRaw.findIndex(h => _normColName(h) === _normColName(n));
            if (idx >= 0) return idx;
          }
          return -1;
        };

        const IDX = {
          tipo:        fi(['Descripcion','Tipo']),
          fechaOrigen: fi(['Fecha Origen','Fecha']),
          fechaPago:   fi(['Fecha Pago']),
          orden:       fi(['Número de Orden','Numero de Orden','NroOrden','Orden']),
          nombre:      fi(['ApellidoNombre','Nombre']),
          cuotas:      fi(['Cuotas']),
          importe:     fi(['Importe',' Importe ']),
          totalCobrar: fi(['Total a cobrar',' Total a cobrar ']),
          sucId:       fi(['Sucursal ID']),
          sucNombre:   fi(['Sucursal Nombre']),
          refExt:      fi(['Referencia Externa']),
        };
        console.log('[GOC-PAGOS] Índices columnas:', IDX);

        _GOC_PAGOS = lines.slice(1).map((line, i) => {
          const cols = line.split(sep).map(c => c.trim());
          const g = (key) => IDX[key] >= 0 ? cols[IDX[key]] : '';
          const tipo = g('tipo').toLowerCase();
          if (tipo.includes('total') || tipo.includes('subtotal')) return null; // fila de totales
          return {
            idx:         i,
            tipo:        g('tipo'),
            fechaOrigen: _parseFechaGoC(g('fechaOrigen')),
            fechaPago:   _parseFechaGoC(g('fechaPago')),
            orden:       _gNorm(g('orden')),
            nombre:      g('nombre'),
            cuotas:      parseInt(g('cuotas'))||0,
            importe:     _gParseMonto(g('importe')),
            totalCobrar: _gParseMonto(g('totalCobrar')),
            sucId:       g('sucId'),
            sucNombre:   g('sucNombre'),
            refExt:      g('refExt'),
          };
        }).filter(Boolean).filter(r => r.importe > 0 || r.orden !== '0');

        console.log('[GOC-PAGOS] Go Cuotas CSV parseados:', _GOC_PAGOS.length);
        res(_GOC_PAGOS);
      } catch(err) { rej(err); }
    };
    reader.onerror = () => rej(new Error('Error leyendo CSV'));
    reader.readAsText(file, 'utf-8');
  });
}

function _parseFechaGoC(s) {
  // "30/4/2026" o "30/04/2026" → "2026-04-30"
  if (!s) return '';
  const m = String(s).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return `${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`;
  return s.slice(0,10);
}

// ── Parser: Ventas XLSX ──────────────────────────────────────────────
function parseGocVentas(wb) {
  const ws   = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: false });
  if (!rows.length) { _GOC_VENTAS = []; return []; }

  const allKeys = Object.keys(rows[0]);
  const K = {
    fecha:       _resolveKey(allKeys, 'Fecha'),
    comprobante: _resolveKey(allKeys, 'Comprobante'),
    descripcion: _resolveKey(allKeys, 'Descripcion del Articulo', 'Descripción del Artículo', 'Articulo', 'Descripcion'),
    trazabilidad:_resolveKey(allKeys, 'Id. Trazabilidad', 'Id Trazabilidad', 'IMEI', 'Trazabilidad'),
    sucId:       _resolveKey(allKeys, 'Id Sucursal'),
    sucNombre:   _resolveKey(allKeys, 'Nombre de la Sucursal'),
    monto:       _resolveKey(allKeys, 'Total Bruto', 'Importe', 'Monto'),
    idArticulo:  _resolveKey(allKeys, 'Id Articulo', 'Id. Articulo'),
    vendedor:    _resolveKey(allKeys, 'Nombre del Vendedor', 'Vendedor'),
    cliente:     _resolveKey(allKeys, 'Razon Social del Cliente', 'Cliente'),
  };
  const g = (r,k) => (k && r[k] !== undefined) ? r[k] : null;

  _GOC_VENTAS = rows.map((r,i) => {
    const compRaw = String(g(r, K.comprobante) || '').trim();
    // Extraer n° final del comprobante: "FA A 0049 00059680" → "59680" (sin ceros líderes)
    const compNum = _gNorm(compRaw.split(/\s+/).pop());
    return {
      idx:          i,
      comprobante:  compRaw,
      compNum,                  // número limpio → comparar con SkyGC.Id
      descripcion:  String(g(r, K.descripcion)  || '').trim(),
      trazabilidad: String(g(r, K.trazabilidad) || '').trim(),
      sucId:        String(g(r, K.sucId)        || '').trim(),
      sucNombre:    String(g(r, K.sucNombre)    || '').trim(),
      monto:        _gParseMonto(g(r, K.monto)),
      idArticulo:   String(g(r, K.idArticulo)   || '').trim(),
      vendedor:     String(g(r, K.vendedor)      || '').trim(),
      cliente:      String(g(r, K.cliente)       || '').trim(),
    };
  }).filter(r => r.comprobante);

  // Índice por compNum para búsqueda rápida
  window._GOC_VENTAS_IDX = {};
  _GOC_VENTAS.forEach(v => {
    window._GOC_VENTAS_IDX[v.compNum] = v;
  });

  console.log('[GOC-VENTAS] Ventas parseadas:', _GOC_VENTAS.length);
  return _GOC_VENTAS;
}

// ════════════════════════════════════════════════════════════════════
// CRUCE
// ════════════════════════════════════════════════════════════════════
function cruzarGoCuotas() {
  _GOC_RESULT = [];
  if (!_GOC_SKY.length) return;

  // Índice de pagos por Número de Orden
  const pagoIdx = {};
  _GOC_PAGOS.forEach(p => { pagoIdx[p.orden] = p; });

  // Índice de ventas por compNum
  const ventaIdx = window._GOC_VENTAS_IDX || {};

  _GOC_SKY.forEach(sky => {
    // 1. Buscar pago en Go Cuotas CSV
    const pago = pagoIdx[sky.cupon] || null;

    // 2. Buscar venta en Ventas (link por Id → Comprobante)
    const idNorm = _gNorm(sky.id);
    const venta  = ventaIdx[idNorm] || null;

    // 3. Determinar estado de cobro
    let estadoCobro;
    if (!_GOC_PAGOS.length)        estadoCobro = 'SIN ARCHIVO PAGOS';
    else if (pago)                 estadoCobro = 'COBRADO';
    else                           estadoCobro = 'PENDIENTE';

    // 4. Verificar IMEI si el artículo lo requiere
    const desc       = venta?.descripcion || '';
    const imei       = venta?.trazabilidad || '';
    const reqImei    = _gRequiereImei(sky.plan, desc);
    let estadoImei   = null;
    if (reqImei) {
      if (!venta)         estadoImei = 'SIN VENTAS';
      else if (!imei || imei === '' || imei === '0')
                          estadoImei = 'IMEI FALTANTE';
      else if (!_gEsImei(imei))
                          estadoImei = 'IMEI INVÁLIDO';
      else                estadoImei = 'IMEI OK';
    }

    sky._pagos  = pago;
    sky._venta  = venta;
    sky._estado = estadoCobro;
    sky._imeiOk = estadoImei;

    _GOC_RESULT.push({ sky, pago, venta, estadoCobro, estadoImei, reqImei });
  });

  // Pagos en GoC sin match en SkyGC (cobrados pero no facturados en Skylab)
  const skyOrders = new Set(_GOC_SKY.map(s => s.cupon));
  _GOC_PAGOS.forEach(p => {
    if (!skyOrders.has(p.orden)) {
      _GOC_RESULT.push({ sky:null, pago:p, venta:null, estadoCobro:'EN GOC - SIN SKY', estadoImei:null, reqImei:false });
    }
  });

  console.log('[GOC] Cruce completado:', _GOC_RESULT.length, 'registros');
  return _GOC_RESULT;
}

// ════════════════════════════════════════════════════════════════════
// RENDER MÓDULO
// ════════════════════════════════════════════════════════════════════
function renderModuloGoCuotas() {
  const panel = document.getElementById('mod-goc');
  if (!panel) return;

  const hasSky  = _GOC_SKY.length   > 0;
  const hasPag  = _GOC_PAGOS.length > 0;
  const hasVen  = _GOC_VENTAS.length > 0;

  if (!hasSky) {
    panel.innerHTML = `
      <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;
        height:100%;gap:14px;color:var(--m2);text-align:center;padding:40px">
        <div style="font-size:42px;opacity:.15">💳</div>
        <div style="font-family:var(--head);font-size:16px;font-weight:700;color:var(--txt);opacity:.4">Go Cuotas</div>
        <p style="font-size:10px;max-width:420px;line-height:1.8">
          Cargá los archivos desde el panel izquierdo:<br>
          <b style="color:var(--grn)">Skylab Go Cuotas</b> (XLSX, requerido) ·
          <b style="color:var(--acc)">Go Cuotas CSV</b> (pagos) ·
          <b style="color:var(--yel)">Ventas</b> (artículos + IMEI)
        </p>
      </div>`;
    return;
  }

  cruzarGoCuotas();

  const cobrados   = _GOC_RESULT.filter(r => r.estadoCobro === 'COBRADO');
  const pendientes = _GOC_RESULT.filter(r => r.estadoCobro === 'PENDIENTE');
  const sinSky     = _GOC_RESULT.filter(r => r.estadoCobro === 'EN GOC - SIN SKY');
  const imeiIssues = _GOC_RESULT.filter(r => r.estadoImei && r.estadoImei !== 'IMEI OK');
  const imeiOk     = _GOC_RESULT.filter(r => r.estadoImei === 'IMEI OK');

  const mCob = cobrados.reduce((s,r) => s+(r.sky?.importe||0), 0);
  const mPen = pendientes.reduce((s,r) => s+(r.sky?.importe||0), 0);
  const total = _GOC_SKY.length;
  const pctCob = total ? (cobrados.length/total*100).toFixed(1) : '0.0';

  panel.innerHTML = `
  <div style="display:flex;flex-direction:column;flex:1;min-height:0;overflow:hidden">

    <!-- KPIs -->
    <div style="display:grid;grid-template-columns:repeat(6,1fr);gap:8px;
      padding:12px 18px;background:var(--s1);border-bottom:1px solid var(--b1);flex-shrink:0">
      <div class="dif-kpi"><div class="dif-kpi-lbl">Total Skylab GC</div>
        <div class="dif-kpi-val cyn">${total.toLocaleString('es-AR')}</div>
        <div style="font-size:8px;color:var(--m2)">${_gFmt(_GOC_SKY.reduce((s,r)=>s+r.importe,0))}</div>
      </div>
      <div class="dif-kpi" style="border-color:rgba(52,211,153,.25)">
        <div class="dif-kpi-lbl">✓ Cobrados GoC</div>
        <div class="dif-kpi-val grn">${cobrados.length.toLocaleString('es-AR')}</div>
        <div style="font-size:8px;color:var(--grn)">${_gFmt(mCob)}</div>
      </div>
      <div class="dif-kpi" style="border-color:rgba(248,113,113,.25)">
        <div class="dif-kpi-lbl">⏳ Pendientes</div>
        <div class="dif-kpi-val red">${pendientes.length.toLocaleString('es-AR')}</div>
        <div style="font-size:8px;color:var(--red)">${_gFmt(mPen)}</div>
      </div>
      <div class="dif-kpi" style="border-color:rgba(251,191,36,.25)">
        <div class="dif-kpi-lbl">⚠ En GoC sin SKY</div>
        <div class="dif-kpi-val yel">${sinSky.length}</div>
        <div style="font-size:8px;color:var(--yel)">cobrados no facturados</div>
      </div>
      <div class="dif-kpi">
        <div class="dif-kpi-lbl">% Cobrado</div>
        <div class="dif-kpi-val ${parseFloat(pctCob)>=85?'grn':parseFloat(pctCob)>=60?'yel':'red'}">${pctCob}%</div>
      </div>
      <div class="dif-kpi" style="border-color:rgba(248,113,113,.25)">
        <div class="dif-kpi-lbl">📱 IMEI issues</div>
        <div class="dif-kpi-val ${imeiIssues.length>0?'red':'grn'}">${imeiIssues.length}</div>
        <div style="font-size:8px;color:var(--m2)">${imeiOk.length} ok</div>
      </div>
    </div>

    <!-- Tab strip -->
    <div class="tab-strip" id="tab-strip-goc">
      <button class="tb active" onclick="showTab('goc-todo','tab-strip-goc',this)">
        📋 Todo <span class="cnt">${_GOC_RESULT.length}</span>
      </button>
      <button class="tb" onclick="showTab('goc-cobrados','tab-strip-goc',this)">
        ✓ Cobrados <span class="cnt" style="background:rgba(52,211,153,.15);color:var(--grn)">${cobrados.length}</span>
      </button>
      <button class="tb" onclick="showTab('goc-pendientes','tab-strip-goc',this)">
        ⏳ Pendientes <span class="cnt" style="background:rgba(248,113,113,.15);color:var(--red)">${pendientes.length}</span>
      </button>
      ${sinSky.length ? `<button class="tb" onclick="showTab('goc-sinsky','tab-strip-goc',this)" style="color:var(--yel)">
        ⚠ En GoC sin SKY <span class="cnt" style="background:rgba(251,191,36,.15);color:var(--yel)">${sinSky.length}</span>
      </button>` : ''}
      ${imeiIssues.length ? `<button class="tb" onclick="showTab('goc-imei','tab-strip-goc',this)" style="color:var(--red)">
        📱 IMEI <span class="cnt" style="background:rgba(248,113,113,.15);color:var(--red)">${imeiIssues.length}</span>
      </button>` : ''}
    </div>

    <!-- Tab bodies -->
    <div class="tab-body active" id="goc-todo" style="flex-direction:column;flex:1;min-height:0">
      ${_gToolbar('todo')}
      <div class="tbl-wrap"><table id="tbl-goc-todo"><thead></thead><tbody></tbody></table></div>
    </div>
    <div class="tab-body" id="goc-cobrados" style="flex-direction:column;flex:1;min-height:0">
      ${_gToolbar('cobrados')}
      <div class="tbl-wrap"><table id="tbl-goc-cobrados"><thead></thead><tbody></tbody></table></div>
    </div>
    <div class="tab-body" id="goc-pendientes" style="flex-direction:column;flex:1;min-height:0">
      ${_gToolbar('pendientes')}
      <div class="tbl-wrap"><table id="tbl-goc-pendientes"><thead></thead><tbody></tbody></table></div>
    </div>
    <div class="tab-body" id="goc-sinsky" style="flex-direction:column;flex:1;min-height:0">
      <div class="cor-hdr-bar" style="border-left:3px solid var(--yel)">
        <span class="cor-hdr-title" style="color:var(--yel)">En Go Cuotas sin factura en Skylab</span>
        <span class="cor-stats">${sinSky.length} órdenes cobradas sin registro en Skylab</span>
        <button class="dl-btn" style="background:#2d2000;color:#fbbf24;border:1px solid rgba(251,191,36,.3);margin-left:auto"
          onclick="exportarGoCuotas('sinsky')">⬇ Exportar</button>
      </div>
      <div class="tbl-wrap"><table id="tbl-goc-sinsky"><thead></thead><tbody></tbody></table></div>
    </div>
    <div class="tab-body" id="goc-imei" style="flex-direction:column;flex:1;min-height:0">
      <div class="cor-hdr-bar" style="border-left:3px solid var(--red)">
        <span class="cor-hdr-title" style="color:var(--red)">📱 IMEI pendiente o inválido</span>
        <span class="cor-stats">${imeiIssues.length} ventas de celulares con IMEI faltante</span>
        <button class="dl-btn" style="background:#2d0808;color:#f87171;border:1px solid rgba(248,113,113,.3);margin-left:auto"
          onclick="exportarGoCuotas('imei')">⬇ Exportar</button>
      </div>
      <div class="tbl-wrap"><table id="tbl-goc-imei"><thead></thead><tbody></tbody></table></div>
    </div>

  </div>`;

  // Renderizar tablas
  _renderGocTabla('todo');
  _renderGocTabla('cobrados');
  _renderGocTabla('pendientes');
  if (sinSky.length)     _renderGocSinSky(sinSky);
  if (imeiIssues.length) _renderGocImei(imeiIssues);

  const badge = document.getElementById('mcnt-goc');
  if (badge) badge.textContent = _GOC_RESULT.length || '—';
}

function _gToolbar(tipo) {
  return `<div class="cor-hdr-bar" style="border-left:3px solid var(--acc)">
    <span class="cor-hdr-title" style="color:var(--acc)">Go Cuotas — ${tipo==='todo'?'Vista completa':tipo==='cobrados'?'Cobrados':'Pendientes de cobro'}</span>
    <div style="margin-left:auto;display:flex;gap:6px">
      <select class="filter-sel" id="goc-flt-${tipo}-suc" onchange="_renderGocTabla('${tipo}')" style="font-size:8px">
        <option value="">Todas las suc.</option>
        ${[...new Set(_GOC_RESULT.filter(r=>r.sky).map(r=>r.sky.sucursal).filter(Boolean))].sort()
          .map(s=>`<option value="${s}">${s}</option>`).join('')}
      </select>
      <input class="filter-inp" id="goc-flt-${tipo}-search" placeholder="Nombre, orden, plan..."
        oninput="_renderGocTabla('${tipo}')" style="width:160px">
      <button class="dl-btn" style="background:#14532d;color:#86efac"
        onclick="exportarGoCuotas('${tipo}')">⬇ Exportar</button>
    </div>
  </div>`;
}

function _renderGocTabla(tipo) {
  const tbl = document.getElementById(`tbl-goc-${tipo}`); if (!tbl) return;
  const fltSuc = document.getElementById(`goc-flt-${tipo}-suc`)?.value || '';
  const fSearch= (document.getElementById(`goc-flt-${tipo}-search`)?.value||'').toLowerCase();

  let filas = _GOC_RESULT.filter(r => r.sky); // solo los que tienen Skylab
  if (tipo === 'cobrados')   filas = filas.filter(r => r.estadoCobro === 'COBRADO');
  if (tipo === 'pendientes') filas = filas.filter(r => r.estadoCobro === 'PENDIENTE');
  if (fltSuc)  filas = filas.filter(r => r.sky.sucursal === fltSuc);
  if (fSearch) filas = filas.filter(r => {
    const hay = [r.sky.asiento, r.sky.cupon, r.sky.plan, r.sky.vendedor,
                 r.pago?.nombre||'', r.venta?.descripcion||'', r.venta?.trazabilidad||''].join(' ').toLowerCase();
    return hay.includes(fSearch);
  });

  const HDR = ['Estado cobro','Fecha venta','Suc.','Vendedor','Plan','Nro. Orden GoC',
               'Importe SKY','Importe GoC','Fecha pago GoC','Artículo','IMEI / Trazabilidad'];
  tbl.querySelector('thead').innerHTML = `<tr>${HDR.map(h=>`<th>${h}</th>`).join('')}</tr>`;

  if (!filas.length) {
    tbl.querySelector('tbody').innerHTML =
      `<tr><td colspan="${HDR.length}" style="padding:20px;text-align:center;color:var(--m2);font-size:10px">Sin registros.</td></tr>`;
    return;
  }

  tbl.querySelector('tbody').innerHTML = filas.map(({ sky, pago, venta, estadoCobro, estadoImei }) => {
    const cobColor = estadoCobro==='COBRADO' ? 'var(--grn)' : estadoCobro==='PENDIENTE' ? 'var(--red)' : 'var(--yel)';
    const imeiCell = estadoImei
      ? `<span style="font-size:8px;padding:1px 5px;border-radius:3px;
          color:${estadoImei==='IMEI OK'?'var(--grn)':estadoImei==='IMEI FALTANTE'?'var(--red)':'var(--yel)'};
          border:1px solid currentColor;background:currentColor22">${estadoImei}</span>
         <span style="font-size:9px;color:var(--m2)"> ${venta?.trazabilidad||''}</span>`
      : `<span style="color:var(--m2);font-size:9px">${venta?.trazabilidad||'—'}</span>`;

    return `<tr class="${estadoCobro==='COBRADO'?'row-ok':estadoCobro==='PENDIENTE'?'row-mal':'row-com'}">
      <td><span style="font-size:8px;padding:2px 7px;border-radius:3px;color:${cobColor};
        border:1px solid ${cobColor}55;background:${cobColor}18">${estadoCobro}</span></td>
      <td>${sky.fecha}</td>
      <td>${sky.sucursal}</td>
      <td style="max-width:130px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${sky.vendedor}</td>
      <td>${sky.plan}</td>
      <td class="num" style="font-family:var(--mono);color:var(--cyn)">${sky.cupon}</td>
      <td class="num" style="font-weight:600">${_gFmt(sky.importe)}</td>
      <td class="num" style="color:${pago?'var(--grn)':'var(--m2)'}">${pago?_gFmt(pago.importe):'—'}</td>
      <td>${pago?.fechaPago||'—'}</td>
      <td style="font-size:9px;max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap"
        title="${venta?.descripcion||''}">${venta?.descripcion||'—'}</td>
      <td>${imeiCell}</td>
    </tr>`;
  }).join('');
}

function _renderGocSinSky(filas) {
  const tbl = document.getElementById('tbl-goc-sinsky'); if (!tbl) return;
  const HDR = ['Fecha origen','Fecha pago','Nro. Orden GoC','Nombre','Cuotas','Importe','Total a cobrar','Sucursal','Ref. Externa'];
  tbl.querySelector('thead').innerHTML = `<tr>${HDR.map(h=>`<th>${h}</th>`).join('')}</tr>`;
  tbl.querySelector('tbody').innerHTML = filas.map(({ pago }) => `<tr class="row-com">
    <td>${pago.fechaOrigen}</td>
    <td>${pago.fechaPago}</td>
    <td class="num" style="font-family:var(--mono);color:var(--yel)">${pago.orden}</td>
    <td>${pago.nombre}</td>
    <td class="num">${pago.cuotas}</td>
    <td class="num" style="font-weight:600">${_gFmt(pago.importe)}</td>
    <td class="num" style="color:var(--grn)">${_gFmt(pago.totalCobrar)}</td>
    <td>${pago.sucNombre}</td>
    <td style="font-size:9px;color:var(--m2)">${pago.refExt||'—'}</td>
  </tr>`).join('');
}

function _renderGocImei(filas) {
  const tbl = document.getElementById('tbl-goc-imei'); if (!tbl) return;
  const HDR = ['Estado IMEI','Fecha venta','Suc.','Vendedor','Plan','Nro. Orden GoC','Importe','Artículo','IMEI registrado','Comprobante'];
  tbl.querySelector('thead').innerHTML = `<tr>${HDR.map(h=>`<th>${h}</th>`).join('')}</tr>`;
  tbl.querySelector('tbody').innerHTML = filas.map(({ sky, venta, estadoImei }) => {
    const c = estadoImei==='IMEI FALTANTE'?'var(--red)':estadoImei==='IMEI INVÁLIDO'?'var(--yel)':'var(--org)';
    return `<tr class="row-mal">
      <td><span style="font-size:8px;padding:2px 7px;border-radius:3px;color:${c};
        border:1px solid ${c}55;background:${c}18">${estadoImei}</span></td>
      <td>${sky.fecha}</td>
      <td>${sky.sucursal}</td>
      <td style="max-width:130px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${sky.vendedor}</td>
      <td>${sky.plan}</td>
      <td class="num" style="font-family:var(--mono)">${sky.cupon}</td>
      <td class="num" style="font-weight:600">${_gFmt(sky.importe)}</td>
      <td style="max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap"
        title="${venta?.descripcion||''}">${venta?.descripcion||'—'}</td>
      <td style="font-family:var(--mono);font-size:9px;color:${venta?.trazabilidad?'var(--m1)':'var(--red)'}">
        ${venta?.trazabilidad||'(vacío)'}</td>
      <td style="font-size:9px;color:var(--m2)">${venta?.comprobante||'—'}</td>
    </tr>`;
  }).join('');
}

// ════════════════════════════════════════════════════════════════════
// EXPORT
// ════════════════════════════════════════════════════════════════════
function exportarGoCuotas(tipo) {
  let filas = _GOC_RESULT;
  if (tipo === 'cobrados')   filas = filas.filter(r => r.estadoCobro === 'COBRADO');
  if (tipo === 'pendientes') filas = filas.filter(r => r.estadoCobro === 'PENDIENTE');
  if (tipo === 'sinsky')     filas = filas.filter(r => r.estadoCobro === 'EN GOC - SIN SKY');
  if (tipo === 'imei')       filas = filas.filter(r => r.estadoImei && r.estadoImei !== 'IMEI OK');

  if (!filas.length) { alert('Sin datos para exportar.'); return; }

  const HDR = ['Estado cobro','Estado IMEI','Nro. Asiento SKY','Fecha venta','Sucursal','Vendedor',
               'Plan','Nro. Orden GoC','Importe SKY','Neto SKY','Importe GoC','Total a cobrar GoC',
               'Fecha pago GoC','Nombre cliente GoC','Artículo','IMEI / Trazabilidad','Comprobante'];

  const data = filas.map(({ sky, pago, venta, estadoCobro, estadoImei }) => [
    estadoCobro,
    estadoImei || '',
    sky?.asiento || '',
    sky?.fecha   || pago?.fechaOrigen || '',
    sky?.sucursal|| pago?.sucNombre   || '',
    sky?.vendedor|| pago?.nombre      || '',
    sky?.plan    || '',
    sky?.cupon   || pago?.orden       || '',
    sky?.importe || 0,
    sky?.neto    || 0,
    pago?.importe     || 0,
    pago?.totalCobrar || 0,
    pago?.fechaPago   || '',
    pago?.nombre      || '',
    venta?.descripcion || '',
    venta?.trazabilidad|| '',
    venta?.comprobante || '',
  ]);

  _exportXlsx([HDR, ...data], 'Go Cuotas', `GoCuotas_${tipo}_${hoy()}.xlsx`);
}
