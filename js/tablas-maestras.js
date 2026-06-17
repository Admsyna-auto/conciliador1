// ═══════════════════════════════════════════════════════════════════
// TABLAS-MAESTRAS.JS — CRUD de todas las tablas de referencia
// ═══════════════════════════════════════════════════════════════════


// ── DESCARTAR CAMBIOS — recargar desde IndexedDB
let _tmSnapshot = null; // snapshot de TM al abrir cada tabla

function tomarSnapshotTM() {
  _tmSnapshot = JSON.parse(JSON.stringify(TM));
}

function descartarTM(key) {
  if (!_tmSnapshot) {
    alert('No hay snapshot previo para descartar.');
    return;
  }
  if (!confirm('¿Descartás los cambios y volvés al último estado guardado?')) return;
  if (_tmSnapshot[key] !== undefined) {
    TM[key] = JSON.parse(JSON.stringify(_tmSnapshot[key]));
  }
  showTM(key);
}

// ── Renderizar el módulo completo de Tablas Maestras
function renderModuloTM() {
  const cont = document.getElementById('mod-tm');
  if (!cont) return;
  cont.innerHTML = `
    <div class="tm-layout">
      <div class="tm-sidebar">
        ${['sucursales','vendedores','terminales','comercios','tarjetas','planes','tasas',
           'plazos','feriados','equivalencias','motivos','estados'].map(k => `
          <button class="tm-nav-btn" id="tmnav-${k}" onclick="showTM('${k}')">${tmLabel(k)}</button>
        `).join('')}
      </div>
      <div class="tm-body" id="tm-body"></div>
    </div>`;
  showTM('sucursales');
}

function tmLabel(k) {
  return {
    sucursales:'Sucursales', vendedores:'Vendedores', terminales:'Terminales',
    comercios:'Nros. Comercio', tarjetas:'Tarjetas', planes:'Planes/Cuotas',
    tasas:'Tasas / Acuerdos', plazos:'Plazos Acreditación', feriados:'Feriados',
    equivalencias:'Equiv. Comercio', motivos:'Motivos', estados:'Estados',
  }[k] || k;
}

function showTM(key) {
  tomarSnapshotTM(); // snapshot antes de cualquier cambio
  document.querySelectorAll('.tm-nav-btn').forEach(b => b.classList.remove('active'));
  const btn = document.getElementById(`tmnav-${key}`);
  if (btn) btn.classList.add('active');
  const body = document.getElementById('tm-body');
  if (!body) return;

  switch(key) {
    case 'sucursales':   body.innerHTML = renderTMGeneric(key, ['id','nombre','estado'],
      ['ID','Nombre','Estado'], () => ({id:'',nombre:'',estado:'Activo'})); break;
    case 'vendedores':   body.innerHTML = renderTMGeneric(key, ['id','nombre','sucursal','legajo'],
      ['ID','Nombre','Sucursal','Legajo'], () => ({id:'',nombre:'',sucursal:'',legajo:''})); break;
    case 'terminales':   body.innerHTML = renderTMTerminales(); break;
    case 'comercios':    body.innerHTML = renderTMGeneric(key, ['nroCom','procesadora','acuerdo','vigDesde','vigHasta'],
      ['Nro. Comercio','Procesadora','Acuerdo','Desde','Hasta'],
      () => ({nroCom:'',procesadora:'FISERV',acuerdo:'',vigDesde:'',vigHasta:''})); break;
    case 'tarjetas':     body.innerHTML = renderTMTarjetas(); break;
    case 'planes':       body.innerHTML = renderTMGeneric(key, ['plan','cuotas','tarjeta','procesadora','codigos'],
      ['Plan','Cuotas','Tarjeta','Procesadora','Códigos'], () => ({plan:'',cuotas:'',tarjeta:'',procesadora:'',codigos:''})); break;
    case 'tasas':        body.innerHTML = renderTMTasas(); break;
    case 'plazos':       body.innerHTML = renderTMGeneric(key,
      ['procesadora','comercio','tarjeta','dias_habiles','vigDesde','vigHasta'],
      ['Procesadora','Comercio','Tarjeta','Días Háb.','Desde','Hasta'],
      () => ({procesadora:'FISERV',comercio:'',tarjeta:'',dias_habiles:'2',vigDesde:'',vigHasta:''})); break;
    case 'feriados':     body.innerHTML = renderTMFeriados(); break;
    case 'equivalencias':body.innerHTML = renderTMEquiv(); break;
    case 'motivos':      body.innerHTML = renderTMCatalogo('motivos','Motivos de diferencia'); break;
    case 'estados':      body.innerHTML = renderTMCatalogo('estados','Estados de conciliación'); break;
  }

  // Bindings post-render
  bindTMImport(key);
}

// ── Tabla genérica editable
function renderTMGeneric(key, fields, headers, newRowFn) {
  const rows = TM[key] || [];
  const hdr = headers.map(h => `<th>${h}</th>`).join('') + '<th style="width:70px"></th>';
  const bodyRows = rows.map((r, i) => `
    <tr id="tmr-${key}-${i}">
      ${fields.map(f => `<td><input class="tm-inp" value="${esc(r[f]??'')}"
        onchange="updateTMRow('${key}',${i},'${f}',this.value)"></td>`).join('')}
      <td><button class="tm-del-btn" onclick="deleteTMRow('${key}',${i})">×</button></td>
    </tr>`).join('');

  return `
    <div class="tm-panel">
      <div class="tm-panel-hdr">
        <span class="tm-panel-title">${tmLabel(key)}</span>
        <div class="tm-panel-actions">
          <label class="tm-import-btn" title="Importar desde Excel">
            📥 Importar Excel
            <input type="file" accept=".xlsx,.xls,.csv" style="display:none"
              onchange="importarTMExcel('${key}',this)">
          </label>
          <button class="tm-export-btn" onclick="exportarTM('${key}')">📤 Exportar</button>
          <button class="tm-add-btn" onclick="addTMRow('${key}')">+ Agregar</button>
        </div>
      </div>
      <div class="tm-table-wrap">
        <table class="tm-table">
          <thead><tr>${hdr}</tr></thead>
          <tbody id="tmbody-${key}">${bodyRows}</tbody>
        </table>
      </div>
      <div class="tm-footer">
        <span class="tm-count">${rows.length} registros</span>
        <button class="tm-discard-btn" onclick="descartarTM('${key}')">↩ Descartar</button>
        <button class="tm-save-btn" onclick="saveTM()">💾 Guardar cambios</button>
      </div>
    </div>`;
}

// ── Terminales (tiene campos especiales)
function renderTMTerminales() {
  const rows = TM.terminales || [];
  const fields = ['terminal','procesadora','sucursal','vigDesde','vigHasta'];
  const headers = ['Terminal','Procesadora','Sucursal','Desde','Hasta'];
  return renderTMGeneric('terminales', fields, headers,
    () => ({terminal:'',procesadora:'FISERV',sucursal:'',vigDesde:'',vigHasta:''}));
}

// ── Tasas (tabla compleja)
function renderTMTasas() {
  const rows = TM.tasas || [];
  const fields = ['acuerdo','procesadora','comercio','tarjeta','plan','cuotas','tasa','coef','vigDesde','vigHasta'];
  const headers = ['Acuerdo','Procesadora','Comercio','Tarjeta','Plan','Cuotas','Tasa %','Coef.','Desde','Hasta'];
  return renderTMGeneric('tasas', fields, headers,
    () => ({acuerdo:'',procesadora:'',comercio:'',tarjeta:'',plan:'',cuotas:'',tasa:'',coef:'',vigDesde:'',vigHasta:''}));
}

// ── Equivalencias de tarjetas Procesadora ↔ Skylab ───────────────────
function renderTMTarjetas() {
  if (!TM.tarjetas) TM.tarjetas = [];

  // Tarjetas únicas conocidas en Skylab (para datalist de sugerencias)
  const skyTarjetas = [...new Set(
    (typeof RESULTADO !== 'undefined' ? RESULTADO : [])
      .map(r => r.sky?.tarjeta).filter(Boolean)
  )].sort();

  // Tarjetas únicas detectadas en archivos de liquidación cargados
  const liqTarjetasDetectadas = [...new Set(
    (typeof _LIQ_CUPONES !== 'undefined' ? _LIQ_CUPONES : [])
      .map(r => r.tarjeta).filter(Boolean)
  )].sort();

  const rows = TM.tarjetas;
  const esc  = v => String(v||'').replace(/"/g,'&quot;').replace(/</g,'&lt;');

  const bodyRows = rows.length ? rows.map((r, i) => `
    <tr style="border-bottom:1px solid var(--b0)">
      <td style="padding:4px 6px">
        <input class="tm-inp" value="${esc(r.tarjeta)}"
          list="tm-liq-tarjetas-list"
          placeholder="Ej: TARJETA VISA BANCARIZADA"
          onchange="updateTMRow('tarjetas',${i},'tarjeta',this.value)"
          style="width:260px">
      </td>
      <td style="padding:4px 6px;text-align:center;color:var(--m2);font-size:18px">→</td>
      <td style="padding:4px 6px">
        <input class="tm-inp" value="${esc(r.equivSkylab)}"
          list="tm-sky-tarjetas-list"
          placeholder="Ej: VISA"
          onchange="updateTMRow('tarjetas',${i},'equivSkylab',this.value)"
          style="width:200px">
      </td>
      <td style="padding:4px 6px">
        <input class="tm-inp" value="${esc(r.equivProc||'')}"
          placeholder="(opcional)"
          onchange="updateTMRow('tarjetas',${i},'equivProc',this.value)"
          style="width:160px">
      </td>
      <td style="padding:4px 6px">
        <button class="tm-del-btn" onclick="deleteTMRow('tarjetas',${i})">×</button>
      </td>
    </tr>`).join('')
  : `<tr><td colspan="5" style="padding:20px;text-align:center;color:var(--m2);font-size:10px">
      Sin equivalencias configuradas — usá "Detectar desde liquidaciones" para cargar automáticamente
    </td></tr>`;

  const detectedBadge = liqTarjetasDetectadas.length
    ? `<span style="font-size:9px;color:var(--grn)">${liqTarjetasDetectadas.length} tarjetas en archivo cargado</span>`
    : `<span style="font-size:9px;color:var(--m2)">Cargá un archivo de liquidaciones para auto-detectar</span>`;

  return `
    <datalist id="tm-sky-tarjetas-list">
      ${skyTarjetas.map(t => `<option value="${esc(t)}">`).join('')}
    </datalist>
    <datalist id="tm-liq-tarjetas-list">
      ${liqTarjetasDetectadas.map(t => `<option value="${esc(t)}">`).join('')}
    </datalist>

    <div style="display:flex;flex-direction:column;height:100%;overflow:hidden">

      <!-- Header -->
      <div style="padding:14px 18px 10px;flex-shrink:0;border-bottom:1px solid var(--b1)">
        <div style="font-size:12px;font-weight:700;color:var(--txt);margin-bottom:4px">
          Equivalencias de Tarjetas
        </div>
        <p style="font-size:9px;color:var(--m2);margin:0;line-height:1.6">
          Mapea como aparece el nombre de tarjeta en el archivo de liquidaciones (FISERV/GETPOS)
          al nombre equivalente en Skylab. Se usa para detectar diferencias reales de tarjeta.
        </p>
      </div>

      <!-- Toolbar -->
      <div style="display:flex;align-items:center;gap:10px;padding:8px 18px;
        flex-shrink:0;border-bottom:1px solid var(--b0);background:var(--s1)">
        ${detectedBadge}
        <button class="tm-add-btn" onclick="_tmDetectarTarjetas()"
          ${liqTarjetasDetectadas.length ? '' : 'disabled'}
          style="${liqTarjetasDetectadas.length ? '' : 'opacity:.4;cursor:not-allowed'}">
          ⬇ Detectar desde liquidaciones
        </button>
        <div style="flex:1"></div>
        <span class="tm-count">${rows.length} equivalencias</span>
        <button class="tm-add-btn" onclick="addTMRow('tarjetas')">+ Agregar manual</button>
        <button class="tm-discard-btn" onclick="descartarTM('tarjetas')">↩ Descartar</button>
        <button class="tm-save-btn" onclick="saveTM()">💾 Guardar</button>
      </div>

      <!-- Tabla -->
      <div style="flex:1;overflow-y:auto">
        <table class="tm-table" style="width:100%;border-collapse:collapse">
          <thead><tr style="background:var(--s2)">
            <th style="padding:8px 6px;text-align:left;font-size:9px;color:var(--m2);
              white-space:nowrap;border-bottom:1px solid var(--b1)">
              Tarjeta en Liquidaciones (Procesadora)
            </th>
            <th style="width:32px"></th>
            <th style="padding:8px 6px;text-align:left;font-size:9px;color:var(--m2);
              white-space:nowrap;border-bottom:1px solid var(--b1)">
              Equivalente en Skylab
            </th>
            <th style="padding:8px 6px;text-align:left;font-size:9px;color:var(--m2);
              white-space:nowrap;border-bottom:1px solid var(--b1)">
              Equiv. Procesadora (opcional)
            </th>
            <th style="width:40px;border-bottom:1px solid var(--b1)"></th>
          </tr></thead>
          <tbody>${bodyRows}</tbody>
        </table>
      </div>
    </div>`;
}

window._tmDetectarTarjetas = function() {
  const liqTarjetas = [...new Set(
    (typeof _LIQ_CUPONES !== 'undefined' ? _LIQ_CUPONES : [])
      .map(r => r.tarjeta).filter(Boolean)
  )].sort();
  if (!liqTarjetas.length) return;
  if (!TM.tarjetas) TM.tarjetas = [];
  const existentes = new Set(TM.tarjetas.map(r => (r.tarjeta||'').toUpperCase().trim()));
  let added = 0;
  for (const t of liqTarjetas) {
    if (!existentes.has(t.toUpperCase().trim())) {
      TM.tarjetas.push({ tarjeta: t, equivSkylab: '', equivProc: '' });
      added++;
    }
  }
  scheduleAutoSave();
  showTM('tarjetas');
  if (added === 0) alert('Todas las tarjetas detectadas ya están en la tabla.');
};

// ── Feriados nacionales
function renderTMFeriados() {
  const rows = TM.feriados || [];
  const bodyRows = rows.map((r, i) => `
    <tr>
      <td><input class="tm-inp" value="${esc(r.fecha||'')}" placeholder="YYYY-MM-DD"
        onchange="updateTMRow('feriados',${i},'fecha',this.value)"></td>
      <td><input class="tm-inp" value="${esc(r.descripcion||'')}" placeholder="Descripción"
        onchange="updateTMRow('feriados',${i},'descripcion',this.value)"></td>
      <td><button class="tm-del-btn" onclick="deleteTMRow('feriados',${i})">×</button></td>
    </tr>`).join('');

  return `
    <div class="tm-panel">
      <div class="tm-panel-hdr">
        <span class="tm-panel-title">Feriados Nacionales</span>
        <div class="tm-panel-actions">
          <label class="tm-import-btn" title="Importar desde Excel">
            📥 Importar Excel
            <input type="file" accept=".xlsx,.xls,.csv" style="display:none"
              onchange="importarTMExcel('feriados',this)">
          </label>
          <button class="tm-export-btn" onclick="exportarTM('feriados')">📤 Exportar</button>
          <button class="tm-add-btn" onclick="addTMRow('feriados')">+ Agregar</button>
        </div>
      </div>
      <div class="tm-table-wrap">
        <table class="tm-table">
          <thead><tr><th>Fecha</th><th>Descripción</th><th style="width:50px"></th></tr></thead>
          <tbody id="tmbody-feriados">${bodyRows}</tbody>
        </table>
      </div>
      <div class="tm-footer">
        <span class="tm-count">${rows.length} feriados</span>
        <button class="tm-discard-btn" onclick="descartarTM('feriados')">↩ Descartar</button>
        <button class="tm-save-btn" onclick="saveTM()">💾 Guardar cambios</button>
      </div>
    </div>`;
}

// ── Equivalencias de comercio
function renderTMEquiv() {
  const equivs = TM.equivalencias || {};
  const entries = Object.entries(equivs);
  const rows = entries.map(([k, v], i) => {
    const [com, suc] = k.split('|');
    return `
      <tr>
        <td><input class="tm-inp" id="ec-${i}" value="${esc(com)}" placeholder="Com. SKY"></td>
        <td><input class="tm-inp" id="es-${i}" value="${esc(suc)}" placeholder="Suc."></td>
        <td style="text-align:center;color:var(--m2)">→</td>
        <td><input class="tm-inp" id="ev-${i}" value="${esc(v)}" placeholder="Com. FISERV"></td>
        <td><button class="tm-del-btn" onclick="delEquivRow(${i})">×</button></td>
      </tr>`;
  }).join('');

  return `
    <div class="tm-panel">
      <div class="tm-panel-hdr">
        <span class="tm-panel-title">Equivalencias de comercio</span>
        <div class="tm-panel-actions">
          <button class="tm-add-btn" onclick="addEquivRow()">+ Agregar</button>
        </div>
      </div>
      <div class="tm-table-wrap">
        <table class="tm-table">
          <thead><tr><th>Com. Skylab</th><th>Sucursal</th><th></th><th>Com. FISERV equiv.</th><th style="width:50px"></th></tr></thead>
          <tbody id="tmbody-equiv">${rows}</tbody>
        </table>
      </div>
      <div class="tm-footer">
        <span class="tm-count">${entries.length} equivalencias</span>
        <button class="tm-discard-btn" onclick="descartarTM('equivalencias')">↩ Descartar</button>
        <button class="tm-save-btn" onclick="saveEquivRows()">💾 Guardar y aplicar</button>
      </div>
    </div>`;
}

// ── Catálogo (motivos/estados)
function renderTMCatalogo(key, titulo) {
  const items = TM[key] || [];
  const rows = items.map((item, i) => `
    <tr>
      <td><input class="tm-inp" value="${esc(item)}"
        onchange="TM['${key}'][${i}]=this.value;scheduleAutoSave()"></td>
      <td><button class="tm-del-btn" onclick="TM['${key}'].splice(${i},1);showTM('${key}')">×</button></td>
    </tr>`).join('');

  return `
    <div class="tm-panel">
      <div class="tm-panel-hdr">
        <span class="tm-panel-title">${titulo}</span>
        <div class="tm-panel-actions">
          <button class="tm-add-btn" onclick="TM['${key}'].push('Nuevo');showTM('${key}')">+ Agregar</button>
        </div>
      </div>
      <div class="tm-table-wrap">
        <table class="tm-table">
          <thead><tr><th>Descripción</th><th style="width:50px"></th></tr></thead>
          <tbody>${rows}</tbody>
        </table>
      </div>
      <div class="tm-footer">
        <span class="tm-count">${items.length} ítems</span>
        <button class="tm-discard-btn" onclick="descartarTM('${key}')">↩ Descartar</button>
        <button class="tm-save-btn" onclick="saveTM()">💾 Guardar</button>
      </div>
    </div>`;
}

// ════════════════════════════════════════════════════════════════════
// CRUD GENÉRICO
// ════════════════════════════════════════════════════════════════════
function updateTMRow(key, idx, field, value) {
  if (!TM[key]?.[idx]) return;
  TM[key][idx][field] = value;
  scheduleAutoSave();
}

function deleteTMRow(key, idx) {
  if (!TM[key]) return;
  TM[key].splice(idx, 1);
  showTM(key);
  scheduleAutoSave();
}

function addTMRow(key) {
  const defaults = {
    sucursales:  { id:'', nombre:'', estado:'Activo' },
    vendedores:  { id:'', nombre:'', sucursal:'', legajo:'' },
    terminales:  { terminal:'', procesadora:'FISERV', sucursal:'', vigDesde:'', vigHasta:'' },
    comercios:   { nroCom:'', procesadora:'FISERV', acuerdo:'', vigDesde:'', vigHasta:'' },
    tarjetas:    { tarjeta:'', equivSkylab:'', equivProc:'' },
    planes:      { plan:'', cuotas:'', tarjeta:'', procesadora:'', codigos:'' },
    tasas:       { acuerdo:'', procesadora:'', comercio:'', tarjeta:'', plan:'', cuotas:'', tasa:'', coef:'', vigDesde:'', vigHasta:'' },
    plazos:      { procesadora:'FISERV', comercio:'', tarjeta:'', dias_habiles:'2', vigDesde:'', vigHasta:'' },
    feriados:    { fecha:'', descripcion:'' },
  };
  if (!TM[key]) TM[key] = [];
  TM[key].push({ ...(defaults[key] || {}) });
  showTM(key);
  // Scroll al final
  setTimeout(() => {
    const wrap = document.querySelector('.tm-table-wrap');
    if (wrap) wrap.scrollTop = wrap.scrollHeight;
  }, 50);
}

async function saveTM() {
  const ok = await guardarTM();
  const footer = document.querySelector('.tm-footer .tm-save-btn');
  if (footer) {
    footer.textContent = ok ? '✓ Guardado' : '⚠ Error';
    setTimeout(() => { footer.textContent = '💾 Guardar cambios'; }, 2000);
  }
  scheduleAutoSave();
}

// ── Equivalencias especiales
function addEquivRow() {
  const tbody = document.getElementById('tmbody-equiv');
  if (!tbody) return;
  const i = tbody.children.length;
  const tr = document.createElement('tr');
  tr.innerHTML = `
    <td><input class="tm-inp" id="ec-${i}" value="" placeholder="Com. SKY"></td>
    <td><input class="tm-inp" id="es-${i}" value="" placeholder="Suc."></td>
    <td style="text-align:center;color:var(--m2)">→</td>
    <td><input class="tm-inp" id="ev-${i}" value="" placeholder="Com. FISERV"></td>
    <td><button class="tm-del-btn" onclick="this.closest('tr').remove()">×</button></td>`;
  tbody.appendChild(tr);
}

function delEquivRow(i) {
  document.querySelector(`#tmbody-equiv tr:nth-child(${i+1})`)?.remove();
}

function saveEquivRows() {
  const rows = document.querySelectorAll('#tmbody-equiv tr');
  const nuevo = {};
  rows.forEach((tr, i) => {
    const com = tr.querySelector(`[id^="ec-"]`)?.value?.trim();
    const suc = tr.querySelector(`[id^="es-"]`)?.value?.trim();
    const val = tr.querySelector(`[id^="ev-"]`)?.value?.trim();
    if (com && val) nuevo[`${com}|${suc??''}`] = val;
  });
  TM.equivalencias = nuevo;
  // Sincronizar con EQUIV_COMERCIO global
  if (typeof EQUIV_COMERCIO !== 'undefined') {
    Object.keys(EQUIV_COMERCIO).forEach(k => delete EQUIV_COMERCIO[k]);
    Object.assign(EQUIV_COMERCIO, nuevo);
  }
  saveTM();
  const btn = document.querySelector('.tm-footer .tm-save-btn');
  if (btn) { btn.textContent = '✓ Aplicado'; setTimeout(() => btn.textContent = '💾 Guardar y aplicar', 2000); }
}

// ════════════════════════════════════════════════════════════════════
// IMPORTAR EXCEL → TABLA MAESTRA
// ════════════════════════════════════════════════════════════════════
function importarTMExcel(key, input) {
  const file = input.files[0]; if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb   = XLSX.read(e.target.result, { type:'array', raw:false });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval:'' });

      // Mapeo automático de columnas por nombre normalizado
      const mapCols = {
        sucursales:  { 'id':'id','nombre':'nombre','estado':'estado' },
        vendedores:  { 'id':'id','nombre':'nombre','sucursal':'sucursal','legajo':'legajo' },
        terminales:  { 'terminal':'terminal','procesadora':'procesadora','sucursal':'sucursal',
                       'nro comercio':'nroCom','nrocomercio':'nroCom','desde':'vigDesde','hasta':'vigHasta' },
        comercios:   { 'nro comercio':'nroCom','procesadora':'procesadora','acuerdo':'acuerdo',
                       'desde':'vigDesde','hasta':'vigHasta' },
        tarjetas:    { 'tarjeta':'tarjeta','equiv skylab':'equivSkylab','equiv proc':'equivProc' },
        planes:      { 'plan':'plan','cuotas':'cuotas','tarjeta':'tarjeta','procesadora':'procesadora','codigos':'codigos' },
        tasas:       { 'acuerdo':'acuerdo','procesadora':'procesadora','comercio':'comercio',
                       'tarjeta':'tarjeta','plan':'plan','cuotas':'cuotas',
                       'tasa':'tasa','tasa%':'tasa',
                       'coef':'coef','coeficiente':'coef',
                       'desde':'vigDesde','hasta':'vigHasta' },
        plazos:      { 'procesadora':'procesadora','comercio':'comercio','tarjeta':'tarjeta',
                       'diashabiles':'dias_habiles','diashab':'dias_habiles','dias':'dias_habiles',
                       'desde':'vigDesde','hasta':'vigHasta' },
        feriados:    { 'fecha':'fecha','descripcion':'descripcion','descripción':'descripcion' },
      };

      const map = mapCols[key] || {};
      // Normaliza header: minúsculas, reemplaza vocales acentuadas, sin espacios/puntos
      const normHdr = h => String(h ?? '').toLowerCase()
        .replace(/[áàâä]/g,'a').replace(/[éèêë]/g,'e')
        .replace(/[íìîï]/g,'i').replace(/[óòôö]/g,'o')
        .replace(/[úùûü]/g,'u').replace(/[ñ]/g,'n')
        .replace(/[\s.()]/g,'');
      const imported = rows.map(r => {
        const obj = {};
        for (const [col, field] of Object.entries(map)) {
          const found = Object.keys(r).find(k => normHdr(k) === col.replace(/[\s.]/g,''));
          if (found !== undefined && obj[field] === undefined) obj[field] = String(r[found] ?? '');
        }
        return obj;
      }).filter(r => Object.values(r).some(v => v !== ''));

      // Post-proceso: fechas serial Excel → ISO  (también acepta DD/MM/YYYY texto)
      const serialToISO = v => {
        const n = parseFloat(v);
        if (!isNaN(n) && n >= 10000)
          return new Date(Math.round((n - 25569) * 86400000)).toISOString().slice(0, 10);
        // Fecha en texto DD/MM/YYYY o D/M/YYYY (formato argentino)
        const m = String(v).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
        if (m) return `${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`;
        return v;
      };
      if (key === 'tasas') {
        imported.forEach(row => {
          if (row.tasa) {
            const n = parseFloat(row.tasa);
            if (!isNaN(n) && n > 0 && n < 1) row.tasa = (n * 100).toFixed(4).replace(/\.?0+$/, '');
          }
          if (row.vigDesde) row.vigDesde = serialToISO(row.vigDesde);
          if (row.vigHasta) row.vigHasta = serialToISO(row.vigHasta);
        });
      }
      if (key === 'plazos') {
        imported.forEach(row => {
          if (row.vigDesde) row.vigDesde = serialToISO(row.vigDesde);
          if (row.vigHasta) row.vigHasta = serialToISO(row.vigHasta);
        });
      }
      if (key === 'feriados') {
        imported.forEach(row => {
          if (row.fecha) row.fecha = serialToISO(row.fecha);
        });
      }

      if (!TM[key]) TM[key] = [];
      TM[key].push(...imported);
      showTM(key);
      guardarTM();   // persiste a tablasMaestras inmediatamente (no depende de Guardar cambios)
      scheduleAutoSave();
      alert(`✓ ${imported.length} registros importados en ${tmLabel(key)}`);
    } catch(err) {
      alert('Error importando: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
  input.value = ''; // reset para permitir reimportar mismo archivo
}

// ════════════════════════════════════════════════════════════════════
// EXPORTAR TABLA MAESTRA A EXCEL
// ════════════════════════════════════════════════════════════════════
function exportarTM(key) {
  const data = TM[key];
  if (!data || !data.length) { alert('No hay datos para exportar.'); return; }
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, tmLabel(key));
  XLSX.writeFile(wb, `TM_${key}_${new Date().toISOString().slice(0,10)}.xlsx`);
}

// ════════════════════════════════════════════════════════════════════
// HELPER
// ════════════════════════════════════════════════════════════════════
function esc(v) {
  return String(v ?? '').replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;');
}

// ════════════════════════════════════════════════════════════════════
// BUSCAR CUOTAS EN TM.PLANES
// ════════════════════════════════════════════════════════════════════
// Devuelve el nº de cuotas según TM.planes para un plan/tarjeta/procesadora dado.
// Retorna null si no hay entrada en la tabla maestra (no modifica el valor del archivo).
function buscarCuotasEnTM(plan, tarjeta, procesadora) {
  if (!TM.planes || !TM.planes.length) return null;

  const normP = String(plan        || '').trim().toUpperCase();
  const normT = String(tarjeta     || '').trim().toUpperCase();
  const normR = String(procesadora || '').trim().toUpperCase();

  // Coincidencia con límite de palabra para evitar que "4 CUOTAS" coincida
  // como subcadena dentro de "14 CUOTAS" (sin este control "14".includes("4") = true).
  // Busca `needle` en `haystack` respetando inicio/fin de token (espacio o extremo).
  function wordMatch(haystack, needle) {
    if (!needle) return true;
    if (haystack === needle) return true;
    const esc = needle.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    return new RegExp('(?:^|\\s)' + esc + '(?:\\s|$)').test(haystack);
  }

  const candidatos = TM.planes.filter(p => {
    if (!p.cuotas) return false;
    const tmPlan = String(p.plan        || '').trim().toUpperCase();
    const tmTarj = String(p.tarjeta     || '').trim().toUpperCase();
    const tmProc = String(p.procesadora || '').trim().toUpperCase();

    // El nombre del plan debe coincidir respetando límites de palabra
    const matchPlan = !tmPlan || wordMatch(normP, tmPlan) || wordMatch(tmPlan, normP);
    // Tarjeta y procesadora son filtros opcionales
    const matchTarj = !tmTarj || normT.includes(tmTarj) || tmTarj.includes(normT);
    const matchProc = !tmProc || tmProc === normR;

    return matchPlan && matchTarj && matchProc;
  });

  if (!candidatos.length) return null;

  // Ordenar: 1º match exacto de plan, 2º sky-contiene-TM, 3º TM-contiene-sky
  // En igualdad, más campos específicos (tarjeta + procesadora) ganan
  candidatos.sort((a, b) => {
    const pA = String(a.plan || '').trim().toUpperCase();
    const pB = String(b.plan || '').trim().toUpperCase();
    const qA = !pA ? 0 : pA === normP ? 2 : wordMatch(normP, pA) ? 1 : 0;
    const qB = !pB ? 0 : pB === normP ? 2 : wordMatch(normP, pB) ? 1 : 0;
    if (qB !== qA) return qB - qA;
    const sA = [a.tarjeta, a.procesadora].filter(Boolean).length;
    const sB = [b.tarjeta, b.procesadora].filter(Boolean).length;
    return sB - sA;
  });

  const c = parseInt(candidatos[0].cuotas);
  return isNaN(c) ? null : c;
}

// Función para obtener tasa según operación
function buscarTasaEnTM(tarjeta, cuotas, comercio, procesadora) {
  if (!TM.tasas || !TM.tasas.length) return null;
  const hoy = new Date().toISOString().slice(0,10);
  // Buscar match más específico primero
  const candidatos = TM.tasas.filter(t => {
    const matchTarjeta  = !t.tarjeta  || t.tarjeta.toUpperCase()  === tarjeta?.toUpperCase();
    const matchCuotas   = !t.cuotas   || parseInt(t.cuotas)       === parseInt(cuotas);
    const matchComercio = !t.comercio || t.comercio               === String(comercio);
    const matchProc     = !t.procesadora || t.procesadora.toUpperCase() === procesadora?.toUpperCase();
    const vigente = (!t.vigDesde || t.vigDesde <= hoy) && (!t.vigHasta || t.vigHasta >= hoy);
    return matchTarjeta && matchCuotas && matchComercio && matchProc && vigente;
  });
  if (!candidatos.length) return null;
  // Priorizar el más específico (más campos completos)
  candidatos.sort((a,b) => {
    const scoreA = [a.tarjeta,a.cuotas,a.comercio,a.procesadora].filter(Boolean).length;
    const scoreB = [b.tarjeta,b.cuotas,b.comercio,b.procesadora].filter(Boolean).length;
    return scoreB - scoreA;
  });
  return candidatos[0];
}

function bindTMImport(key) {
  // ya están bindados inline en el HTML generado
}
