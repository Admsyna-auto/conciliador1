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
           'equivalencias','motivos','estados'].map(k => `
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
    tasas:'Tasas / Acuerdos', equivalencias:'Equiv. Comercio',
    motivos:'Motivos', estados:'Estados',
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
    case 'tarjetas':     body.innerHTML = renderTMGeneric(key, ['tarjeta','equivSkylab','equivProc'],
      ['Tarjeta','Equiv. Skylab','Equiv. Proc.'], () => ({tarjeta:'',equivSkylab:'',equivProc:''})); break;
    case 'planes':       body.innerHTML = renderTMGeneric(key, ['plan','cuotas','tarjeta','procesadora','codigos'],
      ['Plan','Cuotas','Tarjeta','Procesadora','Códigos'], () => ({plan:'',cuotas:'',tarjeta:'',procesadora:'',codigos:''})); break;
    case 'tasas':        body.innerHTML = renderTMTasas(); break;
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
      };

      const map = mapCols[key] || {};
      const imported = rows.map(r => {
        const obj = {};
        for (const [col, field] of Object.entries(map)) {
          // Buscar columna case-insensitive
          const found = Object.keys(r).find(k => k.toLowerCase().replace(/[\s.]/g,'') === col.replace(/[\s.]/g,''));
          if (found !== undefined && obj[field] === undefined) obj[field] = String(r[found] ?? '');
        }
        return obj;
      }).filter(r => Object.values(r).some(v => v !== ''));

      // Post-proceso para tasas: convertir tasa decimal→%, fechas serial→ISO
      if (key === 'tasas') {
        const serialToISO = v => {
          const n = parseFloat(v);
          if (isNaN(n) || n < 10000) return v;
          return new Date(Math.round((n - 25569) * 86400000)).toISOString().slice(0, 10);
        };
        imported.forEach(row => {
          if (row.tasa) {
            const n = parseFloat(row.tasa);
            if (!isNaN(n) && n > 0 && n < 1) row.tasa = (n * 100).toFixed(4).replace(/\.?0+$/, '');
          }
          if (row.vigDesde) row.vigDesde = serialToISO(row.vigDesde);
          if (row.vigHasta) row.vigHasta = serialToISO(row.vigHasta);
        });
      }

      if (!TM[key]) TM[key] = [];
      TM[key].push(...imported);
      showTM(key);
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
