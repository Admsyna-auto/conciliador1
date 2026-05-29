// ═══════════════════════════════════════════════════════════════════
// DIFERENCIAS.JS — Cálculo de diferencias económicas por tasa
// ═══════════════════════════════════════════════════════════════════

// ── Acción sugerida según magnitud y tipo de diferencia
function accionSugerida(difMonto, estado) {
  if (!difMonto || Math.abs(difMonto) < 100) return 'OK — Sin diferencia significativa';
  if (estado?.includes('INTEGRADO')) return 'OK — Operación integrada';
  if (difMonto > 0)  return difMonto > 5000 ? 'COBRAR diferencia al comercio' : 'REVISAR operación';
  if (difMonto < 0)  return 'ABSORBER — procesadora cobró menos de lo acordado';
  return 'REVISAR';
}

// ── Calcular diferencia para UNA fila conciliada
function calcularDiferencia(fila) {
  const s = fila.sky;
  const p = fila.proc;

  // No calcular si no está conciliada, es integrada o es devolución
  if (!p || s.integrado || s.esNeg) {
    fila.tasaCobrada    = null;
    fila.tasaAcordada   = null;
    fila.difTasa        = null;
    fila.difMonto       = null;
    fila.accionSugerida = s.integrado ? 'OK — Integrado' : null;
    return;
  }

  const importe = Math.abs(s.monto);
  if (!importe) return;

  // Tasa cobrada: si la procesadora informa arancel+cfo, calcularla
  // En el nuevo formato FISERV no vienen esos campos → null
  const tasaCob = (p.arancel != null && p.cfo != null)
    ? (p.arancel + p.cfo) / importe
    : null;

  // Tasa acordada: buscar en tabla de tasas
  const tmTasa = buscarTasaEnTM(
    s.tarjeta,
    s.cuotas || extraerCuotasDePlan(s.plan),
    p.comFis || s.nroCom,
    fila.procEncontrada
  );

  const tasaAco = tmTasa ? parseFloat(tmTasa.tasa) / 100 : null;

  fila.tasaCobrada  = tasaCob !== null ? +tasaCob.toFixed(6) : null;
  fila.tasaAcordada = tasaAco;
  fila.grupoTasa    = tmTasa ? `${tmTasa.acuerdo || tmTasa.tarjeta} ${tmTasa.cuotas}c` : null;

  if (tasaCob !== null && tasaAco !== null) {
    fila.difTasa  = +(tasaCob - tasaAco).toFixed(6);
    fila.difMonto = +(fila.difTasa * importe).toFixed(2);
  } else if (tasaAco !== null) {
    // Solo tenemos tasa acordada — la diferencia no se puede calcular con precisión
    fila.difTasa  = null;
    fila.difMonto = null;
  }

  fila.accionSugerida = accionSugerida(fila.difMonto, fila.estado);
}

function extraerCuotasDePlan(plan) {
  const s = String(plan || '').toUpperCase();
  const m = s.match(/(\d+)\s*CUOTA/);
  if (m) return parseInt(m[1]);
  if (s.includes('DEBITO') || s.includes('1 ')) return 1;
  return null;
}

// ── Calcular diferencias para todo el RESULTADO
function calcularTodasDiferencias() {
  for (const fila of RESULTADO) {
    calcularDiferencia(fila);
  }
}

// ════════════════════════════════════════════════════════════════════
// RENDER MÓDULO 3 — CÁLCULO DE DIFERENCIAS
// ════════════════════════════════════════════════════════════════════
function renderModuloDiferencias() {
  calcularTodasDiferencias();

  // Incluir COM.ERRADO solo si fueron marcadas como CON_DIF por el usuario
  const conciliadas = RESULTADO.filter(r => {
    if (!r.proc || r.sky.integrado || r.sky.esNeg) return false;
    if (r.estado === 'COM. ERRADO') {
      return (typeof COM_ERRADO_MARCAS !== 'undefined')
        ? COM_ERRADO_MARCAS[r.sky.idx] === 'CON_DIF'
        : false;
    }
    return ['OK','OK (equiv.)','CORREGIDO MANUAL (FISERV)','CORREGIDO MANUAL (GETPOS)']
      .some(e => r.estado?.startsWith(e.split('(')[0].trim()));
  });

  const conDif = conciliadas.filter(r => r.difMonto != null && Math.abs(r.difMonto) >= 100);
  const totalDif = conDif.reduce((s,r) => s + (r.difMonto||0), 0);
  const totalFav  = conDif.filter(r => (r.difMonto||0) < 0).reduce((s,r) => s + r.difMonto, 0);
  const totalContra = conDif.filter(r => (r.difMonto||0) > 0).reduce((s,r) => s + r.difMonto, 0);

  const cont = document.getElementById('mod-dif');
  if (!cont) return;

  cont.innerHTML = `
    <div class="dif-layout">
      <!-- KPIs -->
      <div class="dif-kpis">
        <div class="dif-kpi">
          <div class="dif-kpi-lbl">Operaciones analizadas</div>
          <div class="dif-kpi-val cyn">${conciliadas.length.toLocaleString()}</div>
        </div>
        <div class="dif-kpi">
          <div class="dif-kpi-lbl">Con diferencia ≥ $100</div>
          <div class="dif-kpi-val yel">${conDif.length.toLocaleString()}</div>
        </div>
        <div class="dif-kpi">
          <div class="dif-kpi-lbl">Total diferencia</div>
          <div class="dif-kpi-val ${totalDif > 0 ? 'red' : 'grn'}">${fmtARS(totalDif)}</div>
        </div>
        <div class="dif-kpi">
          <div class="dif-kpi-lbl">A favor empresa</div>
          <div class="dif-kpi-val grn">${fmtARS(Math.abs(totalFav))}</div>
        </div>
        <div class="dif-kpi">
          <div class="dif-kpi-lbl">En contra empresa</div>
          <div class="dif-kpi-val red">${fmtARS(totalContra)}</div>
        </div>
      </div>

      <!-- Filtros -->
      <div class="dif-filters">
        <span class="fix-filters-lbl">Filtrar</span>
        <select class="filter-sel" id="dflt-suc" onchange="renderTablaDif()">
          <option value="">Todas las sucursales</option>
          ${[...new Set(conciliadas.map(r=>r.sky.suc))].sort((a,b)=>+a-+b)
            .map(s=>`<option value="${s}">${s}</option>`).join('')}
        </select>
        <select class="filter-sel" id="dflt-proc" onchange="renderTablaDif()">
          <option value="">Ambas procesadoras</option>
          <option value="FISERV">FISERV</option>
          <option value="GETPOS">GETPOS</option>
        </select>
        <select class="filter-sel" id="dflt-accion" onchange="renderTablaDif()">
          <option value="">Todas las acciones</option>
          <option value="COBRAR">COBRAR</option>
          <option value="ABSORBER">ABSORBER</option>
          <option value="REVISAR">REVISAR</option>
        </select>
        <label style="display:flex;align-items:center;gap:6px;font-size:10px;color:var(--m1)">
          <input type="checkbox" id="dflt-solo-dif" onchange="renderTablaDif()" checked>
          Solo con diferencia
        </label>
        <button class="fix-clear" onclick="exportarDiferencias()">📤 Exportar Excel</button>
        <button class="fix-clear" onclick="exportarResumenes()">📊 Resúmenes</button>
      </div>

      <!-- Tabla -->
      <div class="tbl-wrap" id="dif-tbl-wrap">
        <table id="tbl-dif"><thead></thead><tbody></tbody></table>
      </div>
    </div>`;

  renderTablaDif();

  // Mostrar tab strip y actualizar todos los contadores
  const strip = document.getElementById('tab-strip-dif');
  if (strip) strip.style.display = 'flex';
  const cntFis = document.getElementById('cnt-dif-fis');
  if (cntFis) cntFis.textContent = (window._FIS_NO_CRUZADAS || []).length;
  const cntGp = document.getElementById('cnt-dif-gp');
  if (cntGp) cntGp.textContent = (window._GP_NO_CRUZADAS || []).length;

  // Contadores nuevas tabs
  const cntCuotas = document.getElementById('cnt-dif-cuotas');
  if (cntCuotas) cntCuotas.textContent = _getDifCuotasRows().length;
  const cntProc = document.getElementById('cnt-dif-proc');
  if (cntProc) cntProc.textContent = _getDifProcRows().length;
}

function renderTablaDif() {
  const filSuc    = document.getElementById('dflt-suc')?.value    || '';
  const filProc   = document.getElementById('dflt-proc')?.value   || '';
  const filAccion = document.getElementById('dflt-accion')?.value || '';
  const soloDif   = document.getElementById('dflt-solo-dif')?.checked ?? true;

  let rows = RESULTADO.filter(r => {
    if (!r.proc || r.sky.integrado || r.sky.esNeg) return false;
    if (r.estado === 'COM. ERRADO') {
      return (typeof COM_ERRADO_MARCAS !== 'undefined')
        ? COM_ERRADO_MARCAS[r.sky.idx] === 'CON_DIF'
        : false;
    }
    return true;
  });

  if (filSuc)    rows = rows.filter(r => r.sky.suc === filSuc);
  if (filProc)   rows = rows.filter(r => r.procEncontrada === filProc);
  if (filAccion) rows = rows.filter(r => r.accionSugerida?.toUpperCase().includes(filAccion));
  // "Solo con diferencia" filtra por difTasa calculada,
  // pero COM.ERRADO CON_DIF siempre aparecen (todavía sin tasa cargada)
  if (soloDif) rows = rows.filter(r =>
    (r.estado === 'COM. ERRADO' && COM_ERRADO_MARCAS?.[r.sky.idx] === 'CON_DIF') ||
    (r.difTasa != null && Math.abs(r.difTasa) > 0)
  );

  const HDR_DIF = ['Fecha','Suc','Vendedor','Tarjeta','Plan','Cuotas','Monto','Proc.',
    'Com. FIS','Tasa Cobrada','Tasa Acordada','Grupo','DIF %','DIF $','Acción'];
  const tbl = document.getElementById('tbl-dif');
  if (!tbl) return;

  tbl.querySelector('thead').innerHTML = `<tr>${HDR_DIF.map(h=>`<th>${h}</th>`).join('')}</tr>`;
  tbl.querySelector('tbody').innerHTML = rows.slice(0, 3000).map(r => {
    const s = r.sky, p = r.proc;
    // COM.ERRADO CON_DIF: resaltar como pendiente de tasa si no hay difTasa aún
    const esComErrCON = r.estado === 'COM. ERRADO' &&
      (typeof COM_ERRADO_MARCAS !== 'undefined') &&
      COM_ERRADO_MARCAS[r.sky.idx] === 'CON_DIF';
    const difClass = r.difTasa > 0 ? 'num-pos' : r.difTasa < 0 ? 'num-neg' : '';
    const accion = esComErrCON && !r.tasaAcordada
      ? '⏳ Pendiente tasa'
      : (r.accionSugerida || '—');
    const accionClass = accion.includes('COBRAR') ? 'badge-red' :
                        accion.includes('ABSORBER') ? 'badge-grn' :
                        accion.includes('REVISAR') ? 'badge-yel' :
                        accion.includes('Pendiente') ? 'badge-org' : '';
    return `<tr>
      <td>${s.fecha}</td>
      <td>${s.suc}</td>
      <td style="max-width:140px;overflow:hidden;text-overflow:ellipsis">${s.vendedor??'—'}</td>
      <td>${s.tarjeta}</td>
      <td>${s.plan}</td>
      <td class="num">${s.cuotas??'—'}</td>
      <td class="num">$${Math.abs(s.monto).toLocaleString('es-AR',{minimumFractionDigits:2})}</td>
      <td><span class="st ${r.procEncontrada==='FISERV'?'st-fis':'st-gp'}">${r.procEncontrada}</span></td>
      <td class="num" style="font-size:9px">${p?.comFis??'—'}</td>
      <td class="num">${r.tasaCobrada!=null ? (r.tasaCobrada*100).toFixed(4)+'%' : '—'}</td>
      <td class="num">${r.tasaAcordada!=null ? (r.tasaAcordada*100).toFixed(4)+'%' : '—'}</td>
      <td style="font-size:9px;color:var(--m1)">${r.grupoTasa??'—'}</td>
      <td class="num ${difClass}">${r.difTasa!=null ? (r.difTasa*100).toFixed(4)+'%' : '—'}</td>
      <td class="num ${difClass}">${r.difMonto!=null ? fmtARS(r.difMonto) : '—'}</td>
      <td><span class="badge-accion ${accionClass}">${accion}</span></td>
    </tr>`;
  }).join('');
}

// ════════════════════════════════════════════════════════════════════
// EXPORTACIONES
// ════════════════════════════════════════════════════════════════════

// Export 1 — Diferencias finales
function exportarDiferencias() {
  calcularTodasDiferencias();
  const rows = RESULTADO.filter(r => r.proc && !r.sky.integrado);
  const HDR = ['Fecha','Sucursal','Vendedor','Asiento','Procesadora','Tarjeta','Plan','Cuotas',
    'Monto Facturado','Monto Procesadora','Com. SKY','Com. FIS','Tasa Cobrada','Tasa Acordada',
    'Grupo Tasa','DIF %','DIF $','Acción Sugerida','Estado Conciliación','Método'];
  const data = rows.map(r => {
    const s = r.sky, p = r.proc;
    return [
      s.fecha, s.suc, s.vendedor??'', s.asiento??'',
      r.procEncontrada??'', s.tarjeta, s.plan, s.cuotas??'',
      s.monto, p?.monto??'', s.nroCom, p?.comFis??'',
      r.tasaCobrada!=null?(r.tasaCobrada*100).toFixed(4)+'%':'',
      r.tasaAcordada!=null?(r.tasaAcordada*100).toFixed(4)+'%':'',
      r.grupoTasa??'', r.difTasa??'', r.difMonto??'',
      r.accionSugerida??'', r.estado??'', r.metodo??'',
    ];
  });
  _exportXlsx([HDR, ...data], 'Diferencias', `Diferencias_${hoy()}.xlsx`);
}

// Export 2 — Log de correcciones
function exportarLogCorrecciones() {
  const HDR = ['Fecha/Hora','Usuario','Asiento','Sucursal','Tarjeta','Monto',
    'Campo','Valor Original','Valor Corregido','Motivo','Observaciones','Estado Sistema'];
  const data = LOG_AUDIT.map(e => [
    e.ts, e.usuario, e.asiento, e.suc, e.tarjeta, e.monto,
    e.campo, e.valorAntes, e.valorDespues, e.motivo, e.obs, e.estadoSistema,
  ]);
  _exportXlsx([HDR, ...data], 'Log Correcciones', `LogCorrecciones_${hoy()}.xlsx`);
}

// Export 3 — No conciliadas
function exportarNoConciliadas() {
  const rows = RESULTADO.filter(r => r.estado === 'SIN MATCH');
  const HDR = ['Fecha','Sucursal','Vendedor','Asiento','Tarjeta','Plan','Cuotas',
    'Monto','Cupón','Lote','Nro. Comercio','Procesadora Esperada'];
  const data = rows.map(r => {
    const s = r.sky;
    return [s.fecha, s.suc, s.vendedor??'', s.asiento??'',
      s.tarjeta, s.plan, s.cuotas??'', s.monto,
      s.cupon, s.lote, s.nroCom, r.procEsperada??''];
  });
  _exportXlsx([HDR, ...data], 'No Conciliadas', `NoConciliadas_${hoy()}.xlsx`);
}

// Export 4 — Resúmenes
function exportarResumenes() {
  calcularTodasDiferencias();
  const wb = XLSX.utils.book_new();

  // Por sucursal
  const porSuc = {};
  for (const r of RESULTADO) {
    const k = r.sky.suc || '—';
    if (!porSuc[k]) porSuc[k] = { suc:k, total:0, conciliadas:0, sinMatch:0, dif:0 };
    porSuc[k].total++;
    if (r.estado !== 'SIN MATCH') porSuc[k].conciliadas++;
    else porSuc[k].sinMatch++;
    porSuc[k].dif += r.difMonto || 0;
  }
  const dataSuc = Object.values(porSuc).sort((a,b) => Math.abs(b.dif)-Math.abs(a.dif))
    .map(r => [r.suc, r.total, r.conciliadas, r.sinMatch, +r.dif.toFixed(2)]);
  const wsSuc = XLSX.utils.aoa_to_sheet([['Sucursal','Total','Conciliadas','Sin Match','Dif $'],...dataSuc]);
  XLSX.utils.book_append_sheet(wb, wsSuc, 'Por Sucursal');

  // Por vendedor
  const porVend = {};
  for (const r of RESULTADO) {
    const k = r.sky.vendedor || '—';
    if (!porVend[k]) porVend[k] = { vend:k, suc:r.sky.suc, total:0, dif:0 };
    porVend[k].total++;
    porVend[k].dif += r.difMonto || 0;
  }
  const dataVend = Object.values(porVend).sort((a,b) => Math.abs(b.dif)-Math.abs(a.dif)).slice(0,200)
    .map(r => [r.vend, r.suc, r.total, +r.dif.toFixed(2)]);
  const wsVend = XLSX.utils.aoa_to_sheet([['Vendedor','Sucursal','Total','Dif $'],...dataVend]);
  XLSX.utils.book_append_sheet(wb, wsVend, 'Por Vendedor');

  // Por procesadora
  const porProc = {};
  for (const r of RESULTADO) {
    const k = r.procEncontrada || 'Sin match';
    if (!porProc[k]) porProc[k] = { proc:k, total:0, conDif:0, dif:0 };
    porProc[k].total++;
    if (r.difMonto && Math.abs(r.difMonto) >= 100) { porProc[k].conDif++; porProc[k].dif += r.difMonto; }
  }
  const dataProc = Object.values(porProc).map(r => [r.proc, r.total, r.conDif, +r.dif.toFixed(2)]);
  const wsProc = XLSX.utils.aoa_to_sheet([['Procesadora','Total','Con Diferencia','Dif $'],...dataProc]);
  XLSX.utils.book_append_sheet(wb, wsProc, 'Por Procesadora');

  // Por motivo (correcciones)
  const porMotivo = {};
  for (const e of LOG_AUDIT) {
    const k = e.motivo || 'Sin motivo';
    if (!porMotivo[k]) porMotivo[k] = { motivo:k, count:0 };
    porMotivo[k].count++;
  }
  const dataMotivo = Object.values(porMotivo).sort((a,b) => b.count-a.count)
    .map(r => [r.motivo, r.count]);
  const wsMotivo = XLSX.utils.aoa_to_sheet([['Motivo','Cantidad'],...dataMotivo]);
  XLSX.utils.book_append_sheet(wb, wsMotivo, 'Por Motivo');

  XLSX.writeFile(wb, `Resumenes_Conciliacion_${hoy()}.xlsx`);
}

function _exportXlsx(data, sheetName, filename) {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, filename);
}

// ════════════════════════════════════════════════════════════════════
// DIF. CUOTAS — operaciones cruzadas con cuotas distintas
// ════════════════════════════════════════════════════════════════════

function _getDifCuotasRows() {
  // Compatibilidad con sesiones guardadas antes de esta feature:
  // recalcular difCuotas on-the-fly si no está en la fila
  return RESULTADO.filter(r => {
    if (!r.proc) return false;
    if (r.difCuotas !== undefined) return r.difCuotas;
    const sc = Math.max(1, parseInt(r.sky?.cuotas) || 1);
    const pc = Math.max(1, parseInt(r.proc?.cuotas) || 1);
    return sc !== pc;
  });
}

function renderDifCuotas() {
  const rows  = _getDifCuotasRows();
  const tbl   = document.getElementById('tbl-dif-cuotas');
  const stats = document.getElementById('dif-cuotas-stats');
  const cnt   = document.getElementById('cnt-dif-cuotas');
  if (cnt)   cnt.textContent = rows.length;
  if (stats) stats.innerHTML =
    `<b>${rows.length}</b> operación${rows.length!==1?'es':''} con ` +
    `<b style="color:var(--yel)">cuotas diferentes</b> entre Skylab y procesadora`;
  if (!tbl) return;

  const HDR = ['Estado','Fecha SKY','Suc.','Vendedor','Tarjeta','Plan','Cuotas SKY','Cuotas Proc.','Dif.','Monto SKY','Procesadora','Cód.Auth. Proc.','Lote Proc.','Cupón SKY'];
  tbl.querySelector('thead').innerHTML = `<tr>${HDR.map(h=>`<th>${h}</th>`).join('')}</tr>`;
  tbl.querySelector('tbody').innerHTML = rows.map(r => {
    const s  = r.sky, p = r.proc;
    // Recalcular on-the-fly con TM si el dato no estaba serializado (sesiones anteriores)
    let sc = r.skyCuotas;
    let fromTM = r.skyCuotasTM ?? false;
    if (sc === undefined) {
      const tmC = (typeof buscarCuotasEnTM === 'function')
        ? buscarCuotasEnTM(s.plan, s.tarjeta, r.procEncontrada) : null;
      sc = tmC !== null ? tmC : Math.max(1, parseInt(s.cuotas) || 1);
      fromTM = tmC !== null;
    }
    const pc = r.procCuotas ?? Math.max(1, parseInt(p?.cuotas) || 1);
    const dif = sc - pc;
    const difStr = dif > 0 ? `+${dif}` : String(dif);
    // Indicador de fuente
    const tmBadge = fromTM
      ? `<span title="Valor tomado de TM Planes/Cuotas" style="font-size:7px;background:rgba(79,142,247,.15);color:var(--acc);border:1px solid rgba(79,142,247,.3);border-radius:2px;padding:1px 3px;margin-left:3px">TM</span>`
      : `<span title="Valor extraído del archivo Skylab" style="font-size:7px;background:rgba(107,114,128,.1);color:var(--m2);border:1px solid rgba(107,114,128,.2);border-radius:2px;padding:1px 3px;margin-left:3px">Arch.</span>`;
    return `<tr class="${rowClass(r.estado)}">
      <td>${estadoBadge(r.estado)}</td>
      <td>${s.fecha}</td>
      <td>${s.suc}</td>
      <td style="max-width:130px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${s.vendedor||''}">${s.vendedor||'—'}</td>
      <td>${s.tarjeta}</td>
      <td style="font-size:9px;max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${s.plan||''}">${s.plan||'—'}</td>
      <td class="num" style="color:var(--acc);font-weight:700">${sc}${tmBadge}</td>
      <td class="num" style="color:var(--org);font-weight:700">${pc}</td>
      <td class="num" style="color:var(--red);font-weight:700">${difStr} cuota${Math.abs(dif)!==1?'s':''}</td>
      <td class="num">${fmtARS(s.monto)}</td>
      <td><span class="st ${r.procEncontrada==='FISERV'?'st-fis':'st-gp'}">${r.procEncontrada||'—'}</span></td>
      <td class="num" style="font-size:9px">${p?.aut||'—'}</td>
      <td class="num" style="font-size:9px">${p?.lote||'—'}</td>
      <td class="num" style="font-size:9px">${s.cupon||'—'}</td>
    </tr>`;
  }).join('');
}

function exportarDifCuotas() {
  const rows = _getDifCuotasRows();
  if (!rows.length) { alert('No hay operaciones con diferencia de cuotas.'); return; }
  const HDR = ['Estado','Fecha SKY','Suc.','Vendedor','Tarjeta','Plan',
    'Cuotas SKY','Cuotas Proc.','Dif. Cuotas','Monto SKY',
    'Procesadora','Cód.Auth. Proc.','Lote Proc.','Ticket Proc.','Cupón SKY','Asiento'];
  const data = rows.map(r => {
    const s = r.sky, p = r.proc;
    const sc = r.skyCuotas  ?? Math.max(1, parseInt(s.cuotas)  || 1);
    const pc = r.procCuotas ?? Math.max(1, parseInt(p?.cuotas) || 1);
    return [r.estado, s.fecha, s.suc, s.vendedor||'', s.tarjeta, s.plan,
      sc, pc, sc - pc, s.monto,
      r.procEncontrada||'', p?.aut||'', p?.lote||'', p?.ticket||'', s.cupon||'', s.asiento||''];
  });
  _exportXlsx([HDR, ...data], 'Dif. Cuotas', `DifCuotas_${hoy()}.xlsx`);
}

// ════════════════════════════════════════════════════════════════════
// DIF. PROCESADORA — facturado por una proc, cobrado por otra
// ════════════════════════════════════════════════════════════════════

function _getDifProcRows() {
  return RESULTADO.filter(r => r.estado?.startsWith('MAL FACTURADO') && r.proc);
}

function renderDifProcesadora() {
  const rows  = _getDifProcRows();
  const tbl   = document.getElementById('tbl-dif-proc');
  const stats = document.getElementById('dif-proc-stats');
  const cnt   = document.getElementById('cnt-dif-proc');
  if (cnt)   cnt.textContent = rows.length;
  if (stats) stats.innerHTML =
    `<b>${rows.length}</b> operación${rows.length!==1?'es':''} ` +
    `<b style="color:var(--red)">facturadas por una procesadora pero cobradas por otra</b>`;
  if (!tbl) return;

  const HDR = ['Estado','Fecha SKY','Suc.','Vendedor','Tarjeta','Plan','Cuotas','Monto SKY',
    'Proc. Esperada','Proc. Real','Com. SKY','Com. FIS','Cód.Auth. Proc.','Lote Proc.','Cupón SKY'];
  tbl.querySelector('thead').innerHTML = `<tr>${HDR.map(h=>`<th>${h}</th>`).join('')}</tr>`;
  tbl.querySelector('tbody').innerHTML = rows.map(r => {
    const s = r.sky, p = r.proc;
    const espCls = r.procEsperada==='FISERV'?'st-fis':'st-gp';
    const realCls= r.procEncontrada==='FISERV'?'st-fis':'st-gp';
    return `<tr class="${rowClass(r.estado)}">
      <td>${estadoBadge(r.estado)}</td>
      <td>${s.fecha}</td>
      <td>${s.suc}</td>
      <td style="max-width:130px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${s.vendedor||''}">${s.vendedor||'—'}</td>
      <td>${s.tarjeta}</td>
      <td style="font-size:9px">${s.plan}</td>
      <td class="num">${s.cuotas||1}</td>
      <td class="num">${fmtARS(s.monto)}</td>
      <td><span class="st ${espCls}">${r.procEsperada||'—'}</span></td>
      <td><span class="st ${realCls}" style="outline:1px solid var(--red)">${r.procEncontrada||'—'}</span></td>
      <td class="num" style="font-size:9px">${s.nroCom||'—'}</td>
      <td class="num" style="font-size:9px">${p?.comFis||'—'}</td>
      <td class="num" style="font-size:9px">${p?.aut||'—'}</td>
      <td class="num" style="font-size:9px">${p?.lote||'—'}</td>
      <td class="num" style="font-size:9px">${s.cupon||'—'}</td>
    </tr>`;
  }).join('');
}

function exportarDifProcesadora() {
  const rows = _getDifProcRows();
  if (!rows.length) { alert('No hay operaciones con diferencia de procesadora.'); return; }
  const HDR = ['Estado','Fecha SKY','Suc.','Vendedor','Tarjeta','Plan','Cuotas','Monto SKY',
    'Proc. Esperada','Proc. Real','Com. SKY','Com. FIS',
    'Cód.Auth. Proc.','Lote Proc.','Ticket Proc.','Cupón SKY','Asiento'];
  const data = rows.map(r => {
    const s = r.sky, p = r.proc;
    return [r.estado, s.fecha, s.suc, s.vendedor||'', s.tarjeta, s.plan,
      s.cuotas||1, s.monto,
      r.procEsperada||'', r.procEncontrada||'',
      s.nroCom||'', p?.comFis||'',
      p?.aut||'', p?.lote||'', p?.ticket||'', s.cupon||'', s.asiento||''];
  });
  _exportXlsx([HDR, ...data], 'Dif. Procesadora', `DifProcesadora_${hoy()}.xlsx`);
}

function hoy() { return new Date().toISOString().slice(0,10); }

// ════════════════════════════════════════════════════════════════════
// FISERV Y GETPOS SIN CRUCE CON SKYLAB
// ════════════════════════════════════════════════════════════════════
const HDR_FIS_NC = ['Fecha','Terminal','Suc.','Tarjeta','Cuotas','Monto','Lote','Ticket','Aut.','Cód.Comercio'];
const HDR_GP_NC  = ['Fecha','Nombre Establecimiento','Suc.','Marca','Plan','Monto','Cód.Aut.','Nro.Cupón'];

function renderNoCruzadasFis() {
  const rows  = window._FIS_NO_CRUZADAS || [];
  const tbl   = document.getElementById('tbl-dif-fis');
  const stats = document.getElementById('dif-fis-stats');
  const cnt   = document.getElementById('cnt-dif-fis');

  if (cnt)   cnt.textContent = rows.length;
  if (stats) stats.innerHTML =
    `<b>${rows.length}</b> operaciones FISERV <b style="color:var(--red)">sin correspondencia en Skylab</b>`;

  // Mostrar el tab-strip
  const strip = document.getElementById('tab-strip-dif');
  if (strip) strip.style.display = 'flex';

  if (!tbl) return;
  if (!rows.length) {
    tbl.outerHTML = `<table id="tbl-dif-fis"><thead></thead><tbody><tr>
      <td colspan="${HDR_FIS_NC.length}" style="padding:30px;text-align:center;color:var(--m2);font-size:10px">
        Todas las operaciones FISERV tienen correspondencia en Skylab.</td></tr></tbody></table>`;
    return;
  }
  tbl.querySelector('thead').innerHTML =
    `<tr>${HDR_FIS_NC.map(h => `<th>${h}</th>`).join('')}</tr>`;
  tbl.querySelector('tbody').innerHTML = rows.map(r => `<tr>
    <td>${r.fecha||'—'}</td>
    <td style="font-size:9px">${r.equipo||'—'}</td>
    <td>${r.suc||'—'}</td>
    <td>${r.tarjeta||'—'}</td>
    <td class="num">${r.cuotas||'—'}</td>
    <td class="num" style="color:var(--org)">$${Math.abs(r.monto||0).toLocaleString('es-AR',{minimumFractionDigits:2})}</td>
    <td>${r.lote||'—'}</td>
    <td>${r.ticket||'—'}</td>
    <td style="font-size:9px">${r.aut||'—'}</td>
    <td style="font-size:9px">${r.comFis||'—'}</td>
  </tr>`).join('');
}

function renderNoCruzadasGp() {
  const rows  = window._GP_NO_CRUZADAS || [];
  const tbl   = document.getElementById('tbl-dif-gp');
  const stats = document.getElementById('dif-gp-stats');
  const cnt   = document.getElementById('cnt-dif-gp');

  if (cnt)   cnt.textContent = rows.length;
  if (stats) stats.innerHTML =
    `<b>${rows.length}</b> operaciones GETPOS <b style="color:var(--red)">sin correspondencia en Skylab</b>`;

  const strip = document.getElementById('tab-strip-dif');
  if (strip) strip.style.display = 'flex';

  if (!tbl) return;
  if (!rows.length) {
    tbl.outerHTML = `<table id="tbl-dif-gp"><thead></thead><tbody><tr>
      <td colspan="${HDR_GP_NC.length}" style="padding:30px;text-align:center;color:var(--m2);font-size:10px">
        Todas las operaciones GETPOS tienen correspondencia en Skylab.</td></tr></tbody></table>`;
    return;
  }
  tbl.querySelector('thead').innerHTML =
    `<tr>${HDR_GP_NC.map(h => `<th>${h}</th>`).join('')}</tr>`;
  tbl.querySelector('tbody').innerHTML = rows.map(r => `<tr>
    <td>${r.fecha||'—'}</td>
    <td style="max-width:160px;overflow:hidden;text-overflow:ellipsis">${r.nombre||'—'}</td>
    <td>${r.suc||'—'}</td>
    <td>${r.marca||r.tarjeta||'—'}</td>
    <td>${r.plan||'—'}</td>
    <td class="num" style="color:var(--org)">$${Math.abs(r.monto||0).toLocaleString('es-AR',{minimumFractionDigits:2})}</td>
    <td>${r.aut||'—'}</td>
    <td>${r.cupon||'—'}</td>
  </tr>`).join('');
}

function exportarNoCruzadasFis() {
  const rows = window._FIS_NO_CRUZADAS || [];
  if (!rows.length) { alert('No hay operaciones FISERV sin cruce.'); return; }
  const data = rows.map(r => [r.fecha||'', r.equipo||'', r.suc||'', r.tarjeta||'',
    r.cuotas||'', r.monto||0, r.lote||'', r.ticket||'', r.aut||'', r.comFis||'']);
  _exportXlsx([HDR_FIS_NC, ...data], 'FISERV Sin Cruce', `FISERV_SinCruce_${hoy()}.xlsx`);
}

function exportarNoCruzadasGp() {
  const rows = window._GP_NO_CRUZADAS || [];
  if (!rows.length) { alert('No hay operaciones GETPOS sin cruce.'); return; }
  const data = rows.map(r => [r.fecha||'', r.nombre||'', r.suc||'', r.marca||r.tarjeta||'',
    r.plan||'', r.monto||0, r.aut||'', r.cupon||'']);
  _exportXlsx([HDR_GP_NC, ...data], 'GETPOS Sin Cruce', `GETPOS_SinCruce_${hoy()}.xlsx`);
}
