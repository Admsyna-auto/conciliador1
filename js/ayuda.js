// ════════════════════════════════════════════════════════════════════
// GUÍA DE USO — Conciliador Skylab
// ════════════════════════════════════════════════════════════════════

const _GUIA_TABS = [
  { id: 'objetivo',       label: 'Objetivo'        },
  { id: 'flujo',          label: 'Flujo de trabajo' },
  { id: 'cruce',          label: 'Cruce automático' },
  { id: 'revision',       label: 'Revisión manual'  },
  { id: 'diferencias',    label: 'Diferencias'      },
  { id: 'liquidaciones',  label: 'Liquidaciones'    },
  { id: 'historial',      label: 'Historial'        },
  { id: 'dashboard',      label: 'Dashboard'        },
  { id: 'tm',             label: 'Config / TM'      },
];

const _GUIA_CONTENIDO = {

objetivo: `
<h3>¿Para qué sirve el Conciliador Skylab?</h3>
<p>El <b>Conciliador Skylab</b> es una herramienta de conciliación de operaciones con tarjeta de crédito y débito.
Compara lo que registra el sistema interno de ventas <b>Skylab</b> contra los archivos que reportan las <b>procesadoras</b>
(Fiserv, Getpos, Go Cuotas) y detecta automáticamente cualquier diferencia.</p>

<div class="gy-bloque">
  <div class="gy-bloque-titulo">Problemas que detecta</div>
  <ul>
    <li><b>SIN MATCH</b> — operaciones de Skylab que la procesadora no reconoce (posibles cobros perdidos).</li>
    <li><b>MAL FACTURADO</b> — el monto que cobró la procesadora difiere del que registró Skylab.</li>
    <li><b>COM. ERRADO</b> — la comisión aplicada no coincide con lo pactado.</li>
    <li><b>DIF. CUOTAS</b> — plan de cuotas distinto entre Skylab y la procesadora.</li>
    <li><b>ANULACIÓN SIN COBRO</b> — Skylab registró una anulación pero la procesadora igualmente acreditó el monto.</li>
    <li><b>Diferencias de tasa</b> — la tasa efectiva cobrada difiere de la tasa pactada en Tablas Maestras.</li>
    <li><b>Sin liquidar</b> — operaciones conciliadas que aún no fueron acreditadas en cuenta bancaria.</li>
  </ul>
</div>

<div class="gy-bloque">
  <div class="gy-bloque-titulo">Procesadoras soportadas</div>
  <div style="display:flex;gap:10px;flex-wrap:wrap;margin-top:6px">
    <span class="gy-badge gy-badge-fis">Fiserv</span>
    <span class="gy-badge gy-badge-gp">Getpos</span>
    <span class="gy-badge gy-badge-goc">Go Cuotas</span>
  </div>
  <p style="margin-top:8px;font-size:10px;color:var(--m2)">Cada procesadora se habilita/deshabilita individualmente según los archivos disponibles en el período.</p>
</div>

<div class="gy-bloque">
  <div class="gy-bloque-titulo">Todo funciona 100% offline</div>
  <p>No hay servidor. Los archivos se procesan en el navegador. Los datos del período se guardan en
  <b>IndexedDB</b> (base de datos local del navegador) y persisten entre sesiones en el mismo dispositivo.</p>
</div>
`,

// ─────────────────────────────────────────────────────────────────────
flujo: `
<h3>Flujo de trabajo recomendado</h3>
<p>Seguí estos pasos en orden cada vez que cerrás un período.</p>

<div class="gy-steps">
  <div class="gy-step">
    <div class="gy-step-num">1</div>
    <div class="gy-step-body">
      <b>Configurar Tablas Maestras</b> (primera vez y cuando haya cambios)<br>
      <span>Cargá plazos de acreditación, feriados, tasas pactadas y equivalencias de tarjetas.
      Se guardan en IndexedDB y se reutilizan en todos los períodos.</span>
    </div>
  </div>
  <div class="gy-step">
    <div class="gy-step-num">2</div>
    <div class="gy-step-body">
      <b>Cargar archivos del período</b><br>
      <span>Subí: Skylab, archivos de cada procesadora habilitada, Terminales (mapeo sucursal) y, si lo tenés, el archivo de Liquidaciones.</span>
    </div>
  </div>
  <div class="gy-step">
    <div class="gy-step-num">3</div>
    <div class="gy-step-body">
      <b>Ejecutar Cruce Automático</b> (botón ▶ Conciliar)<br>
      <span>El sistema intenta hacer match de cada operación Skylab con la procesadora correspondiente.
      Ver resultados en las pestañas: Completo, Sin Match, Mal Facturadas, etc.</span>
    </div>
  </div>
  <div class="gy-step">
    <div class="gy-step-num">4</div>
    <div class="gy-step-body">
      <b>Revisión Manual</b><br>
      <span>Para los SIN MATCH que en realidad tienen una razón conocida, aplicá correcciones manuales con el motivo correspondiente.
      Estas correcciones se arrastran automáticamente al próximo período.</span>
    </div>
  </div>
  <div class="gy-step">
    <div class="gy-step-num">5</div>
    <div class="gy-step-body">
      <b>Revisar Diferencias</b> (módulo DIFERENCIAS)<br>
      <span>Analizá diferencias de tasa, cuotas, procesadora y monto real. Clasificá cada diferencia de tasa (vendedor, procesadora, ok) para el dashboard.</span>
    </div>
  </div>
  <div class="gy-step">
    <div class="gy-step-num">6</div>
    <div class="gy-step-body">
      <b>Verificar Liquidaciones y Cobros</b><br>
      <span>Revisá qué operaciones ya fueron acreditadas, cuáles están pendientes y si hay extras.
      Las operaciones sin liquidar pasan automáticamente al próximo período.</span>
    </div>
  </div>
  <div class="gy-step">
    <div class="gy-step-num">7</div>
    <div class="gy-step-body">
      <b>Revisar el Dashboard</b><br>
      <span>Verificá los KPIs de cierre: % OK, montos no cobrados, diferencias de tasa pendientes, evolución histórica.</span>
    </div>
  </div>
  <div class="gy-step">
    <div class="gy-step-num">8</div>
    <div class="gy-step-body">
      <b>Cerrar Período</b> (botón en el historial)<br>
      <span>Guarda todo en IndexedDB y descarga un backup JSON. En el próximo período, el arrastre (pendientes + correcciones) se aplica automáticamente.</span>
    </div>
  </div>
</div>
`,

// ─────────────────────────────────────────────────────────────────────
cruce: `
<h3>Cruce Automático</h3>
<p>El cruce es el motor central del sistema. Compara cada fila de Skylab contra las procesadoras y asigna un estado.</p>

<div class="gy-bloque">
  <div class="gy-bloque-titulo">Archivos necesarios</div>
  <table class="gy-tabla">
    <tr><th>Archivo</th><th>Formato</th><th>Obligatorio</th><th>Para qué</th></tr>
    <tr><td>Skylab</td><td>.xlsx</td><td>Sí</td><td>Fuente principal de operaciones</td></tr>
    <tr><td>Terminales</td><td>.xlsx</td><td>Sí</td><td>Mapea terminal → sucursal</td></tr>
    <tr><td>Fiserv</td><td>.xlsx</td><td>Si habilitado</td><td>Detalle de operaciones Fiserv</td></tr>
    <tr><td>Getpos</td><td>.xlsx</td><td>Si habilitado</td><td>Detalle de operaciones Getpos</td></tr>
    <tr><td>Go Cuotas</td><td>Portal web / Excel</td><td>Si habilitado</td><td>Pagos y celular Go Cuotas</td></tr>
    <tr><td>Liquidaciones</td><td>.xlsx</td><td>No</td><td>Módulo de cobros y arrastre</td></tr>
  </table>
</div>

<div class="gy-bloque">
  <div class="gy-bloque-titulo">Estados del resultado</div>
  <table class="gy-tabla">
    <tr><th>Estado</th><th>Significado</th></tr>
    <tr><td><span class="gy-estado gy-ok">OK</span></td><td>Match exacto entre Skylab y procesadora</td></tr>
    <tr><td><span class="gy-estado gy-ok">OK (equiv.)</span></td><td>Match usando equivalencia de tarjeta</td></tr>
    <tr><td><span class="gy-estado gy-sin">SIN MATCH</span></td><td>Skylab tiene la operación pero la procesadora no</td></tr>
    <tr><td><span class="gy-estado gy-mal">MAL FACTURADO</span></td><td>Montos difieren entre Skylab y procesadora</td></tr>
    <tr><td><span class="gy-estado gy-com">COM. ERRADO</span></td><td>Comisión aplicada incorrecta</td></tr>
    <tr><td><span class="gy-estado gy-dif">DIF. CUOTAS</span></td><td>Plan de cuotas distinto</td></tr>
    <tr><td><span class="gy-estado gy-urg">REV. URGENTE</span></td><td>Requiere revisión inmediata</td></tr>
    <tr><td><span class="gy-estado gy-ref">REFACTURADO</span></td><td>Operación refacturada en Skylab</td></tr>
    <tr><td><span class="gy-estado gy-dev">DEVOLUCIÓN</span></td><td>Operación negativa / reverso</td></tr>
    <tr><td><span class="gy-estado gy-anul">ANUL.S/COBRO</span></td><td>Anulada en Skylab pero cobrada por procesadora</td></tr>
  </table>
</div>

<div class="gy-bloque">
  <div class="gy-bloque-titulo">Lógica de matching</div>
  <ol>
    <li>Intenta match por <b>cupón + lote + monto + sucursal</b> (exacto).</li>
    <li>Si no encuentra, intenta match por <b>equivalencias de tarjeta</b> configuradas en TM.</li>
    <li>Si hay reversos en la procesadora, los cruza contra negativos de Skylab.</li>
    <li>Las operaciones integradas (cruzadas internamente en Skylab) se marcan como <b>OK (integrado)</b>.</li>
  </ol>
</div>
`,

// ─────────────────────────────────────────────────────────────────────
revision: `
<h3>Revisión Manual</h3>
<p>Algunas operaciones quedan como SIN MATCH por razones legítimas. La revisión manual permite registrar el motivo y marcarlas como resueltas.</p>

<div class="gy-bloque">
  <div class="gy-bloque-titulo">Cómo aplicar una corrección</div>
  <ol>
    <li>Ir a <b>Operaciones → Revisión Manual</b>.</li>
    <li>Buscar la operación por cupón, sucursal, monto o fecha.</li>
    <li>Hacer clic en el botón <b>Corregir</b> de la fila correspondiente.</li>
    <li>Completar: resultado, motivo, método, observaciones y usuario.</li>
    <li>Guardar. La corrección queda asociada a esa operación.</li>
  </ol>
</div>

<div class="gy-bloque">
  <div class="gy-bloque-titulo">Arrastre automático de correcciones</div>
  <p>Al <b>cerrar el período</b>, todas las correcciones manuales se guardan en IndexedDB.
  Al inicio del <b>período siguiente</b>, el sistema intenta re-aplicarlas automáticamente
  sobre las operaciones del nuevo período usando dos estrategias:</p>
  <ul>
    <li><b>Match exacto</b> — por número de asiento Skylab (mismo ticket).</li>
    <li><b>Match fuzzy</b> — si el asiento cambió: misma sucursal + tarjeta + cuotas + monto ±2%.</li>
  </ul>
  <p>Las correcciones solo se auto-aplican si no hiciste ninguna corrección manual en el período nuevo todavía.</p>
</div>

<div class="gy-bloque">
  <div class="gy-bloque-titulo">Importación desde período anterior</div>
  <p>También podés importar correcciones manualmente desde el <b>Historial</b>
  (botón 📋 por período) o cargando un backup JSON desde el toolbar de Correcciones.</p>
</div>
`,

// ─────────────────────────────────────────────────────────────────────
diferencias: `
<h3>Diferencias</h3>
<p>El módulo de Diferencias analiza discrepancias que no impiden el match pero que afectan el monto efectivamente cobrado.</p>

<div class="gy-bloque">
  <div class="gy-bloque-titulo">Tipos de diferencias</div>
  <table class="gy-tabla">
    <tr><th>Tipo</th><th>Qué compara</th><th>Impacto</th></tr>
    <tr>
      <td><b>Tasas</b></td>
      <td>Tasa cobrada por la procesadora vs. tasa pactada en Tablas Maestras (por procesadora, tarjeta, comercio)</td>
      <td>Diferencia de monto retenido por comisión</td>
    </tr>
    <tr>
      <td><b>Cuotas</b></td>
      <td>Plan de cuotas registrado en Skylab vs. el informado por la procesadora</td>
      <td>Puede afectar el flujo de cobros si el plazo es diferente</td>
    </tr>
    <tr>
      <td><b>Procesadora</b></td>
      <td>La operación aparece en una procesadora diferente a la que indica Skylab</td>
      <td>Riesgo de doble cobro o de no cobro</td>
    </tr>
    <tr>
      <td><b>Monto real</b></td>
      <td>El monto acreditado por la procesadora difiere del monto vendido</td>
      <td>Diferencia directa en el cobro recibido</td>
    </tr>
  </table>
</div>

<div class="gy-bloque">
  <div class="gy-bloque-titulo">Clasificación de diferencias de tasa</div>
  <p>Cada diferencia de tasa puede clasificarse manualmente para entender su origen:</p>
  <ul>
    <li><span class="gy-badge" style="background:rgba(52,211,153,.15);color:#34d399">Vendedor</span> — el vendedor aplicó una tasa incorrecta al cliente.</li>
    <li><span class="gy-badge" style="background:rgba(248,113,113,.15);color:#f87171">Procesadora</span> — la procesadora aplicó una tasa diferente a la pactada.</li>
    <li><span class="gy-badge" style="background:rgba(56,189,248,.15);color:var(--cyn)">OK</span> — diferencia esperada (ej.: promoción vigente).</li>
    <li><span class="gy-badge" style="background:var(--s3);color:var(--m2)">Sin clasificar</span> — pendiente de revisión.</li>
  </ul>
  <p style="margin-top:6px">La clasificación se guarda en localStorage y se muestra en el Dashboard.</p>
</div>
`,

// ─────────────────────────────────────────────────────────────────────
liquidaciones: `
<h3>Liquidaciones y Cobros</h3>
<p>El módulo de Liquidaciones verifica qué operaciones conciliadas ya fueron acreditadas efectivamente en la cuenta bancaria.</p>

<div class="gy-bloque">
  <div class="gy-bloque-titulo">Cómo funciona</div>
  <ol>
    <li>Cargá el archivo de <b>Liquidaciones</b> en el panel lateral izquierdo.</li>
    <li>El sistema cruza las operaciones OK del cruce contra los cupones del archivo de liquidaciones.</li>
    <li>El resultado muestra: <b>Liquidadas</b>, <b>Sin liquidar</b>, <b>Extras</b> (en liq pero no en cruce) y <b>Fuera de plazo</b>.</li>
  </ol>
</div>

<div class="gy-bloque">
  <div class="gy-bloque-titulo">Plazo de acreditación esperado</div>
  <p>Para cada operación sin liquidar, el sistema calcula la <b>fecha de acreditación esperada</b>
  en base a los <b>días hábiles</b> configurados en Tablas Maestras (por procesadora y tarjeta),
  descontando fines de semana y feriados.</p>
  <ul>
    <li><span style="color:#34d399">Verde</span> — dentro del plazo.</li>
    <li><span style="color:#f87171">Rojo</span> — plazo vencido (puede ser un cobro perdido).</li>
  </ul>
</div>

<div class="gy-bloque">
  <div class="gy-bloque-titulo">Arrastre al mes siguiente</div>
  <p>Las operaciones sin liquidar al cerrar el período se guardan automáticamente en IndexedDB.
  En el próximo período, aparecen en la pestaña <b>"Arrastre mes ant."</b> del módulo Cobros
  <b>sin necesidad de subir ningún archivo</b>. El sistema re-intenta el match contra las nuevas liquidaciones
  para ver cuáles finalmente se acreditaron.</p>
</div>

<div class="gy-bloque">
  <div class="gy-bloque-titulo">Pestañas del módulo Cobros</div>
  <table class="gy-tabla">
    <tr><th>Pestaña</th><th>Qué muestra</th></tr>
    <tr><td>Sin liquidar</td><td>Ops conciliadas que no aparecen en el archivo de liq del período</td></tr>
    <tr><td>Extras</td><td>Ops en liq que no tienen match en el cruce (posible error o doble pago)</td></tr>
    <tr><td>Arrastre mes ant.</td><td>Ops sin liquidar del período anterior, con su nuevo estado</td></tr>
  </table>
</div>
`,

// ─────────────────────────────────────────────────────────────────────
historial: `
<h3>Historial de Períodos</h3>
<p>Cada período cerrado queda registrado en el historial con un resumen completo.</p>

<div class="gy-bloque">
  <div class="gy-bloque-titulo">Cerrar un período</div>
  <ol>
    <li>Completar el cruce y todas las revisiones del período.</li>
    <li>Hacer clic en <b>Cerrar período</b> (en el sidebar o en el Historial).</li>
    <li>El sistema guarda en <b>IndexedDB</b> y descarga automáticamente un <b>backup JSON</b>.</li>
    <li>El arrastre (pendientes de liquidación + correcciones manuales) queda disponible para el período siguiente.</li>
  </ol>
</div>

<div class="gy-bloque">
  <div class="gy-bloque-titulo">Qué se guarda por período</div>
  <ul>
    <li>Estadísticas del cruce (total ops, % OK, sin match, correcciones).</li>
    <li>Resumen de cobros y contracargos.</li>
    <li>Operaciones sin liquidar (<code>pendientesArrastre</code>).</li>
    <li>Correcciones manuales con metadata para re-matching (<code>correccionesArrastre</code>).</li>
  </ul>
</div>

<div class="gy-bloque">
  <div class="gy-bloque-titulo">Importar correcciones desde el historial</div>
  <p>Desde la tabla del historial, hacé clic en el botón <b>📋 N</b> de cualquier período
  para importar sus correcciones manuales al período actual. Útil si el arrastre automático
  no encontró el match por alguna razón.</p>
</div>

<div class="gy-bloque">
  <div class="gy-bloque-titulo">Portabilidad entre dispositivos</div>
  <p>IndexedDB es local al navegador/dispositivo. Para usar el historial en otra computadora,
  cargá el <b>backup JSON</b> exportado al cerrar el período. El JSON contiene todo lo necesario
  para reproducir el resumen y el arrastre.</p>
</div>
`,

// ─────────────────────────────────────────────────────────────────────
dashboard: `
<h3>Dashboard de Resumen</h3>
<p>Vista consolidada de los indicadores más importantes del período activo.</p>

<div class="gy-bloque">
  <div class="gy-bloque-titulo">Pestañas del Dashboard</div>
  <table class="gy-tabla">
    <tr><th>Pestaña</th><th>Indicadores principales</th></tr>
    <tr>
      <td><b>Transacciones</b></td>
      <td>Total de operaciones, % OK, SIN MATCH, distribución por estado y por sucursal</td>
    </tr>
    <tr>
      <td><b>Pagos / Cobros</b></td>
      <td>Monto cobrado vs. pendiente, distribución por tarjeta, evolución histórica de cobros</td>
    </tr>
    <tr>
      <td><b>Financiero</b></td>
      <td>Análisis de contracargos: montos en riesgo, distribución por tarjeta y estado</td>
    </tr>
    <tr>
      <td><b>Por tarjeta</b></td>
      <td>Operaciones y montos desglosados por marca de tarjeta</td>
    </tr>
    <tr>
      <td><b>Liquidaciones</b></td>
      <td>Liquidadas vs. sin liquidar vs. extras por procesadora, diferencias de tasas clasificadas, top sucursales sin liquidar</td>
    </tr>
  </table>
</div>

<div class="gy-bloque">
  <div class="gy-bloque-titulo">Actualización</div>
  <p>El dashboard se recalcula haciendo clic en el botón <b>↺ Actualizar</b>.
  No se actualiza automáticamente para evitar recálculos innecesarios al navegar entre módulos.</p>
</div>
`,

// ─────────────────────────────────────────────────────────────────────
tm: `
<h3>Configuración — Tablas Maestras</h3>
<p>Las Tablas Maestras son la configuración central del sistema. Se cargan una vez y persisten en IndexedDB.</p>

<div class="gy-bloque">
  <div class="gy-bloque-titulo">Secciones de configuración</div>
  <table class="gy-tabla">
    <tr><th>Sección</th><th>Qué configura</th><th>Formato</th></tr>
    <tr>
      <td><b>Plazos</b></td>
      <td>Días hábiles de acreditación por procesadora, tarjeta y comercio. Incluye vigencia.</td>
      <td>Excel (.xlsx)</td>
    </tr>
    <tr>
      <td><b>Feriados</b></td>
      <td>Calendario de feriados nacionales y provinciales. Se usa para calcular días hábiles.</td>
      <td>Excel (.xlsx)</td>
    </tr>
    <tr>
      <td><b>Tasas</b></td>
      <td>Tasas pactadas por procesadora, tarjeta y plan de cuotas. Base para detectar diferencias.</td>
      <td>Excel (.xlsx)</td>
    </tr>
    <tr>
      <td><b>Equivalencias</b></td>
      <td>Mapeo de nombres de tarjeta entre Skylab y cada procesadora (ej.: "VISA DEBITO" → "VD").</td>
      <td>Excel (.xlsx)</td>
    </tr>
    <tr>
      <td><b>Comercios</b></td>
      <td>Datos de los comercios (CUIT, nombre, código). Se usan para filtrar tasas y plazos.</td>
      <td>Excel (.xlsx)</td>
    </tr>
  </table>
</div>

<div class="gy-bloque">
  <div class="gy-bloque-titulo">Persistencia</div>
  <p>Las Tablas Maestras se guardan en IndexedDB al cargarlas. Se restauran automáticamente
  al abrir la aplicación. No es necesario volver a cargarlas cada período salvo que haya cambios
  (nuevas tasas, feriados, equivalencias).</p>
</div>

<div class="gy-bloque">
  <div class="gy-bloque-titulo">Apariencia</div>
  <p>El botón <b>⚙</b> en la barra superior permite cambiar el fondo y la tipografía de la aplicación.
  La configuración de apariencia se guarda en <b>localStorage</b>.</p>
</div>
`,

};

// ── Render ────────────────────────────────────────────────────────────

function abrirGuia() {
  let modal = document.getElementById('modal-guia');
  if (!modal) {
    modal = document.createElement('div');
    modal.id = 'modal-guia';
    modal.className = 'modal-overlay';
    modal.innerHTML = _guiaHTML();
    document.body.appendChild(modal);
    modal.addEventListener('click', e => { if (e.target === modal) cerrarGuia(); });
  }
  modal.classList.add('open');
  _showGuiaTab('objetivo', modal.querySelector('[data-tab="objetivo"]'));
}

function cerrarGuia() {
  document.getElementById('modal-guia')?.classList.remove('open');
}

function _showGuiaTab(tabId, btn) {
  const modal = document.getElementById('modal-guia');
  if (!modal) return;
  modal.querySelectorAll('.gy-tab-content').forEach(el => el.style.display = 'none');
  modal.querySelectorAll('.gy-tab-btn').forEach(el => el.classList.remove('active'));
  const content = modal.querySelector(`#gy-content-${tabId}`);
  if (content) content.style.display = 'block';
  if (btn) btn.classList.add('active');
}

function _guiaHTML() {
  const tabBtns = _GUIA_TABS.map(t =>
    `<button class="gy-tab-btn" data-tab="${t.id}" onclick="_showGuiaTab('${t.id}',this)">${t.label}</button>`
  ).join('');

  const tabContents = _GUIA_TABS.map(t =>
    `<div id="gy-content-${t.id}" class="gy-tab-content" style="display:none">${_GUIA_CONTENIDO[t.id] || ''}</div>`
  ).join('');

  return `
<div class="modal-guia" role="dialog" aria-label="Guía de uso">
  <div class="gy-header">
    <div class="gy-header-title">
      <span style="color:var(--acc);font-size:16px">?</span>
      Guía del Conciliador Skylab
    </div>
    <button class="gy-close-btn" onclick="cerrarGuia()" title="Cerrar">✕</button>
  </div>

  <div class="gy-tabs">${tabBtns}</div>

  <div class="gy-body">${tabContents}</div>
</div>`;
}
