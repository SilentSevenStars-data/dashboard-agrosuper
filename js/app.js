/* Panel de gestión — lee un Excel publicado en el mismo repositorio (GitHub Pages)
   y arma KPIs, gráficos y tabla de forma automática según las columnas del archivo. */
(() => {
'use strict';

const AZUL = '#003087', GRILLA = '#D9DFEB';
const nf  = new Intl.NumberFormat('es-CL', { maximumFractionDigits: 2 });
const nf0 = new Intl.NumberFormat('es-CL', { maximumFractionDigits: 0 });
const nfC = new Intl.NumberFormat('es-CL', { notation: 'compact', maximumFractionDigits: 1 });
const $ = (s) => document.querySelector(s);
const pad = (n) => String(n).padStart(2, '0');

const state = {
  cfg: {}, wb: null, sheet: null, origen: 'publicado', lastMod: null, archivo: '',
  rows: [], cols: [], types: {},
  metric: null, agg: 'sum', dim1: null, dim2: null, dateCol: null, grain: 'month',
  desde: '', hasta: '', filters: {}, search: '', sortCol: null, sortDir: 1, page: 0, pageSize: 25,
};
const charts = {};

/* ---------- Utilidades de fecha ---------- */
const keyDia = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
function normDate(d) {
  const r = new Date(Math.round(d.getTime() / 60000) * 60000);   // corrige desfases de segundos de Excel
  return new Date(r.getFullYear(), r.getMonth(), r.getDate());
}
function parseDateStr(s) {
  s = String(s).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{2,4})$/);
  if (m) { const y = +m[3] < 100 ? 2000 + +m[3] : +m[3]; return new Date(y, +m[2] - 1, +m[1]); }
  return null;
}
function periodKey(d, g) {
  if (g === 'year') return String(d.getFullYear());
  if (g === 'month') return `${d.getFullYear()}-${pad(d.getMonth() + 1)}`;
  if (g === 'week') { const m = new Date(d); m.setDate(d.getDate() - ((d.getDay() + 6) % 7)); return keyDia(m); }
  return keyDia(d);
}
function nextKey(key, g) {
  if (g === 'year') return String(+key + 1);
  if (g === 'month') { const [y, m] = key.split('-').map(Number); const d = new Date(y, m, 1); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}`; }
  const [y, m, dd] = key.split('-').map(Number);
  const d = new Date(y, m - 1, dd + (g === 'week' ? 7 : 1));
  return keyDia(d);
}
function periodEnd(key, g) {
  if (g === 'year') return `${key}-12-31`;
  if (g === 'month') { const [y, m] = key.split('-').map(Number); return keyDia(new Date(y, m, 0)); }
  if (g === 'week') { const [y, m, d] = key.split('-').map(Number); return keyDia(new Date(y, m - 1, d + 6)); }
  return key;
}
function periodLabel(key, g) {
  if (g === 'year') return key;
  if (g === 'month') { const [y, m] = key.split('-').map(Number); return new Date(y, m - 1, 1).toLocaleDateString('es-CL', { month: 'short', year: '2-digit' }).replace('.', ''); }
  const [y, m, d] = key.split('-').map(Number);
  const t = `${pad(d)}-${pad(m)}-${String(y).slice(2)}`;
  return g === 'week' ? `Sem ${t}` : t;
}
const fmtFecha = (d) => `${pad(d.getDate())}-${pad(d.getMonth() + 1)}-${d.getFullYear()}`;

/* ---------- Carga ---------- */
async function init() {
  try { state.cfg = await (await fetch('config.json', { cache: 'no-store' })).json(); } catch { state.cfg = {}; }
  $('#titulo').textContent = state.cfg.titulo || 'Panel de gestión';
  document.title = state.cfg.titulo || 'Panel de gestión';
  $('#subtitulo').textContent = state.cfg.subtitulo || '';
  try { await Promise.all(['350 12px Biennale', '400 12px Biennale', '700 12px Biennale', '900 12px Biennale'].map((f) => document.fonts.load(f))); } catch {}
  setupChartDefaults();
  bindStaticEvents();
  await cargarPublicado();
}

async function cargarPublicado() {
  setEstado('Cargando datos…');
  $('#aviso').hidden = true;
  const url = (state.cfg.archivo || 'data/datos.xlsx');
  try {
    const res = await fetch(`${url}?v=${Date.now()}`, { cache: 'no-store' });
    if (!res.ok) throw new Error(`No se encontró "${url}" (HTTP ${res.status}).`);
    const lm = res.headers.get('Last-Modified');
    const buf = await res.arrayBuffer();
    state.origen = 'publicado';
    state.lastMod = lm ? new Date(lm) : null;
    state.archivo = url;
    cargarLibro(buf, state.cfg.hoja);
  } catch (e) {
    const file = location.protocol === 'file:';
    setEstado(`<strong>No pude cargar los datos.</strong><br>${esc(e.message)}<br>` +
      (file ? 'Estás abriendo el archivo directamente desde tu computador; el navegador bloquea esa lectura. Publica el sitio en GitHub Pages o usa “Probar archivo local”.'
            : 'Revisa que el archivo exista en la carpeta <code>data/</code> del repositorio y que el nombre coincida con <code>config.json</code>.'), true);
  }
}

function cargarLibro(buf, hojaPref) {
  const wb = XLSX.read(buf, { type: 'array', cellDates: true });
  state.wb = wb;
  const names = wb.SheetNames;
  state.sheet = names.includes(hojaPref) ? hojaPref : names[0];
  const sel = $('#selHoja');
  sel.innerHTML = names.map((n) => `<option>${esc(n)}</option>`).join('');
  sel.value = state.sheet;
  $('#campoHoja').hidden = names.length < 2;
  cargarHoja(true);
}

function cargarHoja(primera) {
  const ws = state.wb.Sheets[state.sheet];
  let rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: true });
  rows = rows.map((r) => { const o = {}; for (const k in r) o[String(k).trim()] = r[k]; return o; })
             .filter((r) => Object.values(r).some((v) => v !== null && v !== ''));
  if (!rows.length) { setEstado('<strong>El archivo no tiene datos.</strong> La primera fila debe contener los nombres de las columnas.', true); return; }
  state.rows = rows;
  inferirTipos();
  defaults(primera);
  construirControles();
  $('#estado').hidden = true;
  $('#app').hidden = false;
  render();
}

function inferirTipos() {
  const rows = state.rows;
  const cols = Object.keys(rows[0]);
  state.cols = cols; state.types = {};
  for (const c of cols) {
    const vals = rows.map((r) => r[c]).filter((v) => v !== null && v !== '' && v !== undefined);
    let t = 'text';
    if (!vals.length) t = 'text';
    else if (vals.every((v) => v instanceof Date)) t = 'date';
    else if (vals.every((v) => typeof v === 'number')) {
      const uniq = new Set(vals).size;
      const esId = /^(id|cod|c[oó]digo|folio|rut|n[uú]mero|nro|n°)\b/i.test(c);
      const esCodigo = vals.every(Number.isInteger) && uniq <= 15 && /(a[nñ]o|year|mes|semana|week|trimestre|periodo|per[ií]odo)/i.test(c);
      t = esId ? 'text' : esCodigo ? 'cat' : 'num';
    } else {
      const strs = vals.map(String);
      const fechas = strs.map(parseDateStr);
      if (fechas.filter(Boolean).length / strs.length >= 0.9) t = 'date';
      else {
        const uniq = new Set(strs).size;
        t = (uniq <= 60 && uniq < vals.length * 0.6) || uniq <= 12 ? 'cat' : 'text';
      }
    }
    state.types[c] = t;
  }
  for (const r of rows) for (const c of cols) {
    const t = state.types[c], v = r[c];
    if (t === 'date') r[c] = v instanceof Date ? normDate(v) : (v === null || v === '' ? null : parseDateStr(v));
    else if (t === 'cat') r[c] = v === null || v === '' ? '(sin dato)' : String(v);
    else if (t === 'num' && typeof v !== 'number') r[c] = null;
  }
}

function colsDe(tipo) { return state.cols.filter((c) => state.types[c] === tipo); }

function defaults(primera) {
  const cfg = state.cfg, nums = colsDe('num'), cats = colsDe('cat'), dates = colsDe('date');
  const pick = (pref, list, i = 0) => (pref && list.includes(pref) ? pref : list[i] || null);
  state.metric = pick(primera ? cfg.metricaPorDefecto : null, nums);
  state.agg = state.metric ? 'sum' : 'count';
  state.dim1 = pick(primera ? cfg.dimensionPorDefecto : null, cats);
  state.dim2 = pick(primera ? cfg.dimension2PorDefecto : null, cats.filter((c) => c !== state.dim1), 0);
  state.dateCol = pick(primera ? cfg.columnaFecha : null, dates);
  state.filters = {}; state.search = ''; state.sortCol = null; state.page = 0;
  const f = state.dateCol ? state.rows.map((r) => r[state.dateCol]).filter(Boolean) : [];
  if (f.length) {
    const min = new Date(Math.min(...f)), max = new Date(Math.max(...f));
    state.desde = keyDia(min); state.hasta = keyDia(max);
    state.dateMin = state.desde; state.dateMax = state.hasta;
    const dias = (max - min) / 86400000;
    state.grain = dias > 120 ? 'month' : dias > 31 ? 'week' : 'day';
    if (dias > 365 * 4) state.grain = 'year';
  } else { state.desde = state.hasta = ''; }
}

/* ---------- Controles ---------- */
function opts(list, sel, vacio) {
  if (!list.length) return `<option value="">${vacio || '—'}</option>`;
  return list.map((c) => `<option value="${esc(c)}"${c === sel ? ' selected' : ''}>${esc(c)}</option>`).join('');
}
function construirControles() {
  const nums = colsDe('num'), cats = colsDe('cat');
  $('#selMetrica').innerHTML = opts(nums, state.metric, 'Sin columnas numéricas');
  $('#campoMetrica').hidden = !nums.length;
  $('#selAgg').value = state.agg;
  $('#selAgg').querySelectorAll('option').forEach((o) => { if (o.value !== 'count') o.disabled = !nums.length; });
  $('#selDim1').innerHTML = opts(cats, state.dim1, 'Sin columnas de categoría');
  $('#selDim2').innerHTML = opts(cats, state.dim2, 'Sin columnas de categoría');
  $('#campoDim1').hidden = !cats.length;
  $('#campoDim2').hidden = cats.length < 2;
  $('#selGrano').value = state.grain;
  const hayFecha = !!state.dateCol;
  ['#campoGrano', '#campoDesde', '#campoHasta'].forEach((s) => ($(s).hidden = !hayFecha));
  $('#inDesde').value = state.desde; $('#inHasta').value = state.hasta;
  construirFiltros();
}

function construirFiltros() {
  const cont = $('#filtros');
  const cats = colsDe('cat').filter((c) => new Set(state.rows.map((r) => r[c])).size <= 60).slice(0, 6);
  if (!cats.length) { cont.innerHTML = ''; return; }
  cont.innerHTML = '<span class="lbl">Filtros</span>';
  for (const c of cats) {
    const valores = [...new Set(state.rows.map((r) => r[c]))].sort((a, b) => String(a).localeCompare(String(b), 'es'));
    const d = document.createElement('details');
    d.className = 'ms'; d.dataset.col = c;
    d.innerHTML = `<summary><span>${esc(c)}</span><span class="cnt"></span></summary>
      <div class="ms-panel"><div class="ms-acc"><button type="button" data-a="all">Todos</button><button type="button" data-a="none">Ninguno</button></div>
      ${valores.map((v) => `<label class="ms-op"><input type="checkbox" value="${esc(v)}" checked> <span>${esc(v)}</span></label>`).join('')}</div>`;
    cont.appendChild(d);
    const act = () => {
      const checks = [...d.querySelectorAll('input[type=checkbox]')];
      const sel = checks.filter((x) => x.checked).map((x) => x.value);
      if (sel.length === checks.length) delete state.filters[c]; else state.filters[c] = new Set(sel);
      d.classList.toggle('activo', !!state.filters[c]);
      d.querySelector('.cnt').textContent = state.filters[c] ? `(${sel.length}/${checks.length})` : '';
      state.page = 0; render();
    };
    d.addEventListener('change', act);
    d.querySelectorAll('[data-a]').forEach((b) => b.addEventListener('click', () => {
      d.querySelectorAll('input[type=checkbox]').forEach((x) => (x.checked = b.dataset.a === 'all')); act();
    }));
  }
  const lim = document.createElement('button');
  lim.type = 'button'; lim.className = 'limpiar'; lim.textContent = 'Limpiar filtros';
  lim.addEventListener('click', () => {
    state.filters = {};
    cont.querySelectorAll('input[type=checkbox]').forEach((x) => (x.checked = true));
    cont.querySelectorAll('details').forEach((d) => { d.classList.remove('activo'); d.querySelector('.cnt').textContent = ''; });
    state.page = 0; render();
  });
  cont.appendChild(lim);
}

function bindStaticEvents() {
  $('#selHoja').addEventListener('change', (e) => { state.sheet = e.target.value; cargarHoja(false); });
  $('#selMetrica').addEventListener('change', (e) => { state.metric = e.target.value; render(); });
  $('#selAgg').addEventListener('change', (e) => { state.agg = e.target.value; render(); });
  $('#selDim1').addEventListener('change', (e) => { state.dim1 = e.target.value; render(); });
  $('#selDim2').addEventListener('change', (e) => { state.dim2 = e.target.value; render(); });
  $('#selGrano').addEventListener('change', (e) => { state.grain = e.target.value; render(); });
  $('#inDesde').addEventListener('change', (e) => { state.desde = e.target.value; state.page = 0; render(); });
  $('#inHasta').addEventListener('change', (e) => { state.hasta = e.target.value; state.page = 0; render(); });
  $('#inBuscar').addEventListener('input', (e) => { state.search = e.target.value.toLowerCase(); state.page = 0; renderTabla(); });
  $('#pgPrev').addEventListener('click', () => { state.page = Math.max(0, state.page - 1); renderTabla(); });
  $('#pgNext').addEventListener('click', () => { state.page++; renderTabla(); });
  $('#btnRefrescar').addEventListener('click', cargarPublicado);
  $('#btnExportar').addEventListener('click', exportar);
  $('#fileLocal').addEventListener('change', async (e) => {
    const f = e.target.files[0]; if (!f) return;
    try {
      const buf = await f.arrayBuffer();
      state.origen = 'local'; state.archivo = f.name; state.lastMod = f.lastModified ? new Date(f.lastModified) : null;
      cargarLibro(buf, null);
      const av = $('#aviso');
      av.innerHTML = `<span>Estás viendo el archivo local “${esc(f.name)}”. No está publicado: solo lo ves tú.</span><button class="btn" type="button" id="btnVolver">Volver a datos publicados</button>`;
      av.hidden = false;
      $('#btnVolver').addEventListener('click', cargarPublicado);
    } catch (err) { setEstado(`<strong>No pude leer el archivo.</strong><br>${esc(err.message)}`, true); }
    e.target.value = '';
  });
  document.addEventListener('click', (e) => {
    document.querySelectorAll('details.ms[open]').forEach((d) => { if (!d.contains(e.target)) d.open = false; });
  });
}

/* ---------- Cálculo ---------- */
function filtradas() {
  const { filters, dateCol, desde, hasta } = state;
  return state.rows.filter((r) => {
    for (const c in filters) if (!filters[c].has(r[c])) return false;
    if (dateCol && (desde || hasta)) {
      const d = r[dateCol]; if (!d) return false;
      const k = keyDia(d);
      if (desde && k < desde) return false;
      if (hasta && k > hasta) return false;
    }
    return true;
  });
}
function agregar(rows) {
  if (state.agg === 'count' || !state.metric) return rows.length;
  const v = rows.map((r) => r[state.metric]).filter((x) => typeof x === 'number');
  if (!v.length) return state.agg === 'sum' ? 0 : null;
  const s = v.reduce((a, b) => a + b, 0);
  return state.agg === 'sum' ? s : s / v.length;
}
function agrupar(rows, col) {
  const m = new Map();
  for (const r of rows) { const k = r[col]; if (!m.has(k)) m.set(k, []); m.get(k).push(r); }
  return m;
}
const nombreMetrica = () => {
  const base = state.agg === 'count' || !state.metric ? 'Registros' : state.metric;
  return state.agg === 'avg' ? `${base} (promedio)` : state.agg === 'sum' && state.metric ? `${base} (suma)` : base;
};
const fmtVal = (v) => (v === null || v === undefined ? '—' : Math.abs(v) >= 1000 ? nf0.format(v) : nf.format(v));

function serieTiempo(rows) {
  const g = state.grain, m = agrupar(rows.filter((r) => r[state.dateCol]).map((r) => ({ ...r, __k: periodKey(r[state.dateCol], g) })), '__k');
  const keys = [...m.keys()].sort();
  if (!keys.length) return [];
  const todas = [];
  let k = keys[0], guard = 0;
  while (k <= keys[keys.length - 1] && guard++ < 800) { todas.push(k); k = nextKey(k, g); }
  const lista = guard < 800 ? todas : keys;
  return lista.map((key) => ({ key, v: m.has(key) ? agregar(m.get(key)) : (state.agg === 'avg' ? null : 0) }));
}
function topN(rows, col, n = 10) {
  const g = [...agrupar(rows, col).entries()].map(([k, rs]) => ({ k, v: agregar(rs) })).filter((x) => x.v !== null).sort((a, b) => b.v - a.v);
  if (g.length <= n) return g;
  const top = g.slice(0, n);
  if (state.agg !== 'avg') {
    const resto = g.slice(n);
    top.push({ k: `Otros (${resto.length})`, v: resto.reduce((a, b) => a + b.v, 0) });
  }
  return top;
}

/* ---------- Render ---------- */
function render() {
  const rows = filtradas();
  renderKpis(rows);
  renderCharts(rows);
  renderTabla();
  const lm = state.lastMod ? state.lastMod.toLocaleString('es-CL', { dateStyle: 'medium', timeStyle: 'short' }) : '—';
  $('#pie').textContent = `Fuente: ${state.archivo} · hoja “${state.sheet}” · ${nf0.format(state.rows.length)} registros en el archivo · archivo actualizado: ${lm}`;
}

function kpi(lbl, val, sub, cls = '') {
  return `<div class="kpi"><div class="k-lbl">${esc(lbl)}</div><div class="k-val ${cls}">${val}</div><div class="k-sub">${sub}</div></div>`;
}
function renderKpis(rows) {
  const out = [];
  const total = agregar(rows);
  out.push(kpi(nombreMetrica(), fmtVal(total), `${nf0.format(rows.length)} registros con los filtros actuales`));
  if (state.metric && state.agg !== 'avg') {
    const v = rows.map((r) => r[state.metric]).filter((x) => typeof x === 'number');
    out.push(kpi(`Promedio por registro`, fmtVal(v.length ? v.reduce((a, b) => a + b, 0) / v.length : null), esc(state.metric)));
  } else if (state.metric) {
    const v = rows.map((r) => r[state.metric]).filter((x) => typeof x === 'number');
    out.push(kpi(`Suma total`, fmtVal(v.reduce((a, b) => a + b, 0)), esc(state.metric)));
  }
  if (state.dim1) {
    const t = topN(rows, state.dim1, 1000);
    if (t.length) {
      const lider = t[0];
      const suma = state.agg !== 'avg' ? t.reduce((a, b) => a + b.v, 0) : null;
      out.push(kpi(`Líder en ${state.dim1}`, esc(String(lider.k)), suma ? `${nf.format((lider.v / suma) * 100)} % del total · ${fmtVal(lider.v)}` : fmtVal(lider.v), 'txt'));
      out.push(kpi(`Categorías en ${state.dim1}`, nf0.format(t.length), 'con datos en el filtro actual'));
    }
  }
  if (state.dateCol) {
    let s = serieTiempo(rows);
    let parcial = false;
    const fin = state.hasta || state.dateMax;
    if (s.length && fin && periodEnd(s[s.length - 1].key, state.grain) > fin) { s = s.slice(0, -1); parcial = true; }
    if (s.length >= 2) {
      const a = s[s.length - 1], b = s[s.length - 2];
      if (a.v !== null && b.v) {
        const pct = ((a.v - b.v) / Math.abs(b.v)) * 100;
        const sube = pct >= 0;
        out.push(kpi('Variación último período cerrado', `${sube ? '▲' : '▼'} ${nf.format(Math.abs(pct))} %`,
          `<span class="${sube ? 'ok' : 'alerta'}">${sube ? 'Sube' : 'Baja'}</span> · ${periodLabel(a.key, state.grain)} vs ${periodLabel(b.key, state.grain)}` +
          (parcial ? '<br>(se excluye el período en curso, aún incompleto)' : '')));
      }
    }
  }
  $('#kpis').innerHTML = out.join('');
}

function setupChartDefaults() {
  const C = Chart.defaults;
  C.font.family = "'Biennale', system-ui, sans-serif"; C.font.size = 12; C.color = AZUL;
  C.maintainAspectRatio = false; C.animation = { duration: 250 };
  C.plugins.legend.display = false;
  Object.assign(C.plugins.tooltip, { backgroundColor: AZUL, titleColor: '#FFF6E6', bodyColor: '#FFF6E6', cornerRadius: 6, padding: 10, displayColors: false,
    titleFont: { weight: 700, family: "'Biennale', sans-serif" } });
}
const valueLabels = {
  id: 'valueLabels',
  afterDatasetsDraw(chart) {
    if (chart.options.indexAxis !== 'y') return;
    const { ctx } = chart; ctx.save();
    ctx.font = "400 12px 'Biennale', sans-serif"; ctx.fillStyle = AZUL; ctx.textBaseline = 'middle';
    chart.getDatasetMeta(0).data.forEach((bar, i) => {
      const v = chart.data.datasets[0].data[i]; if (v === null || v === undefined) return;
      ctx.textAlign = v < 0 ? 'right' : 'left';
      ctx.fillText(nfC.format(v), bar.x + (v < 0 ? -6 : 6), bar.y);
    });
    ctx.restore();
  },
};
function drawChart(id, cfg) { if (charts[id]) charts[id].destroy(); charts[id] = new Chart($('#' + id), cfg); }

function renderCharts(rows) {
  const nm = nombreMetrica();
  // Tiempo
  $('#cardTiempo').hidden = !state.dateCol;
  if (state.dateCol) {
    const s = serieTiempo(rows);
    $('#tTiempo').textContent = `${nm} por ${({ day: 'día', week: 'semana', month: 'mes', year: 'año' })[state.grain]}`;
    drawChart('chTiempo', {
      type: 'line',
      data: { labels: s.map((x) => periodLabel(x.key, state.grain)), datasets: [{ data: s.map((x) => x.v), borderColor: AZUL, backgroundColor: AZUL, borderWidth: 2, tension: 0,
        pointRadius: s.length <= 40 ? 3 : 0, pointHoverRadius: 5, pointBackgroundColor: AZUL, spanGaps: true }] },
      options: { interaction: { mode: 'index', intersect: false },
        plugins: { tooltip: { callbacks: { label: (c) => `${nm}: ${fmtVal(c.parsed.y)}` } } },
        scales: { x: { grid: { display: false }, ticks: { maxRotation: 0, autoSkipPadding: 14 }, border: { color: GRILLA } },
                  y: { beginAtZero: true, grid: { color: GRILLA }, border: { display: false }, ticks: { callback: (v) => nfC.format(v) } } } },
    });
  }
  barras('chDim1', 'tDim1', state.dim1, rows);
  const c2 = $('#chDim2').closest('.card'); c2.hidden = !state.dim2;
  if (!c2.hidden) barras('chDim2', 'tDim2', state.dim2, rows);
  $('#chDim1').closest('.card').hidden = !state.dim1;
}
function barras(id, titulo, col, rows) {
  if (!col) return;
  const t = topN(rows, col, 10);
  $('#' + titulo).textContent = `${nombreMetrica()} por ${col}`;
  drawChart(id, {
    type: 'bar', plugins: [valueLabels],
    data: { labels: t.map((x) => x.k), datasets: [{ data: t.map((x) => x.v), backgroundColor: AZUL, borderRadius: 4, borderSkipped: 'start', maxBarThickness: 22 }] },
    options: { indexAxis: 'y', layout: { padding: { right: 56 } },
      plugins: { tooltip: { callbacks: { label: (c) => `${nombreMetrica()}: ${fmtVal(c.parsed.x)}` } } },
      scales: { x: { grid: { color: GRILLA }, border: { display: false }, ticks: { callback: (v) => nfC.format(v) } },
                y: { grid: { display: false }, border: { color: GRILLA }, ticks: { autoSkip: false } } } },
  });
}

/* ---------- Tabla ---------- */
function fmtCelda(c, v) {
  if (v === null || v === undefined || v === '') return '';
  const t = state.types[c];
  if (t === 'date') return fmtFecha(v);
  if (t === 'num') return nf.format(v);
  return String(v);
}
function filasTabla() {
  let rows = filtradas();
  if (state.search) rows = rows.filter((r) => state.cols.some((c) => fmtCelda(c, r[c]).toLowerCase().includes(state.search)));
  if (state.sortCol) {
    const c = state.sortCol, t = state.types[c], d = state.sortDir;
    rows = [...rows].sort((a, b) => {
      let x = a[c], y = b[c];
      if (x === null || x === undefined) return 1; if (y === null || y === undefined) return -1;
      if (t === 'date' || t === 'num') return (x - y) * d;
      return String(x).localeCompare(String(y), 'es') * d;
    });
  }
  return rows;
}
function renderTabla() {
  const rows = filasTabla();
  const pages = Math.max(1, Math.ceil(rows.length / state.pageSize));
  state.page = Math.min(state.page, pages - 1);
  const slice = rows.slice(state.page * state.pageSize, (state.page + 1) * state.pageSize);
  const head = state.cols.map((c) => {
    const flecha = state.sortCol === c ? (state.sortDir === 1 ? '▲' : '▼') : '';
    const aria = state.sortCol === c ? (state.sortDir === 1 ? 'ascending' : 'descending') : 'none';
    return `<th class="${state.types[c] === 'num' ? 'num' : ''}" data-c="${esc(c)}" aria-sort="${aria}" tabindex="0">${esc(c)}<span class="flecha">${flecha}</span></th>`;
  }).join('');
  const body = slice.map((r) => `<tr>${state.cols.map((c) => `<td class="${state.types[c] === 'num' ? 'num' : ''}">${esc(fmtCelda(c, r[c]))}</td>`).join('')}</tr>`).join('');
  $('#tabla').innerHTML = `<thead><tr>${head}</tr></thead><tbody>${body}</tbody>`;
  $('#tabla').querySelectorAll('th').forEach((th) => {
    const go = () => { const c = th.dataset.c; if (state.sortCol === c) state.sortDir *= -1; else { state.sortCol = c; state.sortDir = 1; } renderTabla(); };
    th.addEventListener('click', go);
    th.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); go(); } });
  });
  const desde = rows.length ? state.page * state.pageSize + 1 : 0;
  $('#pagInfo').textContent = `${nf0.format(desde)}–${nf0.format(Math.min(rows.length, (state.page + 1) * state.pageSize))} de ${nf0.format(rows.length)} registros`;
  $('#pgPrev').disabled = state.page === 0;
  $('#pgNext').disabled = state.page >= pages - 1;
}
function exportar() {
  const rows = filasTabla().map((r) => { const o = {}; state.cols.forEach((c) => (o[c] = r[c])); return o; });
  const ws = XLSX.utils.json_to_sheet(rows, { cellDates: true });
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Filtrado');
  XLSX.writeFile(wb, `datos_filtrados_${keyDia(new Date())}.xlsx`);
}

/* ---------- Varios ---------- */
function esc(s) { return String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
function setEstado(html, error) {
  const e = $('#estado'); e.innerHTML = html; e.hidden = false; e.className = 'estado' + (error ? ' error' : '');
  $('#app').hidden = true;
}

init();
})();
