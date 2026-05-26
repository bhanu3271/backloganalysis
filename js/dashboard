/* =========================================================
   EXCEL DATA DASHBOARD — Full JavaScript Engine
   ========================================================= */

// ── State ──────────────────────────────────────────────────
const state = {
  rawData: [],          // original parsed rows
  filteredData: [],     // after filters applied
  columns: [],          // column definitions [{name, type, uniq}]
  fileName: '',
  currentTab: 'overview',
  currentPage: 1,
  pageSize: 50,
  sortCol: null,
  sortDir: 'asc',
  hiddenCols: new Set(),
  filters: {},          // col -> value
  customChart: null,
  overviewCharts: [],
  autoCharts: [],
  autoCompCharts: [],
};

const COLORS = [
  '#6c63ff','#00d4aa','#ff6b6b','#ffd166','#06d6a0',
  '#118ab2','#ef476f','#a8dadc','#e9c46a','#f4a261',
  '#264653','#2a9d8f','#e76f51','#457b9d','#1d3557',
];

// ── DOM Refs ────────────────────────────────────────────────
const $ = id => document.getElementById(id);
const uploadScreen = $('upload-screen');
const dashboard    = $('dashboard');
const fileInput    = $('file-input');
const sidebar      = document.querySelector('.sidebar');

// ── Boot ────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  fileInput.addEventListener('change', handleFileUpload);
  $('change-file-btn').addEventListener('click', () => {
    fileInput.value = '';
    fileInput.click();
  });
  $('sidebar-toggle').addEventListener('click', () => sidebar.classList.toggle('collapsed'));
  $('global-search').addEventListener('input', debounce(applyFilters, 250));
  $('export-btn').addEventListener('click', exportCSV);
  $('reset-filters-btn').addEventListener('click', resetFilters);
  $('toggle-filters-btn').addEventListener('click', toggleFilterPanel);
  $('col-toggle-btn').addEventListener('click', e => { e.stopPropagation(); $('col-toggle-panel').classList.toggle('hidden'); });
  document.addEventListener('click', () => $('col-toggle-panel').classList.add('hidden'));
  $('col-toggle-panel').addEventListener('click', e => e.stopPropagation());
  $('cb-generate').addEventListener('click', buildCustomChart);
  $('cmp-generate').addEventListener('click', buildComparison);

  document.querySelectorAll('.nav-item').forEach(item => {
    item.addEventListener('click', e => {
      e.preventDefault();
      switchTab(item.dataset.tab);
    });
  });
});

// ── File Upload ─────────────────────────────────────────────
function handleFileUpload(e) {
  const file = e.target.files[0];
  if (!file) return;
  const status = $('upload-status');
  status.textContent = '⏳ Reading file…';
  status.className = 'upload-status';
  status.classList.remove('hidden');

  const reader = new FileReader();
  reader.onload = ev => {
    try {
      const wb = XLSX.read(ev.target.result, { type: 'binary', cellDates: true });
      const sheetName = wb.SheetNames[0];
      const ws = wb.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
      if (!rows.length) throw new Error('Sheet is empty or unreadable.');

      state.rawData = rows;
      state.fileName = file.name;
      analyzeColumns(rows);
      init();
      status.textContent = `✅ ${rows.length.toLocaleString()} rows loaded from "${sheetName}"`;
      status.className = 'upload-status success';
      uploadScreen.classList.add('hidden');
      dashboard.classList.remove('hidden');
    } catch (err) {
      status.textContent = '❌ Error: ' + err.message;
      status.className = 'upload-status error';
    }
  };
  reader.readAsBinaryString(file);
}

// ── Column Analysis ─────────────────────────────────────────
function analyzeColumns(rows) {
  const keys = Object.keys(rows[0]);
  state.columns = keys.map(name => {
    const vals = rows.map(r => r[name]).filter(v => v !== '' && v !== null && v !== undefined);
    const isNum = vals.length > 0 && vals.every(v => !isNaN(parseFloat(v)) && isFinite(v));
    const isDate = !isNum && vals.some(v => v instanceof Date || (typeof v === 'string' && /\d{1,4}[-/]\d{1,2}[-/]\d{1,4}/.test(v)));
    const type = isNum ? 'number' : isDate ? 'date' : 'text';
    const uniq = [...new Set(vals.map(v => String(v).trim()))].sort((a,b) => {
      if (!isNaN(parseFloat(a)) && !isNaN(parseFloat(b))) return parseFloat(a) - parseFloat(b);
      return a.localeCompare(b);
    });
    const min = isNum ? Math.min(...vals.map(Number)) : null;
    const max = isNum ? Math.max(...vals.map(Number)) : null;
    return { name, type, uniq, min, max };
  });
}

// ── Init Dashboard ──────────────────────────────────────────
function init() {
  $('file-info-badge').textContent = state.fileName;
  buildFilters();
  applyFilters();
  buildChartBuilderSelects();
  buildComparisonSelects();
}

// ── Filters ─────────────────────────────────────────────────
function buildFilters() {
  const panel = $('filter-controls');
  panel.innerHTML = '';
  state.filters = {};

  state.columns.forEach(col => {
    const grp = document.createElement('div');
    grp.className = 'filter-group';
    const lbl = document.createElement('label');
    lbl.textContent = col.name;

    if (col.type === 'number') {
      lbl.textContent += ` (${fmtNum(col.min)} – ${fmtNum(col.max)})`;
      const rangeWrap = document.createElement('div');
      rangeWrap.className = 'filter-range';
      const minI = document.createElement('input');
      minI.type = 'number'; minI.placeholder = 'Min'; minI.dataset.col = col.name; minI.dataset.edge = 'min';
      const maxI = document.createElement('input');
      maxI.type = 'number'; maxI.placeholder = 'Max'; maxI.dataset.col = col.name; maxI.dataset.edge = 'max';
      const sep = document.createElement('span'); sep.textContent = '–';
      [minI, maxI].forEach(i => i.addEventListener('input', debounce(applyFilters, 300)));
      rangeWrap.append(minI, sep, maxI);
      grp.append(lbl, rangeWrap);
    } else if (col.type === 'date') {
      const rangeWrap = document.createElement('div');
      rangeWrap.className = 'filter-range';
      const fromI = document.createElement('input');
      fromI.type = 'date'; fromI.dataset.col = col.name; fromI.dataset.edge = 'from';
      const toI = document.createElement('input');
      toI.type = 'date'; toI.dataset.col = col.name; toI.dataset.edge = 'to';
      const sep = document.createElement('span'); sep.textContent = '–';
      [fromI, toI].forEach(i => i.addEventListener('change', applyFilters));
      rangeWrap.append(fromI, sep, toI);
      grp.append(lbl, rangeWrap);
    } else {
      // Text / Categorical
      if (col.uniq.length <= 60) {
        const sel = document.createElement('select');
        sel.dataset.col = col.name;
        const opt0 = document.createElement('option');
        opt0.value = ''; opt0.textContent = 'All';
        sel.appendChild(opt0);
        col.uniq.forEach(v => {
          const opt = document.createElement('option');
          opt.value = v; opt.textContent = v || '(blank)';
          sel.appendChild(opt);
        });
        sel.addEventListener('change', applyFilters);
        grp.append(lbl, sel);
      } else {
        const inp = document.createElement('input');
        inp.type = 'text'; inp.placeholder = `Search ${col.name}…`; inp.dataset.col = col.name;
        inp.addEventListener('input', debounce(applyFilters, 300));
        grp.append(lbl, inp);
      }
    }
    panel.appendChild(grp);
  });
}

function applyFilters() {
  const globalQ = $('global-search').value.trim().toLowerCase();
  const filterInputs = document.querySelectorAll('#filter-controls [data-col]');
  const colFilters = {};
  let activeCount = 0;

  filterInputs.forEach(el => {
    const col = el.dataset.col;
    const edge = el.dataset.edge;
    const val = el.value.trim();
    if (!val) return;
    if (!colFilters[col]) colFilters[col] = {};
    if (edge) colFilters[col][edge] = val;
    else colFilters[col].exact = val;
    activeCount++;
  });
  if (globalQ) activeCount++;

  // Badge
  const badge = $('filter-count-badge');
  if (activeCount > 0) { badge.textContent = activeCount; badge.classList.remove('hidden'); }
  else badge.classList.add('hidden');

  state.filters = colFilters;

  let data = state.rawData;

  // Apply column filters
  Object.entries(colFilters).forEach(([colName, f]) => {
    const colDef = state.columns.find(c => c.name === colName);
    data = data.filter(row => {
      const val = row[colName];
      if (colDef && colDef.type === 'number') {
        const n = parseFloat(val);
        if (f.min !== undefined && f.min !== '' && n < parseFloat(f.min)) return false;
        if (f.max !== undefined && f.max !== '' && n > parseFloat(f.max)) return false;
      } else if (colDef && colDef.type === 'date') {
        const d = new Date(val);
        if (f.from && d < new Date(f.from)) return false;
        if (f.to && d > new Date(f.to)) return false;
      } else {
        if (f.exact && String(val).toLowerCase().indexOf(f.exact.toLowerCase()) === -1) return false;
      }
      return true;
    });
  });

  // Global search
  if (globalQ) {
    data = data.filter(row =>
      Object.values(row).some(v => String(v).toLowerCase().includes(globalQ))
    );
  }

  state.filteredData = data;
  state.currentPage = 1;

  // Result bar
  const rb = $('result-bar');
  rb.classList.remove('hidden');
  $('result-count').innerHTML = `<strong>${data.length.toLocaleString()}</strong> of ${state.rawData.length.toLocaleString()} rows`;

  renderCurrentTab();
}

function resetFilters() {
  document.querySelectorAll('#filter-controls [data-col]').forEach(el => { el.value = ''; });
  $('global-search').value = '';
  applyFilters();
  showToast('Filters cleared', 'success');
}

function toggleFilterPanel() {
  const fc = $('filter-controls');
  const btn = $('toggle-filters-btn');
  const hidden = fc.classList.toggle('hidden');
  btn.innerHTML = hidden ? '<i class="fas fa-chevron-down"></i>' : '<i class="fas fa-chevron-up"></i>';
}

// ── Tab Switching ────────────────────────────────────────────
function switchTab(tab) {
  state.currentTab = tab;
  document.querySelectorAll('.tab-content').forEach(el => el.classList.add('hidden'));
  document.querySelectorAll('.nav-item').forEach(el => el.classList.remove('active'));
  $('tab-' + tab).classList.remove('hidden');
  document.querySelector(`.nav-item[data-tab="${tab}"]`).classList.add('active');
  $('topbar-title').textContent = { overview: 'Overview', 'data-table': 'Data Table', charts: 'Charts', comparison: 'Comparison' }[tab];
  renderCurrentTab();
}

function renderCurrentTab() {
  switch (state.currentTab) {
    case 'overview':   renderOverview(); break;
    case 'data-table': renderTable(); break;
    case 'charts':     renderAutoCharts(); break;
    case 'comparison': renderAutoComparisons(); break;
  }
}

// ── Overview ────────────────────────────────────────────────
function renderOverview() {
  renderKPIs();
  renderOverviewCharts();
}

function renderKPIs() {
  const grid = $('kpi-grid');
  grid.innerHTML = '';
  const data = state.filteredData;
  const total = data.length;

  // Total Rows KPI
  addKPI(grid, 'fas fa-database', 'Total Records', total.toLocaleString(), `of ${state.rawData.length.toLocaleString()} total`);

  const numCols = state.columns.filter(c => c.type === 'number');
  numCols.slice(0, 6).forEach(col => {
    const vals = data.map(r => parseFloat(r[col.name])).filter(v => !isNaN(v));
    const sum = vals.reduce((a, b) => a + b, 0);
    const avg = vals.length ? sum / vals.length : 0;
    addKPI(grid, 'fas fa-sigma', `Sum of ${col.name}`, fmtNum(sum), `Avg: ${fmtNum(avg)}`);
  });

  const catCols = state.columns.filter(c => c.type === 'text');
  catCols.slice(0, 3).forEach(col => {
    const freq = frequency(data, col.name);
    const top = Object.entries(freq).sort((a,b)=>b[1]-a[1])[0];
    if (top) addKPI(grid, 'fas fa-tag', `Top ${col.name}`, top[0].length > 16 ? top[0].substring(0,16)+'…' : top[0], `${top[1]} records`);
  });
}

function addKPI(grid, icon, label, value, sub) {
  const card = document.createElement('div');
  card.className = 'kpi-card';
  card.innerHTML = `<div class="kpi-icon"><i class="${icon}" style="color:var(--primary)"></i></div>
    <div class="kpi-label">${label}</div>
    <div class="kpi-value">${value}</div>
    <div class="kpi-sub">${sub}</div>`;
  grid.appendChild(card);
}

function renderOverviewCharts() {
  // Destroy old
  state.overviewCharts.forEach(c => c.destroy());
  state.overviewCharts = [];
  const container = $('overview-charts');
  container.innerHTML = '';

  const catCols = state.columns.filter(c => c.type === 'text' && c.uniq.length >= 2 && c.uniq.length <= 50);
  const numCols = state.columns.filter(c => c.type === 'number');
  const data = state.filteredData;

  // Top-5 category breakdowns
  catCols.slice(0, 4).forEach((col, i) => {
    const freq = frequency(data, col.name);
    const sorted = Object.entries(freq).sort((a,b)=>b[1]-a[1]).slice(0,15);
    const type = sorted.length <= 7 ? 'pie' : 'bar';
    const c = createChartCard(container, `<i class="fas fa-chart-${type === 'pie' ? 'pie' : 'bar'}"></i> ${col.name} Distribution`, 300);
    const chart = new Chart(c, {
      type: type,
      data: {
        labels: sorted.map(e=>e[0]||'(blank)'),
        datasets: [{ data: sorted.map(e=>e[1]), backgroundColor: COLORS, borderWidth: type==='pie'?2:0, borderRadius: type!=='pie'?6:0 }]
      },
      options: chartOpts(type, col.name, 'Count')
    });
    state.overviewCharts.push(chart);
  });

  // Number distribution (histogram) for first 2 numeric columns
  numCols.slice(0, 2).forEach(col => {
    const vals = data.map(r => parseFloat(r[col.name])).filter(v => !isNaN(v));
    if (!vals.length) return;
    const { labels, counts } = histogram(vals, 12);
    const c = createChartCard(container, `<i class="fas fa-chart-area"></i> ${col.name} Distribution`, 280);
    const chart = new Chart(c, {
      type: 'bar',
      data: { labels, datasets: [{ data: counts, backgroundColor: 'rgba(108,99,255,0.7)', borderRadius: 4, barPercentage: 0.95, categoryPercentage: 1 }] },
      options: chartOpts('bar', col.name, 'Frequency', false)
    });
    state.overviewCharts.push(chart);
  });

  // Numeric vs Category scatter (first num vs first cat)
  if (numCols.length && catCols.length) {
    const numCol = numCols[0];
    const catCol = catCols[0];
    const freq = {};
    const sums = {};
    data.forEach(r => {
      const k = String(r[catCol.name] || '(blank)');
      const v = parseFloat(r[numCol.name]);
      if (!isNaN(v)) { freq[k] = (freq[k]||0)+1; sums[k] = (sums[k]||0)+v; }
    });
    const entries = Object.entries(sums).sort((a,b)=>b[1]-a[1]).slice(0,12);
    if (entries.length) {
      const c = createChartCard(container, `<i class="fas fa-chart-bar"></i> ${numCol.name} by ${catCol.name}`, 300);
      const chart = new Chart(c, {
        type: 'bar',
        data: {
          labels: entries.map(e=>e[0]),
          datasets: [{ label: `Sum of ${numCol.name}`, data: entries.map(e=>e[1].toFixed(2)), backgroundColor: COLORS, borderRadius: 6 }]
        },
        options: chartOpts('bar', catCol.name, `Sum of ${numCol.name}`)
      });
      state.overviewCharts.push(chart);
    }
  }
}

// ── Data Table ───────────────────────────────────────────────
function renderTable() {
  const data = state.filteredData;
  const cols = state.columns.filter(c => !state.hiddenCols.has(c.name));

  // Column toggle panel
  buildColTogglePanel();

  // Sort
  let sorted = [...data];
  if (state.sortCol) {
    sorted.sort((a, b) => {
      let va = a[state.sortCol], vb = b[state.sortCol];
      const col = state.columns.find(c=>c.name===state.sortCol);
      if (col && col.type === 'number') { va = parseFloat(va)||0; vb = parseFloat(vb)||0; }
      else { va = String(va||'').toLowerCase(); vb = String(vb||'').toLowerCase(); }
      return state.sortDir === 'asc' ? (va > vb ? 1 : va < vb ? -1 : 0) : (va < vb ? 1 : va > vb ? -1 : 0);
    });
  }

  // Paginate
  const total = sorted.length;
  const pages = Math.ceil(total / state.pageSize) || 1;
  state.currentPage = Math.min(state.currentPage, pages);
  const start = (state.currentPage - 1) * state.pageSize;
  const pageData = sorted.slice(start, start + state.pageSize);

  $('table-count-info').textContent = `Showing ${start+1}–${Math.min(start+state.pageSize, total)} of ${total.toLocaleString()}`;

  // Build header
  const thead = $('table-head');
  thead.innerHTML = '';
  const tr = document.createElement('tr');
  cols.forEach(col => {
    const th = document.createElement('th');
    const isSorted = state.sortCol === col.name;
    th.className = isSorted ? (state.sortDir === 'asc' ? 'sorted-asc' : 'sorted-desc') : '';
    th.innerHTML = `${col.name} <span class="sort-icon"></span>`;
    th.title = `Click to sort by ${col.name}`;
    th.addEventListener('click', () => {
      if (state.sortCol === col.name) state.sortDir = state.sortDir === 'asc' ? 'desc' : 'asc';
      else { state.sortCol = col.name; state.sortDir = 'asc'; }
      renderTable();
    });
    tr.appendChild(th);
  });
  thead.appendChild(tr);

  // Build body
  const tbody = $('table-body');
  tbody.innerHTML = '';
  pageData.forEach(row => {
    const tr = document.createElement('tr');
    cols.forEach(col => {
      const td = document.createElement('td');
      const v = row[col.name];
      if (col.type === 'number') {
        td.className = 'cell-num';
        td.textContent = fmtNum(parseFloat(v));
      } else {
        const s = String(v ?? '');
        if (col.uniq && col.uniq.length <= 40 && s) {
          td.innerHTML = `<span class="cell-tag">${s}</span>`;
        } else {
          td.textContent = s;
          td.title = s;
        }
      }
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  // Pagination
  renderPagination(pages);
}

function buildColTogglePanel() {
  const panel = $('col-toggle-panel');
  panel.innerHTML = '';
  state.columns.forEach(col => {
    const item = document.createElement('label');
    item.className = 'col-toggle-item';
    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.checked = !state.hiddenCols.has(col.name);
    cb.addEventListener('change', () => {
      if (cb.checked) state.hiddenCols.delete(col.name);
      else state.hiddenCols.add(col.name);
      renderTable();
    });
    item.append(cb, document.createTextNode(col.name));
    panel.appendChild(item);
  });
}

function renderPagination(pages) {
  const pg = $('pagination');
  pg.innerHTML = '';
  if (pages <= 1) return;

  const addBtn = (label, page, disabled=false, active=false) => {
    const btn = document.createElement('button');
    btn.className = 'page-btn' + (active?' active':'');
    btn.innerHTML = label;
    btn.disabled = disabled;
    btn.addEventListener('click', () => { state.currentPage = page; renderTable(); });
    pg.appendChild(btn);
  };

  addBtn('<i class="fas fa-angles-left"></i>', 1, state.currentPage===1);
  addBtn('<i class="fas fa-angle-left"></i>', state.currentPage-1, state.currentPage===1);

  let start = Math.max(1, state.currentPage-2);
  let end   = Math.min(pages, start+4);
  start = Math.max(1, end-4);
  for (let p = start; p <= end; p++) addBtn(p, p, false, p===state.currentPage);

  addBtn('<i class="fas fa-angle-right"></i>', state.currentPage+1, state.currentPage===pages);
  addBtn('<i class="fas fa-angles-right"></i>', pages, state.currentPage===pages);
}

// ── Chart Builder ────────────────────────────────────────────
function buildChartBuilderSelects() {
  const xSel = $('cb-x'), ySel = $('cb-y');
  xSel.innerHTML = ''; ySel.innerHTML = '<option value="">— Count Only —</option>';
  state.columns.forEach(col => {
    const o1 = new Option(col.name, col.name);
    const o2 = new Option(col.name, col.name);
    xSel.add(o1);
    if (col.type === 'number') ySel.add(o2);
  });
}

function buildCustomChart() {
  const type = $('cb-type').value;
  const xCol = $('cb-x').value;
  const yCol = $('cb-y').value;
  const agg  = $('cb-agg').value;
  if (!xCol) return showToast('Please select an X axis column', 'error');

  const data = state.filteredData;
  if (state.customChart) { state.customChart.destroy(); state.customChart = null; }

  const canvas = $('custom-chart');
  const ctx = canvas.getContext('2d');

  if (type === 'scatter' && yCol) {
    const pts = data.map(r => ({ x: parseFloat(r[xCol]), y: parseFloat(r[yCol]) })).filter(p => !isNaN(p.x) && !isNaN(p.y));
    state.customChart = new Chart(ctx, {
      type: 'scatter',
      data: { datasets: [{ label: `${xCol} vs ${yCol}`, data: pts, backgroundColor: 'rgba(108,99,255,0.6)' }] },
      options: { ...baseChartOptions(), plugins: { legend: { display: false } }, scales: { x: scaleStyle(xCol), y: scaleStyle(yCol) } }
    });
    return;
  }

  const aggData = aggregate(data, xCol, yCol, agg);
  const sorted = Object.entries(aggData).sort((a,b) => b[1]-a[1]).slice(0, 30);
  const labels = sorted.map(e=>e[0]||'(blank)');
  const vals   = sorted.map(e=>+e[1].toFixed(2));
  const chartType = type === 'horizontalBar' ? 'bar' : type;
  const indexAxis = type === 'horizontalBar' ? 'y' : 'x';

  state.customChart = new Chart(ctx, {
    type: chartType,
    data: {
      labels,
      datasets: [{ label: yCol ? `${agg} of ${yCol}` : 'Count', data: vals, backgroundColor: COLORS, borderRadius: chartType==='bar'?6:0, borderWidth: chartType==='pie'||chartType==='doughnut'?2:0 }]
    },
    options: { ...baseChartOptions(), indexAxis, plugins: { legend: { display: chartType!=='bar' }, datalabels: { display: false } },
      scales: chartType==='pie'||chartType==='doughnut' ? {} : { x: scaleStyle(xCol), y: scaleStyle(yCol||'Count') } }
  });
}

function renderAutoCharts() {
  state.autoCharts.forEach(c => c.destroy());
  state.autoCharts = [];
  const grid = $('auto-charts-grid');
  grid.innerHTML = '';
  const data = state.filteredData;
  const numCols = state.columns.filter(c => c.type === 'number');
  const catCols = state.columns.filter(c => c.type === 'text' && c.uniq.length >= 2 && c.uniq.length <= 40);

  // For each cat col × num col combo (limited)
  catCols.slice(0,4).forEach(cat => {
    numCols.slice(0,4).forEach(num => {
      const sums = {};
      data.forEach(r => { const k=String(r[cat.name]||'(blank)'); const v=parseFloat(r[num.name]); if(!isNaN(v)){sums[k]=(sums[k]||0)+v;} });
      const entries = Object.entries(sums).sort((a,b)=>b[1]-a[1]).slice(0,12);
      if (!entries.length) return;
      const c = createChartCard(grid, `<i class="fas fa-chart-bar"></i> ${num.name} by ${cat.name}`, 280);
      const ch = new Chart(c, {
        type: 'bar',
        data: { labels: entries.map(e=>e[0]), datasets: [{ data: entries.map(e=>+e[1].toFixed(2)), backgroundColor: COLORS, borderRadius: 6 }] },
        options: chartOpts('bar', cat.name, `Sum of ${num.name}`)
      });
      state.autoCharts.push(ch);
    });
  });

  // Pie for each categorical
  catCols.slice(0,4).forEach(cat => {
    const freq = frequency(data, cat.name);
    const top = Object.entries(freq).sort((a,b)=>b[1]-a[1]).slice(0,10);
    if (!top.length) return;
    const c = createChartCard(grid, `<i class="fas fa-chart-pie"></i> ${cat.name} Share`, 280);
    const ch = new Chart(c, {
      type: 'doughnut',
      data: { labels: top.map(e=>e[0]||'(blank)'), datasets: [{ data: top.map(e=>e[1]), backgroundColor: COLORS, borderWidth: 2 }] },
      options: chartOpts('doughnut', cat.name, 'Count')
    });
    state.autoCharts.push(ch);
  });
}

// ── Comparison ───────────────────────────────────────────────
function buildComparisonSelects() {
  const grpSel = $('cmp-group'), metSel = $('cmp-metric');
  grpSel.innerHTML = ''; metSel.innerHTML = '<option value="">— Count Only —</option>';
  state.columns.forEach(col => {
    if (col.type === 'text') grpSel.add(new Option(col.name, col.name));
    if (col.type === 'number') metSel.add(new Option(col.name, col.name));
  });
}

function buildComparison() {
  const grpCol = $('cmp-group').value;
  const metCol = $('cmp-metric').value;
  const agg    = $('cmp-agg').value;
  if (!grpCol) return showToast('Please select a Group By column', 'error');

  const data = state.filteredData;
  const aggData = aggregate(data, grpCol, metCol, agg);
  const entries = Object.entries(aggData).sort((a,b)=>b[1]-a[1]);
  if (!entries.length) return showToast('No data to compare', 'error');

  const maxVal = Math.max(...entries.map(e=>e[1]));
  const output = $('comparison-output');
  output.innerHTML = '';

  // Table
  const wrap = document.createElement('div');
  wrap.className = 'comparison-table-wrap';
  const metLabel = metCol ? `${agg.charAt(0).toUpperCase()+agg.slice(1)} of ${metCol}` : 'Count';
  wrap.innerHTML = `<table class="comparison-table">
    <thead><tr><th>#</th><th>${grpCol}</th><th>${metLabel}</th><th style="min-width:160px">Proportion</th></tr></thead>
    <tbody>${entries.map((e,i) => {
      const pct = maxVal ? (e[1]/maxVal*100).toFixed(1) : 0;
      return `<tr>
        <td>${i+1}</td><td>${e[0]||'(blank)'}</td>
        <td class="cell-num">${fmtNum(e[1])}</td>
        <td><div class="bar-cell"><div class="mini-bar" style="width:${pct}%"></div><span style="font-size:0.78rem;color:var(--text-muted)">${pct}%</span></div></td>
      </tr>`;
    }).join('')}</tbody>
  </table>`;
  output.appendChild(wrap);

  // Chart
  const chartWrap = document.createElement('div');
  chartWrap.className = 'chart-card';
  const h4 = document.createElement('h4');
  h4.innerHTML = `<i class="fas fa-code-compare"></i> ${metLabel} by ${grpCol}`;
  const canvasWrap = document.createElement('div');
  canvasWrap.style.height = '320px'; canvasWrap.style.position = 'relative';
  const canvas = document.createElement('canvas');
  canvasWrap.appendChild(canvas); chartWrap.append(h4, canvasWrap); output.appendChild(chartWrap);

  const top20 = entries.slice(0,20);
  new Chart(canvas, {
    type: 'bar',
    data: { labels: top20.map(e=>e[0]||'(blank)'), datasets: [{ data: top20.map(e=>+e[1].toFixed(2)), backgroundColor: COLORS, borderRadius: 6 }] },
    options: chartOpts('bar', grpCol, metLabel)
  });
}

function renderAutoComparisons() {
  state.autoCompCharts.forEach(c => c.destroy());
  state.autoCompCharts = [];
  const grid = $('auto-comparison-grid');
  grid.innerHTML = '';
  const data = state.filteredData;
  const catCols = state.columns.filter(c => c.type === 'text' && c.uniq.length >= 2 && c.uniq.length <= 40);
  const numCols = state.columns.filter(c => c.type === 'number');

  // Multi-metric grouped bar per category
  catCols.slice(0,3).forEach(cat => {
    if (numCols.length < 2) return;
    const groups = [...new Set(data.map(r=>String(r[cat.name]||'(blank)')))].slice(0,15);
    const datasets = numCols.slice(0,4).map((num, idx) => {
      const vals = groups.map(g => {
        const rows = data.filter(r=>String(r[cat.name]||'(blank)')===g);
        const ns = rows.map(r=>parseFloat(r[num.name])).filter(v=>!isNaN(v));
        return ns.length ? +(ns.reduce((a,b)=>a+b,0)/ns.length).toFixed(2) : 0;
      });
      return { label: `Avg ${num.name}`, data: vals, backgroundColor: COLORS[idx], borderRadius: 4 };
    });
    const c = createChartCard(grid, `<i class="fas fa-layer-group"></i> Multi-Metric by ${cat.name}`, 300);
    const ch = new Chart(c, {
      type: 'bar',
      data: { labels: groups, datasets },
      options: { ...baseChartOptions(), plugins: { legend: { labels: { color: '#8892b0', font:{size:11} } } }, scales: { x: scaleStyle(cat.name), y: scaleStyle('Value') } }
    });
    state.autoCompCharts.push(ch);
  });

  // Stacked comparison
  if (catCols.length >= 2) {
    const cat1 = catCols[0], cat2 = catCols[1];
    const keys1 = [...new Set(data.map(r=>String(r[cat1.name]||'(blank)')))].slice(0,10);
    const keys2 = [...new Set(data.map(r=>String(r[cat2.name]||'(blank)')))].slice(0,8);
    const datasets = keys2.map((k2,i) => ({
      label: k2,
      data: keys1.map(k1 => data.filter(r=>String(r[cat1.name]||'(blank)')===k1 && String(r[cat2.name]||'(blank)')===k2).length),
      backgroundColor: COLORS[i], borderRadius: 4
    }));
    const c = createChartCard(grid, `<i class="fas fa-chart-bar"></i> ${cat1.name} × ${cat2.name}`, 300);
    const ch = new Chart(c, {
      type: 'bar',
      data: { labels: keys1, datasets },
      options: { ...baseChartOptions(), scales: { x: { ...scaleStyle(cat1.name), stacked:true }, y: { ...scaleStyle('Count'), stacked:true } }, plugins: { legend: { labels: { color:'#8892b0', font:{size:10} } } } }
    });
    state.autoCompCharts.push(ch);
  }
}

// ── Chart Helpers ────────────────────────────────────────────
function createChartCard(parent, titleHTML, height=280) {
  const card = document.createElement('div');
  card.className = 'chart-card';
  const h4 = document.createElement('h4');
  h4.innerHTML = titleHTML;
  const wrap = document.createElement('div');
  wrap.className = 'chart-canvas-wrap';
  wrap.style.height = height + 'px';
  const canvas = document.createElement('canvas');
  wrap.appendChild(canvas); card.append(h4, wrap); parent.appendChild(card);
  return canvas;
}

function baseChartOptions() {
  return {
    responsive: true, maintainAspectRatio: false,
    plugins: {
      legend: { display: false },
      tooltip: { backgroundColor: '#1a1d27', titleColor: '#e8eaf6', bodyColor: '#8892b0', borderColor: '#2e3250', borderWidth: 1 }
    }
  };
}
function scaleStyle(label='') {
  return { grid: { color: 'rgba(46,50,80,0.7)' }, ticks: { color: '#8892b0', font:{size:11} }, title: { display:!!label, text:label, color:'#8892b0', font:{size:11} } };
}
function chartOpts(type, xLabel='', yLabel='', showLegend=false) {
  if (type==='pie'||type==='doughnut') {
    return { ...baseChartOptions(), plugins: { ...baseChartOptions().plugins, legend: { display:true, position:'right', labels:{color:'#8892b0',boxWidth:12,font:{size:11}} } } };
  }
  return { ...baseChartOptions(), plugins: { ...baseChartOptions().plugins, legend: { display:showLegend } }, scales: { x: scaleStyle(xLabel), y: scaleStyle(yLabel) } };
}

// ── Data Helpers ─────────────────────────────────────────────
function frequency(data, col) {
  const f = {};
  data.forEach(r => { const k = String(r[col]??'(blank)'); f[k] = (f[k]||0)+1; });
  return f;
}

function aggregate(data, groupCol, valCol, agg) {
  const groups = {};
  const counts = {};
  data.forEach(r => {
    const k = String(r[groupCol]??'(blank)');
    const v = valCol ? parseFloat(r[valCol]) : 1;
    if (!groups[k]) { groups[k] = []; counts[k] = 0; }
    if (!isNaN(v)) groups[k].push(v);
    counts[k]++;
  });
  const result = {};
  Object.entries(groups).forEach(([k, vals]) => {
    if (!valCol) { result[k] = counts[k]; return; }
    if (!vals.length) { result[k] = 0; return; }
    switch (agg) {
      case 'sum':   result[k] = vals.reduce((a,b)=>a+b,0); break;
      case 'avg':   result[k] = vals.reduce((a,b)=>a+b,0)/vals.length; break;
      case 'count': result[k] = counts[k]; break;
      case 'min':   result[k] = Math.min(...vals); break;
      case 'max':   result[k] = Math.max(...vals); break;
      default:      result[k] = vals.reduce((a,b)=>a+b,0);
    }
  });
  return result;
}

function histogram(vals, bins=10) {
  const min = Math.min(...vals), max = Math.max(...vals);
  const step = (max-min)/bins;
  const counts = Array(bins).fill(0);
  const labels = [];
  for (let i=0;i<bins;i++) labels.push(`${fmtNum(min+i*step)}–${fmtNum(min+(i+1)*step)}`);
  vals.forEach(v => {
    const idx = Math.min(Math.floor((v-min)/step), bins-1);
    counts[idx]++;
  });
  return { labels, counts };
}

// ── Export ───────────────────────────────────────────────────
function exportCSV() {
  const data = state.filteredData;
  if (!data.length) return showToast('No data to export', 'error');
  const cols = state.columns.map(c=>c.name);
  const rows = [cols.join(','), ...data.map(r => cols.map(c => JSON.stringify(r[c]??'')).join(','))];
  const blob = new Blob([rows.join('\n')], { type: 'text/csv' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = 'filtered_data.csv'; a.click();
  URL.revokeObjectURL(url);
  showToast(`Exported ${data.length.toLocaleString()} rows`, 'success');
}

// ── Utility ──────────────────────────────────────────────────
function fmtNum(n) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  if (Math.abs(n) >= 1e9) return (n/1e9).toFixed(2)+'B';
  if (Math.abs(n) >= 1e6) return (n/1e6).toFixed(2)+'M';
  if (Math.abs(n) >= 1e3) return (n/1e3).toFixed(2)+'K';
  if (Number.isInteger(n)) return n.toLocaleString();
  return parseFloat(n.toFixed(2)).toLocaleString();
}

function debounce(fn, ms) {
  let t; return (...args) => { clearTimeout(t); t = setTimeout(()=>fn(...args), ms); };
}

function showToast(msg, type='') {
  const toast = $('toast');
  toast.textContent = msg;
  toast.className = 'toast' + (type ? ' '+type : '');
  toast.classList.remove('hidden');
  setTimeout(() => toast.classList.add('hidden'), 3000);
}
