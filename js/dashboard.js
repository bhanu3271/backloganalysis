/* =========================================================
EXCEL DATA DASHBOARD — OPTIMIZED VERSION
FIXED: Maximum call stack size exceeded
========================================================= */

// ── State ──────────────────────────────────────────────────
const state = {
rawData: [],
filteredData: [],
columns: [],
fileName: '',
currentTab: 'overview',
currentPage: 1,
pageSize: 25,
sortCol: null,
sortDir: 'asc',
hiddenCols: new Set(),
filters: {},
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
const dashboard = $('dashboard');
const fileInput = $('file-input');
const sidebar = document.querySelector('.sidebar');

// ── Boot ────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
fileInput.addEventListener('change', handleFileUpload);

$('change-file-btn').addEventListener('click', () => {
fileInput.value = '';
fileInput.click();
});

$('sidebar-toggle').addEventListener('click', () => {
sidebar.classList.toggle('collapsed');
});

$('global-search').addEventListener(
'input',
debounce(applyFilters, 250)
);

$('export-btn').addEventListener('click', exportCSV);
$('reset-filters-btn').addEventListener('click', resetFilters);
$('toggle-filters-btn').addEventListener('click', toggleFilterPanel);

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

// FILE SIZE LIMIT
if (file.size > 20 * 1024 * 1024) {
status.textContent = '❌ File too large. Upload below 20MB';
status.className = 'upload-status error';
status.classList.remove('hidden');
return;
}

status.textContent = '⏳ Reading file...';
status.className = 'upload-status';
status.classList.remove('hidden');

const reader = new FileReader();

reader.onload = ev => {
try {
const wb = XLSX.read(ev.target.result, {
type: 'array',
cellDates: true,
});

```
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];

  const rows = XLSX.utils.sheet_to_json(ws, {
    defval: '',
  });

  if (!rows.length) {
    throw new Error('Sheet is empty');
  }

  // ROW LIMIT
  if (rows.length > 50000) {
    throw new Error('Maximum supported rows: 50,000');
  }

  state.rawData = rows;
  state.fileName = file.name;

  analyzeColumns(rows);

  init();

  status.textContent = `✅ ${rows.length.toLocaleString()} rows loaded`;
  status.className = 'upload-status success';

  uploadScreen.classList.add('hidden');
  dashboard.classList.remove('hidden');

} catch (err) {
  console.error(err);

  status.textContent = '❌ Error: ' + err.message;
  status.className = 'upload-status error';
}
```

};

reader.readAsArrayBuffer(file);
}

// ── Column Analysis ─────────────────────────────────────────
function analyzeColumns(rows) {

const keys = Object.keys(rows[0]);

state.columns = keys.map(name => {

```
const vals = [];

for (const row of rows) {
  const v = row[name];

  if (v !== '' && v !== null && v !== undefined) {
    vals.push(v);
  }
}

const isNum =
  vals.length > 0 &&
  vals.every(v => !isNaN(parseFloat(v)) && isFinite(v));

const isDate =
  !isNum &&
  vals.some(v =>
    v instanceof Date ||
    (
      typeof v === 'string' &&
      /\d{1,4}[-/]\d{1,2}[-/]\d{1,4}/.test(v)
    )
  );

const type = isNum
  ? 'number'
  : isDate
  ? 'date'
  : 'text';

// SAFE UNIQUE VALUES
const uniqSet = new Set();

for (const v of vals) {
  uniqSet.add(String(v).trim());

  if (uniqSet.size > 500) break;
}

const uniq = [...uniqSet].sort((a, b) => {
  if (
    !isNaN(parseFloat(a)) &&
    !isNaN(parseFloat(b))
  ) {
    return parseFloat(a) - parseFloat(b);
  }

  return a.localeCompare(b);
});

// SAFE MIN MAX
let min = null;
let max = null;

if (isNum) {
  min = Infinity;
  max = -Infinity;

  for (const v of vals) {
    const num = Number(v);

    if (!isNaN(num)) {
      if (num < min) min = num;
      if (num > max) max = num;
    }
  }

  if (min === Infinity) min = null;
  if (max === -Infinity) max = null;
}

return {
  name,
  type,
  uniq,
  min,
  max,
};
```

});
}

// ── Init ────────────────────────────────────────────────────
function init() {
$('file-info-badge').textContent = state.fileName;

buildFilters();
applyFilters();
buildChartBuilderSelects();
buildComparisonSelects();
}

// ── Filters ─────────────────────────────────────────────────
function applyFilters() {

const globalQ = $('global-search')
.value
.trim()
.toLowerCase();

let data = state.rawData;

// SAFE GLOBAL SEARCH
if (globalQ) {
data = data.filter(row =>
Object.values(row).some(v => {

```
    if (v === null || v === undefined) {
      return false;
    }

    return String(v)
      .toLowerCase()
      .includes(globalQ);
  })
);
```

}

state.filteredData = data;
state.currentPage = 1;

renderCurrentTab();
}

// ── Tabs ────────────────────────────────────────────────────
function switchTab(tab) {

state.currentTab = tab;

document
.querySelectorAll('.tab-content')
.forEach(el => el.classList.add('hidden'));

document
.querySelectorAll('.nav-item')
.forEach(el => el.classList.remove('active'));

$('tab-' + tab).classList.remove('hidden');

document
.querySelector(`.nav-item[data-tab="${tab}"]`)
.classList.add('active');

renderCurrentTab();
}

function renderCurrentTab() {

switch (state.currentTab) {

```
case 'overview':
  renderOverview();
  break;

case 'data-table':
  renderTable();
  break;

case 'charts':
  renderAutoCharts();
  break;

case 'comparison':
  renderAutoComparisons();
  break;
```

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

addKPI(
grid,
'fas fa-database',
'Total Records',
data.length.toLocaleString(),
'Filtered Rows'
);

const numCols = state.columns
.filter(c => c.type === 'number')
.slice(0, 4);

numCols.forEach(col => {

```
let sum = 0;
let count = 0;

data.forEach(r => {
  const v = parseFloat(r[col.name]);

  if (!isNaN(v)) {
    sum += v;
    count++;
  }
});

const avg = count ? sum / count : 0;

addKPI(
  grid,
  'fas fa-chart-line',
  col.name,
  fmtNum(sum),
  `Avg: ${fmtNum(avg)}`
);
```

});
}

function addKPI(grid, icon, label, value, sub) {

const card = document.createElement('div');

card.className = 'kpi-card';

card.innerHTML = ` <div class="kpi-icon"> <i class="${icon}"></i> </div>

```
<div class="kpi-label">${label}</div>

<div class="kpi-value">${value}</div>

<div class="kpi-sub">${sub}</div>
```

`;

grid.appendChild(card);
}

// ── Table ───────────────────────────────────────────────────
function renderTable() {

const tbody = $('table-body');
const thead = $('table-head');

tbody.innerHTML = '';
thead.innerHTML = '';

const cols = state.columns;

// TABLE HEADER
const headerRow = document.createElement('tr');

cols.forEach(col => {

```
const th = document.createElement('th');
th.textContent = col.name;

headerRow.appendChild(th);
```

});

thead.appendChild(headerRow);

// PAGINATION
const start = (state.currentPage - 1) * state.pageSize;
const end = start + state.pageSize;

const rows = state.filteredData.slice(start, end);

rows.forEach(row => {

```
const tr = document.createElement('tr');

cols.forEach(col => {

  const td = document.createElement('td');

  td.textContent = row[col.name];

  tr.appendChild(td);
});

tbody.appendChild(tr);
```

});
}

// ── Auto Charts ─────────────────────────────────────────────
function renderAutoCharts() {

state.autoCharts.forEach(c => c.destroy());
state.autoCharts = [];

const grid = $('auto-charts-grid');
grid.innerHTML = '';

const data = state.filteredData;

const numCols = state.columns
.filter(c => c.type === 'number')
.slice(0, 2);

const catCols = state.columns
.filter(c => c.type === 'text' && c.uniq.length <= 40)
.slice(0, 2);

catCols.forEach(cat => {

```
numCols.forEach(num => {

  const sums = {};

  data.forEach(r => {

    const key = String(r[cat.name] || '(blank)');
    const val = parseFloat(r[num.name]);

    if (!isNaN(val)) {
      sums[key] = (sums[key] || 0) + val;
    }
  });

  const entries = Object.entries(sums)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10);

  if (!entries.length) return;

  const card = document.createElement('div');
  card.className = 'chart-card';

  const canvas = document.createElement('canvas');

  card.appendChild(canvas);
  grid.appendChild(card);

  const chart = new Chart(canvas, {
    type: 'bar',

    data: {
      labels: entries.map(e => e[0]),

      datasets: [{
        data: entries.map(e => e[1]),
        backgroundColor: COLORS,
      }]
    },

    options: {
      responsive: true,
      maintainAspectRatio: false,
    }
  });

  state.autoCharts.push(chart);
});
```

});
}

// ── Histogram ───────────────────────────────────────────────
function histogram(vals, bins = 10) {

if (!vals.length) {
return {
labels: [],
counts: []
};
}

let min = Infinity;
let max = -Infinity;

for (const v of vals) {
if (v < min) min = v;
if (v > max) max = v;
}

if (min === max) {
return {
labels: [fmtNum(min)],
counts: [vals.length]
};
}

const step = (max - min) / bins;

const counts = Array(bins).fill(0);
const labels = [];

for (let i = 0; i < bins; i++) {

```
labels.push(
  `${fmtNum(min + i * step)}–${fmtNum(min + (i + 1) * step)}`
);
```

}

vals.forEach(v => {

```
let idx = Math.floor((v - min) / step);

if (idx >= bins) idx = bins - 1;
if (idx < 0) idx = 0;

counts[idx]++;
```

});

return {
labels,
counts
};
}

// ── Export CSV ──────────────────────────────────────────────
function exportCSV() {

const data = state.filteredData;

if (!data.length) {
showToast('No data to export', 'error');
return;
}

const cols = state.columns.map(c => c.name);

const rows = [
cols.join(','),

```
...data.map(r =>
  cols
    .map(c => JSON.stringify(r[c] ?? ''))
    .join(',')
)
```

];

const blob = new Blob([
rows.join('\n')
], {
type: 'text/csv'
});

const url = URL.createObjectURL(blob);

const a = document.createElement('a');

a.href = url;
a.download = 'filtered_data.csv';
a.click();

URL.revokeObjectURL(url);

showToast('CSV Exported', 'success');
}

// ── Utilities ───────────────────────────────────────────────
function fmtNum(n) {

if (
n === null ||
n === undefined ||
isNaN(n)
) {
return '—';
}

if (Math.abs(n) >= 1e9) {
return (n / 1e9).toFixed(2) + 'B';
}

if (Math.abs(n) >= 1e6) {
return (n / 1e6).toFixed(2) + 'M';
}

if (Math.abs(n) >= 1e3) {
return (n / 1e3).toFixed(2) + 'K';
}

if (Number.isInteger(n)) {
return n.toLocaleString();
}

return parseFloat(n.toFixed(2)).toLocaleString();
}

function debounce(fn, ms) {

let t;

return (...args) => {
clearTimeout(t);

```
t = setTimeout(() => {
  fn(...args);
}, ms);
```

};
}

function showToast(msg, type = '') {

const toast = $('toast');

toast.textContent = msg;

toast.className = 'toast' + (type ? ' ' + type : '');

toast.classList.remove('hidden');

setTimeout(() => {
toast.classList.add('hidden');
}, 3000);
}
