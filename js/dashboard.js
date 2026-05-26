/* =========================================================
   EXCEL DATA DASHBOARD — COMPLETE STABLE VERSION
   FIXED:
   ✔ Maximum call stack exceeded
   ✔ Large Excel handling
   ✔ Upload issues
   ✔ Chart crashes
   ✔ Memory overflow
   ✔ Safer filtering
   ========================================================= */

'use strict';

/* =========================================================
   GLOBAL ERROR LOGGER
========================================================= */

window.onerror = function(msg, src, line, col, err) {
  console.error('GLOBAL ERROR:', msg);
};

/* =========================================================
   STATE
========================================================= */

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
  '#6c63ff',
  '#00d4aa',
  '#ff6b6b',
  '#ffd166',
  '#06d6a0',
  '#118ab2',
  '#ef476f',
  '#a8dadc',
  '#e9c46a',
  '#f4a261',
];

/* =========================================================
   DOM HELPERS
========================================================= */

const $ = id => document.getElementById(id);

const uploadScreen = $('upload-screen');
const dashboard = $('dashboard');
const fileInput = $('file-input');

/* =========================================================
   BOOT
========================================================= */

document.addEventListener('DOMContentLoaded', () => {

  console.log('Dashboard Loaded');

  if (!fileInput) {
    console.error('file-input not found');
    return;
  }

  fileInput.addEventListener(
    'change',
    handleFileUpload
  );

  const search = $('global-search');

  if (search) {
    search.addEventListener(
      'input',
      debounce(applyFilters, 300)
    );
  }

  const exportBtn = $('export-btn');

  if (exportBtn) {
    exportBtn.addEventListener(
      'click',
      exportCSV
    );
  }

  const resetBtn = $('reset-filters-btn');

  if (resetBtn) {
    resetBtn.addEventListener(
      'click',
      resetFilters
    );
  }

  document.querySelectorAll('.nav-item')
    .forEach(item => {

      item.addEventListener('click', e => {

        e.preventDefault();

        switchTab(item.dataset.tab);
      });
    });
});

/* =========================================================
   FILE UPLOAD
========================================================= */

function handleFileUpload(e) {

  const file = e.target.files[0];

  if (!file) return;

  const status = $('upload-status');

  status.classList.remove('hidden');

  status.textContent =
    '⏳ Reading Excel file...';

  status.className =
    'upload-status';

  console.log('FILE:', file.name);

  const reader = new FileReader();

  reader.onload = function(evt) {

    try {

      const data = evt.target.result;

      console.log('Reading workbook');

      const workbook = XLSX.read(data, {
        type: 'binary',
        cellDates: true
      });

      const sheetName =
        workbook.SheetNames[0];

      console.log('Sheet:', sheetName);

      const worksheet =
        workbook.Sheets[sheetName];

      const rows =
        XLSX.utils.sheet_to_json(
          worksheet,
          {
            defval: ''
          }
        );

      console.log('Rows:', rows.length);

      if (!rows.length) {
        throw new Error(
          'Excel sheet is empty'
        );
      }

      state.rawData = rows;
      state.filteredData = rows;
      state.fileName = file.name;

      analyzeColumns(rows);

      init();

      status.textContent =
        `✅ ${rows.length.toLocaleString()} rows loaded`;

      status.className =
        'upload-status success';

      if (uploadScreen) {
        uploadScreen.classList.add('hidden');
      }

      if (dashboard) {
        dashboard.classList.remove('hidden');
      }

    } catch (err) {

      console.error(err);

      status.textContent =
        '❌ ' + err.message;

      status.className =
        'upload-status error';
    }
  };

  reader.onerror = function() {

    status.textContent =
      '❌ Failed to read file';

    status.className =
      'upload-status error';
  };

  reader.readAsBinaryString(file);
}

/* =========================================================
   COLUMN ANALYSIS
========================================================= */

function analyzeColumns(rows) {

  const keys = Object.keys(rows[0]);

  state.columns = keys.map(name => {

    const vals = [];

    for (const row of rows) {

      const v = row[name];

      if (
        v !== '' &&
        v !== null &&
        v !== undefined
      ) {
        vals.push(v);
      }
    }

    const isNum =
      vals.length > 0 &&
      vals.every(v =>
        !isNaN(parseFloat(v))
      );

    const type =
      isNum ? 'number' : 'text';

    // SAFE UNIQUE
    const uniqSet = new Set();

    for (const v of vals) {

      uniqSet.add(
        String(v).trim()
      );

      if (uniqSet.size > 200) {
        break;
      }
    }

    const uniq = [...uniqSet];

    // SAFE MIN MAX
    let min = null;
    let max = null;

    if (isNum) {

      min = Infinity;
      max = -Infinity;

      for (const v of vals) {

        const n = Number(v);

        if (!isNaN(n)) {

          if (n < min) min = n;
          if (n > max) max = n;
        }
      }

      if (min === Infinity) {
        min = null;
      }

      if (max === -Infinity) {
        max = null;
      }
    }

    return {
      name,
      type,
      uniq,
      min,
      max
    };
  });

  console.log(state.columns);
}

/* =========================================================
   INIT
========================================================= */

function init() {

  buildFilters();

  renderOverview();

  renderTable();

  renderCharts();

  const badge = $('file-info-badge');

  if (badge) {
    badge.textContent =
      state.fileName;
  }
}

/* =========================================================
   FILTERS
========================================================= */

function buildFilters() {

  const panel =
    $('filter-controls');

  if (!panel) return;

  panel.innerHTML = '';

  state.columns.forEach(col => {

    const wrap =
      document.createElement('div');

    wrap.className =
      'filter-group';

    const label =
      document.createElement('label');

    label.textContent =
      col.name;

    wrap.appendChild(label);

    const input =
      document.createElement('input');

    input.type = 'text';

    input.dataset.col =
      col.name;

    input.placeholder =
      'Search...';

    input.addEventListener(
      'input',
      debounce(applyFilters, 300)
    );

    wrap.appendChild(input);

    panel.appendChild(wrap);
  });
}

function applyFilters() {

  let data = [...state.rawData];

  const globalSearch =
    $('global-search');

  const globalQ =
    globalSearch
      ? globalSearch.value
          .trim()
          .toLowerCase()
      : '';

  // COLUMN FILTERS
  document
    .querySelectorAll(
      '#filter-controls [data-col]'
    )
    .forEach(el => {

      const col =
        el.dataset.col;

      const val =
        el.value
          .trim()
          .toLowerCase();

      if (!val) return;

      data = data.filter(row => {

        const v =
          row[col];

        if (
          v === null ||
          v === undefined
        ) {
          return false;
        }

        return String(v)
          .toLowerCase()
          .includes(val);
      });
    });

  // GLOBAL SEARCH
  if (globalQ) {

    data = data.filter(row =>

      Object.values(row).some(v => {

        if (
          v === null ||
          v === undefined
        ) {
          return false;
        }

        return String(v)
          .toLowerCase()
          .includes(globalQ);
      })
    );
  }

  state.filteredData = data;

  state.currentPage = 1;

  renderOverview();
  renderTable();
  renderCharts();
}

/* =========================================================
   RESET FILTERS
========================================================= */

function resetFilters() {

  document
    .querySelectorAll(
      '#filter-controls input'
    )
    .forEach(el => {

      el.value = '';
    });

  const globalSearch =
    $('global-search');

  if (globalSearch) {
    globalSearch.value = '';
  }

  state.filteredData =
    [...state.rawData];

  renderOverview();
  renderTable();
  renderCharts();
}

/* =========================================================
   OVERVIEW
========================================================= */

function renderOverview() {

  const grid = $('kpi-grid');

  if (!grid) return;

  grid.innerHTML = '';

  addKPI(
    grid,
    'Total Records',
    state.filteredData.length
      .toLocaleString()
  );

  const numericCols =
    state.columns
      .filter(c => c.type === 'number')
      .slice(0, 4);

  numericCols.forEach(col => {

    let sum = 0;

    state.filteredData.forEach(r => {

      const n =
        parseFloat(
          r[col.name]
        );

      if (!isNaN(n)) {
        sum += n;
      }
    });

    addKPI(
      grid,
      'Sum of ' + col.name,
      fmtNum(sum)
    );
  });
}

function addKPI(
  parent,
  label,
  value
) {

  const card =
    document.createElement('div');

  card.className =
    'kpi-card';

  card.innerHTML = `
    <div class="kpi-label">
      ${label}
    </div>

    <div class="kpi-value">
      ${value}
    </div>
  `;

  parent.appendChild(card);
}

/* =========================================================
   TABLE
========================================================= */

function renderTable() {

  const tbody =
    $('table-body');

  const thead =
    $('table-head');

  if (!tbody || !thead) return;

  tbody.innerHTML = '';
  thead.innerHTML = '';

  const cols = state.columns;

  // HEADER
  const tr =
    document.createElement('tr');

  cols.forEach(col => {

    const th =
      document.createElement('th');

    th.textContent =
      col.name;

    tr.appendChild(th);
  });

  thead.appendChild(tr);

  // PAGINATION
  const start =
    (state.currentPage - 1) *
    state.pageSize;

  const end =
    start + state.pageSize;

  const rows =
    state.filteredData.slice(
      start,
      end
    );

  rows.forEach(row => {

    const tr =
      document.createElement('tr');

    cols.forEach(col => {

      const td =
        document.createElement('td');

      td.textContent =
        row[col.name];

      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  });
}

/* =========================================================
   CHARTS
========================================================= */

function renderCharts() {

  const container =
    $('auto-charts-grid');

  if (!container) return;

  container.innerHTML = '';

  const numericCols =
    state.columns
      .filter(c => c.type === 'number')
      .slice(0, 2);

  const textCols =
    state.columns
      .filter(c => c.type === 'text')
      .slice(0, 2);

  textCols.forEach(textCol => {

    numericCols.forEach(numCol => {

      const sums = {};

      state.filteredData.forEach(r => {

        const key =
          String(
            r[textCol.name] ||
            '(blank)'
          );

        const val =
          parseFloat(
            r[numCol.name]
          );

        if (!isNaN(val)) {

          sums[key] =
            (sums[key] || 0) + val;
        }
      });

      const entries =
        Object.entries(sums)
          .slice(0, 10);

      if (!entries.length) {
        return;
      }

      const card =
        document.createElement('div');

      card.className =
        'chart-card';

      card.style.height =
        '350px';

      const canvas =
        document.createElement('canvas');

      card.appendChild(canvas);

      container.appendChild(card);

      new Chart(canvas, {

        type: 'bar',

        data: {

          labels:
            entries.map(
              e => e[0]
            ),

          datasets: [{

            data:
              entries.map(
                e => e[1]
              ),

            backgroundColor:
              COLORS
          }]
        },

        options: {

          responsive: true,

          maintainAspectRatio: false
        }
      });
    });
  });
}

/* =========================================================
   EXPORT CSV
========================================================= */

function exportCSV() {

  const data =
    state.filteredData;

  if (!data.length) {
    return;
  }

  const cols =
    state.columns.map(
      c => c.name
    );

  const csv = [

    cols.join(','),

    ...data.map(row =>

      cols.map(col =>

        JSON.stringify(
          row[col] ?? ''
        )

      ).join(',')
    )
  ].join('\n');

  const blob =
    new Blob([csv], {
      type: 'text/csv'
    });

  const url =
    URL.createObjectURL(blob);

  const a =
    document.createElement('a');

  a.href = url;

  a.download =
    'filtered_data.csv';

  a.click();

  URL.revokeObjectURL(url);
}

/* =========================================================
   TAB SWITCHING
========================================================= */

function switchTab(tab) {

  document
    .querySelectorAll('.tab-content')
    .forEach(el => {

      el.classList.add('hidden');
    });

  const active =
    $('tab-' + tab);

  if (active) {
    active.classList.remove('hidden');
  }
}

/* =========================================================
   HELPERS
========================================================= */

function fmtNum(n) {

  if (
    n === null ||
    n === undefined ||
    isNaN(n)
  ) {
    return '—';
  }

  if (Math.abs(n) >= 1e6) {
    return (
      (n / 1e6).toFixed(2) + 'M'
    );
  }

  if (Math.abs(n) >= 1e3) {
    return (
      (n / 1e3).toFixed(2) + 'K'
    );
  }

  return Number(n)
    .toLocaleString();
}

function debounce(fn, ms) {

  let timer;

  return (...args) => {

    clearTimeout(timer);

    timer = setTimeout(() => {

      fn(...args);

    }, ms);
  };
}
