const Dashboard = (() => {
  let dashboards = [];
  let currentDashboardId = '';
  let currentData = null;
  let chartInstances = [];

  function destroyCharts() {
    chartInstances.forEach((chart) => {
      try {
        chart.destroy();
      } catch (err) {}
    });
    chartInstances = [];
  }

  async function loadDashboards() {
    const result = await Api.dashboards();
    dashboards = result.dashboards || [];

    const select = document.getElementById('dashboardSelect');
    select.innerHTML = '';

    if (dashboards.length === 0) {
      select.innerHTML = '<option value="">ไม่พบ Dashboard</option>';
      return;
    }

    dashboards.forEach((db) => {
      const opt = document.createElement('option');
      opt.value = db.dashboardId;
      opt.textContent = db.dashboardName;
      select.appendChild(opt);
    });

    currentDashboardId = dashboards[0].dashboardId;
    select.value = currentDashboardId;
  }

  function getSelectedDashboardId() {
    const select = document.getElementById('dashboardSelect');
    return select ? select.value : currentDashboardId;
  }

  function buildParamsFromFilters() {
    const params = {};

    const dateFrom = document.getElementById('dateFrom')?.value || '';
    const dateTo = document.getElementById('dateTo')?.value || '';

    if (dateFrom) params.dateFrom = dateFrom;
    if (dateTo) params.dateTo = dateTo;

    const dynamicInputs = document.querySelectorAll('[data-filter-field]');
    dynamicInputs.forEach((el) => {
      const fieldKey = el.getAttribute('data-filter-field');
      const value = el.value;

      if (fieldKey && value && value !== 'ALL') {
        params[`filter_${fieldKey}`] = value;
      }
    });

    params.tableLimit = 500;

    return params;
  }

  async function loadDashboardData() {
    const dashboardId = getSelectedDashboardId();

    if (!dashboardId) {
      return;
    }

    currentDashboardId = dashboardId;
    setLoading(true);

    try {
      const params = buildParamsFromFilters();
      const data = await Api.dashboardData(dashboardId, params);
      currentData = data;

      renderDashboard(data);
    } catch (err) {
      Swal.fire({
        icon: 'error',
        title: 'โหลด Dashboard ไม่สำเร็จ',
        text: err.message || String(err)
      });
    } finally {
      setLoading(false);
    }
  }

  function renderDashboard(data) {
    renderMeta(data);
    renderDynamicFilters(data);
    renderMetrics(data.metrics || []);
    renderCharts(data.charts || []);
    renderTable(data.table || { columns: [], rows: [] });
  }

  function renderMeta(data) {
    const title = document.getElementById('dashboardTitle');
    const meta = document.getElementById('dashboardMeta');

    const dbName = data.dashboard?.dashboardName || '-';
    const sourceName = data.source?.sourceName || '-';
    const totalRows = data.meta?.totalRows ?? 0;
    const filteredRows = data.meta?.filteredRows ?? 0;
    const generatedAt = data.meta?.generatedAt || '-';

    if (title) title.textContent = dbName;
    if (meta) {
      meta.textContent = `แหล่งข้อมูล: ${sourceName} | ทั้งหมด ${formatNumber(totalRows)} แถว | หลังกรอง ${formatNumber(filteredRows)} แถว | อัปเดต ${generatedAt}`;
    }
  }

  function renderDynamicFilters(data) {
    const wrap = document.getElementById('dynamicFilters');
    if (!wrap) return;

    const configs = data.filters?.config || [];
    const optionsMap = data.filters?.options || {};

    wrap.innerHTML = '';

    configs.forEach((filter) => {
      if (filter.filterType === 'dateRange') return;

      const label = document.createElement('label');
      const span = document.createElement('span');
      span.textContent = filter.label || filter.fieldKey;

      let input;

      if (filter.filterType === 'select' || filter.filterType === 'multiSelect' || filter.filterType === 'status') {
        input = document.createElement('select');
        input.setAttribute('data-filter-field', filter.fieldKey);

        const defaultOpt = document.createElement('option');
        defaultOpt.value = 'ALL';
        defaultOpt.textContent = 'ทั้งหมด';
        input.appendChild(defaultOpt);

        const options = optionsMap[filter.fieldKey] || [];
        options.forEach((item) => {
          const opt = document.createElement('option');
          opt.value = item.value;
          opt.textContent = `${item.label} (${item.count})`;
          input.appendChild(opt);
        });
      } else {
        input = document.createElement('input');
        input.type = filter.filterType === 'numberRange' ? 'text' : 'text';
        input.placeholder = filter.filterType === 'numberRange' ? 'เช่น 10:100' : 'ค้นหา...';
        input.setAttribute('data-filter-field', filter.fieldKey);
      }

      label.appendChild(span);
      label.appendChild(input);
      wrap.appendChild(label);
    });
  }

  function renderMetrics(metrics) {
    const grid = document.getElementById('metricGrid');
    if (!grid) return;

    grid.innerHTML = '';

    if (!metrics.length) {
      grid.innerHTML = `<div class="emptyState">ไม่พบ KPI ที่เปิดใช้งาน</div>`;
      return;
    }

    metrics.forEach((metric) => {
      const card = document.createElement('article');
      card.className = 'metricCard';

      card.innerHTML = `
        <div class="label">${escapeHtml(metric.title || '-')}</div>
        <div class="value">${escapeHtml(String(metric.displayValue ?? metric.value ?? '-'))}</div>
        <div class="sub">${escapeHtml(metric.metricType || '')}${metric.fieldKey ? ' · ' + escapeHtml(metric.fieldKey) : ''}</div>
      `;

      grid.appendChild(card);
    });
  }

  function renderCharts(charts) {
    destroyCharts();

    const grid = document.getElementById('chartGrid');
    if (!grid) return;

    grid.innerHTML = '';

    if (!charts.length) {
      grid.innerHTML = `<div class="emptyState">ไม่พบกราฟที่เปิดใช้งาน</div>`;
      return;
    }

    charts.forEach((chart, index) => {
      const card = document.createElement('article');
      card.className = 'chartCard';

      const canvasId = `chart_${chart.chartId || index}`;

      card.innerHTML = `
        <h3>${escapeHtml(chart.title || '-')}</h3>
        <div class="chartBox">
          <canvas id="${canvasId}"></canvas>
        </div>
      `;

      grid.appendChild(card);

      const ctx = document.getElementById(canvasId);
      if (!ctx) return;

      const chartType = normalizeChartType(chart.type);

      const instance = new Chart(ctx, {
        type: chartType,
        data: {
          labels: chart.labels || [],
          datasets: (chart.datasets || []).map((ds) => ({
            label: ds.label || chart.title || '',
            data: ds.data || [],
            tension: 0.25
          }))
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: {
              display: chartType === 'pie' || chartType === 'doughnut'
            },
            tooltip: {
              callbacks: {
                label(context) {
                  const label = context.dataset.label || '';
                  const value = context.parsed.y ?? context.parsed ?? 0;
                  return `${label}: ${formatNumber(value)}`;
                }
              }
            }
          },
          scales: chartType === 'pie' || chartType === 'doughnut'
            ? {}
            : {
                y: {
                  beginAtZero: true
                }
              }
        }
      });

      chartInstances.push(instance);
    });
  }

  function renderTable(table) {
    const thead = document.querySelector('#dataTable thead');
    const tbody = document.querySelector('#dataTable tbody');
    const meta = document.getElementById('tableMeta');

    if (!thead || !tbody) return;

    const columns = table.columns || [];
    const rows = table.rows || [];

    if (meta) {
      meta.textContent = `แสดง ${formatNumber(rows.length)} จาก ${formatNumber(table.totalRows || rows.length)} แถว`;
    }

    if (!columns.length) {
      thead.innerHTML = '';
      tbody.innerHTML = `<tr><td class="emptyState">ไม่พบคอลัมน์ตาราง</td></tr>`;
      return;
    }

    thead.innerHTML = `
      <tr>
        ${columns.map((col) => `<th>${escapeHtml(col.title || col.field)}</th>`).join('')}
      </tr>
    `;

    tbody.innerHTML = rows.map((row) => {
      return `
        <tr>
          ${columns.map((col) => `<td>${escapeHtml(String(row[col.field] ?? ''))}</td>`).join('')}
        </tr>
      `;
    }).join('');

    applyTableSearch();
  }

  function applyTableSearch() {
    const input = document.getElementById('tableSearch');
    const tbody = document.querySelector('#dataTable tbody');

    if (!input || !tbody) return;

    const q = input.value.trim().toLowerCase();
    const trs = tbody.querySelectorAll('tr');

    trs.forEach((tr) => {
      const text = tr.textContent.toLowerCase();
      tr.style.display = !q || text.includes(q) ? '' : 'none';
    });
  }

  function setDefaultDateThisMonth() {
    const from = document.getElementById('dateFrom');
    const to = document.getElementById('dateTo');

    if (!from || !to) return;

    const now = new Date();
    const first = new Date(now.getFullYear(), now.getMonth(), 1);
    const last = new Date(now.getFullYear(), now.getMonth() + 1, 0);

    from.value = toDateInputValue(first);
    to.value = toDateInputValue(last);
  }

  function clearFilters() {
    const from = document.getElementById('dateFrom');
    const to = document.getElementById('dateTo');

    if (from) from.value = '';
    if (to) to.value = '';

    document.querySelectorAll('[data-filter-field]').forEach((el) => {
      el.value = el.tagName === 'SELECT' ? 'ALL' : '';
    });
  }

  function setLoading(isLoading) {
    const el = document.getElementById('loadingState');
    if (!el) return;
    el.classList.toggle('hidden', !isLoading);
  }

  function normalizeChartType(type) {
    if (type === 'horizontalBar') return 'bar';
    if (type === 'area') return 'line';
    if (type === 'ranking') return 'bar';
    if (type === 'pareto') return 'bar';
    if (['line', 'bar', 'pie', 'doughnut'].includes(type)) return type;
    return 'bar';
  }

  function toDateInputValue(date) {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  function formatNumber(value) {
    const n = Number(value);
    if (Number.isNaN(n)) return String(value ?? '');
    return n.toLocaleString('en-US', { maximumFractionDigits: 2 });
  }

  function escapeHtml(value) {
    return String(value ?? '')
      .replaceAll('&', '&amp;')
      .replaceAll('<', '&lt;')
      .replaceAll('>', '&gt;')
      .replaceAll('"', '&quot;')
      .replaceAll("'", '&#039;');
  }

  return {
    loadDashboards,
    loadDashboardData,
    setDefaultDateThisMonth,
    clearFilters,
    applyTableSearch
  };
})();
