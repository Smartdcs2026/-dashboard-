(function () {
  'use strict';

  const API_BASE = 'https://dashboard.somchaibutphon.workers.dev';
  const TOKEN_KEY = 'analytics_dashboard_token';

  const API_TIMEOUT = {
    DEFAULT: 30000,
    HEALTH: 15000,
    LOGIN: 20000,
    READ: 45000,
    DASHBOARD: 90000,
    EXPORT: 120000
  };

  function getToken() {
    return localStorage.getItem(TOKEN_KEY) || '';
  }

  function setToken(token) {
    if (token) {
      localStorage.setItem(TOKEN_KEY, token);
    }
  }

  function clearToken() {
    localStorage.removeItem(TOKEN_KEY);
  }

  function buildQuery(params) {
    const search = new URLSearchParams();

    Object.keys(params || {}).forEach(function (key) {
      const value = params[key];

      if (value !== undefined && value !== null && String(value).trim() !== '') {
        search.set(key, String(value));
      }
    });

    const text = search.toString();
    return text ? '?' + text : '';
  }

  function fetchWithTimeout(url, fetchOptions, timeoutMs) {
    const controller = new AbortController();

    const timer = setTimeout(function () {
      controller.abort();
    }, timeoutMs || API_TIMEOUT.DEFAULT);

    return fetch(url, {
      ...fetchOptions,
      signal: controller.signal
    }).finally(function () {
      clearTimeout(timer);
    });
  }

  async function request(path, options = {}) {
    const token = getToken();
    const method = String(options.method || 'GET').toUpperCase();
    const query = options.query ? buildQuery(options.query) : '';
    const url = API_BASE + path + query;

    const headers = {
      Accept: 'application/json',
      ...(options.headers || {})
    };

    let body = options.body;

    if (body && typeof body === 'object' && !(body instanceof FormData)) {
      headers['Content-Type'] = 'application/json';
      body = JSON.stringify(body);
    }

    if (token) {
      headers.Authorization = 'Bearer ' + token;
    }

    const fetchOptions = {
      method,
      headers
    };

    if (method !== 'GET' && method !== 'HEAD' && body !== undefined) {
      fetchOptions.body = body;
    }

    let response;
    let data;

    try {
      response = await fetchWithTimeout(
        url,
        fetchOptions,
        options.timeoutMs || API_TIMEOUT.DEFAULT
      );
    } catch (error) {
      if (error && error.name === 'AbortError') {
        throw new Error('API ใช้เวลานานเกินกำหนด กรุณาลองใหม่อีกครั้ง หรือลดจำนวนแถวที่โหลด');
      }

      throw new Error('ไม่สามารถเชื่อมต่อ API ได้: ' + error.message);
    }

    try {
      data = await response.json();
    } catch (error) {
      throw new Error('API ตอบกลับไม่ใช่ JSON');
    }

    if (!response.ok || data.ok === false) {
      const message = data.message || 'API Error';
      const err = new Error(message);
      err.payload = data;
      err.status = response.status;
      throw err;
    }

    return data;
  }

  function health() {
    return request('/api/health', {
      timeoutMs: API_TIMEOUT.HEALTH
    });
  }

  function setupStatus() {
    return request('/api/setup-status', {
      timeoutMs: API_TIMEOUT.HEALTH
    });
  }

  async function login(username, password) {
    const data = await request('/api/login', {
      method: 'POST',
      timeoutMs: API_TIMEOUT.LOGIN,
      body: {
        username,
        password
      }
    });

    if (data.token) {
      setToken(data.token);
    }

    return data;
  }

  function me() {
    return request('/api/me', {
      timeoutMs: API_TIMEOUT.READ
    });
  }

  async function logout() {
    try {
      await request('/api/logout', {
        method: 'POST',
        timeoutMs: API_TIMEOUT.READ,
        body: {}
      });
    } finally {
      clearToken();
    }
  }

  function changePassword(oldPassword, newPassword, confirmPassword) {
    return request('/api/change-password', {
      method: 'POST',
      timeoutMs: API_TIMEOUT.READ,
      body: {
        oldPassword,
        newPassword,
        confirmPassword
      }
    });
  }

  function listSources() {
    return request('/api/sources', {
      timeoutMs: API_TIMEOUT.READ
    });
  }

  function createSource(payload) {
    return request('/api/sources', {
      method: 'POST',
      timeoutMs: API_TIMEOUT.READ,
      body: payload
    });
  }

  function listSourceSheets(payload) {
    return request('/api/source-sheets', {
      method: 'POST',
      timeoutMs: API_TIMEOUT.READ,
      body: payload
    });
  }

  function readHeaders(payload) {
    return request('/api/headers', {
      method: 'POST',
      timeoutMs: API_TIMEOUT.READ,
      body: payload
    });
  }

  function getMapping(payload) {
    return request('/api/mapping', {
      method: 'POST',
      timeoutMs: API_TIMEOUT.READ,
      body: {
        mode: 'get',
        ...payload
      }
    });
  }

  function saveMapping(payload) {
    return request('/api/mapping', {
      method: 'POST',
      timeoutMs: API_TIMEOUT.READ,
      body: {
        mode: 'save',
        ...payload
      }
    });
  }

  function dashboardPreview(payload) {
    return request('/api/dashboard-preview', {
      method: 'POST',
      timeoutMs: API_TIMEOUT.DASHBOARD,
      body: payload
    });
  }

  function createDashboardFromPreview(payload) {
    return request('/api/dashboard-create', {
      method: 'POST',
      timeoutMs: API_TIMEOUT.DASHBOARD,
      body: payload
    });
  }

  function dashboardView(payload) {
    return request('/api/dashboard-view', {
      method: 'POST',
      timeoutMs: API_TIMEOUT.DASHBOARD,
      body: payload
    });
  }

  function dashboardExport(payload) {
    return request('/api/dashboard-export', {
      method: 'POST',
      timeoutMs: API_TIMEOUT.EXPORT,
      body: payload
    });
  }

  function listDashboards(options = {}) {
    const includeDeleted = !!options.includeDeleted;

    if (includeDeleted) {
      return request('/api/dashboards', {
        timeoutMs: API_TIMEOUT.READ,
        query: {
          includeDeleted: true
        }
      });
    }

    return request('/api/dashboards', {
      timeoutMs: API_TIMEOUT.READ
    });
  }

  function manageDashboard(payload) {
    return request('/api/dashboards', {
      method: 'POST',
      timeoutMs: API_TIMEOUT.READ,
      body: payload
    });
  }

  function updateDashboard(payload) {
    payload = payload || {};

    if (!payload.dashboardId) {
      return Promise.reject(new Error('ไม่พบ dashboardId สำหรับแก้ไข Dashboard'));
    }

    return manageDashboard({
      mode: 'update',
      ...payload
    });
  }

  function updateDashboardBasic(dashboardId, dashboardName, dashboardType, description) {
    return updateDashboard({
      dashboardId,
      dashboardName,
      dashboardType,
      description
    });
  }

  function setDashboardPublish(dashboardId, publish) {
    return updateDashboard({
      dashboardId,
      publish: !!publish
    });
  }

  function setDashboardExport(dashboardId, allowExport) {
    return updateDashboard({
      dashboardId,
      allowExport: !!allowExport
    });
  }

  function setDashboardHidden(dashboardId, hide) {
    return updateDashboard({
      dashboardId,
      hide: !!hide
    });
  }

  function setDashboardVisibility(dashboardId, visibility) {
    return updateDashboard({
      dashboardId,
      visibility: visibility
    });
  }

  function regenerateDashboard(dashboardId) {
    if (!dashboardId) {
      return Promise.reject(new Error('ไม่พบ dashboardId สำหรับ Regenerate Dashboard'));
    }

    return manageDashboard({
      mode: 'regenerate',
      dashboardId
    });
  }

  function deleteDashboard(dashboardId) {
    if (!dashboardId) {
      return Promise.reject(new Error('ไม่พบ dashboardId สำหรับลบ Dashboard'));
    }

    return manageDashboard({
      mode: 'delete',
      dashboardId
    });
  }

  function auditLog(payload = {}) {
    return request('/api/audit-log', {
      method: 'POST',
      timeoutMs: API_TIMEOUT.READ,
      body: payload
    });
  }

  window.AnalyticsAPI = {
    API_BASE,
    TOKEN_KEY,
    API_TIMEOUT,

    getToken,
    setToken,
    clearToken,

    request,

    health,
    setupStatus,

    login,
    me,
    logout,
    changePassword,

    listSources,
    createSource,
    listSourceSheets,
    readHeaders,

    getMapping,
    saveMapping,

    dashboardPreview,
    createDashboardFromPreview,
    dashboardView,
    dashboardExport,

    listDashboards,

    manageDashboard,
    updateDashboard,
    updateDashboardBasic,
    setDashboardPublish,
    setDashboardExport,
    setDashboardHidden,
    setDashboardVisibility,
    regenerateDashboard,
    deleteDashboard,

    auditLog
  };
})();
