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
function normalizeApiErrorMessage(message, status) {
  const text = String(message || '').trim();

  if (!text) {
    return 'เกิดข้อผิดพลาดจาก API';
  }

  if (
    text.includes('Token หมดอายุ') ||
    text.includes('กรุณาเข้าสู่ระบบก่อนใช้งาน') ||
    text.includes('Token ไม่ถูกต้อง')
  ) {
    return 'เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่';
  }

  if (
    text.includes('ไม่มีสิทธิ์') ||
    status === 401 ||
    status === 403
  ) {
    return 'คุณไม่มีสิทธิ์ใช้งานส่วนนี้';
  }

  if (
    text.includes('ยังไม่มี Mapping') ||
    text.includes('ไม่มี Mapping')
  ) {
    return 'ยังไม่มี Mapping ของชีทนี้ กรุณาบันทึก Mapping ก่อน';
  }

  if (
    text.includes('ไม่พบ Dashboard') ||
    text.includes('Dashboard นี้')
  ) {
    return text;
  }

  if (
    text.includes('ไม่พบชีท') ||
    text.includes('ไม่สามารถเปิด Google Sheet') ||
    text.includes('ไม่สามารถเปิดแหล่งข้อมูล')
  ) {
    return 'ไม่สามารถอ่านข้อมูลจาก Google Sheet ได้ กรุณาตรวจสอบชื่อชีท / สิทธิ์การเข้าถึง / Sheet ID';
  }

  if (
    text.includes('Apps Script ตอบกลับไม่ใช่ JSON') ||
    text.includes('Worker Error') ||
    text.includes('เชื่อมต่อ Apps Script ไม่สำเร็จ')
  ) {
    return 'ระบบ API Gateway หรือ Apps Script มีปัญหา กรุณาลองใหม่อีกครั้ง';
  }

  if (
    text.includes('API ใช้เวลานานเกินกำหนด') ||
    text.includes('timeout') ||
    text.includes('aborted')
  ) {
    return 'API ใช้เวลานานเกินกำหนด กรุณาลดจำนวนแถวที่โหลด หรือกดใหม่อีกครั้ง';
  }

  return text;
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
  const err = new Error('API ใช้เวลานานเกินกำหนด กรุณาลดจำนวนแถวที่โหลด หรือกดใหม่อีกครั้ง');
  err.status = 408;
  err.isTimeout = true;
  throw err;
}

      throw new Error('ไม่สามารถเชื่อมต่อ API ได้: ' + error.message);
    }

    try {
      data = await response.json();
    } catch (error) {
      throw new Error('API ตอบกลับไม่ใช่ JSON');
    }

    if (!response.ok || data.ok === false) {
  const rawMessage = data.message || 'API Error';
  const message = normalizeApiErrorMessage(rawMessage, response.status);

  const err = new Error(message);
  err.rawMessage = rawMessage;
  err.payload = data;
  err.status = response.status;
  err.isAuthError =
    message.includes('เซสชันหมดอายุ') ||
    rawMessage.includes('Token หมดอายุ') ||
    rawMessage.includes('กรุณาเข้าสู่ระบบก่อนใช้งาน') ||
    rawMessage.includes('Token ไม่ถูกต้อง');

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
function clearCache(payload = {}) {
  return request('/api/clear-cache', {
    method: 'POST',
    timeoutMs: API_TIMEOUT.READ,
    body: payload
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
  normalizeApiErrorMessage,

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
    clearCache,

    auditLog
  };
})();
