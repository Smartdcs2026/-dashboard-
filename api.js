(function () {
  'use strict';

  const API_BASE = 'https://dashboard.somchaibutphon.workers.dev';
  const TOKEN_KEY = 'analytics_dashboard_token';

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
      response = await fetch(url, fetchOptions);
    } catch (error) {
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
    return request('/api/health');
  }

  function setupStatus() {
    return request('/api/setup-status');
  }

  async function login(username, password) {
    const data = await request('/api/login', {
      method: 'POST',
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
    return request('/api/me');
  }

  async function logout() {
    try {
      await request('/api/logout', {
        method: 'POST',
        body: {}
      });
    } finally {
      clearToken();
    }
  }

  function changePassword(oldPassword, newPassword, confirmPassword) {
    return request('/api/change-password', {
      method: 'POST',
      body: {
        oldPassword,
        newPassword,
        confirmPassword
      }
    });
  }

  function listSources() {
    return request('/api/sources');
  }

  function createSource(payload) {
    return request('/api/sources', {
      method: 'POST',
      body: payload
    });
  }

  function listSourceSheets(payload) {
    return request('/api/source-sheets', {
      method: 'POST',
      body: payload
    });
  }

  function readHeaders(payload) {
    return request('/api/headers', {
      method: 'POST',
      body: payload
    });
  }

  function listDashboards() {
    return request('/api/dashboards');
  }

  window.AnalyticsAPI = {
    API_BASE,
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
    listDashboards
  };
})();
