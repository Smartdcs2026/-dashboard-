(function () {
  'use strict';

  /*
   * สำคัญ:
   * API_BASE ต้องเป็น URL Worker หลักเท่านั้น
   * ห้ามต่อท้าย /api
   */
  const API_BASE = 'https://dataanalyticsdashboard.somchaibutphon.workers.dev';

  const TOKEN_KEY = 'analytics_token';

  function getBaseUrl() {
    return String(API_BASE || '').replace(/\/+$/, '');
  }

  function normalizePath(path) {
    const p = String(path || '').trim();

    if (!p) return '/';

    // บังคับให้ path เริ่มด้วย /
    return p.startsWith('/') ? p : `/${p}`;
  }

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

  async function request(path, options = {}) {
    const token = getToken();
    const url = `${getBaseUrl()}${normalizePath(path)}`;

    const headers = {
      Accept: 'application/json',
      'Content-Type': 'application/json',
      ...(options.headers || {})
    };

    if (token) {
      headers.Authorization = `Bearer ${token}`;
    }

    let res;

    try {
      res = await fetch(url, {
        ...options,
        headers
      });
    } catch (err) {
      throw new Error(`เชื่อมต่อ API ไม่สำเร็จ: ${err.message || err}`);
    }

    let data;

    try {
      data = await res.json();
    } catch (err) {
      throw new Error(`API ตอบกลับไม่ใช่ JSON หรือ Worker URL ไม่ถูกต้อง: ${url}`);
    }

    if (!res.ok || data.ok === false) {
      throw new Error(data.message || `API Error: ${res.status}`);
    }

    return data.data;
  }

  function buildQuery(params = {}) {
    const q = new URLSearchParams();

    Object.entries(params).forEach(([key, value]) => {
      if (value === undefined || value === null || value === '') return;
      q.set(key, value);
    });

    const s = q.toString();
    return s ? `?${s}` : '';
  }

  window.Api = {
    getToken,
    setToken,
    clearToken,

    health() {
      return request('/api/health');
    },

    login(username, password) {
      return request('/api/login', {
        method: 'POST',
        body: JSON.stringify({
          username,
          password
        })
      });
    },

    logout() {
      return request('/api/logout', {
        method: 'POST',
        body: JSON.stringify({})
      });
    },

    authCheck() {
      return request('/api/auth-check');
    },

    me() {
      return request('/api/me');
    },

    config() {
      return request('/api/config');
    },

    dashboards() {
      return request('/api/dashboards');
    },

    dashboardConfig(dashboardId) {
      return request(`/api/dashboard-config/${encodeURIComponent(dashboardId)}`);
    },

    dashboardData(dashboardId, params = {}) {
      return request(
        `/api/dashboard-data/${encodeURIComponent(dashboardId)}${buildQuery(params)}`
      );
    },

    sourceHeaders(sourceId) {
      return request(`/api/source-headers/${encodeURIComponent(sourceId)}`);
    },

    sourcePreview(sourceId, limit = 20) {
      return request(
        `/api/source-preview/${encodeURIComponent(sourceId)}${buildQuery({ limit })}`
      );
    },

    fieldMap(sourceId) {
      return request(`/api/field-map/${encodeURIComponent(sourceId)}`);
    }
  };
})();
