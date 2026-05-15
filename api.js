(function () {
  'use strict';

  const API_BASE = 'https://dataanalyticsdashboard.somchaibutphon.workers.dev';

  function getToken() {
    return localStorage.getItem('analytics_token') || '';
  }

  function setToken(token) {
    if (token) {
      localStorage.setItem('analytics_token', token);
    }
  }

  function clearToken() {
    localStorage.removeItem('analytics_token');
  }

  async function request(path, options = {}) {
    const token = getToken();

    const headers = {
      'Content-Type': 'application/json',
      ...(options.headers || {})
    };

    if (token) {
      headers.Authorization = `Bearer ${token}`;
    }

    const res = await fetch(`${API_BASE}${path}`, {
      ...options,
      headers
    });

    let data;
    try {
      data = await res.json();
    } catch (err) {
      throw new Error('API ตอบกลับไม่ใช่ JSON');
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
        body: JSON.stringify({ username, password })
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

    dashboards() {
      return request('/api/dashboards');
    },

    dashboardConfig(dashboardId) {
      return request(`/api/dashboard-config/${encodeURIComponent(dashboardId)}`);
    },

    dashboardData(dashboardId, params = {}) {
      return request(`/api/dashboard-data/${encodeURIComponent(dashboardId)}${buildQuery(params)}`);
    }
  };
})();
