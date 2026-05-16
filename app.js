(function () {
  'use strict';

  const el = {
    loginView: document.getElementById('loginView'),
    dashboardView: document.getElementById('dashboardView'),

    apiStatus: document.getElementById('apiStatus'),
    loginForm: document.getElementById('loginForm'),
    username: document.getElementById('username'),
    password: document.getElementById('password'),
    loginBtn: document.getElementById('loginBtn'),
    loginMessage: document.getElementById('loginMessage'),

    logoutBtn: document.getElementById('logoutBtn'),
    refreshMeBtn: document.getElementById('refreshMeBtn'),
    loadSourcesBtn: document.getElementById('loadSourcesBtn'),
    loadDashboardsBtn: document.getElementById('loadDashboardsBtn'),
    clearLogBtn: document.getElementById('clearLogBtn'),

    userDisplayName: document.getElementById('userDisplayName'),
    userRole: document.getElementById('userRole'),
    systemStatusText: document.getElementById('systemStatusText'),

    healthResult: document.getElementById('healthResult'),
    meResult: document.getElementById('meResult'),
    sourceResult: document.getElementById('sourceResult'),
    dashboardResult: document.getElementById('dashboardResult'),
    debugLog: document.getElementById('debugLog')
  };

  let currentUser = null;

  document.addEventListener('DOMContentLoaded', init);

  async function init() {
    bindEvents();
    showLogin();
    await checkApiHealth();
    await restoreSession();
  }

  function bindEvents() {
    el.loginForm.addEventListener('submit', handleLogin);
    el.logoutBtn.addEventListener('click', handleLogout);
    el.refreshMeBtn.addEventListener('click', loadMe);
    el.loadSourcesBtn.addEventListener('click', loadSources);
    el.loadDashboardsBtn.addEventListener('click', loadDashboards);
    el.clearLogBtn.addEventListener('click', clearLog);
  }

  async function checkApiHealth() {
    setApiStatus('กำลังตรวจสอบสถานะระบบ...', 'muted');

    try {
      const health = await window.AnalyticsAPI.health();
      const setup = await window.AnalyticsAPI.setupStatus();

      setApiStatus('เชื่อมต่อ API สำเร็จ: ' + (health.message || 'พร้อมใช้งาน'), 'success');

      el.healthResult.textContent =
        'API พร้อมใช้งาน | Setup: ' + (setup.ok ? 'ครบถ้วน' : 'ยังไม่ครบ');

      setSystemStatus('API พร้อมใช้งาน');

      writeLog({
        step: 'health',
        health,
        setup
      });

    } catch (error) {
      setApiStatus(error.message, 'error');
      el.healthResult.textContent = 'เชื่อมต่อ API ไม่สำเร็จ';
      setSystemStatus('API มีปัญหา');
      writeLog({
        step: 'health_error',
        message: error.message,
        payload: error.payload || null
      });
    }
  }

  async function restoreSession() {
    const token = window.AnalyticsAPI.getToken();

    if (!token) {
      showLogin();
      return;
    }

    try {
      await loadMe();
      showDashboard();
    } catch (error) {
      window.AnalyticsAPI.clearToken();
      showLogin();
      writeLog({
        step: 'restore_session_error',
        message: error.message,
        payload: error.payload || null
      });
    }
  }

  async function handleLogin(event) {
    event.preventDefault();

    const username = el.username.value.trim();
    const password = el.password.value;

    setLoginMessage('');
    setLoading(true);

    try {
      const data = await window.AnalyticsAPI.login(username, password);

      currentUser = data.user || null;

      renderUser(currentUser);
      showDashboard();

      writeLog({
        step: 'login_success',
        response: data
      });

      await loadMe();

    } catch (error) {
      setLoginMessage(error.message || 'เข้าสู่ระบบไม่สำเร็จ');

      writeLog({
        step: 'login_error',
        message: error.message,
        payload: error.payload || null
      });

    } finally {
      setLoading(false);
    }
  }

  async function handleLogout() {
    try {
      await window.AnalyticsAPI.logout();
    } catch (error) {
      writeLog({
        step: 'logout_error',
        message: error.message,
        payload: error.payload || null
      });
    }

    currentUser = null;
    el.username.value = '';
    el.password.value = '';
    showLogin();
  }

  async function loadMe() {
    const data = await window.AnalyticsAPI.me();

    currentUser = data.user || null;
    renderUser(currentUser);

    el.meResult.textContent = currentUser
      ? `${currentUser.displayName || currentUser.username} (${currentUser.role})`
      : '-';

    writeLog({
      step: 'me',
      response: data
    });

    return data;
  }

  async function loadSources() {
    el.sourceResult.textContent = 'กำลังโหลด...';

    try {
      const data = await window.AnalyticsAPI.listSources();

      el.sourceResult.textContent = `พบแหล่งข้อมูล ${data.total || 0} รายการ`;

      writeLog({
        step: 'sources',
        response: data
      });

    } catch (error) {
      el.sourceResult.textContent = error.message;

      writeLog({
        step: 'sources_error',
        message: error.message,
        payload: error.payload || null
      });
    }
  }

  async function loadDashboards() {
    el.dashboardResult.textContent = 'กำลังโหลด...';

    try {
      const data = await window.AnalyticsAPI.listDashboards();

      el.dashboardResult.textContent = `พบ Dashboard ${data.total || 0} รายการ`;

      writeLog({
        step: 'dashboards',
        response: data
      });

    } catch (error) {
      el.dashboardResult.textContent = error.message;

      writeLog({
        step: 'dashboards_error',
        message: error.message,
        payload: error.payload || null
      });
    }
  }

  function renderUser(user) {
    if (!user) {
      el.userDisplayName.textContent = '-';
      el.userRole.textContent = '-';
      return;
    }

    el.userDisplayName.textContent = user.displayName || user.username || '-';
    el.userRole.textContent = user.role || '-';

    if (user.mustChangePassword) {
      setSystemStatus('เข้าสู่ระบบสำเร็จ - แนะนำให้เปลี่ยนรหัสผ่านเริ่มต้น');
    } else {
      setSystemStatus('เข้าสู่ระบบสำเร็จ');
    }
  }

  function showLogin() {
    el.loginView.classList.remove('hidden');
    el.dashboardView.classList.add('hidden');
  }

  function showDashboard() {
    el.loginView.classList.add('hidden');
    el.dashboardView.classList.remove('hidden');
  }

  function setApiStatus(message, type) {
    el.apiStatus.textContent = message || '';

    el.apiStatus.classList.remove('status-muted', 'status-success', 'status-error');

    if (type === 'success') {
      el.apiStatus.classList.add('status-success');
    } else if (type === 'error') {
      el.apiStatus.classList.add('status-error');
    } else {
      el.apiStatus.classList.add('status-muted');
    }
  }

  function setSystemStatus(message) {
    el.systemStatusText.textContent = message || '';
  }

  function setLoginMessage(message) {
    el.loginMessage.textContent = message || '';
  }

  function setLoading(isLoading) {
    el.loginBtn.disabled = !!isLoading;
    el.loginBtn.textContent = isLoading ? 'กำลังเข้าสู่ระบบ...' : 'เข้าสู่ระบบ';
  }

  function writeLog(data) {
    const text = JSON.stringify(data, null, 2);
    el.debugLog.textContent = text;
  }

  function clearLog() {
    el.debugLog.textContent = 'ยังไม่มีข้อมูล';
  }
})();
