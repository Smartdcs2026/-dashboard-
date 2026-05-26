(function () {
  'use strict';

  const FORCE_LOGIN_EVERY_OPEN = true;

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

    sourceForm: document.getElementById('sourceForm'),
    sourceName: document.getElementById('sourceName'),
    sourceSpreadsheetId: document.getElementById('sourceSpreadsheetId'),
    sourceDataType: document.getElementById('sourceDataType'),
    sourceOwner: document.getElementById('sourceOwner'),
    sourceDescription: document.getElementById('sourceDescription'),
    createSourceBtn: document.getElementById('createSourceBtn'),
    reloadSourcesBtn: document.getElementById('reloadSourcesBtn'),
    sourceFormMessage: document.getElementById('sourceFormMessage'),

    sourcesList: document.getElementById('sourcesList'),
    sourceSheetsList: document.getElementById('sourceSheetsList'),
    headersResult: document.getElementById('headersResult'),

    saveMappingBtn: document.getElementById('saveMappingBtn'),
    mappingMessage: document.getElementById('mappingMessage'),

    previewDashboardBtn: document.getElementById('previewDashboardBtn'),
    previewResult: document.getElementById('previewResult'),

    dashboardName: document.getElementById('dashboardName'),
    dashboardType: document.getElementById('dashboardType'),
    createDashboardBtn: document.getElementById('createDashboardBtn'),
    createDashboardMessage: document.getElementById('createDashboardMessage'),

    reloadDashboardListBtn: document.getElementById('reloadDashboardListBtn'),
    dashboardSelect: document.getElementById('dashboardSelect'),
    dashboardLimit: document.getElementById('dashboardLimit'),
    dashboardFilterBox: document.getElementById('dashboardFilterBox'),
    openDashboardBtn: document.getElementById('openDashboardBtn'),
    applyDashboardFilterBtn: document.getElementById('applyDashboardFilterBtn'),
  resetDashboardFilterBtn: document.getElementById('resetDashboardFilterBtn'),
    refreshDashboardBtn: document.getElementById('refreshDashboardBtn'),
exportDashboardBtn: document.getElementById('exportDashboardBtn'),
    exportDashboardExcelBtn: document.getElementById('exportDashboardExcelBtn'),
dashboardViewMessage: document.getElementById('dashboardViewMessage'),
    dashboardViewResult: document.getElementById('dashboardViewResult'),
        reloadManageDashboardsBtn: document.getElementById('reloadManageDashboardsBtn'),
    manageDashboardSelect: document.getElementById('manageDashboardSelect'),
    manageDashboardSearch: document.getElementById('manageDashboardSearch'),
manageDashboardTypeFilter: document.getElementById('manageDashboardTypeFilter'),
manageDashboardStatusFilter: document.getElementById('manageDashboardStatusFilter'),
manageDashboardCount: document.getElementById('manageDashboardCount'),
    manageDashboardName: document.getElementById('manageDashboardName'),
    manageDashboardType: document.getElementById('manageDashboardType'),
    manageDashboardDescription: document.getElementById('manageDashboardDescription'),
    manageDashboardPublish: document.getElementById('manageDashboardPublish'),
    manageDashboardExport: document.getElementById('manageDashboardExport'),
    manageDashboardHidden: document.getElementById('manageDashboardHidden'),
    saveManageDashboardBtn: document.getElementById('saveManageDashboardBtn'),
    regenerateManageDashboardBtn: document.getElementById('regenerateManageDashboardBtn'),
    hideManageDashboardBtn: document.getElementById('hideManageDashboardBtn'),
    deleteManageDashboardBtn: document.getElementById('deleteManageDashboardBtn'),
    manageDashboardMessage: document.getElementById('manageDashboardMessage'),
    manageDashboardSummary: document.getElementById('manageDashboardSummary'),
    manageDashboardCardList: document.getElementById('manageDashboardCardList'),
manageDashboardCardCount: document.getElementById('manageDashboardCardCount'),

    reloadUsersBtn: document.getElementById('reloadUsersBtn'),
userManageUsername: document.getElementById('userManageUsername'),
userManageDisplayName: document.getElementById('userManageDisplayName'),
userManageEmail: document.getElementById('userManageEmail'),
userManageRole: document.getElementById('userManageRole'),
userManageStatus: document.getElementById('userManageStatus'),
userManagePassword: document.getElementById('userManagePassword'),
userManageDashboards: document.getElementById('userManageDashboards'),
userCanExport: document.getElementById('userCanExport'),
userCanAddSource: document.getElementById('userCanAddSource'),
userCanEditDashboard: document.getElementById('userCanEditDashboard'),
userCanViewAuditLog: document.getElementById('userCanViewAuditLog'),
userCanManageUser: document.getElementById('userCanManageUser'),
createUserBtn: document.getElementById('createUserBtn'),
updateUserBtn: document.getElementById('updateUserBtn'),
resetUserPasswordBtn: document.getElementById('resetUserPasswordBtn'),
clearUserFormBtn: document.getElementById('clearUserFormBtn'),
userManageMessage: document.getElementById('userManageMessage'),
userManageCount: document.getElementById('userManageCount'),
userManageList: document.getElementById('userManageList'),

    userDisplayName: document.getElementById('userDisplayName'),
    userRole: document.getElementById('userRole'),
    systemStatusText: document.getElementById('systemStatusText'),
    modeSwitcher: document.getElementById('modeSwitcher'),
adminModeBtn: document.getElementById('adminModeBtn'),
userModeBtn: document.getElementById('userModeBtn'),

    healthResult: document.getElementById('healthResult'),
    meResult: document.getElementById('meResult'),
    sourceResult: document.getElementById('sourceResult'),
    dashboardResult: document.getElementById('dashboardResult'),
    globalLoadingOverlay: document.getElementById('globalLoadingOverlay'),
globalLoadingText: document.getElementById('globalLoadingText'),
    reloadAuditLogBtn: document.getElementById('reloadAuditLogBtn'),
auditLimitSelect: document.getElementById('auditLimitSelect'),
auditLogResult: document.getElementById('auditLogResult'),
    
    debugLog: document.getElementById('debugLog')
  };

  let currentUser = null;
  let sourcesCache = [];
  let selectedSourceId = '';
  let selectedSheetName = '';
  let lastHeadersData = null;
  let currentDashboardId = '';
  let currentDashboardFilters = [];
    let manageDashboardsCache = [];
  let selectedManageDashboard = null;
  let chartRenderQueue = [];
let chartInstances = [];
let chartUid = 0;
  let currentMode = 'admin';
  let usersCache = [];
let selectedManageUserId = '';

  let currentDashboardTable = null;
let currentDashboardTablePage = 1;
const DASHBOARD_TABLE_PAGE_SIZE = 50;

  document.addEventListener('DOMContentLoaded', init);

  async function init() {
    bindEvents();
    showLogin();
    resetWorkingState();

    if (FORCE_LOGIN_EVERY_OPEN && window.AnalyticsAPI && window.AnalyticsAPI.clearToken) {
      window.AnalyticsAPI.clearToken();
    }

    await checkApiHealth();

    if (!FORCE_LOGIN_EVERY_OPEN) {
      await restoreSession();
    } else {
      showLogin();
    }
  }

  function bindEvents() {
    if (el.loginForm) {
      el.loginForm.addEventListener('submit', handleLogin);
    }

    if (el.logoutBtn) {
      el.logoutBtn.addEventListener('click', handleLogout);
    }

    if (el.refreshMeBtn) {
      el.refreshMeBtn.addEventListener('click', loadMe);
    }

    if (el.loadSourcesBtn) {
      el.loadSourcesBtn.addEventListener('click', loadSources);
    }

    if (el.reloadSourcesBtn) {
      el.reloadSourcesBtn.addEventListener('click', loadSources);
    }

    if (el.loadDashboardsBtn) {
      el.loadDashboardsBtn.addEventListener('click', loadDashboards);
    }

    if (el.clearLogBtn) {
      el.clearLogBtn.addEventListener('click', clearLog);
    }

    if (el.sourceForm) {
      el.sourceForm.addEventListener('submit', handleCreateSource);
    }

    if (el.saveMappingBtn) {
      el.saveMappingBtn.addEventListener('click', handleSaveMapping);
      el.saveMappingBtn.disabled = true;
    }

    if (el.previewDashboardBtn) {
      el.previewDashboardBtn.addEventListener('click', handleDashboardPreview);
    }

    if (el.createDashboardBtn) {
      el.createDashboardBtn.addEventListener('click', handleCreateDashboardFromPreview);
    }

    if (el.reloadDashboardListBtn) {
      el.reloadDashboardListBtn.addEventListener('click', loadDashboardOptions);
    }

    if (el.openDashboardBtn) {
      el.openDashboardBtn.addEventListener('click', handleOpenDashboard);
    }

    if (el.applyDashboardFilterBtn) {
      el.applyDashboardFilterBtn.addEventListener('click', handleApplyDashboardFilter);
    }
      if (el.adminModeBtn) {
  el.adminModeBtn.addEventListener('click', function () {
    setAppMode('admin');
  });
}

if (el.userModeBtn) {
  el.userModeBtn.addEventListener('click', function () {
    setAppMode('user');
  });
}
if (el.reloadUsersBtn) {
  el.reloadUsersBtn.addEventListener('click', loadUsers);
}

if (el.createUserBtn) {
  el.createUserBtn.addEventListener('click', handleCreateUser);
}

if (el.updateUserBtn) {
  el.updateUserBtn.addEventListener('click', handleUpdateUser);
}

if (el.resetUserPasswordBtn) {
  el.resetUserPasswordBtn.addEventListener('click', handleResetUserPassword);
}

if (el.clearUserFormBtn) {
  el.clearUserFormBtn.addEventListener('click', resetUserForm);
}
    if (el.resetDashboardFilterBtn) {
      el.resetDashboardFilterBtn.addEventListener('click', handleResetDashboardFilter);
    }
     if (el.exportDashboardBtn) {
  el.exportDashboardBtn.addEventListener('click', handleExportDashboard);
}
    if (el.exportDashboardExcelBtn) {
  el.exportDashboardExcelBtn.addEventListener('click', handleExportDashboardExcel);
}
    if (el.refreshDashboardBtn) {
  el.refreshDashboardBtn.addEventListener('click', handleRefreshDashboard);
}
        if (el.reloadManageDashboardsBtn) {
      el.reloadManageDashboardsBtn.addEventListener('click', loadManageDashboards);
    }

    if (el.manageDashboardSelect) {
      el.manageDashboardSelect.addEventListener('change', handleSelectManageDashboard);
    }
   if (el.manageDashboardSearch) {
  el.manageDashboardSearch.addEventListener('input', renderFilteredManageDashboardOptions);
}

if (el.manageDashboardTypeFilter) {
  el.manageDashboardTypeFilter.addEventListener('change', renderFilteredManageDashboardOptions);
}

if (el.manageDashboardStatusFilter) {
  el.manageDashboardStatusFilter.addEventListener('change', renderFilteredManageDashboardOptions);
}
    if (el.saveManageDashboardBtn) {
      el.saveManageDashboardBtn.addEventListener('click', handleSaveManageDashboard);
    }
   if (el.regenerateManageDashboardBtn) {
  el.regenerateManageDashboardBtn.addEventListener('click', handleRegenerateManageDashboard);
}
    if (el.hideManageDashboardBtn) {
      el.hideManageDashboardBtn.addEventListener('click', handleToggleManageDashboardHidden);
    }
    if (el.reloadAuditLogBtn) {
  el.reloadAuditLogBtn.addEventListener('click', loadAuditLog);
}

    if (el.deleteManageDashboardBtn) {
      el.deleteManageDashboardBtn.addEventListener('click', handleDeleteManageDashboard);
    }
  }

  function resetWorkingState() {
    currentUser = null;
    sourcesCache = [];
    selectedSourceId = '';
    selectedSheetName = '';
    lastHeadersData = null;
    currentDashboardId = '';
    currentDashboardFilters = [];
        manageDashboardsCache = [];
    selectedManageDashboard = null;

    renderSources([]);
    renderSourceSheets([]);
    renderHeaders(null);

    if (el.previewResult) {
      el.previewResult.classList.add('empty');
      el.previewResult.textContent = 'กรุณาบันทึก Mapping ก่อน แล้วกดสร้าง Dashboard Preview';
    }

    if (el.dashboardViewResult) {
      el.dashboardViewResult.classList.add('empty');
      el.dashboardViewResult.textContent = 'กรุณาเลือก Dashboard แล้วกดเปิด Dashboard';
    }

    if (el.dashboardFilterBox) {
      el.dashboardFilterBox.classList.add('empty');
      el.dashboardFilterBox.textContent = 'เมื่อเปิด Dashboard แล้ว ระบบจะแสดงตัวกรองที่ตั้งไว้ตรงนี้';
    }

    if (el.dashboardSelect) {
      el.dashboardSelect.innerHTML = '<option value="">ยังไม่มีรายการ Dashboard</option>';
    }

    usersCache = [];
selectedManageUserId = '';
resetUserForm();

if (el.userManageList) {
  el.userManageList.classList.add('empty');
  el.userManageList.textContent = 'กดโหลดรายชื่อผู้ใช้เพื่อแสดงรายการ';
}

if (el.userManageCount) {
  el.userManageCount.textContent = '0 รายการ';
}
        resetManageDashboardForm();
    setLoginMessage('');
    setSourceFormMessage('');
    setMappingMessage('');
    setPreviewMessage('');
    setCreateDashboardMessage('');
    setDashboardViewMessage('');
  }

  async function checkApiHealth() {
    setApiStatus('กำลังตรวจสอบสถานะระบบ...', 'muted');

    try {
      const health = await window.AnalyticsAPI.health();
      const setup = await window.AnalyticsAPI.setupStatus();

      setApiStatus('เชื่อมต่อ API สำเร็จ: ' + (health.message || 'พร้อมใช้งาน'), 'success');

      if (el.healthResult) {
        el.healthResult.textContent =
          'API พร้อมใช้งาน | Setup: ' + (setup.ok ? 'ครบถ้วน' : 'ยังไม่ครบ');
      }

      setSystemStatus('API พร้อมใช้งาน');

      writeLog({
        step: 'health',
        health,
        setup
      });

    } catch (error) {
      setApiStatus(error.message, 'error');

      if (el.healthResult) {
        el.healthResult.textContent = 'เชื่อมต่อ API ไม่สำเร็จ';
      }

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
            await loadSources();
      await loadDashboardOptions();
      await loadManageDashboards();
      if (currentUser && ['Super Admin', 'Admin'].includes(String(currentUser.role || ''))) {
  await loadUsers();
}
      if (currentUser && ['Super Admin', 'Admin'].includes(String(currentUser.role || ''))) {
  await loadAuditLog();
}
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

    const username = el.username ? el.username.value.trim() : '';
    const password = el.password ? el.password.value : '';

    if (!username || !password) {
      setLoginMessage('กรุณากรอกชื่อผู้ใช้และรหัสผ่าน');
      return;
    }

    setLoginMessage('');
    setLoading(true);

    try {
      const data = await window.AnalyticsAPI.login(username, password);

      currentUser = data.user || null;

      renderUser(currentUser);
showDashboard();
applyRoleUi(currentUser);

      writeLog({
        step: 'login_success',
        response: data
      });

      await loadMe();
            await loadSources();
      await loadDashboardOptions();
      await loadManageDashboards();
      if (currentUser && ['Super Admin', 'Admin'].includes(String(currentUser.role || ''))) {
  await loadUsers();
}
      if (currentUser && ['Super Admin', 'Admin'].includes(String(currentUser.role || ''))) {
  await loadAuditLog();
}

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

    if (window.AnalyticsAPI && window.AnalyticsAPI.clearToken) {
      window.AnalyticsAPI.clearToken();
    }

    if (el.username) el.username.value = '';
    if (el.password) el.password.value = '';

    resetWorkingState();
    showLogin();
  }

  async function loadMe() {
    const data = await window.AnalyticsAPI.me();

    currentUser = data.user || null;
    renderUser(currentUser);

    if (el.meResult) {
      el.meResult.textContent = currentUser
        ? `${currentUser.displayName || currentUser.username} (${currentUser.role})`
        : '-';
    }

    writeLog({
      step: 'me',
      response: data
    });

    return data;
  }

  async function handleCreateSource(event) {
    event.preventDefault();

    const payload = {
      name: el.sourceName ? el.sourceName.value.trim() : '',
      spreadsheetId: el.sourceSpreadsheetId ? el.sourceSpreadsheetId.value.trim() : '',
      dataType: el.sourceDataType ? el.sourceDataType.value : 'Other',
      owner: el.sourceOwner ? el.sourceOwner.value.trim() : '',
      description: el.sourceDescription ? el.sourceDescription.value.trim() : ''
    };

    if (!payload.name || !payload.spreadsheetId) {
      setSourceFormMessage('กรุณากรอกชื่อแหล่งข้อมูลและ Google Sheet ID');
      return;
    }

    setSourceFormMessage('');
    setButtonLoading(el.createSourceBtn, true, 'กำลังเพิ่ม...');

    try {
      const data = await window.AnalyticsAPI.createSource(payload);

      setSourceFormMessage('เพิ่มแหล่งข้อมูลสำเร็จ');

      if (el.sourceName) el.sourceName.value = '';
      if (el.sourceSpreadsheetId) el.sourceSpreadsheetId.value = '';
      if (el.sourceOwner) el.sourceOwner.value = '';
      if (el.sourceDescription) el.sourceDescription.value = '';

      writeLog({
        step: 'create_source',
        response: data
      });

      await loadSources();

    } catch (error) {
      setSourceFormMessage(error.message);

      writeLog({
        step: 'create_source_error',
        message: error.message,
        payload: error.payload || null
      });

    } finally {
      setButtonLoading(el.createSourceBtn, false, 'เพิ่มแหล่งข้อมูล');
    }
  }

  async function loadSources() {
    if (el.sourceResult) {
      el.sourceResult.textContent = 'กำลังโหลด...';
    }

    if (el.sourcesList) {
      el.sourcesList.textContent = 'กำลังโหลด...';
      el.sourcesList.classList.add('empty');
    }

    try {
      const data = await window.AnalyticsAPI.listSources();

      sourcesCache = data.sources || [];

      if (el.sourceResult) {
        el.sourceResult.textContent = `พบแหล่งข้อมูล ${data.total || 0} รายการ`;
      }

      renderSources(sourcesCache);

      writeLog({
        step: 'sources',
        response: data
      });

   } catch (error) {
  const message = getFriendlyErrorMessage(error);

  if (handleAuthErrorIfNeeded(error)) {
    return;
  }

  if (el.sourceResult) {
    el.sourceResult.textContent = message;
  }

  if (el.sourcesList) {
    el.sourcesList.textContent = message;
    el.sourcesList.classList.add('empty');
  }

  logApiError('sources_error', error);
}
  }

  function renderSources(sources) {
    if (!el.sourcesList) return;

    el.sourcesList.innerHTML = '';

    if (!sources || !sources.length) {
      el.sourcesList.textContent = 'ยังไม่มีแหล่งข้อมูล ให้เพิ่ม Google Sheet ID ก่อน';
      el.sourcesList.classList.add('empty');
      return;
    }

    el.sourcesList.classList.remove('empty');

    sources.forEach(function (source) {
      const sourceId = source['รหัสแหล่งข้อมูล'] || '';
      const name = source['ชื่อแหล่งข้อมูล'] || '-';
      const sheetId = source['Google Sheet ID'] || '-';
      const type = source['ประเภทข้อมูล'] || '-';
      const active = toBool(source['สถานะใช้งาน']);

      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'source-item' + (sourceId === selectedSourceId ? ' active' : '');

      btn.innerHTML = `
        <div class="item-title">
          <span>${escapeHtml(name)}</span>
          <span class="badge ${active ? 'badge-success' : 'badge-muted'}">${active ? 'Active' : 'Inactive'}</span>
        </div>
        <div class="item-meta">
          <div>ประเภท: ${escapeHtml(type)}</div>
          <div>Source ID: ${escapeHtml(sourceId)}</div>
          <div>Sheet ID: ${escapeHtml(sheetId)}</div>
        </div>
      `;

      btn.addEventListener('click', function () {
        selectedSourceId = sourceId;
        selectedSheetName = '';
        renderSources(sourcesCache);
        renderHeaders(null);
        loadSourceSheets(sourceId);
      });

      el.sourcesList.appendChild(btn);
    });
  }

  async function loadSourceSheets(sourceId) {
  if (!el.sourceSheetsList) return;

  selectedSheetName = '';
  renderHeaders(null);

  el.sourceSheetsList.textContent = 'กำลังอ่านรายชื่อชีท...';
  el.sourceSheetsList.classList.add('empty');

  setSystemStatus('กำลังโหลดรายชื่อชีท');

  try {
    const data = await window.AnalyticsAPI.listSourceSheets({
      sourceId: sourceId,
      includeMeta: false
    });

    renderSourceSheets(data.sheets || []);

    setSystemStatus('โหลดรายชื่อชีทสำเร็จ');

    writeLog({
      step: 'source_sheets',
      response: data
    });

  } catch (error) {
    const message = getFriendlyErrorMessage(error);

    if (handleAuthErrorIfNeeded(error)) {
      return;
    }

    el.sourceSheetsList.innerHTML = `
      <div class="retry-box">
        <div class="retry-title">โหลดรายชื่อชีทไม่สำเร็จ</div>
        <div class="retry-message">${escapeHtml(message)}</div>
        <button id="retryLoadSourceSheetsBtn" class="btn btn-secondary" type="button">
          ลองโหลดรายชื่อชีทอีกครั้ง
        </button>
      </div>
    `;

    el.sourceSheetsList.classList.remove('empty');

    const retryBtn = document.getElementById('retryLoadSourceSheetsBtn');

    if (retryBtn) {
      retryBtn.addEventListener('click', function () {
        loadSourceSheets(sourceId);
      });
    }

    setSystemStatus('โหลดรายชื่อชีทไม่สำเร็จ');

    logApiError('source_sheets_error', error, {
      sourceId: sourceId
    });
  }
}

  function renderSourceSheets(sheets) {
  if (!el.sourceSheetsList) return;

  el.sourceSheetsList.innerHTML = '';

  if (!sheets || !sheets.length) {
    el.sourceSheetsList.textContent = selectedSourceId
      ? 'ไม่พบชีทในแหล่งข้อมูลนี้'
      : 'กรุณาเลือกแหล่งข้อมูลก่อน';

    el.sourceSheetsList.classList.add('empty');
    return;
  }

  el.sourceSheetsList.classList.remove('empty');

  sheets.forEach(function (sheet) {
    const sheetName = sheet.sheetName || '';

    const rowText = sheet.lastRow
      ? Number(sheet.lastRow || 0).toLocaleString() + ' rows'
      : 'เลือกชีท';

    const metaText = sheet.lastColumn
      ? 'คอลัมน์: ' + Number(sheet.lastColumn || 0).toLocaleString()
      : 'กดเลือกเพื่ออ่านหัวคอลัมน์';

    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'sheet-item' + (sheetName === selectedSheetName ? ' active' : '');

    btn.innerHTML = `
      <div class="item-title">
        <span>${escapeHtml(sheetName)}</span>
        <span class="badge badge-muted">${escapeHtml(rowText)}</span>
      </div>
      <div class="item-meta">
        <div>${escapeHtml(metaText)}</div>
      </div>
    `;

    btn.addEventListener('click', function () {
      selectedSheetName = sheetName;
      renderSourceSheets(sheets);
      readHeaders(selectedSourceId, selectedSheetName);
    });

    el.sourceSheetsList.appendChild(btn);
  });
}

  async function readHeaders(sourceId, sheetName) {
  if (!el.headersResult) return;

  el.headersResult.textContent = 'กำลังอ่านหัวคอลัมน์...';
  el.headersResult.classList.add('empty');

  setMappingMessage('');
  setSystemStatus('กำลังอ่านหัวคอลัมน์');

  if (el.saveMappingBtn) {
    el.saveMappingBtn.disabled = true;
  }

  try {
    const data = await window.AnalyticsAPI.readHeaders({
      sourceId: sourceId,
      sheetName: sheetName,
      headerRow: 1
    });

    renderHeaders(data);

    setSystemStatus('อ่านหัวคอลัมน์สำเร็จ');

    writeLog({
      step: 'headers',
      response: data
    });

  } catch (error) {
    const message = getFriendlyErrorMessage(error);

    if (handleAuthErrorIfNeeded(error)) {
      return;
    }

    el.headersResult.innerHTML = `
      <div class="retry-box">
        <div class="retry-title">อ่านหัวคอลัมน์ไม่สำเร็จ</div>
        <div class="retry-message">${escapeHtml(message)}</div>
        <button id="retryReadHeadersBtn" class="btn btn-secondary" type="button">
          ลองอ่านหัวคอลัมน์อีกครั้ง
        </button>
      </div>
    `;

    el.headersResult.classList.remove('empty');

    const retryBtn = document.getElementById('retryReadHeadersBtn');

    if (retryBtn) {
      retryBtn.addEventListener('click', function () {
        readHeaders(sourceId, sheetName);
      });
    }

    setSystemStatus('อ่านหัวคอลัมน์ไม่สำเร็จ');

    logApiError('headers_error', error, {
      sourceId: sourceId,
      sheetName: sheetName
    });
  }
}

  function renderHeaders(data) {
    if (!el.headersResult) return;

    el.headersResult.innerHTML = '';
    lastHeadersData = data || null;

    if (el.saveMappingBtn) {
      el.saveMappingBtn.disabled = !data || !data.headers;
    }

    setMappingMessage('');

    if (!data || !data.headers) {
      el.headersResult.textContent = 'กรุณาเลือกชีทก่อน';
      el.headersResult.classList.add('empty');
      return;
    }

    el.headersResult.classList.remove('empty');

    const box = document.createElement('div');
    box.className = 'mapping-card-list';

    const cards = data.headers.map(function (header, idx) {
      const columnName = header.name || '';
      const guessedType = header.guessedType || 'text';
      const guessedMeaning = header.guessedMeaning || 'text';
      const sampleValue = getFirstSampleValue(data.sampleRows || [], idx);

      return `
        <div class="mapping-card" data-map-row="${idx}">
          <div class="mapping-card-head">
            <div>
              <div class="mapping-index">
                ${escapeHtml(header.index)} / ${escapeHtml(header.columnLetter)}
              </div>
              <div class="mapping-title">
                ${escapeHtml(columnName || '(ไม่มีหัวคอลัมน์)')}
              </div>
            </div>

            <div class="mapping-badges">
              <span class="badge badge-muted">${escapeHtml(guessedType)}</span>
              <span class="badge badge-muted">${escapeHtml(guessedMeaning)}</span>
            </div>
          </div>

          <div class="mapping-form-grid">
            <label class="mapping-control">
              <span>ชื่อแสดงผล</span>
              <input class="map-display-name" type="text" value="${escapeAttr(columnName)}" placeholder="ชื่อแสดงผล">
            </label>

            <label class="mapping-control">
              <span>ประเภทข้อมูล</span>
              <select class="map-data-type">
                ${renderDataTypeOptions(guessedType)}
              </select>
            </label>

            <label class="mapping-control">
              <span>ความหมายข้อมูล</span>
              <select class="map-meaning-type">
                ${renderMeaningTypeOptions(guessedMeaning)}
              </select>
            </label>

            <label class="mapping-control">
              <span>หน่วย</span>
              <input class="map-unit" type="text" placeholder="เช่น รายการ / บาท / ชั่วโมง" value="">
            </label>

            <label class="mapping-control">
              <span>รูปแบบวันที่</span>
              <input
                class="map-date-format"
                type="text"
                placeholder="dd/MM/yyyy HH:mm:ss"
                value="${guessedType === 'datetime' ? 'dd/MM/yyyy HH:mm:ss' : ''}"
              >
            </label>

            <label class="mapping-control">
              <span>หมายเหตุ</span>
              <input class="map-note" type="text" placeholder="หมายเหตุ" value="">
            </label>
          </div>

          <div class="mapping-check-section">
            <div class="mapping-check-title">เลือกการใช้งานของคอลัมน์นี้</div>

            <div class="mapping-checks">
              ${renderCheck('isPrimaryDate', 'วันที่หลัก', guessedMeaning === 'date')}
              ${renderCheck('isPrimaryTime', 'เวลาหลัก', guessedType === 'time')}
              ${renderCheck('isMeasure', 'ตัวเลข', guessedMeaning === 'measure' || guessedType === 'number')}
              ${renderCheck('isCategory', 'หมวดหมู่', guessedMeaning === 'category')}
              ${renderCheck('isStatus', 'สถานะ', guessedMeaning === 'status')}
              ${renderCheck('isPerson', 'บุคคล', guessedMeaning === 'person')}
              ${renderCheck('isTeam', 'ทีม', false)}
              ${renderCheck('isLocation', 'สถานที่', guessedMeaning === 'location')}
              ${renderCheck('isDocument', 'เอกสาร', guessedMeaning === 'document')}
              ${renderCheck('isImage', 'รูปภาพ', guessedMeaning === 'image')}
              ${renderCheck('isUrl', 'URL', guessedMeaning === 'url')}
              ${renderCheck('useFilter', 'ใช้เป็น Filter', shouldDefaultFilter(guessedMeaning, guessedType))}
              ${renderCheck('useExport', 'ใช้ Export', true)}
              ${renderCheck('required', 'บังคับกรอก', false)}
            </div>
          </div>

          <div class="mapping-sample">
            <strong>ตัวอย่างข้อมูล:</strong>
            <span>${escapeHtml(sampleValue || '-')}</span>
          </div>

          <input class="map-column-name" type="hidden" value="${escapeAttr(columnName)}">
          <input class="map-sample-value" type="hidden" value="${escapeAttr(sampleValue)}">
        </div>
      `;
    }).join('');

    box.innerHTML = cards;
    el.headersResult.appendChild(box);

    const sampleBox = document.createElement('div');
    sampleBox.className = 'sample-box';

    sampleBox.innerHTML = `
      <h4>ตัวอย่างข้อมูล 20 แถวแรก</h4>
      <pre>${escapeHtml(JSON.stringify(data.sampleRows || [], null, 2))}</pre>
    `;

    el.headersResult.appendChild(sampleBox);
  }

  async function handleSaveMapping() {
    if (!lastHeadersData || !lastHeadersData.headers || !selectedSourceId || !selectedSheetName) {
      setMappingMessage('กรุณาเลือกแหล่งข้อมูลและชีทก่อน');
      return;
    }

    const selectedSource = sourcesCache.find(function (src) {
      return String(src['รหัสแหล่งข้อมูล'] || '') === selectedSourceId;
    }) || {};

    const fields = collectMappingFields();

    if (!fields.length) {
      setMappingMessage('ไม่มี Mapping ให้บันทึก');
      return;
    }

    setButtonLoading(el.saveMappingBtn, true, 'กำลังบันทึก...');
    setMappingMessage('');

    try {
      const data = await window.AnalyticsAPI.saveMapping({
        sourceId: selectedSourceId,
        sourceName: selectedSource['ชื่อแหล่งข้อมูล'] || '',
        sourceSheetId: '',
        sheetName: selectedSheetName,
        fields: fields
      });

      setMappingMessage('บันทึก Mapping สำเร็จ จำนวน ' + (data.total || fields.length) + ' ฟิลด์');

      if (el.previewResult) {
        el.previewResult.classList.remove('empty');
        el.previewResult.textContent = 'บันทึก Mapping แล้ว สามารถกด “สร้าง Dashboard Preview” ได้';
      }

      writeLog({
        step: 'save_mapping',
        response: data
      });

    } catch (error) {
  showApiError(setMappingMessage, error, 'บันทึก Mapping ไม่สำเร็จ');
  logApiError('save_mapping_error', error, {
    sourceId: selectedSourceId,
    sheetName: selectedSheetName
  });

} finally {
      setButtonLoading(el.saveMappingBtn, false, 'บันทึก Mapping');
    }
  }

  function collectMappingFields() {
    const rows = Array.from(
      el.headersResult.querySelectorAll('.mapping-card[data-map-row]')
    );

    return rows.map(function (row, index) {
      const getValue = function (selector) {
        const input = row.querySelector(selector);
        return input ? input.value : '';
      };

      const getChecked = function (name) {
        const input = row.querySelector('input[data-map-check="' + name + '"]');
        return !!(input && input.checked);
      };

      return {
        columnName: getValue('.map-column-name'),
        displayName: getValue('.map-display-name'),
        dataType: getValue('.map-data-type'),
        meaningType: getValue('.map-meaning-type'),

        isPrimaryDate: getChecked('isPrimaryDate'),
        isPrimaryTime: getChecked('isPrimaryTime'),
        isMeasure: getChecked('isMeasure'),
        isCategory: getChecked('isCategory'),
        isStatus: getChecked('isStatus'),
        isPerson: getChecked('isPerson'),
        isTeam: getChecked('isTeam'),
        isLocation: getChecked('isLocation'),
        isDocument: getChecked('isDocument'),
        isImage: getChecked('isImage'),
        isUrl: getChecked('isUrl'),
        useFilter: getChecked('useFilter'),
        useExport: getChecked('useExport'),
        required: getChecked('required'),

        unit: getValue('.map-unit'),
        dateFormat: getValue('.map-date-format'),
        note: getValue('.map-note'),
        sampleValue: getValue('.map-sample-value'),
        displayOrder: index + 1
      };
    });
  }

 async function handleDashboardPreview() {
  if (!selectedSourceId || !selectedSheetName) {
    setPreviewMessage('กรุณาเลือกแหล่งข้อมูลและชีทก่อน');
    return;
  }

  if (!window.AnalyticsAPI.dashboardPreview) {
    setPreviewMessage('ยังไม่พบฟังก์ชัน dashboardPreview ใน api.js');
    return;
  }

  setButtonLoading(el.previewDashboardBtn, true, 'กำลังสร้าง Preview...');
  setPreviewMessage('กำลังสร้าง Dashboard Preview...');
  setPanelLoading(el.previewResult, 'กำลังสร้าง Dashboard Preview...');

  try {
    const data = await window.AnalyticsAPI.dashboardPreview({
      sourceId: selectedSourceId,
      sheetName: selectedSheetName,
      limit: 1000
    });

    renderDashboardPreview(data);

    writeLog({
      step: 'dashboard_preview',
      response: data
    });

  } catch (error) {
    showApiError(setPreviewMessage, error, 'สร้าง Dashboard Preview ไม่สำเร็จ');

    logApiError('dashboard_preview_error', error, {
      sourceId: selectedSourceId,
      sheetName: selectedSheetName
    });

  } finally {
    setButtonLoading(el.previewDashboardBtn, false, 'สร้าง Dashboard Preview');
  }
}

  function renderDashboardView(data) {
  if (!el.dashboardViewResult) {
    return;
  }

  disposeDashboardCharts();

  const dashboard = data.dashboard || {};
  const source = data.source || {};

  const dashboardName =
    dashboard['ชื่อ Dashboard'] ||
    dashboard.dashboardName ||
    'Dashboard';

  const dashboardType =
    dashboard['ประเภท Dashboard'] ||
    dashboard.dashboardType ||
    '-';

  const sourceName =
    source.sourceName ||
    source['ชื่อแหล่งข้อมูล'] ||
    '-';

  const sheetName =
    source.sheetName ||
    dashboard['รหัสชีทย่อยหลัก'] ||
    '-';

  const totalRowsRead = Number(data.totalRowsRead || 0);
  const totalRowsAfterFilter = Number(data.totalRowsAfterFilter || totalRowsRead || 0);
  const loadedAt = formatClientDateTime(new Date());

const kpisHtml = renderPreviewKpis(data.kpis || []);
const chartsHtml = renderPreviewCharts(data.chartResults || data.charts || []);

 currentDashboardTable = data.table || {
  fields: [],
  rows: []
};

currentDashboardTablePage = 1;

const tableHtml = renderDashboardPagedTable(currentDashboardTable, currentDashboardTablePage);

  el.dashboardViewResult.classList.remove('empty');

  el.dashboardViewResult.innerHTML = `
    <div class="dashboard-title-box dashboard-live-header">
      <div class="dashboard-live-title-row">
        <div>
          <h3>${escapeHtml(dashboardName)}</h3>
          <p>
            ประเภท: ${escapeHtml(dashboardType)}
            · แหล่งข้อมูล: ${escapeHtml(sourceName)}
            · ชีท: ${escapeHtml(sheetName)}
          </p>
        </div>

        <div class="dashboard-live-time">
          <span>โหลดล่าสุด</span>
          <strong>${escapeHtml(loadedAt)}</strong>
        </div>
      </div>

      <div class="dashboard-live-meta-grid">
        <div class="dashboard-live-meta-item">
          <span>จำนวนแถวทั้งหมด</span>
          <strong>${totalRowsRead.toLocaleString()}</strong>
        </div>

        <div class="dashboard-live-meta-item">
          <span>หลังกรอง</span>
          <strong>${totalRowsAfterFilter.toLocaleString()}</strong>
        </div>

        <div class="dashboard-live-meta-item">
          <span>KPI</span>
          <strong>${Number((data.kpis || []).length).toLocaleString()}</strong>
        </div>

        <div class="dashboard-live-meta-item">
          <span>กราฟ</span>
          <strong>${Number((data.chartResults || data.charts || []).length).toLocaleString()}</strong>
        </div>
      </div>
    </div>

    <div class="dashboard-section-title">ตัวชี้วัดสำคัญ</div>
    ${kpisHtml || '<div class="preview-chart-empty">ยังไม่มี KPI สำหรับ Dashboard นี้</div>'}

    <div class="dashboard-section-title">กราฟวิเคราะห์ข้อมูล</div>
    ${chartsHtml || '<div class="preview-chart-empty">ยังไม่มีกราฟสำหรับ Dashboard นี้</div>'}

    <div class="dashboard-section-title">ตารางข้อมูล</div>
<div data-dashboard-table-section>
  ${tableHtml || '<div class="preview-chart-empty">ยังไม่มีข้อมูลตาราง</div>'}
</div>
  `;

  mountQueuedCharts();
}

  function renderDashboardPreview(data) {
  if (!el.previewResult) {
    return;
  }

  if (!data || data.ok === false) {
    setPreviewMessage((data && data.message) || 'ไม่สามารถสร้าง Dashboard Preview ได้');
    return;
  }

  disposeDashboardCharts();
  chartRenderQueue = [];

  const kpis = data.kpis || data.metrics || [];
  const charts = data.charts || data.chartResults || [];
  const table = data.table || {
    fields: [],
    rows: []
  };

  const sourceName =
    data.sourceName ||
    (data.source && data.source.sourceName) ||
    '';

  const sheetName =
    data.sheetName ||
    (data.source && data.source.sheetName) ||
    selectedSheetName ||
    '';

  const totalRowsRead = Number(data.totalRowsRead || 0);
  const totalRowsAfterFilter = Number(data.totalRowsAfterFilter || totalRowsRead || 0);

  el.previewResult.classList.remove('empty');

  el.previewResult.innerHTML = `
    <div class="dashboard-title-box">
      <h3>Dashboard Preview</h3>
      <p>
        แหล่งข้อมูล: ${escapeHtml(sourceName || '-')}
        / ชีท: ${escapeHtml(sheetName || '-')}
        / อ่านข้อมูล ${totalRowsRead.toLocaleString()} แถว
        / หลังกรอง ${totalRowsAfterFilter.toLocaleString()} แถว
      </p>
    </div>

    <div class="dashboard-section-title">ตัวชี้วัดตัวอย่าง</div>
    ${renderPreviewKpis(kpis)}

    <div class="dashboard-section-title">กราฟตัวอย่าง</div>
    ${renderPreviewCharts(charts)}

    <div class="dashboard-section-title">ตารางตัวอย่าง</div>
    ${renderPreviewTable(table)}
  `;

  mountQueuedCharts();
}
  
  async function handleCreateDashboardFromPreview() {
    if (!selectedSourceId || !selectedSheetName) {
      setCreateDashboardMessage('กรุณาเลือกแหล่งข้อมูลและชีทก่อน');
      return;
    }

    const dashboardName = el.dashboardName ? el.dashboardName.value.trim() : '';
    const dashboardType = el.dashboardType ? el.dashboardType.value : 'Custom';

    if (!dashboardName) {
      setCreateDashboardMessage('กรุณากรอกชื่อ Dashboard');
      return;
    }

    if (!window.AnalyticsAPI.createDashboardFromPreview) {
      setCreateDashboardMessage('ยังไม่พบฟังก์ชัน createDashboardFromPreview ใน api.js');
      return;
    }

    setButtonLoading(el.createDashboardBtn, true, 'กำลังบันทึก Dashboard...');
    setCreateDashboardMessage('');

    try {
      const data = await window.AnalyticsAPI.createDashboardFromPreview({
        sourceId: selectedSourceId,
        sheetName: selectedSheetName,
        dashboardName: dashboardName,
        dashboardType: dashboardType,
        description: 'สร้างจาก Mapping ผ่าน Admin Console'
      });

      setCreateDashboardMessage(
        'สร้าง Dashboard สำเร็จ: ' + data.dashboardId +
        ' | KPI ' + data.totalMetrics +
        ' | กราฟ ' + data.totalCharts +
        ' | ตัวกรอง ' + data.totalFilters
      );

      writeLog({
        step: 'create_dashboard_from_preview',
        response: data
      });

           await loadDashboards();
      await loadDashboardOptions();
      await loadManageDashboards();

    } catch (error) {
      setCreateDashboardMessage(error.message);

      writeLog({
        step: 'create_dashboard_from_preview_error',
        message: error.message,
        payload: error.payload || null
      });

    } finally {
      setButtonLoading(el.createDashboardBtn, false, 'บันทึกเป็น Dashboard จริง');
    }
  }

  async function loadDashboardOptions() {
    if (!el.dashboardSelect) {
      return;
    }

    el.dashboardSelect.innerHTML = '<option value="">กำลังโหลด...</option>';
    setDashboardViewMessage('กำลังโหลดรายการ Dashboard...');

    try {
      const data = await window.AnalyticsAPI.listDashboards();
      const dashboards = data.dashboards || [];

      if (!dashboards.length) {
        el.dashboardSelect.innerHTML = '<option value="">ยังไม่มี Dashboard</option>';
        setDashboardViewMessage('ยังไม่มี Dashboard ที่บันทึกไว้');
        return;
      }

      el.dashboardSelect.innerHTML =
        '<option value="">เลือก Dashboard</option>' +
        dashboards.map(function (dash) {
          const id = dash['รหัส Dashboard'] || '';
          const name = dash['ชื่อ Dashboard'] || id;
          const type = dash['ประเภท Dashboard'] || '';

          return `<option value="${escapeAttr(id)}">${escapeHtml(name)} (${escapeHtml(type)})</option>`;
        }).join('');

      setDashboardViewMessage('โหลด Dashboard แล้ว ' + dashboards.length + ' รายการ');

      writeLog({
        step: 'dashboard_options',
        response: data
      });

    } catch (error) {
      el.dashboardSelect.innerHTML = '<option value="">โหลด Dashboard ไม่สำเร็จ</option>';
      setDashboardViewMessage(error.message);

      writeLog({
        step: 'dashboard_options_error',
        message: error.message,
        payload: error.payload || null
      });
    }
  }
async function handleRefreshDashboard() {
  const dashboardId = el.dashboardSelect ? el.dashboardSelect.value : '';

  if (!dashboardId && !currentDashboardId) {
    setDashboardViewMessage('กรุณาเลือก Dashboard ก่อน Refresh');
    return;
  }

  if (!dashboardId && currentDashboardId && el.dashboardSelect) {
    el.dashboardSelect.value = currentDashboardId;
  }

  setDashboardViewMessage('กำลัง Refresh Dashboard...');

  if (el.refreshDashboardBtn) {
    setButtonLoading(el.refreshDashboardBtn, true, 'กำลัง Refresh...');
  }

  try {
    await handleOpenDashboard();

    writeLog({
      step: 'dashboard_refresh',
      dashboardId: dashboardId || currentDashboardId,
      status: 'success'
    });

  } catch (error) {
    setDashboardViewMessage(error.message || 'Refresh Dashboard ไม่สำเร็จ');

    writeLog({
      step: 'dashboard_refresh_error',
      message: error.message,
      payload: error.payload || null
    });

  } finally {
    if (el.refreshDashboardBtn) {
      setButtonLoading(el.refreshDashboardBtn, false, 'Refresh');
    }
  }
}
  async function handleOpenDashboard() {
    const dashboardId = el.dashboardSelect ? el.dashboardSelect.value : '';
    const limit = el.dashboardLimit ? Number(el.dashboardLimit.value || 1000) : 1000;

    if (!dashboardId) {
      setDashboardViewMessage('กรุณาเลือก Dashboard');
      return;
    }

    if (!window.AnalyticsAPI.dashboardView) {
      setDashboardViewMessage('ยังไม่พบฟังก์ชัน dashboardView ใน api.js');
      return;
    }

    currentDashboardId = dashboardId;

    setButtonLoading(el.openDashboardBtn, true, 'กำลังเปิด Dashboard...');
setDashboardViewMessage('กำลังโหลด Dashboard...');
setPanelLoading(el.dashboardViewResult, 'กำลังโหลด Dashboard...');

    try {
      const data = await window.AnalyticsAPI.dashboardView({
        dashboardId: dashboardId,
        limit: limit,
        filters: []
      });

      renderDashboardView(data);
      renderDashboardFilters(data.filters || []);
      setDashboardViewMessage('โหลด Dashboard สำเร็จ');

      writeLog({
        step: 'dashboard_view',
        response: data
      });

    } catch (error) {
  const message = getFriendlyErrorMessage(error);

  showApiError(setDashboardViewMessage, error, 'เปิด Dashboard ไม่สำเร็จ');

  if (el.dashboardViewResult) {
    el.dashboardViewResult.textContent = message;
    el.dashboardViewResult.classList.add('empty');
  }

  logApiError('dashboard_view_error', error, {
    dashboardId: dashboardId,
    limit: limit
  });

} finally {
      setButtonLoading(el.openDashboardBtn, false, 'เปิด Dashboard');
    }
  }

  async function handleApplyDashboardFilter() {
    const dashboardId = currentDashboardId || (el.dashboardSelect ? el.dashboardSelect.value : '');
    const limit = el.dashboardLimit ? Number(el.dashboardLimit.value || 1000) : 1000;

    if (!dashboardId) {
      setDashboardViewMessage('กรุณาเลือก Dashboard ก่อน');
      return;
    }

    const filters = collectDashboardFilters();

    setButtonLoading(el.applyDashboardFilterBtn, true, 'กำลังกรอง...');
    setDashboardViewMessage('กำลังกรองข้อมูล...');
    setPanelLoading(el.dashboardViewResult, 'กำลังกรองข้อมูล Dashboard...');

    try {
      const data = await window.AnalyticsAPI.dashboardView({
        dashboardId: dashboardId,
        limit: limit,
        filters: filters
      });

      renderDashboardView(data);
      renderDashboardFilters(data.filters || [], filters);

      setDashboardViewMessage(
        'กรองข้อมูลสำเร็จ: จาก ' +
        Number(data.totalRowsRead || 0).toLocaleString() +
        ' แถว เหลือ ' +
        Number(data.totalRowsAfterFilter || 0).toLocaleString() +
        ' แถว'
      );

      writeLog({
        step: 'dashboard_filter',
        filters: filters,
        response: data
      });

    } catch (error) {
  showApiError(setDashboardViewMessage, error, 'กรองข้อมูล Dashboard ไม่สำเร็จ');
  logApiError('dashboard_filter_error', error, {
    dashboardId: dashboardId,
    filters: filters
  });

} finally {
      setButtonLoading(el.applyDashboardFilterBtn, false, 'กรองข้อมูล Dashboard');
    }
  }

  async function handleResetDashboardFilter() {
    if (!el.dashboardFilterBox) {
      return;
    }

    const inputs = el.dashboardFilterBox.querySelectorAll('input, select');

    inputs.forEach(function (input) {
      input.value = '';
    });

    await handleApplyDashboardFilter();
  }

  function renderSavedDashboard(data) {
  if (!el.dashboardViewResult) {
    return;
  }

  if (!data || !data.ok) {
    el.dashboardViewResult.textContent = (data && data.message) || 'โหลด Dashboard ไม่สำเร็จ';
    el.dashboardViewResult.classList.add('empty');
    return;
  }

  disposeDashboardCharts();
  chartRenderQueue = [];

  const dashboard = data.dashboard || {};
  const dashboardName = dashboard['ชื่อ Dashboard'] || '-';
  const dashboardDesc = dashboard['คำอธิบาย'] || '';
  const source = data.source || {};

  el.dashboardViewResult.classList.remove('empty');

  el.dashboardViewResult.innerHTML = `
    <div class="dashboard-title-box">
      <h3>${escapeHtml(dashboardName)}</h3>
      <p>
        ${escapeHtml(dashboardDesc)}
        ${dashboardDesc ? ' | ' : ''}
        แหล่งข้อมูล: ${escapeHtml(source.sourceName || '')}
        / ชีท: ${escapeHtml(source.sheetName || '')}
        / อ่านข้อมูล ${Number(data.totalRowsRead || 0).toLocaleString()} แถว
        / หลังกรอง ${Number(data.totalRowsAfterFilter || data.totalRowsRead || 0).toLocaleString()} แถว
      </p>
    </div>

    <div class="dashboard-section-title">ตัวชี้วัด</div>
    ${renderPreviewKpis(data.kpis || [])}

    <div class="dashboard-section-title">กราฟ</div>
    ${renderPreviewCharts(data.chartResults || [])}

    <div class="dashboard-section-title">ตารางข้อมูล</div>
    ${renderPreviewTable(data.table || { fields: [], rows: [] })}
  `;
   
  mountQueuedCharts();
    bindDashboardTablePagination();
}

  function renderDashboardFilters(filters, appliedFilters) {
    if (!el.dashboardFilterBox) {
      return;
    }

    currentDashboardFilters = filters || [];
    appliedFilters = appliedFilters || [];

    if (!filters || !filters.length) {
      el.dashboardFilterBox.classList.add('empty');
      el.dashboardFilterBox.textContent = 'Dashboard นี้ยังไม่มีตัวกรองที่ตั้งไว้';
      return;
    }

    el.dashboardFilterBox.classList.remove('empty');

    const html = filters.map(function (filter, index) {
      const name = filter['ชื่อตัวกรอง'] || filter['คอลัมน์ที่ใช้กรอง'] || '-';
      const column = filter['คอลัมน์ที่ใช้กรอง'] || '';
      const type = filter['ประเภทตัวกรอง'] || 'text_search';
      const applied = findAppliedFilter(appliedFilters, column);

      if (type === 'date_range') {
        return `
          <label class="dashboard-filter-item" data-filter-index="${index}">
            <span>${escapeHtml(name)}</span>
            <div class="dashboard-filter-date-pair">
              <input type="date" data-dashboard-filter-field="${escapeAttr(column)}" data-dashboard-filter-type="date_range" data-dashboard-filter-part="from" value="${escapeAttr(applied.from || '')}">
              <input type="date" data-dashboard-filter-field="${escapeAttr(column)}" data-dashboard-filter-type="date_range" data-dashboard-filter-part="to" value="${escapeAttr(applied.to || '')}">
            </div>
          </label>
        `;
      }

      if (type === 'number_range') {
        return `
          <label class="dashboard-filter-item" data-filter-index="${index}">
            <span>${escapeHtml(name)}</span>
            <div class="dashboard-filter-number-pair">
              <input type="number" data-dashboard-filter-field="${escapeAttr(column)}" data-dashboard-filter-type="number_range" data-dashboard-filter-part="min" placeholder="ต่ำสุด" value="${escapeAttr(applied.min || '')}">
              <input type="number" data-dashboard-filter-field="${escapeAttr(column)}" data-dashboard-filter-type="number_range" data-dashboard-filter-part="max" placeholder="สูงสุด" value="${escapeAttr(applied.max || '')}">
            </div>
          </label>
        `;
      }

      return `
        <label class="dashboard-filter-item" data-filter-index="${index}">
          <span>${escapeHtml(name)}</span>
          <input
            type="text"
            data-dashboard-filter-field="${escapeAttr(column)}"
            data-dashboard-filter-type="${escapeAttr(type)}"
            placeholder="ค้นหา / กรองค่า"
            value="${escapeAttr(applied.value || '')}"
          >
        </label>
      `;
    }).join('');

    el.dashboardFilterBox.innerHTML = `
      <div class="dashboard-filter-grid">
        ${html}
      </div>
    `;
  }

  function collectDashboardFilters() {
    if (!el.dashboardFilterBox) {
      return [];
    }

    const map = {};
    const inputs = Array.from(el.dashboardFilterBox.querySelectorAll('[data-dashboard-filter-field]'));

    inputs.forEach(function (input) {
      const field = input.getAttribute('data-dashboard-filter-field') || '';
      const type = input.getAttribute('data-dashboard-filter-type') || 'text_search';
      const part = input.getAttribute('data-dashboard-filter-part') || '';
      const value = input.value;

      if (!field) {
        return;
      }

      if (!map[field]) {
        map[field] = {
          field: field,
          type: type,
          value: ''
        };
      }

      if (type === 'date_range') {
        if (part === 'from') map[field].from = value;
        if (part === 'to') map[field].to = value;
        map[field].value = map[field].from || map[field].to || '';
        return;
      }

      if (type === 'number_range') {
        if (part === 'min') map[field].min = value;
        if (part === 'max') map[field].max = value;
        map[field].value = map[field].min || map[field].max || '';
        return;
      }

      map[field].value = value;
    });

    return Object.keys(map)
      .map(function (key) {
        return map[key];
      })
      .filter(function (f) {
        if (f.type === 'date_range') {
          return String(f.from || '').trim() || String(f.to || '').trim();
        }

        if (f.type === 'number_range') {
          return String(f.min || '').trim() || String(f.max || '').trim();
        }

        return String(f.value || '').trim();
      });
  }

  function findAppliedFilter(appliedFilters, field) {
    for (let i = 0; i < appliedFilters.length; i++) {
      if (String(appliedFilters[i].field || '') === String(field || '')) {
        return appliedFilters[i];
      }
    }

    return {};
  }

  async function loadDashboards() {
    if (el.dashboardResult) {
      el.dashboardResult.textContent = 'กำลังโหลด...';
    }

    try {
      const data = await window.AnalyticsAPI.listDashboards();

      if (el.dashboardResult) {
        el.dashboardResult.textContent = `พบ Dashboard ${data.total || 0} รายการ`;
      }

      writeLog({
        step: 'dashboards',
        response: data
      });

    } catch (error) {
      if (el.dashboardResult) {
        el.dashboardResult.textContent = error.message;
      }

      writeLog({
        step: 'dashboards_error',
        message: error.message,
        payload: error.payload || null
      });
    }
  }

  function renderPreviewKpis(kpis) {
    if (!kpis.length) {
      return '';
    }

    const html = kpis.map(function (kpi) {
      return `
        <div class="preview-kpi">
          <div class="preview-kpi-title">${escapeHtml(kpi.title || '-')}</div>
          <div class="preview-kpi-value">${formatNumber(kpi.value)}</div>
          <div class="preview-kpi-unit">${escapeHtml(kpi.unit || '')}</div>
        </div>
      `;
    }).join('');

    return `<div class="preview-kpi-grid">${html}</div>`;
  }

  function renderPreviewCharts(charts) {
  if (!charts || !charts.length) {
    return '';
  }

  const html = charts.map(function (chart) {
    const id = 'echart_' + (++chartUid);
    const type = String(chart.type || 'bar').toLowerCase();
    const sizeClass = type === 'line' ? 'chart-large' : '';

    chartRenderQueue.push({
      id: id,
      chart: chart
    });

    return `
      <div class="preview-chart ${sizeClass}">
        <h4>${escapeHtml(chart.title || '-')}</h4>
        <div id="${escapeAttr(id)}" class="echart-box"></div>
      </div>
    `;
  }).join('');

  return `
    <div class="preview-chart-grid">
      ${html}
    </div>
  `;
}

  function renderSimpleBarChart(items) {
    if (!items.length) {
      return '<p class="empty">ไม่มีข้อมูลสำหรับกราฟ</p>';
    }

    const max = Math.max.apply(null, items.map(function (x) {
      return Number(x.value || 0);
    })) || 1;

    const html = items.map(function (item) {
      const value = Number(item.value || 0);
      const percent = Math.max(2, Math.round((value / max) * 100));

      return `
        <div class="simple-bar-item">
          <div class="simple-bar-label" title="${escapeAttr(item.name || '')}">
            ${escapeHtml(item.name || '(ว่าง)')}
          </div>
          <div class="simple-bar-track">
            <div class="simple-bar-fill" style="width:${percent}%"></div>
          </div>
          <div class="simple-bar-value">${formatNumber(value)}</div>
        </div>
      `;
    }).join('');

    return `<div class="simple-bar-list">${html}</div>`;
  }
function renderDashboardPagedTable(table, page) {
  table = table || {
    fields: [],
    rows: []
  };

  const fields = table.fields || [];
  const rows = table.rows || [];

  if (!fields.length || !rows.length) {
    return '<div class="preview-chart-empty">ยังไม่มีข้อมูลตาราง</div>';
  }

  page = Math.max(1, Number(page || 1));
  const pageSize = DASHBOARD_TABLE_PAGE_SIZE;
  const totalRows = rows.length;
  const totalPages = Math.max(1, Math.ceil(totalRows / pageSize));

  if (page > totalPages) {
    page = totalPages;
  }

  const start = (page - 1) * pageSize;
  const end = Math.min(start + pageSize, totalRows);
  const pageRows = rows.slice(start, end);

  const headersHtml = fields.map(function (field) {
    return `<th>${escapeHtml(field.displayName || field.columnName || '')}</th>`;
  }).join('');

  const rowsHtml = pageRows.map(function (row) {
    const cells = fields.map(function (field) {
      const key = field.displayName || field.columnName;
      return `<td>${escapeHtml(row[key] || '')}</td>`;
    }).join('');

    return `<tr>${cells}</tr>`;
  }).join('');

  return `
    <div class="dashboard-table-toolbar">
      <div>
        แสดง ${Number(start + 1).toLocaleString()}-${Number(end).toLocaleString()}
        จาก ${Number(totalRows).toLocaleString()} แถว
      </div>

      <div class="dashboard-table-pager">
        <button
          id="dashboardTablePrevBtn"
          class="btn btn-ghost"
          type="button"
          ${page <= 1 ? 'disabled' : ''}
        >
          ก่อนหน้า
        </button>

        <span>หน้า ${page.toLocaleString()} / ${totalPages.toLocaleString()}</span>

        <button
          id="dashboardTableNextBtn"
          class="btn btn-ghost"
          type="button"
          ${page >= totalPages ? 'disabled' : ''}
        >
          ถัดไป
        </button>
      </div>
    </div>

    <div class="preview-table-wrap">
      <table class="preview-table">
        <thead>
          <tr>${headersHtml}</tr>
        </thead>
        <tbody>${rowsHtml}</tbody>
      </table>
    </div>
  `;
}
  function bindDashboardTablePagination() {
  const prevBtn = document.getElementById('dashboardTablePrevBtn');
  const nextBtn = document.getElementById('dashboardTableNextBtn');

  if (prevBtn) {
    prevBtn.addEventListener('click', function () {
      changeDashboardTablePage(-1);
    });
  }

  if (nextBtn) {
    nextBtn.addEventListener('click', function () {
      changeDashboardTablePage(1);
    });
  }
}


function changeDashboardTablePage(direction) {
  if (!currentDashboardTable || !el.dashboardViewResult) {
    return;
  }

  const rows = currentDashboardTable.rows || [];
  const totalPages = Math.max(1, Math.ceil(rows.length / DASHBOARD_TABLE_PAGE_SIZE));

  currentDashboardTablePage = Math.max(
    1,
    Math.min(totalPages, currentDashboardTablePage + Number(direction || 0))
  );

  const tableSection = el.dashboardViewResult.querySelector('[data-dashboard-table-section]');

  if (!tableSection) {
    return;
  }

  tableSection.innerHTML = renderDashboardPagedTable(
    currentDashboardTable,
    currentDashboardTablePage
  );

  bindDashboardTablePagination();
}
  function renderPreviewTable(table) {
    const fields = table.fields || [];
    const rows = table.rows || [];

    if (!fields.length || !rows.length) {
      return '';
    }

    const headersHtml = fields.map(function (field) {
      return `<th>${escapeHtml(field.displayName || field.columnName || '')}</th>`;
    }).join('');

    const rowsHtml = rows.map(function (row) {
      const cells = fields.map(function (field) {
        const key = field.displayName || field.columnName;
        return `<td>${escapeHtml(row[key] || '')}</td>`;
      }).join('');

      return `<tr>${cells}</tr>`;
    }).join('');

    return `
      <div class="preview-table-wrap">
        <table class="preview-table">
          <thead>
            <tr>${headersHtml}</tr>
          </thead>
          <tbody>${rowsHtml}</tbody>
        </table>
      </div>
    `;
  }

  function renderUser(user) {
    if (!user) {
      if (el.userDisplayName) el.userDisplayName.textContent = '-';
      if (el.userRole) el.userRole.textContent = '-';
      return;
    }

    if (el.userDisplayName) el.userDisplayName.textContent = user.displayName || user.username || '-';
    if (el.userRole) el.userRole.textContent = user.role || '-';

    if (user.mustChangePassword) {
  setSystemStatus('เข้าสู่ระบบสำเร็จ - แนะนำให้เปลี่ยนรหัสผ่านเริ่มต้น');
} else {
  setSystemStatus('เข้าสู่ระบบสำเร็จ');
}

applyRoleUi(user);
    
  }

 function showLogin() {
  if (el.loginView) el.loginView.classList.remove('hidden');
  if (el.dashboardView) el.dashboardView.classList.add('hidden');

  document.body.classList.remove(
    'mode-admin',
    'mode-user',
    'role-viewer-only',
    'role-super-admin',
    'role-admin',
    'role-editor',
    'role-user',
    'role-viewer'
  );

  if (el.modeSwitcher) {
    el.modeSwitcher.classList.add('hidden');
  }

  if (el.adminModeBtn) {
    el.adminModeBtn.classList.remove('hidden', 'active');
  }

  if (el.userModeBtn) {
    el.userModeBtn.classList.remove('hidden', 'active');
  }
}
 function showDashboard() {
  if (el.loginView) el.loginView.classList.add('hidden');
  if (el.dashboardView) el.dashboardView.classList.remove('hidden');

  if (el.modeSwitcher) {
    el.modeSwitcher.classList.remove('hidden');
  }

  if (currentUser) {
    applyRoleUi(currentUser);
  }
}
  function ensureGlobalLoadingOverlay() {
  let overlay = document.getElementById('globalLoadingOverlay');

  if (overlay) {
    el.globalLoadingOverlay = overlay;
    el.globalLoadingText = document.getElementById('globalLoadingText');
    return overlay;
  }

  overlay = document.createElement('div');
  overlay.id = 'globalLoadingOverlay';
  overlay.className = 'global-loading-overlay hidden';
  overlay.innerHTML = `
    <div class="global-loading-card">
      <div class="loading-spinner"></div>
      <div id="globalLoadingText" class="global-loading-text">
        กำลังประมวลผล...
      </div>
    </div>
  `;

  document.body.appendChild(overlay);

  el.globalLoadingOverlay = overlay;
  el.globalLoadingText = document.getElementById('globalLoadingText');

  return overlay;
}


function showGlobalLoading(message) {
  const overlay = ensureGlobalLoadingOverlay();

  if (el.globalLoadingText) {
    el.globalLoadingText.textContent = message || 'กำลังประมวลผล...';
  }

  overlay.classList.remove('hidden');
  document.body.classList.add('is-loading');
}


function hideGlobalLoading() {
  const overlay = ensureGlobalLoadingOverlay();
  overlay.classList.add('hidden');
  document.body.classList.remove('is-loading');
}


function setPanelLoading(target, message) {
  if (!target) {
    return;
  }

  target.classList.remove('empty');
  target.innerHTML = `
    <div class="panel-loading-box">
      <div class="loading-spinner small"></div>
      <div>${escapeHtml(message || 'กำลังโหลดข้อมูล...')}</div>
    </div>

    <div class="skeleton-grid">
      <div class="skeleton-card"></div>
      <div class="skeleton-card"></div>
      <div class="skeleton-card"></div>
    </div>
  `;
}


function setInlineLoading(target, message) {
  if (!target) {
    return;
  }

  target.innerHTML = `
    <span class="inline-loading">
      <span class="loading-spinner mini"></span>
      ${escapeHtml(message || 'กำลังโหลด...')}
    </span>
  `;
}
  function setButtonLoading(button, isLoading, text) {
    if (!button) return;

    button.disabled = !!isLoading;

    if (text) {
      button.textContent = text;
    }
  }

  function setApiStatus(message, type) {
    if (!el.apiStatus) return;

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
    if (el.systemStatusText) {
      el.systemStatusText.textContent = message || '';
    }
  }

  function setLoginMessage(message) {
    if (el.loginMessage) {
      el.loginMessage.textContent = message || '';
    }
  }

  function setSourceFormMessage(message) {
    if (el.sourceFormMessage) {
      el.sourceFormMessage.textContent = message || '';
    }
  }

  function setMappingMessage(message) {
    if (el.mappingMessage) {
      el.mappingMessage.textContent = message || '';
    }
  }

function setPreviewMessage(message) {
  if (!el.previewResult) {
    return;
  }

  if (!message) {
    return;
  }

  el.previewResult.textContent = message;
  el.previewResult.classList.add('empty');
}
  function setCreateDashboardMessage(message) {
    if (el.createDashboardMessage) {
      el.createDashboardMessage.textContent = message || '';
    }
  }

  function setDashboardViewMessage(message) {
    if (el.dashboardViewMessage) {
      el.dashboardViewMessage.textContent = message || '';
    }
  }

  function setLoading(isLoading) {
    if (!el.loginBtn) return;

    el.loginBtn.disabled = !!isLoading;
    el.loginBtn.textContent = isLoading ? 'กำลังเข้าสู่ระบบ...' : 'เข้าสู่ระบบ';
  }

  function getFriendlyErrorMessage(error) {
  if (!error) {
    return 'เกิดข้อผิดพลาดที่ไม่ทราบสาเหตุ';
  }

  const message = String(error.message || '').trim();

  if (message) {
    return message;
  }

  if (error.payload && error.payload.message) {
    return String(error.payload.message);
  }

  return 'เกิดข้อผิดพลาด กรุณาลองใหม่อีกครั้ง';
}


function handleAuthErrorIfNeeded(error) {
  if (!error) {
    return false;
  }

  const message = String(error.message || '');
  const rawMessage = String(error.rawMessage || '');

  const isAuthError =
    error.isAuthError ||
    message.includes('เซสชันหมดอายุ') ||
    rawMessage.includes('Token หมดอายุ') ||
    rawMessage.includes('กรุณาเข้าสู่ระบบก่อนใช้งาน') ||
    rawMessage.includes('Token ไม่ถูกต้อง');

  if (!isAuthError) {
    return false;
  }

  if (window.AnalyticsAPI && window.AnalyticsAPI.clearToken) {
    window.AnalyticsAPI.clearToken();
  }

  setSystemStatus('เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่');

  window.alert('เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่');

  resetWorkingState();
  showLogin();

  return true;
}


function showApiError(targetSetter, error, fallbackMessage) {
  if (handleAuthErrorIfNeeded(error)) {
    return;
  }

  const message = getFriendlyErrorMessage(error) || fallbackMessage || 'เกิดข้อผิดพลาด';

  if (typeof targetSetter === 'function') {
    targetSetter(message);
  } else {
    setSystemStatus(message);
  }
}


function logApiError(step, error, extra) {
  writeLog({
    step: step,
    message: getFriendlyErrorMessage(error),
    rawMessage: error && error.rawMessage ? error.rawMessage : '',
    status: error && error.status ? error.status : '',
    payload: error && error.payload ? error.payload : null,
    extra: extra || null
  });
}

  async function loadUsers() {
  if (!el.userManageList) {
    return;
  }

  el.userManageList.classList.remove('empty');
  setPanelLoading(el.userManageList, 'กำลังโหลดรายชื่อผู้ใช้...');

  if (el.reloadUsersBtn) {
    setButtonLoading(el.reloadUsersBtn, true, 'กำลังโหลด...');
  }

  try {
    const data = await window.AnalyticsAPI.listUsers();

    usersCache = data.users || [];
    renderUsers(usersCache);

    setUserManageMessage('โหลดรายชื่อผู้ใช้สำเร็จ');

    writeLog({
      step: 'users_load',
      response: {
        ok: data.ok,
        total: data.total
      }
    });

  } catch (error) {
    const message = getFriendlyErrorMessage(error);

    if (handleAuthErrorIfNeeded(error)) {
      return;
    }

    el.userManageList.classList.add('empty');
    el.userManageList.textContent = message;

    logApiError('users_load_error', error);

  } finally {
    if (el.reloadUsersBtn) {
      setButtonLoading(el.reloadUsersBtn, false, 'โหลดรายชื่อผู้ใช้');
    }
  }
}


function renderUsers(users) {
  if (!el.userManageList) {
    return;
  }

  users = users || [];

  if (el.userManageCount) {
    el.userManageCount.textContent = users.length.toLocaleString() + ' รายการ';
  }

  if (!users.length) {
    el.userManageList.classList.add('empty');
    el.userManageList.textContent = 'ยังไม่มีผู้ใช้งาน';
    return;
  }

  el.userManageList.classList.remove('empty');

  el.userManageList.innerHTML = users.map(function (user) {
    const active = String(user.status || '').toLowerCase() === 'active';

    return `
      <article class="user-card" data-user-id="${escapeAttr(user.userId || '')}">
        <div class="user-card-head">
          <div>
            <h4>${escapeHtml(user.displayName || user.username || '-')}</h4>
            <p>${escapeHtml(user.username || '-')} · ${escapeHtml(user.email || '-')}</p>
          </div>

          <span class="user-role-pill">${escapeHtml(user.role || '-')}</span>
        </div>

        <div class="user-card-meta">
          <div>
            <span>User ID</span>
            <strong>${escapeHtml(user.userId || '-')}</strong>
          </div>

          <div>
            <span>Status</span>
            <strong class="${active ? 'success-text' : 'error-text'}">
              ${escapeHtml(user.status || '-')}
            </strong>
          </div>

          <div>
            <span>Login</span>
            <strong>${Number(user.loginCount || 0).toLocaleString()} ครั้ง</strong>
          </div>

          <div>
            <span>Last Login</span>
            <strong>${escapeHtml(user.lastLoginAt || '-')}</strong>
          </div>
        </div>

        <div class="user-card-permissions">
          ${renderUserPermissionBadge('Export', user.canExport)}
          ${renderUserPermissionBadge('Add Source', user.canAddSource)}
          ${renderUserPermissionBadge('Edit Dashboard', user.canEditDashboard)}
          ${renderUserPermissionBadge('Audit Log', user.canViewAuditLog)}
          ${renderUserPermissionBadge('Manage User', user.canManageUser)}
          ${user.mustChangePassword ? '<span class="user-permission-badge warning">ต้องเปลี่ยนรหัสผ่าน</span>' : ''}
        </div>

        <div class="user-card-actions">
          <button class="btn btn-secondary btn-select-user" type="button" data-user-id="${escapeAttr(user.userId || '')}">
            เลือกแก้ไข
          </button>

          <button class="btn btn-ghost btn-reset-user-password" type="button" data-user-id="${escapeAttr(user.userId || '')}">
            Reset Password
          </button>

          <button class="btn ${active ? 'btn-danger' : 'btn-primary'} btn-toggle-user-status" type="button" data-user-id="${escapeAttr(user.userId || '')}" data-status="${active ? 'inactive' : 'active'}">
            ${active ? 'ปิดใช้งาน' : 'เปิดใช้งาน'}
          </button>
        </div>
      </article>
    `;
  }).join('');

  bindUserCardActions();
}


function renderUserPermissionBadge(label, value) {
  return `
    <span class="user-permission-badge ${value ? 'on' : 'off'}">
      ${escapeHtml(label)}: ${value ? 'On' : 'Off'}
    </span>
  `;
}


function bindUserCardActions() {
  if (!el.userManageList) {
    return;
  }

  el.userManageList.querySelectorAll('.btn-select-user').forEach(function (btn) {
    btn.addEventListener('click', function () {
      const userId = btn.getAttribute('data-user-id') || '';
      selectUserForEdit(userId);
    });
  });

  el.userManageList.querySelectorAll('.btn-reset-user-password').forEach(function (btn) {
    btn.addEventListener('click', function () {
      const userId = btn.getAttribute('data-user-id') || '';
      resetPasswordByUserId(userId);
    });
  });

  el.userManageList.querySelectorAll('.btn-toggle-user-status').forEach(function (btn) {
    btn.addEventListener('click', function () {
      const userId = btn.getAttribute('data-user-id') || '';
      const status = btn.getAttribute('data-status') || '';
      setUserStatusById(userId, status);
    });
  });
}


function selectUserForEdit(userId) {
  userId = String(userId || '').trim();

  const user = usersCache.find(function (item) {
    return String(item.userId || '') === userId;
  });

  if (!user) {
    setUserManageMessage('ไม่พบผู้ใช้ที่เลือก');
    return;
  }

  selectedManageUserId = userId;

  if (el.userManageUsername) {
    el.userManageUsername.value = user.username || '';
    el.userManageUsername.disabled = true;
  }

  if (el.userManageDisplayName) el.userManageDisplayName.value = user.displayName || '';
  if (el.userManageEmail) el.userManageEmail.value = user.email || '';
  if (el.userManageRole) el.userManageRole.value = user.role || 'User';
  if (el.userManageStatus) el.userManageStatus.value = user.status || 'active';
  if (el.userManagePassword) el.userManagePassword.value = '';
  if (el.userManageDashboards) el.userManageDashboards.value = user.dashboards || '';
  setUserManageMessage(
  'เลือกผู้ใช้แล้ว: ' +
  (user.displayName || user.username || userId) +
  ' | ช่อง Dashboard ที่ดูได้ ใช้รูปแบบ DASH_xxx,DASH_yyy'
);

  if (el.userCanExport) el.userCanExport.checked = !!user.canExport;
  if (el.userCanAddSource) el.userCanAddSource.checked = !!user.canAddSource;
  if (el.userCanEditDashboard) el.userCanEditDashboard.checked = !!user.canEditDashboard;
  if (el.userCanViewAuditLog) el.userCanViewAuditLog.checked = !!user.canViewAuditLog;
  if (el.userCanManageUser) el.userCanManageUser.checked = !!user.canManageUser;

  setUserManageMessage('เลือกผู้ใช้แล้ว: ' + (user.displayName || user.username || userId));

  const target = document.querySelector('.user-management-section');

  if (target) {
    try {
      target.scrollIntoView({
        behavior: 'smooth',
        block: 'start'
      });
    } catch (err) {
      target.scrollIntoView();
    }
  }
}


function collectUserPayload() {
  return {
    username: el.userManageUsername ? el.userManageUsername.value.trim() : '',
    displayName: el.userManageDisplayName ? el.userManageDisplayName.value.trim() : '',
    email: el.userManageEmail ? el.userManageEmail.value.trim() : '',
    role: el.userManageRole ? el.userManageRole.value : 'User',
    status: el.userManageStatus ? el.userManageStatus.value : 'active',
    password: el.userManagePassword ? el.userManagePassword.value.trim() : '',
    dashboards: el.userManageDashboards ? el.userManageDashboards.value.trim() : '',

    canExport: el.userCanExport ? el.userCanExport.checked : false,
    canAddSource: el.userCanAddSource ? el.userCanAddSource.checked : false,
    canEditDashboard: el.userCanEditDashboard ? el.userCanEditDashboard.checked : false,
    canViewAuditLog: el.userCanViewAuditLog ? el.userCanViewAuditLog.checked : false,
    canManageUser: el.userCanManageUser ? el.userCanManageUser.checked : false
  };
}


async function handleCreateUser() {
  const payload = collectUserPayload();

  if (!payload.username || !payload.displayName) {
    setUserManageMessage('กรุณากรอกชื่อผู้ใช้และชื่อแสดงผล');
    return;
  }

  setButtonLoading(el.createUserBtn, true, 'กำลังเพิ่มผู้ใช้...');
  setUserManageMessage('');

  try {
    const data = await window.AnalyticsAPI.createUser(payload);

    setUserManageMessage(
      'เพิ่มผู้ใช้สำเร็จ: ' +
      data.username +
      ' | รหัสผ่านเริ่มต้น: ' +
      data.temporaryPassword
    );

    window.alert(
      'เพิ่มผู้ใช้สำเร็จ\n\n' +
      'Username: ' + data.username + '\n' +
      'Password: ' + data.temporaryPassword + '\n\n' +
      'กรุณาจดรหัสผ่านนี้ไว้ ระบบจะแสดงครั้งเดียว'
    );

    resetUserForm();
    await loadUsers();

  } catch (error) {
    showApiError(setUserManageMessage, error, 'เพิ่มผู้ใช้ไม่สำเร็จ');
    logApiError('create_user_error', error, {
      username: payload.username
    });

  } finally {
    setButtonLoading(el.createUserBtn, false, 'เพิ่มผู้ใช้ใหม่');
  }
}


async function handleUpdateUser() {
  if (!selectedManageUserId) {
    setUserManageMessage('กรุณาเลือกผู้ใช้ก่อนแก้ไข');
    return;
  }

  const payload = collectUserPayload();
  payload.userId = selectedManageUserId;

  setButtonLoading(el.updateUserBtn, true, 'กำลังบันทึก...');
  setUserManageMessage('');

  try {
    const data = await window.AnalyticsAPI.updateUser(payload);

    setUserManageMessage(data.message || 'แก้ไขผู้ใช้สำเร็จ');

    await loadUsers();

  } catch (error) {
    showApiError(setUserManageMessage, error, 'แก้ไขผู้ใช้ไม่สำเร็จ');
    logApiError('update_user_error', error, {
      userId: selectedManageUserId
    });

  } finally {
    setButtonLoading(el.updateUserBtn, false, 'บันทึกการแก้ไข');
  }
}


async function handleResetUserPassword() {
  if (!selectedManageUserId) {
    setUserManageMessage('กรุณาเลือกผู้ใช้ก่อน Reset Password');
    return;
  }

  await resetPasswordByUserId(selectedManageUserId);
}


async function resetPasswordByUserId(userId) {
  userId = String(userId || '').trim();

  if (!userId) {
    setUserManageMessage('ไม่พบ User ID');
    return;
  }

  const user = usersCache.find(function (item) {
    return String(item.userId || '') === userId;
  });

  const confirmed = window.confirm(
    'ยืนยัน Reset Password ผู้ใช้นี้หรือไม่?\n\n' +
    (user ? user.displayName + ' (' + user.username + ')' : userId)
  );

  if (!confirmed) {
    return;
  }

  setButtonLoading(el.resetUserPasswordBtn, true, 'กำลัง Reset...');

  try {
    const data = await window.AnalyticsAPI.resetUserPassword(userId, '');

    setUserManageMessage(
      'Reset Password สำเร็จ: ' +
      (data.username || userId) +
      ' | Password: ' +
      data.temporaryPassword
    );

    window.alert(
      'Reset Password สำเร็จ\n\n' +
      'Username: ' + (data.username || '') + '\n' +
      'Password: ' + data.temporaryPassword + '\n\n' +
      'กรุณาจดรหัสผ่านนี้ไว้ ระบบจะแสดงครั้งเดียว'
    );

    await loadUsers();

  } catch (error) {
    showApiError(setUserManageMessage, error, 'Reset Password ไม่สำเร็จ');
    logApiError('reset_user_password_error', error, {
      userId: userId
    });

  } finally {
    setButtonLoading(el.resetUserPasswordBtn, false, 'Reset Password');
  }
}


async function setUserStatusById(userId, status) {
  userId = String(userId || '').trim();
  status = String(status || '').trim();

  if (!userId || !status) {
    setUserManageMessage('ข้อมูล User ID หรือสถานะไม่ครบ');
    return;
  }

  const user = usersCache.find(function (item) {
    return String(item.userId || '') === userId;
  });

  const confirmed = window.confirm(
    'ยืนยันเปลี่ยนสถานะผู้ใช้นี้เป็น ' + status + ' หรือไม่?\n\n' +
    (user ? user.displayName + ' (' + user.username + ')' : userId)
  );

  if (!confirmed) {
    return;
  }

  try {
    const data = await window.AnalyticsAPI.setUserStatus(userId, status);

    setUserManageMessage(data.message || 'เปลี่ยนสถานะผู้ใช้สำเร็จ');

    await loadUsers();

  } catch (error) {
    showApiError(setUserManageMessage, error, 'เปลี่ยนสถานะผู้ใช้ไม่สำเร็จ');
    logApiError('set_user_status_error', error, {
      userId: userId,
      status: status
    });
  }
}


function resetUserForm() {
  selectedManageUserId = '';

  if (el.userManageUsername) {
    el.userManageUsername.value = '';
    el.userManageUsername.disabled = false;
  }

  if (el.userManageDisplayName) el.userManageDisplayName.value = '';
  if (el.userManageEmail) el.userManageEmail.value = '';
  if (el.userManageRole) el.userManageRole.value = 'Viewer';
  if (el.userManageStatus) el.userManageStatus.value = 'active';
  if (el.userManagePassword) el.userManagePassword.value = '';
  if (el.userManageDashboards) el.userManageDashboards.value = '';

  if (el.userCanExport) el.userCanExport.checked = false;
  if (el.userCanAddSource) el.userCanAddSource.checked = false;
  if (el.userCanEditDashboard) el.userCanEditDashboard.checked = false;
  if (el.userCanViewAuditLog) el.userCanViewAuditLog.checked = false;
  if (el.userCanManageUser) el.userCanManageUser.checked = false;

  setUserManageMessage('');
}


function setUserManageMessage(message) {
  if (el.userManageMessage) {
    el.userManageMessage.textContent = message || '';
  }
}
async function loadAuditLog() {
  if (!el.auditLogResult) {
    return;
  }

  const limit = el.auditLimitSelect ? Number(el.auditLimitSelect.value || 100) : 100;

  el.auditLogResult.classList.remove('empty');
  setPanelLoading(el.auditLogResult, 'กำลังโหลดประวัติการใช้งาน...');

  if (el.reloadAuditLogBtn) {
    setButtonLoading(el.reloadAuditLogBtn, true, 'กำลังโหลด...');
  }

  try {
    const data = await window.AnalyticsAPI.auditLog({
      limit: limit
    });

    renderAuditLog(data.logs || []);

    writeLog({
      step: 'audit_log',
      response: {
        ok: data.ok,
        total: data.total
      }
    });

  } catch (error) {
    const message = getFriendlyErrorMessage(error);

    if (handleAuthErrorIfNeeded(error)) {
      return;
    }

    el.auditLogResult.classList.add('empty');
    el.auditLogResult.textContent = message;

    logApiError('audit_log_error', error, {
      limit: limit
    });

  } finally {
    if (el.reloadAuditLogBtn) {
      setButtonLoading(el.reloadAuditLogBtn, false, 'โหลดประวัติ');
    }
  }
}


function renderAuditLog(logs) {
  if (!el.auditLogResult) {
    return;
  }

  if (!logs || !logs.length) {
    el.auditLogResult.classList.add('empty');
    el.auditLogResult.textContent = 'ยังไม่มีประวัติการใช้งาน';
    return;
  }

  el.auditLogResult.classList.remove('empty');

  const rowsHtml = logs.map(function (log) {
    const time =
      log['วันที่เวลา'] ||
      log['Timestamp'] ||
      log.timestamp ||
      log.createdAt ||
      '-';

    const username =
      log['ชื่อผู้ใช้'] ||
      log.username ||
      '-';

    const role =
      log['สิทธิ์'] ||
      log.role ||
      '-';

    const action =
      log['Action'] ||
      log['การทำงาน'] ||
      log.action ||
      '-';

    const status =
      log['Status'] ||
      log['สถานะ'] ||
      log.status ||
      '-';

    const detail =
      log['รายละเอียด'] ||
      log.detail ||
      '-';

    const sourceId =
      log['รหัสแหล่งข้อมูล'] ||
      log.sourceId ||
      '';

    const dashboardId =
      log['รหัส Dashboard'] ||
      log.dashboardId ||
      '';

    return `
      <tr>
        <td>${escapeHtml(time)}</td>
        <td>
          <strong>${escapeHtml(username)}</strong>
          <small>${escapeHtml(role)}</small>
        </td>
        <td>${renderAuditActionBadge(action)}</td>
        <td>${renderAuditStatusBadge(status)}</td>
        <td>
          <div class="audit-detail">${escapeHtml(detail)}</div>
          ${
            sourceId || dashboardId
              ? `<div class="audit-meta">
                  ${sourceId ? 'Source: ' + escapeHtml(sourceId) : ''}
                  ${sourceId && dashboardId ? ' · ' : ''}
                  ${dashboardId ? 'Dashboard: ' + escapeHtml(dashboardId) : ''}
                </div>`
              : ''
          }
        </td>
      </tr>
    `;
  }).join('');

  el.auditLogResult.innerHTML = `
    <div class="audit-table-wrap">
      <table class="audit-table">
        <thead>
          <tr>
            <th>เวลา</th>
            <th>ผู้ใช้</th>
            <th>Action</th>
            <th>Status</th>
            <th>รายละเอียด</th>
          </tr>
        </thead>
        <tbody>
          ${rowsHtml}
        </tbody>
      </table>
    </div>
  `;
}


function renderAuditActionBadge(action) {
  const text = String(action || '-');
  const lower = text.toLowerCase();

  let cls = 'audit-badge muted';

  if (lower.includes('login')) cls = 'audit-badge info';
  if (lower.includes('error')) cls = 'audit-badge danger';
  if (lower.includes('export')) cls = 'audit-badge warning';
  if (lower.includes('create') || lower.includes('save') || lower.includes('update')) cls = 'audit-badge success';
  if (lower.includes('delete')) cls = 'audit-badge danger';

  return `<span class="${cls}">${escapeHtml(text)}</span>`;
}


function renderAuditStatusBadge(status) {
  const text = String(status || '-');
  const lower = text.toLowerCase();

  let cls = 'audit-badge muted';

  if (lower === 'success' || lower === 'สำเร็จ') cls = 'audit-badge success';
  if (lower === 'fail' || lower === 'failed') cls = 'audit-badge warning';
  if (lower === 'error') cls = 'audit-badge danger';

  return `<span class="${cls}">${escapeHtml(text)}</span>`;
}
  function writeLog(data) {
    if (el.debugLog) {
      el.debugLog.textContent = JSON.stringify(data, null, 2);
    }
  }

  function clearLog() {
    if (el.debugLog) {
      el.debugLog.textContent = 'ยังไม่มีข้อมูล';
    }
  }

    function toBool(value) {
    if (typeof value === 'boolean') {
      return value;
    }

    const text = String(value || '').trim().toLowerCase();

    return ['true', '1', 'yes', 'y', 'on', 'เปิด', 'ใช่'].indexOf(text) >= 0;
  }

  function renderDataTypeOptions(selected) {
    const options = [
      ['text', 'ข้อความ'],
      ['number', 'ตัวเลข'],
      ['currency', 'จำนวนเงิน'],
      ['percent', 'เปอร์เซ็นต์'],
      ['date', 'วันที่'],
      ['datetime', 'วันที่และเวลา'],
      ['time', 'เวลา'],
      ['duration', 'ระยะเวลา'],
      ['category', 'หมวดหมู่'],
      ['status', 'สถานะ'],
      ['person', 'บุคคล'],
      ['location', 'สถานที่'],
      ['boolean', 'จริง/เท็จ'],
      ['url', 'URL'],
      ['image', 'รูปภาพ'],
      ['document', 'เอกสาร'],
      ['id', 'รหัสอ้างอิง'],
      ['remark', 'หมายเหตุ']
    ];

    return options.map(function (opt) {
      return `<option value="${escapeAttr(opt[0])}" ${opt[0] === selected ? 'selected' : ''}>${escapeHtml(opt[1])}</option>`;
    }).join('');
  }

  function renderMeaningTypeOptions(selected) {
    const options = [
      ['text', 'ข้อความทั่วไป'],
      ['date', 'วันที่หลัก/เวลา'],
      ['measure', 'ค่าตัวเลขวิเคราะห์'],
      ['category', 'หมวดหมู่'],
      ['status', 'สถานะ'],
      ['person', 'บุคคล'],
      ['team', 'ทีม/แผนก'],
      ['location', 'สถานที่/DC/สาขา'],
      ['asset', 'ทรัพย์สิน/รถ/อุปกรณ์'],
      ['document', 'เอกสาร'],
      ['image', 'รูปภาพ'],
      ['url', 'URL'],
      ['id', 'รหัสอ้างอิง'],
      ['remark', 'หมายเหตุ']
    ];

    return options.map(function (opt) {
      return `<option value="${escapeAttr(opt[0])}" ${opt[0] === selected ? 'selected' : ''}>${escapeHtml(opt[1])}</option>`;
    }).join('');
  }

  function renderCheck(name, label, checked) {
    return `
      <label class="mapping-check">
        <input data-map-check="${escapeAttr(name)}" type="checkbox" ${checked ? 'checked' : ''}>
        <span>${escapeHtml(label)}</span>
      </label>
    `;
  }

  function shouldDefaultFilter(meaning, dataType) {
    return [
      'date',
      'category',
      'status',
      'person',
      'team',
      'location'
    ].includes(meaning) || ['date', 'datetime', 'status', 'category'].includes(dataType);
  }

  function getFirstSampleValue(sampleRows, colIndex) {
    for (let i = 0; i < sampleRows.length; i++) {
      const row = sampleRows[i] || [];
      const value = row[colIndex];

      if (value !== undefined && value !== null && String(value).trim() !== '') {
        return String(value);
      }
    }

    return '';
  }

  function formatNumber(value) {
    const num = Number(value);

    if (isNaN(num)) {
      return escapeHtml(value);
    }

    return num.toLocaleString('th-TH', {
      maximumFractionDigits: 2
    });
  }

  function escapeHtml(value) {
    return String(value == null ? '' : value)
      .replaceAll('&', '&amp;')
      .replaceAll('<', '&lt;')
      .replaceAll('>', '&gt;')
      .replaceAll('"', '&quot;')
      .replaceAll("'", '&#039;');
  }

  function escapeAttr(value) {
    return escapeHtml(value).replaceAll('`', '&#096;');
  }
  async function handleExportDashboard() {
  const dashboardId = currentDashboardId || (el.dashboardSelect ? el.dashboardSelect.value : '');
  const limit = el.dashboardLimit ? Number(el.dashboardLimit.value || 5000) : 5000;

  if (!dashboardId) {
    setDashboardViewMessage('กรุณาเลือก Dashboard ก่อน Export');
    return;
  }

  if (!window.AnalyticsAPI.dashboardExport) {
    setDashboardViewMessage('ยังไม่พบฟังก์ชัน dashboardExport ใน api.js');
    return;
  }

  const filters = collectDashboardFilters();

  setButtonLoading(el.exportDashboardBtn, true, 'กำลัง Export...');
  setDashboardViewMessage('กำลังเตรียมไฟล์ Export...');
    showGlobalLoading('กำลังเตรียมไฟล์ Export CSV...');

  try {
    const data = await window.AnalyticsAPI.dashboardExport({
      dashboardId: dashboardId,
      limit: limit,
      filters: filters
    });

    downloadTextFile(
      data.filename || 'dashboard_export.csv',
      data.csv || '',
      data.mimeType || 'text/csv;charset=utf-8'
    );

    setDashboardViewMessage(
      'Export สำเร็จ ' +
      Number(data.totalExportRows || data.totalRowsAfterFilter || 0).toLocaleString() +
      ' แถว'
    );

   writeLog({
  step: 'dashboard_export',
  response: {
    ok: data.ok,
    filename: data.filename,
    mimeType: data.mimeType,
    totalExportRows: data.totalExportRows,
    totalRowsAfterFilter: data.totalRowsAfterFilter,
    csv: '[CSV_CONTENT_HIDDEN]'
  }
});

 } catch (error) {
  showApiError(setDashboardViewMessage, error, 'Export CSV ไม่สำเร็จ');
  logApiError('dashboard_export_error', error, {
    dashboardId: dashboardId,
    limit: limit
  });

} finally {
  hideGlobalLoading();
  setButtonLoading(el.exportDashboardBtn, false, 'Export CSV');
}
}

 async function handleExportDashboardExcel() {
  const dashboardId = currentDashboardId || (el.dashboardSelect ? el.dashboardSelect.value : '');
  const limit = el.dashboardLimit ? Number(el.dashboardLimit.value || 5000) : 5000;

  if (!dashboardId) {
    setDashboardViewMessage('กรุณาเลือก Dashboard ก่อน Export Excel');
    return;
  }

  if (!window.AnalyticsAPI || !window.AnalyticsAPI.dashboardExport) {
    setDashboardViewMessage('ยังไม่พบฟังก์ชัน dashboardExport ใน api.js');
    return;
  }

  const filters = collectDashboardFilters();

  setButtonLoading(el.exportDashboardExcelBtn, true, 'กำลัง Export Excel...');
  setDashboardViewMessage('กำลังเตรียมไฟล์ Excel...');
  showGlobalLoading('กำลังสร้างไฟล์ Excel...');

  try {
    const data = await window.AnalyticsAPI.dashboardExport({
      dashboardId: dashboardId,
      limit: limit,
      filters: filters
    });

    if (!data.csv) {
      throw new Error('ไม่พบข้อมูล CSV สำหรับแปลงเป็น Excel');
    }

    const filename = buildExcelFilename(data.filename || 'dashboard-export.csv');

    /**
     * วิธีหลัก: ใช้ SheetJS ถ้าโหลดสำเร็จ
     */
    if (window.XLSX && typeof window.XLSX.read === 'function' && typeof window.XLSX.writeFile === 'function') {
      const workbook = window.XLSX.read(data.csv, {
        type: 'string'
      });

      window.XLSX.writeFile(workbook, filename);
    } else {
      /**
       * วิธีสำรอง: ถ้า SheetJS โหลดไม่ได้ ให้สร้างไฟล์ Excel-compatible .xls
       * เปิดด้วย Microsoft Excel ได้ และรองรับภาษาไทย
       */
      downloadExcelCompatibleHtmlFile(
        filename.replace(/\.xlsx$/i, '.xls'),
        data.csv
      );
    }

    setDashboardViewMessage(
      'Export Excel สำเร็จ: ' +
      filename +
      ' | ' +
      Number(data.totalExportRows || data.totalRowsAfterFilter || 0).toLocaleString() +
      ' แถว'
    );

    writeLog({
      step: 'dashboard_export_excel',
      response: {
        ok: data.ok,
        filename: filename,
        sourceFilename: data.filename,
        totalExportRows: data.totalExportRows,
        totalRowsAfterFilter: data.totalRowsAfterFilter,
        usedSheetJS: !!window.XLSX,
        csv: '[CSV_CONTENT_HIDDEN]'
      }
    });

  } catch (error) {
    showApiError(setDashboardViewMessage, error, 'Export Excel ไม่สำเร็จ');

    logApiError('dashboard_export_excel_error', error, {
      dashboardId: dashboardId,
      limit: limit,
      filters: filters
    });

  } finally {
    hideGlobalLoading();
    setButtonLoading(el.exportDashboardExcelBtn, false, 'Export Excel');
  }
}


function buildExcelFilename(filename) {
  filename = String(filename || 'dashboard-export.csv').trim();

  filename = filename
    .replace(/\.csv$/i, '')
    .replace(/[\\/:*?"<>|]/g, '_');

  if (!filename) {
    filename = 'dashboard-export';
  }

  return filename + '.xlsx';
}

  function downloadExcelCompatibleHtmlFile(filename, csvText) {
  const rows = parseCsvRows(csvText);

  if (!rows.length) {
    throw new Error('ไม่มีข้อมูลสำหรับ Export Excel');
  }

  const htmlRows = rows.map(function (row) {
    const cells = row.map(function (cell) {
      return '<td style="mso-number-format:\\@;">' + escapeHtmlForExcel(cell) + '</td>';
    }).join('');

    return '<tr>' + cells + '</tr>';
  }).join('');

  const html =
    '<html xmlns:o="urn:schemas-microsoft-com:office:office" ' +
    'xmlns:x="urn:schemas-microsoft-com:office:excel" ' +
    'xmlns="http://www.w3.org/TR/REC-html40">' +
    '<head>' +
    '<meta charset="UTF-8">' +
    '<!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet>' +
    '<x:Name>Dashboard Export</x:Name>' +
    '<x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions>' +
    '</x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->' +
    '</head>' +
    '<body>' +
    '<table border="1">' +
    htmlRows +
    '</table>' +
    '</body>' +
    '</html>';

  downloadTextFile(
    filename || 'dashboard-export.xls',
    html,
    'application/vnd.ms-excel;charset=utf-8'
  );
}


function parseCsvRows(csvText) {
  csvText = String(csvText || '').replace(/^\uFEFF/, '');

  const rows = [];
  let row = [];
  let cell = '';
  let insideQuote = false;

  for (let i = 0; i < csvText.length; i++) {
    const char = csvText[i];
    const next = csvText[i + 1];

    if (char === '"' && insideQuote && next === '"') {
      cell += '"';
      i++;
      continue;
    }

    if (char === '"') {
      insideQuote = !insideQuote;
      continue;
    }

    if (char === ',' && !insideQuote) {
      row.push(cell);
      cell = '';
      continue;
    }

    if ((char === '\n' || char === '\r') && !insideQuote) {
      if (char === '\r' && next === '\n') {
        i++;
      }

      row.push(cell);

      if (row.some(function (v) { return String(v || '').trim() !== ''; })) {
        rows.push(row);
      }

      row = [];
      cell = '';
      continue;
    }

    cell += char;
  }

  row.push(cell);

  if (row.some(function (v) { return String(v || '').trim() !== ''; })) {
    rows.push(row);
  }

  return rows;
}


function escapeHtmlForExcel(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
  async function loadManageDashboards() {
    if (!el.manageDashboardSelect) {
      return;
    }

    setManageDashboardMessage('');
    selectedManageDashboard = null;

    el.manageDashboardSelect.innerHTML = '<option value="">กำลังโหลด Dashboard...</option>';
    if (el.manageDashboardSummary) {
  setPanelLoading(el.manageDashboardSummary, 'กำลังโหลดรายการ Dashboard...');
}
    resetManageDashboardForm(false);

    try {
      const data = await window.AnalyticsAPI.listDashboards();
manageDashboardsCache = data.dashboards || [];

renderFilteredManageDashboardOptions();

if (el.manageDashboardSummary) {
  el.manageDashboardSummary.classList.add('empty');
  el.manageDashboardSummary.textContent = manageDashboardsCache.length
    ? 'เลือก Dashboard เพื่อดูรายละเอียดและแก้ไข'
    : 'ยังไม่มี Dashboard ที่สร้างไว้';
}

      writeLog({
        step: 'manage_dashboards_load',
        response: data
      });

    } catch (error) {
      manageDashboardsCache = [];
      el.manageDashboardSelect.innerHTML = '<option value="">โหลด Dashboard ไม่สำเร็จ</option>';
      setManageDashboardMessage(error.message);

      writeLog({
        step: 'manage_dashboards_load_error',
        message: error.message,
        payload: error.payload || null
      });
    }
  }
 function renderFilteredManageDashboardOptions() {
  const dashboards = getFilteredManageDashboards();

  renderManageDashboardOptions(dashboards);
  renderManageDashboardCards(dashboards);
  renderManageDashboardCount(dashboards.length, manageDashboardsCache.length);
}

function getFilteredManageDashboards() {
  const keyword = el.manageDashboardSearch
    ? String(el.manageDashboardSearch.value || '').trim().toLowerCase()
    : '';

  const typeFilter = el.manageDashboardTypeFilter
    ? String(el.manageDashboardTypeFilter.value || '').trim()
    : '';

  const statusFilter = el.manageDashboardStatusFilter
    ? String(el.manageDashboardStatusFilter.value || '').trim()
    : '';

  return (manageDashboardsCache || []).filter(function (dash) {
    const id = String(dash['รหัส Dashboard'] || '').trim();
    const name = String(dash['ชื่อ Dashboard'] || '').trim();
    const type = String(dash['ประเภท Dashboard'] || '').trim();

    const publish = toBool(dash['สถานะ Publish']);
    const allowExport = toBool(dash['อนุญาต Export']);

    const text = [
      id,
      name,
      type,
      String(dash['คำอธิบาย'] || ''),
      String(dash['สิทธิ์ที่มองเห็น'] || '')
    ].join(' ').toLowerCase();

    if (keyword && text.indexOf(keyword) < 0) {
      return false;
    }

    if (typeFilter && type !== typeFilter) {
      return false;
    }

    if (statusFilter === 'published' && !publish) {
      return false;
    }

    if (statusFilter === 'unpublished' && publish) {
      return false;
    }

    if (statusFilter === 'export_on' && !allowExport) {
      return false;
    }

    if (statusFilter === 'export_off' && allowExport) {
      return false;
    }

    return true;
  });
}


function renderManageDashboardCount(filteredCount, totalCount) {
  if (!el.manageDashboardCount) {
    return;
  }

  filteredCount = Number(filteredCount || 0);
  totalCount = Number(totalCount || 0);

  if (filteredCount === totalCount) {
    el.manageDashboardCount.textContent =
      'Dashboard ทั้งหมด ' + totalCount.toLocaleString() + ' รายการ';
    return;
  }

  el.manageDashboardCount.textContent =
    'พบ ' + filteredCount.toLocaleString() +
    ' จากทั้งหมด ' + totalCount.toLocaleString() +
    ' รายการ';
}
  function renderManageDashboardCards(dashboards) {
  if (!el.manageDashboardCardList) {
    return;
  }

  dashboards = dashboards || [];

  if (el.manageDashboardCardCount) {
    el.manageDashboardCardCount.textContent =
      dashboards.length.toLocaleString() + ' รายการ';
  }

  if (!dashboards.length) {
    el.manageDashboardCardList.classList.add('empty');
    el.manageDashboardCardList.textContent = 'ไม่พบ Dashboard ตามเงื่อนไขที่เลือก';
    return;
  }

  el.manageDashboardCardList.classList.remove('empty');

  el.manageDashboardCardList.innerHTML = dashboards.map(function (dash) {
    const dashboardId = String(dash['รหัส Dashboard'] || '').trim();
    const name = String(dash['ชื่อ Dashboard'] || dashboardId).trim();
    const type = String(dash['ประเภท Dashboard'] || '-').trim();
    const description = String(dash['คำอธิบาย'] || '').trim();
    const sourceId = String(dash['รหัสแหล่งข้อมูลหลัก'] || '-').trim();
    const sheetName = String(dash['รหัสชีทย่อยหลัก'] || '-').trim();
    const publish = toBool(dash['สถานะ Publish']);
    const allowExport = toBool(dash['อนุญาต Export']);
    const visibility = String(dash['สิทธิ์ที่มองเห็น'] || '-').trim();
    const updatedAt = String(dash['วันที่แก้ไขล่าสุด'] || '-').trim();
    const updatedBy = String(dash['แก้ไขโดย'] || '-').trim();

    return `
      <article class="dashboard-manage-card" data-dashboard-id="${escapeAttr(dashboardId)}">
        <div class="dashboard-manage-card-head">
          <div>
            <h4>${escapeHtml(name || '-')}</h4>
            <p>${escapeHtml(description || 'ไม่มีคำอธิบาย')}</p>
          </div>

          <span class="dashboard-type-pill">${escapeHtml(type || '-')}</span>
        </div>

        <div class="dashboard-manage-card-meta">
          <div>
            <span>Dashboard ID</span>
            <strong>${escapeHtml(dashboardId || '-')}</strong>
          </div>

          <div>
            <span>Source</span>
            <strong>${escapeHtml(sourceId || '-')}</strong>
          </div>

          <div>
            <span>Sheet</span>
            <strong>${escapeHtml(sheetName || '-')}</strong>
          </div>

          <div>
            <span>Visibility</span>
            <strong>${escapeHtml(visibility || '-')}</strong>
          </div>
        </div>

        <div class="dashboard-manage-card-status">
          <span class="dashboard-status-pill ${publish ? 'is-on' : 'is-off'}">
            ${publish ? 'Published' : 'Unpublished'}
          </span>

          <span class="dashboard-status-pill ${allowExport ? 'is-on' : 'is-off'}">
            ${allowExport ? 'Export On' : 'Export Off'}
          </span>

          <span class="dashboard-status-pill is-muted">
            Update: ${escapeHtml(updatedAt || '-')}
          </span>

          <span class="dashboard-status-pill is-muted">
            By: ${escapeHtml(updatedBy || '-')}
          </span>
        </div>

        <div class="dashboard-manage-card-actions">
          <button class="btn btn-secondary btn-card-select-dashboard" type="button" data-dashboard-id="${escapeAttr(dashboardId)}">
            เลือกแก้ไข
          </button>

          <button class="btn btn-ghost btn-card-open-dashboard" type="button" data-dashboard-id="${escapeAttr(dashboardId)}">
            เปิดดู
          </button>

          <button class="btn btn-ghost btn-card-regenerate-dashboard" type="button" data-dashboard-id="${escapeAttr(dashboardId)}">
            Regenerate
          </button>
        </div>
      </article>
    `;
  }).join('');

  bindManageDashboardCardActions();
}


function bindManageDashboardCardActions() {
  if (!el.manageDashboardCardList) {
    return;
  }

  el.manageDashboardCardList
    .querySelectorAll('.btn-card-select-dashboard')
    .forEach(function (btn) {
      btn.addEventListener('click', function () {
        const dashboardId = btn.getAttribute('data-dashboard-id') || '';
        selectManageDashboardFromCard(dashboardId);
      });
    });

  el.manageDashboardCardList
    .querySelectorAll('.btn-card-open-dashboard')
    .forEach(function (btn) {
      btn.addEventListener('click', function () {
        const dashboardId = btn.getAttribute('data-dashboard-id') || '';
        openDashboardFromManageCard(dashboardId);
      });
    });

  el.manageDashboardCardList
    .querySelectorAll('.btn-card-regenerate-dashboard')
    .forEach(function (btn) {
      btn.addEventListener('click', function () {
        const dashboardId = btn.getAttribute('data-dashboard-id') || '';
        selectManageDashboardFromCard(dashboardId);
        handleRegenerateManageDashboard();
      });
    });
}


function selectManageDashboardFromCard(dashboardId) {
  dashboardId = String(dashboardId || '').trim();

  if (!dashboardId) {
    setManageDashboardMessage('ไม่พบ Dashboard ID');
    return;
  }

  if (el.manageDashboardSelect) {
    el.manageDashboardSelect.value = dashboardId;
  }

  handleSelectManageDashboard();

  const target = document.querySelector('.dashboard-management-section');

  if (target) {
    try {
      target.scrollIntoView({
        behavior: 'smooth',
        block: 'start'
      });
    } catch (err) {
      target.scrollIntoView();
    }
  }
}


async function openDashboardFromManageCard(dashboardId) {
  dashboardId = String(dashboardId || '').trim();

  if (!dashboardId) {
    setManageDashboardMessage('ไม่พบ Dashboard ID');
    return;
  }

  if (el.dashboardSelect) {
    el.dashboardSelect.value = dashboardId;
  }

  currentDashboardId = dashboardId;
  setAppMode('user');

  await handleOpenDashboard();
}

  function renderManageDashboardOptions(dashboards) {
  if (!el.manageDashboardSelect) {
    return;
  }

  const currentValue = el.manageDashboardSelect.value || '';

  if (!dashboards || !dashboards.length) {
    el.manageDashboardSelect.innerHTML = '<option value="">ไม่พบ Dashboard ตามเงื่อนไข</option>';
    selectedManageDashboard = null;

    if (el.manageDashboardSummary) {
      el.manageDashboardSummary.classList.add('empty');
      el.manageDashboardSummary.textContent = 'ไม่พบ Dashboard ตามเงื่อนไขที่เลือก';
    }

    return;
  }

  el.manageDashboardSelect.innerHTML =
    '<option value="">เลือก Dashboard ที่ต้องการจัดการ</option>' +
    dashboards.map(function (dash) {
      const id = String(dash['รหัส Dashboard'] || '').trim();
      const name = String(dash['ชื่อ Dashboard'] || id).trim();
      const type = String(dash['ประเภท Dashboard'] || '').trim();
      const publish = toBool(dash['สถานะ Publish']);
      const allowExport = toBool(dash['อนุญาต Export']);

      const statusText = [
        publish ? 'Published' : 'Unpublished',
        allowExport ? 'Export On' : 'Export Off'
      ].join(' / ');

      return (
        '<option value="' + escapeAttr(id) + '">' +
        escapeHtml(name) +
        (type ? ' (' + escapeHtml(type) + ')' : '') +
        ' - ' + escapeHtml(statusText) +
        '</option>'
      );
    }).join('');

  const stillExists = dashboards.some(function (dash) {
    return String(dash['รหัส Dashboard'] || '').trim() === currentValue;
  });

  if (currentValue && stillExists) {
    el.manageDashboardSelect.value = currentValue;
  }
}


  function handleSelectManageDashboard() {
    const dashboardId = el.manageDashboardSelect ? el.manageDashboardSelect.value : '';

    if (!dashboardId) {
      selectedManageDashboard = null;
      resetManageDashboardForm();
      return;
    }

    selectedManageDashboard = findManageDashboardById(dashboardId);

    if (!selectedManageDashboard) {
      resetManageDashboardForm();
      setManageDashboardMessage('ไม่พบ Dashboard ที่เลือก');
      return;
    }

    fillManageDashboardForm(selectedManageDashboard);
    renderManageDashboardSummary(selectedManageDashboard);
    setManageDashboardMessage('');
  }


  function findManageDashboardById(dashboardId) {
    dashboardId = String(dashboardId || '').trim();

    return (manageDashboardsCache || []).find(function (dash) {
      return String(dash['รหัส Dashboard'] || '').trim() === dashboardId;
    }) || null;
  }


  function fillManageDashboardForm(dash) {
    if (!dash) {
      resetManageDashboardForm();
      return;
    }

    if (el.manageDashboardName) {
      el.manageDashboardName.value = String(dash['ชื่อ Dashboard'] || '');
    }

    if (el.manageDashboardType) {
      el.manageDashboardType.value = String(dash['ประเภท Dashboard'] || 'Custom') || 'Custom';
    }

    if (el.manageDashboardDescription) {
      el.manageDashboardDescription.value = String(dash['คำอธิบาย'] || '');
    }

    if (el.manageDashboardPublish) {
      el.manageDashboardPublish.checked = toBool(dash['สถานะ Publish']);
    }

    if (el.manageDashboardExport) {
      el.manageDashboardExport.checked = toBool(dash['อนุญาต Export']);
    }

    if (el.manageDashboardHidden) {
      el.manageDashboardHidden.checked = !toBool(dash['สถานะ Publish']);
    }

    setManageDashboardRoles(String(dash['สิทธิ์ที่มองเห็น'] || ''));
  }


  function resetManageDashboardForm(clearSelect = true) {
    selectedManageDashboard = null;

    if (clearSelect && el.manageDashboardSelect) {
      el.manageDashboardSelect.value = '';
    }

    if (el.manageDashboardName) {
      el.manageDashboardName.value = '';
    }

    if (el.manageDashboardType) {
      el.manageDashboardType.value = 'Custom';
    }

    if (el.manageDashboardDescription) {
      el.manageDashboardDescription.value = '';
    }

    if (el.manageDashboardPublish) {
      el.manageDashboardPublish.checked = false;
    }

    if (el.manageDashboardExport) {
      el.manageDashboardExport.checked = false;
    }

    if (el.manageDashboardHidden) {
      el.manageDashboardHidden.checked = false;
    }
    if (el.manageDashboardSearch) {
  el.manageDashboardSearch.value = '';
}

if (el.manageDashboardTypeFilter) {
  el.manageDashboardTypeFilter.value = '';
}

if (el.manageDashboardStatusFilter) {
  el.manageDashboardStatusFilter.value = '';
}

if (el.manageDashboardCount) {
  el.manageDashboardCount.textContent = 'Dashboard ทั้งหมด 0 รายการ';
}
    setManageDashboardRoles('');

    if (el.manageDashboardSummary) {
      el.manageDashboardSummary.classList.add('empty');
      el.manageDashboardSummary.textContent = 'เลือก Dashboard เพื่อดูรายละเอียดและแก้ไข';
    }
   if (el.manageDashboardCardList) {
  el.manageDashboardCardList.classList.add('empty');
  el.manageDashboardCardList.textContent = 'กดโหลดรายการ Dashboard เพื่อแสดงรายการ';
}

if (el.manageDashboardCardCount) {
  el.manageDashboardCardCount.textContent = '0 รายการ';
}
    setManageDashboardMessage('');
  }


  function getManageDashboardRoles() {
    return Array.from(document.querySelectorAll('.manage-dashboard-role'))
      .filter(function (input) {
        return input.checked;
      })
      .map(function (input) {
        return input.value;
      });
  }


  function setManageDashboardRoles(value) {
    const roles = String(value || '')
      .split(',')
      .map(function (role) {
        return String(role || '').trim();
      })
      .filter(function (role) {
        return role !== '';
      });

    document.querySelectorAll('.manage-dashboard-role').forEach(function (input) {
      input.checked = roles.indexOf(input.value) >= 0;
    });
  }


  async function handleSaveManageDashboard() {
    const dashboardId = el.manageDashboardSelect ? el.manageDashboardSelect.value : '';

    if (!dashboardId) {
      setManageDashboardMessage('กรุณาเลือก Dashboard ก่อน');
      return;
    }

    const dashboardName = el.manageDashboardName ? el.manageDashboardName.value.trim() : '';
    const dashboardType = el.manageDashboardType ? el.manageDashboardType.value : 'Custom';
    const description = el.manageDashboardDescription ? el.manageDashboardDescription.value.trim() : '';
    const publish = el.manageDashboardPublish ? el.manageDashboardPublish.checked : false;
    const allowExport = el.manageDashboardExport ? el.manageDashboardExport.checked : false;
    const visibility = getManageDashboardRoles();

    if (!dashboardName) {
      setManageDashboardMessage('กรุณากรอกชื่อ Dashboard');
      return;
    }

    if (!visibility.length) {
      setManageDashboardMessage('กรุณาเลือกสิทธิ์ที่มองเห็นอย่างน้อย 1 สิทธิ์');
      return;
    }

    setButtonLoading(el.saveManageDashboardBtn, true, 'กำลังบันทึก...');
    setManageDashboardMessage('');

    try {
      const data = await window.AnalyticsAPI.updateDashboard({
        dashboardId: dashboardId,
        dashboardName: dashboardName,
        dashboardType: dashboardType,
        description: description,
        publish: publish,
        allowExport: allowExport,
        visibility: visibility
      });

      setManageDashboardMessage(data.message || 'บันทึกการแก้ไข Dashboard สำเร็จ');

      writeLog({
        step: 'manage_dashboard_save',
        response: data
      });

      await loadDashboardOptions();
      await loadManageDashboards();

      if (el.manageDashboardSelect) {
        el.manageDashboardSelect.value = dashboardId;
        handleSelectManageDashboard();
      }

    } catch (error) {
      setManageDashboardMessage(error.message);

      writeLog({
        step: 'manage_dashboard_save_error',
        message: error.message,
        payload: error.payload || null
      });

    } finally {
      setButtonLoading(el.saveManageDashboardBtn, false, 'บันทึกการแก้ไข');
    }
  }


  async function handleToggleManageDashboardHidden() {
    const dashboardId = el.manageDashboardSelect ? el.manageDashboardSelect.value : '';

    if (!dashboardId) {
      setManageDashboardMessage('กรุณาเลือก Dashboard ก่อน');
      return;
    }

    const hide = el.manageDashboardHidden ? el.manageDashboardHidden.checked : false;

    setButtonLoading(el.hideManageDashboardBtn, true, 'กำลังอัปเดต...');

    try {
      const data = await window.AnalyticsAPI.setDashboardHidden(dashboardId, hide);

      setManageDashboardMessage(
        hide
          ? 'ซ่อน Dashboard สำเร็จ'
          : 'เปิด Dashboard กลับมาแล้ว'
      );

      writeLog({
        step: 'manage_dashboard_toggle_hidden',
        response: data
      });

      await loadDashboardOptions();
      await loadManageDashboards();

      if (el.manageDashboardSelect) {
        el.manageDashboardSelect.value = dashboardId;
        handleSelectManageDashboard();
      }

    } catch (error) {
      setManageDashboardMessage(error.message);

      writeLog({
        step: 'manage_dashboard_toggle_hidden_error',
        message: error.message,
        payload: error.payload || null
      });

    } finally {
      setButtonLoading(el.hideManageDashboardBtn, false, 'ซ่อน / เปิดกลับ');
    }
  }

   async function handleRegenerateManageDashboard() {
  const dashboardId = el.manageDashboardSelect ? el.manageDashboardSelect.value : '';

  if (!dashboardId) {
    setManageDashboardMessage('กรุณาเลือก Dashboard ก่อน Regenerate');
    return;
  }

  const dash = findManageDashboardById(dashboardId);
  const dashboardName = dash ? String(dash['ชื่อ Dashboard'] || dashboardId) : dashboardId;

  const confirmed = window.confirm(
    'ยืนยัน Regenerate Dashboard จาก Mapping ล่าสุดหรือไม่?\n\n' +
    dashboardName +
    '\n\nระบบจะสร้าง KPI / กราฟ / ตัวกรอง / Layout ใหม่จาก Mapping ล่าสุด แต่จะคงชื่อ Dashboard, Publish, Export และสิทธิ์เดิมไว้'
  );

  if (!confirmed) {
    return;
  }

  if (!window.AnalyticsAPI.regenerateDashboard) {
    setManageDashboardMessage('ยังไม่พบฟังก์ชัน regenerateDashboard ใน api.js');
    return;
  }

  setButtonLoading(el.regenerateManageDashboardBtn, true, 'กำลัง Regenerate...');
  setManageDashboardMessage('กำลัง Regenerate Dashboard จาก Mapping ล่าสุด...');
     showGlobalLoading('กำลัง Regenerate Dashboard จาก Mapping ล่าสุด...');

  try {
    const data = await window.AnalyticsAPI.regenerateDashboard(dashboardId);

    setManageDashboardMessage(
      (data.message || 'Regenerate สำเร็จ') +
      ' | KPI ' + Number(data.totalMetrics || 0).toLocaleString() +
      ' | กราฟ ' + Number(data.totalCharts || 0).toLocaleString() +
      ' | ตัวกรอง ' + Number(data.totalFilters || 0).toLocaleString()
    );

    writeLog({
      step: 'manage_dashboard_regenerate',
      response: data
    });

    await loadDashboardOptions();
    await loadManageDashboards();

    if (el.manageDashboardSelect) {
      el.manageDashboardSelect.value = dashboardId;
      handleSelectManageDashboard();
    }

    if (currentDashboardId === dashboardId) {
      await handleRefreshDashboard();
    }

  } catch (error) {
  showApiError(setManageDashboardMessage, error, 'Regenerate Dashboard ไม่สำเร็จ');
  logApiError('manage_dashboard_regenerate_error', error, {
    dashboardId: dashboardId
  });

} finally {
  hideGlobalLoading();
  setButtonLoading(el.regenerateManageDashboardBtn, false, 'Regenerate จาก Mapping ล่าสุด');
}
}
  async function handleDeleteManageDashboard() {
    const dashboardId = el.manageDashboardSelect ? el.manageDashboardSelect.value : '';

    if (!dashboardId) {
      setManageDashboardMessage('กรุณาเลือก Dashboard ก่อน');
      return;
    }

    const dash = findManageDashboardById(dashboardId);
    const dashboardName = dash ? String(dash['ชื่อ Dashboard'] || dashboardId) : dashboardId;

    const confirmed = window.confirm(
      'ยืนยัน Soft Delete Dashboard นี้หรือไม่?\n\n' +
      dashboardName +
      '\n\nระบบจะไม่ลบแถวจริง แต่จะซ่อนและปิด Export'
    );

    if (!confirmed) {
      return;
    }

    setButtonLoading(el.deleteManageDashboardBtn, true, 'กำลังลบ...');

    try {
      const data = await window.AnalyticsAPI.deleteDashboard(dashboardId);

      setManageDashboardMessage(data.message || 'Soft Delete Dashboard สำเร็จ');

      writeLog({
        step: 'manage_dashboard_soft_delete',
        response: data
      });

      await loadDashboardOptions();
      await loadManageDashboards();

    } catch (error) {
      setManageDashboardMessage(error.message);

      writeLog({
        step: 'manage_dashboard_soft_delete_error',
        message: error.message,
        payload: error.payload || null
      });

    } finally {
      setButtonLoading(el.deleteManageDashboardBtn, false, 'Soft Delete');
    }
  }


  function renderManageDashboardSummary(dash) {
    if (!el.manageDashboardSummary || !dash) {
      return;
    }

    const dashboardId = String(dash['รหัส Dashboard'] || '');
    const name = String(dash['ชื่อ Dashboard'] || '');
    const type = String(dash['ประเภท Dashboard'] || '');
    const sourceId = String(dash['รหัสแหล่งข้อมูลหลัก'] || '');
    const sheetName = String(dash['รหัสชีทย่อยหลัก'] || '');
    const publish = toBool(dash['สถานะ Publish']);
    const allowExport = toBool(dash['อนุญาต Export']);
    const visibility = String(dash['สิทธิ์ที่มองเห็น'] || '-');
    const updatedAt = String(dash['วันที่แก้ไขล่าสุด'] || '-');
    const updatedBy = String(dash['แก้ไขโดย'] || '-');

    el.manageDashboardSummary.classList.remove('empty');

    el.manageDashboardSummary.innerHTML = `
      <div class="dashboard-manage-summary-grid">
        <div class="dashboard-manage-summary-item">
          <span>Dashboard ID</span>
          <strong>${escapeHtml(dashboardId || '-')}</strong>
        </div>

        <div class="dashboard-manage-summary-item">
          <span>ชื่อ Dashboard</span>
          <strong>${escapeHtml(name || '-')}</strong>
        </div>

        <div class="dashboard-manage-summary-item">
          <span>ประเภท</span>
          <strong>${escapeHtml(type || '-')}</strong>
        </div>

        <div class="dashboard-manage-summary-item">
          <span>แหล่งข้อมูล</span>
          <strong>${escapeHtml(sourceId || '-')}</strong>
        </div>

        <div class="dashboard-manage-summary-item">
          <span>ชีทหลัก</span>
          <strong>${escapeHtml(sheetName || '-')}</strong>
        </div>

        <div class="dashboard-manage-summary-item">
          <span>Publish</span>
          <strong>
            <span class="dashboard-status-pill ${publish ? 'is-on' : 'is-off'}">
              ${publish ? 'เปิด' : 'ปิด'}
            </span>
          </strong>
        </div>

        <div class="dashboard-manage-summary-item">
          <span>Export</span>
          <strong>
            <span class="dashboard-status-pill ${allowExport ? 'is-on' : 'is-off'}">
              ${allowExport ? 'อนุญาต' : 'ไม่อนุญาต'}
            </span>
          </strong>
        </div>

        <div class="dashboard-manage-summary-item">
          <span>สิทธิ์ที่มองเห็น</span>
          <strong>${escapeHtml(visibility)}</strong>
        </div>

        <div class="dashboard-manage-summary-item">
          <span>แก้ไขล่าสุด</span>
          <strong>${escapeHtml(updatedAt)}</strong>
        </div>

        <div class="dashboard-manage-summary-item">
          <span>แก้ไขโดย</span>
          <strong>${escapeHtml(updatedBy)}</strong>
        </div>
      </div>
    `;
  }


  function setManageDashboardMessage(message) {
    if (el.manageDashboardMessage) {
      el.manageDashboardMessage.textContent = message || '';
    }
  }
function downloadTextFile(filename, content, mimeType) {
  const blob = new Blob([content], {
    type: mimeType || 'text/plain;charset=utf-8'
  });

  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');

  a.href = url;
  a.download = filename || 'dashboard_export.csv';

  document.body.appendChild(a);
  a.click();

  setTimeout(function () {
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, 0);
}
  function applyRoleUi(user) {
  const role = String(user && user.role ? user.role : '').trim();
  const isAdminRole = ['Super Admin', 'Admin', 'Editor'].includes(role);

  document.body.classList.remove(
    'role-super-admin',
    'role-admin',
    'role-editor',
    'role-user',
    'role-viewer',
    'role-viewer-only'
  );

  if (role === 'Super Admin') {
    document.body.classList.add('role-super-admin');
  } else if (role === 'Admin') {
    document.body.classList.add('role-admin');
  } else if (role === 'Editor') {
    document.body.classList.add('role-editor');
  } else if (role === 'User') {
    document.body.classList.add('role-user');
  } else if (role === 'Viewer') {
    document.body.classList.add('role-viewer');
  }

  if (el.modeSwitcher) {
    el.modeSwitcher.classList.remove('hidden');
  }

  if (!isAdminRole) {
    document.body.classList.add('role-viewer-only');

    if (el.adminModeBtn) {
      el.adminModeBtn.classList.add('hidden');
    }

    if (el.userModeBtn) {
      el.userModeBtn.classList.remove('hidden');
    }

    setAppMode('user');
    return;
  }

  document.body.classList.remove('role-viewer-only');

  if (el.adminModeBtn) {
    el.adminModeBtn.classList.remove('hidden');
  }

  if (el.userModeBtn) {
    el.userModeBtn.classList.remove('hidden');
  }

  setAppMode('admin');
}


function setAppMode(mode) {
  const nextMode = mode === 'user' ? 'user' : 'admin';
  const role = String(currentUser && currentUser.role ? currentUser.role : '').trim();
  const isAdminRole = ['Super Admin', 'Admin', 'Editor'].includes(role);

  if (!isAdminRole && nextMode === 'admin') {
    currentMode = 'user';
  } else {
    currentMode = nextMode;
  }

  document.body.classList.remove('mode-admin', 'mode-user');

  if (currentMode === 'user') {
    document.body.classList.add('mode-user');
  } else {
    document.body.classList.add('mode-admin');
  }

  if (el.adminModeBtn) {
    el.adminModeBtn.classList.toggle('active', currentMode === 'admin');
  }

  if (el.userModeBtn) {
    el.userModeBtn.classList.toggle('active', currentMode === 'user');
  }

  if (currentMode === 'user') {
    setSystemStatus('โหมด User Dashboard');

    setTimeout(function () {
      scrollDashboardViewerIntoView();
    }, 80);
  } else {
    setSystemStatus('โหมด Admin Console');
  }
}
  function scrollDashboardViewerIntoView() {
  const target = document.querySelector('.user-dashboard-section');

  if (!target) {
    return;
  }

  try {
    target.scrollIntoView({
      behavior: 'smooth',
      block: 'start'
    });
  } catch (error) {
    target.scrollIntoView();
  }
}
  function mountQueuedCharts() {
    const queue = chartRenderQueue.slice();
    chartRenderQueue = [];

    if (!queue.length) {
      return;
    }

    if (!window.echarts) {
      mountFallbackCharts(queue);
      return;
    }

    requestAnimationFrame(function () {
      queue.forEach(function (item) {
        const container = document.getElementById(item.id);

        if (!container) {
          return;
        }

        renderEChartByType(container, item.chart);
      });

      window.removeEventListener('resize', resizeDashboardCharts);
      window.addEventListener('resize', resizeDashboardCharts, { passive: true });
    });
  }


  function disposeDashboardCharts() {
    chartInstances.forEach(function (chart) {
      try {
        chart.dispose();
      } catch (error) {
        // ignore
      }
    });

    chartInstances = [];
    chartRenderQueue = [];
  }


  function resizeDashboardCharts() {
    chartInstances.forEach(function (chart) {
      try {
        chart.resize();
      } catch (error) {
        // ignore
      }
    });
  }


  function mountFallbackCharts(queue) {
    (queue || chartRenderQueue || []).forEach(function (item) {
      const dom = document.getElementById(item.id);

      if (!dom) {
        return;
      }

      dom.outerHTML = renderSimpleBarChart((item.chart && item.chart.data) || []);
    });
  }


  function formatClientDateTime(date) {
  const d = date instanceof Date ? date : new Date();

  const pad = function (n) {
    return String(n).padStart(2, '0');
  };

  const day = pad(d.getDate());
  const month = pad(d.getMonth() + 1);
  const year = d.getFullYear();
  const hour = pad(d.getHours());
  const minute = pad(d.getMinutes());
  const second = pad(d.getSeconds());

  return day + '/' + month + '/' + year + ' ' + hour + ':' + minute + ':' + second;
}
  function renderEChartByType(container, chart) {
  if (!container || !chart) {
    return;
  }

  if (!window.echarts) {
    container.innerHTML = '<div class="empty">ยังไม่พบ ECharts library</div>';
    return;
  }

  const type = String(chart.type || chart.chartType || '').trim().toLowerCase();
  const title = String(chart.title || chart.name || chart.chartName || 'Chart');
  const data = Array.isArray(chart.data) ? chart.data : [];

  container.innerHTML = '';

  const chartInstance = echarts.init(container);
  chartInstances.push(chartInstance);

  let option;

  if (type === 'gauge') {
    option = buildGaugeOption(title, data, chart);
  } else if (type === 'area') {
    option = buildAreaOption(title, data, chart);
  } else if (type === 'stacked_bar' || type === 'stackedbar') {
    option = buildStackedBarOption(title, data, chart);
  } else if (type === 'radar') {
    option = buildRadarOption(title, data, chart);
  } else if (type === 'heatmap') {
    option = buildHeatmapOption(title, data, chart);
  } else if (type === 'pie') {
    option = buildPieOption(title, data, chart, false);
  } else if (type === 'donut' || type === 'doughnut') {
    option = buildPieOption(title, data, chart, true);
  } else if (type === 'horizontal_bar' || type === 'horizontalbar') {
    option = buildHorizontalBarOption(title, data, chart);
  } else if (type === 'line') {
    option = buildLineOption(title, data, chart);
  } else {
    option = buildBarOption(title, data, chart);
  }

  chartInstance.setOption(option);

  window.addEventListener('resize', function () {
    chartInstance.resize();
  });
}
  function buildBaseTooltip() {
  return {
    trigger: 'item',
    confine: true
  };
}


function normalizeChartData(data) {
  return (data || []).map(function (item) {
    if (Array.isArray(item)) {
      return {
        name: String(item[0] || ''),
        value: Number(item[1] || 0)
      };
    }

    return {
      name: String(item.name || item.label || item.category || ''),
      value: Number(item.value || item.count || item.total || 0),
      group: String(item.group || item.series || ''),
      x: String(item.x || item.name || item.label || ''),
      y: String(item.y || item.value || '')
    };
  });
}


function buildBarOption(title, data, chart) {
  const rows = normalizeChartData(data);
  const names = rows.map(function (x) { return x.name; });
  const values = rows.map(function (x) { return x.value; });

  return {
    title: {
      text: title,
      left: 'center',
      textStyle: {
        fontSize: 14,
        fontWeight: 700
      }
    },
    tooltip: {
      trigger: 'axis',
      confine: true
    },
    grid: {
      left: 42,
      right: 20,
      top: 56,
      bottom: 56
    },
    xAxis: {
      type: 'category',
      data: names,
      axisLabel: {
        rotate: names.length > 6 ? 35 : 0
      }
    },
    yAxis: {
      type: 'value'
    },
    series: [
      {
        type: 'bar',
        data: values,
        barMaxWidth: 46
      }
    ]
  };
}


function buildHorizontalBarOption(title, data, chart) {
  const rows = normalizeChartData(data).slice().reverse();
  const names = rows.map(function (x) { return x.name; });
  const values = rows.map(function (x) { return x.value; });

  return {
    title: {
      text: title,
      left: 'center',
      textStyle: {
        fontSize: 14,
        fontWeight: 700
      }
    },
    tooltip: {
      trigger: 'axis',
      confine: true
    },
    grid: {
      left: 110,
      right: 22,
      top: 56,
      bottom: 24
    },
    xAxis: {
      type: 'value'
    },
    yAxis: {
      type: 'category',
      data: names
    },
    series: [
      {
        type: 'bar',
        data: values,
        barMaxWidth: 34
      }
    ]
  };
}


function buildLineOption(title, data, chart) {
  const rows = normalizeChartData(data);
  const names = rows.map(function (x) { return x.name; });
  const values = rows.map(function (x) { return x.value; });

  return {
    title: {
      text: title,
      left: 'center',
      textStyle: {
        fontSize: 14,
        fontWeight: 700
      }
    },
    tooltip: {
      trigger: 'axis',
      confine: true
    },
    grid: {
      left: 42,
      right: 22,
      top: 56,
      bottom: 48
    },
    xAxis: {
      type: 'category',
      data: names
    },
    yAxis: {
      type: 'value'
    },
    series: [
      {
        type: 'line',
        data: values,
        smooth: true
      }
    ]
  };
}


function buildAreaOption(title, data, chart) {
  const rows = normalizeChartData(data);
  const names = rows.map(function (x) { return x.name; });
  const values = rows.map(function (x) { return x.value; });

  return {
    title: {
      text: title,
      left: 'center',
      textStyle: {
        fontSize: 14,
        fontWeight: 700
      }
    },
    tooltip: {
      trigger: 'axis',
      confine: true
    },
    grid: {
      left: 42,
      right: 22,
      top: 56,
      bottom: 48
    },
    xAxis: {
      type: 'category',
      boundaryGap: false,
      data: names
    },
    yAxis: {
      type: 'value'
    },
    series: [
      {
        type: 'line',
        data: values,
        smooth: true,
        areaStyle: {}
      }
    ]
  };
}


function buildPieOption(title, data, chart, donut) {
  const rows = normalizeChartData(data);

  return {
    title: {
      text: title,
      left: 'center',
      textStyle: {
        fontSize: 14,
        fontWeight: 700
      }
    },
    tooltip: buildBaseTooltip(),
    legend: {
      bottom: 0,
      type: 'scroll'
    },
    series: [
      {
        type: 'pie',
        radius: donut ? ['42%', '68%'] : '68%',
        center: ['50%', '48%'],
        data: rows.map(function (x) {
          return {
            name: x.name,
            value: x.value
          };
        }),
        label: {
          formatter: '{b}: {d}%'
        }
      }
    ]
  };
}


function buildGaugeOption(title, data, chart) {
  const rows = normalizeChartData(data);
  const first = rows[0] || {};
  const value = Number(first.value || chart.value || 0);
  const max = Number(chart.max || chart.target || 100) || 100;

  return {
    title: {
      text: title,
      left: 'center',
      textStyle: {
        fontSize: 14,
        fontWeight: 700
      }
    },
    tooltip: buildBaseTooltip(),
    series: [
      {
        type: 'gauge',
        min: 0,
        max: max,
        progress: {
          show: true
        },
        detail: {
          valueAnimation: true,
          formatter: function (v) {
            return Number(v || 0).toLocaleString();
          }
        },
        data: [
          {
            value: value,
            name: first.name || 'Value'
          }
        ]
      }
    ]
  };
}


function buildStackedBarOption(title, data, chart) {
  const rows = normalizeChartData(data);
  const categories = [];
  const groups = [];

  rows.forEach(function (item) {
    if (categories.indexOf(item.x || item.name) < 0) {
      categories.push(item.x || item.name);
    }

    if (groups.indexOf(item.group || 'รายการ') < 0) {
      groups.push(item.group || 'รายการ');
    }
  });

  const series = groups.map(function (group) {
    return {
      name: group,
      type: 'bar',
      stack: 'total',
      data: categories.map(function (cat) {
        const found = rows.find(function (item) {
          return (item.x || item.name) === cat && (item.group || 'รายการ') === group;
        });

        return found ? found.value : 0;
      })
    };
  });

  return {
    title: {
      text: title,
      left: 'center',
      textStyle: {
        fontSize: 14,
        fontWeight: 700
      }
    },
    tooltip: {
      trigger: 'axis',
      axisPointer: {
        type: 'shadow'
      },
      confine: true
    },
    legend: {
      bottom: 0,
      type: 'scroll'
    },
    grid: {
      left: 42,
      right: 22,
      top: 56,
      bottom: 70
    },
    xAxis: {
      type: 'category',
      data: categories
    },
    yAxis: {
      type: 'value'
    },
    series: series
  };
}


function buildRadarOption(title, data, chart) {
  const rows = normalizeChartData(data).slice(0, 8);
  const maxValue = Math.max.apply(null, rows.map(function (x) {
    return x.value;
  }).concat([100]));

  return {
    title: {
      text: title,
      left: 'center',
      textStyle: {
        fontSize: 14,
        fontWeight: 700
      }
    },
    tooltip: buildBaseTooltip(),
    radar: {
      radius: '62%',
      indicator: rows.map(function (x) {
        return {
          name: x.name,
          max: maxValue
        };
      })
    },
    series: [
      {
        type: 'radar',
        data: [
          {
            value: rows.map(function (x) { return x.value; }),
            name: title
          }
        ]
      }
    ]
  };
}


function buildHeatmapOption(title, data, chart) {
  const rows = normalizeChartData(data);

  const xLabels = [];
  const yLabels = [];

  rows.forEach(function (item) {
    const x = item.x || item.name || '';
    const y = item.group || 'รายการ';

    if (xLabels.indexOf(x) < 0) xLabels.push(x);
    if (yLabels.indexOf(y) < 0) yLabels.push(y);
  });

  const heatData = rows.map(function (item) {
    const x = item.x || item.name || '';
    const y = item.group || 'รายการ';

    return [
      xLabels.indexOf(x),
      yLabels.indexOf(y),
      item.value
    ];
  });

  return {
    title: {
      text: title,
      left: 'center',
      textStyle: {
        fontSize: 14,
        fontWeight: 700
      }
    },
    tooltip: {
      position: 'top',
      confine: true
    },
    grid: {
      left: 90,
      right: 24,
      top: 56,
      bottom: 60
    },
    xAxis: {
      type: 'category',
      data: xLabels,
      splitArea: {
        show: true
      }
    },
    yAxis: {
      type: 'category',
      data: yLabels,
      splitArea: {
        show: true
      }
    },
    visualMap: {
      min: 0,
      max: Math.max.apply(null, rows.map(function (x) { return x.value; }).concat([1])),
      calculable: true,
      orient: 'horizontal',
      left: 'center',
      bottom: 6
    },
    series: [
      {
        type: 'heatmap',
        data: heatData,
        label: {
          show: true
        }
      }
    ]
  };
}
})();
