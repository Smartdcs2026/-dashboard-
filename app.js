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
    runSystemCheckBtn: document.getElementById('runSystemCheckBtn'),
systemCheckResult: document.getElementById('systemCheckResult'),
    
    debugLog: document.getElementById('debugLog'),
        /**
     * Dashboard Designer / Builder
     */
    openDesignerBtn: document.getElementById('openDesignerBtn'),
    dashboardDesignerSection: document.getElementById('dashboardDesignerSection'),

    designerReloadBtn: document.getElementById('designerReloadBtn'),
    designerResetBtn: document.getElementById('designerResetBtn'),

    designerDashboardSelect: document.getElementById('designerDashboardSelect'),
    designerDashboardType: document.getElementById('designerDashboardType'),
    designerSourceSelect: document.getElementById('designerSourceSelect'),
    designerSheetSelect: document.getElementById('designerSheetSelect'),
    designerSampleLimit: document.getElementById('designerSampleLimit'),

    designerAnalyzeBtn: document.getElementById('designerAnalyzeBtn'),
    designerSuggestBtn: document.getElementById('designerSuggestBtn'),
    designerMessage: document.getElementById('designerMessage'),
    designerDetectedSummary: document.getElementById('designerDetectedSummary'),

    designerSaveFieldAnalysisBtn: document.getElementById('designerSaveFieldAnalysisBtn'),
    designerFieldList: document.getElementById('designerFieldList'),

    designerWidgetGallery: document.getElementById('designerWidgetGallery'),
    designerSuggestedWidgets: document.getElementById('designerSuggestedWidgets'),

    designerWidgetType: document.getElementById('designerWidgetType'),
    designerWidgetTitle: document.getElementById('designerWidgetTitle'),
    designerDateField: document.getElementById('designerDateField'),
    designerXField: document.getElementById('designerXField'),
    designerValueField: document.getElementById('designerValueField'),
    designerStackField: document.getElementById('designerStackField'),
    designerAggregate: document.getElementById('designerAggregate'),
    designerUnit: document.getElementById('designerUnit'),
    designerSortMethod: document.getElementById('designerSortMethod'),
    designerLimit: document.getElementById('designerLimit'),
    designerTheme: document.getElementById('designerTheme'),
    designerTargetValue: document.getElementById('designerTargetValue'),

    designerShowLegend: document.getElementById('designerShowLegend'),
    designerShowLabel: document.getElementById('designerShowLabel'),
    designerShowPercent: document.getElementById('designerShowPercent'),
    designerEnableExport: document.getElementById('designerEnableExport'),

    designerCompareEnabled: document.getElementById('designerCompareEnabled'),
    designerCompareType: document.getElementById('designerCompareType'),
    designerCompareCategoryField: document.getElementById('designerCompareCategoryField'),
    designerCompareCategoryA: document.getElementById('designerCompareCategoryA'),
    designerCompareCategoryB: document.getElementById('designerCompareCategoryB'),
    designerCompareTargetValue: document.getElementById('designerCompareTargetValue'),
    designerCompareAverageWindow: document.getElementById('designerCompareAverageWindow'),
    designerComparisonPreviewBtn: document.getElementById('designerComparisonPreviewBtn'),
    designerComparisonResult: document.getElementById('designerComparisonResult'),

    designerMiniPreviewBtn: document.getElementById('designerMiniPreviewBtn'),
    designerRealPreviewBtn: document.getElementById('designerRealPreviewBtn'),
    designerLivePreview: document.getElementById('designerLivePreview'),

    designerSaveWidgetBtn: document.getElementById('designerSaveWidgetBtn'),
    designerUpdateWidgetBtn: document.getElementById('designerUpdateWidgetBtn'),

    designerDesktopW: document.getElementById('designerDesktopW'),
    designerDesktopH: document.getElementById('designerDesktopH'),
    designerTabletOrder: document.getElementById('designerTabletOrder'),
    designerMobileOrder: document.getElementById('designerMobileOrder'),
    designerMobileHidden: document.getElementById('designerMobileHidden'),

    designerLoadSavedWidgetsBtn: document.getElementById('designerLoadSavedWidgetsBtn'),
    designerSavedWidgetList: document.getElementById('designerSavedWidgetList')
  };

  let currentUser = null;
  let sourcesCache = [];
  let selectedSourceId = '';
  let selectedSheetName = '';
  let lastHeadersData = null;
 let currentDashboardId = '';
let currentDashboardFilters = [];
let dashboardsCache = [];
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
 

 
  /**
   * Dashboard Designer State
   */
  let designerThemesCache = [];
  let designerWidgetTemplatesCache = [];
    let designerCurrentPreviewData = null;
  let designerSavedWidgetsCache = [];
  let designerSavedComparisonsCache = [];
  let designerSavedLayoutsCache = [];
  let designerFieldsCache = [];
  let designerSuggestedWidgetsCache = [];
  let designerSelectedSourceId = '';
  let designerSelectedSheetName = '';
  let designerSelectedDashboardId = '';
  let designerSelectedWidgetType = '';
  let designerEditingWidgetId = '';
  let currentBuilderTheme = null;
let currentBuilderPalette = [];
  
const DASHBOARD_TABLE_PAGE_SIZE = 50;

 window.DASHBOARD_APP_READY = false;

document.addEventListener('DOMContentLoaded', init);

async function init() {
  window.DASHBOARD_APP_READY = true;

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
  el.openDashboardBtn.addEventListener('click', handleOpenDashboardSmart);
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
  el.exportDashboardBtn.addEventListener('click', handleExportDashboardSmart);
}
    if (el.exportDashboardExcelBtn) {
  el.exportDashboardExcelBtn.addEventListener('click', handleExportDashboardExcel);
}
 if (el.refreshDashboardBtn) {
  el.refreshDashboardBtn.addEventListener('click', handleRefreshDashboardSmart);
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
 if (el.runSystemCheckBtn) {
  el.runSystemCheckBtn.addEventListener('click', runSystemCheck);
}
    if (el.dashboardViewResult) {
      el.dashboardViewResult.addEventListener('click', function (event) {
        const btn = event.target.closest('[data-dashboard-table-page-action]');

        if (!btn || btn.disabled) {
          return;
        }

        const action = btn.getAttribute('data-dashboard-table-page-action');

        if (action === 'prev') {
          changeDashboardTablePage(-1);
        }

        if (action === 'next') {
          changeDashboardTablePage(1);
        }
      });
    }
  bindDashboardDesignerEvents();
}

  function resetWorkingState() {
    currentUser = null;
    sourcesCache = [];
    selectedSourceId = '';
    selectedSheetName = '';
    lastHeadersData = null;
    currentDashboardId = '';
    currentDashboardFilters = [];
    currentDashboardTable = null;
    currentDashboardTablePage = 1;
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

    dashboardsCache = [];
manageDashboardsCache = [];

if (el.dashboardSelect) {
  el.dashboardSelect.innerHTML = '<option value="">ยังไม่มีรายการ Dashboard</option>';
}

if (el.designerDashboardSelect) {
  el.designerDashboardSelect.innerHTML = '<option value="">ยังไม่มี Dashboard ให้เลือก</option>';
}

if (el.manageDashboardSelect) {
  el.manageDashboardSelect.innerHTML = '<option value="">ยังไม่มี Dashboard ให้เลือก</option>';
}

if (el.manageDashboardCardList) {
  el.manageDashboardCardList.classList.add('empty');
  el.manageDashboardCardList.textContent = 'ยังไม่มี Dashboard ให้จัดการ';
}

if (el.manageDashboardCount) {
  el.manageDashboardCount.textContent = '0 รายการ';
}

if (el.manageDashboardCardCount) {
  el.manageDashboardCardCount.textContent = '0 รายการ';
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
    if (!window.AnalyticsAPI || typeof window.AnalyticsAPI.health !== 'function') {
      throw new Error('ไม่พบ window.AnalyticsAPI.health กรุณาตรวจสอบ api.js');
    }

    const health = await window.AnalyticsAPI.health();

    let setup = {
      ok: false,
      message: 'ยังไม่ได้ตรวจ Setup'
    };

    if (typeof window.AnalyticsAPI.setupStatus === 'function') {
      setup = await window.AnalyticsAPI.setupStatus();
    }

    setApiStatus(
      'เชื่อมต่อ API สำเร็จ: ' + (health.message || 'พร้อมใช้งาน'),
      'success'
    );

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
    const message = error.message || 'ตรวจสอบ API ไม่สำเร็จ';

    setApiStatus(message, 'error');

    if (el.healthResult) {
      el.healthResult.textContent = message;
    }

    setSystemStatus('API มีปัญหา');

    writeLog({
      step: 'health_error',
      message: message,
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
        await initDashboardDesigner();
      await loadManageDashboards();
      if (currentUser && ['Super Admin', 'Admin'].includes(String(currentUser.role || ''))) {
  await loadUsers();
}
      if (currentUser && ['Super Admin', 'Admin'].includes(String(currentUser.role || ''))) {
  await loadAuditLog();
}
      if (currentUser && ['Super Admin', 'Admin'].includes(String(currentUser.role || ''))) {
  runSystemCheck();
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
      await initDashboardDesigner();
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

  data = data || {};

  disposeDashboardCharts();
  chartRenderQueue = [];

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
  const limitWarning = String(data.limitWarning || '').trim();
  const tableLimit = Number(data.tableLimit || 0);

  const kpisHtml = renderPreviewKpis(data.kpis || []);
  const chartsHtml = renderPreviewCharts(data.chartResults || data.charts || []);

  currentDashboardTable = data.table || {
    fields: [],
    rows: []
  };

  currentDashboardTablePage = 1;

  const tableHtml = renderDashboardPagedTable(
    currentDashboardTable,
    currentDashboardTablePage
  );

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

      ${limitWarning ? `
        <div class="dashboard-warning-box">
          ${escapeHtml(limitWarning)}
        </div>
      ` : ''}

      ${tableLimit ? `
        <div class="dashboard-note-box">
          ตารางด้านล่างแสดงสูงสุด ${tableLimit.toLocaleString()} แถวแรก เพื่อให้หน้าเว็บทำงานเร็วขึ้น
        </div>
      ` : ''}
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
      

const newDashboardId = String(
  data.dashboardId ||
  data.id ||
  data['รหัส Dashboard'] ||
  ''
).trim();

if (newDashboardId) {
  if (el.dashboardSelect) {
    el.dashboardSelect.value = newDashboardId;
  }

  if (el.designerDashboardSelect) {
    el.designerDashboardSelect.value = newDashboardId;
  }

  if (el.manageDashboardSelect) {
    el.manageDashboardSelect.value = newDashboardId;
    handleSelectManageDashboard();
  }
}

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

  function normalizeDashboardRows(data) {
  if (!data) {
    return [];
  }

  if (Array.isArray(data)) {
    return data;
  }

  if (Array.isArray(data.dashboards)) {
    return data.dashboards;
  }

  if (Array.isArray(data.rows)) {
    return data.rows;
  }

  if (Array.isArray(data.items)) {
    return data.items;
  }

  if (data.data && Array.isArray(data.data.dashboards)) {
    return data.data.dashboards;
  }

  return [];
}

function getDashboardId(dash) {
  dash = dash || {};

  return String(
    dash.dashboardId ||
    dash.id ||
    dash['รหัส Dashboard'] ||
    ''
  ).trim();
}

function getDashboardName(dash) {
  dash = dash || {};
  const id = getDashboardId(dash);

  return String(
    dash.dashboardName ||
    dash.name ||
    dash.title ||
    dash['ชื่อ Dashboard'] ||
    id ||
    ''
  ).trim();
}

function getDashboardType(dash) {
  dash = dash || {};

  return String(
    dash.dashboardType ||
    dash.type ||
    dash['ประเภท Dashboard'] ||
    ''
  ).trim();
}

function getDashboardDescription(dash) {
  dash = dash || {};

  return String(
    dash.description ||
    dash.dashboardDescription ||
    dash['คำอธิบาย'] ||
    ''
  ).trim();
}

function getDashboardPublish(dash) {
  dash = dash || {};

  return toBool(
    dash.publish !== undefined
      ? dash.publish
      : dash.published !== undefined
        ? dash.published
        : dash['สถานะ Publish']
  );
}

function getDashboardExport(dash) {
  dash = dash || {};

  return toBool(
    dash.allowExport !== undefined
      ? dash.allowExport
      : dash.export !== undefined
        ? dash.export
        : dash['อนุญาต Export']
  );
}

function getDashboardHidden(dash) {
  dash = dash || {};

  return toBool(
    dash.hide !== undefined
      ? dash.hide
      : dash.hidden !== undefined
        ? dash.hidden
        : dash['ซ่อน Dashboard']
  );
}

function getDashboardVisibility(dash) {
  dash = dash || {};

  return String(
    dash.visibility ||
    dash.allowedRoles ||
    dash['สิทธิ์ที่มองเห็น'] ||
    ''
  ).trim();
}
  function renderAllDashboardDropdowns(dashboards) {
  dashboards = Array.isArray(dashboards) ? dashboards : [];

  renderDashboardViewerOptions(dashboards);
  renderDesignerDashboardOptions(dashboards);
  renderManageDashboardOptions(dashboards);

  if (typeof renderManageDashboardCards === 'function') {
    renderManageDashboardCards(dashboards);
  }

  if (typeof renderManageDashboardCount === 'function') {
    renderManageDashboardCount(dashboards.length, dashboards.length);
  } else {
    if (el.manageDashboardCount) {
      el.manageDashboardCount.textContent = dashboards.length.toLocaleString() + ' รายการ';
    }

    if (el.manageDashboardCardCount) {
      el.manageDashboardCardCount.textContent = dashboards.length.toLocaleString() + ' รายการ';
    }
  }

  if (el.dashboardResult) {
    el.dashboardResult.textContent = dashboards.length
      ? 'โหลด Dashboard แล้ว ' + dashboards.length.toLocaleString() + ' รายการ'
      : 'ไม่พบ Dashboard';
  }
}

function renderDashboardViewerOptions(dashboards) {
  if (!el.dashboardSelect) {
    return;
  }

  const currentValue = el.dashboardSelect.value || '';

  if (!dashboards || !dashboards.length) {
    el.dashboardSelect.innerHTML = '<option value="">ยังไม่มีรายการ Dashboard</option>';
    return;
  }

  el.dashboardSelect.innerHTML =
    '<option value="">เลือก Dashboard</option>' +
    dashboards.map(function (dash) {
      const id = getDashboardId(dash);
      const name = getDashboardName(dash);
      const type = getDashboardType(dash);
      const publish = getDashboardPublish(dash);

      return (
        '<option value="' + escapeAttr(id) + '">' +
        escapeHtml(name) +
        (type ? ' (' + escapeHtml(type) + ')' : '') +
        (publish ? '' : ' - Draft') +
        '</option>'
      );
    }).join('');

  if (currentValue && dashboards.some(function (dash) {
    return getDashboardId(dash) === currentValue;
  })) {
    el.dashboardSelect.value = currentValue;
  }
}

function renderDesignerDashboardOptions(dashboards) {
  if (!el.designerDashboardSelect) {
    return;
  }

  const currentValue = el.designerDashboardSelect.value || '';

  if (!dashboards || !dashboards.length) {
    el.designerDashboardSelect.innerHTML = '<option value="">ยังไม่มี Dashboard ให้เลือก</option>';
    return;
  }

  el.designerDashboardSelect.innerHTML =
    '<option value="">เลือก Dashboard ที่ต้องการออกแบบ</option>' +
    dashboards.map(function (dash) {
      const id = getDashboardId(dash);
      const name = getDashboardName(dash);
      const type = getDashboardType(dash);
      const publish = getDashboardPublish(dash);

      return (
        '<option value="' + escapeAttr(id) + '">' +
        escapeHtml(name) +
        (type ? ' (' + escapeHtml(type) + ')' : '') +
        (publish ? '' : ' - Draft') +
        '</option>'
      );
    }).join('');

  if (currentValue && dashboards.some(function (dash) {
    return getDashboardId(dash) === currentValue;
  })) {
    el.designerDashboardSelect.value = currentValue;
  }
}
  

  async function loadDashboardOptions() {
  try {
    if (!window.AnalyticsAPI || typeof window.AnalyticsAPI.listDashboards !== 'function') {
      throw new Error('ไม่พบ window.AnalyticsAPI.listDashboards กรุณาตรวจสอบ api.js');
    }

    if (el.dashboardResult) {
      el.dashboardResult.textContent = 'กำลังโหลด Dashboard...';
    }

    const data = await window.AnalyticsAPI.listDashboards({
      includeDeleted: true
    });

    const dashboards = normalizeDashboardRows(data);

    dashboardsCache = dashboards;
    manageDashboardsCache = dashboards;

    renderAllDashboardDropdowns(dashboards);

    writeLog({
      step: 'dashboard_options_loaded',
      total: dashboards.length,
      response: data
    });

    return dashboards;

  } catch (error) {
    dashboardsCache = [];
    manageDashboardsCache = [];

    renderAllDashboardDropdowns([]);

    const message = getFriendlyErrorMessage
      ? getFriendlyErrorMessage(error)
      : (error.message || 'โหลด Dashboard ไม่สำเร็จ');

    if (el.dashboardResult) {
      el.dashboardResult.textContent = message;
    }

    setDashboardViewMessage(message);

    writeLog({
      step: 'dashboard_options_load_error',
      message: message,
      payload: error.payload || null
    });

    return [];
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
      setButtonLoading(el.refreshDashboardBtn, false, 'Refresh Dashboard');
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

  currentDashboardTable = data.table || {
    fields: [],
    rows: []
  };

  currentDashboardTablePage = 1;

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
    <div data-dashboard-table-section>
      ${renderDashboardPagedTable(currentDashboardTable, currentDashboardTablePage)}
    </div>
  `;

  mountQueuedCharts();
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
  return loadDashboardOptions();
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
      const displayKey = field.displayName || '';
      const columnKey = field.columnName || '';

      const value =
        row[displayKey] !== undefined
          ? row[displayKey]
          : row[columnKey] !== undefined
            ? row[columnKey]
            : '';

      return `<td>${escapeHtml(value)}</td>`;
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
          class="btn btn-ghost"
          type="button"
          data-dashboard-table-page-action="prev"
          ${page <= 1 ? 'disabled' : ''}
        >
          ก่อนหน้า
        </button>

        <span>หน้า ${page.toLocaleString()} / ${totalPages.toLocaleString()}</span>

        <button
          class="btn btn-ghost"
          type="button"
          data-dashboard-table-page-action="next"
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
  async function runSystemCheck() {
  if (!el.systemCheckResult) {
    return;
  }

  el.systemCheckResult.classList.remove('empty');
  setPanelLoading(el.systemCheckResult, 'กำลังตรวจสอบระบบ...');

  if (el.runSystemCheckBtn) {
    setButtonLoading(el.runSystemCheckBtn, true, 'กำลังตรวจ...');
  }

  const checks = [];

  function addCheck(name, ok, detail) {
    checks.push({
      name: name,
      ok: !!ok,
      detail: detail || ''
    });
  }

  try {
    addCheck(
      'Frontend Version',
      !!window.DASHBOARD_FRONTEND_VERSION,
      window.DASHBOARD_FRONTEND_VERSION || 'ไม่พบ version'
    );

    addCheck(
      'AnalyticsAPI',
      !!window.AnalyticsAPI,
      window.AnalyticsAPI ? 'พร้อมใช้งาน' : 'ไม่พบ window.AnalyticsAPI'
    );

    addCheck(
      'ECharts',
      !!window.echarts,
      window.echarts ? 'โหลด ECharts แล้ว' : 'ไม่พบ ECharts'
    );

    addCheck(
      'SheetJS / XLSX',
      !!window.XLSX,
      window.XLSX ? 'โหลด SheetJS แล้ว' : 'ไม่พบ SheetJS ใช้ Export Excel แบบ fallback ได้'
    );

    addCheck(
      'Token Key',
      !!(window.AnalyticsAPI && window.AnalyticsAPI.TOKEN_KEY === 'analytics_dashboard_token'),
      window.AnalyticsAPI ? window.AnalyticsAPI.TOKEN_KEY : 'ไม่พบ TOKEN_KEY'
    );

    addCheck(
      'Dashboard Viewer Buttons',
      !!(
        el.openDashboardBtn &&
        el.refreshDashboardBtn &&
        el.applyDashboardFilterBtn &&
        el.exportDashboardBtn &&
        el.exportDashboardExcelBtn
      ),
      'Open / Refresh / Filter / Export CSV / Export Excel'
    );

    addCheck(
      'Table Pagination Container',
      !!el.dashboardViewResult,
      el.dashboardViewResult ? 'พร้อมใช้งาน' : 'ไม่พบ dashboardViewResult'
    );

    addCheck(
      'User Management UI',
      !!(el.reloadUsersBtn && el.userManageList),
      el.reloadUsersBtn && el.userManageList ? 'พร้อมใช้งาน' : 'ไม่พบ User Management UI'
    );

    addCheck(
      'Audit Log UI',
      !!(el.reloadAuditLogBtn && el.auditLogResult),
      el.reloadAuditLogBtn && el.auditLogResult ? 'พร้อมใช้งาน' : 'ไม่พบ Audit Log UI'
    );

    if (window.AnalyticsAPI && window.AnalyticsAPI.health) {
      try {
        const health = await window.AnalyticsAPI.health();

        addCheck(
          'API Health',
          !!health.ok,
          health.message || 'API ตอบกลับแล้ว'
        );
      } catch (err) {
        addCheck(
          'API Health',
          false,
          getFriendlyErrorMessage(err)
        );
      }
    }

    if (window.AnalyticsAPI && window.AnalyticsAPI.setupStatus) {
      try {
        const setup = await window.AnalyticsAPI.setupStatus();

        addCheck(
          'Master Sheet Setup',
          !!setup.ok,
          setup.ok
            ? 'โครงสร้างชีทครบถ้วน'
            : 'ยังขาดชีท: ' + (setup.missingSheetNames || []).join(', ')
        );
      } catch (err) {
        addCheck(
          'Master Sheet Setup',
          false,
          getFriendlyErrorMessage(err)
        );
      }
    }

    renderSystemCheckResult(checks);

    writeLog({
      step: 'system_check',
      checks: checks
    });

  } catch (error) {
    el.systemCheckResult.classList.add('empty');
    el.systemCheckResult.textContent = getFriendlyErrorMessage(error);

    logApiError('system_check_error', error);

  } finally {
    if (el.runSystemCheckBtn) {
      setButtonLoading(el.runSystemCheckBtn, false, 'Run System Check');
    }
  }
}


function renderSystemCheckResult(checks) {
  if (!el.systemCheckResult) {
    return;
  }

  checks = checks || [];

  const passed = checks.filter(function (item) {
    return item.ok;
  }).length;

  const total = checks.length;

  const rowsHtml = checks.map(function (item) {
    return `
      <div class="system-check-item ${item.ok ? 'is-pass' : 'is-fail'}">
        <div>
          <strong>${escapeHtml(item.name)}</strong>
          <span>${escapeHtml(item.detail || '-')}</span>
        </div>

        <b>${item.ok ? 'PASS' : 'FAIL'}</b>
      </div>
    `;
  }).join('');

  el.systemCheckResult.classList.remove('empty');
  el.systemCheckResult.innerHTML = `
    <div class="system-check-summary ${passed === total ? 'is-pass' : 'is-warn'}">
      ผ่าน ${passed.toLocaleString()} / ${total.toLocaleString()} รายการ
    </div>

    <div class="system-check-list">
      ${rowsHtml}
    </div>
  `;
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
  try {
    if (!window.AnalyticsAPI || typeof window.AnalyticsAPI.listDashboards !== 'function') {
      throw new Error('ไม่พบ window.AnalyticsAPI.listDashboards กรุณาตรวจสอบ api.js');
    }

    if (el.manageDashboardSelect) {
      el.manageDashboardSelect.innerHTML = '<option value="">กำลังโหลด Dashboard...</option>';
    }

    setManageDashboardMessage('กำลังโหลดรายการ Dashboard...');
    selectedManageDashboard = null;
    resetManageDashboardForm(false);

    if (el.manageDashboardSummary) {
      setPanelLoading(el.manageDashboardSummary, 'กำลังโหลดรายการ Dashboard...');
    }

    const data = await window.AnalyticsAPI.listDashboards({
      includeDeleted: true
    });

    const dashboards = normalizeDashboardRows(data);

    dashboardsCache = dashboards;
    manageDashboardsCache = dashboards;

    renderAllDashboardDropdowns(dashboards);
    setManageDashboardMessage('');

    if (el.manageDashboardSummary) {
      el.manageDashboardSummary.classList.add('empty');
      el.manageDashboardSummary.textContent = dashboards.length
        ? 'เลือก Dashboard เพื่อดูรายละเอียดและแก้ไข'
        : 'ยังไม่มี Dashboard ที่สร้างไว้';
    }

    writeLog({
      step: 'manage_dashboards_loaded',
      total: dashboards.length,
      response: data
    });

    return dashboards;

  } catch (error) {
    dashboardsCache = [];
    manageDashboardsCache = [];

    renderAllDashboardDropdowns([]);

    const message = getFriendlyErrorMessage(error);

    if (el.manageDashboardSelect) {
      el.manageDashboardSelect.innerHTML = '<option value="">โหลด Dashboard ไม่สำเร็จ</option>';
    }

    if (el.manageDashboardSummary) {
      el.manageDashboardSummary.classList.add('empty');
      el.manageDashboardSummary.textContent = message;
    }

    setManageDashboardMessage(message);

    writeLog({
      step: 'manage_dashboards_load_error',
      message: message,
      payload: error.payload || null
    });

    return [];
  }
}

 function renderFilteredManageDashboardOptions() {
  let dashboards = Array.isArray(manageDashboardsCache) ? manageDashboardsCache.slice() : [];

  const searchText = el.manageDashboardSearch
    ? String(el.manageDashboardSearch.value || '').trim().toLowerCase()
    : '';

  const typeFilter = el.manageDashboardTypeFilter
    ? String(el.manageDashboardTypeFilter.value || '').trim()
    : '';

  const statusFilter = el.manageDashboardStatusFilter
    ? String(el.manageDashboardStatusFilter.value || '').trim()
    : '';

  if (searchText) {
    dashboards = dashboards.filter(function (dash) {
      const text = [
        getDashboardId(dash),
        getDashboardName(dash),
        getDashboardType(dash),
        getDashboardDescription(dash),
        getDashboardVisibility(dash)
      ].join(' ').toLowerCase();

      return text.indexOf(searchText) >= 0;
    });
  }

  if (typeFilter) {
    dashboards = dashboards.filter(function (dash) {
      return getDashboardType(dash) === typeFilter;
    });
  }

  if (statusFilter) {
    dashboards = dashboards.filter(function (dash) {
      const publish = getDashboardPublish(dash);
      const hidden = getDashboardHidden(dash);

      if (statusFilter === 'published') {
        return publish && !hidden;
      }

      if (statusFilter === 'draft') {
        return !publish && !hidden;
      }

      if (statusFilter === 'hidden') {
        return hidden;
      }

      return true;
    });
  }

  renderManageDashboardOptions(dashboards);

  if (typeof renderManageDashboardCards === 'function') {
    renderManageDashboardCards(dashboards);
  }

  if (typeof renderManageDashboardCount === 'function') {
    renderManageDashboardCount(dashboards.length, manageDashboardsCache.length);
  } else if (el.manageDashboardCount) {
    el.manageDashboardCount.textContent =
      dashboards.length.toLocaleString() + ' / ' +
      manageDashboardsCache.length.toLocaleString() + ' รายการ';
  }
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

  dashboards = Array.isArray(dashboards) ? dashboards : [];
  const currentValue = el.manageDashboardSelect.value || '';

  if (!dashboards.length) {
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
      const id = getDashboardId(dash);
      const name = getDashboardName(dash);
      const type = getDashboardType(dash);
      const publish = getDashboardPublish(dash);
      const allowExport = getDashboardExport(dash);
      const hidden = getDashboardHidden(dash);

      const statusText = [
        hidden ? 'Hidden' : (publish ? 'Published' : 'Draft'),
        allowExport ? 'Export On' : 'Export Off'
      ].join(' / ');

      return (
        '<option value="' + escapeAttr(id) + '">' +
        escapeHtml(name || id) +
        (type ? ' (' + escapeHtml(type) + ')' : '') +
        ' - ' + escapeHtml(statusText) +
        '</option>'
      );
    }).join('');

  if (currentValue && dashboards.some(function (dash) {
    return getDashboardId(dash) === currentValue;
  })) {
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
    return getDashboardId(dash) === dashboardId;
  }) || null;
}



  function fillManageDashboardForm(dash) {
  if (!dash) {
    resetManageDashboardForm();
    return;
  }

  if (el.manageDashboardName) {
    el.manageDashboardName.value = getDashboardName(dash);
  }

  if (el.manageDashboardType) {
    el.manageDashboardType.value = getDashboardType(dash) || 'Custom';
  }

  if (el.manageDashboardDescription) {
    el.manageDashboardDescription.value = getDashboardDescription(dash);
  }

  if (el.manageDashboardPublish) {
    el.manageDashboardPublish.checked = getDashboardPublish(dash);
  }

  if (el.manageDashboardExport) {
    el.manageDashboardExport.checked = getDashboardExport(dash);
  }

  if (el.manageDashboardHidden) {
    el.manageDashboardHidden.checked = getDashboardHidden(dash);
  }

  setManageDashboardRoles(getDashboardVisibility(dash));
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

  const dashboardId = getDashboardId(dash);
  const name = getDashboardName(dash);
  const type = getDashboardType(dash);
  const sourceId = String(
    dash.sourceId ||
    dash.mainSourceId ||
    dash['รหัสแหล่งข้อมูลหลัก'] ||
    ''
  ).trim();
  const sheetName = String(
    dash.sheetName ||
    dash.mainSheetName ||
    dash['รหัสชีทย่อยหลัก'] ||
    ''
  ).trim();
  const publish = getDashboardPublish(dash);
  const allowExport = getDashboardExport(dash);
  const hidden = getDashboardHidden(dash);
  const visibility = getDashboardVisibility(dash) || '-';
  const updatedAt = String(
    dash.updatedAt ||
    dash['วันที่แก้ไขล่าสุด'] ||
    '-'
  ).trim();
  const updatedBy = String(
    dash.updatedBy ||
    dash['แก้ไขโดย'] ||
    '-'
  ).trim();

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
        <span>สถานะ</span>
        <strong>${hidden ? 'Hidden' : (publish ? 'Published' : 'Draft')}</strong>
      </div>

      <div class="dashboard-manage-summary-item">
        <span>Export</span>
        <strong>${allowExport ? 'เปิดใช้งาน' : 'ปิดใช้งาน'}</strong>
      </div>

      <div class="dashboard-manage-summary-item">
        <span>สิทธิ์ที่มองเห็น</span>
        <strong>${escapeHtml(visibility)}</strong>
      </div>

      <div class="dashboard-manage-summary-item">
        <span>แก้ไขล่าสุด</span>
        <strong>${escapeHtml(updatedAt || '-')}</strong>
      </div>

      <div class="dashboard-manage-summary-item">
        <span>แก้ไขโดย</span>
        <strong>${escapeHtml(updatedBy || '-')}</strong>
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
    function bindDashboardDesignerEvents() {
    if (el.openDesignerBtn) {
      el.openDesignerBtn.addEventListener('click', function () {
        scrollToDesignerSection();
      });
    }

    if (el.designerReloadBtn) {
      el.designerReloadBtn.addEventListener('click', initDashboardDesigner);
    }

    if (el.designerResetBtn) {
      el.designerResetBtn.addEventListener('click', resetDashboardDesigner);
    }

    if (el.designerSourceSelect) {
      el.designerSourceSelect.addEventListener('change', handleDesignerSourceChange);
    }

    if (el.designerSheetSelect) {
      el.designerSheetSelect.addEventListener('change', function () {
        designerSelectedSheetName = el.designerSheetSelect.value || '';
      });
    }

    if (el.designerDashboardSelect) {
      el.designerDashboardSelect.addEventListener('change', function () {
        designerSelectedDashboardId = el.designerDashboardSelect.value || '';
      });
    }

    if (el.designerAnalyzeBtn) {
      el.designerAnalyzeBtn.addEventListener('click', handleDesignerAnalyzeSheet);
    }

    if (el.designerSuggestBtn) {
      el.designerSuggestBtn.addEventListener('click', handleDesignerSuggestWidgets);
    }

    if (el.designerWidgetType) {
      el.designerWidgetType.addEventListener('change', function () {
        designerSelectedWidgetType = el.designerWidgetType.value || '';
        renderDesignerMiniPreviewByType(designerSelectedWidgetType);
      });
    }

    if (el.designerSaveFieldAnalysisBtn) {
      el.designerSaveFieldAnalysisBtn.addEventListener('click', handleDesignerSaveFieldAnalysis);
    }

    if (el.designerMiniPreviewBtn) {
      el.designerMiniPreviewBtn.addEventListener('click', function () {
        renderDesignerMiniPreviewByType(el.designerWidgetType ? el.designerWidgetType.value : '');
      });
    }
        if (el.designerRealPreviewBtn) {
    el.designerRealPreviewBtn.addEventListener('click', handleDesignerRealPreview);
  }

  if (el.designerComparisonPreviewBtn) {
    el.designerComparisonPreviewBtn.addEventListener('click', handleDesignerComparisonPreview);
  }

  if (el.designerSaveWidgetBtn) {
    el.designerSaveWidgetBtn.addEventListener('click', handleDesignerSaveWidget);
  }

  if (el.designerUpdateWidgetBtn) {
    el.designerUpdateWidgetBtn.addEventListener('click', handleDesignerSaveWidget);
  }

  if (el.designerLoadSavedWidgetsBtn) {
    el.designerLoadSavedWidgetsBtn.addEventListener('click', handleDesignerLoadSavedWidgets);
  }
  }
    /**
   * =====================================================
   * Dashboard Designer / Builder
   * =====================================================
   */

 async function initDashboardDesigner() {
  try {
    setDesignerMessage('กำลังโหลดข้อมูล Dashboard Designer...', 'muted');

    /**
     * โหลดข้อมูลหลักของ Designer
     * - Theme
     * - Widget Templates
     * - Sources
     * - Dashboards
     */
    const results = await Promise.allSettled([
      loadDesignerThemes(),
      loadDesignerWidgetTemplates(),
      loadDesignerSources(),
      loadDesignerDashboards()
    ]);

    const failed = results.filter(function (item) {
      return item.status === 'rejected';
    });

    if (failed.length) {
      const firstError = failed[0].reason;
      throw firstError || new Error('โหลดข้อมูลบางส่วนของ Dashboard Designer ไม่สำเร็จ');
    }

    /**
     * สำคัญ:
     * ป้องกันกรณี loadDesignerDashboards โหลดแล้ว
     * แต่ dropdown designerDashboardSelect ยังไม่ถูกเติมข้อมูล
     */
    if (el.designerDashboardSelect) {
      if (dashboardsCache && dashboardsCache.length) {
        renderDesignerDashboardOptions(dashboardsCache);
      } else if (manageDashboardsCache && manageDashboardsCache.length) {
        dashboardsCache = manageDashboardsCache.slice();
        renderDesignerDashboardOptions(dashboardsCache);
      } else {
        const loadedDashboards = await loadDashboardOptions();

        if (loadedDashboards && loadedDashboards.length) {
          renderDesignerDashboardOptions(loadedDashboards);
        } else {
          el.designerDashboardSelect.innerHTML =
            '<option value="">ยังไม่มี Dashboard ให้เลือก</option>';
        }
      }
    }

    /**
     * ถ้ามี Dashboard ถูกเลือกอยู่แล้ว ให้ sync state ของ Designer
     */
    if (el.designerDashboardSelect && el.designerDashboardSelect.value) {
      designerSelectedDashboardId = el.designerDashboardSelect.value;
    } else {
      designerSelectedDashboardId = '';
    }

    /**
     * ถ้ามี Source ถูกเลือกอยู่แล้ว ให้ sync state ของ Designer
     */
    if (el.designerSourceSelect && el.designerSourceSelect.value) {
      designerSelectedSourceId = el.designerSourceSelect.value;
    }

    /**
     * ถ้ามี Sheet ถูกเลือกอยู่แล้ว ให้ sync state ของ Designer
     */
    if (el.designerSheetSelect && el.designerSheetSelect.value) {
      designerSelectedSheetName = el.designerSheetSelect.value;
    }

    setDesignerMessage('โหลด Dashboard Designer สำเร็จ', 'success');

    writeLog({
      step: 'designer_init_success',
      dashboards: dashboardsCache ? dashboardsCache.length : 0,
      sources: sourcesCache ? sourcesCache.length : 0,
      themes: designerThemesCache ? designerThemesCache.length : 0,
      templates: designerWidgetTemplatesCache ? designerWidgetTemplatesCache.length : 0
    });

  } catch (error) {
    const message = getFriendlyErrorMessage
      ? getFriendlyErrorMessage(error)
      : (error.message || 'โหลด Dashboard Designer ไม่สำเร็จ');

    setDesignerMessage(message, 'error');

    if (el.designerDashboardSelect && (!dashboardsCache || !dashboardsCache.length)) {
      el.designerDashboardSelect.innerHTML =
        '<option value="">โหลดรายการ Dashboard ไม่สำเร็จ</option>';
    }

    writeLog({
      step: 'designer_init_error',
      message: message,
      payload: error.payload || null
    });
  }
}
  

  function resetDashboardDesigner() {
    designerFieldsCache = [];
    designerSuggestedWidgetsCache = [];
    designerSelectedSourceId = '';
    designerSelectedSheetName = '';
    designerSelectedDashboardId = '';
    designerSelectedWidgetType = '';
    designerEditingWidgetId = '';

    if (el.designerDashboardSelect) el.designerDashboardSelect.value = '';
    if (el.designerDashboardType) el.designerDashboardType.value = 'operation';
    if (el.designerSourceSelect) el.designerSourceSelect.value = '';
    if (el.designerSheetSelect) {
      el.designerSheetSelect.innerHTML = '<option value="">เลือก Sheet</option>';
    }

    if (el.designerDetectedSummary) {
      el.designerDetectedSummary.classList.add('empty');
      el.designerDetectedSummary.textContent = 'ยังไม่ได้วิเคราะห์ข้อมูล';
    }

    if (el.designerFieldList) {
      el.designerFieldList.classList.add('empty');
      el.designerFieldList.textContent = 'กด “วิเคราะห์ข้อมูล” เพื่อแสดง Field ที่ตรวจพบ';
    }

    if (el.designerSuggestedWidgets) {
      el.designerSuggestedWidgets.classList.add('empty');
      el.designerSuggestedWidgets.textContent = 'กด “แนะนำ Widget” เพื่อให้ระบบสร้างรายการแนะนำ';
    }

    if (el.designerLivePreview) {
      el.designerLivePreview.classList.add('empty');
      el.designerLivePreview.textContent = 'เลือก Widget และกด Preview เพื่อแสดงตัวอย่าง';
    }

    if (el.designerComparisonResult) {
      el.designerComparisonResult.className = 'comparison-result empty';
      el.designerComparisonResult.textContent = 'ยังไม่ได้คำนวณ Comparison';
    }

    resetDesignerWidgetForm();
    setDesignerMessage('');
  }


  function resetDesignerWidgetForm() {
    if (el.designerWidgetType) el.designerWidgetType.value = '';
    if (el.designerWidgetTitle) el.designerWidgetTitle.value = '';
    if (el.designerDateField) el.designerDateField.innerHTML = '<option value="">ไม่ใช้</option>';
    if (el.designerXField) el.designerXField.innerHTML = '<option value="">ไม่ใช้</option>';
    if (el.designerValueField) el.designerValueField.innerHTML = '<option value="">นับจำนวนรายการ</option>';
    if (el.designerStackField) el.designerStackField.innerHTML = '<option value="">ไม่ใช้</option>';
    if (el.designerAggregate) el.designerAggregate.value = 'count';
    if (el.designerUnit) el.designerUnit.value = '';
    if (el.designerSortMethod) el.designerSortMethod.value = 'desc';
    if (el.designerLimit) el.designerLimit.value = '';
    if (el.designerTargetValue) el.designerTargetValue.value = '';
    if (el.designerShowLegend) el.designerShowLegend.checked = true;
    if (el.designerShowLabel) el.designerShowLabel.checked = false;
    if (el.designerShowPercent) el.designerShowPercent.checked = false;
    if (el.designerEnableExport) el.designerEnableExport.checked = true;
    if (el.designerUpdateWidgetBtn) el.designerUpdateWidgetBtn.disabled = true;
  }


  function scrollToDesignerSection() {
    if (!el.dashboardDesignerSection) {
      return;
    }

    el.dashboardDesignerSection.scrollIntoView({
      behavior: 'smooth',
      block: 'start'
    });
  }


  async function loadDesignerThemes() {
    if (!window.AnalyticsAPI || !window.AnalyticsAPI.themes) {
      return;
    }

    const result = await window.AnalyticsAPI.themes();
    designerThemesCache = result.themes || [];

    renderDesignerThemes(designerThemesCache);
  }


  function renderDesignerThemes(themes) {
    if (!el.designerTheme) {
      return;
    }

    const list = Array.isArray(themes) ? themes : [];

    if (!list.length) {
      el.designerTheme.innerHTML = '<option value="executive_blue">Executive Blue</option>';
      return;
    }

    el.designerTheme.innerHTML = list.map(function (theme) {
      return '<option value="' + escapeHtml(theme.themeId) + '">' +
        escapeHtml(theme.themeName || theme.themeId) +
        '</option>';
    }).join('');

    const defaultTheme = list.find(function (theme) {
      return theme.isDefault;
    });

    if (defaultTheme) {
      el.designerTheme.value = defaultTheme.themeId;
    }
  }


  async function loadDesignerWidgetTemplates() {
    if (!window.AnalyticsAPI || !window.AnalyticsAPI.widgetTemplates) {
      return;
    }

    const result = await window.AnalyticsAPI.widgetTemplates();
    designerWidgetTemplatesCache = result.templates || [];

    renderDesignerWidgetGallery(designerWidgetTemplatesCache);
  }


  function renderDesignerWidgetGallery(templates) {
    if (!el.designerWidgetGallery) {
      return;
    }

    const list = Array.isArray(templates) ? templates : [];

    if (!list.length) {
      el.designerWidgetGallery.classList.add('empty');
      el.designerWidgetGallery.textContent = 'ยังไม่มี Widget Template';
      return;
    }

    el.designerWidgetGallery.classList.remove('empty');

    el.designerWidgetGallery.innerHTML = list.map(function (tpl) {
      return [
        '<button class="widget-option-card" type="button" data-widget-type="' + escapeHtml(tpl.widgetType) + '">',
          '<div class="widget-option-head">',
            '<div class="widget-option-title">',
              '<strong>' + escapeHtml(tpl.templateName || tpl.widgetType) + '</strong>',
              '<small>' + escapeHtml(tpl.bestFor || '') + '</small>',
            '</div>',
            '<span class="widget-type-pill">' + escapeHtml(tpl.widgetType || '') + '</span>',
          '</div>',
          '<div class="widget-mini-preview">' + escapeHtml(tpl.miniPreviewText || '') + '</div>',
          '<div class="widget-option-meta">',
            'ต้องใช้: ' + escapeHtml(tpl.requiredFields || '-'),
          '</div>',
        '</button>'
      ].join('');
    }).join('');

    el.designerWidgetGallery.querySelectorAll('[data-widget-type]').forEach(function (btn) {
      btn.addEventListener('click', function () {
        const type = btn.getAttribute('data-widget-type') || '';
        selectDesignerWidgetType(type);
      });
    });
  }


 async function loadDesignerSources() {
  if (!el.designerSourceSelect) {
    return;
  }

  try {
    let list = sourcesCache || [];

    if (!list.length) {
      let result = null;

      if (window.AnalyticsAPI && typeof window.AnalyticsAPI.sources === 'function') {
        result = await window.AnalyticsAPI.sources();
      } else if (window.AnalyticsAPI && typeof window.AnalyticsAPI.listSources === 'function') {
        result = await window.AnalyticsAPI.listSources();
      }

      if (result) {
        list =
          result.sources ||
          result.items ||
          result.data ||
          result.rows ||
          [];
      }

      sourcesCache = list;
    }

    if (!Array.isArray(list) || !list.length) {
      el.designerSourceSelect.innerHTML = '<option value="">ยังไม่มี Source</option>';
      return;
    }

    el.designerSourceSelect.innerHTML =
      '<option value="">เลือก Source</option>' +
      list.map(function (source) {
        const id =
          source.sourceId ||
          source.sourceID ||
          source['Source ID'] ||
          source['รหัสแหล่งข้อมูล'] ||
          source.id ||
          '';

        const name =
          source.sourceName ||
          source.name ||
          source['Source Name'] ||
          source['ชื่อแหล่งข้อมูล'] ||
          id;

        if (!id) {
          return '';
        }

        return (
          '<option value="' + escapeHtml(id) + '">' +
          escapeHtml(name) +
          '</option>'
        );
      }).join('');

  } catch (error) {
    el.designerSourceSelect.innerHTML = '<option value="">โหลด Source ไม่สำเร็จ</option>';

    setDesignerMessage(
      error.message || 'โหลด Source ไม่สำเร็จ',
      'error'
    );
  }
}


  async function loadDesignerDashboards() {
  if (!el.designerDashboardSelect) {
    return [];
  }

  try {
    let dashboards = Array.isArray(dashboardsCache) ? dashboardsCache.slice() : [];

    if (!dashboards.length) {
      if (!window.AnalyticsAPI || typeof window.AnalyticsAPI.listDashboards !== 'function') {
        el.designerDashboardSelect.innerHTML = '<option value="">ไม่พบ API สำหรับโหลด Dashboard</option>';
        return [];
      }

      const result = await window.AnalyticsAPI.listDashboards({
        includeDeleted: true
      });

      dashboards = normalizeDashboardRows(result);
      dashboardsCache = dashboards;
      manageDashboardsCache = dashboards;
    }

    renderDesignerDashboardOptions(dashboards);

    writeLog({
      step: 'designer_load_dashboards',
      total: dashboards.length
    });

    return dashboards;

  } catch (error) {
    const message = getFriendlyErrorMessage(error);

    el.designerDashboardSelect.innerHTML =
      '<option value="">โหลด Dashboard ไม่สำเร็จ</option>';

    setDesignerMessage(message || 'โหลดรายการ Dashboard ไม่สำเร็จ', 'error');

    writeLog({
      step: 'designer_load_dashboards_error',
      message: message,
      payload: error.payload || null
    });

    return [];
  }
}



  async function handleDesignerSourceChange() {
  designerSelectedSourceId = el.designerSourceSelect ? el.designerSourceSelect.value : '';
  designerSelectedSheetName = '';

  if (!designerSelectedSourceId) {
    if (el.designerSheetSelect) {
      el.designerSheetSelect.innerHTML = '<option value="">เลือก Sheet</option>';
    }
    return;
  }

  try {
    setDesignerMessage('กำลังโหลดรายชื่อ Sheet...', 'muted');

    let result = null;

    if (window.AnalyticsAPI && typeof window.AnalyticsAPI.sourceSheets === 'function') {
      result = await window.AnalyticsAPI.sourceSheets({
        sourceId: designerSelectedSourceId
      });
    } else if (window.AnalyticsAPI && typeof window.AnalyticsAPI.listSourceSheets === 'function') {
      result = await window.AnalyticsAPI.listSourceSheets({
        sourceId: designerSelectedSourceId
      });
    } else {
      throw new Error('ไม่พบ API สำหรับโหลดรายชื่อ Sheet');
    }

    const sheets =
      result.sheets ||
      result.items ||
      result.data ||
      result.rows ||
      [];

    renderDesignerSheetOptions(sheets);

    setDesignerMessage('โหลดรายชื่อ Sheet สำเร็จ', 'success');

  } catch (error) {
    setDesignerMessage(error.message || 'โหลดรายชื่อ Sheet ไม่สำเร็จ', 'error');

    if (el.designerSheetSelect) {
      el.designerSheetSelect.innerHTML = '<option value="">โหลด Sheet ไม่สำเร็จ</option>';
    }
  }
}


  function renderDesignerSheetOptions(sheets) {
    if (!el.designerSheetSelect) {
      return;
    }

    const list = Array.isArray(sheets) ? sheets : [];

    if (!list.length) {
      el.designerSheetSelect.innerHTML = '<option value="">ไม่พบ Sheet</option>';
      return;
    }

    el.designerSheetSelect.innerHTML = '<option value="">เลือก Sheet</option>' + list.map(function (sheet) {
      const name = typeof sheet === 'string'
        ? sheet
        : (sheet.sheetName || sheet.name || sheet['ชื่อชีท'] || '');

      return '<option value="' + escapeHtml(name) + '">' +
        escapeHtml(name) +
        '</option>';
    }).join('');
  }


  async function handleDesignerAnalyzeSheet() {
    const sourceId = getDesignerSourceId();
    const sheetName = getDesignerSheetName();
    const sampleLimit = Number(el.designerSampleLimit ? el.designerSampleLimit.value : 200);

    if (!sourceId || !sheetName) {
      setDesignerMessage('กรุณาเลือก Source และ Sheet ก่อนวิเคราะห์ข้อมูล', 'error');
      return;
    }

    try {
      setDesignerMessage('กำลังวิเคราะห์ข้อมูล...', 'muted');
      setDesignerBusy(true);

      const result = await window.AnalyticsAPI.analyzeSheet({
        sourceId: sourceId,
        sheetName: sheetName,
        sampleLimit: sampleLimit
      });

      designerFieldsCache = result.fields || [];

      renderDesignerDetectedSummary(result);
      renderDesignerFieldList(designerFieldsCache);
      populateDesignerFieldSelects(designerFieldsCache);

      setDesignerMessage('วิเคราะห์ข้อมูลสำเร็จ', 'success');

      writeLog({
        step: 'designer_analyze_sheet',
        result: result
      });

    } catch (error) {
      setDesignerMessage(error.message || 'วิเคราะห์ข้อมูลไม่สำเร็จ', 'error');

      writeLog({
        step: 'designer_analyze_error',
        message: error.message,
        payload: error.payload || null
      });

    } finally {
      setDesignerBusy(false);
    }
  }


  function renderDesignerDetectedSummary(result) {
    if (!el.designerDetectedSummary) {
      return;
    }

    const fields = result.fields || [];

    const dateFields = fields.filter(function (f) { return f.isDateField; });
    const numberFields = fields.filter(function (f) { return f.isNumberField; });
    const categoryFields = fields.filter(function (f) { return f.isCategoryField; });
    const statusFields = fields.filter(function (f) { return f.isStatusField; });
    const filterFields = fields.filter(function (f) { return f.isFilterField; });

    el.designerDetectedSummary.classList.remove('empty');

    el.designerDetectedSummary.innerHTML = [
      '<div><strong>จำนวนแถวทั้งหมด:</strong> ' + escapeHtml(result.totalRows || '-') + '</div>',
      '<div><strong>จำนวนคอลัมน์:</strong> ' + escapeHtml(result.totalColumns || '-') + '</div>',
      '<div><strong>วันที่:</strong> ' + escapeHtml(dateFields.map(function (f) { return f.columnName; }).join(', ') || '-') + '</div>',
      '<div><strong>ตัวเลข:</strong> ' + escapeHtml(numberFields.map(function (f) { return f.columnName; }).join(', ') || '-') + '</div>',
      '<div><strong>หมวดหมู่:</strong> ' + escapeHtml(categoryFields.map(function (f) { return f.columnName; }).join(', ') || '-') + '</div>',
      '<div><strong>สถานะ:</strong> ' + escapeHtml(statusFields.map(function (f) { return f.columnName; }).join(', ') || '-') + '</div>',
      '<div><strong>Filter:</strong> ' + escapeHtml(filterFields.map(function (f) { return f.columnName; }).join(', ') || '-') + '</div>'
    ].join('');
  }


  function renderDesignerFieldList(fields) {
    if (!el.designerFieldList) {
      return;
    }

    const list = Array.isArray(fields) ? fields : [];

    if (!list.length) {
      el.designerFieldList.classList.add('empty');
      el.designerFieldList.textContent = 'ไม่พบ Field ที่วิเคราะห์ได้';
      return;
    }

    el.designerFieldList.classList.remove('empty');

    el.designerFieldList.innerHTML = list.map(function (field, index) {
      const type = field.manualType || field.detectedType || 'text';
      const semantic = field.semanticType || '';

      return [
        '<div class="detected-field-card" data-field-index="' + index + '">',
          '<div class="detected-field-head">',
            '<div class="detected-field-name">',
              '<strong>' + escapeHtml(field.columnName || '') + '</strong>',
              '<small>ตัวอย่าง: ' + escapeHtml((field.sampleValues || []).join(' | ') || '-') + '</small>',
            '</div>',
            '<span class="field-badge ' + escapeHtml(type) + '">' + escapeHtml(type) + '</span>',
          '</div>',

          '<div class="field-badge-row">',
            field.isDateField ? '<span class="field-badge">Date</span>' : '',
            field.isNumberField ? '<span class="field-badge number">Number</span>' : '',
            field.isCategoryField ? '<span class="field-badge category">Category</span>' : '',
            field.isStatusField ? '<span class="field-badge status">Status</span>' : '',
            field.isFilterField ? '<span class="field-badge text">Filter</span>' : '',
          '</div>',

          '<div class="detected-field-edit">',
            '<select data-field-manual-type="' + index + '">',
              renderTypeOptions(type),
            '</select>',
            '<input data-field-display-name="' + index + '" type="text" value="' + escapeHtml(field.displayName || field.columnName || '') + '" placeholder="ชื่อแสดงผล">',
          '</div>',
        '</div>'
      ].join('');
    }).join('');
  }


  function renderTypeOptions(selected) {
    const types = [
      'datetime',
      'date',
      'number',
      'currency',
      'percent',
      'category',
      'status',
      'person',
      'location',
      'vehicle',
      'text',
      'boolean',
      'image',
      'url',
      'id'
    ];

    return types.map(function (type) {
      return '<option value="' + type + '"' + (type === selected ? ' selected' : '') + '>' + type + '</option>';
    }).join('');
  }


  function populateDesignerFieldSelects(fields) {
    const list = Array.isArray(fields) ? fields : [];

    const allOptions = '<option value="">ไม่ใช้</option>' + list.map(function (field) {
      return '<option value="' + escapeHtml(field.columnName) + '">' +
        escapeHtml(field.displayName || field.columnName) +
        '</option>';
    }).join('');

    const valueOptions = '<option value="">นับจำนวนรายการ</option>' + list
      .filter(function (field) {
        return field.isNumberField || field.semanticType === 'measure';
      })
      .map(function (field) {
        return '<option value="' + escapeHtml(field.columnName) + '">' +
          escapeHtml(field.displayName || field.columnName) +
          '</option>';
      }).join('');

    if (el.designerDateField) el.designerDateField.innerHTML = allOptions;
    if (el.designerXField) el.designerXField.innerHTML = allOptions;
    if (el.designerStackField) el.designerStackField.innerHTML = allOptions;
    if (el.designerCompareCategoryField) el.designerCompareCategoryField.innerHTML = allOptions;
    if (el.designerValueField) el.designerValueField.innerHTML = valueOptions;

    const firstDate = list.find(function (field) {
      return field.isDateField;
    });

    const firstNumber = list.find(function (field) {
      return field.isNumberField;
    });

    const firstCategory = list.find(function (field) {
      return field.isCategoryField || field.semanticType === 'location' || field.semanticType === 'category';
    });

    if (firstDate && el.designerDateField) el.designerDateField.value = firstDate.columnName;
    if (firstNumber && el.designerValueField) el.designerValueField.value = firstNumber.columnName;
    if (firstCategory && el.designerXField) el.designerXField.value = firstCategory.columnName;
    if (firstCategory && el.designerCompareCategoryField) el.designerCompareCategoryField.value = firstCategory.columnName;
  }


  async function handleDesignerSaveFieldAnalysis() {
    if (!designerFieldsCache.length) {
      setDesignerMessage('ยังไม่มี Field Analysis ให้บันทึก', 'error');
      return;
    }

    const sourceId = getDesignerSourceId();
    const sheetName = getDesignerSheetName();

    const fields = designerFieldsCache.map(function (field, index) {
      const typeEl = el.designerFieldList.querySelector('[data-field-manual-type="' + index + '"]');
      const nameEl = el.designerFieldList.querySelector('[data-field-display-name="' + index + '"]');

      return {
        ...field,
        sourceId: sourceId,
        sheetName: sheetName,
        columnName: field.columnName,
        displayName: nameEl ? nameEl.value : field.displayName,
        manualType: typeEl ? typeEl.value : field.manualType
      };
    });

    try {
      setDesignerMessage('กำลังบันทึก Field Analysis...', 'muted');

      await window.AnalyticsAPI.fieldAnalysis({
        mode: 'save',
        fields: fields
      });

      designerFieldsCache = fields;
      renderDesignerFieldList(designerFieldsCache);
      populateDesignerFieldSelects(designerFieldsCache);

      setDesignerMessage('บันทึก Field Analysis สำเร็จ', 'success');

    } catch (error) {
      setDesignerMessage(error.message || 'บันทึก Field Analysis ไม่สำเร็จ', 'error');
    }
  }


  async function handleDesignerSuggestWidgets() {
    const sourceId = getDesignerSourceId();
    const sheetName = getDesignerSheetName();
    const dashboardType = el.designerDashboardType ? el.designerDashboardType.value : 'operation';

    if (!sourceId || !sheetName) {
      setDesignerMessage('กรุณาเลือก Source และ Sheet ก่อนแนะนำ Widget', 'error');
      return;
    }

    try {
      setDesignerMessage('กำลังสร้าง Widget แนะนำ...', 'muted');
      setDesignerBusy(true);

      const result = await window.AnalyticsAPI.suggestWidgets({
        sourceId: sourceId,
        sheetName: sheetName,
        dashboardType: dashboardType
      });

      designerSuggestedWidgetsCache = result.suggestions || [];

      renderDesignerSuggestedWidgets(designerSuggestedWidgetsCache);

      setDesignerMessage('แนะนำ Widget สำเร็จ', 'success');

      writeLog({
        step: 'designer_suggest_widgets',
        result: result
      });

    } catch (error) {
      setDesignerMessage(error.message || 'แนะนำ Widget ไม่สำเร็จ', 'error');

      writeLog({
        step: 'designer_suggest_error',
        message: error.message,
        payload: error.payload || null
      });

    } finally {
      setDesignerBusy(false);
    }
  }


  function renderDesignerSuggestedWidgets(suggestions) {
    if (!el.designerSuggestedWidgets) {
      return;
    }

    const list = Array.isArray(suggestions) ? suggestions : [];

    if (!list.length) {
      el.designerSuggestedWidgets.classList.add('empty');
      el.designerSuggestedWidgets.textContent = 'ยังไม่มี Widget ที่ระบบแนะนำ';
      return;
    }

    el.designerSuggestedWidgets.classList.remove('empty');

    el.designerSuggestedWidgets.innerHTML = list.map(function (item, index) {
      return [
        '<div class="suggested-widget-card" data-suggest-index="' + index + '">',
          '<div class="suggested-widget-top">',
            '<div>',
              '<strong>' + escapeHtml(item.title || item.widgetType || '') + '</strong>',
              '<small>' + escapeHtml(item.reason || item.description || '') + '</small>',
            '</div>',
            '<span class="widget-type-pill">' + escapeHtml(item.widgetType || '') + '</span>',
          '</div>',
          '<div class="suggested-widget-actions">',
            '<button class="btn btn-primary" type="button" data-use-suggestion="' + index + '">ใช้รายการนี้</button>',
            '<button class="btn btn-ghost" type="button" data-mini-suggestion="' + index + '">ดูตัวอย่างเล็ก</button>',
          '</div>',
        '</div>'
      ].join('');
    }).join('');

    el.designerSuggestedWidgets.querySelectorAll('[data-use-suggestion]').forEach(function (btn) {
      btn.addEventListener('click', function () {
        const index = Number(btn.getAttribute('data-use-suggestion'));
        applyDesignerSuggestion(index);
      });
    });

    el.designerSuggestedWidgets.querySelectorAll('[data-mini-suggestion]').forEach(function (btn) {
      btn.addEventListener('click', function () {
        const index = Number(btn.getAttribute('data-mini-suggestion'));
        const item = designerSuggestedWidgetsCache[index];
        renderDesignerSuggestionMiniPreview(item);
      });
    });
  }


  function selectDesignerWidgetType(type) {
    designerSelectedWidgetType = type;

    if (el.designerWidgetType) {
      el.designerWidgetType.value = type;
    }

    if (el.designerWidgetGallery) {
      el.designerWidgetGallery.querySelectorAll('.widget-option-card').forEach(function (card) {
        card.classList.toggle('active', card.getAttribute('data-widget-type') === type);
      });
    }

    renderDesignerMiniPreviewByType(type);
  }


  function applyDesignerSuggestion(index) {
    const item = designerSuggestedWidgetsCache[index];

    if (!item) {
      return;
    }

    if (el.designerWidgetType) el.designerWidgetType.value = item.widgetType || '';
    if (el.designerWidgetTitle) el.designerWidgetTitle.value = item.title || '';
    if (el.designerDateField && item.dateField) el.designerDateField.value = item.dateField;
    if (el.designerXField && (item.xField || item.categoryField)) el.designerXField.value = item.xField || item.categoryField;
    if (el.designerValueField && item.valueField !== undefined) el.designerValueField.value = item.valueField || '';
    if (el.designerStackField && item.stackField) el.designerStackField.value = item.stackField;
    if (el.designerAggregate) el.designerAggregate.value = item.aggregate || 'count';
    if (el.designerLimit && item.limit !== undefined) el.designerLimit.value = item.limit || '';
    if (el.designerTheme && item.theme) el.designerTheme.value = item.theme;

    designerSelectedWidgetType = item.widgetType || '';

    renderDesignerSuggestionMiniPreview(item);
    setDesignerMessage('เลือก Widget แนะนำแล้ว สามารถแก้ไขค่าก่อน Preview ได้', 'success');
  }


  function renderDesignerMiniPreviewByType(type) {
    if (!el.designerLivePreview) {
      return;
    }

    if (!type) {
      el.designerLivePreview.classList.add('empty');
      el.designerLivePreview.textContent = 'เลือก Widget เพื่อแสดง Mini Preview';
      return;
    }

    const template = designerWidgetTemplatesCache.find(function (tpl) {
      return tpl.widgetType === type;
    });

    const title = template ? template.templateName : type;
    const preview = template ? template.miniPreviewText : type;

    el.designerLivePreview.classList.remove('empty');

    el.designerLivePreview.innerHTML = [
      '<div class="preview-widget-card">',
        '<div class="widget-option-head">',
          '<div class="widget-option-title">',
            '<strong>' + escapeHtml(title) + '</strong>',
            '<small>Mini Preview</small>',
          '</div>',
          '<span class="widget-type-pill">' + escapeHtml(type) + '</span>',
        '</div>',
        '<div class="widget-mini-preview" style="min-height:120px;margin-top:14px;">',
          escapeHtml(preview || ''),
        '</div>',
        '<p class="widget-option-meta" style="margin-top:12px;">',
          escapeHtml(template ? template.bestFor : ''),
        '</p>',
      '</div>'
    ].join('');
  }


  function renderDesignerSuggestionMiniPreview(item) {
    if (!item || !el.designerLivePreview) {
      return;
    }

    el.designerLivePreview.classList.remove('empty');

    el.designerLivePreview.innerHTML = [
      '<div class="preview-widget-card">',
        '<div class="widget-option-head">',
          '<div class="widget-option-title">',
            '<strong>' + escapeHtml(item.title || item.widgetType || '') + '</strong>',
            '<small>' + escapeHtml(item.description || item.reason || '') + '</small>',
          '</div>',
          '<span class="widget-type-pill">' + escapeHtml(item.widgetType || '') + '</span>',
        '</div>',
        '<div class="designer-summary-box" style="margin-top:14px;">',
          '<div><strong>Source:</strong> ' + escapeHtml(item.sourceId || '-') + '</div>',
          '<div><strong>Sheet:</strong> ' + escapeHtml(item.sheetName || '-') + '</div>',
          '<div><strong>Date:</strong> ' + escapeHtml(item.dateField || '-') + '</div>',
          '<div><strong>X/Category:</strong> ' + escapeHtml(item.xField || item.categoryField || '-') + '</div>',
          '<div><strong>Value:</strong> ' + escapeHtml(item.valueField || 'Count rows') + '</div>',
          '<div><strong>Aggregate:</strong> ' + escapeHtml(item.aggregate || '-') + '</div>',
        '</div>',
      '</div>'
    ].join('');
  }


  function getDesignerSourceId() {
    return el.designerSourceSelect ? String(el.designerSourceSelect.value || '').trim() : '';
  }


  function getDesignerSheetName() {
    return el.designerSheetSelect ? String(el.designerSheetSelect.value || '').trim() : '';
  }


  function setDesignerMessage(message, type) {
    if (!el.designerMessage) {
      return;
    }

    el.designerMessage.textContent = message || '';

    el.designerMessage.classList.remove('success', 'error', 'muted');

    if (type) {
      el.designerMessage.classList.add(type);
    }
  }


  function setDesignerBusy(isBusy) {
    const buttons = [
      el.designerAnalyzeBtn,
      el.designerSuggestBtn,
      el.designerReloadBtn,
      el.designerResetBtn
    ];

    buttons.forEach(function (btn) {
      if (btn) {
        btn.disabled = !!isBusy;
      }
    });
  }
    function escapeHtml(value) {
    return String(value === undefined || value === null ? '' : value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }
    /**
   * =====================================================
   * Dashboard Designer / Builder - Round 2
   * Preview / Compare / Save / Load Saved Widgets
   * =====================================================
   */

  function collectDesignerWidgetConfig() {
    const widgetType = el.designerWidgetType ? el.designerWidgetType.value : '';
    const title = el.designerWidgetTitle ? el.designerWidgetTitle.value.trim() : '';

    return {
      widgetId: designerEditingWidgetId || '',
      dashboardId: getDesignerDashboardId(),
      widgetType: widgetType,
      title: title || getDefaultWidgetTitle(widgetType),
      description: '',
      sourceId: getDesignerSourceId(),
      sourceName: getDesignerSourceName(),
      sheetName: getDesignerSheetName(),

      dateField: el.designerDateField ? el.designerDateField.value : '',
      xField: el.designerXField ? el.designerXField.value : '',
      yField: '',
      valueField: el.designerValueField ? el.designerValueField.value : '',
      categoryField: el.designerXField ? el.designerXField.value : '',
      groupField: '',
      stackField: el.designerStackField ? el.designerStackField.value : '',

      aggregate: el.designerAggregate ? el.designerAggregate.value : 'count',
      unit: el.designerUnit ? el.designerUnit.value.trim() : '',
      sortMethod: el.designerSortMethod ? el.designerSortMethod.value : 'desc',
      limit: el.designerLimit ? Number(el.designerLimit.value || 0) : 0,

      theme: el.designerTheme ? el.designerTheme.value : 'executive_blue',
      colorMode: 'auto',
      customColors: '',

      showLegend: el.designerShowLegend ? el.designerShowLegend.checked : true,
      showLabel: el.designerShowLabel ? el.designerShowLabel.checked : false,
      showPercent: el.designerShowPercent ? el.designerShowPercent.checked : false,
      showTotal: true,
      enableExport: el.designerEnableExport ? el.designerEnableExport.checked : true,
      refreshMode: 'manual',

      groupBy: 'day',
      targetValue: el.designerTargetValue ? el.designerTargetValue.value : '',

      compare: collectDesignerCompareConfig(),
      layout: collectDesignerLayoutConfig()
    };
  }


  function collectDesignerCompareConfig() {
    return {
      enabled: el.designerCompareEnabled ? el.designerCompareEnabled.checked : false,
      type: el.designerCompareType ? el.designerCompareType.value : 'none',
      categoryField: el.designerCompareCategoryField ? el.designerCompareCategoryField.value : '',
      categoryA: el.designerCompareCategoryA ? el.designerCompareCategoryA.value.trim() : '',
      categoryB: el.designerCompareCategoryB ? el.designerCompareCategoryB.value.trim() : '',
      targetValue: el.designerCompareTargetValue ? el.designerCompareTargetValue.value : '',
      averageWindow: el.designerCompareAverageWindow ? el.designerCompareAverageWindow.value : '',
      dateField: el.designerDateField ? el.designerDateField.value : '',
      displayMode: 'summary'
    };
  }


  function collectDesignerLayoutConfig() {
    return {
      desktop: {
        x: 0,
        y: 0,
        w: el.designerDesktopW ? Number(el.designerDesktopW.value || 4) : 4,
        h: el.designerDesktopH ? Number(el.designerDesktopH.value || 3) : 3
      },
      tablet: {
        order: el.designerTabletOrder ? Number(el.designerTabletOrder.value || 1) : 1,
        w: 2,
        h: el.designerDesktopH ? Number(el.designerDesktopH.value || 3) : 3
      },
      mobile: {
        order: el.designerMobileOrder ? Number(el.designerMobileOrder.value || 1) : 1,
        display: 'full',
        hidden: el.designerMobileHidden ? el.designerMobileHidden.checked : false
      }
    };
  }


  function getDesignerDashboardId() {
    return el.designerDashboardSelect ? String(el.designerDashboardSelect.value || '').trim() : '';
  }


  function getDesignerSourceName() {
    if (!el.designerSourceSelect) {
      return '';
    }

    const opt = el.designerSourceSelect.options[el.designerSourceSelect.selectedIndex];

    return opt ? String(opt.textContent || '').trim() : '';
  }


  function getDefaultWidgetTitle(widgetType) {
    const map = {
      kpi: 'KPI Summary',
      line: 'Trend Chart',
      bar: 'Bar Chart',
      donut: 'Donut Chart',
      stacked_bar: 'Stacked Bar Chart',
      gauge: 'Gauge Chart',
      ranking: 'Ranking',
      table: 'Table',
      heatmap: 'Heatmap'
    };

    return map[widgetType] || 'Untitled Widget';
  }


  function validateDesignerWidgetConfig(config, mode) {
    if (!config.dashboardId) {
      return 'กรุณาเลือก Dashboard ที่ต้องการออกแบบ';
    }

    if (!config.sourceId || !config.sheetName) {
      return 'กรุณาเลือก Source และ Sheet';
    }

    if (!config.widgetType) {
      return 'กรุณาเลือก Widget Type';
    }

    if (config.widgetType === 'line' && !(config.dateField || config.xField)) {
      return 'กราฟเส้นต้องเลือก Date Field';
    }

    if (config.widgetType === 'bar' && !(config.xField || config.categoryField)) {
      return 'กราฟแท่งต้องเลือก X / Category Field';
    }

    if (config.widgetType === 'donut' && !(config.categoryField || config.xField)) {
      return 'Donut Chart ต้องเลือก Category Field';
    }

    if (config.widgetType === 'stacked_bar' && (!(config.xField || config.categoryField) || !config.stackField)) {
      return 'Stacked Bar ต้องเลือก X Field และ Stack Field';
    }

    if (config.widgetType === 'ranking' && !(config.categoryField || config.xField)) {
      return 'Ranking ต้องเลือก Field สำหรับจัดอันดับ';
    }

    if (config.widgetType === 'gauge' && !config.targetValue && !(config.compare && config.compare.targetValue)) {
      return 'Gauge ต้องกำหนด Target Value';
    }

    if (mode === 'save' && !config.dashboardId) {
      return 'ต้องเลือก Dashboard ก่อนบันทึก Widget';
    }

    return '';
  }


  async function handleDesignerRealPreview() {
    const config = collectDesignerWidgetConfig();
    const error = validateDesignerWidgetConfig(config, 'preview');

    if (error) {
      setDesignerMessage(error, 'error');
      return;
    }

    try {
      setDesignerMessage('กำลังสร้าง Preview จากข้อมูลจริง...', 'muted');
      setDesignerBusy(true);

      const result = await window.AnalyticsAPI.widgetPreview(config);

      designerCurrentPreviewData = result;

      renderDesignerRealPreview(result);

      setDesignerMessage('สร้าง Preview สำเร็จ', 'success');

      writeLog({
        step: 'designer_widget_preview',
        result: result
      });

    } catch (error) {
      setDesignerMessage(error.message || 'สร้าง Preview ไม่สำเร็จ', 'error');

      writeLog({
        step: 'designer_widget_preview_error',
        message: error.message,
        payload: error.payload || null
      });

    } finally {
      setDesignerBusy(false);
    }
  }


  function renderDesignerRealPreview(result) {
    if (!el.designerLivePreview) {
      return;
    }

    const preview = result.preview || {};
    const type = result.widgetType || preview.type || '';

    el.designerLivePreview.classList.remove('empty');

    if (type === 'kpi') {
      renderDesignerKpiPreview(preview);
      return;
    }

    if (type === 'bar') {
      renderDesignerChartPreview(preview, 'bar');
      return;
    }

    if (type === 'line') {
      renderDesignerChartPreview(preview, 'line');
      return;
    }

    if (type === 'donut') {
      renderDesignerChartPreview(preview, 'pie');
      return;
    }

    if (type === 'stacked_bar') {
      renderDesignerStackedBarPreview(preview);
      return;
    }

    if (type === 'ranking') {
      renderDesignerRankingPreview(preview);
      return;
    }

    if (type === 'table') {
      renderDesignerTablePreview(preview);
      return;
    }

    if (type === 'gauge') {
      renderDesignerGaugePreview(preview);
      return;
    }

    el.designerLivePreview.innerHTML = [
      '<div class="preview-widget-card">',
        '<strong>' + escapeHtml(result.title || 'Preview') + '</strong>',
        '<pre style="white-space:pre-wrap;word-break:break-word;">',
          escapeHtml(JSON.stringify(preview, null, 2)),
        '</pre>',
      '</div>'
    ].join('');
  }


  function renderDesignerKpiPreview(preview) {
    el.designerLivePreview.innerHTML = [
      '<div class="preview-kpi-card">',
        '<h5>' + escapeHtml(preview.title || 'KPI') + '</h5>',
        '<div class="preview-kpi-value">' + escapeHtml(preview.displayValue || preview.value || '0') + '</div>',
        '<div class="preview-kpi-unit">',
          escapeHtml(preview.aggregate ? 'Aggregate: ' + preview.aggregate : ''),
        '</div>',
      '</div>'
    ].join('');
  }


  function renderDesignerChartPreview(preview, chartType) {
    const chartId = 'designerPreviewChart_' + Date.now();

    el.designerLivePreview.innerHTML = [
      '<div class="preview-widget-card">',
        '<h5 style="margin:0 0 12px;">' + escapeHtml(preview.title || 'Chart Preview') + '</h5>',
        '<div id="' + chartId + '" class="preview-chart"></div>',
      '</div>'
    ].join('');

    const chartEl = document.getElementById(chartId);

    if (!chartEl || !window.echarts) {
      return;
    }

    const chart = echarts.init(chartEl);

    let option;

    if (chartType === 'line') {
      option = {
        tooltip: { trigger: 'axis' },
        legend: { show: true },
        grid: { left: 36, right: 16, top: 36, bottom: 36 },
        xAxis: {
          type: 'category',
          data: preview.labels || []
        },
        yAxis: {
          type: 'value'
        },
        series: preview.series || [
          {
            name: preview.title || 'Value',
            type: 'line',
            smooth: true,
            data: preview.data || []
          }
        ]
      };
    } else if (chartType === 'pie') {
      option = {
        tooltip: { trigger: 'item' },
        legend: { bottom: 0 },
        series: [
          {
            name: preview.title || 'Share',
            type: 'pie',
            radius: ['42%', '68%'],
            center: ['50%', '44%'],
            data: preview.data || []
          }
        ]
      };
    } else {
      const data = preview.data || [];

      option = {
        tooltip: { trigger: 'axis' },
        grid: { left: 42, right: 16, top: 28, bottom: 48 },
        xAxis: {
          type: 'category',
          data: data.map(function (item) {
            return item.name;
          }),
          axisLabel: {
            rotate: data.length > 5 ? 35 : 0
          }
        },
        yAxis: {
          type: 'value'
        },
        series: [
          {
            name: preview.title || 'Value',
            type: 'bar',
            data: data.map(function (item) {
              return item.value;
            })
          }
        ]
      };
    }

    chart.setOption(option);

    setTimeout(function () {
      chart.resize();
    }, 120);
  }


  function renderDesignerStackedBarPreview(preview) {
    const chartId = 'designerPreviewChart_' + Date.now();

    el.designerLivePreview.innerHTML = [
      '<div class="preview-widget-card">',
        '<h5 style="margin:0 0 12px;">' + escapeHtml(preview.title || 'Stacked Bar Preview') + '</h5>',
        '<div id="' + chartId + '" class="preview-chart"></div>',
      '</div>'
    ].join('');

    const chartEl = document.getElementById(chartId);

    if (!chartEl || !window.echarts) {
      return;
    }

    const chart = echarts.init(chartEl);

    chart.setOption({
      tooltip: { trigger: 'axis' },
      legend: { top: 0 },
      grid: { left: 42, right: 16, top: 48, bottom: 48 },
      xAxis: {
        type: 'category',
        data: preview.categories || []
      },
      yAxis: {
        type: 'value'
      },
      series: preview.series || []
    });

    setTimeout(function () {
      chart.resize();
    }, 120);
  }


  function renderDesignerRankingPreview(preview) {
    const data = preview.data || [];

    el.designerLivePreview.innerHTML = [
      '<div class="preview-widget-card">',
        '<h5 style="margin:0 0 12px;">' + escapeHtml(preview.title || 'Ranking') + '</h5>',
        '<div class="saved-widget-list" style="padding:0;background:transparent;">',
          data.map(function (item, index) {
            return [
              '<div class="saved-widget-card">',
                '<div class="saved-widget-top">',
                  '<strong>' + (index + 1) + '. ' + escapeHtml(item.name) + '</strong>',
                  '<span class="widget-type-pill">' + escapeHtml(formatDesignerNumber(item.value)) + '</span>',
                '</div>',
              '</div>'
            ].join('');
          }).join(''),
        '</div>',
      '</div>'
    ].join('');
  }


  function renderDesignerTablePreview(preview) {
    const columns = preview.columns || [];
    const rows = preview.rows || [];

    if (!columns.length) {
      el.designerLivePreview.innerHTML = '<div class="preview-widget-card">ไม่พบคอลัมน์สำหรับแสดงตาราง</div>';
      return;
    }

    el.designerLivePreview.innerHTML = [
      '<div class="preview-widget-card">',
        '<h5 style="margin:0 0 12px;">' + escapeHtml(preview.title || 'Table') + '</h5>',
        '<div style="overflow:auto;">',
          '<table class="preview-table">',
            '<thead>',
              '<tr>',
                columns.map(function (col) {
                  return '<th>' + escapeHtml(col) + '</th>';
                }).join(''),
              '</tr>',
            '</thead>',
            '<tbody>',
              rows.map(function (row) {
                return [
                  '<tr>',
                    columns.map(function (col) {
                      return '<td>' + escapeHtml(formatDesignerCell(row[col])) + '</td>';
                    }).join(''),
                  '</tr>'
                ].join('');
              }).join(''),
            '</tbody>',
          '</table>',
        '</div>',
      '</div>'
    ].join('');
  }


  function renderDesignerGaugePreview(preview) {
    const chartId = 'designerPreviewChart_' + Date.now();

    el.designerLivePreview.innerHTML = [
      '<div class="preview-widget-card">',
        '<h5 style="margin:0 0 12px;">' + escapeHtml(preview.title || 'Gauge') + '</h5>',
        '<div id="' + chartId + '" class="preview-chart"></div>',
      '</div>'
    ].join('');

    const chartEl = document.getElementById(chartId);

    if (!chartEl || !window.echarts) {
      return;
    }

    const chart = echarts.init(chartEl);

    chart.setOption({
      tooltip: {
        formatter: '{a}<br/>{b}: {c}%'
      },
      series: [
        {
          name: preview.title || 'Gauge',
          type: 'gauge',
          progress: { show: true },
          detail: {
            valueAnimation: true,
            formatter: '{value}%'
          },
          data: [
            {
              value: Number(preview.percent || 0),
              name: 'Progress'
            }
          ]
        }
      ]
    });

    setTimeout(function () {
      chart.resize();
    }, 120);
  }


  async function handleDesignerComparisonPreview() {
    const config = collectDesignerWidgetConfig();
    const compare = collectDesignerCompareConfig();

    if (!compare.enabled || compare.type === 'none') {
      renderDesignerComparisonResult({
        enabled: false,
        message: 'ยังไม่ได้เปิดการเปรียบเทียบ'
      });
      return;
    }

    try {
      setDesignerMessage('กำลังคำนวณ Comparison...', 'muted');

      const result = await window.AnalyticsAPI.comparisonPreview({
        widget: config,
        compare: compare
      });

      renderDesignerComparisonResult(result);

      setDesignerMessage('คำนวณ Comparison สำเร็จ', 'success');

    } catch (error) {
      setDesignerMessage(error.message || 'คำนวณ Comparison ไม่สำเร็จ', 'error');

      if (el.designerComparisonResult) {
        el.designerComparisonResult.className = 'comparison-result danger';
        el.designerComparisonResult.textContent = error.message || 'Comparison Error';
      }
    }
  }


  function renderDesignerComparisonResult(result) {
    if (!el.designerComparisonResult) {
      return;
    }

    if (!result || result.enabled === false) {
      el.designerComparisonResult.className = 'comparison-result empty';
      el.designerComparisonResult.textContent = result && result.message ? result.message : 'ยังไม่ได้เปิดการเปรียบเทียบ';
      return;
    }

    const data = result.result || result;

    el.designerComparisonResult.className = 'comparison-result success';

    if (data.type === 'target') {
      el.designerComparisonResult.innerHTML = [
        '<div><strong>ประเภท:</strong> Target</div>',
        '<div><strong>ค่าปัจจุบัน:</strong> ' + escapeHtml(formatDesignerNumber(data.current)) + '</div>',
        '<div><strong>Target:</strong> ' + escapeHtml(formatDesignerNumber(data.target)) + '</div>',
        '<div><strong>ส่วนต่าง:</strong> ' + escapeHtml(formatDesignerNumber(data.diff)) + '</div>',
        '<div><strong>% สำเร็จ:</strong> ' + escapeHtml(data.percent) + '%</div>'
      ].join('');
      return;
    }

    if (data.type === 'category_vs_category') {
      el.designerComparisonResult.innerHTML = [
        '<div><strong>ประเภท:</strong> Category A vs B</div>',
        '<div><strong>' + escapeHtml(data.categoryA) + ':</strong> ' + escapeHtml(formatDesignerNumber(data.valueA)) + '</div>',
        '<div><strong>' + escapeHtml(data.categoryB) + ':</strong> ' + escapeHtml(formatDesignerNumber(data.valueB)) + '</div>',
        '<div><strong>ส่วนต่าง:</strong> ' + escapeHtml(formatDesignerNumber(data.diff)) + '</div>',
        '<div><strong>% เปลี่ยนแปลง:</strong> ' + escapeHtml(data.percent) + '%</div>'
      ].join('');
      return;
    }

    el.designerComparisonResult.innerHTML = '<pre style="white-space:pre-wrap;">' +
      escapeHtml(JSON.stringify(data, null, 2)) +
      '</pre>';
  }


  async function handleDesignerSaveWidget() {
    const config = collectDesignerWidgetConfig();
    const error = validateDesignerWidgetConfig(config, 'save');

    if (error) {
      setDesignerMessage(error, 'error');
      return;
    }

    try {
      setDesignerMessage('กำลังบันทึก Widget...', 'muted');
      setDesignerBusy(true);

      const payload = {
        ...config,
        compare: config.compare,
        layout: config.layout
      };

      const result = designerEditingWidgetId
        ? await window.AnalyticsAPI.updateWidget(payload)
        : await window.AnalyticsAPI.saveWidget(payload);

      designerEditingWidgetId = result.widgetId || designerEditingWidgetId || '';

      setDesignerMessage(result.message || 'บันทึก Widget สำเร็จ', 'success');

      await handleDesignerLoadSavedWidgets();

      writeLog({
        step: 'designer_save_widget',
        result: result
      });

    } catch (error) {
      setDesignerMessage(error.message || 'บันทึก Widget ไม่สำเร็จ', 'error');

      writeLog({
        step: 'designer_save_widget_error',
        message: error.message,
        payload: error.payload || null
      });

    } finally {
      setDesignerBusy(false);
    }
  }


  async function handleDesignerLoadSavedWidgets() {
    const dashboardId = getDesignerDashboardId();

    if (!dashboardId) {
      setDesignerMessage('กรุณาเลือก Dashboard ก่อนโหลด Widget เดิม', 'error');
      return;
    }

    try {
      setDesignerMessage('กำลังโหลด Widget เดิม...', 'muted');

      const result = await window.AnalyticsAPI.dashboardDesignerLoad(dashboardId);

      designerSavedWidgetsCache = result.widgets || [];
      designerSavedComparisonsCache = result.comparisons || [];
      designerSavedLayoutsCache = result.layouts || [];

      renderDesignerSavedWidgets(designerSavedWidgetsCache);

      setDesignerMessage('โหลด Widget เดิมสำเร็จ', 'success');

      writeLog({
        step: 'designer_load_saved_widgets',
        result: result
      });

    } catch (error) {
      setDesignerMessage(error.message || 'โหลด Widget เดิมไม่สำเร็จ', 'error');

      writeLog({
        step: 'designer_load_saved_widgets_error',
        message: error.message,
        payload: error.payload || null
      });
    }
  }


  function renderDesignerSavedWidgets(widgets) {
    if (!el.designerSavedWidgetList) {
      return;
    }

    const list = Array.isArray(widgets) ? widgets : [];

    if (!list.length) {
      el.designerSavedWidgetList.classList.add('empty');
      el.designerSavedWidgetList.textContent = 'ยังไม่มี Widget ที่บันทึกไว้';
      return;
    }

    el.designerSavedWidgetList.classList.remove('empty');

    el.designerSavedWidgetList.innerHTML = list.map(function (item, index) {
      return [
        '<div class="saved-widget-card" data-saved-widget-index="' + index + '">',
          '<div class="saved-widget-top">',
            '<div>',
              '<strong>' + escapeHtml(item.title || item.widgetType || '') + '</strong>',
              '<small>' + escapeHtml(item.widgetId || '') + '</small>',
            '</div>',
            '<span class="widget-type-pill">' + escapeHtml(item.widgetType || '') + '</span>',
          '</div>',
          '<div class="saved-widget-meta">',
            '<span class="field-badge">Source: ' + escapeHtml(item.sourceName || item.sourceId || '-') + '</span>',
            '<span class="field-badge">Sheet: ' + escapeHtml(item.sheetName || '-') + '</span>',
            '<span class="field-badge">Aggregate: ' + escapeHtml(item.aggregate || '-') + '</span>',
          '</div>',
          '<div class="saved-widget-actions">',
            '<button class="btn btn-primary" type="button" data-edit-saved-widget="' + index + '">แก้ไข</button>',
            '<button class="btn btn-secondary" type="button" data-preview-saved-widget="' + index + '">Preview</button>',
            '<button class="btn btn-danger" type="button" data-delete-saved-widget="' + index + '">ลบ</button>',
          '</div>',
        '</div>'
      ].join('');
    }).join('');

    el.designerSavedWidgetList.querySelectorAll('[data-edit-saved-widget]').forEach(function (btn) {
      btn.addEventListener('click', function () {
        const index = Number(btn.getAttribute('data-edit-saved-widget'));
        applySavedWidgetToDesignerForm(index);
      });
    });

    el.designerSavedWidgetList.querySelectorAll('[data-preview-saved-widget]').forEach(function (btn) {
      btn.addEventListener('click', async function () {
        const index = Number(btn.getAttribute('data-preview-saved-widget'));
        await previewSavedWidget(index);
      });
    });

    el.designerSavedWidgetList.querySelectorAll('[data-delete-saved-widget]').forEach(function (btn) {
      btn.addEventListener('click', async function () {
        const index = Number(btn.getAttribute('data-delete-saved-widget'));
        await deleteSavedWidget(index);
      });
    });
  }


  function applySavedWidgetToDesignerForm(index) {
    const item = designerSavedWidgetsCache[index];

    if (!item) {
      return;
    }

    designerEditingWidgetId = item.widgetId || '';

    if (el.designerWidgetType) el.designerWidgetType.value = item.widgetType || '';
    if (el.designerWidgetTitle) el.designerWidgetTitle.value = item.title || '';
    if (el.designerDateField) el.designerDateField.value = item.dateField || '';
    if (el.designerXField) el.designerXField.value = item.xField || item.categoryField || '';
    if (el.designerValueField) el.designerValueField.value = item.valueField || '';
    if (el.designerStackField) el.designerStackField.value = item.stackField || item.groupField || '';
    if (el.designerAggregate) el.designerAggregate.value = item.aggregate || 'count';
    if (el.designerUnit) el.designerUnit.value = item.unit || '';
    if (el.designerSortMethod) el.designerSortMethod.value = item.sortMethod || 'desc';
    if (el.designerLimit) el.designerLimit.value = item.limit || '';
    if (el.designerTheme) el.designerTheme.value = item.theme || 'executive_blue';

    if (el.designerShowLegend) el.designerShowLegend.checked = !!item.showLegend;
    if (el.designerShowLabel) el.designerShowLabel.checked = !!item.showLabel;
    if (el.designerShowPercent) el.designerShowPercent.checked = !!item.showPercent;
    if (el.designerEnableExport) el.designerEnableExport.checked = item.enableExport !== false;

    applySavedComparisonToForm(item.widgetId);
    applySavedLayoutToForm(item.widgetId);

    if (el.designerUpdateWidgetBtn) {
      el.designerUpdateWidgetBtn.disabled = false;
    }

    renderDesignerMiniPreviewByType(item.widgetType || '');
    setDesignerMessage('โหลด Widget มาแก้ไขแล้ว', 'success');
    scrollToDesignerSection();
  }


  function applySavedComparisonToForm(widgetId) {
    const compare = designerSavedComparisonsCache.find(function (item) {
      return item.widgetId === widgetId;
    });

    if (!compare) {
      if (el.designerCompareEnabled) el.designerCompareEnabled.checked = false;
      if (el.designerCompareType) el.designerCompareType.value = 'none';
      return;
    }

    if (el.designerCompareEnabled) el.designerCompareEnabled.checked = !!compare.enabled;
    if (el.designerCompareType) el.designerCompareType.value = compare.type || 'none';
    if (el.designerCompareCategoryField) el.designerCompareCategoryField.value = compare.categoryField || '';
    if (el.designerCompareCategoryA) el.designerCompareCategoryA.value = compare.categoryA || '';
    if (el.designerCompareCategoryB) el.designerCompareCategoryB.value = compare.categoryB || '';
    if (el.designerCompareTargetValue) el.designerCompareTargetValue.value = compare.targetValue || '';
    if (el.designerCompareAverageWindow) el.designerCompareAverageWindow.value = compare.averageWindow || '';
  }


  function applySavedLayoutToForm(widgetId) {
    const layout = designerSavedLayoutsCache.find(function (item) {
      return item.widgetId === widgetId;
    });

    if (!layout) {
      return;
    }

    if (el.designerDesktopW) el.designerDesktopW.value = layout.desktop ? layout.desktop.w || 4 : 4;
    if (el.designerDesktopH) el.designerDesktopH.value = layout.desktop ? layout.desktop.h || 3 : 3;
    if (el.designerTabletOrder) el.designerTabletOrder.value = layout.tablet ? layout.tablet.order || 1 : 1;
    if (el.designerMobileOrder) el.designerMobileOrder.value = layout.mobile ? layout.mobile.order || 1 : 1;
    if (el.designerMobileHidden) el.designerMobileHidden.checked = layout.mobile ? !!layout.mobile.hidden : false;
  }


  async function previewSavedWidget(index) {
    const item = designerSavedWidgetsCache[index];

    if (!item) {
      return;
    }

    try {
      setDesignerMessage('กำลัง Preview Widget เดิม...', 'muted');

      const result = await window.AnalyticsAPI.widgetPreview(item);
      renderDesignerRealPreview(result);

      setDesignerMessage('Preview Widget เดิมสำเร็จ', 'success');

    } catch (error) {
      setDesignerMessage(error.message || 'Preview Widget เดิมไม่สำเร็จ', 'error');
    }
  }


  async function deleteSavedWidget(index) {
    const item = designerSavedWidgetsCache[index];

    if (!item || !item.widgetId) {
      return;
    }

    const ok = window.confirm('ต้องการลบ Widget นี้หรือไม่?\n' + (item.title || item.widgetId));

    if (!ok) {
      return;
    }

    try {
      setDesignerMessage('กำลังลบ Widget...', 'muted');

      await window.AnalyticsAPI.deleteWidget(item.widgetId);

      setDesignerMessage('ลบ Widget สำเร็จ', 'success');

      await handleDesignerLoadSavedWidgets();

    } catch (error) {
      setDesignerMessage(error.message || 'ลบ Widget ไม่สำเร็จ', 'error');
    }
  }


  function formatDesignerNumber(value) {
    const num = Number(value || 0);

    return num.toLocaleString('en-US', {
      maximumFractionDigits: 2
    });
  }


  function formatDesignerCell(value) {
    if (value instanceof Date) {
      return value.toLocaleString('th-TH');
    }

    if (value === null || value === undefined) {
      return '';
    }

    return String(value);
  }
  /**
 * =====================================================
 * Dashboard Builder Viewer
 * ใช้แสดง Widget ที่สร้างจาก Dashboard Designer
 * =====================================================
 */
let currentBuilderDashboard = null;
let builderChartInstances = [];
let currentBuilderFilterOptions = [];

let currentBuilderFilters = {
  dateField: '',
  dateFrom: '',
  dateTo: '',
  categoryFilters: [],
  keyword: ''
};
async function handleOpenDashboardSmart() {
  const dashboardId = el.dashboardSelect ? String(el.dashboardSelect.value || '').trim() : '';

  if (!dashboardId) {
    setDashboardViewMessage('กรุณาเลือก Dashboard ก่อน', 'error');
    return;
  }

  if (window.AnalyticsAPI && window.AnalyticsAPI.dashboardBuilderView) {
    try {
      await handleOpenDashboardBuilderView();
      return;
    } catch (error) {
      const msg = String(error && error.message ? error.message : '');

      /**
       * ถ้า Dashboard ยังไม่มี Widget จาก Designer
       * ให้ fallback ไปใช้ระบบ dashboardView เดิม
       */
      if (
        msg.includes('ยังไม่มี Widget') ||
        msg.includes('ยังไม่มี') ||
        msg.includes('ไม่พบ') ||
        msg.includes('endpoint')
      ) {
        writeLog({
          step: 'builder_view_fallback_to_legacy',
          message: msg,
          payload: error.payload || null
        });

        await handleOpenDashboard();
        return;
      }

      setDashboardViewMessage(msg || 'เปิด Dashboard Builder ไม่สำเร็จ', 'error');

      writeLog({
        step: 'builder_view_error',
        message: msg,
        payload: error.payload || null
      });

      return;
    }
  }

  await handleOpenDashboard();
}


async function handleOpenDashboardBuilderView() {
  const dashboardId = el.dashboardSelect ? String(el.dashboardSelect.value || '').trim() : '';
  const limit = el.dashboardLimit ? Number(el.dashboardLimit.value || 5000) : 5000;

  if (!dashboardId) {
    throw new Error('กรุณาเลือก Dashboard ก่อน');
  }

  setDashboardViewMessage('กำลังโหลด Dashboard Builder...', 'muted');
  showGlobalLoading('กำลังโหลด Dashboard Builder...');

  try {
    clearBuilderChartInstances();

    const result = await window.AnalyticsAPI.dashboardBuilderView({
  dashboardId: dashboardId,
  limit: limit,
  filters: currentBuilderFilters || {}
});

    if (!result.ok) {
      throw new Error(result.message || 'โหลด Dashboard Builder ไม่สำเร็จ');
    }

    currentBuilderDashboard = result;

    renderDashboardBuilderView(result);

    setDashboardViewMessage('เปิด Dashboard Builder สำเร็จ', 'success');

    writeLog({
      step: 'dashboard_builder_view',
      result: result
    });

  } finally {
    hideGlobalLoading();
  }
}


function renderDashboardBuilderView(result) {
  if (!el.dashboardViewResult) {
    return;
  }

  const widgets = Array.isArray(result.widgets) ? result.widgets : [];
  const theme = result.theme || {};

  if (!widgets.length) {
    el.dashboardViewResult.classList.add('empty');
    el.dashboardViewResult.textContent = 'Dashboard นี้ยังไม่มี Widget จาก Dashboard Designer';
    return;
  }

  applyDashboardBuilderTheme(theme);
  renderBuilderFilterBox(result);

  const visibleWidgets = widgets
    .filter(function (widget) {
      return widget && widget.ok !== false;
    })
    .sort(sortBuilderWidgetsByLayout);

  const errorWidgets = widgets.filter(function (widget) {
    return widget && widget.ok === false;
  });

  el.dashboardViewResult.classList.remove('empty');

  el.dashboardViewResult.innerHTML = [
    '<div class="builder-dashboard-shell">',
      renderBuilderDashboardHeader(result),
      '<div class="builder-dashboard-grid">',
        visibleWidgets.map(function (widget, index) {
          return renderBuilderWidgetShell(widget, index);
        }).join(''),
      '</div>',
      errorWidgets.length ? renderBuilderWidgetErrors(errorWidgets) : '',
    '</div>'
  ].join('');

  visibleWidgets.forEach(function (widget, index) {
    renderBuilderWidgetContent(widget, index);
  });

  setTimeout(function () {
    resizeBuilderCharts();
  }, 150);
}


function renderBuilderDashboardHeader(result) {
  const theme = result.theme || {};
  const widgets = Array.isArray(result.widgets) ? result.widgets : [];

  return [
    '<div class="builder-dashboard-header">',
      '<div>',
        '<h3>Dashboard Builder View</h3>',
        '<p>',
          'Dashboard ID: ', escapeHtml(result.dashboardId || '-'),
          ' · Widgets: ', escapeHtml(widgets.length),
          ' · Theme: ', escapeHtml(theme.themeName || theme.themeId || '-'),
        '</p>',
      '</div>',
      '<div class="builder-dashboard-badge">',
        escapeHtml(result.message || 'พร้อมใช้งาน'),
      '</div>',
    '</div>'
  ].join('');
}


function renderBuilderWidgetShell(widget, index) {
  const layout = widget.layout || {};
  const desktop = layout.desktop || {};
  const mobile = layout.mobile || {};

  const w = clampNumber(desktop.w || 4, 1, 12);
  const h = clampNumber(desktop.h || 3, 1, 12);
  const mobileHidden = mobile.hidden ? ' builder-widget-mobile-hidden' : '';

  return [
    '<article class="builder-widget builder-widget-', escapeHtml(widget.widgetType || 'unknown'), mobileHidden, '" ',
      'style="--builder-widget-span:', w, ';--builder-widget-height:', h, ';" ',
      'data-builder-widget-id="', escapeHtml(widget.widgetId || ''), '">',
      '<div class="builder-widget-header">',
        '<div>',
          '<h4>', escapeHtml(widget.title || 'Untitled Widget'), '</h4>',
          '<p>',
            escapeHtml(widget.widgetType || ''),
            widget.sourceName ? ' · ' + escapeHtml(widget.sourceName) : '',
            widget.sheetName ? ' · ' + escapeHtml(widget.sheetName) : '',
          '</p>',
        '</div>',
        '<span class="widget-type-pill">', escapeHtml(widget.widgetType || ''), '</span>',
      '</div>',
      '<div id="builderWidgetBody_', index, '" class="builder-widget-body">',
        '<div class="builder-widget-loading">กำลังแสดงผล...</div>',
      '</div>',
    '</article>'
  ].join('');
}


function renderBuilderWidgetContent(widget, index) {
  const body = document.getElementById('builderWidgetBody_' + index);

  if (!body) {
    return;
  }

  const type = widget.widgetType || '';
  const data = widget.data || {};
  const comparisonResult = widget.comparisonResult || null;

  if (type === 'kpi') {
    body.innerHTML = renderBuilderKpi(widget, data, comparisonResult);
    return;
  }

  if (type === 'ranking') {
    body.innerHTML = renderBuilderRanking(widget, data);
    return;
  }

  if (type === 'table') {
    body.innerHTML = renderBuilderTable(widget, data);
    return;
  }

  if (type === 'gauge') {
    body.innerHTML = '<div id="builderChart_' + index + '" class="builder-chart"></div>';
    renderBuilderGaugeChart('builderChart_' + index, widget, data);
    return;
  }

  if (type === 'bar' || type === 'line' || type === 'donut' || type === 'stacked_bar') {
    body.innerHTML = '<div id="builderChart_' + index + '" class="builder-chart"></div>';
    renderBuilderEChart('builderChart_' + index, widget, data);
    return;
  }

  body.innerHTML = [
    '<div class="builder-widget-empty">',
      'ยังไม่รองรับ Widget Type: ',
      escapeHtml(type || '-'),
    '</div>'
  ].join('');
}


function renderBuilderKpi(widget, data, comparisonResult) {
  return [
    '<div class="builder-kpi">',
      '<div class="builder-kpi-value">',
        escapeHtml(data.displayValue || formatBuilderNumber(data.value || 0)),
      '</div>',
      '<div class="builder-kpi-meta">',
        escapeHtml(data.aggregate ? 'Aggregate: ' + data.aggregate : ''),
      '</div>',
      comparisonResult ? renderBuilderComparisonSummary(comparisonResult) : '',
    '</div>'
  ].join('');
}


function renderBuilderComparisonSummary(comparisonResult) {
  const type = comparisonResult.type || '';

  if (type === 'target') {
    const statusClass = Number(comparisonResult.diff || 0) >= 0 ? 'success' : 'warning';

    return [
      '<div class="builder-compare ', statusClass, '">',
        '<strong>Target:</strong> ', escapeHtml(formatBuilderNumber(comparisonResult.target || 0)),
        ' · <strong>ทำได้:</strong> ', escapeHtml(comparisonResult.percent || 0), '%',
      '</div>'
    ].join('');
  }

  if (type === 'category_vs_category') {
    return [
      '<div class="builder-compare">',
        '<strong>', escapeHtml(comparisonResult.categoryA || 'A'), ':</strong> ',
        escapeHtml(formatBuilderNumber(comparisonResult.valueA || 0)),
        ' · <strong>', escapeHtml(comparisonResult.categoryB || 'B'), ':</strong> ',
        escapeHtml(formatBuilderNumber(comparisonResult.valueB || 0)),
        ' · <strong>ต่าง:</strong> ',
        escapeHtml(formatBuilderNumber(comparisonResult.diff || 0)),
      '</div>'
    ].join('');
  }

  return '';
}


function renderBuilderRanking(widget, data) {
  const rows = Array.isArray(data.data) ? data.data : [];

  if (!rows.length) {
    return '<div class="builder-widget-empty">ไม่มีข้อมูล Ranking</div>';
  }

  return [
    '<div class="builder-ranking">',
      rows.map(function (item, index) {
        return [
          '<div class="builder-ranking-row">',
            '<span class="builder-ranking-no">', index + 1, '</span>',
            '<span class="builder-ranking-name">', escapeHtml(item.name || '-'), '</span>',
            '<strong>', escapeHtml(formatBuilderNumber(item.value || 0)), '</strong>',
          '</div>'
        ].join('');
      }).join(''),
    '</div>'
  ].join('');
}


function renderBuilderTable(widget, data) {
  const columns = Array.isArray(data.columns) ? data.columns : [];
  const rows = Array.isArray(data.rows) ? data.rows : [];

  if (!columns.length) {
    return '<div class="builder-widget-empty">ไม่มีคอลัมน์สำหรับแสดงตาราง</div>';
  }

  return [
    '<div class="builder-table-wrap">',
      '<table class="builder-table">',
        '<thead>',
          '<tr>',
            columns.map(function (col) {
              return '<th>' + escapeHtml(col) + '</th>';
            }).join(''),
          '</tr>',
        '</thead>',
        '<tbody>',
          rows.map(function (row) {
            return [
              '<tr>',
                columns.map(function (col) {
                  return '<td>' + escapeHtml(formatBuilderCell(row[col])) + '</td>';
                }).join(''),
              '</tr>'
            ].join('');
          }).join(''),
        '</tbody>',
      '</table>',
    '</div>',
    '<div class="builder-table-count">แสดง ', escapeHtml(rows.length), ' จาก ', escapeHtml(data.totalRows || rows.length), ' รายการ</div>'
  ].join('');
}


function renderBuilderEChart(elementId, widget, data) {
  const chartEl = document.getElementById(elementId);

  if (!chartEl || !window.echarts) {
    return;
  }

  const chart = echarts.init(chartEl);
  const type = widget.widgetType || '';
  const colors = getBuilderChartColors();

  let option = {};

  if (type === 'line') {
    option = {
      color: colors,
      tooltip: {
        trigger: 'axis'
      },
      legend: {
        show: widget.config ? widget.config.showLegend !== false : true,
        top: 0
      },
      grid: {
        left: 42,
        right: 18,
        top: 48,
        bottom: 42
      },
      xAxis: {
        type: 'category',
        data: data.labels || [],
        axisLabel: {
          color: '#64748b'
        }
      },
      yAxis: {
        type: 'value',
        axisLabel: {
          color: '#64748b'
        },
        splitLine: {
          lineStyle: {
            type: 'dashed',
            color: '#e2e8f0'
          }
        }
      },
      series: (data.series || [
        {
          name: widget.title || 'Value',
          type: 'line',
          smooth: true,
          data: data.data || []
        }
      ]).map(function (series) {
        return {
          ...series,
          smooth: true,
          symbolSize: 7,
          lineStyle: {
            width: 3
          },
          areaStyle: {
            opacity: 0.08
          }
        };
      })
    };
  }

  if (type === 'bar') {
    const rows = Array.isArray(data.data) ? data.data : [];

    option = {
      color: colors,
      tooltip: {
        trigger: 'axis'
      },
      grid: {
        left: 42,
        right: 18,
        top: 32,
        bottom: 54
      },
      xAxis: {
        type: 'category',
        data: rows.map(function (item) {
          return item.name;
        }),
        axisLabel: {
          rotate: rows.length > 5 ? 35 : 0,
          color: '#64748b'
        }
      },
      yAxis: {
        type: 'value',
        axisLabel: {
          color: '#64748b'
        },
        splitLine: {
          lineStyle: {
            type: 'dashed',
            color: '#e2e8f0'
          }
        }
      },
      series: [
        {
          name: widget.title || 'Value',
          type: 'bar',
          barMaxWidth: 44,
          itemStyle: {
            borderRadius: [8, 8, 0, 0]
          },
          data: rows.map(function (item) {
            return item.value;
          })
        }
      ]
    };
  }

  if (type === 'donut') {
    option = {
      color: colors,
      tooltip: {
        trigger: 'item'
      },
      legend: {
        bottom: 0,
        type: 'scroll'
      },
      series: [
        {
          name: widget.title || 'Share',
          type: 'pie',
          radius: ['46%', '70%'],
          center: ['50%', '43%'],
          avoidLabelOverlap: true,
          itemStyle: {
            borderRadius: 8,
            borderColor: '#ffffff',
            borderWidth: 2
          },
          label: {
            formatter: '{b}\n{d}%'
          },
          data: data.data || []
        }
      ]
    };
  }

  if (type === 'stacked_bar') {
    option = {
      color: colors,
      tooltip: {
        trigger: 'axis'
      },
      legend: {
        top: 0,
        type: 'scroll'
      },
      grid: {
        left: 42,
        right: 18,
        top: 56,
        bottom: 54
      },
      xAxis: {
        type: 'category',
        data: data.categories || [],
        axisLabel: {
          color: '#64748b'
        }
      },
      yAxis: {
        type: 'value',
        axisLabel: {
          color: '#64748b'
        },
        splitLine: {
          lineStyle: {
            type: 'dashed',
            color: '#e2e8f0'
          }
        }
      },
      series: data.series || []
    };
  }

  chart.setOption(option);
  builderChartInstances.push(chart);
}


function renderBuilderGaugeChart(elementId, widget, data) {
  const chartEl = document.getElementById(elementId);

  if (!chartEl || !window.echarts) {
    return;
  }

  const chart = echarts.init(chartEl);
  const colors = getBuilderChartColors();

  chart.setOption({
    color: colors,
    tooltip: {
      formatter: '{a}<br/>{b}: {c}%'
    },
    series: [
      {
        name: widget.title || 'Gauge',
        type: 'gauge',
        min: 0,
        max: 100,
        progress: {
          show: true,
          width: 14
        },
        axisLine: {
          lineStyle: {
            width: 14,
            color: [
              [0.6, colors[4] || '#dc2626'],
              [0.85, colors[3] || '#f59e0b'],
              [1, colors[2] || '#16a34a']
            ]
          }
        },
        axisTick: {
          show: false
        },
        splitLine: {
          length: 10,
          lineStyle: {
            width: 2,
            color: '#94a3b8'
          }
        },
        axisLabel: {
          color: '#64748b'
        },
        detail: {
          valueAnimation: true,
          formatter: '{value}%',
          fontSize: 28,
          fontWeight: 800,
          color: colors[0] || '#2563eb'
        },
        data: [
          {
            value: Number(data.percent || 0),
            name: 'Progress'
          }
        ]
      }
    ]
  });

  builderChartInstances.push(chart);
}


function renderBuilderWidgetErrors(errors) {
  return [
    '<div class="builder-error-panel">',
      '<h4>Widget ที่แสดงผลไม่สำเร็จ</h4>',
      errors.map(function (widget) {
        return [
          '<div class="builder-error-item">',
            '<strong>', escapeHtml(widget.title || widget.widgetId || '-'), '</strong>',
            '<span>', escapeHtml(widget.message || 'ไม่ทราบสาเหตุ'), '</span>',
          '</div>'
        ].join('');
      }).join(''),
    '</div>'
  ].join('');
}


function sortBuilderWidgetsByLayout(a, b) {
  const la = a.layout || {};
  const lb = b.layout || {};

  const ma = la.mobile || {};
  const mb = lb.mobile || {};

  const da = la.desktop || {};
  const db = lb.desktop || {};

  const orderA = Number(ma.order || da.y || 0);
  const orderB = Number(mb.order || db.y || 0);

  if (orderA !== orderB) {
    return orderA - orderB;
  }

  return Number(da.x || 0) - Number(db.x || 0);
}


function applyDashboardBuilderTheme(theme) {
  if (!el.dashboardViewResult || !theme) {
    return;
  }

  const target = el.dashboardViewResult;

  target.style.setProperty('--builder-bg', theme.backgroundColor || '#f8fafc');
  target.style.setProperty('--builder-card', theme.cardColor || '#ffffff');
  target.style.setProperty('--builder-primary', theme.primaryColor || '#2563eb');
  target.style.setProperty('--builder-secondary', theme.secondaryColor || '#38bdf8');
  target.style.setProperty('--builder-success', theme.successColor || '#16a34a');
  target.style.setProperty('--builder-warning', theme.warningColor || '#f59e0b');
  target.style.setProperty('--builder-danger', theme.dangerColor || '#dc2626');
  target.style.setProperty('--builder-text', theme.textColor || '#0f172a');
  target.style.setProperty('--builder-muted', theme.mutedTextColor || '#64748b');
  target.style.setProperty('--builder-border', theme.borderColor || '#e2e8f0');
}


function clearBuilderChartInstances() {
  builderChartInstances.forEach(function (chart) {
    try {
      chart.dispose();
    } catch (err) {}
  });

  builderChartInstances = [];
}


function resizeBuilderCharts() {
  builderChartInstances.forEach(function (chart) {
    try {
      chart.resize();
    } catch (err) {}
  });
}


function clampNumber(value, min, max) {
  const n = Number(value || 0);

  if (n < min) return min;
  if (n > max) return max;

  return n;
}


function formatBuilderNumber(value) {
  const num = Number(value || 0);

  return num.toLocaleString('en-US', {
    maximumFractionDigits: 2
  });
}


function formatBuilderCell(value) {
  if (value instanceof Date) {
    return value.toLocaleString('th-TH');
  }

  if (value === null || value === undefined) {
    return '';
  }

  return String(value);
}
  /**
 * =====================================================
 * Dashboard Builder Viewer - Round 2
 * Theme Palette / Refresh / Export
 * =====================================================
 */

async function handleRefreshDashboardSmart() {
  if (currentBuilderDashboard && currentBuilderDashboard.dashboardId) {
    try {
      await handleOpenDashboardBuilderView();
      return;
    } catch (error) {
      setDashboardViewMessage(error.message || 'Refresh Dashboard Builder ไม่สำเร็จ', 'error');
      return;
    }
  }

  await handleRefreshDashboard();
}


async function handleExportDashboardSmart() {
  if (currentBuilderDashboard && currentBuilderDashboard.widgets) {
    exportBuilderDashboardCsv(currentBuilderDashboard);
    return;
  }

  await handleExportDashboard();
}


function applyDashboardBuilderTheme(theme) {
  if (!el.dashboardViewResult || !theme) {
    return;
  }

  currentBuilderTheme = theme;
  currentBuilderPalette = Array.isArray(theme.chartPalette) && theme.chartPalette.length
    ? theme.chartPalette
    : [
        theme.primaryColor || '#2563eb',
        theme.secondaryColor || '#38bdf8',
        theme.successColor || '#16a34a',
        theme.warningColor || '#f59e0b',
        theme.dangerColor || '#dc2626'
      ];

  const target = el.dashboardViewResult;

  target.style.setProperty('--builder-bg', theme.backgroundColor || '#f8fafc');
  target.style.setProperty('--builder-card', theme.cardColor || '#ffffff');
  target.style.setProperty('--builder-primary', theme.primaryColor || '#2563eb');
  target.style.setProperty('--builder-secondary', theme.secondaryColor || '#38bdf8');
  target.style.setProperty('--builder-success', theme.successColor || '#16a34a');
  target.style.setProperty('--builder-warning', theme.warningColor || '#f59e0b');
  target.style.setProperty('--builder-danger', theme.dangerColor || '#dc2626');
  target.style.setProperty('--builder-text', theme.textColor || '#0f172a');
  target.style.setProperty('--builder-muted', theme.mutedTextColor || '#64748b');
  target.style.setProperty('--builder-border', theme.borderColor || '#e2e8f0');
}


function getBuilderChartColors() {
  if (currentBuilderPalette && currentBuilderPalette.length) {
    return currentBuilderPalette;
  }

  return ['#2563eb', '#38bdf8', '#16a34a', '#f59e0b', '#dc2626'];
}


function renderBuilderDashboardHeader(result) {
  const theme = result.theme || {};
  const widgets = Array.isArray(result.widgets) ? result.widgets : [];

  const okCount = widgets.filter(function (widget) {
    return widget && widget.ok !== false;
  }).length;

  const errorCount = widgets.filter(function (widget) {
    return widget && widget.ok === false;
  }).length;

  const totalRowsRead = widgets.reduce(function (sum, widget) {
    return sum + Number(widget.rowsRead || 0);
  }, 0);

  return [
    '<div class="builder-dashboard-header">',
      '<div>',
        '<h3>Dashboard Builder View</h3>',
        '<p>',
          'Dashboard ID: ', escapeHtml(result.dashboardId || '-'),
          ' · Widgets: ', escapeHtml(widgets.length),
          ' · Theme: ', escapeHtml(theme.themeName || theme.themeId || '-'),
        '</p>',
      '</div>',

      '<div class="builder-header-stats">',
        '<div class="builder-stat-chip">',
          '<strong>', escapeHtml(okCount), '</strong>',
          '<span>Widget ใช้งานได้</span>',
        '</div>',
        '<div class="builder-stat-chip">',
          '<strong>', escapeHtml(errorCount), '</strong>',
          '<span>Widget Error</span>',
        '</div>',
        '<div class="builder-stat-chip">',
          '<strong>', escapeHtml(formatBuilderNumber(totalRowsRead)), '</strong>',
          '<span>Rows Read</span>',
        '</div>',
      '</div>',
    '</div>'
  ].join('');
}
  function exportBuilderDashboardCsv(result) {
  const widgets = Array.isArray(result.widgets) ? result.widgets : [];

  if (!widgets.length) {
    setDashboardViewMessage('ไม่มี Widget สำหรับ Export', 'error');
    return;
  }

  const lines = [];

  lines.push([
    'Dashboard ID',
    'Widget ID',
    'Widget Type',
    'Widget Title',
    'Name',
    'Value'
  ].join(','));

  widgets.forEach(function (widget) {
    if (!widget || widget.ok === false) {
      return;
    }

    const data = widget.data || {};
    const type = widget.widgetType || '';

    if (type === 'kpi' || type === 'gauge') {
      lines.push(csvJoin([
        result.dashboardId,
        widget.widgetId,
        type,
        widget.title,
        data.title || widget.title,
        data.displayValue || data.value || data.percent || ''
      ]));
      return;
    }

    if (type === 'bar' || type === 'donut' || type === 'ranking') {
      const rows = Array.isArray(data.data) ? data.data : [];

      rows.forEach(function (item) {
        lines.push(csvJoin([
          result.dashboardId,
          widget.widgetId,
          type,
          widget.title,
          item.name || '',
          item.value || ''
        ]));
      });

      return;
    }

    if (type === 'line') {
      const labels = data.labels || [];
      const values = data.data || [];

      labels.forEach(function (label, index) {
        lines.push(csvJoin([
          result.dashboardId,
          widget.widgetId,
          type,
          widget.title,
          label,
          values[index] || 0
        ]));
      });

      return;
    }

    if (type === 'stacked_bar') {
      const categories = data.categories || [];
      const series = data.series || [];

      series.forEach(function (s) {
        categories.forEach(function (category, index) {
          lines.push(csvJoin([
            result.dashboardId,
            widget.widgetId,
            type,
            widget.title + ' - ' + s.name,
            category,
            s.data ? s.data[index] || 0 : 0
          ]));
        });
      });

      return;
    }

    if (type === 'table') {
      const columns = data.columns || [];
      const rows = data.rows || [];

      lines.push('');
      lines.push('Table Widget: ' + csvEscape(widget.title || widget.widgetId));
      lines.push(columns.map(csvEscape).join(','));

      rows.forEach(function (row) {
        lines.push(columns.map(function (col) {
          return csvEscape(row[col]);
        }).join(','));
      });
    }
  });

  const csv = '\ufeff' + lines.join('\n');
  const blob = new Blob([csv], {
    type: 'text/csv;charset=utf-8;'
  });

  const fileName = 'dashboard-builder-' + (result.dashboardId || 'export') + '-' + Date.now() + '.csv';

  downloadBlob(blob, fileName);
  setDashboardViewMessage('Export CSV สำเร็จ', 'success');
}


function csvJoin(values) {
  return values.map(csvEscape).join(',');
}


function csvEscape(value) {
  const text = String(value === undefined || value === null ? '' : value);

  if (/[",\n\r]/.test(text)) {
    return '"' + text.replace(/"/g, '""') + '"';
  }

  return text;
}


function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');

  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();

  setTimeout(function () {
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, 0);
}
  /**
 * =====================================================
 * Dashboard Builder Filters
 * =====================================================
 */

function renderBuilderFilterBox(result) {
  currentBuilderFilterOptions = Array.isArray(result.filterOptions)
  ? result.filterOptions
  : [];
  
  if (!el.dashboardFilterBox) {
    return;
  }

  const fields = collectBuilderFilterFields(result);
  const dateFields = fields.dateFields;
  const categoryFields = fields.categoryFields;

  el.dashboardFilterBox.classList.remove('empty');

  el.dashboardFilterBox.innerHTML = [
    '<div class="builder-filter-panel">',
      '<div class="builder-filter-head">',
        '<div>',
          '<h4>ตัวกรอง Dashboard Builder</h4>',
          '<p>เลือกช่วงวันที่ หมวดหมู่ หรือค้นหาข้อมูล เพื่อคำนวณ Widget ใหม่จากข้อมูลที่กรองแล้ว</p>',
        '</div>',
        '<button id="builderClearFiltersBtn" class="btn btn-ghost" type="button">ล้าง Filter</button>',
      '</div>',

      '<div class="builder-filter-grid">',
        '<label class="field">',
          '<span>Date Field</span>',
          '<select id="builderFilterDateField">',
            '<option value="">ไม่ใช้วันที่</option>',
            dateFields.map(function (field) {
              return '<option value="' + escapeHtml(field) + '">' + escapeHtml(field) + '</option>';
            }).join(''),
          '</select>',
        '</label>',

        '<label class="field">',
          '<span>ตั้งแต่วันที่</span>',
          '<input id="builderFilterDateFrom" type="date">',
        '</label>',

        '<label class="field">',
          '<span>ถึงวันที่</span>',
          '<input id="builderFilterDateTo" type="date">',
        '</label>',

        '<label class="field">',
          '<span>ค้นหา Keyword</span>',
          '<input id="builderFilterKeyword" type="text" placeholder="ค้นหาทุกคอลัมน์">',
        '</label>',
      '</div>',

    '<div class="builder-category-filter-box">',
  '<div class="builder-category-filter-title">Category Filters</div>',
  (
    categoryFields.length
      ? categoryFields.map(renderBuilderCategoryFilterControlV2).join('')
      : '<div class="builder-filter-empty">ไม่พบ Field ที่เหมาะกับ Category Filter</div>'
  ),
'</div>',

      '<div class="actions-row builder-filter-actions">',
        '<button id="builderApplyFiltersBtn" class="btn btn-primary" type="button">ใช้ Filter</button>',
        '<span id="builderFilterSummary" class="mapping-message"></span>',
      '</div>',
    '</div>'
  ].join('');

  setBuilderFilterValuesToUi();
  bindBuilderFilterEvents();
}


function collectBuilderFilterFields(result) {
  const widgets = Array.isArray(result.widgets) ? result.widgets : [];
  const dateMap = {};
  const categoryMap = {};

  widgets.forEach(function (widget) {
    if (!widget || !widget.config) {
      return;
    }

    const config = widget.config || {};

    if (config.dateField) {
      dateMap[config.dateField] = true;
    }

    if (config.xField && config.widgetType !== 'line') {
      categoryMap[config.xField] = true;
    }

    if (config.categoryField) {
      categoryMap[config.categoryField] = true;
    }

    if (config.stackField) {
      categoryMap[config.stackField] = true;
    }
  });

  return {
    dateFields: Object.keys(dateMap),
    categoryFields: Object.keys(categoryMap)
  };
}


function renderBuilderCategoryFilterControl(field) {
  return [
    '<label class="builder-category-filter-item">',
      '<span>', escapeHtml(field), '</span>',
      '<input ',
        'type="text" ',
        'data-builder-category-filter="', escapeHtml(field), '" ',
        'placeholder="เช่น 906,903 หรือ Open,Closed"',
      '>',
    '</label>'
  ].join('');
}


function bindBuilderFilterEvents() {
  const applyBtn = document.getElementById('builderApplyFiltersBtn');
  const clearBtn = document.getElementById('builderClearFiltersBtn');

  if (applyBtn) {
    applyBtn.addEventListener('click', async function () {
      currentBuilderFilters = collectBuilderFiltersFromUi();
      setBuilderFilterSummary('กำลังใช้ Filter...');
      await handleOpenDashboardBuilderView();
    });
  }

  if (clearBtn) {
    clearBtn.addEventListener('click', async function () {
      currentBuilderFilters = {
        dateField: '',
        dateFrom: '',
        dateTo: '',
        categoryFilters: [],
        keyword: ''
      };

      setBuilderFilterSummary('ล้าง Filter แล้ว');
      await handleOpenDashboardBuilderView();
    });
  }
}

function collectBuilderFiltersFromUi() {
  const dateFieldEl = document.getElementById('builderFilterDateField');
  const dateFromEl = document.getElementById('builderFilterDateFrom');
  const dateToEl = document.getElementById('builderFilterDateTo');
  const keywordEl = document.getElementById('builderFilterKeyword');

  const categoryFilterMap = {};

  document.querySelectorAll('[data-builder-filter-checkbox]').forEach(function (input) {
    if (!input.checked) {
      return;
    }

    const field = input.getAttribute('data-builder-filter-checkbox') || '';
    const value = String(input.value || '').trim();

    if (!field || !value) {
      return;
    }

    if (!categoryFilterMap[field]) {
      categoryFilterMap[field] = [];
    }

    categoryFilterMap[field].push(value);
  });

  document.querySelectorAll('[data-builder-category-filter]').forEach(function (input) {
    const field = input.getAttribute('data-builder-category-filter') || '';
    const raw = String(input.value || '').trim();

    if (!field || !raw) {
      return;
    }

    const values = raw
      .split(',')
      .map(function (v) {
        return String(v || '').trim();
      })
      .filter(function (v) {
        return !!v;
      });

    if (!values.length) {
      return;
    }

    if (!categoryFilterMap[field]) {
      categoryFilterMap[field] = [];
    }

    values.forEach(function (value) {
      if (categoryFilterMap[field].indexOf(value) < 0) {
        categoryFilterMap[field].push(value);
      }
    });
  });

  const categoryFilters = Object.keys(categoryFilterMap).map(function (field) {
    return {
      field: field,
      values: categoryFilterMap[field]
    };
  });

  return {
    dateField: dateFieldEl ? dateFieldEl.value : '',
    dateFrom: dateFromEl ? dateFromEl.value : '',
    dateTo: dateToEl ? dateToEl.value : '',
    keyword: keywordEl ? keywordEl.value.trim() : '',
    categoryFilters: categoryFilters
  };
}

function setBuilderFilterValuesToUi() {
  const filters = currentBuilderFilters || {};

  const dateFieldEl = document.getElementById('builderFilterDateField');
  const dateFromEl = document.getElementById('builderFilterDateFrom');
  const dateToEl = document.getElementById('builderFilterDateTo');
  const keywordEl = document.getElementById('builderFilterKeyword');

  if (dateFieldEl) dateFieldEl.value = filters.dateField || '';
  if (dateFromEl) dateFromEl.value = filters.dateFrom || '';
  if (dateToEl) dateToEl.value = filters.dateTo || '';
  if (keywordEl) keywordEl.value = filters.keyword || '';

  const categoryFilters = Array.isArray(filters.categoryFilters)
    ? filters.categoryFilters
    : [];

  categoryFilters.forEach(function (filter) {
    const values = Array.isArray(filter.values)
      ? filter.values.map(String)
      : [];

    document.querySelectorAll('[data-builder-filter-checkbox="' + cssEscapeSafe(filter.field) + '"]').forEach(function (checkbox) {
      checkbox.checked = values.indexOf(String(checkbox.value)) >= 0;
    });

    const input = document.querySelector('[data-builder-category-filter="' + cssEscapeSafe(filter.field) + '"]');

    if (input) {
      const uncheckedValues = values.filter(function (value) {
        const found = Array.from(
          document.querySelectorAll('[data-builder-filter-checkbox="' + cssEscapeSafe(filter.field) + '"]')
        ).some(function (checkbox) {
          return String(checkbox.value) === String(value);
        });

        return !found;
      });

      input.value = uncheckedValues.join(',');
    }
  });

  updateBuilderFilterSummary();
}

function setBuilderFilterSummary(text) {
  const box = document.getElementById('builderFilterSummary');

  if (box) {
    box.textContent = text || '';
  }
}


function updateBuilderFilterSummary() {
  const filters = currentBuilderFilters || {};
  const parts = [];

  if (filters.dateField && (filters.dateFrom || filters.dateTo)) {
    parts.push('วันที่: ' + filters.dateField);
  }

  if (filters.keyword) {
    parts.push('ค้นหา: ' + filters.keyword);
  }

  if (Array.isArray(filters.categoryFilters) && filters.categoryFilters.length) {
    parts.push('Category: ' + filters.categoryFilters.length + ' เงื่อนไข');
  }

  setBuilderFilterSummary(parts.length ? parts.join(' · ') : 'ยังไม่ได้ใช้ Filter');
}


function cssEscapeSafe(value) {
  if (window.CSS && CSS.escape) {
    return CSS.escape(String(value || ''));
  }

  return String(value || '').replace(/"/g, '\\"');
}
  function renderBuilderCategoryFilterControlV2(field) {
  const optionGroup = currentBuilderFilterOptions.find(function (item) {
    return item.field === field;
  });

  const options = optionGroup && Array.isArray(optionGroup.options)
    ? optionGroup.options
    : [];

  const topOptions = options.slice(0, 20);

  return [
    '<div class="builder-category-filter-group" data-builder-filter-group="' + escapeHtml(field) + '">',
      '<div class="builder-category-filter-group-head">',
        '<strong>' + escapeHtml(field) + '</strong>',
        '<span>' + escapeHtml(options.length) + ' ค่า</span>',
      '</div>',

      topOptions.length
        ? '<div class="builder-filter-option-list">' +
            topOptions.map(function (item) {
              return renderBuilderFilterOptionCheckbox(field, item);
            }).join('') +
          '</div>'
        : '<div class="builder-filter-empty">ไม่มีค่าตัวอย่าง</div>',

      '<label class="builder-category-filter-item manual">',
        '<span>ระบุเอง</span>',
        '<input ',
          'type="text" ',
          'data-builder-category-filter="' + escapeHtml(field) + '" ',
          'placeholder="เช่น 906,903 หรือ Open,Closed"',
        '>',
      '</label>',
    '</div>'
  ].join('');
}


function renderBuilderFilterOptionCheckbox(field, item) {
  const id = 'bf_' + safeId(field) + '_' + safeId(item.value);

  return [
    '<label class="builder-filter-option" for="' + escapeHtml(id) + '">',
      '<input ',
        'id="' + escapeHtml(id) + '" ',
        'type="checkbox" ',
        'data-builder-filter-checkbox="' + escapeHtml(field) + '" ',
        'value="' + escapeAttr(item.value) + '"',
      '>',
      '<span class="builder-filter-option-text">' + escapeHtml(item.value) + '</span>',
      '<span class="builder-filter-option-count">' + escapeHtml(formatBuilderNumber(item.count || 0)) + '</span>',
    '</label>'
  ].join('');
}


function safeId(value) {
  return String(value || '')
    .replace(/[^a-zA-Z0-9ก-๙_-]/g, '_')
    .substring(0, 60);
}
})();
