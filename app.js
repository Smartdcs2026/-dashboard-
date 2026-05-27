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

  const DASHBOARD_BUILDER_STORAGE_KEY = 'analytics_dashboard_builder_widgets';

  const DASHBOARD_WIDGET_CATALOG = [
    {
      id: 'kpi_total',
      type: 'metric',
      subType: 'summary',
      title: 'KPI Summary Card',
      shortTitle: 'KPI Card',
      description: 'แสดงยอดรวม จำนวนรายการ ค่าเฉลี่ย หรือค่าหลักของระบบ',
      bestFor: 'ข้อมูลตัวเลข / จำนวนรายการ / KPI หลัก',
      icon: 'KPI',
      recommended: true
    },
    {
      id: 'kpi_trend',
      type: 'metric',
      subType: 'trend',
      title: 'KPI Trend Card',
      shortTitle: 'Trend Card',
      description: 'แสดงตัวเลขพร้อมแนวโน้มเพิ่มขึ้นหรือลดลง',
      bestFor: 'ข้อมูลที่มีวันที่ เช่น รายวัน รายเดือน',
      icon: 'TRD',
      recommended: true
    },
    {
      id: 'kpi_target',
      type: 'metric',
      subType: 'target',
      title: 'KPI Target Card',
      shortTitle: 'Target Card',
      description: 'แสดงผลเทียบเป้าหมายพร้อม Progress Bar',
      bestFor: 'งานที่มีเป้าหมาย เช่น SLA / Performance / Target',
      icon: 'TGT',
      recommended: false
    },
    {
      id: 'bar_chart',
      type: 'chart',
      chartType: 'bar',
      title: 'Bar Chart',
      shortTitle: 'Bar Chart',
      description: 'เปรียบเทียบจำนวนหรือผลรวมแยกตามหมวดหมู่',
      bestFor: 'DC / สาขา / ประเภท / สถานะ',
      icon: 'BAR',
      recommended: true
    },
    {
      id: 'horizontal_bar',
      type: 'chart',
      chartType: 'horizontal_bar',
      title: 'Horizontal Bar',
      shortTitle: 'Horizontal Bar',
      description: 'เหมาะกับชื่อหมวดหมู่ยาว หรือ Top Ranking',
      bestFor: 'Top 10 สาขา / พนักงาน / Location',
      icon: 'H-B',
      recommended: false
    },
    {
      id: 'line_chart',
      type: 'chart',
      chartType: 'line',
      title: 'Line Chart',
      shortTitle: 'Line Chart',
      description: 'ดูแนวโน้มข้อมูลตามช่วงเวลา',
      bestFor: 'Timestamp / วันที่ตรวจ / วันที่บันทึก',
      icon: 'LIN',
      recommended: true
    },
    {
      id: 'area_chart',
      type: 'chart',
      chartType: 'area',
      title: 'Area Chart',
      shortTitle: 'Area Chart',
      description: 'ดูแนวโน้มพร้อมน้ำหนักของปริมาณข้อมูล',
      bestFor: 'ยอดสะสม / ปริมาณงานตามเวลา',
      icon: 'ARE',
      recommended: false
    },
    {
      id: 'donut_chart',
      type: 'chart',
      chartType: 'donut',
      title: 'Donut Chart',
      shortTitle: 'Donut Chart',
      description: 'แสดงสัดส่วนแบบอ่านง่ายและดูเป็นมืออาชีพ',
      bestFor: 'สถานะ / ประเภท / แบรนด์ / กลุ่มข้อมูล',
      icon: 'DON',
      recommended: true
    },
    {
      id: 'pie_chart',
      type: 'chart',
      chartType: 'pie',
      title: 'Pie Chart',
      shortTitle: 'Pie Chart',
      description: 'แสดงสัดส่วนข้อมูลแบบพื้นฐาน',
      bestFor: 'ข้อมูลไม่เกิน 5-7 กลุ่ม',
      icon: 'PIE',
      recommended: false
    },
    {
      id: 'table_detail',
      type: 'table',
      subType: 'detail',
      title: 'Data Table',
      shortTitle: 'Table',
      description: 'แสดงรายการข้อมูลจริง พร้อม Pagination',
      bestFor: 'ตรวจสอบรายการละเอียด / Export',
      icon: 'TBL',
      recommended: true
    },
    {
      id: 'table_ranking',
      type: 'table',
      subType: 'ranking',
      title: 'Ranking Table',
      shortTitle: 'Ranking',
      description: 'แสดงอันดับ Top 5 / Top 10 จากข้อมูล',
      bestFor: 'อันดับ DC / สาขา / ผู้รับผิดชอบ',
      icon: 'TOP',
      recommended: false
    }
  ];

  let dashboardBuilderState = {
    selectedWidgetIds: [],
    lastSelectedAt: '',
    mode: 'auto'
  };

  document.addEventListener('DOMContentLoaded', init);

  async function init() {
    bindEvents();
    showLogin();
    resetWorkingState();
    initDashboardBuilderDesigner();

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
      showToast('warning', 'กรอกข้อมูลไม่ครบ', 'กรุณากรอกชื่อผู้ใช้และรหัสผ่าน');
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

      showToast('success', 'เข้าสู่ระบบสำเร็จ', currentUser ? (currentUser.displayName || currentUser.username || '') : '');

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
      showToast('error', 'เข้าสู่ระบบไม่สำเร็จ', error.message || '');

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
    showToast('info', 'ออกจากระบบแล้ว', '');
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
      showToast('warning', 'ข้อมูลไม่ครบ', 'กรุณากรอกชื่อแหล่งข้อมูลและ Google Sheet ID');
      return;
    }

    setSourceFormMessage('');
    setButtonLoading(el.createSourceBtn, true, 'กำลังเพิ่ม...');

    try {
      const data = await window.AnalyticsAPI.createSource(payload);

      setSourceFormMessage('เพิ่มแหล่งข้อมูลสำเร็จ');
      showToast('success', 'เพิ่มแหล่งข้อมูลสำเร็จ', payload.name);

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
      showToast('error', 'เพิ่มแหล่งข้อมูลไม่สำเร็จ', error.message || '');

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

      showToast('error', 'โหลด Source ไม่สำเร็จ', message);
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
      const sourceId = String(source['รหัสแหล่งข้อมูล'] || '');
      const name = String(source['ชื่อแหล่งข้อมูล'] || '-');
      const sheetId = String(source['Google Sheet ID'] || '-');
      const type = String(source['ประเภทข้อมูล'] || '-');
      const owner = String(source['เจ้าของข้อมูล'] || '-');
      const active = toBool(source['สถานะใช้งาน']);

      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'source-item' + (sourceId === selectedSourceId ? ' active' : '');
      btn.setAttribute('title', name + '\n' + sheetId);

      btn.innerHTML = `
        <div class="item-title">
          <span>${escapeHtml(name)}</span>
          <span class="badge ${active ? 'badge-success' : 'badge-muted'}">
            ${active ? 'Active' : 'Off'}
          </span>
        </div>

        <div class="item-meta">
          <div>ประเภท: ${escapeHtml(type)}</div>
          <div>เจ้าของ: ${escapeHtml(owner || '-')}</div>
          <div>ID: ${escapeHtml(compactText(sourceId, 18))}</div>
          <div>Sheet: ${escapeHtml(compactText(sheetId, 22))}</div>
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
      showToast('error', 'โหลดรายชื่อชีทไม่สำเร็จ', message);

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
      const sheetName = String(sheet.sheetName || '');

      const rowText = sheet.lastRow
        ? Number(sheet.lastRow || 0).toLocaleString() + ' rows'
        : 'เลือกชีท';

      const colText = sheet.lastColumn
        ? Number(sheet.lastColumn || 0).toLocaleString() + ' cols'
        : 'read header';

      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'sheet-item' + (sheetName === selectedSheetName ? ' active' : '');
      btn.setAttribute('title', sheetName);

      btn.innerHTML = `
        <div class="item-title">
          <span>${escapeHtml(compactText(sheetName, 28))}</span>
          <span class="badge badge-muted">${escapeHtml(rowText)}</span>
        </div>

        <div class="item-meta">
          <div>${escapeHtml(colText)}</div>
          <div>กดเพื่ออ่านหัวคอลัมน์</div>
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
      showToast('error', 'อ่านหัวคอลัมน์ไม่สำเร็จ', message);

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
      showToast('warning', 'ยังเลือกข้อมูลไม่ครบ', 'กรุณาเลือกแหล่งข้อมูลและชีทก่อน');
      return;
    }

    const selectedSource = sourcesCache.find(function (src) {
      return String(src['รหัสแหล่งข้อมูล'] || '') === selectedSourceId;
    }) || {};

    const fields = collectMappingFields();

    if (!fields.length) {
      setMappingMessage('ไม่มี Mapping ให้บันทึก');
      showToast('warning', 'ไม่มี Mapping ให้บันทึก', '');
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
      showToast('success', 'บันทึก Mapping สำเร็จ', 'พร้อมสร้าง Dashboard Preview แล้ว');

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
      showToast('warning', 'ยังเลือกข้อมูลไม่ครบ', 'กรุณาเลือกแหล่งข้อมูลและชีทก่อน');
      return;
    }

    if (!window.AnalyticsAPI.dashboardPreview) {
      setPreviewMessage('ยังไม่พบฟังก์ชัน dashboardPreview ใน api.js');
      showToast('error', 'ไม่พบ API', 'ยังไม่พบฟังก์ชัน dashboardPreview ใน api.js');
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
      showToast('success', 'สร้าง Preview สำเร็จ', 'ตรวจสอบผลลัพธ์ก่อนบันทึก Dashboard จริง');

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
      showToast('warning', 'ยังเลือกข้อมูลไม่ครบ', 'กรุณาเลือกแหล่งข้อมูลและชีทก่อน');
      return;
    }

    const dashboardName = el.dashboardName ? el.dashboardName.value.trim() : '';
    const dashboardType = el.dashboardType ? el.dashboardType.value : 'Custom';

    if (!dashboardName) {
      setCreateDashboardMessage('กรุณากรอกชื่อ Dashboard');
      showToast('warning', 'ยังไม่ได้ตั้งชื่อ Dashboard', 'กรุณากรอกชื่อ Dashboard ก่อนบันทึก');
      return;
    }

    if (!window.AnalyticsAPI.createDashboardFromPreview) {
      setCreateDashboardMessage('ยังไม่พบฟังก์ชัน createDashboardFromPreview ใน api.js');
      showToast('error', 'ไม่พบ API', 'ยังไม่พบฟังก์ชัน createDashboardFromPreview ใน api.js');
      return;
    }

    setButtonLoading(el.createDashboardBtn, true, 'กำลังบันทึก Dashboard...');
    setCreateDashboardMessage('');

    try {
      const builderPayload = getDashboardBuilderPayload();

      const data = await window.AnalyticsAPI.createDashboardFromPreview({
        sourceId: selectedSourceId,
        sheetName: selectedSheetName,
        dashboardName: dashboardName,
        dashboardType: dashboardType,
        description: 'สร้างจาก Mapping ผ่าน Admin Console',
        builderConfig: builderPayload
      });

      setCreateDashboardMessage(
        'สร้าง Dashboard สำเร็จ: ' + data.dashboardId +
        ' | KPI ' + data.totalMetrics +
        ' | กราฟ ' + data.totalCharts +
        ' | ตัวกรอง ' + data.totalFilters
      );

      showToast('success', 'สร้าง Dashboard สำเร็จ', data.dashboardId || dashboardName);

      writeLog({
        step: 'create_dashboard_from_preview',
        response: data,
        builderConfig: getDashboardBuilderPayload()
      });

      await loadDashboards();
      await loadDashboardOptions();
      await loadManageDashboards();

    } catch (error) {
      setCreateDashboardMessage(error.message);
      showToast('error', 'สร้าง Dashboard ไม่สำเร็จ', error.message || '');

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
      showToast('error', 'โหลด Dashboard ไม่สำเร็จ', error.message || '');

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
      showToast('warning', 'กรุณาเลือก Dashboard', 'เลือก Dashboard ก่อน Refresh');
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
      showToast('error', 'Refresh Dashboard ไม่สำเร็จ', error.message || '');

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
      showToast('warning', 'กรุณาเลือก Dashboard', '');
      return;
    }

    if (!window.AnalyticsAPI.dashboardView) {
      setDashboardViewMessage('ยังไม่พบฟังก์ชัน dashboardView ใน api.js');
      showToast('error', 'ไม่พบ API', 'ยังไม่พบฟังก์ชัน dashboardView ใน api.js');
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
      showToast('success', 'โหลด Dashboard สำเร็จ', '');

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
      showToast('warning', 'กรุณาเลือก Dashboard', 'เลือก Dashboard ก่อนใช้ Filter');
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

      showToast('success', 'กรองข้อมูลสำเร็จ', 'เหลือ ' + Number(data.totalRowsAfterFilter || 0).toLocaleString() + ' แถว');

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

  function compactText(value, maxLength) {
    const text = String(value || '').trim();
    const limit = Number(maxLength || 24);

    if (!text || text.length <= limit) {
      return text;
    }

    if (limit <= 6) {
      return text.slice(0, limit);
    }

    const head = Math.ceil((limit - 3) * 0.6);
    const tail = Math.floor((limit - 3) * 0.4);

    return text.slice(0, head) + '...' + text.slice(text.length - tail);
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

      showToast('error', 'โหลด Dashboard ไม่สำเร็จ', error.message || '');

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

    showToast('warning', 'เซสชันหมดอายุ', 'กรุณาเข้าสู่ระบบใหม่');

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

    showToast('error', fallbackMessage || 'เกิดข้อผิดพลาด', message);
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

      showToast('error', 'โหลดรายชื่อผู้ใช้ไม่สำเร็จ', message);
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
      showToast('warning', 'ไม่พบผู้ใช้ที่เลือก', '');
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

    if (el.userCanExport) el.userCanExport.checked = !!user.canExport;
    if (el.userCanAddSource) el.userCanAddSource.checked = !!user.canAddSource;
    if (el.userCanEditDashboard) el.userCanEditDashboard.checked = !!user.canEditDashboard;
    if (el.userCanViewAuditLog) el.userCanViewAuditLog.checked = !!user.canViewAuditLog;
    if (el.userCanManageUser) el.userCanManageUser.checked = !!user.canManageUser;

    setUserManageMessage(
      'เลือกผู้ใช้แล้ว: ' +
      (user.displayName || user.username || userId) +
      ' | ช่อง Dashboard ที่ดูได้ ใช้รูปแบบ DASH_xxx,DASH_yyy'
    );

    const target = document.querySelector('#sectionUsers') || document.querySelector('.user-management-section');

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
      showToast('warning', 'ข้อมูลไม่ครบ', 'กรุณากรอกชื่อผู้ใช้และชื่อแสดงผล');
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

      showCredentialDialog({
        title: 'เพิ่มผู้ใช้สำเร็จ',
        username: data.username,
        password: data.temporaryPassword,
        note: 'กรุณาจดรหัสผ่านนี้ไว้ ระบบจะแสดงครั้งเดียว'
      });

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
      showToast('warning', 'ยังไม่ได้เลือกผู้ใช้', 'กรุณาเลือกผู้ใช้ก่อนแก้ไข');
      return;
    }

    const payload = collectUserPayload();
    payload.userId = selectedManageUserId;

    setButtonLoading(el.updateUserBtn, true, 'กำลังบันทึก...');
    setUserManageMessage('');

    try {
      const data = await window.AnalyticsAPI.updateUser(payload);

      setUserManageMessage(data.message || 'แก้ไขผู้ใช้สำเร็จ');
      showToast('success', 'แก้ไขผู้ใช้สำเร็จ', payload.displayName || payload.username);

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
      showToast('warning', 'ยังไม่ได้เลือกผู้ใช้', 'กรุณาเลือกผู้ใช้ก่อน Reset Password');
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

    const confirmed = await showConfirmDialog({
      icon: 'warning',
      title: 'ยืนยัน Reset Password',
      text: user ? user.displayName + ' (' + user.username + ')' : userId,
      confirmText: 'Reset Password',
      cancelText: 'ยกเลิก'
    });

    if (!confirmed || !confirmed.isConfirmed) {
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

      showCredentialDialog({
        title: 'Reset Password สำเร็จ',
        username: data.username || '',
        password: data.temporaryPassword,
        note: 'กรุณาจดรหัสผ่านนี้ไว้ ระบบจะแสดงครั้งเดียว'
      });

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

    const confirmed = await showConfirmDialog({
      icon: status === 'active' ? 'question' : 'warning',
      title: 'ยืนยันเปลี่ยนสถานะผู้ใช้',
      text: 'เปลี่ยนสถานะเป็น ' + status + ' : ' + (user ? user.displayName + ' (' + user.username + ')' : userId),
      confirmText: 'ยืนยัน',
      cancelText: 'ยกเลิก'
    });

    if (!confirmed || !confirmed.isConfirmed) {
      return;
    }

    try {
      const data = await window.AnalyticsAPI.setUserStatus(userId, status);

      setUserManageMessage(data.message || 'เปลี่ยนสถานะผู้ใช้สำเร็จ');
      showToast('success', 'เปลี่ยนสถานะผู้ใช้สำเร็จ', status);

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


          response: data
      });

      await loadManageDashboards();
      await loadDashboardOptions();

      if (el.manageDashboardSelect) {
        el.manageDashboardSelect.value = dashboardId;
        handleSelectManageDashboard();
      }

    } catch (error) {
      showApiError(setManageDashboardMessage, error, 'บันทึก Dashboard ไม่สำเร็จ');
      logApiError('manage_dashboard_save_error', error, {
        dashboardId: dashboardId
      });

    } finally {
      setButtonLoading(el.saveManageDashboardBtn, false, 'Save Dashboard');
    }
  }


  async function handleRegenerateManageDashboard() {
    const dashboardId = el.manageDashboardSelect ? el.manageDashboardSelect.value : '';

    if (!dashboardId) {
      setManageDashboardMessage('กรุณาเลือก Dashboard ก่อน Regenerate');
      return;
    }

    const dashboard = findManageDashboardById(dashboardId);
    const dashboardName = dashboard
      ? String(dashboard['ชื่อ Dashboard'] || dashboardId)
      : dashboardId;

    const result = await showConfirmDialog({
      icon: 'warning',
      title: 'ยืนยัน Regenerate Dashboard',
      text:
        'ยืนยัน Regenerate Dashboard หรือไม่?\n\n' +
        dashboardName +
        '\n\nระบบจะสร้าง KPI / กราฟ / ตัวกรอง / Layout ใหม่จาก Mapping ล่าสุด\n' +
        'ถ้า Dashboard นี้เคยสร้างจาก Dashboard Builder ระบบจะใช้ Widget เดิมที่ Super Admin เลือกไว้\n' +
        'และจะคงชื่อ Dashboard, Publish, Export และสิทธิ์เดิมไว้',
      confirmText: 'Regenerate',
      cancelText: 'ยกเลิก'
    });

    if (!result || !result.isConfirmed) {
      return;
    }

    setButtonLoading(el.regenerateManageDashboardBtn, true, 'กำลัง Regenerate...');
    setManageDashboardMessage('กำลัง Regenerate Dashboard จาก Mapping ล่าสุด / Builder Config เดิม.');
    showGlobalLoading('กำลัง Regenerate Dashboard จาก Mapping ล่าสุด / Builder Config เดิม.');

    try {
      const data = await window.AnalyticsAPI.regenerateDashboard(dashboardId);

      const modeText = data.regenerateMode === 'builder'
        ? 'โหมด Builder'
        : 'โหมด Auto Mapping';

      setManageDashboardMessage(
        (data.message || 'Regenerate สำเร็จ') +
        ' | ' + modeText +
        ' | KPI ' + Number(data.totalMetrics || 0).toLocaleString() +
        ' | กราฟ ' + Number(data.totalCharts || 0).toLocaleString() +
        ' | ตัวกรอง ' + Number(data.totalFilters || 0).toLocaleString()
      );

      showToast('success', 'Regenerate สำเร็จ', modeText);

      writeLog({
        step: 'manage_dashboard_regenerate',
        response: data
      });

      await loadManageDashboards();
      await loadDashboardOptions();

      if (el.manageDashboardSelect) {
        el.manageDashboardSelect.value = dashboardId;
        handleSelectManageDashboard();
      }

    } catch (error) {
      showApiError(setManageDashboardMessage, error, 'Regenerate Dashboard ไม่สำเร็จ');
      logApiError('manage_dashboard_regenerate_error', error, {
        dashboardId: dashboardId
      });

    } finally {
      hideGlobalLoading();
      setButtonLoading(el.regenerateManageDashboardBtn, false, 'Regenerate');
    }
  }


  async function handleToggleManageDashboardHidden() {
    const dashboardId = el.manageDashboardSelect ? el.manageDashboardSelect.value : '';

    if (!dashboardId) {
      setManageDashboardMessage('กรุณาเลือก Dashboard ก่อน');
      return;
    }

    const currentPublish = el.manageDashboardPublish ? el.manageDashboardPublish.checked : false;
    const nextPublish = !currentPublish;

    const result = await showConfirmDialog({
      icon: nextPublish ? 'question' : 'warning',
      title: nextPublish ? 'ยืนยันเปิดใช้งาน Dashboard' : 'ยืนยันซ่อน Dashboard',
      text: nextPublish
        ? 'ต้องการเปิด Publish Dashboard นี้หรือไม่?'
        : 'ต้องการซ่อน Dashboard นี้หรือไม่?',
      confirmText: nextPublish ? 'เปิดใช้งาน' : 'ซ่อน',
      cancelText: 'ยกเลิก'
    });

    if (!result || !result.isConfirmed) {
      return;
    }

    setButtonLoading(el.hideManageDashboardBtn, true, 'กำลังบันทึก...');

    try {
      const data = await window.AnalyticsAPI.toggleDashboardPublish(
        dashboardId,
        nextPublish
      );

      setManageDashboardMessage(data.message || 'เปลี่ยนสถานะ Dashboard สำเร็จ');
      showToast('success', 'เปลี่ยนสถานะ Dashboard สำเร็จ', nextPublish ? 'Published' : 'Hidden');

      writeLog({
        step: 'manage_dashboard_toggle_publish',
        response: data
      });

      await loadManageDashboards();
      await loadDashboardOptions();

      if (el.manageDashboardSelect) {
        el.manageDashboardSelect.value = dashboardId;
        handleSelectManageDashboard();
      }

    } catch (error) {
      showApiError(setManageDashboardMessage, error, 'เปลี่ยนสถานะ Dashboard ไม่สำเร็จ');
      logApiError('manage_dashboard_toggle_publish_error', error, {
        dashboardId: dashboardId,
        publish: nextPublish
      });

    } finally {
      setButtonLoading(el.hideManageDashboardBtn, false, 'Hide / Show');
    }
  }


  async function handleDeleteManageDashboard() {
    const dashboardId = el.manageDashboardSelect ? el.manageDashboardSelect.value : '';

    if (!dashboardId) {
      setManageDashboardMessage('กรุณาเลือก Dashboard ก่อนลบ');
      return;
    }

    const dashboard = findManageDashboardById(dashboardId);
    const dashboardName = dashboard
      ? String(dashboard['ชื่อ Dashboard'] || dashboardId)
      : dashboardId;

    const result = await showConfirmDialog({
      icon: 'warning',
      title: 'ยืนยัน Soft Delete Dashboard',
      text:
        'ต้องการลบ Dashboard แบบ Soft Delete หรือไม่?\n\n' +
        dashboardName +
        '\n\nระบบจะซ่อน Dashboard และปิด Publish แต่ยังเก็บข้อมูลไว้ในระบบ',
      confirmText: 'Soft Delete',
      cancelText: 'ยกเลิก'
    });

    if (!result || !result.isConfirmed) {
      return;
    }

    setButtonLoading(el.deleteManageDashboardBtn, true, 'กำลังลบ...');
    setManageDashboardMessage('');

    try {
      const data = await window.AnalyticsAPI.deleteDashboard(dashboardId);

      setManageDashboardMessage(data.message || 'Soft Delete Dashboard สำเร็จ');
      showToast('success', 'Soft Delete Dashboard สำเร็จ', dashboardName);

      writeLog({
        step: 'manage_dashboard_delete',
        response: data
      });

      await loadManageDashboards();
      await loadDashboardOptions();

      resetManageDashboardForm();

    } catch (error) {
      showApiError(setManageDashboardMessage, error, 'Soft Delete Dashboard ไม่สำเร็จ');
      logApiError('manage_dashboard_delete_error', error, {
        dashboardId: dashboardId
      });

    } finally {
      setButtonLoading(el.deleteManageDashboardBtn, false, 'Soft Delete');
    }
  }


  function renderManageDashboardSummary(dash) {
    if (!el.manageDashboardSummary) {
      return;
    }

    if (!dash) {
      el.manageDashboardSummary.classList.add('empty');
      el.manageDashboardSummary.textContent = 'เลือก Dashboard เพื่อดูรายละเอียด';
      return;
    }

    const dashboardId = String(dash['รหัส Dashboard'] || '').trim();
    const name = String(dash['ชื่อ Dashboard'] || dashboardId).trim();
    const type = String(dash['ประเภท Dashboard'] || '-').trim();
    const sourceId = String(dash['รหัสแหล่งข้อมูลหลัก'] || '-').trim();
    const sheetName = String(dash['รหัสชีทย่อยหลัก'] || '-').trim();
    const visibility = String(dash['สิทธิ์ที่มองเห็น'] || '-').trim();
    const publish = toBool(dash['สถานะ Publish']);
    const allowExport = toBool(dash['อนุญาต Export']);

    el.manageDashboardSummary.classList.remove('empty');

    el.manageDashboardSummary.innerHTML = `
      <div class="dashboard-manage-summary">
        <div class="dashboard-manage-summary-head">
          <div>
            <h4>${escapeHtml(name || '-')}</h4>
            <p>${escapeHtml(dash['คำอธิบาย'] || 'ไม่มีคำอธิบาย')}</p>
          </div>

          <span class="dashboard-type-pill">${escapeHtml(type || '-')}</span>
        </div>

        <div class="dashboard-manage-summary-grid">
          <div class="dashboard-manage-summary-item">
            <span>Dashboard ID</span>
            <strong>${escapeHtml(dashboardId || '-')}</strong>
          </div>

          <div class="dashboard-manage-summary-item">
            <span>Source ID</span>
            <strong>${escapeHtml(sourceId || '-')}</strong>
          </div>

          <div class="dashboard-manage-summary-item">
            <span>Sheet</span>
            <strong>${escapeHtml(sheetName || '-')}</strong>
          </div>

          <div class="dashboard-manage-summary-item">
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
            Created: ${escapeHtml(dash['วันที่สร้าง'] || '-')}
          </span>

          <span class="dashboard-status-pill is-muted">
            Updated: ${escapeHtml(dash['วันที่แก้ไขล่าสุด'] || '-')}
          </span>
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
    const blob = new Blob([content || ''], {
      type: mimeType || 'text/plain;charset=utf-8'
    });

    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');

    link.href = url;
    link.download = filename || 'download.txt';

    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    setTimeout(function () {
      URL.revokeObjectURL(url);
    }, 1000);
  }


  function setAppMode(mode) {
    mode = mode === 'user' ? 'user' : 'admin';
    currentMode = mode;

    document.body.classList.remove('mode-admin', 'mode-user');
    document.body.classList.add(mode === 'admin' ? 'mode-admin' : 'mode-user');

    if (el.adminModeBtn) {
      el.adminModeBtn.classList.toggle('active', mode === 'admin');
    }

    if (el.userModeBtn) {
      el.userModeBtn.classList.toggle('active', mode === 'user');
    }

    if (mode === 'user') {
      const viewer = document.getElementById('sectionViewer');

      if (viewer) {
        try {
          viewer.scrollIntoView({
            behavior: 'smooth',
            block: 'start'
          });
        } catch (err) {
          viewer.scrollIntoView();
        }
      }
    }
  }


  function applyRoleUi(user) {
    user = user || {};
    const role = String(user.role || '').trim();

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
    } else {
      document.body.classList.add('role-viewer', 'role-viewer-only');
    }

    const isAdmin = role === 'Super Admin' || role === 'Admin';
    const canEditDashboard = isAdmin || !!user.canEditDashboard;
    const canAddSource = isAdmin || !!user.canAddSource;
    const canManageUser = isAdmin || !!user.canManageUser;
    const canAudit = isAdmin || !!user.canViewAuditLog;

    document.querySelectorAll('.admin-only').forEach(function (node) {
      node.classList.toggle('hidden', !isAdmin && !canEditDashboard && !canAddSource);
    });

    document.querySelectorAll('.source-admin-only').forEach(function (node) {
      node.classList.toggle('hidden', !canAddSource);
    });

    document.querySelectorAll('.dashboard-admin-only').forEach(function (node) {
      node.classList.toggle('hidden', !canEditDashboard);
    });

    document.querySelectorAll('.user-admin-only').forEach(function (node) {
      node.classList.toggle('hidden', !canManageUser);
    });

    document.querySelectorAll('.audit-admin-only').forEach(function (node) {
      node.classList.toggle('hidden', !canAudit);
    });

    if (el.modeSwitcher) {
      el.modeSwitcher.classList.remove('hidden');
    }

    if (role === 'Viewer') {
      setAppMode('user');
    } else {
      setAppMode('admin');
    }
  }


  function mountQueuedCharts() {
    if (!chartRenderQueue.length) {
      return;
    }

    if (!window.echarts) {
      writeLog({
        step: 'echarts_missing',
        message: 'ไม่พบ ECharts'
      });
      return;
    }

    chartRenderQueue.forEach(function (item) {
      const target = document.getElementById(item.id);

      if (!target) {
        return;
      }

      const instance = window.echarts.init(target);
      const option = buildEChartOption(item.chart || {});

      instance.setOption(option);

      chartInstances.push(instance);
    });

    chartRenderQueue = [];

    setTimeout(function () {
      resizeDashboardCharts();
    }, 80);
  }


  function resizeDashboardCharts() {
    chartInstances.forEach(function (chart) {
      try {
        chart.resize();
      } catch (err) {}
    });
  }


  function disposeDashboardCharts() {
    chartInstances.forEach(function (chart) {
      try {
        chart.dispose();
      } catch (err) {}
    });

    chartInstances = [];
  }


  window.addEventListener('resize', function () {
    resizeDashboardCharts();
  });


  function buildEChartOption(chart) {
    const type = String(chart.type || 'bar').toLowerCase();
    const labels = chart.labels || chart.names || [];
    const values = chart.values || chart.data || [];
    const title = chart.title || '';

    if (type === 'pie' || type === 'donut') {
      const pieData = labels.map(function (label, index) {
        return {
          name: label,
          value: Number(values[index] || 0)
        };
      });

      return {
        tooltip: {
          trigger: 'item'
        },
        legend: {
          bottom: 0,
          type: 'scroll'
        },
        series: [
          {
            name: title,
            type: 'pie',
            radius: type === 'donut' ? ['45%', '70%'] : '70%',
            center: ['50%', '45%'],
            data: pieData,
            label: {
              formatter: '{b}: {d}%'
            }
          }
        ]
      };
    }

    if (type === 'line' || type === 'area') {
      return {
        tooltip: {
          trigger: 'axis'
        },
        grid: {
          left: 42,
          right: 20,
          top: 30,
          bottom: 45
        },
        xAxis: {
          type: 'category',
          data: labels
        },
        yAxis: {
          type: 'value'
        },
        series: [
          {
            name: title,
            type: 'line',
            smooth: true,
            areaStyle: type === 'area' ? {} : undefined,
            data: values
          }
        ]
      };
    }
      async function loadManageDashboards() {
    if (!el.manageDashboardCardList) {
      return;
    }

    el.manageDashboardCardList.classList.remove('empty');
    setPanelLoading(el.manageDashboardCardList, 'กำลังโหลดรายการ Dashboard...');

    if (el.reloadManageDashboardsBtn) {
      setButtonLoading(el.reloadManageDashboardsBtn, true, 'กำลังโหลด...');
    }

    try {
      const data = await window.AnalyticsAPI.listDashboards();

      manageDashboardsCache = data.dashboards || [];

      renderFilteredManageDashboardOptions();

      setManageDashboardMessage('โหลดรายการ Dashboard สำเร็จ');

      writeLog({
        step: 'manage_dashboards_load',
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

      el.manageDashboardCardList.classList.add('empty');
      el.manageDashboardCardList.textContent = message;

      showToast('error', 'โหลดรายการ Dashboard ไม่สำเร็จ', message);

      logApiError('manage_dashboards_load_error', error);

    } finally {
      if (el.reloadManageDashboardsBtn) {
        setButtonLoading(el.reloadManageDashboardsBtn, false, 'โหลดรายการ Dashboard');
      }
    }
  }

  function renderFilteredManageDashboardOptions() {
    const search = el.manageDashboardSearch
      ? String(el.manageDashboardSearch.value || '').toLowerCase().trim()
      : '';

    const typeFilter = el.manageDashboardTypeFilter
      ? String(el.manageDashboardTypeFilter.value || '').trim()
      : '';

    const statusFilter = el.manageDashboardStatusFilter
      ? String(el.manageDashboardStatusFilter.value || '').trim()
      : '';

    const filtered = manageDashboardsCache.filter(function (dash) {
      const dashboardId = String(dash['รหัส Dashboard'] || '').trim();
      const name = String(dash['ชื่อ Dashboard'] || '').trim();
      const type = String(dash['ประเภท Dashboard'] || '').trim();

      const text = [
        dashboardId,
        name,
        type,
        dash['คำอธิบาย'] || '',
        dash['รหัสแหล่งข้อมูลหลัก'] || '',
        dash['รหัสชีทย่อยหลัก'] || ''
      ].join(' ').toLowerCase();

      if (search && text.indexOf(search) < 0) {
        return false;
      }

      if (typeFilter && type !== typeFilter) {
        return false;
      }

      if (statusFilter) {
        const publish = toBool(dash['สถานะ Publish']);
        const allowExport = toBool(dash['อนุญาต Export']);
        const note = String(dash['หมายเหตุ'] || '').toLowerCase();
        const hidden = note.indexOf('soft delete') >= 0 || note.indexOf('hidden') >= 0;

        if (statusFilter === 'published' && !publish) return false;
        if (statusFilter === 'unpublished' && publish) return false;
        if (statusFilter === 'export_on' && !allowExport) return false;
        if (statusFilter === 'export_off' && allowExport) return false;
        if (statusFilter === 'hidden' && !hidden) return false;
      }

      return true;
    });

    renderManageDashboardSelect(filtered);
    renderManageDashboardCards(filtered);

    if (el.manageDashboardCount) {
      el.manageDashboardCount.textContent = filtered.length.toLocaleString() + ' รายการ';
    }

    if (el.manageDashboardCardCount) {
      el.manageDashboardCardCount.textContent = filtered.length.toLocaleString() + ' รายการ';
    }
  }

  function renderManageDashboardSelect(dashboards) {
    if (!el.manageDashboardSelect) {
      return;
    }

    const currentValue = el.manageDashboardSelect.value || '';

    if (!dashboards || !dashboards.length) {
      el.manageDashboardSelect.innerHTML = '<option value="">ยังไม่มี Dashboard</option>';
      return;
    }

    el.manageDashboardSelect.innerHTML =
      '<option value="">เลือก Dashboard</option>' +
      dashboards.map(function (dash) {
        const id = dash['รหัส Dashboard'] || '';
        const name = dash['ชื่อ Dashboard'] || id;
        const type = dash['ประเภท Dashboard'] || '';

        return `<option value="${escapeAttr(id)}">${escapeHtml(name)} (${escapeHtml(type)})</option>`;
      }).join('');

    if (currentValue) {
      el.manageDashboardSelect.value = currentValue;
    }
  }

  function renderManageDashboardCards(dashboards) {
    if (!el.manageDashboardCardList) {
      return;
    }

    dashboards = dashboards || [];

    if (!dashboards.length) {
      el.manageDashboardCardList.classList.add('empty');
      el.manageDashboardCardList.textContent = 'ไม่พบ Dashboard ตามเงื่อนไขที่ค้นหา';
      return;
    }

    el.manageDashboardCardList.classList.remove('empty');

    el.manageDashboardCardList.innerHTML = dashboards.map(function (dash) {
      const dashboardId = String(dash['รหัส Dashboard'] || '').trim();
      const name = String(dash['ชื่อ Dashboard'] || dashboardId).trim();
      const type = String(dash['ประเภท Dashboard'] || '-').trim();
      const sourceId = String(dash['รหัสแหล่งข้อมูลหลัก'] || '-').trim();
      const sheetName = String(dash['รหัสชีทย่อยหลัก'] || '-').trim();
      const publish = toBool(dash['สถานะ Publish']);
      const allowExport = toBool(dash['อนุญาต Export']);

      return `
        <article class="dashboard-manage-card" data-dashboard-id="${escapeAttr(dashboardId)}">
          <div class="dashboard-manage-card-head">
            <div>
              <h4>${escapeHtml(name || '-')}</h4>
              <p>${escapeHtml(dash['คำอธิบาย'] || 'ไม่มีคำอธิบาย')}</p>
            </div>

            <span class="dashboard-type-pill">${escapeHtml(type || '-')}</span>
          </div>

          <div class="dashboard-manage-card-meta">
            <div>
              <span>Dashboard ID</span>
              <strong>${escapeHtml(compactText(dashboardId, 24))}</strong>
            </div>

            <div>
              <span>Source</span>
              <strong>${escapeHtml(compactText(sourceId, 24))}</strong>
            </div>

            <div>
              <span>Sheet</span>
              <strong>${escapeHtml(sheetName || '-')}</strong>
            </div>

            <div>
              <span>Updated</span>
              <strong>${escapeHtml(dash['วันที่แก้ไขล่าสุด'] || '-')}</strong>
            </div>
          </div>

          <div class="dashboard-manage-card-status">
            <span class="dashboard-status-pill ${publish ? 'is-on' : 'is-off'}">
              ${publish ? 'Published' : 'Unpublished'}
            </span>

            <span class="dashboard-status-pill ${allowExport ? 'is-on' : 'is-off'}">
              ${allowExport ? 'Export On' : 'Export Off'}
            </span>
          </div>

          <div class="dashboard-manage-card-actions">
            <button class="btn btn-secondary btn-select-dashboard" type="button" data-dashboard-id="${escapeAttr(dashboardId)}">
              เลือกแก้ไข
            </button>

            <button class="btn btn-ghost btn-open-dashboard" type="button" data-dashboard-id="${escapeAttr(dashboardId)}">
              เปิดดู
            </button>
          </div>
        </article>
      `;
    }).join('');

    el.manageDashboardCardList.querySelectorAll('.btn-select-dashboard').forEach(function (btn) {
      btn.addEventListener('click', function () {
        const dashboardId = btn.getAttribute('data-dashboard-id') || '';

        if (el.manageDashboardSelect) {
          el.manageDashboardSelect.value = dashboardId;
        }

        handleSelectManageDashboard();
      });
    });

    el.manageDashboardCardList.querySelectorAll('.btn-open-dashboard').forEach(function (btn) {
      btn.addEventListener('click', function () {
        const dashboardId = btn.getAttribute('data-dashboard-id') || '';

        if (el.dashboardSelect) {
          el.dashboardSelect.value = dashboardId;
        }

        setAppMode('user');
        handleOpenDashboard();
      });
    });
  }

  function handleSelectManageDashboard() {
    const dashboardId = el.manageDashboardSelect ? el.manageDashboardSelect.value : '';

    if (!dashboardId) {
      selectedManageDashboard = null;
      resetManageDashboardForm();
      return;
    }

    const dash = findManageDashboardById(dashboardId);

    if (!dash) {
      selectedManageDashboard = null;
      setManageDashboardMessage('ไม่พบ Dashboard ที่เลือก');
      showToast('warning', 'ไม่พบ Dashboard ที่เลือก', '');
      return;
    }

    selectedManageDashboard = dash;

    if (el.manageDashboardName) el.manageDashboardName.value = dash['ชื่อ Dashboard'] || '';
    if (el.manageDashboardType) el.manageDashboardType.value = dash['ประเภท Dashboard'] || 'Custom';
    if (el.manageDashboardDescription) el.manageDashboardDescription.value = dash['คำอธิบาย'] || '';
    if (el.manageDashboardPublish) el.manageDashboardPublish.checked = toBool(dash['สถานะ Publish']);
    if (el.manageDashboardExport) el.manageDashboardExport.checked = toBool(dash['อนุญาต Export']);

    if (el.manageDashboardHidden) {
      const note = String(dash['หมายเหตุ'] || '').toLowerCase();
      el.manageDashboardHidden.checked = note.indexOf('soft delete') >= 0 || note.indexOf('hidden') >= 0;
    }

    renderManageDashboardSummary(dash);

    setManageDashboardMessage('เลือก Dashboard แล้ว: ' + (dash['ชื่อ Dashboard'] || dashboardId));
  }

  function findManageDashboardById(dashboardId) {
    dashboardId = String(dashboardId || '').trim();

    return manageDashboardsCache.find(function (dash) {
      return String(dash['รหัส Dashboard'] || '').trim() === dashboardId;
    }) || null;
  }

  function resetManageDashboardForm() {
    selectedManageDashboard = null;

    if (el.manageDashboardSelect) el.manageDashboardSelect.value = '';
    if (el.manageDashboardName) el.manageDashboardName.value = '';
    if (el.manageDashboardType) el.manageDashboardType.value = 'Operation';
    if (el.manageDashboardDescription) el.manageDashboardDescription.value = '';
    if (el.manageDashboardPublish) el.manageDashboardPublish.checked = false;
    if (el.manageDashboardExport) el.manageDashboardExport.checked = false;
    if (el.manageDashboardHidden) el.manageDashboardHidden.checked = false;

    renderManageDashboardSummary(null);
    setManageDashboardMessage('');
  }

  function collectManageDashboardPayload() {
    return {
      dashboardId: el.manageDashboardSelect ? el.manageDashboardSelect.value : '',
      dashboardName: el.manageDashboardName ? el.manageDashboardName.value.trim() : '',
      dashboardType: el.manageDashboardType ? el.manageDashboardType.value : 'Custom',
      description: el.manageDashboardDescription ? el.manageDashboardDescription.value.trim() : '',
      publish: el.manageDashboardPublish ? el.manageDashboardPublish.checked : false,
      allowExport: el.manageDashboardExport ? el.manageDashboardExport.checked : false,
      hidden: el.manageDashboardHidden ? el.manageDashboardHidden.checked : false
    };
  }

  async function handleSaveManageDashboard() {
    const payload = collectManageDashboardPayload();
    const dashboardId = payload.dashboardId;

    if (!dashboardId) {
      setManageDashboardMessage('กรุณาเลือก Dashboard ก่อนบันทึก');
      showToast('warning', 'ยังไม่ได้เลือก Dashboard', '');
      return;
    }

    if (!payload.dashboardName) {
      setManageDashboardMessage('กรุณากรอกชื่อ Dashboard');
      showToast('warning', 'กรุณากรอกชื่อ Dashboard', '');
      return;
    }

    setButtonLoading(el.saveManageDashboardBtn, true, 'กำลังบันทึก...');
    setManageDashboardMessage('');

    try {
      const data = await window.AnalyticsAPI.updateDashboard(payload);

      setManageDashboardMessage(data.message || 'บันทึก Dashboard สำเร็จ');
      showToast('success', 'บันทึก Dashboard สำเร็จ', payload.dashboardName);

      writeLog({
        step: 'manage_dashboard_save',
        response: data
      });

      await loadManageDashboards();
      await loadDashboardOptions();

      if (el.manageDashboardSelect) {
        el.manageDashboardSelect.value = dashboardId;
        handleSelectManageDashboard();
      }

    } catch (error) {
      showApiError(setManageDashboardMessage, error, 'บันทึก Dashboard ไม่สำเร็จ');
      logApiError('manage_dashboard_save_error', error, {
        dashboardId: dashboardId
      });

    } finally {
      setButtonLoading(el.saveManageDashboardBtn, false, 'Save Dashboard');
    }
  }

  async function handleRegenerateManageDashboard() {
    const dashboardId = el.manageDashboardSelect ? el.manageDashboardSelect.value : '';

    if (!dashboardId) {
      setManageDashboardMessage('กรุณาเลือก Dashboard ก่อน Regenerate');
      showToast('warning', 'ยังไม่ได้เลือก Dashboard', '');
      return;
    }

    const dashboard = findManageDashboardById(dashboardId);
    const dashboardName = dashboard
      ? String(dashboard['ชื่อ Dashboard'] || dashboardId)
      : dashboardId;

    const result = await showConfirmDialog({
      icon: 'warning',
      title: 'ยืนยัน Regenerate Dashboard',
      text:
        'ยืนยัน Regenerate Dashboard หรือไม่?\n\n' +
        dashboardName +
        '\n\nระบบจะสร้าง KPI / กราฟ / ตัวกรอง / Layout ใหม่จาก Mapping ล่าสุด\n' +
        'ถ้า Dashboard นี้เคยสร้างจาก Dashboard Builder ระบบจะใช้ Widget เดิมที่ Super Admin เลือกไว้\n' +
        'และจะคงชื่อ Dashboard, Publish, Export และสิทธิ์เดิมไว้',
      confirmText: 'Regenerate',
      cancelText: 'ยกเลิก'
    });

    if (!result || !result.isConfirmed) {
      return;
    }

    setButtonLoading(el.regenerateManageDashboardBtn, true, 'กำลัง Regenerate...');
    setManageDashboardMessage('กำลัง Regenerate Dashboard จาก Mapping ล่าสุด / Builder Config เดิม.');
    showGlobalLoading('กำลัง Regenerate Dashboard จาก Mapping ล่าสุด / Builder Config เดิม.');

    try {
      const data = await window.AnalyticsAPI.regenerateDashboard(dashboardId);

      const modeText = data.regenerateMode === 'builder'
        ? 'โหมด Builder'
        : 'โหมด Auto Mapping';

      setManageDashboardMessage(
        (data.message || 'Regenerate สำเร็จ') +
        ' | ' + modeText +
        ' | KPI ' + Number(data.totalMetrics || 0).toLocaleString() +
        ' | กราฟ ' + Number(data.totalCharts || 0).toLocaleString() +
        ' | ตัวกรอง ' + Number(data.totalFilters || 0).toLocaleString()
      );

      showToast('success', 'Regenerate สำเร็จ', modeText);

      writeLog({
        step: 'manage_dashboard_regenerate',
        response: data
      });

      await loadManageDashboards();
      await loadDashboardOptions();

      if (el.manageDashboardSelect) {
        el.manageDashboardSelect.value = dashboardId;
        handleSelectManageDashboard();
      }

    } catch (error) {
      showApiError(setManageDashboardMessage, error, 'Regenerate Dashboard ไม่สำเร็จ');
      logApiError('manage_dashboard_regenerate_error', error, {
        dashboardId: dashboardId
      });

    } finally {
      hideGlobalLoading();
      setButtonLoading(el.regenerateManageDashboardBtn, false, 'Regenerate');
    }
  }

  async function handleToggleManageDashboardHidden() {
    const dashboardId = el.manageDashboardSelect ? el.manageDashboardSelect.value : '';

    if (!dashboardId) {
      setManageDashboardMessage('กรุณาเลือก Dashboard ก่อน');
      showToast('warning', 'ยังไม่ได้เลือก Dashboard', '');
      return;
    }

    const currentPublish = el.manageDashboardPublish ? el.manageDashboardPublish.checked : false;
    const nextPublish = !currentPublish;

    const result = await showConfirmDialog({
      icon: nextPublish ? 'question' : 'warning',
      title: nextPublish ? 'ยืนยันเปิดใช้งาน Dashboard' : 'ยืนยันซ่อน Dashboard',
      text: nextPublish
        ? 'ต้องการเปิด Publish Dashboard นี้หรือไม่?'
        : 'ต้องการซ่อน Dashboard นี้หรือไม่?',
      confirmText: nextPublish ? 'เปิดใช้งาน' : 'ซ่อน',
      cancelText: 'ยกเลิก'
    });

    if (!result || !result.isConfirmed) {
      return;
    }

    setButtonLoading(el.hideManageDashboardBtn, true, 'กำลังบันทึก...');

    try {
      const data = await window.AnalyticsAPI.toggleDashboardPublish(
        dashboardId,
        nextPublish
      );

      setManageDashboardMessage(data.message || 'เปลี่ยนสถานะ Dashboard สำเร็จ');
      showToast('success', 'เปลี่ยนสถานะ Dashboard สำเร็จ', nextPublish ? 'Published' : 'Hidden');

      writeLog({
        step: 'manage_dashboard_toggle_publish',
        response: data
      });

      await loadManageDashboards();
      await loadDashboardOptions();

      if (el.manageDashboardSelect) {
        el.manageDashboardSelect.value = dashboardId;
        handleSelectManageDashboard();
      }

    } catch (error) {
      showApiError(setManageDashboardMessage, error, 'เปลี่ยนสถานะ Dashboard ไม่สำเร็จ');
      logApiError('manage_dashboard_toggle_publish_error', error, {
        dashboardId: dashboardId,
        publish: nextPublish
      });

    } finally {
      setButtonLoading(el.hideManageDashboardBtn, false, 'Hide / Show');
    }
  }
      async function handleDeleteManageDashboard() {
    const dashboardId = el.manageDashboardSelect ? el.manageDashboardSelect.value : '';

    if (!dashboardId) {
      setManageDashboardMessage('กรุณาเลือก Dashboard ก่อนลบ');
      showToast('warning', 'ยังไม่ได้เลือก Dashboard', '');
      return;
    }

    const dashboard = findManageDashboardById(dashboardId);
    const dashboardName = dashboard
      ? String(dashboard['ชื่อ Dashboard'] || dashboardId)
      : dashboardId;

    const result = await showConfirmDialog({
      icon: 'warning',
      title: 'ยืนยัน Soft Delete Dashboard',
      text:
        'ต้องการลบ Dashboard แบบ Soft Delete หรือไม่?\n\n' +
        dashboardName +
        '\n\nระบบจะซ่อน Dashboard และปิด Publish แต่ยังเก็บข้อมูลไว้ในระบบ',
      confirmText: 'Soft Delete',
      cancelText: 'ยกเลิก'
    });

    if (!result || !result.isConfirmed) {
      return;
    }

    setButtonLoading(el.deleteManageDashboardBtn, true, 'กำลังลบ...');
    setManageDashboardMessage('');

    try {
      const data = await window.AnalyticsAPI.deleteDashboard(dashboardId);

      setManageDashboardMessage(data.message || 'Soft Delete Dashboard สำเร็จ');
      showToast('success', 'Soft Delete Dashboard สำเร็จ', dashboardName);

      writeLog({
        step: 'manage_dashboard_delete',
        response: data
      });

      await loadManageDashboards();
      await loadDashboardOptions();

      resetManageDashboardForm();

    } catch (error) {
      showApiError(setManageDashboardMessage, error, 'Soft Delete Dashboard ไม่สำเร็จ');
      logApiError('manage_dashboard_delete_error', error, {
        dashboardId: dashboardId
      });

    } finally {
      setButtonLoading(el.deleteManageDashboardBtn, false, 'Soft Delete');
    }
  }

  function renderManageDashboardSummary(dash) {
    if (!el.manageDashboardSummary) {
      return;
    }

    if (!dash) {
      el.manageDashboardSummary.classList.add('empty');
      el.manageDashboardSummary.textContent = 'เลือก Dashboard เพื่อดูรายละเอียด';
      return;
    }

    const dashboardId = String(dash['รหัส Dashboard'] || '').trim();
    const name = String(dash['ชื่อ Dashboard'] || dashboardId).trim();
    const type = String(dash['ประเภท Dashboard'] || '-').trim();
    const sourceId = String(dash['รหัสแหล่งข้อมูลหลัก'] || '-').trim();
    const sheetName = String(dash['รหัสชีทย่อยหลัก'] || '-').trim();
    const visibility = String(dash['สิทธิ์ที่มองเห็น'] || '-').trim();
    const publish = toBool(dash['สถานะ Publish']);
    const allowExport = toBool(dash['อนุญาต Export']);

    el.manageDashboardSummary.classList.remove('empty');

    el.manageDashboardSummary.innerHTML = `
      <div class="dashboard-manage-summary">
        <div class="dashboard-manage-summary-head">
          <div>
            <h4>${escapeHtml(name || '-')}</h4>
            <p>${escapeHtml(dash['คำอธิบาย'] || 'ไม่มีคำอธิบาย')}</p>
          </div>

          <span class="dashboard-type-pill">${escapeHtml(type || '-')}</span>
        </div>

        <div class="dashboard-manage-summary-grid">
          <div class="dashboard-manage-summary-item">
            <span>Dashboard ID</span>
            <strong>${escapeHtml(dashboardId || '-')}</strong>
          </div>

          <div class="dashboard-manage-summary-item">
            <span>Source ID</span>
            <strong>${escapeHtml(sourceId || '-')}</strong>
          </div>

          <div class="dashboard-manage-summary-item">
            <span>Sheet</span>
            <strong>${escapeHtml(sheetName || '-')}</strong>
          </div>

          <div class="dashboard-manage-summary-item">
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
            Created: ${escapeHtml(dash['วันที่สร้าง'] || '-')}
          </span>

          <span class="dashboard-status-pill is-muted">
            Updated: ${escapeHtml(dash['วันที่แก้ไขล่าสุด'] || '-')}
          </span>
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
    const blob = new Blob([content || ''], {
      type: mimeType || 'text/plain;charset=utf-8'
    });

    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');

    link.href = url;
    link.download = filename || 'download.txt';

    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    setTimeout(function () {
      URL.revokeObjectURL(url);
    }, 1000);
  }

  function setAppMode(mode) {
    mode = mode === 'user' ? 'user' : 'admin';
    currentMode = mode;

    document.body.classList.remove('mode-admin', 'mode-user');
    document.body.classList.add(mode === 'admin' ? 'mode-admin' : 'mode-user');

    if (el.adminModeBtn) {
      el.adminModeBtn.classList.toggle('active', mode === 'admin');
    }

    if (el.userModeBtn) {
      el.userModeBtn.classList.toggle('active', mode === 'user');
    }

    if (mode === 'user') {
      const viewer = document.getElementById('sectionViewer');

      if (viewer) {
        try {
          viewer.scrollIntoView({
            behavior: 'smooth',
            block: 'start'
          });
        } catch (err) {
          viewer.scrollIntoView();
        }
      }
    }
  }

  function applyRoleUi(user) {
    user = user || {};
    const role = String(user.role || '').trim();

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
    } else {
      document.body.classList.add('role-viewer', 'role-viewer-only');
    }

    const isAdmin = role === 'Super Admin' || role === 'Admin';
    const canEditDashboard = isAdmin || !!user.canEditDashboard;
    const canAddSource = isAdmin || !!user.canAddSource;
    const canManageUser = isAdmin || !!user.canManageUser;
    const canAudit = isAdmin || !!user.canViewAuditLog;

    document.querySelectorAll('.admin-only').forEach(function (node) {
      node.classList.toggle('hidden', !isAdmin && !canEditDashboard && !canAddSource);
    });

    document.querySelectorAll('.source-admin-only').forEach(function (node) {
      node.classList.toggle('hidden', !canAddSource);
    });

    document.querySelectorAll('.dashboard-admin-only').forEach(function (node) {
      node.classList.toggle('hidden', !canEditDashboard);
    });

    document.querySelectorAll('.user-admin-only').forEach(function (node) {
      node.classList.toggle('hidden', !canManageUser);
    });

    document.querySelectorAll('.audit-admin-only').forEach(function (node) {
      node.classList.toggle('hidden', !canAudit);
    });

    if (el.modeSwitcher) {
      el.modeSwitcher.classList.remove('hidden');
    }

    if (role === 'Viewer') {
      setAppMode('user');
    } else {
      setAppMode('admin');
    }
  }

  function mountQueuedCharts() {
    if (!chartRenderQueue.length) {
      return;
    }

    if (!window.echarts) {
      writeLog({
        step: 'echarts_missing',
        message: 'ไม่พบ ECharts'
      });
      return;
    }

    chartRenderQueue.forEach(function (item) {
      const target = document.getElementById(item.id);

      if (!target) {
        return;
      }

      const instance = window.echarts.init(target);
      const option = buildEChartOption(item.chart || {});

      instance.setOption(option);

      chartInstances.push(instance);
    });

    chartRenderQueue = [];

    setTimeout(function () {
      resizeDashboardCharts();
    }, 80);
  }

  function resizeDashboardCharts() {
    chartInstances.forEach(function (chart) {
      try {
        chart.resize();
      } catch (err) {}
    });
  }

  function disposeDashboardCharts() {
    chartInstances.forEach(function (chart) {
      try {
        chart.dispose();
      } catch (err) {}
    });

    chartInstances = [];
  }

  window.addEventListener('resize', function () {
    resizeDashboardCharts();
  });

  function buildEChartOption(chart) {
    const type = String(chart.type || 'bar').toLowerCase();
    const labels = chart.labels || chart.names || [];
    const values = chart.values || chart.data || [];
    const title = chart.title || '';

    if (type === 'pie' || type === 'donut') {
      const pieData = labels.map(function (label, index) {
        return {
          name: label,
          value: Number(values[index] || 0)
        };
      });

      return {
        tooltip: {
          trigger: 'item'
        },
        legend: {
          bottom: 0,
          type: 'scroll'
        },
        series: [
          {
            name: title,
            type: 'pie',
            radius: type === 'donut' ? ['45%', '70%'] : '70%',
            center: ['50%', '45%'],
            data: pieData,
            label: {
              formatter: '{b}: {d}%'
            }
          }
        ]
      };
    }

    if (type === 'line' || type === 'area') {
      return {
        tooltip: {
          trigger: 'axis'
        },
        grid: {
          left: 42,
          right: 20,
          top: 30,
          bottom: 45
        },
        xAxis: {
          type: 'category',
          data: labels
        },
        yAxis: {
          type: 'value'
        },
        series: [
          {
            name: title,
            type: 'line',
            smooth: true,
            areaStyle: type === 'area' ? {} : undefined,
            data: values
          }
        ]
      };
    }

    if (type === 'horizontal_bar') {
      return {
        tooltip: {
          trigger: 'axis',
          axisPointer: {
            type: 'shadow'
          }
        },
        grid: {
          left: 90,
          right: 24,
          top: 24,
          bottom: 30
        },
        xAxis: {
          type: 'value'
        },
        yAxis: {
          type: 'category',
          data: labels
        },
        series: [
          {
            name: title,
            type: 'bar',
            data: values,
            barMaxWidth: 28
          }
        ]
      };
    }

    return {
      tooltip: {
        trigger: 'axis',
        axisPointer: {
          type: 'shadow'
        }
      },
      grid: {
        left: 42,
        right: 20,
        top: 30,
        bottom: 50
      },
      xAxis: {
        type: 'category',
        data: labels,
        axisLabel: {
          interval: 0,
          rotate: labels.length > 6 ? 28 : 0
        }
      },
      yAxis: {
        type: 'value'
      },
      series: [
        {
          name: title,
          type: 'bar',
          data: values,
          barMaxWidth: 36
        }
      ]
    };
  }

  async function handleExportDashboard() {
    const dashboardId = currentDashboardId || (el.dashboardSelect ? el.dashboardSelect.value : '');

    if (!dashboardId) {
      setDashboardViewMessage('กรุณาเลือก Dashboard ก่อน Export');
      showToast('warning', 'กรุณาเลือก Dashboard', 'เลือก Dashboard ก่อน Export');
      return;
    }

    setButtonLoading(el.exportDashboardBtn, true, 'กำลัง Export...');

    try {
      const filters = collectDashboardFilters();

      const data = await window.AnalyticsAPI.exportDashboard({
        dashboardId: dashboardId,
        format: 'csv',
        filters: filters
      });

      const filename = data.filename || ('dashboard_' + dashboardId + '.csv');
      const content = data.content || data.csv || '';

      downloadTextFile(filename, content, 'text/csv;charset=utf-8');

      setDashboardViewMessage('Export CSV สำเร็จ');
      showToast('success', 'Export CSV สำเร็จ', filename);

      writeLog({
        step: 'dashboard_export_csv',
        response: {
          ok: data.ok,
          filename: filename
        }
      });

    } catch (error) {
      showApiError(setDashboardViewMessage, error, 'Export CSV ไม่สำเร็จ');
      logApiError('dashboard_export_csv_error', error, {
        dashboardId: dashboardId
      });

    } finally {
      setButtonLoading(el.exportDashboardBtn, false, 'Export CSV');
    }
  }

  async function handleExportDashboardExcel() {
    const dashboardId = currentDashboardId || (el.dashboardSelect ? el.dashboardSelect.value : '');

    if (!dashboardId) {
      setDashboardViewMessage('กรุณาเลือก Dashboard ก่อน Export Excel');
      showToast('warning', 'กรุณาเลือก Dashboard', 'เลือก Dashboard ก่อน Export Excel');
      return;
    }

    setButtonLoading(el.exportDashboardExcelBtn, true, 'กำลัง Export Excel...');

    try {
      const filters = collectDashboardFilters();

      const data = await window.AnalyticsAPI.exportDashboard({
        dashboardId: dashboardId,
        format: 'excel',
        filters: filters
      });

      if (data.base64 && window.XLSX) {
        const binary = atob(data.base64);
        const bytes = new Uint8Array(binary.length);

        for (let i = 0; i < binary.length; i++) {
          bytes[i] = binary.charCodeAt(i);
        }

        const blob = new Blob([bytes], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });

        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');

        link.href = url;
        link.download = data.filename || ('dashboard_' + dashboardId + '.xlsx');
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);

        setTimeout(function () {
          URL.revokeObjectURL(url);
        }, 1000);
      } else if (data.content) {
        downloadTextFile(
          data.filename || ('dashboard_' + dashboardId + '.csv'),
          data.content,
          'text/csv;charset=utf-8'
        );
      }

      setDashboardViewMessage('Export Excel สำเร็จ');
      showToast('success', 'Export Excel สำเร็จ', data.filename || '');

      writeLog({
        step: 'dashboard_export_excel',
        response: {
          ok: data.ok,
          filename: data.filename || ''
        }
      });

    } catch (error) {
      showApiError(setDashboardViewMessage, error, 'Export Excel ไม่สำเร็จ');
      logApiError('dashboard_export_excel_error', error, {
        dashboardId: dashboardId
      });

    } finally {
      setButtonLoading(el.exportDashboardExcelBtn, false, 'Export Excel');
    }
  }
      async function loadAuditLog() {
    if (!el.auditLogResult) {
      return;
    }

    const limit = el.auditLimitSelect ? Number(el.auditLimitSelect.value || 100) : 100;

    setPanelLoading(el.auditLogResult, 'กำลังโหลด Audit Log...');

    if (el.reloadAuditLogBtn) {
      setButtonLoading(el.reloadAuditLogBtn, true, 'กำลังโหลด...');
    }

    try {
      const data = await window.AnalyticsAPI.auditLog({
        limit: limit
      });

      const logs = data.logs || data.rows || [];

      if (!logs.length) {
        el.auditLogResult.classList.add('empty');
        el.auditLogResult.textContent = 'ยังไม่มี Audit Log';
        return;
      }

      el.auditLogResult.classList.remove('empty');

      el.auditLogResult.innerHTML = `
        <div class="audit-list">
          ${logs.map(function (log) {
            return `
              <article class="audit-item">
                <div class="audit-item-head">
                  <strong>${escapeHtml(log.action || log['Action'] || '-')}</strong>
                  <span>${escapeHtml(log.status || log['Status'] || '-')}</span>
                </div>

                <div class="audit-item-meta">
                  <span>${escapeHtml(log.timestamp || log['Timestamp'] || log['วันที่เวลา'] || '-')}</span>
                  <span>${escapeHtml(log.username || log['Username'] || '-')}</span>
                  <span>${escapeHtml(log.role || log['Role'] || '-')}</span>
                </div>

                <p>${escapeHtml(log.detail || log['Detail'] || log.message || '-')}</p>
              </article>
            `;
          }).join('')}
        </div>
      `;

      writeLog({
        step: 'audit_log',
        response: {
          ok: data.ok,
          total: logs.length
        }
      });

    } catch (error) {
      const message = getFriendlyErrorMessage(error);

      if (handleAuthErrorIfNeeded(error)) {
        return;
      }

      el.auditLogResult.classList.add('empty');
      el.auditLogResult.textContent = message;

      showToast('error', 'โหลด Audit Log ไม่สำเร็จ', message);

      logApiError('audit_log_error', error);

    } finally {
      if (el.reloadAuditLogBtn) {
        setButtonLoading(el.reloadAuditLogBtn, false, 'โหลด Audit Log');
      }
    }
  }

  async function runSystemCheck() {
    if (!el.systemCheckResult) {
      return;
    }

    setPanelLoading(el.systemCheckResult, 'กำลังตรวจสอบระบบ...');

    try {
      const result = {
        time: formatClientDateTime(new Date()),
        health: null,
        setupStatus: null,
        me: null,
        sources: null,
        dashboards: null
      };

      try {
        result.health = await window.AnalyticsAPI.health();
      } catch (err) {
        result.health = {
          ok: false,
          message: err.message
        };
      }

      try {
        result.setupStatus = await window.AnalyticsAPI.setupStatus();
      } catch (err) {
        result.setupStatus = {
          ok: false,
          message: err.message
        };
      }

      try {
        result.me = await window.AnalyticsAPI.me();
      } catch (err) {
        result.me = {
          ok: false,
          message: err.message
        };
      }

      try {
        result.sources = await window.AnalyticsAPI.listSources();
      } catch (err) {
        result.sources = {
          ok: false,
          message: err.message
        };
      }

      try {
        result.dashboards = await window.AnalyticsAPI.listDashboards();
      } catch (err) {
        result.dashboards = {
          ok: false,
          message: err.message
        };
      }

      el.systemCheckResult.classList.remove('empty');

      el.systemCheckResult.innerHTML = `
        <div class="system-check-grid">
          ${renderSystemCheckItem('API Health', result.health)}
          ${renderSystemCheckItem('Setup Status', result.setupStatus)}
          ${renderSystemCheckItem('Current User', result.me)}
          ${renderSystemCheckItem('Sources', result.sources)}
          ${renderSystemCheckItem('Dashboards', result.dashboards)}
        </div>

        <pre class="system-check-json">${escapeHtml(JSON.stringify(result, null, 2))}</pre>
      `;

      writeLog({
        step: 'system_check',
        response: result
      });

    } catch (error) {
      const message = getFriendlyErrorMessage(error);

      el.systemCheckResult.classList.add('empty');
      el.systemCheckResult.textContent = message;

      showToast('error', 'ตรวจสอบระบบไม่สำเร็จ', message);

      logApiError('system_check_error', error);
    }
  }

  function renderSystemCheckItem(title, data) {
    data = data || {};

    const ok = data.ok !== false;
    const message = data.message || data.status || (ok ? 'พร้อมใช้งาน' : 'ไม่สำเร็จ');

    return `
      <div class="system-check-item ${ok ? 'is-ok' : 'is-error'}">
        <span>${escapeHtml(title)}</span>
        <strong>${ok ? 'OK' : 'ERROR'}</strong>
        <p>${escapeHtml(message)}</p>
      </div>
    `;
  }

  function initDashboardBuilderDesigner() {
    restoreDashboardBuilderState();
    renderDashboardWidgetDesigner();
    bindDashboardBuilderDesignerEvents();

    window.AnalyticsDashboardBuilder = {
      getSelectedWidgets: getSelectedDashboardWidgets,
      getPayload: getDashboardBuilderPayload,
      reset: resetDashboardBuilderSelection
    };
  }

  function restoreDashboardBuilderState() {
    try {
      const raw = localStorage.getItem(DASHBOARD_BUILDER_STORAGE_KEY);
      const saved = raw ? JSON.parse(raw) : null;

      if (saved && Array.isArray(saved.selectedWidgetIds)) {
        dashboardBuilderState = {
          selectedWidgetIds: saved.selectedWidgetIds,
          lastSelectedAt: saved.lastSelectedAt || '',
          mode: saved.mode || 'manual'
        };
        return;
      }
    } catch (error) {
      console.warn('restoreDashboardBuilderState error:', error);
    }

    dashboardBuilderState = {
      selectedWidgetIds: DASHBOARD_WIDGET_CATALOG
        .filter(function (item) {
          return item.recommended;
        })
        .map(function (item) {
          return item.id;
        }),
      lastSelectedAt: new Date().toISOString(),
      mode: 'auto'
    };

    saveDashboardBuilderState();
  }

  function saveDashboardBuilderState() {
    try {
      localStorage.setItem(
        DASHBOARD_BUILDER_STORAGE_KEY,
        JSON.stringify(dashboardBuilderState)
      );
    } catch (error) {
      console.warn('saveDashboardBuilderState error:', error);
    }
  }

  function bindDashboardBuilderDesignerEvents() {
    const box = document.getElementById('dashboardWidgetDesigner');

    if (!box) {
      return;
    }

    box.addEventListener('click', function (event) {
      const toggleBtn = event.target.closest('[data-widget-toggle]');
      const sampleBtn = event.target.closest('[data-widget-sample]');
      const selectRecommendedBtn = event.target.closest('[data-widget-action="recommended"]');
      const selectAllBtn = event.target.closest('[data-widget-action="all"]');
      const clearBtn = event.target.closest('[data-widget-action="clear"]');

      if (toggleBtn) {
        toggleDashboardWidget(toggleBtn.getAttribute('data-widget-toggle'));
        return;
      }

      if (sampleBtn) {
        showDashboardWidgetSample(sampleBtn.getAttribute('data-widget-sample'));
        return;
      }

      if (selectRecommendedBtn) {
        selectRecommendedDashboardWidgets();
        return;
      }

      if (selectAllBtn) {
        selectAllDashboardWidgets();
        return;
      }

      if (clearBtn) {
        resetDashboardBuilderSelection();
      }
    });
  }

  function renderDashboardWidgetDesigner() {
    const target = document.querySelector('.builder-card-preview');

    if (!target) {
      return;
    }

    target.innerHTML = `
      <div id="dashboardWidgetDesigner" class="dashboard-widget-designer">
        <div class="builder-card-head">
          <span class="builder-step">B</span>
          <div>
            <h4>เลือก Chart / Card ที่ต้องการแสดง</h4>
            <p>
              Super Admin สามารถเลือก Widget ได้มากกว่า 1 แบบจากข้อมูลชุดเดียวกัน
            </p>
          </div>
        </div>

        <div class="widget-designer-toolbar">
          <button class="btn btn-secondary" type="button" data-widget-action="recommended">
            เลือกแบบแนะนำ
          </button>

          <button class="btn btn-ghost" type="button" data-widget-action="all">
            เลือกทั้งหมด
          </button>

          <button class="btn btn-ghost" type="button" data-widget-action="clear">
            ล้างการเลือก
          </button>
        </div>

        <div class="widget-type-group">
          <div class="widget-group-title">
            <strong>Card / KPI</strong>
            <span>ใช้แสดงตัวเลขสำคัญแบบอ่านเร็ว</span>
          </div>

          <div class="widget-choice-grid advanced-widget-grid">
            ${renderWidgetChoiceCards('metric')}
          </div>
        </div>

        <div class="widget-type-group">
          <div class="widget-group-title">
            <strong>Charts</strong>
            <span>ใช้วิเคราะห์แนวโน้ม เปรียบเทียบ และสัดส่วน</span>
          </div>

          <div class="widget-choice-grid advanced-widget-grid">
            ${renderWidgetChoiceCards('chart')}
          </div>
        </div>

        <div class="widget-type-group">
          <div class="widget-group-title">
            <strong>Tables</strong>
            <span>ใช้ตรวจสอบข้อมูลจริงและจัดอันดับ</span>
          </div>

          <div class="widget-choice-grid advanced-widget-grid">
            ${renderWidgetChoiceCards('table')}
          </div>
        </div>

        <div id="dashboardWidgetSummary" class="widget-summary-box"></div>
      </div>
    `;

    renderDashboardWidgetSummary();
  }

  function renderWidgetChoiceCards(type) {
    return DASHBOARD_WIDGET_CATALOG
      .filter(function (item) {
        return item.type === type;
      })
      .map(function (item) {
        const selected = dashboardBuilderState.selectedWidgetIds.indexOf(item.id) >= 0;
        const recommended = item.recommended ? '<span class="widget-badge">แนะนำ</span>' : '';

        return `
          <article class="widget-choice advanced-widget-choice ${selected ? 'is-selected' : ''}">
            <div class="widget-choice-top">
              <div class="widget-icon">${escapeHtml(item.icon)}</div>
              <div>
                <strong>${escapeHtml(item.shortTitle)}</strong>
                ${recommended}
              </div>
            </div>

            <p>${escapeHtml(item.description)}</p>

            <div class="widget-mini-preview ${escapeAttr(item.id)}">
              ${renderWidgetMiniPreview(item)}
            </div>

            <div class="widget-best-for">
              <span>เหมาะกับ:</span>
              ${escapeHtml(item.bestFor)}
            </div>

            <div class="widget-choice-actions">
              <button
                class="btn ${selected ? 'btn-primary' : 'btn-secondary'}"
                type="button"
                data-widget-toggle="${escapeAttr(item.id)}"
              >
                ${selected ? 'เลือกแล้ว' : 'เลือกใช้งาน'}
              </button>

              <button
                class="btn btn-ghost"
                type="button"
                data-widget-sample="${escapeAttr(item.id)}"
              >
                ดูตัวอย่าง
              </button>
            </div>
          </article>
        `;
      })
      .join('');
  }

  function renderWidgetMiniPreview(item) {
    if (item.type === 'metric' && item.subType === 'summary') {
      return `
        <div class="mini-kpi">
          <span>งานทั้งหมด</span>
          <strong>12,450</strong>
        </div>
      `;
    }

    if (item.type === 'metric' && item.subType === 'trend') {
      return `
        <div class="mini-kpi">
          <span>วันนี้</span>
          <strong>850</strong>
          <em>+12.4%</em>
        </div>
      `;
    }

    if (item.type === 'metric' && item.subType === 'target') {
      return `
        <div class="mini-kpi">
          <span>เป้าหมาย</span>
          <strong>86%</strong>
          <div class="mini-progress"><i style="width:86%"></i></div>
        </div>
      `;
    }

    if (item.chartType === 'bar' || item.chartType === 'horizontal_bar') {
      return `
        <div class="mini-bars ${item.chartType === 'horizontal_bar' ? 'is-horizontal' : ''}">
          <i style="height:35%"></i>
          <i style="height:68%"></i>
          <i style="height:48%"></i>
          <i style="height:82%"></i>
          <i style="height:56%"></i>
        </div>
      `;
    }

    if (item.chartType === 'line' || item.chartType === 'area') {
      return `
        <div class="mini-line">
          <svg viewBox="0 0 120 54" preserveAspectRatio="none" aria-hidden="true">
            <path d="M4 42 C20 28, 32 36, 45 25 S72 18, 86 30 S105 12, 116 18"></path>
            ${item.chartType === 'area' ? '<path class="area" d="M4 42 C20 28, 32 36, 45 25 S72 18, 86 30 S105 12, 116 18 L116 52 L4 52 Z"></path>' : ''}
          </svg>
        </div>
      `;
    }

    if (item.chartType === 'donut' || item.chartType === 'pie') {
      return `
        <div class="mini-donut">
          <span></span>
        </div>
      `;
    }

    if (item.type === 'table') {
      return `
        <div class="mini-table">
          <b></b><b></b><b></b>
          <i></i><i></i><i></i>
          <i></i><i></i><i></i>
          <i></i><i></i><i></i>
        </div>
      `;
    }

    return '<div class="mini-placeholder"></div>';
  }

  function renderDashboardWidgetSummary() {
    const summary = document.getElementById('dashboardWidgetSummary');

    if (!summary) {
      return;
    }

    const selected = getSelectedDashboardWidgets();

    if (!selected.length) {
      summary.innerHTML = `
        <div class="widget-summary-empty">
          ยังไม่ได้เลือก Widget ระบบจะใช้รูปแบบอัตโนมัติเดิมจาก Mapping
        </div>
      `;
      return;
    }

    const metricCount = selected.filter(function (item) { return item.type === 'metric'; }).length;
    const chartCount = selected.filter(function (item) { return item.type === 'chart'; }).length;
    const tableCount = selected.filter(function (item) { return item.type === 'table'; }).length;

    summary.innerHTML = `
      <div class="widget-summary-head">
        <strong>Widget ที่เลือกแล้ว</strong>
        <span>
          Card ${metricCount} / Chart ${chartCount} / Table ${tableCount}
        </span>
      </div>

      <div class="selected-widget-list">
        ${selected.map(function (item, index) {
          return `
            <div class="selected-widget-item">
              <span>${index + 1}</span>
              <strong>${escapeHtml(item.shortTitle)}</strong>
              <em>${escapeHtml(item.bestFor)}</em>
            </div>
          `;
        }).join('')}
      </div>
    `;
  }

  function toggleDashboardWidget(widgetId) {
    widgetId = String(widgetId || '').trim();

    if (!widgetId) {
      return;
    }

    const index = dashboardBuilderState.selectedWidgetIds.indexOf(widgetId);

    if (index >= 0) {
      dashboardBuilderState.selectedWidgetIds.splice(index, 1);
    } else {
      dashboardBuilderState.selectedWidgetIds.push(widgetId);
    }

    dashboardBuilderState.mode = 'manual';
    dashboardBuilderState.lastSelectedAt = new Date().toISOString();

    saveDashboardBuilderState();
    renderDashboardWidgetDesigner();
  }

  function selectRecommendedDashboardWidgets() {
    dashboardBuilderState.selectedWidgetIds = DASHBOARD_WIDGET_CATALOG
      .filter(function (item) {
        return item.recommended;
      })
      .map(function (item) {
        return item.id;
      });

    dashboardBuilderState.mode = 'recommended';
    dashboardBuilderState.lastSelectedAt = new Date().toISOString();

    saveDashboardBuilderState();
    renderDashboardWidgetDesigner();
  }

  function selectAllDashboardWidgets() {
    dashboardBuilderState.selectedWidgetIds = DASHBOARD_WIDGET_CATALOG.map(function (item) {
      return item.id;
    });

    dashboardBuilderState.mode = 'all';
    dashboardBuilderState.lastSelectedAt = new Date().toISOString();

    saveDashboardBuilderState();
    renderDashboardWidgetDesigner();
  }

  function resetDashboardBuilderSelection() {
    dashboardBuilderState.selectedWidgetIds = [];
    dashboardBuilderState.mode = 'empty';
    dashboardBuilderState.lastSelectedAt = new Date().toISOString();

    saveDashboardBuilderState();
    renderDashboardWidgetDesigner();
  }

  function getSelectedDashboardWidgets() {
    return DASHBOARD_WIDGET_CATALOG.filter(function (item) {
      return dashboardBuilderState.selectedWidgetIds.indexOf(item.id) >= 0;
    });
  }

  function getDashboardBuilderPayload() {
    return {
      mode: dashboardBuilderState.mode,
      selectedWidgetIds: dashboardBuilderState.selectedWidgetIds.slice(),
      widgets: getSelectedDashboardWidgets().map(function (item, index) {
        return {
          order: index + 1,
          id: item.id,
          type: item.type,
          subType: item.subType || '',
          chartType: item.chartType || '',
          title: item.title,
          shortTitle: item.shortTitle,
          description: item.description,
          bestFor: item.bestFor,
          recommended: !!item.recommended
        };
      }),
      selectedAt: dashboardBuilderState.lastSelectedAt
    };
  }

  function showDashboardWidgetSample(widgetId) {
    const item = DASHBOARD_WIDGET_CATALOG.find(function (x) {
      return x.id === widgetId;
    });

    if (!item) {
      return;
    }

    const sampleHtml = `
      <div class="widget-sample-modal">
        <div class="widget-sample-desc">
          <strong>คำอธิบาย</strong>
          <p>${escapeHtml(item.description)}</p>
          <strong>เหมาะกับ</strong>
          <p>${escapeHtml(item.bestFor)}</p>
        </div>

        <div class="widget-sample-preview">
          ${renderWidgetMiniPreview(item)}
        </div>
      </div>
    `;

    if (window.Swal && window.Swal.fire) {
      window.Swal.fire({
        title: item.title,
        html: sampleHtml,
        width: 720,
        confirmButtonText: 'เข้าใจแล้ว',
        customClass: {
          popup: 'dashboard-widget-swal'
        }
      });
      return;
    }

    alert(item.title + '\n\n' + item.description + '\nเหมาะกับ: ' + item.bestFor);
  }

  function renderDataTypeOptions(selected) {
    const types = [
      ['text', 'Text'],
      ['number', 'Number'],
      ['date', 'Date'],
      ['datetime', 'DateTime'],
      ['time', 'Time'],
      ['boolean', 'Boolean'],
      ['url', 'URL'],
      ['image', 'Image'],
      ['json', 'JSON']
    ];

    selected = String(selected || 'text').toLowerCase();

    return types.map(function (item) {
      return `<option value="${escapeAttr(item[0])}" ${item[0] === selected ? 'selected' : ''}>${escapeHtml(item[1])}</option>`;
    }).join('');
  }

  function renderMeaningTypeOptions(selected) {
    const meanings = [
      ['text', 'ข้อความทั่วไป'],
      ['date', 'วันที่หลัก'],
      ['time', 'เวลา'],
      ['measure', 'ตัวเลข / KPI'],
      ['category', 'หมวดหมู่'],
      ['status', 'สถานะ'],
      ['person', 'บุคคล'],
      ['team', 'ทีม'],
      ['location', 'สถานที่'],
      ['document', 'เอกสาร'],
      ['image', 'รูปภาพ'],
      ['url', 'URL'],
      ['id', 'รหัสอ้างอิง'],
      ['note', 'หมายเหตุ']
    ];

    selected = String(selected || 'text').toLowerCase();

    return meanings.map(function (item) {
      return `<option value="${escapeAttr(item[0])}" ${item[0] === selected ? 'selected' : ''}>${escapeHtml(item[1])}</option>`;
    }).join('');
  }

  function renderCheck(name, label, checked) {
    return `
      <label class="mapping-check">
        <input
          type="checkbox"
          data-map-check="${escapeAttr(name)}"
          ${checked ? 'checked' : ''}
        >
        <span>${escapeHtml(label)}</span>
      </label>
    `;
  }

  function shouldDefaultFilter(meaning, type) {
    meaning = String(meaning || '').toLowerCase();
    type = String(type || '').toLowerCase();

    return [
      'date',
      'category',
      'status',
      'person',
      'team',
      'location',
      'document'
    ].indexOf(meaning) >= 0 || type === 'date' || type === 'datetime';
  }

  function getFirstSampleValue(rows, index) {
    rows = rows || [];

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];

      if (Array.isArray(row)) {
        if (row[index] !== undefined && row[index] !== '') {
          return row[index];
        }
      } else if (row && typeof row === 'object') {
        const values = Object.keys(row).map(function (key) {
          return row[key];
        });

        if (values[index] !== undefined && values[index] !== '') {
          return values[index];
        }
      }
    }

    return '';
  }

  function showToast(type, title, text) {
    if (!window.Swal || !window.Swal.fire) {
      if (text) {
        console.log(title + ': ' + text);
      } else {
        console.log(title);
      }
      return;
    }

    const icon = type || 'info';

    window.Swal.fire({
      toast: true,
      position: 'top-end',
      icon: icon,
      title: title || '',
      text: text || '',
      showConfirmButton: false,
      timer: icon === 'error' ? 4200 : 2600,
      timerProgressBar: true
    });
  }

  function showConfirmDialog(options) {
    options = options || {};

    if (!window.Swal || !window.Swal.fire) {
      return Promise.resolve({
        isConfirmed: window.confirm(options.text || 'ยืนยันการทำรายการ?')
      });
    }

    return window.Swal.fire({
      icon: options.icon || 'question',
      title: options.title || 'ยืนยันการทำรายการ',
      text: options.text || '',
      html: options.html || undefined,
      showCancelButton: true,
      confirmButtonText: options.confirmText || 'ยืนยัน',
      cancelButtonText: options.cancelText || 'ยกเลิก',
      reverseButtons: true,
      focusCancel: true
    });
  }

  function showCredentialDialog(payload) {
    payload = payload || {};

    const html = `
      <div style="text-align:left;display:grid;gap:10px;">
        <div>
          <strong>Username</strong>
          <div style="padding:10px;border-radius:12px;background:#f1f5f9;">${escapeHtml(payload.username || '-')}</div>
        </div>

        <div>
          <strong>Password</strong>
          <div style="padding:10px;border-radius:12px;background:#fef3c7;color:#92400e;font-weight:900;">${escapeHtml(payload.password || '-')}</div>
        </div>

        <p style="margin:0;color:#64748b;">${escapeHtml(payload.note || '')}</p>
      </div>
    `;

    if (window.Swal && window.Swal.fire) {
      window.Swal.fire({
        icon: 'success',
        title: payload.title || 'สำเร็จ',
        html: html,
        confirmButtonText: 'รับทราบ'
      });
      return;
    }

    alert(
      (payload.title || 'สำเร็จ') +
      '\nUsername: ' +
      (payload.username || '-') +
      '\nPassword: ' +
      (payload.password || '-')
    );
  }

  function toBool(value) {
    if (value === true) return true;
    if (value === false) return false;

    const text = String(value || '').trim().toLowerCase();

    return [
      'true',
      'yes',
      'y',
      '1',
      'on',
      'active',
      'published',
      'ใช่',
      'เปิด',
      'ใช้งาน'
    ].indexOf(text) >= 0;
  }

  function formatNumber(value) {
    const number = Number(value);

    if (!isFinite(number)) {
      return escapeHtml(value || '0');
    }

    return number.toLocaleString();
  }

  function formatClientDateTime(value) {
    const date = value instanceof Date ? value : new Date(value);

    if (isNaN(date.getTime())) {
      return '';
    }

    const dd = String(date.getDate()).padStart(2, '0');
    const mm = String(date.getMonth() + 1).padStart(2, '0');
    const yyyy = String(date.getFullYear());
    const hh = String(date.getHours()).padStart(2, '0');
    const mi = String(date.getMinutes()).padStart(2, '0');
    const ss = String(date.getSeconds()).padStart(2, '0');

    return dd + '/' + mm + '/' + yyyy + ' ' + hh + ':' + mi + ':' + ss;
  }

  function escapeHtml(value) {
    return String(value === undefined || value === null ? '' : value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  function escapeAttr(value) {
    return escapeHtml(value);
  }

  function clearLog() {
    if (el.debugLog) {
      el.debugLog.textContent = '';
    }
  }

  function writeLog(payload) {
    if (!el.debugLog) {
      return;
    }

    const time = formatClientDateTime(new Date());
    const text = '[' + time + '] ' + JSON.stringify(payload, null, 2);

    el.debugLog.textContent = text + '\n\n' + el.debugLog.textContent;
  }

})();
