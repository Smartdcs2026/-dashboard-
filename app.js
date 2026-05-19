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
exportDashboardBtn: document.getElementById('exportDashboardBtn'),
dashboardViewMessage: document.getElementById('dashboardViewMessage'),
    dashboardViewResult: document.getElementById('dashboardViewResult'),
        reloadManageDashboardsBtn: document.getElementById('reloadManageDashboardsBtn'),
    manageDashboardSelect: document.getElementById('manageDashboardSelect'),
    manageDashboardName: document.getElementById('manageDashboardName'),
    manageDashboardType: document.getElementById('manageDashboardType'),
    manageDashboardDescription: document.getElementById('manageDashboardDescription'),
    manageDashboardPublish: document.getElementById('manageDashboardPublish'),
    manageDashboardExport: document.getElementById('manageDashboardExport'),
    manageDashboardHidden: document.getElementById('manageDashboardHidden'),
    saveManageDashboardBtn: document.getElementById('saveManageDashboardBtn'),
    hideManageDashboardBtn: document.getElementById('hideManageDashboardBtn'),
    deleteManageDashboardBtn: document.getElementById('deleteManageDashboardBtn'),
    manageDashboardMessage: document.getElementById('manageDashboardMessage'),
    manageDashboardSummary: document.getElementById('manageDashboardSummary'),

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

    if (el.resetDashboardFilterBtn) {
      el.resetDashboardFilterBtn.addEventListener('click', handleResetDashboardFilter);
    }
     if (el.exportDashboardBtn) {
  el.exportDashboardBtn.addEventListener('click', handleExportDashboard);
}
        if (el.reloadManageDashboardsBtn) {
      el.reloadManageDashboardsBtn.addEventListener('click', loadManageDashboards);
    }

    if (el.manageDashboardSelect) {
      el.manageDashboardSelect.addEventListener('change', handleSelectManageDashboard);
    }

    if (el.saveManageDashboardBtn) {
      el.saveManageDashboardBtn.addEventListener('click', handleSaveManageDashboard);
    }

    if (el.hideManageDashboardBtn) {
      el.hideManageDashboardBtn.addEventListener('click', handleToggleManageDashboardHidden);
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
      if (el.sourceResult) {
        el.sourceResult.textContent = error.message;
      }

      if (el.sourcesList) {
        el.sourcesList.textContent = error.message;
        el.sourcesList.classList.add('empty');
      }

      writeLog({
        step: 'sources_error',
        message: error.message,
        payload: error.payload || null
      });
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

    el.sourceSheetsList.textContent = 'กำลังอ่านรายชื่อชีท...';
    el.sourceSheetsList.classList.add('empty');

    try {
      const data = await window.AnalyticsAPI.listSourceSheets({ sourceId });

      renderSourceSheets(data.sheets || []);

      writeLog({
        step: 'source_sheets',
        response: data
      });

    } catch (error) {
      el.sourceSheetsList.textContent = error.message;
      el.sourceSheetsList.classList.add('empty');

      writeLog({
        step: 'source_sheets_error',
        message: error.message,
        payload: error.payload || null
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

      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'sheet-item' + (sheetName === selectedSheetName ? ' active' : '');

      btn.innerHTML = `
        <div class="item-title">
          <span>${escapeHtml(sheetName)}</span>
          <span class="badge badge-muted">${Number(sheet.lastRow || 0).toLocaleString()} rows</span>
        </div>
        <div class="item-meta">
          <div>คอลัมน์: ${Number(sheet.lastColumn || 0).toLocaleString()}</div>
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

    if (el.saveMappingBtn) {
      el.saveMappingBtn.disabled = true;
    }

    try {
      const data = await window.AnalyticsAPI.readHeaders({
        sourceId,
        sheetName,
        headerRow: 1
      });

      renderHeaders(data);

      writeLog({
        step: 'headers',
        response: data
      });

    } catch (error) {
      el.headersResult.textContent = error.message;
      el.headersResult.classList.add('empty');

      writeLog({
        step: 'headers_error',
        message: error.message,
        payload: error.payload || null
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
      setMappingMessage(error.message);

      writeLog({
        step: 'save_mapping_error',
        message: error.message,
        payload: error.payload || null
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
      setPreviewMessage(error.message);

      writeLog({
        step: 'dashboard_preview_error',
        message: error.message,
        payload: error.payload || null
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

  const kpisHtml = renderPreviewKpis({
    kpis: data.kpis || []
  });

  const chartsHtml = renderPreviewCharts({
    charts: data.chartResults || data.charts || []
  });

  const tableHtml = renderPreviewTable(data.table || {
    fields: [],
    rows: []
  });

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
    ${tableHtml || '<div class="preview-chart-empty">ยังไม่มีข้อมูลตาราง</div>'}
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

    try {
      const data = await window.AnalyticsAPI.dashboardView({
        dashboardId: dashboardId,
        limit: limit,
        filters: []
      });

      renderSavedDashboard(data);
      renderDashboardFilters(data.filters || []);
      setDashboardViewMessage('โหลด Dashboard สำเร็จ');

      writeLog({
        step: 'dashboard_view',
        response: data
      });

    } catch (error) {
      setDashboardViewMessage(error.message);

      if (el.dashboardViewResult) {
        el.dashboardViewResult.textContent = error.message;
        el.dashboardViewResult.classList.add('empty');
      }

      writeLog({
        step: 'dashboard_view_error',
        message: error.message,
        payload: error.payload || null
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

    try {
      const data = await window.AnalyticsAPI.dashboardView({
        dashboardId: dashboardId,
        limit: limit,
        filters: filters
      });

      renderSavedDashboard(data);
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
      setDashboardViewMessage(error.message);

      writeLog({
        step: 'dashboard_filter_error',
        message: error.message,
        payload: error.payload || null
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
    if (el.previewResult && message) {
      el.previewResult.textContent = message;
      el.previewResult.classList.add('empty');
    }
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
    totalExportRows: data.totalExportRows,
    totalRowsAfterFilter: data.totalRowsAfterFilter,
    csv: '[CSV_CONTENT_HIDDEN]'
  }
});

  } catch (error) {
    setDashboardViewMessage(error.message);

    writeLog({
      step: 'dashboard_export_error',
      message: error.message,
      payload: error.payload || null
    });

  } finally {
    setButtonLoading(el.exportDashboardBtn, false, 'Export CSV');
  }
}

  async function loadManageDashboards() {
    if (!el.manageDashboardSelect) {
      return;
    }

    setManageDashboardMessage('');
    selectedManageDashboard = null;

    el.manageDashboardSelect.innerHTML = '<option value="">กำลังโหลด Dashboard...</option>';
    resetManageDashboardForm(false);

    try {
      const data = await window.AnalyticsAPI.listDashboards();
      manageDashboardsCache = data.dashboards || [];

      renderManageDashboardOptions(manageDashboardsCache);

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


  function renderManageDashboardOptions(dashboards) {
    if (!el.manageDashboardSelect) {
      return;
    }

    if (!dashboards || !dashboards.length) {
      el.manageDashboardSelect.innerHTML = '<option value="">ยังไม่มี Dashboard</option>';
      return;
    }

    el.manageDashboardSelect.innerHTML =
      '<option value="">เลือก Dashboard ที่ต้องการจัดการ</option>' +
      dashboards.map(function (dash) {
        const id = String(dash['รหัส Dashboard'] || '').trim();
        const name = String(dash['ชื่อ Dashboard'] || id).trim();
        const type = String(dash['ประเภท Dashboard'] || '').trim();
        const publish = toBool(dash['สถานะ Publish']) ? 'เผยแพร่' : 'ไม่เผยแพร่';

        return (
          '<option value="' + escapeAttr(id) + '">' +
          escapeHtml(name) +
          (type ? ' (' + escapeHtml(type) + ')' : '') +
          ' - ' + publish +
          '</option>'
        );
      }).join('');
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

    setManageDashboardRoles('');

    if (el.manageDashboardSummary) {
      el.manageDashboardSummary.classList.add('empty');
      el.manageDashboardSummary.textContent = 'เลือก Dashboard เพื่อดูรายละเอียดและแก้ไข';
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
  if (!window.echarts) {
    mountFallbackCharts();
    return;
  }

  requestAnimationFrame(function () {
    chartRenderQueue.forEach(function (item) {
      const dom = document.getElementById(item.id);

      if (!dom) {
        return;
      }

      const chart = window.echarts.init(dom);
      const option = buildEchartOption(item.chart);

      chart.setOption(option);
      chartInstances.push(chart);
    });

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


function buildEchartOption(chart) {
  const type = String(chart.type || 'bar').toLowerCase();
  const data = Array.isArray(chart.data) ? chart.data : [];

  if (type === 'line') {
    return buildLineChartOption(chart, data);
  }

  if (type === 'donut' || type === 'pie') {
    return buildDonutChartOption(chart, data);
  }

  if (type === 'horizontal_bar') {
    return buildHorizontalBarOption(chart, data);
  }

  return buildBarChartOption(chart, data);
}


function buildLineChartOption(chart, data) {
  return {
    tooltip: {
      trigger: 'axis'
    },
    grid: {
      left: 45,
      right: 24,
      top: 28,
      bottom: 42
    },
    xAxis: {
      type: 'category',
      data: data.map(function (x) { return x.name; }),
      axisLabel: {
        rotate: data.length > 8 ? 35 : 0
      }
    },
    yAxis: {
      type: 'value'
    },
    series: [
      {
        name: chart.title || '',
        type: 'line',
        smooth: true,
        symbolSize: 7,
        areaStyle: {
          opacity: 0.12
        },
        data: data.map(function (x) { return Number(x.value || 0); })
      }
    ]
  };
}


function buildBarChartOption(chart, data) {
  return {
    tooltip: {
      trigger: 'axis',
      axisPointer: {
        type: 'shadow'
      }
    },
    grid: {
      left: 45,
      right: 24,
      top: 28,
      bottom: 72
    },
    xAxis: {
      type: 'category',
      data: data.map(function (x) { return x.name; }),
      axisLabel: {
        rotate: data.length > 5 ? 35 : 0,
        overflow: 'truncate'
      }
    },
    yAxis: {
      type: 'value'
    },
    series: [
      {
        name: chart.title || '',
        type: 'bar',
        barMaxWidth: 42,
        data: data.map(function (x) { return Number(x.value || 0); })
      }
    ]
  };
}


function buildHorizontalBarOption(chart, data) {
  const sorted = data.slice().reverse();

  return {
    tooltip: {
      trigger: 'axis',
      axisPointer: {
        type: 'shadow'
      }
    },
    grid: {
      left: 90,
      right: 28,
      top: 24,
      bottom: 28
    },
    xAxis: {
      type: 'value'
    },
    yAxis: {
      type: 'category',
      data: sorted.map(function (x) { return x.name; }),
      axisLabel: {
        overflow: 'truncate',
        width: 80
      }
    },
    series: [
      {
        name: chart.title || '',
        type: 'bar',
        barMaxWidth: 24,
        data: sorted.map(function (x) { return Number(x.value || 0); })
      }
    ]
  };
}


function buildDonutChartOption(chart, data) {
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
        name: chart.title || '',
        type: 'pie',
        radius: ['45%', '70%'],
        center: ['50%', '45%'],
        avoidLabelOverlap: true,
        label: {
          formatter: '{b}: {c}'
        },
        data: data.map(function (x) {
          return {
            name: x.name || '(ว่าง)',
            value: Number(x.value || 0)
          };
        })
      }
    ]
  };
}


function mountFallbackCharts() {
  chartRenderQueue.forEach(function (item) {
    const dom = document.getElementById(item.id);

    if (!dom) {
      return;
    }

    dom.outerHTML = renderSimpleBarChart(item.chart.data || []);
  });
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
})();
