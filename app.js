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

    userDisplayName: document.getElementById('userDisplayName'),
    userRole: document.getElementById('userRole'),
    systemStatusText: document.getElementById('systemStatusText'),

    healthResult: document.getElementById('healthResult'),
    meResult: document.getElementById('meResult'),
    sourceResult: document.getElementById('sourceResult'),
    dashboardResult: document.getElementById('dashboardResult'),
    previewDashboardBtn: document.getElementById('previewDashboardBtn'),
previewResult: document.getElementById('previewResult'),
    debugLog: document.getElementById('debugLog')
  };

  let currentUser = null;
  let sourcesCache = [];
  let selectedSourceId = '';
  let selectedSheetName = '';
  let lastHeadersData = null;

  document.addEventListener('DOMContentLoaded', init);

  async function init() {
    bindEvents();
    showLogin();
    await checkApiHealth();
    await restoreSession();
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
    if (el.previewDashboardBtn) {
  el.previewDashboardBtn.addEventListener('click', handleDashboardPreview);
}
    if (el.saveMappingBtn) {
      el.saveMappingBtn.addEventListener('click', handleSaveMapping);
      el.saveMappingBtn.disabled = true;
    }
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
      await loadSources();

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
    selectedSourceId = '';
    selectedSheetName = '';
    sourcesCache = [];
    lastHeadersData = null;

    if (el.username) el.username.value = '';
    if (el.password) el.password.value = '';

    renderSources([]);
    renderSourceSheets([]);
    renderHeaders(null);

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
      name: el.sourceName.value.trim(),
      spreadsheetId: el.sourceSpreadsheetId.value.trim(),
      dataType: el.sourceDataType.value,
      owner: el.sourceOwner.value.trim(),
      description: el.sourceDescription.value.trim()
    };

    if (!payload.name || !payload.spreadsheetId) {
      setSourceFormMessage('กรุณากรอกชื่อแหล่งข้อมูลและ Google Sheet ID');
      return;
    }

    setSourceFormMessage('');
    el.createSourceBtn.disabled = true;
    el.createSourceBtn.textContent = 'กำลังเพิ่ม...';

    try {
      const data = await window.AnalyticsAPI.createSource(payload);

      setSourceFormMessage('เพิ่มแหล่งข้อมูลสำเร็จ');

      el.sourceName.value = '';
      el.sourceSpreadsheetId.value = '';
      el.sourceOwner.value = '';
      el.sourceDescription.value = '';

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
      el.createSourceBtn.disabled = false;
      el.createSourceBtn.textContent = 'เพิ่มแหล่งข้อมูล';
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
      const data = await window.AnalyticsAPI.listSourceSheets({
        sourceId
      });

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
              <input
                class="map-display-name"
                type="text"
                value="${escapeAttr(columnName)}"
                placeholder="ชื่อแสดงผล"
              >
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

    el.saveMappingBtn.disabled = true;
    el.saveMappingBtn.textContent = 'กำลังบันทึก...';
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
      el.saveMappingBtn.disabled = false;
      el.saveMappingBtn.textContent = 'บันทึก Mapping';
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
    if (el.loginView) el.loginView.classList.remove('hidden');
    if (el.dashboardView) el.dashboardView.classList.add('hidden');
  }

  function showDashboard() {
    if (el.loginView) el.loginView.classList.add('hidden');
    if (el.dashboardView) el.dashboardView.classList.remove('hidden');
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
    if (value === true) return true;

    const text = String(value == null ? '' : value).trim().toLowerCase();

    return ['true', 'yes', '1', 'active', 'ใช่'].includes(text);
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

  async function handleDashboardPreview() {
  if (!selectedSourceId || !selectedSheetName) {
    setPreviewMessage('กรุณาเลือกแหล่งข้อมูลและชีทก่อน');
    return;
  }

  if (!window.AnalyticsAPI.dashboardPreview) {
    setPreviewMessage('ยังไม่พบฟังก์ชัน dashboardPreview ใน api.js');
    return;
  }

  el.previewDashboardBtn.disabled = true;
  el.previewDashboardBtn.textContent = 'กำลังสร้าง Preview...';

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
    el.previewDashboardBtn.disabled = false;
    el.previewDashboardBtn.textContent = 'สร้าง Dashboard Preview';
  }
}


function renderDashboardPreview(data) {
  if (!el.previewResult) {
    return;
  }

  if (!data || !data.ok) {
    setPreviewMessage((data && data.message) || 'ไม่สามารถสร้าง Preview ได้');
    return;
  }

  el.previewResult.classList.remove('empty');

  const kpisHtml = renderPreviewKpis(data.kpis || []);
  const chartsHtml = renderPreviewCharts(data.charts || []);
  const tableHtml = renderPreviewTable(data.table || { fields: [], rows: [] });

  el.previewResult.innerHTML = `
    <div class="preview-summary">
      <strong>${escapeHtml(data.sourceName || '')}</strong>
      <span> / ${escapeHtml(data.sheetName || '')}</span>
      <span> / อ่านข้อมูล ${Number(data.totalRowsRead || 0).toLocaleString()} แถว</span>
    </div>

    ${kpisHtml}
    ${chartsHtml}
    ${tableHtml}
  `;
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

  return `
    <div class="preview-kpi-grid">
      ${html}
    </div>
  `;
}


function renderPreviewCharts(charts) {
  if (!charts.length) {
    return '';
  }

  const html = charts.map(function (chart) {
    return `
      <div class="preview-chart">
        <h4>${escapeHtml(chart.title || '-')}</h4>
        ${renderSimpleBarChart(chart.data || [])}
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


function setPreviewMessage(message) {
  if (el.previewResult) {
    el.previewResult.textContent = message || '';
    el.previewResult.classList.add('empty');
  }
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
})();
