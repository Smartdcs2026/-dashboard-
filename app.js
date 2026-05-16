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

    userDisplayName: document.getElementById('userDisplayName'),
    userRole: document.getElementById('userRole'),
    systemStatusText: document.getElementById('systemStatusText'),

    healthResult: document.getElementById('healthResult'),
    meResult: document.getElementById('meResult'),
    sourceResult: document.getElementById('sourceResult'),
    dashboardResult: document.getElementById('dashboardResult'),
    saveMappingBtn: document.getElementById('saveMappingBtn'),
    mappingMessage: document.getElementById('mappingMessage'),
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
    el.loginForm.addEventListener('submit', handleLogin);
    el.logoutBtn.addEventListener('click', handleLogout);
    el.refreshMeBtn.addEventListener('click', loadMe);

    el.loadSourcesBtn.addEventListener('click', loadSources);
    el.reloadSourcesBtn.addEventListener('click', loadSources);
    el.loadDashboardsBtn.addEventListener('click', loadDashboards);
    el.clearLogBtn.addEventListener('click', clearLog);

    el.sourceForm.addEventListener('submit', handleCreateSource);
    el.saveMappingBtn.addEventListener('click', handleSaveMapping);
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

    el.username.value = '';
    el.password.value = '';

    renderSources([]);
    renderSourceSheets([]);
    renderHeaders(null);

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
    el.sourceResult.textContent = 'กำลังโหลด...';
    el.sourcesList.textContent = 'กำลังโหลด...';
    el.sourcesList.classList.add('empty');

    try {
      const data = await window.AnalyticsAPI.listSources();

      sourcesCache = data.sources || [];

      el.sourceResult.textContent = `พบแหล่งข้อมูล ${data.total || 0} รายการ`;
      renderSources(sourcesCache);

      writeLog({
        step: 'sources',
        response: data
      });

    } catch (error) {
      el.sourceResult.textContent = error.message;
      el.sourcesList.textContent = error.message;
      el.sourcesList.classList.add('empty');

      writeLog({
        step: 'sources_error',
        message: error.message,
        payload: error.payload || null
      });
    }
  }

  function renderSources(sources) {
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
        readHeaders(sourceIdOrSelected(), selectedSheetName);
      });

      el.sourceSheetsList.appendChild(btn);
    });
  }

  function sourceIdOrSelected() {
    return selectedSourceId;
  }

  async function readHeaders(sourceId, sheetName) {
    el.headersResult.textContent = 'กำลังอ่านหัวคอลัมน์...';
    el.headersResult.classList.add('empty');

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
  el.headersResult.innerHTML = '';
  lastHeadersData = data || null;

  if (el.saveMappingBtn) {
    el.saveMappingBtn.disabled = !data || !data.headers;
  }

  if (el.mappingMessage) {
    el.mappingMessage.textContent = '';
  }

  if (!data || !data.headers) {
    el.headersResult.textContent = 'กรุณาเลือกชีทก่อน';
    el.headersResult.classList.add('empty');
    return;
  }

  el.headersResult.classList.remove('empty');

  const wrap = document.createElement('div');
  wrap.className = 'headers-table-wrap';

  const rows = data.headers.map(function (header, idx) {
    const columnName = header.name || '';
    const guessedType = header.guessedType || 'text';
    const guessedMeaning = header.guessedMeaning || 'text';
    const sampleValue = getFirstSampleValue(data.sampleRows || [], idx);

    return `
      <tr data-map-row="${idx}">
        <td>${escapeHtml(header.index)}</td>
        <td>${escapeHtml(header.columnLetter)}</td>
        <td>
          <div class="header-name">${escapeHtml(columnName || '(ไม่มีหัวคอลัมน์)')}</div>
          <input
            class="map-display-name"
            type="text"
            value="${escapeAttr(columnName)}"
            placeholder="ชื่อแสดงผล"
          >
        </td>
        <td>
          <div class="mapping-control">
            <select class="map-data-type">
              ${renderDataTypeOptions(guessedType)}
            </select>
          </div>
        </td>
        <td>
          <div class="mapping-control">
            <select class="map-meaning-type">
              ${renderMeaningTypeOptions(guessedMeaning)}
            </select>
          </div>
        </td>
        <td>
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
            ${renderCheck('useFilter', 'Filter', shouldDefaultFilter(guessedMeaning, guessedType))}
            ${renderCheck('useExport', 'Export', true)}
            ${renderCheck('required', 'บังคับ', false)}
          </div>
        </td>
        <td>
          <input class="map-unit" type="text" placeholder="หน่วย" value="">
        </td>
        <td>
          <input class="map-date-format" type="text" placeholder="dd/MM/yyyy HH:mm:ss" value="${guessedType === 'datetime' ? 'dd/MM/yyyy HH:mm:ss' : ''}">
        </td>
        <td>
          <input class="map-note" type="text" placeholder="หมายเหตุ" value="">
          <input class="map-column-name" type="hidden" value="${escapeAttr(columnName)}">
          <input class="map-sample-value" type="hidden" value="${escapeAttr(sampleValue)}">
        </td>
      </tr>
    `;
  }).join('');

  wrap.innerHTML = `
    <table class="headers-table mapping-table">
      <thead>
        <tr>
          <th>#</th>
          <th>คอลัมน์</th>
          <th>หัวคอลัมน์ / ชื่อแสดงผล</th>
          <th>ประเภทข้อมูล</th>
          <th>ความหมายข้อมูล</th>
          <th>ใช้ทำอะไร</th>
          <th>หน่วย</th>
          <th>รูปแบบวันที่</th>
          <th>หมายเหตุ</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;

  el.headersResult.appendChild(wrap);

  const sampleBox = document.createElement('div');
  sampleBox.className = 'sample-box';

  sampleBox.innerHTML = `
    <h4>ตัวอย่างข้อมูล 20 แถวแรก</h4>
    <pre>${escapeHtml(JSON.stringify(data.sampleRows || [], null, 2))}</pre>
  `;

  el.headersResult.appendChild(sampleBox);
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

  function setSourceFormMessage(message) {
    el.sourceFormMessage.textContent = message || '';
  }

  function setLoading(isLoading) {
    el.loginBtn.disabled = !!isLoading;
    el.loginBtn.textContent = isLoading ? 'กำลังเข้าสู่ระบบ...' : 'เข้าสู่ระบบ';
  }

  function writeLog(data) {
    el.debugLog.textContent = JSON.stringify(data, null, 2);
  }

  function clearLog() {
    el.debugLog.textContent = 'ยังไม่มีข้อมูล';
  }

  function toBool(value) {
    if (value === true) return true;
    const text = String(value == null ? '' : value).trim().toLowerCase();
    return ['true', 'yes', '1', 'active', 'ใช่'].includes(text);
  }

  function escapeHtml(value) {
    return String(value == null ? '' : value)
      .replaceAll('&', '&amp;')
      .replaceAll('<', '&lt;')
      .replaceAll('>', '&gt;')
      .replaceAll('"', '&quot;')
      .replaceAll("'", '&#039;');
  }
})();
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
  const rows = Array.from(el.headersResult.querySelectorAll('tr[data-map-row]'));

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


function setMappingMessage(message) {
  if (el.mappingMessage) {
    el.mappingMessage.textContent = message || '';
  }
}


function escapeAttr(value) {
  return escapeHtml(value).replaceAll('`', '&#096;');
}
