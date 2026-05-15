document.addEventListener('DOMContentLoaded', async () => {
  bindEvents();

  showBootLoading();

  try {
    const isLoggedIn = await Auth.checkExistingSession();

    if (isLoggedIn) {
      await enterMainApp();
    } else {
      showLogin();
    }
  } catch (err) {
    showLogin();
  }
});

function bindEvents() {
  const loginForm = document.getElementById('loginForm');
  const logoutBtn = document.getElementById('logoutBtn');
  const dashboardSelect = document.getElementById('dashboardSelect');
  const applyFilterBtn = document.getElementById('applyFilterBtn');
  const clearFilterBtn = document.getElementById('clearFilterBtn');
  const refreshBtn = document.getElementById('refreshBtn');
  const tableSearch = document.getElementById('tableSearch');

  if (loginForm) {
    loginForm.addEventListener('submit', handleLoginSubmit);
  }

  if (logoutBtn) {
    logoutBtn.addEventListener('click', handleLogout);
  }

  if (dashboardSelect) {
    dashboardSelect.addEventListener('change', async () => {
      await Dashboard.loadDashboardData();
    });
  }

  if (applyFilterBtn) {
    applyFilterBtn.addEventListener('click', async () => {
      await Dashboard.loadDashboardData();
    });
  }

  if (clearFilterBtn) {
    clearFilterBtn.addEventListener('click', async () => {
      Dashboard.clearFilters();
      await Dashboard.loadDashboardData();
    });
  }

  if (refreshBtn) {
    refreshBtn.addEventListener('click', async () => {
      await Dashboard.loadDashboardData();
    });
  }

  if (tableSearch) {
    tableSearch.addEventListener('input', () => {
      Dashboard.applyTableSearch();
    });
  }
}

async function handleLoginSubmit(event) {
  event.preventDefault();

  const username = document.getElementById('loginUsername')?.value.trim() || '';
  const password = document.getElementById('loginPassword')?.value || '';
  const btn = document.getElementById('loginBtn');

  if (!username || !password) {
    Swal.fire({
      icon: 'warning',
      title: 'กรอกข้อมูลไม่ครบ',
      text: 'กรุณากรอก Username และ Password'
    });
    return;
  }

  setButtonLoading(btn, true, 'กำลังเข้าสู่ระบบ...');

  try {
    await Auth.login(username, password);

    Swal.fire({
      icon: 'success',
      title: 'เข้าสู่ระบบสำเร็จ',
      timer: 900,
      showConfirmButton: false
    });

    await enterMainApp();
  } catch (err) {
    Swal.fire({
      icon: 'error',
      title: 'เข้าสู่ระบบไม่สำเร็จ',
      text: err.message || String(err)
    });
  } finally {
    setButtonLoading(btn, false, 'เข้าสู่ระบบ');
  }
}

async function handleLogout() {
  const confirm = await Swal.fire({
    icon: 'question',
    title: 'ออกจากระบบ?',
    text: 'ต้องการออกจากระบบใช่หรือไม่',
    showCancelButton: true,
    confirmButtonText: 'ออกจากระบบ',
    cancelButtonText: 'ยกเลิก'
  });

  if (!confirm.isConfirmed) return;

  await Auth.logout();
  showLogin();
}

async function enterMainApp() {
  showMain();

  try {
    Dashboard.setDefaultDateThisMonth();
    await Dashboard.loadDashboards();
    await Dashboard.loadDashboardData();
  } catch (err) {
    Swal.fire({
      icon: 'error',
      title: 'โหลดระบบไม่สำเร็จ',
      text: err.message || String(err)
    });
  }
}

function showLogin() {
  document.getElementById('loginView')?.classList.remove('hidden');
  document.getElementById('mainView')?.classList.add('hidden');

  const password = document.getElementById('loginPassword');
  if (password) password.value = '';
}

function showMain() {
  document.getElementById('loginView')?.classList.add('hidden');
  document.getElementById('mainView')?.classList.remove('hidden');
}

function showBootLoading() {
  document.getElementById('loginView')?.classList.add('hidden');
  document.getElementById('mainView')?.classList.add('hidden');
}

function setButtonLoading(btn, loading, text) {
  if (!btn) return;

  btn.disabled = loading;
  btn.textContent = text;
}
