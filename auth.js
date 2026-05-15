(function () {
  'use strict';

  let currentUser = null;

  function getApi_() {
    if (!window.Api) {
      throw new Error('ไม่พบ api.js หรือโหลด api.js ไม่สำเร็จ');
    }
    return window.Api;
  }

  function getUser() {
    return currentUser;
  }

  function setUser(user) {
    currentUser = user || null;

    const userName = document.getElementById('userName');
    const userRole = document.getElementById('userRole');

    if (userName) userName.textContent = currentUser ? currentUser.name : '-';
    if (userRole) userRole.textContent = currentUser ? currentUser.roleId : '-';
  }

  async function checkExistingSession() {
    const Api = getApi_();
    const token = Api.getToken();

    if (!token) {
      setUser(null);
      return false;
    }

    try {
      const result = await Api.authCheck();

      if (!result.valid) {
        Api.clearToken();
        setUser(null);
        return false;
      }

      setUser(result.user);
      return true;
    } catch (err) {
      Api.clearToken();
      setUser(null);
      return false;
    }
  }

  async function login(username, password) {
    const Api = getApi_();
    const result = await Api.login(username, password);

    if (!result.token) {
      throw new Error('ไม่พบ token จากระบบ');
    }

    Api.setToken(result.token);
    setUser(result.user);

    return result;
  }

  async function logout() {
    const Api = getApi_();

    try {
      await Api.logout();
    } catch (err) {}

    Api.clearToken();
    setUser(null);
  }

  window.Auth = {
    getUser,
    setUser,
    checkExistingSession,
    login,
    logout
  };
})();
