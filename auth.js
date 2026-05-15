const Auth = (() => {
  let currentUser = null;

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
    const result = await Api.login(username, password);

    if (!result.token) {
      throw new Error('ไม่พบ token จากระบบ');
    }

    Api.setToken(result.token);
    setUser(result.user);

    return result;
  }

  async function logout() {
    try {
      await Api.logout();
    } catch (err) {
      // ไม่ต้อง block logout ถ้า API error
    }

    Api.clearToken();
    setUser(null);
  }

  return {
    getUser,
    setUser,
    checkExistingSession,
    login,
    logout
  };
})();
