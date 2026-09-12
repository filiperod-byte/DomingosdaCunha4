/* Cliente partilhado: tokens apenas no corpo POST e apenas no endpoint configurado. */
(function () {
  const nativeFetch = window.fetch.bind(window);
  const script = document.currentScript;
  const root = new URL('.', script.src);
  let endpointPromise;
  const storageKey = role => 'dc4_session_' + role;
  const read = role => { try { return JSON.parse(sessionStorage.getItem(storageKey(role)) || 'null'); } catch (_) { return null; } };
  function clear(role) {
    sessionStorage.removeItem(storageKey(role));
    if (role === 'admin') ['dc4_admin_unlocked', 'extintores_bo_unlocked'].forEach(k => sessionStorage.removeItem(k));
  }
  function endpoint() {
    if (!endpointPromise) endpointPromise = nativeFetch(new URL('V2/config.json', root), { cache: 'no-store' }).then(r => r.json()).then(c => new URL(c.backendUrl).href);
    return endpointPromise;
  }
  window.dc4HasSession = role => { const s = read(role); return !!(s && s.expiresAt > Date.now()); };
  window.dc4Logout = async role => {
    const s = read(role); clear(role);
    if (s) try { await nativeFetch(await endpoint(), { method: 'POST', keepalive: true, headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ action: 'auth.logout', token: s.token }) }); } catch (_) { /* Expiração no servidor continua a aplicar-se. */ }
  };
  window.dc4Fetch = async (input, options = {}) => {
    const url = new URL(String(input), location.href);
    if (url.origin !== 'https://script.google.com' || url.pathname !== new URL(await endpoint()).pathname) return nativeFetch(input, options);
    const payload = Object.fromEntries(url.searchParams);
    if ((options.method || 'GET').toUpperCase() === 'GET') payload._method = 'GET';
    else if (options.body instanceof FormData) Object.assign(payload, Object.fromEntries(options.body));
    else if (options.body) Object.assign(payload, JSON.parse(options.body));
    const action = String(payload.action || 'status').replace(/^garage([A-Z])/, (_, c) => 'garage.' + c.toLowerCase());
    const role = ['report', 'garage.getCode', 'garage.loginPin'].includes(action) ? 'resident' : 'admin';
    const session = read(role) || (action === 'report' ? read('admin') : null);
    if (session && session.expiresAt > Date.now()) payload.token = session.token;
    const response = await nativeFetch(await endpoint(), { ...options, method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(payload) });
    const result = await response.clone().json();
    if (result.token && ['admin', 'resident'].includes(result.role)) sessionStorage.setItem(storageKey(result.role), JSON.stringify({ token: result.token, expiresAt: result.expiresAt }));
    if (result.code === 'AUTH_REQUIRED') {
      clear(role);
      if (role === 'admin') location.assign(new URL('V2/admin.html', root));
      throw new Error(result.message || 'Volte a entrar na app.');
    }
    return response;
  };
})();
