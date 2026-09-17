/* Cliente partilhado: tokens apenas no corpo POST e apenas no endpoint configurado. */
(function () {
  const nativeFetch = window.fetch.bind(window);
  const script = document.currentScript;
  const root = new URL('.', script.src);
  let endpointPromise;
  let sessionGeneration=0;
  const measurements=[];
  window.dc4Diagnostics=()=>measurements.map(x=>({...x}));
  const storageKey = role => 'dc4_session_' + role;
  const read = role => { try { return JSON.parse(sessionStorage.getItem(storageKey(role)) || 'null'); } catch (_) { return null; } };
  function clear(role) {
    sessionGeneration++;
    sessionStorage.removeItem(storageKey(role));
    if (role === 'admin') ['dc4_admin_unlocked', 'extintores_bo_unlocked'].forEach(k => sessionStorage.removeItem(k));
  }
  function endpoint() {
    if (!endpointPromise) endpointPromise = nativeFetch(new URL('V2/config.json', root), { cache: 'no-store' }).then(r => r.json()).then(c => new URL(c.backendUrl).href).catch(error=>{endpointPromise=null;throw error});
    return endpointPromise;
  }
  window.dc4HasSession = role => { const s = read(role); return !!(s && s.expiresAt > Date.now()); };
  window.dc4Logout = async role => {
    const s = read(role); clear(role);
    if (s) try { await nativeFetch(await endpoint(), { method: 'POST', keepalive: true, headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ action: 'auth.logout', token: s.token }) }); } catch (_) { /* Expiração no servidor continua a aplicar-se. */ }
  };
  const request = async (input, options = {}) => {
    const generation=sessionGeneration;
    const url = new URL(String(input), location.href);
    if (url.origin !== 'https://script.google.com' || url.pathname !== new URL(await endpoint()).pathname) return nativeFetch(input, options);
    const payload = Object.fromEntries(url.searchParams);
    if ((options.method || 'GET').toUpperCase() === 'GET') payload._method = 'GET';
    else if (options.body instanceof FormData) Object.assign(payload, Object.fromEntries(options.body));
    else if (options.body) Object.assign(payload, JSON.parse(options.body));
    const action = String(payload.action || 'status').replace(/^garage([A-Z])/, (_, c) => 'garage.' + c.toLowerCase());
    payload.action = action;
    const adminPage = /\/(?:V2\/(?:admin|acessos-admin|config-admin|garagem-admin)|backoffice|relatorio-extintores|occurrences-admin|floor-qrcodes)\.html$/.test(new URL(location.href).pathname);
    const role = ['general.report','general.confirm','report', 'garage.getCode', 'garage.loginPin'].includes(action) ? 'resident' : 'admin';
    const session = read(role) || (action === 'report' ? read('admin') : null);
    if (session && session.expiresAt > Date.now()) payload.token = session.token;
    const startedAt=Date.now();
    const response = await nativeFetch(await endpoint(), { ...options, method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(payload) });
    let result;
    try{result=await response.clone().json()}catch(_){throw new Error('O serviço devolveu uma resposta inválida. O resultado do pedido não foi confirmado.')}
    measurements.push({action,totalMs:Date.now()-startedAt,serverMs:result.serverDurationMs??null,serviceVersion:result.serviceVersion||'não identificada'});
    if(measurements.length>20)measurements.shift();
    if(!response.ok)throw new Error('Falha do serviço ('+response.status+'). O resultado do pedido não foi confirmado.');
    if (generation === sessionGeneration && result.token && ['admin', 'resident'].includes(result.role)) sessionStorage.setItem(storageKey(result.role), JSON.stringify({ token: result.token, expiresAt: result.expiresAt }));
    if (result.code === 'AUTH_REQUIRED') {
      clear(role);
      if (role === 'admin' && adminPage) location.assign(new URL('V2/admin.html', root));
      const error = new Error(result.message || 'Volte a entrar na app.');
      error.code = result.code;
      throw error;
    }
    return response;
  };
  const pendingReads=new Map();
  window.dc4Fetch=async(input,options={})=>{
    const url=new URL(String(input),location.href);
    const action=url.searchParams.get('action')||'status';
    const publicRead=(options.method||'GET').toUpperCase()==='GET' &&
      url.origin==='https://script.google.com' &&
      ['general.status','status','garage.publicConfig','garage.structure'].includes(action);
    if(!publicRead)return request(input,options);
    const key=url.href;
    if(!pendingReads.has(key)){
      const promise=request(input,options).finally(()=>pendingReads.delete(key));
      pendingReads.set(key,promise);
    }
    return (await pendingReads.get(key)).clone();
  };
})();
