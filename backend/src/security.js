// Política de acesso aplicada a TODAS as entradas HTTP. Sem dependências Google.
function createSecureApplication_(app, ports) {
  const hash = value => ports.crypto.hashPin(String(value), 'dc4-session-v1');
  const now = () => ports.clock.now().getTime();
  const read = key => JSON.parse(ports.properties.getProperty(key) || 'null');
  const write = (key, value) => ports.properties.setProperties({ [key]: JSON.stringify(value) }, false);
  const canonical = action => action.replace(/^garage([A-Z])/, (_, c) => 'garage.' + c.toLowerCase());
  const deny = (message, code) => ({ ok: false, success: false, valid: false, msg: message, message, code: code || 'AUTH_REQUIRED' });
  const active = r => r && r.status === 'APROVADO' && String(r.pinActive).toLowerCase() === 'true' && !(new Date(r.blockedUntil).getTime() > now());
  function fingerprint(role, resident) {
    return hash(role === 'admin' ? ports.settings.get('ADMIN_EMAIL') + ':' + ports.settings.get('ADMIN_PIN') : resident.pin + ':' + resident.status + ':' + resident.pinActive);
  }
  function issue(role, resident) {
    const id = role === 'admin' ? 'admin' : String(resident.id);
    const subject = hash(role + ':' + id);
    const token = subject + '.' + ports.crypto.uuid() + ports.crypto.uuid();
    const expiresAt = now() + (role === 'admin' ? 30 * 60 : 8 * 60 * 60) * 1000;
    write('DC4_SESSION_' + subject, { digest: hash(token), role, id, expiresAt, fingerprint: fingerprint(role, resident) });
    return { token, expiresAt, role };
  }
  function verify(token) {
    if (typeof token !== 'string' || !/^[a-f0-9]{64}\.[a-f0-9-]{72}$/.test(token)) return null;
    const key = 'DC4_SESSION_' + token.split('.')[0];
    const session = read(key);
    if (!session || session.expiresAt <= now() || session.digest !== hash(token)) return null;
    const resident = session.role === 'resident' ? ports.residents.list().find(r => String(r.id) === session.id) : null;
    if (session.role === 'resident' && !active(resident)) return null;
    if (session.fingerprint !== fingerprint(session.role, resident)) return null;
    return { ...session, resident, key };
  }
  function throttle(bucket, max, duration) {
    return ports.lock.run(() => {
      const key = 'DC4_RATE_' + bucket;
      let state = read(key);
      if (!state || state.until <= now()) state = { count: 0, until: now() + duration };
      if (state.count >= max) return false;
      state.count++; write(key, state); return true;
    });
  }
  const publicGet = new Set(['status', 'garage.publicConfig', 'garage.structure']);
  const adminGet = new Set(['health', 'openOccurrences', 'pendingOccurrences', 'garage.dashboard', 'garage.pending', 'garage.approved', 'garage.history', 'garage.adminConfig']);
  const adminPost = new Set(['approveOccurrence', 'rejectOccurrence', 'closeOccurrence', 'garage.approve', 'garage.reject', 'garage.block', 'garage.unblock', 'garage.regeneratePin', 'garage.changeCode', 'garage.saveConfig']);
  function dispatch(method, action, input) {
    try {
      action = canonical(action);
      const p = Object.assign({}, input || {});
      // O transporte POST permite consultas autenticadas sem colocar tokens no URL.
      if (method === 'POST' && p._method === 'GET') method = 'GET';
      if (method === 'GET' && publicGet.has(action)) return app.dispatch(method, action, p);
      if (method === 'POST' && ['garage.loginAdmin', 'garage.loginPin'].includes(action)) {
        const timings = {};
        let checkpoint = now();
        const mark = key => { const t = now(); timings[key] = Math.max(0, t - checkpoint); checkpoint = t; };
        const admin = action === 'garage.loginAdmin';
        if (!throttle(admin ? 'admin-login' : 'resident-login', admin ? 10 : 30, 15 * 60 * 1000)) return deny('Demasiadas tentativas. Aguarde 15 minutos.', 'RATE_LIMITED');
        mark('rateLimitMs');
        const pin = String(p.pin || '');
        if (!/^\d{6,12}$/.test(pin) || (admin && ['123456', '000000'].includes(pin))) return deny('Credenciais inválidas.', 'INVALID_CREDENTIALS');
        const resident = admin ? null : ports.residents.list().find(r => String(r.pin) === pin);
        if (admin ? !ports.settings.get('ADMIN_PIN') || !ports.settings.get('ADMIN_EMAIL') : !active(resident)) return deny('Credenciais inválidas ou acesso indisponível.', 'INVALID_CREDENTIALS');
        mark('validationMs');
        const result = app.dispatch(method, action, p);
        mark('accessUpdateMs');
        if (!(result.ok || result.success)) return deny('Credenciais inválidas.', 'INVALID_CREDENTIALS');
        const session = issue(admin ? 'admin' : 'resident', resident);
        mark('sessionMs');
        return Object.assign({}, result, session, { loginTimingsMs: timings });
      }
      if (method === 'POST' && ['garage.register', 'garage.resendPin', 'garage.recoverCode'].includes(action)) {
        if (!throttle('public-mail', 10, 60 * 60 * 1000)) return deny('Limite de pedidos atingido. Tente mais tarde.', 'RATE_LIMITED');
        if (action !== 'garage.register') {
          const resident = ports.residents.list().find(r => String(r.email).trim().toLowerCase() === String(p.email || '').trim().toLowerCase());
          if (active(resident)) app.dispatch('POST', 'garage.resendPin', { email: resident.email });
          return { ok: true, success: true, msg: 'Se existir uma conta aprovada, receberá o PIN no email registado.' };
        }
        return app.dispatch(method, action, p);
      }
      // A criação/reset do PIN antigo deixa de estar exposta por HTTP.
      if (['pinStatus', 'setPin', 'validatePin', 'resetPin'].includes(action)) return deny('Entre pela área de administração da app.', 'LEGACY_LOGIN_DISABLED');
      const session = verify(p.token);
      if (!session) return deny('Sessão terminada ou inválida. Volte a entrar.');
      if (method === 'POST' && action === 'auth.logout') {
        ports.properties.deleteProperty(session.key); return { ok: true, success: true };
      }
      if (method === 'POST' && action === 'auth.session') return { ok: true, success: true, role: session.role, expiresAt: session.expiresAt };
      if (adminGet.has(action) && method === 'GET' || adminPost.has(action) && method === 'POST') {
        if (session.role !== 'admin') return deny('Acesso reservado à administração.', 'FORBIDDEN');
        p.adminEmail = ports.settings.get('ADMIN_EMAIL');
        if (action === 'garage.saveConfig') {
          const allowed = new Set(['NOME_CONDOMINIO', 'ADMIN_EMAIL', 'ADMIN_PIN', 'CODIGO_CADEADO', 'TEMPO_VISIVEL_SEGUNDOS', 'MAX_TENTATIVAS_PIN', 'BLOQUEIO_MINUTOS', 'TEXTO_AVISO']);
          const configs = p.configs || {};
          if (Object.keys(configs).some(k => !allowed.has(k))) return deny('Configuração não permitida.', 'INVALID_CONFIG');
          if (configs.ADMIN_PIN && (!/^\d{6,12}$/.test(String(configs.ADMIN_PIN)) || ['123456', '000000'].includes(String(configs.ADMIN_PIN)))) return deny('O PIN administrativo deve ter pelo menos 6 dígitos.', 'INVALID_CONFIG');
          p.configs = configs;
        }
        return app.dispatch(method, action, p);
      }
      if (method === 'POST' && action === 'garage.getCode' && session.role === 'resident') {
        p.pin = session.resident.pin;
        return app.dispatch(method, action, p);
      }
      if (method === 'POST' && action === 'report') {
        p.name = p.reportedBy = session.role === 'resident' ? session.resident.name : 'Administração';
        p.adminEmail = ports.settings.get('ADMIN_EMAIL');
        return app.dispatch(method, action, p);
      }
      return deny('Ação não autorizada.', 'FORBIDDEN');
    } catch (err) { return deny('Não foi possível processar o pedido.', 'SERVER_ERROR'); }
  }
  return { dispatch, initialize: app.initialize, clearLegacyPin: app.clearLegacyPin };
}
