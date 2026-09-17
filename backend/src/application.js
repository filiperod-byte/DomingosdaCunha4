// Composição dos módulos e contrato de pedidos independente do transporte HTTP.
// As ações garage.* são aliases de compatibilidade, não a organização interna da app.
function createCondominiumApplication_(ports, config, domain) {
  const notifications = createNotificationService_(ports, config, domain);
  function adminEmail(payload) {
    const candidate = domain.safeText_(payload && payload.adminEmail);
    if (domain.isRealEmail_(candidate)) return candidate;
    if (domain.isRealEmail_(config.ADMIN_EMAIL)) return config.ADMIN_EMAIL;
    const effective = domain.safeText_(ports.environment.effectiveEmail());
    if (domain.isRealEmail_(effective)) return effective;
    const active = domain.safeText_(ports.environment.activeEmail());
    return domain.isRealEmail_(active) ? active : '';
  }
  const occurrences = createOccurrenceService_(ports, config, domain, notifications, adminEmail);
  const general = createGeneralOccurrenceService_(ports);
  const pins = createLegacyPinService_(ports, config, domain, notifications, adminEmail);
  const residents = createResidentService_(ports, domain, notifications);
  const accesses = createAccessService_(ports);
  const setting = key => ports.settings.get(key);
  const routes = { GET: Object.create(null), POST: Object.create(null) };

  function route(method, names, handler) {
    names.split(' ').forEach(name => routes[method][name] = handler);
  }
  route('GET','general.status',general.list);
  route('GET','general.admin',general.adminList);
  route('POST','general.report',general.report);
  route('POST','general.update',general.update);
  route('POST','general.confirm',general.confirm);
  route('GET', 'status', occurrences.status);
  route('GET', 'pinStatus', pins.status);
  route('GET', 'openOccurrences', occurrences.open);
  route('GET', 'pendingOccurrences', occurrences.pending);
  route('GET', 'health', () => ({
    success: true, appName: config.APP_NAME, spreadsheetId: config.SPREADSHEET_ID,
    rootFolderId: config.ROOT_FOLDER_ID, adminEmail: adminEmail({}), pinConfigured: pins.configured(),
    garageConfigReady: !!setting('NOME_CONDOMINIO'), now: ports.clock.now().toISOString()
  }));
  route('GET', 'garage.publicConfig garagePublicConfig', () => ({
    nomeCondominio: setting('NOME_CONDOMINIO') || 'Condomínio',
    tempoVisivel: parseInt(setting('TEMPO_VISIVEL_SEGUNDOS')) || 15,
    adminEmail: setting('ADMIN_EMAIL') || '', structure: domain.getGarageStructure_().floors
  }));
  route('GET', 'garage.structure garageStructure', domain.getGarageStructure_);
  route('GET', 'garage.dashboard garageDashboard', accesses.dashboard);
  route('GET', 'garage.pending garagePending', residents.pending);
  route('GET', 'garage.approved garageApproved', residents.approved);
  route('GET', 'garage.history garageHistory', () => accesses.history(50));
  route('GET', 'garage.adminConfig garageAdminConfig', () => ({
    nomeCondominio: setting('NOME_CONDOMINIO'), adminEmail: setting('ADMIN_EMAIL'),
    codigoAtual: setting('CODIGO_CADEADO'), tempoVisivel: setting('TEMPO_VISIVEL_SEGUNDOS'),
    maxTentativas: setting('MAX_TENTATIVAS_PIN'), bloqueioMinutos: setting('BLOQUEIO_MINUTOS'), textoAviso: setting('TEXTO_AVISO')
  }));
  route('POST', 'report', occurrences.report);
  route('POST', 'approveOccurrence', occurrences.approve);
  route('POST', 'rejectOccurrence', occurrences.reject);
  route('POST', 'closeOccurrence', occurrences.close);
  route('POST', 'setPin', pins.set);
  route('POST', 'validatePin', pins.validate);
  route('POST', 'resetPin', pins.reset);
  route('POST', 'garage.register garageRegister', p => residents.register({
    nome: p.nome || p.name, piso: p.piso || p.floor, fracao: p.fracao || p.fraction,
    email: p.email, telemovel: p.telemovel || p.telefone || p.phone
  }));
  route('POST', 'garage.loginPin garageLoginPin', p => residents.login(p.pin));
  route('POST', 'garage.failedAttempt garageFailedAttempt', residents.failedAttempt);
  route('POST', 'garage.resendPin garageResendPin garage.recoverCode garageRecoverCode', p => residents.resendPin(p.email));
  route('POST', 'garage.getCode garageGetCode', p => accesses.getCode(p.pin, p.motivo || p.reason, p.motivoOutro || p.reasonOther || '', p.userAgent || p.ua || ''));
  route('POST', 'garage.loginAdmin garageLoginAdmin', p => residents.loginAdmin(p.email, p.pin));
  route('POST', 'garage.approve garageApprove', p => residents.approve(Number(p.row), p.adminEmail || p.email || 'Admin'));
  route('POST', 'garage.reject garageReject', p => residents.reject(Number(p.row)));
  route('POST', 'garage.block garageBlock', p => residents.block(Number(p.row)));
  route('POST', 'garage.unblock garageUnblock', p => residents.unblock(Number(p.row)));
  route('POST', 'garage.regeneratePin garageRegeneratePin', p => residents.regeneratePin(Number(p.row)));
  route('POST', 'garage.changeCode garageChangeCode', p => accesses.changeCode(p.codigo || p.code || p.novoCodigo));
  route('POST', 'garage.saveConfig garageSaveConfig', p => {
    const configs = p.configs && typeof p.configs === 'object' ? p.configs : p;
    Object.keys(configs).forEach(key => { if (configs[key] !== '') ports.settings.set(key, configs[key]); });
    return { ok: true, success: true };
  });

  function dispatch(method, action, payload) {
    try {
      const handler = routes[method] && routes[method][action];
      if (!handler) return { success: false, message: 'Ação ' + method + ' inválida: ' + action };
      return handler(payload || {});
    } catch (err) {
      return { success: false, message: err && err.message ? err.message : 'Erro interno no backend' };
    }
  }
  return { dispatch: dispatch, initialize: () => ports.initialize(), clearLegacyPin: pins.clear };
}
