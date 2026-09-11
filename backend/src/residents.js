// Identidade e ciclo de vida dos moradores: não pertencem ao módulo de garagem.
// Mantém as respostas e regras legadas. Ver as limitações de segurança em README.md.
function createResidentService_(ports, domain, notifications) {
  const { validateGarageResident_ } = domain;
  const repo = ports.residents;
  const setting = key => ports.settings.get(key);
  const byPin = pin => repo.list().find(c => String(c.pin) === String(pin)) || null;
  const dateText = date => ports.clock.format(date, 'dd/MM/yyyy HH:mm');

  function uniquePin() {
    const pins = repo.list().map(c => String(c.pin));
    for (let attempt = 0; attempt < 100; attempt++) {
      const pin = ports.ids.pinCandidate();
      if (pins.indexOf(pin) === -1 && pin !== setting('ADMIN_PIN')) return pin;
    }
    throw new Error('Não foi possível gerar PIN único.');
  }

  function register(d) {
    try {
      if (!d || !d.nome) throw new Error('Nome obrigatório.');
      if (!d.email) throw new Error('Email obrigatório.');
      const location = validateGarageResident_(d.piso, d.fracao);
      d = Object.assign({}, d, { piso: location.piso, fracao: location.fracao });
      const email = String(d.email).toLowerCase().trim();
      for (const c of repo.list()) {
        if (String(c.email).toLowerCase().trim() !== email) continue;
        const status = String(c.status || '').toUpperCase();
        if (status === 'APROVADO') return { ok: false, success: false, tipo: 'email_existente', recoveryAllowed: true, msg: 'Este email já tem acesso aprovado. Pode recuperar o PIN.' };
        if (status === 'PENDENTE') return { ok: false, success: false, tipo: 'pendente', msg: 'Já existe um pedido pendente com este email.' };
        if (status === 'BLOQUEADO') return { ok: false, success: false, tipo: 'bloqueado', msg: 'Este email está bloqueado. Contacte a administração.' };
        if (status === 'REJEITADO') return { ok: false, success: false, tipo: 'rejeitado', msg: 'Este email já teve um pedido não aprovado. Contacte a administração.' };
      }
      repo.add({
        id: ports.ids.entity('COND'), registeredAt: ports.clock.now(), name: d.nome,
        floor: d.piso, fraction: d.fracao, email: d.email, phone: d.telemovel || '',
        status: 'PENDENTE', pin: '', pinActive: false, approvedAt: '', approvedBy: '',
        lastAccess: '', notes: '', failedAttempts: 0, blockedUntil: ''
      });
      ports.log.add('INFO', 'REGISTO', d.email, 'Novo pedido: ' + d.nome);
      notifications.emailRegisto(d.email, d.nome);
      notifications.emailAdminNovoPedido(d);
      return { ok: true, success: true };
    } catch (err) {
      return { ok: false, success: false, msg: 'Erro ao registar: ' + err.message, message: 'Erro ao registar: ' + err.message };
    }
  }

  function login(pin) {
    try {
      const c = byPin(pin);
      if (!c) return { ok: false, tipo: 'invalido', msg: 'PIN inválido.' };
      if (c.status === 'PENDENTE') return { ok: false, tipo: 'pendente', msg: 'O seu pedido ainda está pendente.' };
      if (c.status === 'REJEITADO') return { ok: false, tipo: 'rejeitado', msg: 'Pedido rejeitado.' };
      if (c.status === 'BLOQUEADO') return { ok: false, tipo: 'bloqueado', msg: 'Acesso bloqueado.' };
      if (c.status === 'APROVADO') {
        repo.update(c.legacyRow, { lastAccess: ports.clock.now(), failedAttempts: 0, blockedUntil: '' });
        return { ok: true, success: true, nome: c.name, piso: c.floor, fracao: c.fraction };
      }
      return { ok: false, tipo: 'invalido', msg: 'Estado desconhecido.' };
    } catch (err) {
      return { ok: false, tipo: 'erro', msg: 'Erro interno.' };
    }
  }

  function resendPin(email) {
    try {
      const target = String(email || '').toLowerCase().trim();
      const c = repo.list().find(c => String(c.email).toLowerCase().trim() === target);
      if (!c) return { ok: false, success: false, msg: 'Email não encontrado.' };
      if (String(c.status || '').toUpperCase() !== 'APROVADO' || !c.pin) return { ok: false, success: false, msg: 'Não existe conta aprovada com este email.' };
      notifications.emailPINRecuperacao(email, c.name, c.pin);
      return { ok: true, success: true, msg: 'PIN enviado para o email registado.' };
    } catch (err) {
      return { ok: false, success: false, msg: 'Erro ao reenviar PIN: ' + err.message };
    }
  }

  function loginAdmin(email, pin) {
    return String(email || '').toLowerCase() === String(setting('ADMIN_EMAIL') || '').toLowerCase()
      && String(pin) === String(setting('ADMIN_PIN'))
      ? { ok: true, success: true } : { ok: false, msg: 'Credenciais inválidas.' };
  }

  function pending() {
    return repo.list().filter(c => c.status === 'PENDENTE').map(c => ({
      row: c.legacyRow, id: c.id, nome: c.name, piso: c.floor, fracao: c.fraction,
      email: c.email, telemovel: c.phone, dataRegisto: c.registeredAt ? dateText(c.registeredAt) : ''
    }));
  }

  function approved() {
    return repo.list().filter(c => c.status === 'APROVADO').map(c => ({
      row: c.legacyRow, id: c.id, nome: c.name, piso: c.floor, fracao: c.fraction,
      email: c.email, pin: c.pin, ultimoAcesso: c.lastAccess ? dateText(c.lastAccess) : 'Nunca'
    }));
  }

  function approve(row, adminEmail) {
    const c = repo.get(row);
    const pin = uniquePin();
    repo.update(row, { status: 'APROVADO', pin: pin, pinActive: true, approvedAt: ports.clock.now(), approvedBy: adminEmail || 'Admin' });
    notifications.emailAprovacao(c.email, c.name, pin);
    return { ok: true, success: true };
  }

  function reject(row) {
    const c = repo.get(row);
    repo.update(row, { status: 'REJEITADO' });
    notifications.emailRejeicao(c.email, c.name);
    return { ok: true, success: true };
  }

  function changeStatus(row, status) {
    repo.get(row);
    repo.update(row, { status: status });
    return { ok: true, success: true };
  }

  function regeneratePin(row) {
    const c = repo.get(row);
    const pin = uniquePin();
    repo.update(row, { pin: pin });
    notifications.emailPINRecuperacao(c.email, c.name, pin);
    return { ok: true, success: true };
  }

  return {
    register: register, login: login, loginAdmin: loginAdmin, resendPin: resendPin,
    pending: pending, approved: approved, approve: approve, reject: reject,
    block: row => changeStatus(row, 'BLOQUEADO'), unblock: row => changeStatus(row, 'APROVADO'),
    regeneratePin: regeneratePin,
    // Legado: esta ação nunca contou tentativas. A implementação segura exige alteração coordenada do login.
    failedAttempt: () => ({ ok: true, success: true })
  };
}
