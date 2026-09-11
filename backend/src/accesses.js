// Consulta e gestão de códigos de recursos comuns. Independente da tecnologia de armazenamento.
function createAccessService_(ports) {
  const setting = key => ports.settings.get(key);

  function getCode(pin, reason, reasonOther, userAgent) {
    const c = ports.residents.list().find(c => String(c.pin) === String(pin));
    if (!c || c.status !== 'APROVADO') return { ok: false, msg: 'Acesso inválido.' };
    const code = setting('CODIGO_CADEADO');
    ports.consultations.add({
      id: ports.ids.entity('CONS'), at: ports.clock.now(), pin: pin,
      name: c.name, floor: c.floor, fraction: c.fraction, email: c.email,
      reason: reason, reasonOther: reasonOther || '', code: code, result: 'SUCESSO', userAgent: userAgent || ''
    });
    return { ok: true, success: true, codigo: code, aviso: setting('TEXTO_AVISO'), tempo: parseInt(setting('TEMPO_VISIVEL_SEGUNDOS')) || 15 };
  }

  function dashboard() {
    const residents = ports.residents.list();
    const today = ports.clock.now();
    const todayCount = ports.consultations.list().filter(record => {
      const date = new Date(record.at);
      return date.getDate() === today.getDate() && date.getMonth() === today.getMonth() && date.getFullYear() === today.getFullYear();
    }).length;
    return {
      ok: true, pendentes: residents.filter(c => c.status === 'PENDENTE').length,
      ativos: residents.filter(c => c.status === 'APROVADO').length,
      consultasHoje: todayCount, codigoAtual: setting('CODIGO_CADEADO')
    };
  }

  function history(limit) {
    return ports.consultations.list().slice().reverse().slice(0, limit || 50).map(record => ({
      dataHora: record.at ? ports.clock.format(record.at, 'dd/MM/yyyy HH:mm') : '',
      pin: record.pin, nome: record.name, piso: record.floor, fracao: record.fraction,
      email: record.email, motivo: record.reason, motivoOutro: record.reasonOther,
      codigo: record.code, resultado: record.result
    }));
  }

  function changeCode(code) {
    if (!/^\d+$/.test(String(code)) || String(code).length < 4) return { ok: false, msg: 'Código inválido.' };
    ports.settings.set('CODIGO_CADEADO', code);
    return { ok: true, success: true };
  }

  return { getCode: getCode, dashboard: dashboard, history: history, changeCode: changeCode };
}
