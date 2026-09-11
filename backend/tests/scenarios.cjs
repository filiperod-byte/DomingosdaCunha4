const { harness } = require('./harness.cjs');

const resident = { nome: 'Morador de teste', piso: '9', fracao: 'B', email: 'morador@example.invalid' };
const report = { floor: 9, point: 'E1', name: 'Morador de teste', reason: 'Selo danificado', notes: 'Teste', source: 'app' };
const pin = '478000';

const scenarios = {
  'read-contracts': b => [
    ...['status', 'pinStatus', 'health', 'openOccurrences', 'pendingOccurrences',
      'garage.publicConfig', 'garage.structure', 'garage.dashboard', 'garage.pending', 'garage.approved', 'garage.history', 'garage.adminConfig'].map(b.get),
    b.context.doGet(), b.get('unknown'), b.get('toString'), b.post('unknown'), b.post('toString'),
    b.post('sendExtinguisherReport')
  ],
  'resident-lifecycle': b => {
    const out = [b.post('garage.register', resident), b.get('garage.pending')];
    out.push(b.post('garage.register', resident));
    out.push(b.post('garage.approve', { row: 2, adminEmail: 'admin@example.invalid' }));
    out.push(b.post('garage.loginPin', { pin }), b.get('garage.approved'));
    out.push(b.post('garage.getCode', { pin, motivo: 'Disjuntor', motivoOutro: 'Teste', userAgent: 'Test' }));
    out.push(b.get('garage.dashboard'), b.get('garage.history'));
    out.push(b.post('garage.block', { row: 2 }), b.post('garage.loginPin', { pin }), b.post('garage.getCode', { pin }));
    out.push(b.post('garage.unblock', { row: 2 }), b.post('garage.resendPin', { email: resident.email }));
    out.push(b.post('garage.regeneratePin', { row: 2 })); // Colisão sintética até esgotar as tentativas.
    out.push(b.post('garage.reject', { row: 2 }), b.post('garage.loginPin', { pin }));
    return out;
  },
  'validation-and-recovery': b => [
    b.post('garage.register', {}), b.post('garage.register', { nome: 'Teste' }),
    b.post('garage.register', { ...resident, piso: '999' }),
    b.post('garage.register', { ...resident, fracao: 'ZZ' }),
    b.post('garage.loginPin', { pin: '000000' }), b.post('garage.loginAdmin', { email: 'bad', pin: 'bad' }),
    b.post('garage.loginAdmin', { email: 'admin@example.invalid', pin: '123456' }),
    b.post('garage.resendPin', { email: 'unknown@example.invalid' }),
    b.post('garage.failedAttempt', { pin: '000000' }),
    b.post('garage.changeCode', { codigo: 'abc' }), b.post('garage.changeCode', { codigo: '12' }),
    b.post('garage.register', resident), b.post('garage.resendPin', { email: resident.email })
  ],
  'settings-and-form-body': b => [
    b.post('garage.changeCode', { codigo: '0098' }), b.get('garage.adminConfig'),
    b.post('garage.saveConfig', { configs: { NOME_CONDOMINIO: 'Novo nome', TEMPO_VISIVEL_SEGUNDOS: '30', TEXTO_AVISO: '' } }),
    b.get('garage.publicConfig'),
    b.context.doPost({ parameter: { action: 'garage.changeCode', code: '7654' } }),
    b.context.doPost({ postData: { contents: '{invalid' }, parameter: { action: 'garage.changeCode', novoCodigo: '4567' } }),
    b.get('garage.dashboard')
  ],
  'legacy-pin': b => [
    b.post('validatePin', { pin: '6543' }), b.post('setPin', { pin: '1' }),
    b.post('setPin', { pin: 'abcd' }), b.post('setPin', { pin: '6543' }),
    b.get('pinStatus'), b.post('setPin', { pin: '9876' }),
    b.post('validatePin', { pin: '' }), b.post('validatePin', { pin: '0000' }),
    b.post('validatePin', { pin: '6543' }), b.post('resetPin'), b.get('pinStatus')
  ],
  'occurrence-lifecycle': b => {
    const first = b.post('report', { ...report, photoDataUrl: 'data:image/png;base64,aGVsbG8=' });
    return [first, b.get('pendingOccurrences'), b.get('status'),
      b.post('report', { ...report, name: 'Segundo morador' }),
      b.post('approveOccurrence', { occurrenceId: first.occurrenceId, notes: 'Verificado' }),
      b.get('openOccurrences'), b.get('status'), b.post('report', { ...report, notes: 'Adicional' }),
      b.post('closeOccurrence', { occurrenceId: first.occurrenceId }),
      b.post('closeOccurrence', { occurrenceId: first.occurrenceId, closeNotes: 'Resolvido', closePhotoBase64: 'aGVsbG8=', closePhotoName: 'fecho.png', closePhotoType: 'image/png' }),
      b.get('status'), b.get('openOccurrences'), b.get('pendingOccurrences')];
  },
  'occurrence-validation-rejection': b => {
    const first = b.post('report', report);
    return [first, b.post('rejectOccurrence', { occurrenceId: first.occurrenceId }),
      b.post('approveOccurrence', { occurrenceId: first.occurrenceId }),
      b.post('report', {}), b.post('report', { floor: 9 }),
      b.post('report', { floor: 9, point: 'E1' }), b.post('report', { floor: 9, point: 'E1', name: 'Teste' }),
      b.post('closeOccurrence', { floor: 9, point: 'E1' }), b.get('pendingOccurrences')];
  },
  'aliases': b => [
    ...['garagePublicConfig', 'garageStructure', 'garageDashboard', 'garagePending', 'garageApproved', 'garageHistory', 'garageAdminConfig'].map(b.get),
    b.post('garageRegister', { name: resident.nome, floor: resident.piso, fraction: resident.fracao, email: resident.email }),
    b.post('garageApprove', { row: 2 }), b.post('garageLoginPin', { pin }),
    b.post('garageLoginAdmin', { email: 'admin@example.invalid', pin: '123456' }),
    b.post('garageFailedAttempt', { pin }), b.post('garageResendPin', { email: resident.email }),
    b.post('garageRecoverCode', { email: resident.email }), b.post('garage.recoverCode', { email: resident.email }),
    b.post('garageGetCode', { pin, reason: 'Outro', reasonOther: 'Teste', ua: 'Test' }),
    b.post('garageChangeCode', { code: '4321' }), b.post('garageSaveConfig', { configs: { TEXTO_AVISO: 'Novo aviso' } }),
    b.post('garageBlock', { row: 2 }), b.post('garageUnblock', { row: 2 }),
    b.post('garageRegeneratePin', { row: 2 }), b.post('garageReject', { row: 2 })
  ]
};

function runScenario(source, name) {
  const b = harness(source, { legacyCore: true });
  b.initialize();
  const responses = scenarios[name](b);
  return JSON.parse(JSON.stringify({ responses, state: b.snapshot() }));
}
module.exports = { scenarios, runScenario };
