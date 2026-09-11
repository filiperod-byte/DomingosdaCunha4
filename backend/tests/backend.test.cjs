const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { harness, bundled, FIXED } = require('./harness.cjs');
const { runScenario } = require('./scenarios.cjs');
const expected = require('./compatibility.json');

for (const [name, result] of Object.entries(expected)) {
  test(`compatibilidade com backend original: ${name}`, () => assert.deepEqual(runScenario(bundled(), name), result));
}

test('consultas não escrevem nem voltam a abrir a mesma spreadsheet', () => {
  const b = harness(bundled()); b.initialize();
  for (const action of ['status', 'health', 'garage.dashboard', 'garage.publicConfig', 'garage.approved', 'garage.history']) {
    const before = { ...b.counts }; b.get(action);
    assert.equal(b.counts.writes, before.writes, action);
    assert.equal(b.counts.opens - before.opens, 1, action);
  }
});

test('consultas numa instalação vazia não criam folhas', () => {
  const b = harness(bundled());
  assert.equal(b.get('health').success, false);
  assert.equal(b.counts.writes, 0);
  assert.equal(b.sheets.size, 0);
});

test('instalação repetida preserva dados existentes', () => {
  const b = harness(bundled()); b.initialize();
  b.post('garage.register', { nome: 'Teste', email: 'teste@example.invalid', piso: '9', fracao: 'B' });
  const before = b.snapshot(); b.initialize();
  assert.deepEqual(b.snapshot(), before);
});

test('definir PIN preserva outras propriedades e liberta o lock em erro', () => {
  const b = harness(bundled()); b.initialize(); b.props.UNRELATED = 'preservar';
  assert.equal(b.post('setPin', { pin: '123456' }).success, true);
  assert.equal(b.props.UNRELATED, 'preservar');
  assert.equal(b.post('setPin', { pin: '654321' }).success, false);
  assert.equal(b.counts.locks, b.counts.releases);
});

test('operações sobre linhas inválidas não alteram cabeçalhos ou dados', () => {
  const b = harness(bundled()); b.initialize(); const before = b.snapshot();
  for (const action of ['garage.approve', 'garage.reject', 'garage.block', 'garage.unblock', 'garage.regeneratePin']) {
    for (const row of [1, 0, -1, 2, 1.5]) assert.equal(b.post(action, { row }).success, false);
  }
  assert.deepEqual(b.snapshot(), before);
});

test('configuração é atualizada entre pedidos', () => {
  const b = harness(bundled()); b.initialize(); b.get('garage.publicConfig');
  b.post('garage.saveConfig', { configs: { NOME_CONDOMINIO: 'Nome novo' } });
  assert.equal(b.get('garage.publicConfig').nomeCondominio, 'Nome novo');
});

test('moradores funcionam com colunas reordenadas', () => {
  const b = harness(bundled()); b.initialize();
  const sheet = [...b.sheets.values()].find(s => s.rows[0].includes('PINAtivo'));
  sheet.rows.forEach(row => row.reverse());
  assert.equal(b.post('garage.register', { nome: 'Teste', email: 'teste@example.invalid', piso: '9', fracao: 'B' }).success, true);
  assert.equal(b.post('garage.approve', { row: 2 }).success, true);
  assert.equal(b.post('garage.loginPin', { pin: '478000' }).nome, 'Teste');
});

test('aplicação executa um ciclo de morador sem serviços Google', () => {
  const context = vm.createContext({});
  for (const name of ['config', 'domain', 'notifications', 'occurrences', 'legacy-pin', 'residents', 'accesses', 'application']) {
    const source = fs.readFileSync(path.join(__dirname, '../src', name + '.js'), 'utf8');
    assert.doesNotMatch(source, /\b(?:SpreadsheetApp|DriveApp|MailApp|PropertiesService|LockService|Session|Utilities|ContentService)\b/);
    vm.runInContext(source, context);
  }
  const rows = [], mails = [];
  const ports = {
    residents: {
      list: () => rows, get: row => rows.find(r => r.legacyRow === row),
      add: record => rows.push({ ...record, legacyRow: rows.length + 2 }),
      update: (row, patch) => Object.assign(rows.find(r => r.legacyRow === row), patch)
    },
    settings: { get: key => ({ ADMIN_EMAIL: 'admin@example.invalid', ADMIN_PIN: '999999' }[key] || '') },
    clock: { now: () => new Date(FIXED), format: date => date.toISOString() },
    ids: { entity: () => 'resident-1', pinCandidate: () => '123456' },
    log: { add() {} }, mail: { send: (...args) => mails.push(args) }
  };
  context.ports = ports;
  const app = vm.runInContext('createCondominiumApplication_(ports, CONFIG, createCondominiumDomain_(GARAGE_RESIDENTIAL_STRUCTURE, OPEN_STATUSES, PENDING_STATUSES))', context);
  assert.equal(app.dispatch('POST', 'garage.register', { nome: 'Teste', email: 'teste@example.invalid', piso: '9', fracao: 'B' }).success, true);
  assert.equal(app.dispatch('POST', 'garage.approve', { row: 2 }).success, true);
  assert.equal(app.dispatch('POST', 'garage.loginPin', { pin: '123456' }).nome, 'Teste');
  assert.equal(rows[0].status, 'APROVADO');
  assert.ok(mails.length > 0);
});
