const test=require('node:test');
const assert=require('node:assert/strict');
const {harness,bundled}=require('./harness.cjs');

function setup(){
  const b=harness(bundled());b.initialize();
  const config=b.sheets.get('CONFIG');
  config.rows.find(r=>r[0]==='ADMIN_PIN')[1]='98765432';
  config.rows.find(r=>r[0]==='ADMIN_EMAIL')[1]='admin@example.invalid';
  b.props.FINANCE_SPREADSHEET_ID='TEST-FINANCE';
  const Sheet=config.constructor;
  const make=(name,columns,rows=[])=>b.sheets.set(name,new Sheet([columns,...rows]));
  make('WALLBOX_CONFIG',['Chave','Valor','Tipo','Unidade','Início vigência','Atualizado em','Atualizado por','Observações'],[['DESCONTO_FIXOS',25],['NOTIFICACOES_ATIVAS','NAO']]);
  make('WALLBOX_FRACOES',
    ['ID fração','Fração','Estado','Email autorizado','Data adesão','Leitura inicial (kWh)','Data leitura inicial'],
    [['9B','9.º B','Ativa','teste@example.invalid','2026-09-15',0,'2026-09-15']]);
  make('WALLBOX_CICLOS',
    ['ID ciclo','Fatura ID','Início faturação','Fim faturação','Estado','Validado por','Criado em','Notificado em','Desconto aplicado %'],
    [['CYCLE-EDP-1','EDP-1','2026-04-29','2026-06-28','Aberto','admin@example.invalid','','',25]]);
  make('WALLBOX_TARIFAS_CICLO',
    ['ID ciclo','Início faturação','Fim faturação','Consumo prédio kWh','Energia sem IVA','IVA energia','IEC sem IVA','IVA IEC','Potência sem IVA','IVA potência','DGEG sem IVA','IVA DGEG','Audiovisual sem IVA','IVA audiovisual'],
    [['CYCLE-EDP-1','2026-04-29','2026-06-28',777,.1671,.23,.001,.23,.8485,.23,.35,.23,2.85,.06]]);
  make('LEITURAS_WALLBOX',
    ['ID leitura','Fração','Ciclo EDP','Data leitura','Leitura acumulada kWh','Foto Drive ID','Estado','Submetido por','Submetido em','Validado por','Validado em','Observações'],
    [['LR-1','9B','CYCLE-EDP-1','2026-10-28',176,'PHOTO-1','Submetida','teste@example.invalid','','','','']]);
  make('WALLBOX_PAGAMENTOS',
    ['ID apuramento','Fração','Ciclo EDP','Data inicial','Data final','Leitura inicial kWh','Leitura final kWh','Consumo kWh','Energia + IEC €','Custos fixos antes desc. €','Desconto %','Parcela custos fixos €','Total a pagar €','Estado','Emissão','Pago em','Validado por','Recebido por','Foto leitura Drive ID','Observações']);
  make('WALLBOX_DESCONTOS',['ID','Início vigência','Fim vigência','Desconto %','Motivo','Definido por','Registado em','Ciclo aplicado','Notas']);
  make('WALLBOX_NOTIFICACOES',['ID','Ciclo','Fração','Email','Estado','Enviado em','Erro','ID leitor','Observações']);
  make('HISTORICO',['Data','ID documento','Ação','Antes','Depois','Utilizador','Notas','Origem']);
  make('FATURAS',['ID','Tipo','Fornecedor','Descrição','Período início','Período fim','Validação','Pagamento','Validado por','PDF original']);
  make('TARIFAS_EDP',['Parâmetro','Valor']);
  return b;
}
function tokenAdmin(b){const x=b.post('garage.loginAdmin',{email:'admin@example.invalid',pin:'98765432'});assert.ok(x.token);return x.token;}
function tokenResident(b,admin){
  assert.equal(b.post('garage.register',{nome:'Morador real',email:'teste@example.invalid',piso:'9',fracao:'B'}).success,true);
  assert.equal(b.post('garage.approve',{token:admin,row:2}).success,true);
  const x=b.post('garage.loginPin',{pin:'478000'});assert.ok(x.token);
  return x.token;
}
test('financeiro nega sessão anónima, residente não administra e admin não usa área residente',()=>{
  const b=setup(),admin=tokenAdmin(b),resident=tokenResident(b,admin);
  assert.equal(b.post('finance.admin.bootstrap').code,'AUTH_REQUIRED');
  assert.equal(b.post('finance.admin.bootstrap',{token:resident,role:'admin'}).code,'FORBIDDEN');
  assert.equal(b.post('finance.resident.status',{token:admin,role:'resident'}).code,'FORBIDDEN');
  assert.equal(b.post('finance.admin.bootstrap',{token:admin}).ok,true);
});
test('falsificar a fração no corpo não dá acesso a outra fração',()=>{
  const b=setup(),admin=tokenAdmin(b),resident=tokenResident(b,admin);
  const answer=b.post('finance.resident.status',{token:resident,role:'admin',fractionId:'3C',email:'outro@example.invalid'});
  assert.equal(answer.ok,true);
  assert.equal(answer.fraction,'9.º B');
  assert.equal(answer.payments.length,0);
  assert.equal(answer.readings.length,1);
});
test('fórmula do simulador gera 48.38 EUR e validação idempotente não cria segundo pagamento',()=>{
  const b=setup(),admin=tokenAdmin(b);
  const answer=b.post('finance.admin.validateReading',{token:admin,id:'LR-1'});
  assert.equal(answer.ok,true);
  assert.equal(answer.amount,48.38);
  const rows=b.sheets.get('WALLBOX_PAGAMENTOS').rows;
  assert.equal(rows.length,2);
  assert.equal(rows[1][12],48.38);
  assert.equal(rows[1][10],25);
  assert.equal(b.post('finance.admin.validateReading',{token:admin,id:'LR-1'}).code,'FINANCE_ERROR');
  assert.equal(rows.length,2);
});
test('desconto posterior não altera pagamentos fechados',()=>{
  const b=setup(),admin=tokenAdmin(b);
  const approved=b.post('finance.admin.validateReading',{token:admin,id:'LR-1'});
  assert.equal(approved.amount,48.38);
  const discounted=b.post('finance.admin.saveDiscount',{token:admin,percent:12.5,effectiveFrom:'2026-11-01'});
  assert.equal(discounted.ok,true);
  assert.equal(b.sheets.get('WALLBOX_PAGAMENTOS').rows[1][12],48.38);
  assert.equal(b.sheets.get('WALLBOX_PAGAMENTOS').rows[1][10],25);
});
test('emails controlados por sessão, configuração e par ciclo/fração',()=>{
  const b=setup(),admin=tokenAdmin(b);
  assert.equal(b.post('finance.admin.notifyCycle',{token:admin,cycleId:'CYCLE-EDP-1'}).code,'FINANCE_ERROR');
  b.sheets.get('WALLBOX_CONFIG').rows.find(r=>r[0]==='NOTIFICACOES_ATIVAS')[1]='SIM';
  const a=b.post('finance.admin.notifyCycle',{token:admin,cycleId:'CYCLE-EDP-1'});
  assert.equal(a.ok,true);
  assert.equal(a.notified,1);
  const again=b.post('finance.admin.notifyCycle',{token:admin,cycleId:'CYCLE-EDP-1'});
  assert.equal(again.ok,true);
  assert.equal(again.notified,0);
  assert.equal(b.counts.mails,1);
});
test('a política HTTP rejeita GET e _method GET para operações financeiras',()=>{
  const b=setup(),admin=tokenAdmin(b);
  assert.equal(b.post('finance.admin.bootstrap',{token:admin,_method:'GET'}).code,'FORBIDDEN');
  assert.equal(b.get('finance.admin.bootstrap').code,'AUTH_REQUIRED');
});