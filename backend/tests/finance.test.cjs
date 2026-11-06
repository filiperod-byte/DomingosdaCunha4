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
  make('FATURAS',
    ['ID','Tipo','Fornecedor','NIF','Descrição','Documento','Emissão','Vencimento',
     'Período início','Período fim','Total (€)','Método','IBAN fornecedor','Entidade MB',
     'Referência MB','Validação','Pagamento','Data pag.','Validado por','PDF original',
     'Observações','ID Drive'],
    [['DOC-1','Documento por classificar','Por identificar','','PDF novo','PDF-1','','',
      '','',0,'Por confirmar','','','','Por validar','Por confirmar','','','','','DRIVE-1']]);
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
  tokenResident(b,admin);
  const mailsBefore=b.counts.mails; // Registo e aprovação também geram emails legítimos.
  assert.equal(b.post('finance.admin.notifyCycle',{token:admin,cycleId:'CYCLE-EDP-1'}).code,'FINANCE_ERROR');
  b.sheets.get('WALLBOX_CONFIG').rows.find(r=>r[0]==='NOTIFICACOES_ATIVAS')[1]='SIM';
  const a=b.post('finance.admin.notifyCycle',{token:admin,cycleId:'CYCLE-EDP-1'});
  assert.equal(a.ok,true);
  assert.equal(a.notified,1);
  const again=b.post('finance.admin.notifyCycle',{token:admin,cycleId:'CYCLE-EDP-1'});
  assert.equal(again.ok,true);
  assert.equal(again.notified,0);
  assert.equal(b.counts.mails-mailsBefore,1);
});
test('a política HTTP rejeita GET e _method GET para operações financeiras',()=>{
  const b=setup(),admin=tokenAdmin(b);
  assert.equal(b.post('finance.admin.bootstrap',{token:admin,_method:'GET'}).code,'FORBIDDEN');
  assert.equal(b.get('finance.admin.bootstrap').code,'AUTH_REQUIRED');
});
test('documento sem classificação e IBAN inválido não podem entrar no circuito de pagamento',()=>{
  const b=setup(),admin=tokenAdmin(b);
  const validate=()=>b.post('finance.admin.validateInvoice',{token:admin,id:'DOC-1'});
  assert.equal(validate().code,'FINANCE_ERROR');
  const request={token:admin,id:'DOC-1',documentType:'Fatura',supplier:'Fornecedor de teste',
    description:'Serviço de manutenção',value:75.72,due:'2026-11-15',method:'Por confirmar'};
  assert.equal(b.post('finance.admin.updateInvoice',{...request,iban:'PT50INVALID'}).code,'FINANCE_ERROR');
  assert.equal(b.post('finance.admin.updateInvoice',request).ok,true);
  assert.equal(validate().ok,true);
  assert.equal(b.post('finance.admin.markInvoicePaid',{token:admin,id:'DOC-1',paidAt:'2026-11-16'}).ok,true);
  assert.equal(b.post('finance.admin.updateInvoice',request).code,'FINANCE_ERROR');
});

test('diagnóstico apenas de leitura apresenta bloqueios sem alterar pagamentos ou configuração',()=>{
  const b=setup(),admin=tokenAdmin(b);
  const before=b.counts.writes;
  const r=b.post('finance.admin.preflight',{token:admin});
  assert.equal(r.ok,true);
  assert.equal(r.ready,false);
  assert.ok(r.checks.find(c=>c.id==='finance_sheet'&&c.ok));
  assert.ok(r.checks.find(c=>c.id==='residents'&&!c.ok));
  assert.ok(r.checks.find(c=>c.id==='edp_invoice'&&!c.ok));
  assert.equal(b.counts.writes,before);
});
test('ativação exige morador previamente aprovado na fração exata',()=>{
  const b=setup(),admin=tokenAdmin(b);
  const attempt=b.post('finance.admin.upsertFraction',{token:admin,
    fractionId:'8B',label:'8.º B',email:'unverified@example.invalid',
    initialReading:0,initialDate:'2026-11-06',initialPhotoBase64:'data:image/jpeg;base64,YWJj'});
  assert.equal(attempt.code,'FINANCE_ERROR');
  assert.equal(b.sheets.get('WALLBOX_FRACOES').rows.length,2);
});

test('um dashboard financeiro reutiliza a referência à spreadsheet e não escreve',()=>{
  const b=setup(),admin=tokenAdmin(b);
  const before={opens:b.counts.opens,writes:b.counts.writes};
  const r=b.post('finance.admin.bootstrap',{token:admin});
  assert.equal(r.ok,true);
  assert.equal(b.counts.writes,before.writes);
  assert.ok(b.counts.opens-before.opens<=3,'Muitos acessos repetidos à spreadsheet');
});

test('admin mostra uma entrada por fração mesmo após reativação, conservando histórico',()=>{
  const b=setup(),admin=tokenAdmin(b);
  b.sheets.get('WALLBOX_FRACOES').rows.push(
    ['9B','9.º B','Inativa','antigo@example.invalid','2026-01-05',0,'2026-01-05'],
    ['9B','9.º B','Ativa','teste@example.invalid','2026-09-15',0,'2026-09-15']);
  const r=b.post('finance.admin.bootstrap',{token:admin});
  assert.equal(r.ok,true);
  assert.equal(r.fractions.filter(f=>f.id==='9B').length,1);
  assert.equal(r.fractions.find(f=>f.id==='9B').status,'Ativa');
  assert.equal(b.sheets.get('WALLBOX_FRACOES').rows.length,4);
});

test('monitorização pode ser ativada e desativada pelo administrador sem duplicar triggers',()=>{
  const b=setup(),admin=tokenAdmin(b);
  assert.equal(b.post('finance.admin.enableScanner').code,'AUTH_REQUIRED');
  assert.equal(b.post('finance.admin.scannerStatus',{token:admin}).enabled,false);
  assert.equal(b.post('finance.admin.enableScanner',{token:admin}).installed,true);
  assert.equal(b.post('finance.admin.enableScanner',{token:admin}).alreadyInstalled,true);
  assert.equal(b.post('finance.admin.scannerStatus',{token:admin}).count,1);
  assert.equal(b.post('finance.admin.disableScanner',{token:admin}).removed,1);
  assert.equal(b.post('finance.admin.scannerStatus',{token:admin}).enabled,false);
});
test('a procura manual não inventa faturas quando não há ficheiros no Drive',()=>{
  const b=setup(),admin=tokenAdmin(b);
  b.props.FINANCE_FATURAS_ROOT_ID='TEST-DRIVE';
  const r=b.post('finance.admin.scanNow',{token:admin});
  assert.equal(r.ok,true);
  assert.equal(r.newDocuments,0);
  assert.equal(b.sheets.get('FATURAS').rows.length,2);
});
