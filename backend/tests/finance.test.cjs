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
    ['ID fração','Fração','Estado','Email autorizado','Data adesão','Leitura inicial (kWh)',
     'Data leitura inicial','Foto inicial Drive ID','Data desativação','Leitura final (kWh)',
     'Foto final Drive ID','Criado por','Atualizado em','Observações'],
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
     'Observações','ID Drive','IVA (€)','Base tributável (€)','Estado extração','Data extração'],
    [['DOC-1','Documento por classificar','Por identificar','','PDF novo','PDF-1','','',
      '','',0,'Por confirmar','','','','Por validar','Por confirmar','','','','','DRIVE-1',
      '', '', 'Extraído — por validar','']]);
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

test('desconto futuro só entra em vigor na data indicada',()=>{
  const b=setup(),admin=tokenAdmin(b);
  const saved=b.post('finance.admin.saveDiscount',{token:admin,percent:12.5,effectiveFrom:'2026-11-01'});
  assert.equal(saved.ok,true);
  assert.equal(b.context.fncDiscountAt_('2026-09-20'),25);
  assert.equal(b.context.fncDiscountAt_('2026-11-02'),12.5);
});

test('a abertura de Faturas não carrega todo o histórico Wallbox e vice-versa',()=>{
 const b=setup(),admin=tokenAdmin(b);
 const before=b.counts.reads;
 const invoices=b.post('finance.admin.invoices',{token:admin});
 const readsInvoices=b.counts.reads-before;
 assert.equal(invoices.ok,true);assert.equal(invoices.invoices.length,1);
 assert.equal(Object.hasOwn(invoices,'fractions'),false);
 const beforeWB=b.counts.reads;
 const wallboxes=b.post('finance.admin.wallboxes',{token:admin});
 assert.equal(wallboxes.ok,true);
 assert.equal(wallboxes.fractions.length,1);
 assert.equal(Object.hasOwn(wallboxes,'invoices'),false);
 assert.ok(readsInvoices<8);
 assert.ok(b.counts.reads-beforeWB<20);
});
test('corrigir leitura inicial permite valores de fração inativa sem alterar outros ciclos',()=>{
 const b=setup(),admin=tokenAdmin(b);
 b.sheets.get('WALLBOX_FRACOES').rows.push(['3C','3.º C','Inativa','outro@example.invalid','2026-11-01',1777,'2026-11-01']);
 const res=b.post('finance.admin.correctInitial',{token:admin,fractionId:'3C',
   initialReading:1765,initialDate:'2026-11-02',
   email:'emailcorrigido@example.invalid',reason:'Erro na leitura inicial'});
 assert.equal(res.ok,true);
 const row=b.sheets.get('WALLBOX_FRACOES').rows[2];
 assert.equal(row[5],1765);assert.equal(row[6],'2026-11-02');
 assert.equal(row[3],'emailcorrigido@example.invalid');
 assert.equal(b.sheets.get('HISTORICO').rows.length,2);
 assert.equal(b.sheets.get('WALLBOX_PAGAMENTOS').rows.length,1);
});
test('leitura inicial não pode ser alterada após leitura aprovada ou pagamento',()=>{
 const b=setup(),admin=tokenAdmin(b);
 assert.equal(b.post('finance.admin.validateReading',{token:admin,id:'LR-1'}).ok,true);
 const res=b.post('finance.admin.correctInitial',{token:admin,fractionId:'9B',
   initialReading:99,initialDate:'2026-09-15',reason:'Correção após apuramento'});
 assert.equal(res.code,'FINANCE_ERROR');
 assert.equal(b.sheets.get('WALLBOX_FRACOES').rows[1][5],0);
});
test('morador só consegue corrigir leitura submetida da própria fração antes da validação',()=>{
 const b=setup(),admin=tokenAdmin(b),resident=tokenResident(b,admin);
 const r=b.post('finance.resident.amendReading',{token:resident,id:'LR-1',
   kwh:180,readingDate:'2026-10-28',reason:'Valor digitado errado'});
 assert.equal(r.ok,true);
 assert.equal(b.sheets.get('LEITURAS_WALLBOX').rows[1][4],180);
 assert.equal(b.sheets.get('LEITURAS_WALLBOX').rows[1][6],'Submetida');
 assert.equal(b.post('finance.resident.amendReading',{token:resident,id:'LR-1',
   kwh:-1,readingDate:'2026-10-28'}).code,'FINANCE_ERROR');
 const validated=b.post('finance.admin.validateReading',{token:admin,id:'LR-1'});
 assert.equal(validated.ok,true);
 assert.equal(b.post('finance.resident.amendReading',{token:resident,id:'LR-1',
   kwh:190,readingDate:'2026-10-28'}).code,'FINANCE_ERROR');
});
test('morador sem wallbox não consegue acesso financeiro pela simples consulta de elegibilidade',()=>{
 const b=setup(),admin=tokenAdmin(b),resident=tokenResident(b,admin);
 assert.equal(b.post('finance.resident.eligibility',{token:resident}).eligible,true);
 assert.equal(b.post('finance.admin.wallboxes',{token:resident}).code,'FORBIDDEN');
 const row=b.sheets.get('WALLBOX_FRACOES').rows[1];row[2]='Inativa';
 assert.equal(b.post('finance.resident.eligibility',{token:resident}).eligible,false);
});

test('alterar email de wallbox ativa para morador não aprovado é recusado',()=>{
 const b=setup(),admin=tokenAdmin(b);
 const result=b.post('finance.admin.correctInitial',{token:admin,fractionId:'9B',
   initialReading:5,initialDate:'2026-09-15',email:'outra@example.invalid',
   reason:'Correção do destinatário'});
 assert.equal(result.code,'FINANCE_ERROR');
 assert.equal(b.sheets.get('WALLBOX_FRACOES').rows[1][3],'teste@example.invalid');
});

test('consulta dos triggers sem scope OAuth não bloqueia o módulo financeiro',()=>{
  const b=setup(),admin=tokenAdmin(b);
  b.context.ScriptApp.getProjectTriggers=()=>{throw new Error('Authorization is required for script.scriptapp')};
  const status=b.post('finance.admin.scannerStatus',{token:admin});
  assert.equal(status.ok,true);
  assert.equal(status.enabled,false);
  assert.equal(status.verified,false);
  assert.equal(status.authorizationRequired,true);
  assert.match(status.message,/autorização/i);
  const invoices=b.post('finance.admin.invoices',{token:admin});
  assert.equal(invoices.ok,true);
  assert.equal(invoices.invoices.length,1);
});
test('tentativa de instalar trigger sem autorização devolve instrução segura',()=>{
  const b=setup(),admin=tokenAdmin(b);
  b.context.ScriptApp.getProjectTriggers=()=>{throw new Error('You do not have permission to call getProjectTriggers')};
  const res=b.post('finance.admin.enableScanner',{token:admin});
  assert.equal(res.ok,false);
  assert.equal(res.requiresAuthorization,true);
  assert.match(res.message,/ativarSincronizacaoFaturas/);
  assert.equal(b.props.FINANCE_SCANNER_ENABLED,undefined);
});
test('ativar e desativar regista estado apenas depois da operação real',()=>{
  const b=setup(),admin=tokenAdmin(b);
  assert.equal(b.post('finance.admin.enableScanner',{token:admin}).ok,true);
  assert.equal(b.props.FINANCE_SCANNER_ENABLED,'SIM');
  const state=b.post('finance.admin.scannerStatus',{token:admin});
  assert.equal(state.enabled,true);
  assert.equal(state.verified,true);
  assert.equal(b.post('finance.admin.disableScanner',{token:admin}).ok,true);
  assert.equal(b.props.FINANCE_SCANNER_ENABLED,'NAO');
});

test('função pública de configuração aparece no editor e instala apenas um trigger',()=>{
 const b=setup();
 const run=b.context.ativarSincronizacaoFaturas;
 assert.equal(typeof run,'function');
 assert.equal(run().ok,true);
 assert.equal(run().alreadyInstalled,true);
 const admin=tokenAdmin(b);
 assert.equal(b.post('finance.admin.scannerStatus',{token:admin}).count,1);
});
test('configuração no editor informa o erro quando o proprietário ainda não autorizou o ScriptApp',()=>{
 const b=setup();
 b.context.ScriptApp.getProjectTriggers=()=>{throw new Error('Authorization required for script.scriptapp')};
 assert.throws(()=>b.context.ativarSincronizacaoFaturas(),/autoriza/i);
});

test('ativação aceita função interna antiga sem underscore no Apps Script',()=>{
 const b=setup();
 const original=b.context.financeInstallScannerTrigger_;
 assert.equal(typeof original,'function');
 b.context.financeInstallScannerTrigger=original;
 b.context.financeInstallScannerTrigger_=undefined;
 assert.equal(b.context.ativarSincronizacaoFaturas().ok,true);
 assert.equal(b.context.ativarSincronizacaoFaturas().alreadyInstalled,true);
});
test('ativação descreve a falta de ambas as funções sem ReferenceError',()=>{
 const b=setup();
 b.context.financeInstallScannerTrigger_=undefined;
 b.context.financeInstallScannerTrigger=undefined;
 assert.throws(()=>b.context.ativarSincronizacaoFaturas(),/Instalador de faturas não encontrado/);
});

test('instalador solicita a permissão script.scriptapp antes de procurar ou criar triggers',()=>{
 const b=setup();
 const res=b.context.ativarSincronizacaoFaturas();
 assert.equal(res.ok,true);
 assert.equal(b.counts.authChecks,1);
 const admin=tokenAdmin(b);
 assert.equal(b.post('finance.admin.scannerStatus',{token:admin}).enabled,true);
});
test('falta de consentimento interrompe o instalador para exibir o pedido OAuth no editor',()=>{
 const b=harness(bundled(),{missingScannerConsent:true});b.initialize();
 assert.throws(()=>b.context.ativarSincronizacaoFaturas(),/Authorization required/);
 assert.equal(b.counts.authChecks,1);
 assert.equal(b.props.FINANCE_SCANNER_ENABLED,undefined);
});

test('extração manual preenche apenas os dados identificados e mantém por validar',()=>{
  const b=setup(),admin=tokenAdmin(b);
  const pdf=[
    'Factura nº1 2026/52','De','Empresa Fictícia Unipessoal. Lda.',
    'Nº Contribuinte: 500000000','Data de Emissão','2026-10-20',
    'Data de Vencimento','2026-10-24','Para','Condomínio',
    'Resumo','Subtotal da Factura 240,00 €',
    'IVA 23.00% (Incidência: 240,00) 55,20 €',
    'Total da Factura 295,20 €'
  ].join('\n');
  b.context.fncDriveOcrText_=()=>pdf;
  const r=b.post('finance.admin.extractInvoice',{token:admin,id:'DOC-1'});
  assert.equal(r.ok,true);
  assert.equal(r.amount,295.2);
  const headers=b.sheets.get('FATURAS').rows[0];
  const row=b.sheets.get('FATURAS').rows[1];
  assert.equal(row[headers.indexOf('Fornecedor')],'Empresa Fictícia Unipessoal. Lda.');
  assert.equal(row[headers.indexOf('Total (€)')],295.2);
  assert.equal(row[headers.indexOf('Validação')],'Por validar');
  assert.equal(row[headers.indexOf('Pagamento')],'Por confirmar');
});
test('extração é proibida a moradores e não altera documentos validados',()=>{
  const b=setup(),admin=tokenAdmin(b),resident=tokenResident(b,admin);
  assert.equal(b.post('finance.admin.extractInvoice',{token:resident,id:'DOC-1'}).code,'FORBIDDEN');
  b.sheets.get('FATURAS').rows[1][15]='Validada';
  assert.equal(b.post('finance.admin.extractInvoice',{token:admin,id:'DOC-1'}).code,'FINANCE_ERROR');
});

test('releitura repara nome do fornecedor com morada sem mexer nos campos manuais',()=>{
 const b=setup(),admin=tokenAdmin(b);
 const head=b.sheets.get('FATURAS').rows[0],row=b.sheets.get('FATURAS').rows[1];
 row[head.indexOf('Fornecedor')]='Empresa Fictícia Unipessoal. Lda. Impasse Exemplo nº1';
 row[head.indexOf('Descrição')]='fornecedor-original.pdf';
 row[head.indexOf('Estado extração')]='Extração parcial — por validar';
 const pdf=[
  'Factura nº1 2026/52','De','Empresa Fictícia Unipessoal. Lda. Impasse Exemplo nº1',
  'Nº Contribuinte: 500000000','Data de Emissão','2026-10-20',
  'Data de Vencimento','2026-10-24','Para','Condomínio',
  'Lista de Artigos','DESCRIÇÃO DO ARTIGO QUANT. PREÇO DESC. IVA (%) TOTAL',
  'Manutenção portão 1.00 240,00 0,00 23.00 240,00',
  'Subtotal da Factura 240,00 €',
  'IVA 23.00% (Incidência: 240,00) 55,20 €',
  'Total da Factura 295,20 €'
 ].join('\n');
 b.context.fncDriveOcrText_=()=>pdf;
 const result=b.post('finance.admin.extractInvoice',{token:admin,id:'DOC-1'});
 assert.equal(result.ok,true);
 assert.equal(row[head.indexOf('Fornecedor')],'Empresa Fictícia Unipessoal. Lda.');
 assert.match(row[head.indexOf('Descrição')],/Manutenção portão/);
 assert.match(row[head.indexOf('Observações')],/Leitor v2/);
 assert.equal(row[head.indexOf('Validação')],'Por validar');
 row[head.indexOf('Fornecedor')]='Fornecedor confirmado manualmente';
 const again=b.post('finance.admin.extractInvoice',{token:admin,id:'DOC-1'});
 assert.equal(again.ok,true);
 assert.equal(row[head.indexOf('Fornecedor')],'Fornecedor confirmado manualmente');
});
