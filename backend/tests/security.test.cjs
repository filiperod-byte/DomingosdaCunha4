const test = require('node:test');
const assert = require('node:assert/strict');
const { harness, bundled } = require('./harness.cjs');
function setup() {
 const b=harness(bundled()); b.initialize();
 const config=b.sheets.get('CONFIG');
 config.rows.find(r=>r[0]==='ADMIN_PIN')[1]='98765432';
 config.rows.find(r=>r[0]==='ADMIN_EMAIL')[1]='admin@example.invalid';
 return b;
}
function admin(b) { const r=b.post('garage.loginAdmin',{email:'admin@example.invalid',pin:'98765432'}); assert.ok(r.token,JSON.stringify(r));return r.token; }
function resident(b, token) {
 assert.equal(b.post('garage.register',{nome:'Pessoa real',email:'teste@example.invalid',piso:'9',fracao:'B'}).success,true);
 assert.equal(b.post('garage.approve',{row:2,token}).success,true);
 const r=b.post('garage.loginPin',{pin:'478000'});assert.ok(r.token,JSON.stringify(r));return r.token;
}
test('todas as rotas administrativas e aliases recusam chamadas anónimas',()=>{
 const b=setup();
 for(const action of ['health','openOccurrences','pendingOccurrences','garage.dashboard','garage.pending','garage.approved','garage.history','garage.adminConfig','garageDashboard','garageApproved','garageHistory','garageAdminConfig']) assert.equal(b.get(action).code,'AUTH_REQUIRED',action);
 for(const action of ['approveOccurrence','rejectOccurrence','closeOccurrence','garage.approve','garage.reject','garage.block','garage.unblock','garage.regeneratePin','garage.changeCode','garage.saveConfig','garageSaveConfig','report','garage.getCode']) assert.equal(b.post(action,{row:2,pin:'478000'}).code,'AUTH_REQUIRED',action);
 for(const action of ['setPin','resetPin','validatePin']) assert.equal(b.post(action,{pin:'123456'}).code,'LEGACY_LOGIN_DISABLED');
});
test('sessão permite administração, é revogada no logout e não aceita falsificação',()=>{
 const b=setup(),token=admin(b);
 assert.equal(b.post('garage.dashboard',{token,_method:'GET'}).ok,true);
 assert.equal(b.post('garage.dashboard',{token:token+'x',_method:'GET'}).code,'AUTH_REQUIRED');
 assert.equal(b.post('auth.logout',{token}).success,true);
 assert.equal(b.post('garage.dashboard',{token,_method:'GET'}).code,'AUTH_REQUIRED');
});
test('morador não administra, consulta código e autoria do reporte é fixada no servidor',()=>{
 const b=setup(),at=admin(b),token=resident(b,at);
 assert.equal(b.post('garage.saveConfig',{token,configs:{ADMIN_PIN:'999999'}}).code,'FORBIDDEN');
 assert.equal(b.post('garage.getCode',{token,pin:'falso',motivo:'Teste'}).success,true);
 assert.equal(b.post('report',{token,floor:9,point:'E1',name:'Autor falso',reason:'Teste'}).success,true);
 const pending=b.post('pendingOccurrences',{token:at,_method:'GET'});
 assert.equal(pending.occurrences[0].reportedBy,'Pessoa real');
 b.post('garage.block',{token:at,row:2});
 assert.equal(b.post('garage.getCode',{token}).code,'AUTH_REQUIRED');
 assert.equal(b.post('garage.loginPin',{pin:'478000'}).code,'INVALID_CREDENTIALS');
});
test('sessões expiram e alterar credenciais invalida a sessão administrativa',()=>{
 const b=setup();let token=admin(b);
 const key=Object.keys(b.props).find(k=>k.startsWith('DC4_SESSION_'));
 const session=JSON.parse(b.props[key]);session.expiresAt=0;b.props[key]=JSON.stringify(session);
 assert.equal(b.post('auth.session',{token}).code,'AUTH_REQUIRED');
 token=admin(b);
 assert.equal(b.post('garage.saveConfig',{token,configs:{ADMIN_PIN:'87654321'}}).success,true);
 assert.equal(b.post('auth.session',{token}).code,'AUTH_REQUIRED');
});
test('limites de login, configuração restrita e recuperação sem enumeração',()=>{
 const b=setup(),token=admin(b);
 assert.equal(b.post('garage.saveConfig',{token,configs:{ARBITRARY:'bad'}}).code,'INVALID_CONFIG');
 for(let i=0;i<9;i++) assert.equal(b.post('garage.loginAdmin',{email:'admin@example.invalid',pin:'111111'}).code,'INVALID_CREDENTIALS');
 assert.equal(b.post('garage.loginAdmin',{email:'admin@example.invalid',pin:'111111'}).code,'RATE_LIMITED');
 assert.equal(b.post('garage.resendPin',{email:'unknown@example.invalid'}).success,true);
});

test('registo grava fórmulas como texto e a atualização mantém a proteção',()=>{
 const b=setup(),token=admin(b);
 b.post('garage.register',{nome:'=1+1',email:'teste@example.invalid',piso:'9',fracao:'B'});
 const rows=b.sheets.get('CONDOMINOS').rows;
 const col=rows[0].indexOf('Nome');
 assert.equal(rows[1][col],"'=1+1");
 // getValues numa folha real devolve o texto sem o prefixo literal.
 rows[1][col]='=1+1';
 b.post('garage.approve',{row:2,token});
 assert.equal(rows[1][col],"'=1+1");
});

test('histórico omite credenciais novas e antigas, e a limpeza pode repetir-se',()=>{
 const b=setup(),at=admin(b),token=resident(b,at);
 const code=b.post('garage.getCode',{token,motivo:'Disjuntor'}).codigo;
 const rows=b.sheets.get('CONSULTAS').rows;
 const pinCol=rows[0].indexOf('PINUsado'),codeCol=rows[0].indexOf('CodigoMostrado');
 assert.equal(rows[1][pinCol],'');assert.equal(rows[1][codeCol],'');
 const clean=rows[1].slice();
 rows[1][pinCol]='478000';rows[1][codeCol]=code;
 const history=b.post('garage.history',{token:at,_method:'GET'});
 assert.equal(history[0].motivo,'Disjuntor');
 assert.equal('codigo' in history[0] || 'pin' in history[0],false);
 assert.equal('pin' in b.post('garage.approved',{token:at,_method:'GET'})[0],false);
 b.context.migrateSecurity372();b.context.migrateSecurity372();
 assert.deepEqual(rows[1],clean);
 assert.equal(b.post('garage.getCode',{token,motivo:'Teste'}).codigo,code);
});

test('uma nota nova não substitui o texto revisto pela administração',()=>{
 const b=setup(),at=admin(b),token=resident(b,at);
 const report=notes=>b.post('report',{token,floor:9,point:'E1',reason:'Outro',notes});
 const id=report('Texto inicial').occurrenceId;
 const approve=publicNotes=>b.post('approveOccurrence',{token:at,occurrenceId:id,publicNotes});
 const details=()=>b.context.doGet({parameter:{action:'status',details:'public'}}).reported;
 report('Texto revisto');
 assert.equal(approve('Texto inicial').success,false);
 assert.equal(details().length,0);
 assert.equal(approve('Texto revisto').success,true);
 report('Contacto privado');
 assert.equal(details()[0].description,'Texto revisto');
 assert.equal(b.post('openOccurrences',{token:at,_method:'GET'}).occurrences[0].notes,'Contacto privado');
});
