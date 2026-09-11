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
