const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const vm=require('node:vm');
const root=path.join(__dirname,'../..');
function setup(hasSession=false){
 const elements=new Map();
 function el(id){if(!elements.has(id))elements.set(id,{value:'',textContent:'',innerHTML:'',disabled:false,open:false,options:[],classList:{add(){},remove(){}},appendChild(child){(this.children ||= []).push(child)},replaceChildren(){this.children=[]},focus(){},showModal(){this.open=true},close(){this.open=false},reset(){el('reason').value='';el('notes').value=''}});return elements.get(id)}
 const c=vm.createContext({URL,URLSearchParams,Date,console,Response,navigator:{userAgent:'test'},location:{search:'?floor=9&point=E1'},document:{getElementById:el,addEventListener(){},createElement:()=>({style:{}})},window:{dc4HasSession:()=>hasSession},sessionStorage:{setItem(){}},setTimeout:()=>0,clearTimeout(){}});
 vm.runInContext(fs.readFileSync(path.join(root,'occurrence-view.js'),'utf8'),c);
 vm.runInContext(fs.readFileSync(path.join(root,'qr-report.js'),'utf8'),c);
 vm.runInContext("CFG={features:{}};POINT={floor:9,point:'E1',location:'Patamar'};",c);
 el('reason').value='Selo danificado';el('notes').value='Texto a preservar';
 return {c,el};
}
const ev={preventDefault(){}};
test('consulta QR usa apenas estado público e mostra ocorrência existente',async()=>{
 const {c,el}=setup();const actions=[];c.apiGet=async a=>{actions.push(a);return {success:true,reported:[{floor:9,point:'E1'}]}};
 await c.loadExisting();assert.deepEqual(actions,['status']);assert.equal(el('statusBadge').textContent,'Ocorrência aberta');
});
test('falha de consulta nunca aparece como ausência de ocorrência',async()=>{
 const {c,el}=setup();c.apiGet=async()=>{throw Error('offline')};await c.loadExisting();assert.equal(el('statusBadge').textContent,'Estado por confirmar');
});
test('submeter sem sessão abre login de morador sem enviar nem apagar formulário',async()=>{
 const {c,el}=setup();c.apiPost=()=>{throw Error('Não devia enviar')};await c.submit(ev);
 assert.equal(el('authDialog').open,true);assert.equal(el('notes').value,'Texto a preservar');assert.equal(el('pin').value,'');
});
test('sessão existente envia diretamente sem novo login',async()=>{
 const {c,el}=setup(true);const actions=[];c.apiPost=async(a,p)=>{actions.push(a);assert.equal(p.point,'E1');return {success:true}};
 await c.submit(ev);assert.deepEqual(actions,['report']);assert.equal(el('authDialog').open,false);assert.equal(el('notes').value,'');
});
test('login bem sucedido envia reporte e PIN inválido mantém o texto',async()=>{
 const {c,el}=setup();await c.submit(ev);el('pin').value='123456';const actions=[];
 c.apiPost=async a=>{actions.push(a);return a==='garage.loginPin'?{success:true,token:'test',nome:'Morador'}:{success:true}};
 await c.loginAndSend(ev);assert.deepEqual(actions,['garage.loginPin','report']);
 const bad=setup();await bad.c.submit(ev);bad.el('pin').value='123456';bad.c.apiPost=async()=>({ok:false,msg:'PIN inválido'});
 await bad.c.loginAndSend(ev);assert.equal(bad.el('notes').value,'Texto a preservar');assert.equal(bad.el('authDialog').open,true);
});
test('registo pede aprovação sem enviar reporte e preserva fotografia',async()=>{
 const {c,el}=setup();const actions=[];vm.runInContext("FILE={name:'foto.jpg'}",c);
 c.apiPost=async a=>{actions.push(a);return {success:true}};await c.register(ev);
 assert.deepEqual(actions,['garage.register']);assert.match(el('registerMsg').textContent,/ainda não foi enviado/);assert.equal(el('notes').value,'Texto a preservar');assert.equal(vm.runInContext('FILE.name',c),'foto.jpg');
});
test('sessão expirada reabre login e mantém fotografia e texto',async()=>{
 const {c,el}=setup(true);vm.runInContext("FILE={name:'foto.jpg'}",c);c.fileToPayload=async()=>({base64:'fake'});
 c.apiPost=async()=>{const e=Error('Sessão expirada');e.code='AUTH_REQUIRED';throw e};await c.submit(ev);
 assert.equal(el('authDialog').open,true);assert.equal(el('notes').value,'Texto a preservar');assert.equal(vm.runInContext('FILE.name',c),'foto.jpg');
});
test('cancelar durante login não envia reporte',async()=>{
 const {c,el}=setup();await c.submit(ev);el('pin').value='123456';let complete;const actions=[];
 c.apiPost=a=>{actions.push(a);return new Promise(r=>complete=r)};
 const pending=c.loginAndSend(ev);el('authDialog').close();complete({success:true,token:'test'});await pending;
 assert.deepEqual(actions,['garage.loginPin']);assert.equal(el('notes').value,'Texto a preservar');
});
test('resposta HTML nunca é tratada como reporte enviado',async()=>{
 const {c}=setup();await assert.rejects(c.parseJson(new Response('<!DOCTYPE html>Erro')),/não foi confirmado/);
});
test('QR antigo do piso zero é encaminhado para reporte e não para administração',()=>{
 const routes=[];vm.runInNewContext(fs.readFileSync(path.join(root,'qr-entry.js'),'utf8'),{URL,URLSearchParams,document:{currentScript:{src:'https://example.invalid/app/qr-entry.js'}},location:{search:'?floor=0&point=CF1',replace:u=>routes.push(u)}});
 const url=new URL(routes[0]);assert.equal(url.pathname,'/app/qrcode-report.html');assert.equal(url.searchParams.get('floor'),'0');assert.equal(url.searchParams.get('point'),'CF1');
});

test('QR apresenta o motivo público e recomenda não repetir a mesma situação',async()=>{
 const {c,el}=setup();c.apiGet=async()=>({success:true,reported:[{floor:9,point:'E1',reason:'Extintor em falta',reportedAt:'2026-09-14T10:00:00.000Z'}]});
 await c.loadExisting();const text=el('existingBox').children.map(e=>e.textContent).join(' ');
 assert.match(text,/Motivo: Extintor em falta/);assert.match(text,/não precisa de a reportar novamente/);assert.match(text,/Registada em:/);
});

test('QR mostra descrição de Outro como texto, preservando linhas sem interpretar HTML',()=>{
 const {c,el}=setup();const description='Manómetro sem pressão.\n<img src=x onerror=alert(1)>';
 c.renderExisting({reason:'Outro',description});
 const children=el('existingBox').children;
 assert.ok(children.some(e=>e.textContent==='Descrição: '+description));
 assert.ok(children.every(e=>!e.innerHTML));
 assert.doesNotMatch(children.map(e=>e.textContent).join(' '),/descrição completa é consultada/);
 c.renderExisting({reason:'Outro',description:''});
 assert.match(el('existingBox').children.map(e=>e.textContent).join(' '),/não tem uma descrição registada/);
});

test('G4 existe apenas no piso -2 e o QR reconhece o novo ponto',()=>{
 const {c}=setup();const config=JSON.parse(fs.readFileSync(path.join(root,'config.json'),'utf8'));
 assert.deepEqual(config.building.floors.find(f=>f.floor===-2).extinguishers.map(e=>e.point),['G1','G2','G3','G4']);
 for(const floor of [-1,-3])assert.equal(config.building.floors.find(f=>f.floor===floor).extinguishers.length,3);
 c.NEW_CONFIG=config;vm.runInContext('CFG=NEW_CONFIG',c);c.location.search='?floor=-2&point=G4';
 assert.equal(c.findPointFromUrl().point,'G4');
});


test('mapa abre o mesmo formulário QR com ponto e origem explícitos',()=>{
 const {c}=setup(true);const routes=[];
 c.location.href='https://example.invalid/app/index.html';
 c.location.assign=url=>routes.push(url);
 const source=fs.readFileSync(path.join(root,'image-compress.js'),'utf8');
 const start=source.indexOf('openModal = function openModalOverride(ext)');
 const end=source.indexOf("  document.addEventListener('DOMContentLoaded'",start);
 vm.runInContext(source.slice(start,end),c);
 c.openModal({floor:-2,point:'G4'});
 const url=new URL(routes[0]);assert.equal(url.pathname,'/app/qrcode-report.html');
 assert.equal(url.searchParams.get('floor'),'-2');assert.equal(url.searchParams.get('point'),'G4');assert.equal(url.searchParams.get('from'),'map');
});
