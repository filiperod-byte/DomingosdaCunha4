const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm');
function setup(session=true){
 const elements=new Map(),screens=[],events={};
 const el=id=>{if(!elements.has(id))elements.set(id,{value:'',disabled:false,innerHTML:'',textContent:'',style:{},classList:{add(){},remove(){}},setAttribute(){}});return elements.get(id)};
 const c=vm.createContext({URL,URLSearchParams,console,window:{dc4HasSession:()=>session},document:{getElementById:el,querySelectorAll:()=>[]},location:{search:'',href:'https://example.invalid/app/V2/index.html'},sessionStorage:{getItem:()=>null,setItem(){},removeItem(){}},localStorage:{getItem:()=>null,setItem(){}},addEventListener:(name,fn)=>events[name]=fn,scrollTo(){},navigator:{userAgent:'test'},setTimeout,clearInterval});
 vm.runInContext(fs.readFileSync('V2/app.js','utf8'),c);c.irPara=id=>screens.push(id);c.renderStructure=()=>{};c.updateEye=()=>{};
 return {c,el,screens,events};
}
test('sessão local válida abre menu sem esperar pela configuração remota',()=>{
 const {c,screens,events}=setup();c.ensureConfig=()=>new Promise(()=>{});
 events.DOMContentLoaded();assert.deepEqual(screens,['screen-home']);
});
test('login em curso não é submetido duas vezes',async()=>{
 const {c}=setup();let finish,count=0;c.apiPost=()=>{count++;return new Promise(r=>finish=r)};
 vm.runInContext("App.pin='123456'",c);const first=c.submeterLogin();await c.submeterLogin();
 assert.equal(count,1);finish({success:true,nome:'Teste'});await first;
});
test('sessão expirada no cadeado preserva motivo e regressa à tarefa depois do login',async()=>{
 const {c,el,screens}=setup(false);
 vm.runInContext("App.motivo='Outro'",c);el('motivo-outro').value='Disjuntor';
 await c.pedirCodigo();assert.equal(screens.at(-1),'screen-login');
 c.apiPost=async()=>({success:true,nome:'Teste'});vm.runInContext("App.pin='123456'",c);
 await c.submeterLogin();assert.equal(screens.at(-1),'screen-cadeado-motivo');
 assert.equal(el('motivo-outro').value,'Disjuntor');assert.equal(vm.runInContext('App.pin',c),'');
});
