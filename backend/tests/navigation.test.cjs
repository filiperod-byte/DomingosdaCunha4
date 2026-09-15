const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const vm=require('node:vm');
const source=fs.readFileSync(path.join(__dirname,'../../app-navigation.js'),'utf8');
function mount(page,initialScreen,hasSession=true,hash=''){
 let active=initialScreen;
 const elements=new Map(),routes=[],screens=[];let hiddenCodes=0,restores=0;
 function element(){return {children:[],setAttribute(){},appendChild(e){this.children.push(e);if(e.id)elements.set(e.id,e)},prepend(e){this.appendChild(e)},addEventListener(event,callback){this[event]=callback}}}
 const context={URL,URLSearchParams,document:{currentScript:{src:'https://example.invalid/app/app-navigation.js'},readyState:'complete',head:element(),body:element(),getElementById:id=>elements.get(id)||null,createElement:element,querySelector:()=>active?{id:active}:null,querySelectorAll:()=>[]},location:{pathname:'/app/'+page,search:'',hash,assign:url=>routes.push(url)},history:{replaceState(){}},window:{dc4HasSession:()=>hasSession,hideCode:()=>hiddenCodes++,restoreUser:()=>restores++,irPara:id=>{active=id;screens.push(id)}}};
 vm.runInNewContext(source,context);
 return {context,button:elements.get('dc4-back'),routes,screens,hiddenCodes:()=>hiddenCodes,restores:()=>restores};
}
test('ecrã de código volta ao motivo e oculta o código',()=>{
 const m=mount('V2/index.html','screen-codigo');m.button.click();
 assert.deepEqual(m.screens,['screen-cadeado-motivo']);assert.equal(m.hiddenCodes(),1);assert.deepEqual(m.routes,[]);
});
test('mapa e QR voltam ao menu do morador sem depender do histórico',()=>{
 for(const page of ['index.html','qrcode-report.html','dcunha4.html']){
  const m=mount(page,'');m.button.click();assert.deepEqual(m.routes,['https://example.invalid/app/V2/index.html#menu']);
 }
});
test('submenus administrativos voltam à administração',()=>{
 for(const page of ['backoffice.html','relatorio-extintores.html','V2/acessos-admin.html','V2/config-admin.html']){
  const m=mount(page,'');m.button.click();assert.deepEqual(m.routes,['https://example.invalid/app/V2/admin.html']);
 }
});
test('regresso ao menu restaura identidade apenas com sessão ativa',()=>{
 const active=mount('V2/index.html','screen-start',true,'#menu');assert.deepEqual(active.screens,['screen-home']);assert.equal(active.restores(),1);assert.equal(active.button.hidden,true);
 const expired=mount('V2/index.html','screen-start',false,'#menu');assert.deepEqual(expired.screens,['screen-home']);assert.equal(expired.restores(),0);
});
test('registo, recuperação e estados têm destinos definidos; menu principal sem voltar',()=>{
 const m=mount('V2/index.html','screen-home');const parent=m.context.window.DC4Navigation.parentFor;
 assert.equal(m.button.hidden,true);
 assert.equal(parent('V2/index.html','screen-first-use').screen,'screen-start');
 assert.equal(parent('V2/index.html','screen-recover','screen-first-use').screen,'screen-first-use');
 assert.equal(parent('V2/garagem.html','screen-contacto','screen-bloqueado').screen,'screen-bloqueado');
 assert.equal(parent('V2/garagem.html','screen-motivo').menu,true);
});

test('voltar dos extintores mantém o destino menu mesmo sem sessão local',()=>{
 const m=mount('index.html','',false);m.button.click();
 assert.deepEqual(m.routes,['https://example.invalid/app/V2/index.html#menu']);
});
