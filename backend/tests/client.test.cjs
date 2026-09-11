const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const vm=require('node:vm');
const {harness,bundled}=require('./harness.cjs');
const endpoint='https://script.google.com/macros/s/TEST/exec';
function client(b){
 const storage=new Map(),requests=[];
 const fetch=async(url,options={})=>{
  url=String(url);requests.push({url,options});
  if(url.endsWith('config.json'))return new Response(JSON.stringify({backendUrl:endpoint}));
  if(url===endpoint){const p=JSON.parse(options.body);return new Response(JSON.stringify(b.post(p.action,p)));}
  return new Response('static');
 };
 const context=vm.createContext({URL,FormData,Response,Date:class extends Date {static now(){return new Date('2026-09-11T10:00:00.000Z').getTime();}},JSON,window:{fetch},document:{currentScript:{src:'https://example.invalid/app/session-client.js'}},location:{href:'https://example.invalid/app/V2/index.html',assign(){}},sessionStorage:{getItem:k=>storage.get(k)||null,setItem:(k,v)=>storage.set(k,v),removeItem:k=>storage.delete(k)}});
 vm.runInContext(fs.readFileSync(path.join(__dirname,'../../session-client.js'),'utf8'),context);
 return {...context.window,requests,storage};
}
test('cliente e backend: login, consulta por POST, isolamento do token e logout',async()=>{
 const b=harness(bundled());b.initialize();
 const rows=b.sheets.get('CONFIG').rows;
 rows.find(r=>r[0]==='ADMIN_PIN')[1]='98765432';rows.find(r=>r[0]==='ADMIN_EMAIL')[1]='admin@example.invalid';
 const c=client(b);
 const login=await c.dc4Fetch(endpoint,{method:'POST',body:JSON.stringify({action:'garage.loginAdmin',email:'admin@example.invalid',pin:'98765432'})});
 assert.ok((await login.json()).token);
 const response=await c.dc4Fetch(endpoint+'?action=garage.dashboard');assert.equal((await response.json()).ok,true);
 const last=c.requests.at(-1);assert.equal(last.url,endpoint);assert.equal(last.options.method,'POST');assert.ok(JSON.parse(last.options.body).token);
 await c.dc4Fetch('https://script.google.com/macros/s/OTHER/exec');assert.equal(c.requests.at(-1).options.body,undefined);
 await c.dc4Logout('admin');assert.equal(c.dc4HasSession('admin'),false);
 const sessions=Object.keys(b.props).filter(k=>k.startsWith('DC4_SESSION_'));assert.equal(sessions.length,0);
});
test('ficheiros JavaScript e scripts inline alterados têm sintaxe válida',()=>{
 const root=path.join(__dirname,'../..');
 for(const entry of fs.readdirSync(root,{recursive:true})){
  const file=path.join(root,entry);if(!fs.statSync(file).isFile()||entry.startsWith('backend/'))continue;
  if(entry.endsWith('.js'))new vm.Script(fs.readFileSync(file,'utf8'),{filename:entry});
  if(entry.endsWith('.html'))for(const match of fs.readFileSync(file,'utf8').matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/gi))if(match[1].trim())new vm.Script(match[1],{filename:entry});
 }
});
