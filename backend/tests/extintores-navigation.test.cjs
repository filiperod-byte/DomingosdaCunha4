const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const root = path.join(__dirname, '../..');

test('mapa consulta apenas status e mantém os pontos com ocorrências abertas', async () => {
 const source=fs.readFileSync(path.join(root,'image-compress.js'),'utf8');
 const start=source.indexOf('loadStatuses = async function loadStatusesOverride()');
 const end=source.indexOf('  renderBuilding =',start);
 const actions=[],registered=[],openSet=new Set();
 const context=vm.createContext({Set,Date,console,clearOccurrenceState(){openSet.clear()},updateLegendText(){},apiGet:async (action,params)=>{assert.equal(params.details,'public');actions.push(action);return {success:true,reported:[{floor:9,point:'E1'}]}},registerOccurrence(item,type){registered.push({item,type});openSet.add(item.floor+':'+item.point)},openSet,els:{lastRefresh:{}},formatTime:()=> '12:00',showToast(){}});
 vm.runInContext(source.slice(start,end),context);
 await context.loadStatuses();
 assert.deepEqual(actions,['status']);
 assert.equal(registered[0].type,'open');
 assert.deepEqual([...context.REPORTED_SET],['9:E1']);
});

function client(page){
 const redirects=[],requests=[];
 const endpoint='https://script.google.com/macros/s/TEST/exec';
 const context=vm.createContext({URL,FormData,Date,Response,document:{currentScript:{src:'https://example.invalid/app/session-client.js'}},location:{href:'https://example.invalid/app/'+page,assign:url=>redirects.push(String(url))},sessionStorage:{getItem:()=>null,removeItem(){},setItem(){}},window:{fetch:async(url,options={})=>{
  if(String(url).endsWith('config.json'))return new Response(JSON.stringify({backendUrl:endpoint}));
  const p=JSON.parse(options.body);requests.push(p);
  return new Response(JSON.stringify(p.action==='status'?{success:true,reported:[]}:{success:false,code:'AUTH_REQUIRED',message:'Sessão inválida'}));
 }}});
 vm.runInContext(fs.readFileSync(path.join(root,'session-client.js'),'utf8'),context);
 return {fetch:context.window.dc4Fetch,redirects,requests,endpoint};
}
test('recusa administrativa não encaminha o mapa dos moradores para administração', async()=>{
 const c=client('index.html');
 await assert.rejects(c.fetch(c.endpoint+'?action=pendingOccurrences'),/Sessão inválida/);
 assert.deepEqual(c.redirects,[]);
});
test('backoffice continua a encaminhar sessões inválidas para o login administrativo',async()=>{
 const c=client('backoffice.html');
 await assert.rejects(c.fetch(c.endpoint+'?action=pendingOccurrences'),/Sessão inválida/);
 assert.equal(c.redirects.length,1);assert.match(c.redirects[0],/V2\/admin.html$/);
});
test('consulta sem action conserva o status quando convertida para POST',async()=>{
 const c=client('index.html');await c.fetch(c.endpoint);
 assert.equal(c.requests[0].action,'status');assert.equal(c.requests[0]._method,'GET');
});
