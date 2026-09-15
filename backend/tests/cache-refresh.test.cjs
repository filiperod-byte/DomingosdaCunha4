const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm');
test('configuração no-store e navegação não recebem a cache antiga quando a rede responde',async()=>{
 const handlers={};let requests=0;
 const context={URL,Response,location:{origin:'https://example.invalid'},self:{addEventListener:(name,fn)=>handlers[name]=fn},caches:{match:async()=>new Response('antigo')},fetch:async()=>{requests++;return new Response('atual')}};
 vm.runInNewContext(fs.readFileSync('service-worker.js','utf8'),context);
 for(const req of [{method:'GET',url:'https://example.invalid/config.json',cache:'no-store'},{method:'GET',url:'https://example.invalid/index.html',mode:'navigate'}]){
  let result;handlers.fetch({request:req,respondWith:p=>result=p});assert.equal(await(await result).text(),'atual');
 }
 assert.equal(requests,2);
});
