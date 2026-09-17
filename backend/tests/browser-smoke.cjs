const fs=require('fs'),path=require('path'),assert=require('assert/strict');
(async()=>{
 const {chromium}=require(process.env.PLAYWRIGHT_MODULE || 'playwright-core');
 const browser=await chromium.launch({executablePath:process.env.CHROMIUM_EXECUTABLE,args:['--no-sandbox','--disable-gpu','--disable-dev-shm-usage'],headless:true});
 const root=path.resolve(__dirname,'../..'),base='http://localhost:8899/DomingosdaCunha4/';
 const context=await browser.newContext({viewport:{width:390,height:844},serviceWorkers:'block'});
 const page=await context.newPage();const errors=[];page.on('pageerror',e=>errors.push(e.message));
 let offline=false,reports=0,expire=false,lastReport;
 await context.route('**/*',async route=>{
  const req=route.request(),u=new URL(req.url());
  if(u.hostname==='script.google.com'){
   if(offline)return route.abort();
   const p=JSON.parse(req.postData()||'{}');let result;
   if(p.action==='status')result={success:true,reported:[{floor:-2,point:'G4',reason:'Outro',description:'Manómetro sem pressão',reportedAt:'2026-09-15T10:00:00Z'}]};
   else if(p.action==='garage.publicConfig')result={nomeCondominio:'Teste',structure:[]};
   else if(p.action==='garage.loginPin')result={success:true,token:'synthetic',role:'resident',expiresAt:Date.now()+3600000,nome:'Teste'};
   else if(p.action==='report'){lastReport=p;if(expire){expire=false;result={success:false,code:'AUTH_REQUIRED',message:'Sessão terminada'}}else{reports++;result={success:true}}}
   else result={success:true};
   return route.fulfill({contentType:'application/json',body:JSON.stringify(result)});
  }
  const rel=decodeURIComponent(u.pathname).replace('/DomingosdaCunha4/','');
  const file=path.join(root,rel||'index.html');
  if(!file.startsWith(root)||!fs.existsSync(file))return route.fulfill({status:404,body:''});
  const ext=path.extname(file);return route.fulfill({contentType:({'.html':'text/html','.js':'text/javascript','.json':'application/json','.css':'text/css'})[ext]||'text/plain',body:fs.readFileSync(file)});
 });
 await page.addInitScript(()=>{sessionStorage.setItem('dc4_session_resident',JSON.stringify({token:'synthetic',expiresAt:Date.now()+3600000}));sessionStorage.setItem('dc4_user',JSON.stringify({nome:'Teste',piso:'9',fracao:'B'}))});
 await page.goto(base+'V2/index.html');await page.locator('#screen-home.active').waitFor();
 await page.getByText('Ocorrências',{exact:true}).click();
 await page.getByText('Mapa dos extintores',{exact:true}).click();
 await page.locator('.ext-btn[data-floor="-2"][data-point="G4"]').waitFor();
 assert.equal(await page.locator('.ext-btn[data-floor="-2"]').count(),4);
 await page.locator('.ext-btn[data-floor="-2"][data-point="G4"]').click();
 await page.getByText('Descrição: Manómetro sem pressão',{exact:true}).waitFor();
 await page.locator('#dc4-back').click();await page.locator('.ext-btn[data-point="G4"]').waitFor();
 await page.locator('#dc4-back').click();if(page.url().includes('occurrences.html'))await page.locator('#dc4-back').click();await page.locator('#screen-home.active').waitFor();
 // QR sem sessão: preparar e autenticar apenas ao enviar.
 await page.goto(base+'qrcode-report.html?floor=-2&point=G4');
 await page.evaluate(()=>sessionStorage.removeItem('dc4_session_resident'));
 await page.locator('#reason').selectOption('Outro');await page.locator('#notes').fill('Descrição sintética');
 await page.locator('#submitBtn').click();await page.locator('#authDialog[open]').waitFor();
 assert.equal(reports,0);assert.equal(await page.locator('#notes').inputValue(),'Descrição sintética');
 await page.locator('#pin').fill('123456');await page.locator('#loginSubmit').click();
 await page.waitForFunction(()=>document.getElementById('notes').value==='');assert.equal(reports,1);
 // Expiração no servidor após preparar texto e fotografia.
 await page.goto(base+'qrcode-report.html?floor=-2&point=G4');
 await page.locator('#reason').selectOption('Outro');await page.locator('#notes').fill('Preservar fotografia');
 await page.locator('#fileInput').setInputFiles({name:'teste.png',mimeType:'image/png',buffer:Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=','base64')});
 expire=true;await page.locator('#submitBtn').click();await page.locator('#authDialog[open]').waitFor();
 assert.equal(await page.locator('#notes').inputValue(),'Preservar fotografia');assert.ok(lastReport.photoBase64);
 await page.locator('#pin').fill('123456');await page.locator('#loginSubmit').click();
 await page.waitForFunction(()=>document.getElementById('notes').value==='');assert.equal(reports,2);assert.ok(lastReport.photoBase64);
 await page.goto(base+'qrcode-report.html?floor=-2&point=G4');
 await page.locator('#reason').selectOption('Outro');await page.locator('#notes').fill('Rede indisponível');offline=true;
 await page.locator('#submitBtn').click();await page.locator('.msg.err').waitFor();
 assert.equal(await page.locator('#notes').inputValue(),'Rede indisponível');assert.equal(reports,2);
 offline=false;await page.locator('#submitBtn').click();await page.waitForFunction(()=>document.getElementById('notes').value==='');assert.equal(reports,3);
 await page.evaluate(()=>window.dc4Logout('resident'));assert.equal(await page.evaluate(()=>window.dc4HasSession('resident')),false);
 offline=true;await page.goto(base+'qrcode-report.html?floor=-2&point=G4');
 await page.getByText('Estado por confirmar',{exact:true}).waitFor();
 assert.deepEqual(errors,[]);
 await browser.close();console.log('PASS: browser mobile viewport, menu/map/G4/QR/back/session/report/offline. API mocked; no real writes.');
})().catch(e=>{console.error(e);process.exit(1)});
