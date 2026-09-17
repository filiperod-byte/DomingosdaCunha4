const fs=require('fs'),path=require('path'),assert=require('assert/strict');
const {harness,bundled}=require('./harness.cjs');
(async()=>{
const {chromium}=require(process.env.PLAYWRIGHT_MODULE);
const browser=await chromium.launch({executablePath:process.env.CHROMIUM_EXECUTABLE,args:['--no-sandbox','--disable-gpu','--disable-dev-shm-usage'],headless:true});
try{
const b=harness(bundled());b.initialize();b.context.setupGeneralOccurrences();
b.sheets.get('CONFIG').rows.find(r=>r[0]==='ADMIN_PIN')[1]='98765432';
const residents=b.sheets.get('CONDOMINOS');const resident={ID:'r1',Nome:'Teste',Piso:'9',Fracao:'B',Estado:'APROVADO',PIN:'654321',PINAtivo:true};residents.rows.push(residents.rows[0].map(h=>resident[h]??''));
const admin=b.post('garage.loginAdmin',{email:'admin@example.invalid',pin:'98765432'});
const root=path.resolve(__dirname,'../..'),base='http://localhost:8899/DomingosdaCunha4/';
const context=await browser.newContext({viewport:{width:390,height:844},serviceWorkers:'block'});const page=await context.newPage(),errors=[];page.on('pageerror',e=>errors.push(e.message));
let loseResponse=true,lastRequest,offline=false,expire=false;
await context.route('**/*',async route=>{const req=route.request(),u=new URL(req.url());if(u.hostname==='script.google.com'){
 if(offline)return route.abort();const p=JSON.parse(req.postData()||'{}');
 if(p.action==='general.report'){lastRequest=p;if(expire){expire=false;const key=Object.keys(b.props).find(k=>k.startsWith('DC4_SESSION_')&&JSON.parse(b.props[k]).role==='resident');const s=JSON.parse(b.props[key]);s.expiresAt=0;b.props[key]=JSON.stringify(s);}}
 const result=b.post(p.action,p);if(result.token)result.expiresAt=Date.now()+3600000;
 if(p.action==='general.report'&&result.success&&loseResponse){loseResponse=false;return route.abort();}
 return route.fulfill({contentType:'application/json',body:JSON.stringify(result)});
 }
 if(u.hostname!=='localhost')return route.fulfill({status:404,body:''});const file=path.join(root,decodeURIComponent(u.pathname).replace('/DomingosdaCunha4/',''));if(!file.startsWith(root)||!fs.existsSync(file))return route.fulfill({status:404,body:''});return route.fulfill({contentType:({'.html':'text/html','.js':'text/javascript','.css':'text/css','.json':'application/json','.svg':'image/svg+xml','.png':'image/png'})[path.extname(file)]||'text/plain',body:fs.readFileSync(file)});
});
await page.goto(base+'occurrences.html');await page.locator('.facade-row').last().waitFor();assert.equal(await page.locator('.facade-row').count(),13);
await page.screenshot({path:'/tmp/dc4-building.png',fullPage:true});
await page.locator('.facade-row[href="occurrences.html?floor=-2"]').click();await page.getByText('Extintores / Segurança contra incêndio',{exact:true}).waitFor();
await page.getByText('Extintores / Segurança contra incêndio',{exact:true}).click();await page.locator('#equipment').selectOption('G4');await page.locator('#reason').waitFor();assert.ok(page.url().includes('general-report.html'));await page.locator('#dc4-back').click();assert.ok(page.url().includes('occurrences.html?floor=-2'));
await page.getByText('Iluminação',{exact:true}).click();await page.locator('#reason').selectOption('Luz apagada');await page.locator('#notes').fill('Lâmpada fundida junto à escada');
await page.locator('#fileInput').setInputFiles({name:'teste.png',mimeType:'image/png',buffer:Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=','base64')});
await page.locator('#submitBtn').click();await page.locator('#authDialog[open]').waitFor();assert.equal(b.sheets.get('OCORRENCIAS_GERAIS').rows.length,1);
await page.locator('#pin').fill('654321');await page.locator('#loginSubmit').click();await page.getByText('Repetir envio',{exact:true}).waitFor();
const requestId=lastRequest.requestId;assert.ok(lastRequest.photoBase64);assert.equal(b.sheets.get('OCORRENCIAS_GERAIS').rows.length,2);assert.equal(await page.locator('#notes').inputValue(),'Lâmpada fundida junto à escada');
await page.locator('#submitBtn').click();await page.getByText('Reporte recebido. Aguarda validação da administração.',{exact:true}).waitFor();assert.equal(lastRequest.requestId,requestId);assert.equal(b.sheets.get('OCORRENCIAS_GERAIS').rows.length,2);assert.equal(b.counts.files,1);
// Administração publica apenas o texto revisto.
await page.evaluate(s=>sessionStorage.setItem('dc4_session_admin',JSON.stringify(s)),{token:admin.token,expiresAt:Date.now()+3600000});
await page.goto(base+'occurrences-admin.html');await page.locator('form[data-id]').waitFor();await page.locator('[data-copy]').click();await page.locator('[name=status]').selectOption('EM_ANALISE');await page.locator('[name=internalNotes]').fill('NOTA INTERNA PRIVADA');await page.locator('button[type=submit]').click();await page.getByText('Lista atualizada.',{exact:true}).waitFor();
assert.equal(b.get('general.status').occurrences[0].description,'Lâmpada fundida junto à escada');assert.ok(!JSON.stringify(b.get('general.status')).includes('NOTA INTERNA'));
await page.goto(base+'general-report.html?floor=-2&category=iluminacao');await page.getByText('Também verifiquei esta situação',{exact:true}).click();await page.getByText('Confirmação registada. Não foi criada outra ocorrência.',{exact:true}).waitFor();assert.equal(b.get('general.status').occurrences[0].confirmations,1);
// Sessão expira no servidor; formulário e fotografia sobrevivem.
await page.locator('#reason').selectOption('Outro');await page.locator('#notes').fill('Outra luz intermitente');expire=true;await page.locator('#submitBtn').click();await page.locator('#authDialog[open]').waitFor();assert.equal(await page.locator('#notes').inputValue(),'Outra luz intermitente');await page.locator('#pin').fill('654321');await page.locator('#loginSubmit').click();await page.getByText('Reporte recebido. Aguarda validação da administração.',{exact:true}).waitFor();
await page.goto(base+'general-report.html?scope=general');await page.locator('#category').selectOption('elevadores');await page.locator('#reason').selectOption('Fora de serviço');await page.locator('#locationDetail').fill('Elevador esquerdo');await page.locator('#notes').fill('Parado no piso 0');await page.locator('#submitBtn').click();await page.getByText('Reporte recebido. Aguarda validação da administração.',{exact:true}).waitFor();
// QR de piso 0 sem login nem ponto de extintor.
await page.evaluate(()=>sessionStorage.removeItem('dc4_session_resident'));await page.goto(base+'general-report.html?floor=0');await page.locator('#category').waitFor();assert.equal(await page.locator('#authDialog').evaluate(e=>e.open),false);assert.equal(await page.locator('#pointCard .point-title').textContent(),'Piso 0');
await page.locator('#category').selectOption('extintor');await page.locator('#equipment').waitFor();assert.equal(await page.locator('#equipment option[value=CF1]').count(),1);
await page.goto(base+'floor-qrcodes.html');await page.locator('.qr-label').last().waitFor();assert.equal(await page.locator('.qr-label').count(),13);for(const link of await page.locator('.qr-label a').evaluateAll(a=>a.map(x=>x.href))){assert.ok(link.includes('general-report.html?floor='));assert.ok(!link.includes('point='));}
await page.getByText('Limpar seleção',{exact:true}).click();await page.locator('input[aria-label="Imprimir piso -2"]').check();await page.evaluate(()=>window.print=()=>window.printCalled=true);await page.getByText('Imprimir selecionados / Guardar PDF',{exact:true}).click();await page.waitForFunction(()=>window.printCalled===true);await page.emulateMedia({media:'print'});await page.pdf({path:'/tmp/dc4-floor-label.pdf',format:'A4',printBackground:true});assert.equal(await page.locator('.qr-label:visible').count(),1);
await page.emulateMedia({media:'screen'});await page.goto(base+'occurrences.html?floor=-2');await page.getByText('Lâmpada fundida junto à escada',{exact:true}).waitFor();await page.screenshot({path:'/tmp/dc4-floor.png',fullPage:true});

// Existing extinguisher QR aliases preselect the same shared form.
for(const entry of ['qrcode-report.html','index.html','dcunha4.html']){
 await page.goto(base+entry+'?floor=0&point=CF1');await page.locator('#equipment').waitFor();
 assert.equal(await page.locator('#equipment').inputValue(),'CF1');assert.equal(await page.locator('#category').inputValue(),'extintor');
 assert.ok(page.url().includes('general-report.html'));assert.equal(await page.locator('#authDialog').evaluate(e=>e.open),false);
}
const rt=b.post('garage.loginPin',{pin:'654321'}).token;
const ex=b.post('report',{token:rt,floor:-2,point:'G4',reason:'Extintor em falta'});
assert.equal(b.post('approveOccurrence',{token:admin.token,occurrenceId:ex.occurrenceId,floor:-2,point:'G4'}).success,true);
await page.goto(base+'occurrences.html');await page.waitForFunction(()=>document.getElementById('loadState').textContent.startsWith('Atualizado'));
assert.equal(await page.locator('a[href="occurrences.html?floor=-2"] .occ-count.alert').innerText(),'2');
assert.equal(await page.locator('a[href="occurrences.html?floor=10"] .occ-count.ok').innerText(),'0');
assert.equal(await page.getByText('Mapa dos extintores',{exact:true}).count(),0);
assert.equal(await page.locator('[data-point]').count(),0);
await page.screenshot({path:'/tmp/dc4-building.png',fullPage:true});
offline=true;await page.locator('#refresh').click();await page.getByText('Estado por confirmar. Não foi possível atualizar. Tente novamente.',{exact:true}).waitFor();
assert.equal(await page.locator('.occ-count.unknown').count(),13);assert.deepEqual(errors,[]);console.log('PASS: integrated browser + backend emulator: floors, dedicated extinguisher, report/login/photo, lost response replay, administration/privacy, confirmation, expiry, general scope, floor 0, QR print, offline.');
}finally{await browser.close();}
})().catch(e=>{console.error(e);process.exit(1)});
