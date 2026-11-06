const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const base=path.resolve(__dirname,'../..');
const html=fs.readFileSync(path.join(base,'V2/financas-admin.html'),'utf8');
test('faturas: botão explícito procura PDF no Drive; atualizar lista não repete a procura',()=>{
  assert.match(html,/id="scan">Procurar novas faturas no Drive/);
  assert.match(html,/id="refresh">[^<]*Atualizar lista/);
  assert.match(html,/UI\.api\('finance\.admin\.scanNow'\)/);
  assert.match(html,/UI\.api\('finance\.admin\.invoices'\)/);
  assert.match(html,/UI\.api\('finance\.admin\.scannerStatus'\)/);
});
test('faturas: a monitorização pode ser ativada e tem estado visível',()=>{
  assert.match(html,/id="autoStatus"/);
  assert.match(html,/finance\.admin\.enableScanner/);
  assert.match(html,/finance\.admin\.disableScanner/);
  assert.match(html,/Automática · desligada/);
});
test('faturas: PDF não classificado não aparece como valor zero ou pagamento confirmado',()=>{
  assert.match(html,/Por apurar/);
  assert.match(html,/por classificar/i);
});
test('faturas: scripts inline válidos',()=>{
  for(const script of html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g)){
    if(script[1].trim())assert.doesNotThrow(()=>new Function(script[1]));
  }
});
test('faturas: instruções OAuth visíveis quando o backend não pode consultar triggers',()=>{
 assert.match(html,/id="scannerHelp"/);
 assert.match(html,/authorizationRequired/);
 assert.match(html,/financeInstallScannerTrigger_/);
 assert.match(html,/A pesquisa manual de documentos continua disponível/);
});
