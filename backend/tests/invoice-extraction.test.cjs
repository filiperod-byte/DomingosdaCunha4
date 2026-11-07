const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const src=fs.readFileSync(path.join(__dirname,'../src/invoice-extraction.js'),'utf8');
const parse=new Function(src+'\nreturn fncParseInvoiceText_;')();

const samples=[
  ['2026/52','240,00','55,20','295,20','Manutenção portão 17/04/2026',295.2,240,55.2],
  ['2026/56','180,00','41,40','221,40','Manutenção bomba de água',221.4,180,41.4]
];
function invoice(no,net,tax,total,service){
 return [
  'ATCUD: J6KJ8TB5-52',
  'Factura nº1 '+no,
  'ORIGINAL',
  'De',
  'Empresa de Manutenção Unipessoal. Lda.',
  'Nº Contribuinte: 500000000',
  'Data de Emissão',
  '2026-10-20',
  'Data de Vencimento',
  '2026-10-24',
  'Para',
  'Condomínio',
  'Nº Contribuinte: 900000000',
  'Lista de Artigos',
  'DESCRIÇÃO DO ARTIGO QUANT. PREÇO DESC. IVA (%) TOTAL',
  service+' 1.00 '+net+' 0,00 23.00 '+net,
  'Resumo',
  'Subtotal da Factura '+net+' €',
  'IVA 23.00% (Incidência: '+net+') '+tax+' €',
  'Total da Factura '+total+' €'
 ].join('\n');
}
for(const [number,net,tax,total,service,gross,base,iva] of samples){
 test('extrai total com IVA sem confundir subtotal para '+number,()=>{
  const result=parse(invoice(number,net,tax,total,service));
  assert.equal(result.number,number);
  assert.equal(result.value,gross);
  assert.equal(result.iva,iva);
  assert.equal(result.subtotal,base);
  assert.equal(result.nif,'500000000');
  assert.equal(result.supplier,'Empresa de Manutenção Unipessoal. Lda.');
  assert.equal(result.issued,'2026-10-20');
  assert.equal(result.due,'2026-10-24');
  assert.equal(result.state,'Extraído — por validar');
  assert.match(result.description,/Manutenção/);
 });
}
test('nunca extrai o subtotal quando falta o total final',()=>{
 const text=invoice('2026/70','240,00','55,20','295,20','Manutenção')
   .replace('Total da Factura 295,20 €','');
 const result=parse(text);
 assert.equal(result.value,null);
 assert.equal(result.state,'Extração parcial — por validar');
});
test('valores totais diferentes no mesmo PDF obrigam a revisão',()=>{
 const text=invoice('2026/70','240,00','55,20','295,20','Manutenção')+
  '\nTotal a Pagar 19,00 €';
 const result=parse(text);
 assert.equal(result.value,null);
 assert.equal(result.state,'Extração parcial — por validar');
});
test('rejeita datas impossíveis',()=>{
 const text=invoice('2026/52','240,00','55,20','295,20','Manutenção').replace('2026-10-24','2026-02-30');
 assert.equal(parse(text).due,'');
});

function mockOcrDrive(uploadStatus=200, exportStatus=200){
  const calls=[], files=new Map(), original={
    getMimeType:()=> 'application/pdf',
    getBlob:()=>({getBytes:()=>[37,80,68,70,45]})
  };
  const temp={setTrashed:value=>{calls.push('trashed:'+value)}};
  files.set('original',original);files.set('tmp-ocr',temp);
  const DriveApp={getFileById:id=>{if(!files.has(id))throw Error('File missing');return files.get(id)}};
  const Utilities={getUuid:()=> '1234-abcd',newBlob:value=>({getBytes:()=>[...Buffer.from(value)]})};
  const ScriptApp={getOAuthToken:()=> 'FAKE-TEST-TOKEN'};
  const UrlFetchApp={fetch:(url,opts)=>{
    calls.push(url);
    assert.match(opts.headers.Authorization,/^Bearer FAKE-TEST-TOKEN$/);
    if(url.includes('/upload/drive/v3/files')){
      assert.equal(opts.method,'post');
      assert.match(opts.contentType,/multipart\/related/);
      assert.ok(opts.payload.length>5);
      return {getResponseCode:()=>uploadStatus,
        getContentText:()=>JSON.stringify({id:'tmp-ocr'})};
    }
    assert.match(url,/\/export\?mimeType=text%2Fplain/);
    return {getResponseCode:()=>exportStatus,getContentText:()=>invoice(
      '2026/52','240,00','55,20','295,20','Manutenção portão')};
  }};
  const run=new Function('DriveApp','Utilities','ScriptApp','UrlFetchApp',
    src+'\nreturn fncDriveOcrText_;')(DriveApp,Utilities,ScriptApp,UrlFetchApp);
  return {run,calls};
}
test('OCR copia o PDF para Docs, exporta texto e elimina o documento temporário',()=>{
  const m=mockOcrDrive();
  const text=m.run('original');
  assert.match(text,/295,20/);
  assert.equal(m.calls.filter(x=>x==='trashed:true').length,1);
});
test('se a exportação falhar, o documento temporário ainda é eliminado',()=>{
  const m=mockOcrDrive(200,403);
  assert.throws(()=>m.run('original'),/exportação OCR/i);
  assert.equal(m.calls.filter(x=>x==='trashed:true').length,1);
});
test('se o upload falhar, não é criada uma cópia temporária',()=>{
  const m=mockOcrDrive(403);
  assert.throws(()=>m.run('original'),/HTTP 403/);
  assert.equal(m.calls.filter(x=>x==='trashed:true').length,0);
});
