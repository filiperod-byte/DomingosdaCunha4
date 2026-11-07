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
test('soma base e IVA identificados quando o OCR perde a linha do total',()=>{
 const text=invoice('2026/70','240,00','55,20','295,20','Manutenção')
   .replace('Total da Factura 295,20 €','');
 const result=parse(text);
 assert.equal(result.value,295.2);
 assert.equal(result.iva,55.2);
});
test('não deduz montantes se faltar uma das parcelas',()=>{
 const text=invoice('2026/71','240,00','55,20','295,20','Manutenção')
  .replace('Total da Factura 295,20 €','')
  .replace('IVA 23.00% (Incidência: 240,00) 55,20 €','');
 assert.equal(parse(text).value,null);
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

test('OCR desloca morada para a mesma linha do fornecedor e separa corretamente',()=>{
 const doc=invoice('2026/52','240,00','55,20','295,20','Manutenção')
 .replace('Empresa de Manutenção Unipessoal. Lda.',
   'Empresa de Manutenção Unipessoal. Lda. Impasse Exemplo nº18 1111-222 Lisboa');
 const v=parse(doc);
 assert.equal(v.supplier,'Empresa de Manutenção Unipessoal. Lda.');
});
test('OCR pode apresentar total na linha seguinte e euro à esquerda',()=>{
 const doc=invoice('2026/52','240,00','55,20','295,20','Manutenção')
   .replace('Total da Factura 295,20 €','Total da Factura\n€ 295,20');
 const v=parse(doc);
 assert.equal(v.value,295.2);
});
test('fatura de administração com IVA em linhas separadas e IBAN observado',()=>{
 const doc=[
  'Administração do condomínio', 'Fatura','FT 2026/256',
  'Gestora Exemplo, Unipessoal Lda Rua Exemplo 8',
  'Cons. Reg. Comercial: Lisboa - 517817578',
  'NIF: PT517817578','V/ CONTRIBUINTE','901090409',
  'DATA DE EMISSÃO','DATA VENCIMENTO 2026-10-26','2026-10-26',
  'CÓDIGO DESCRIÇÃO QTD. UNI. PREÇO UNI. % DESC. VALOR SEM IVA IVA',
  'Honorários de administração','referentes ao mês corrente',
  '1.0 un 221,50 221,50 23 %','Pagamento por transferência bancária para o IBAN:',
  'BPI: PT50 0010 0000 6261 8490 0011 6',
  'Taxa','23 %','Base','€ 221,50','Valor','€ 50,95',
  'Valor ilíquido € 221,50','IVA € 50,95','Total a pagar € 272,45'
 ].join('\n');
 const v=parse(doc);
 assert.equal(v.supplier,'Gestora Exemplo, Unipessoal Lda');
 assert.equal(v.number,'FT 2026/256');
 assert.equal(v.nif,'517817578');
 assert.equal(v.issued,'2026-10-26');
 assert.equal(v.due,'2026-10-26');
 assert.equal(v.value,272.45);
 assert.equal(v.subtotal,221.5);
 assert.equal(v.iva,50.95);
 assert.equal(v.documentType,'Fatura');
 assert.equal(v.iban,'PT50001000006261849000116');
});
test('orçamento de honorários não é convertido em fatura e retém o valor c/IVA',()=>{
 const doc=[
 'Empresa Obras Lda','Honorários','Data: 24/06/2026',
 'Orc.: Lis24062026','Documento valido por 15 dias',
 'DESIGNAÇÃO DOS TRABALHOS UN QUANT. P. UNIT. TOTAL',
 '1.1 Deslocação – verificação fuga de água garagem un 1 60,00 €',
 'Sem/IVA 60,00 €','GRUPONA C/IVA23% 73,80 €','Com fatura',
 'NIF: 518 832 287','IBAN: PT50 0036 0081 9910 0149 8890 3'
 ].join('\n');
 const v=parse(doc);
 assert.equal(v.documentType,'Orçamento / honorários');
 assert.equal(v.value,73.8);
 assert.equal(v.subtotal,60);
 assert.equal(v.supplier,'Empresa Obras Lda');
 assert.equal(v.issued,'2026-06-24');
 assert.equal(v.iban,'PT50003600819910014988903');
});
test('nota de despesas sem NIF não inventa identificação fiscal',()=>{
 const doc=['ANTÓNIO C. ROSA','FILIPA GERALDES','Advogados',
 'NOTA DE DESPESAS E HONORÁRIOS','Cliente: Condomínio',
 'Assunto: Carta Câmara','Despesas:','Ctt (1 x 5,72€) 5,72',
 'TOTAL 75,72','Estoril, 27 de setembro de 2026'].join('\n');
 const v=parse(doc);
 assert.equal(v.documentType,'Nota despesas/honorários');
 assert.equal(v.supplier,'ANTÓNIO C. ROSA / FILIPA GERALDES');
 assert.equal(v.value,75.72);
 assert.equal(v.description,'Carta Câmara');
 assert.equal(v.nif,'');
});
