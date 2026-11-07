/**
 * Extração conservadora de dados de faturas via OCR do Google Drive.
 * Usa exclusivamente as permissões da conta que executa o Apps Script.
 * O PDF original não é modificado. Documento OCR temporário é enviado para o lixo.
 * Nunca aprova faturas nem cria pagamentos.
 */

/** Formatos portugueses (1.234,56 / 1234.56 / 1234,56). */
function fncParseEuro_(value){
  if(value===null||value===undefined)return null;
  let s=String(value).replace(/[\s\u00a0€]/g,'').replace(/[^\d.,-]/g,'');
  if(!s)return null;
  const c=s.lastIndexOf(','),p=s.lastIndexOf('.');
  if(c>=0&&p>=0)s=c>p?s.replace(/\./g,'').replace(',','.'):s.replace(/,/g,'');
  else if(c>=0)s=s.replace(',','.');
  else if(p>=0&&s.length-p-1===3)s=s.replace(/\./g,'');
  const n=Number(s);
  return Number.isFinite(n)&&n>=0&&n<=1000000?Math.round(n*100)/100:null;
}
function fncISOInvoiceDate_(value){
  const s=String(value||'').trim();
  let y,m,d;
  if(/^\d{4}[-/]\d{2}[-/]\d{2}$/.test(s)){[y,m,d]=s.split(/[-/]/).map(Number);}
  else if(/^\d{2}[-/]\d{2}[-/]\d{4}$/.test(s)){[d,m,y]=s.split(/[-/]/).map(Number);}
  else return '';
  const dt=new Date(Date.UTC(y,m-1,d));
  if(dt.getUTCFullYear()!==y||dt.getUTCMonth()!==m-1||dt.getUTCDate()!==d)return '';
  return [y,String(m).padStart(2,'0'),String(d).padStart(2,'0')].join('-');
}
function fncFirstMatch_(text,expressions){
  for(const exp of expressions){const m=text.match(exp);if(m&&m[1])return String(m[1]).trim();}
  return '';
}
/** Pura: só extrai factos com etiquetas inequívocas, nunca usa nome do PDF como valor. */
function fncParseInvoiceText_(source){
  const text=String(source||'').replace(/\r/g,'').replace(/[\u00a0]/g,' ').slice(0,80000);
  const m=text.match(/\b(?:factura|fatura)\s*(?:n[º°.o]?\s*)?(?:\d+\s+)?(\d{4}\/\d{1,10})\b/i);
  let number=m?m[1]:fncFirstMatch_(text,[
    /\b(?:Factura|Fatura)\s*(?:n[º°.]?\s*)?(FT\s*[\w/. -]*\d)\b/i,
    /\b(FT\s*\d{4}[\w/. -]*\d)\b/i
  ]);
  if(number)number=number.replace(/\s+/g,' ').slice(0,90);
  const issuance=fncISOInvoiceDate_(fncFirstMatch_(text,[
    /Data\s+(?:de\s+)?Emiss[aã]o\s*[:\s]+(\d{4}[-/]\d{2}[-/]\d{2}|\d{2}[-/]\d{2}[-/]\d{4})/i,
    /\bData\s+(?:da\s+)?Fatura\s*[:\s]+(\d{4}[-/]\d{2}[-/]\d{2}|\d{2}[-/]\d{2}[-/]\d{4})/i
  ]));
  const due=fncISOInvoiceDate_(fncFirstMatch_(text,[
    /Data\s+de\s+Vencimento\s*[:\s]+(\d{4}[-/]\d{2}[-/]\d{2}|\d{2}[-/]\d{2}[-/]\d{4})/i,
    /Data\s+Limite\s+(?:de\s+)?Pagamento\s*[:\s]+(\d{4}[-/]\d{2}[-/]\d{2}|\d{2}[-/]\d{2}[-/]\d{4})/i
  ]));
  let gross=fncParseEuro_(fncFirstMatch_(text,[
    /^\s*Total\s+(?:da\s+)?(?:Factura|Fatura)\s*[:\s]+([\d.,]+)\s*€/im,
    /^\s*Total\s+a\s+Pagar\s*[:\s]+([\d.,]+)\s*€/im,
    /^\s*Valor\s+Total\s*[:\s]+([\d.,]+)\s*€/im,
    /^\s*Total\s+(?:do\s+)?Documento\s*[:\s]+([\d.,]+)\s*€/im
  ]));
  // Documentos com várias faturas na mesma página (ex.: eletricidade + audiovisual)
  // não podem receber um total escolhido arbitrariamente.
  const totals=[...text.matchAll(/^\s*(?:Total\s+(?:da\s+)?(?:Factura|Fatura)|Total\s+a\s+Pagar|Valor\s+Total|Total\s+(?:do\s+)?Documento)\s*[:\s]+([\d.,]+)\s*€/gim)]
    .map(m=>fncParseEuro_(m[1])).filter(v=>v!==null);
  const uniqueTotals=[...new Set(totals)];
  if(uniqueTotals.length>1)gross=null;
  let subtotal=null,iva=null;
  const tax=text.match(/\bIVA\s*[\d.,]+\s*%\s*\(\s*Incid[eê]ncia\s*:\s*([\d.,]+)\s*\)\s*([\d.,]+)\s*€/i);
  if(tax){subtotal=fncParseEuro_(tax[1]);iva=fncParseEuro_(tax[2]);}
  else{
    subtotal=fncParseEuro_(fncFirstMatch_(text,[
      /Subtotal\s+(?:da\s+)?(?:Factura|Fatura)\s*[:\s]+([\d.,]+)\s*€/i,
      /Valor\s+sem\s+IVA\s*[:\s]+([\d.,]+)\s*€/i
    ]));
  }
  let vendor='',nif='';
  const issuer=text.split(/\n\s*Para\s*\n/i)[0].split(/\n\s*De\s*\n/i).slice(-1)[0];
  const lines=issuer.split('\n').map(x=>x.trim()).filter(Boolean);
  vendor=lines.find(x=>/\b(?:unipessoal|Lda\.?|S\.?\s*A\.?|Limitada)\b/i.test(x)&&x.length<=150)||'';
  nif=fncFirstMatch_(issuer,[
    /N[ºo°.]?\s*Contribuinte\s*:\s*(\d{9})\b/i,
    /\b(?:NIF|NIPC)\s*[:\s]+(\d{9})\b/i
  ]);
  const article=text.match(/Lista\s+de\s+Artigos\s*[\s\S]*?\n[^\n]*(?:DESCRI[CÇ][AÃ]O|ARTIGO)[^\n]*\n\s*([^\n]{8,180})/i);
  let description=article?article[1].trim():'';
  if(description)description=description.replace(/\s+\d+\.\d{2}\s+[\d.,]+\s+[\d.,]+\s+[\d.,]+\s+[\d.,]+\s*$/,'').slice(0,280);
  const method=/D[eé]bito\s+Diret[oa]/i.test(text)?'Débito direto':'Por confirmar';
  const vatCheck=(gross!==null&&subtotal!==null&&iva!==null&&Math.abs(gross-subtotal-iva)<=0.03);
  if(subtotal!==null&&iva!==null&&!vatCheck){subtotal=null;iva=null;}
  const strong=Boolean(vendor&&gross!==null&&number&&issuance);
  return {supplier:vendor,nif,number,issued:issuance,due,value:gross,iva,subtotal,description,method,
    state:strong?'Extraído — por validar':'Extração parcial — por validar',
    recognized:[vendor,number,issuance,due,gross!==null?String(gross):'',description].filter(Boolean).length};
}
/**
 * Converte PDF/imagem para Docs com OCR da Drive API v3 e exporta texto simples.
 * Não requer serviço avançado "Drive", mas usa UrlFetchApp + token da conta Apps Script.
 */
function fncDriveOcrText_(fileId){
  const file=DriveApp.getFileById(fileId);
  const mime=file.getMimeType();
  if(!['application/pdf','image/jpeg','image/png'].includes(mime))
    throw new Error('Formato não suportado para extração');
  const blob=file.getBlob(),media=blob.getBytes();
  if(media.length>7*1024*1024)throw new Error('PDF acima de 7 MB — revisão manual');
  const token=ScriptApp.getOAuthToken(),boundary='dc4ocr'+Utilities.getUuid().replace(/-/g,'');
  const metadata=JSON.stringify({name:'DC4-OCR-TEMP-'+Utilities.getUuid(),
    mimeType:'application/vnd.google-apps.document'});
  const head='--'+boundary+'\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n'+metadata+
    '\r\n--'+boundary+'\r\nContent-Type: '+mime+'\r\n\r\n';
  const tail='\r\n--'+boundary+'--\r\n';
  const bytes=Utilities.newBlob(head).getBytes().concat(media,Utilities.newBlob(tail).getBytes());
  const options={method:'post',contentType:'multipart/related; boundary='+boundary,
    headers:{Authorization:'Bearer '+token},payload:bytes,muteHttpExceptions:true};
  const upload=UrlFetchApp.fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&ocrLanguage=pt&fields=id,mimeType',
    options);
  if(upload.getResponseCode()<200||upload.getResponseCode()>=300)
    throw new Error('Conversão do PDF indisponível (Google Drive HTTP '+upload.getResponseCode()+
      '). Verificar OAuth Drive e UrlFetchApp no editor.');
  let tempId='';
  try{tempId=JSON.parse(upload.getContentText()).id;}catch(_){}
  if(!tempId)throw new Error('Google Drive não devolveu documento OCR');
  try{
    const txt=UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files/'+encodeURIComponent(tempId)+'/export?mimeType=text%2Fplain',
      {headers:{Authorization:'Bearer '+token},muteHttpExceptions:true});
    if(txt.getResponseCode()!==200)throw new Error('A exportação OCR não foi concluída');
    const content=txt.getContentText('UTF-8');
    if(!String(content).trim())throw new Error('Documento sem texto reconhecido');
    return content;
  }finally{
    try{DriveApp.getFileById(tempId).setTrashed(true);}catch(e){console.warn('OCR temporário não foi eliminado: '+e.message);}
  }
}
/** Aplica somente dados reconhecidos a documento por validar. */

/** Acrescenta só os campos técnicos ausentes, sem alterar os já existentes. */
function fncEnsureExtractionColumns_(){
  const sheet=fncTab_(FNC.sheet.invoices);
  const cols=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(String);
  const extras=['IVA (€)','Base tributável (€)','Estado extração','Data extração'].filter(c=>!cols.includes(c));
  if(extras.length)sheet.getRange(1,sheet.getLastColumn()+1,1,extras.length).setValues([extras]);
}
/** Processa um pequeno lote de faturas não validadas; não precisa de trigger adicional. */
function fncExtractPendingInvoices_(maxInvoices){
  const count=Math.min(2,Math.max(0,Number(maxInvoices)||0));
  if(!count)return {processed:0,extracted:0,failed:0};
  const pending=fncData_(FNC.sheet.invoices)
    .filter(r=>r['Validação']==='Por validar' && r['Pagamento']!=='Pago' &&
      r['ID Drive'] && !r['Estado extração'] &&
      (r['Fornecedor']==='Por identificar'||!r['Total (€)']))
    .slice(0,count);
  if(!pending.length)return {processed:0,extracted:0,failed:0};
  fncEnsureExtractionColumns_();
  let success=0,failed=0;
  for(const invoice of pending){
    try{
      const r=fncExtractOneInvoice_(invoice.ID);
      if(r.ok)success++;
    }catch(err){
      failed++;
      console.error('Falha na extração de '+invoice.ID+': '+String(err.message||err));
      // Falhas não bloqueiam novas faturas nem repetem OCR em cada abertura.
      fncLock_(()=>{
        const row=fncData_(FNC.sheet.invoices).find(x=>x.ID===invoice.ID);
        if(row && row['Validação']==='Por validar')
          fncUpdate_(FNC.sheet.invoices,row.row,{
            'Estado extração':'Falhou — revisão manual',
            'Data extração':fncNow_()
          });
      });
    }
  }
  return {processed:pending.length,extracted:success,failed};
}
function fncExtractOneInvoice_(invoiceId){
  const original=fncData_(FNC.sheet.invoices).find(r=>String(r.ID)===String(invoiceId));
  if(!original)throw new Error('Documento não encontrado');
  if(original['Validação']==='Validada'||original['Pagamento']==='Pago')throw new Error('Documento fechado');
  const fileId=original['ID Drive'];
  if(!fileId)throw new Error('Documento sem ID de Drive');
  // O OCR não mantém bloqueios globais enquanto aguarda o Google Drive.
  fncEnsureExtractionColumns_();
  const parsed=fncParseInvoiceText_(fncDriveOcrText_(String(fileId)));
  return fncLock_(()=>{
    // Conferir novamente o registo ANTES da escrita; pode ter sido validado enquanto o OCR corria.
    const item=fncData_(FNC.sheet.invoices).find(r=>String(r.ID)===String(invoiceId));
    if(!item||item['Validação']==='Validada'||item['Pagamento']==='Pago')
      throw new Error('Documento foi entretanto validado; extração não aplicada');
    if(String(item['ID Drive'])!==String(fileId))throw new Error('Documento original alterado');
    const patch={'Estado extração':parsed.state,'Data extração':fncNow_()};
  const set=(field,value,placeholder)=>{
    if(value===null||value===undefined||value==='')return;
    const current=item[field];
    const empty=current===''||current==null||
      (field==='Total (€)'&&current===0);
    if(empty||(placeholder&&String(current)===String(placeholder)))patch[field]=value;
  };
  if(parsed.recognized>=2){
    if(parsed.number)set('Tipo','Fatura','Documento por classificar');
    set('Fornecedor',parsed.supplier,'Por identificar');
    set('NIF',parsed.nif);
    // Um número de documento corrigido manualmente nunca deve ser substituído.
    set('Documento',parsed.number,item['Descrição']);
    set('Emissão',parsed.issued);
    set('Vencimento',parsed.due);
    set('Total (€)',parsed.value);
    if(parsed.method==='Débito direto')set('Método',parsed.method,'Por confirmar');
    set('Descrição',parsed.description,item['Documento']);
    set('IVA (€)',parsed.iva);
    set('Base tributável (€)',parsed.subtotal);
  }
  const details=['Extração automática — confirmar com o PDF original antes de validar'];
  if(parsed.value!==null)details.push('total '+parsed.value.toFixed(2)+' €');
  if(parsed.iva!==null)details.push('IVA '+parsed.iva.toFixed(2)+' €');
  patch['Observações']=details.join('; ');
  const keys=fncTab_(FNC.sheet.invoices).getRange(1,1,1,fncTab_(FNC.sheet.invoices).getLastColumn())
    .getValues()[0].map(String);
  for(const k of Object.keys(patch))if(!keys.includes(k))delete patch[k];
  fncUpdate_(FNC.sheet.invoices,item.row,patch);
  return {ok:true,id:item.ID,extraction:parsed.state,recognized:parsed.recognized,amount:parsed.value};
  });
}
