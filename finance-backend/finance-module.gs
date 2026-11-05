/**
 * DC4 Finance — módulo para Google Apps Script (pré-produção)
 * NÃO PUBLICAR sem integrar no doPost e validação de sessão existentes.
 * Propriedades do script: FINANCE_SPREADSHEET_ID; FINANCE_PHOTOS_FOLDER_ID.
 * O router do projeto deve VALIDAR O TOKEN ANTES de chamar financeDispatch_(req, actor).
 * actor = { role: 'admin'|'resident', email: email validado, fractionId: fração validada }.
 * Nunca confiar em role, email ou fractionId enviados pelo browser.
 */
const FNC = {
  sheet: {
    invoices: 'FATURAS', settings:'WALLBOX_CONFIG', fractions:'WALLBOX_FRACOES',
    readings:'LEITURAS_WALLBOX', payments:'WALLBOX_PAGAMENTOS',
    discounts:'WALLBOX_DESCONTOS', tariff:'TARIFAS_EDP', history:'HISTORICO',
    cycles:'WALLBOX_CICLOS'
  }
};

function financeDispatch_(req, actor) {
  if (!actor || !actor.role || !actor.email) throw new Error('AUTH_REQUIRED');
  const action = String(req.action || '');
  if (!/^finance\.(admin|resident)\./.test(action)) throw new Error('Ação não autorizada');
  if (action.indexOf('finance.admin.')===0 && actor.role!=='admin') throw new Error('FORBIDDEN');
  if (action.indexOf('finance.resident.')===0 && actor.role!=='resident') throw new Error('FORBIDDEN');
  const map = {
    'finance.admin.bootstrap': () => fncBootstrap_(),
    'finance.admin.saveDiscount': () => fncSaveDiscount_(req, actor),
    'finance.admin.upsertFraction': () => fncUpsertFraction_(req, actor),
    'finance.admin.deactivateFraction': () => fncDeactivateFraction_(req, actor),
    'finance.admin.validateInvoice': () => fncValidateInvoice_(req, actor),
    'finance.admin.markInvoicePaid': () => fncMarkInvoicePaid_(req, actor),
    'finance.admin.validateReading': () => fncValidateReading_(req, actor),
    'finance.admin.markWallboxPaid': () => fncMarkWallboxPaid_(req, actor),
    'finance.resident.status': () => fncResidentStatus_(actor),
    'finance.resident.submitReading': () => fncSubmitReading_(req, actor)
  };
  if (!map[action]) throw new Error('Ação financeira desconhecida');
  return map[action]();
}
function fncDb_() {
  const id=PropertiesService.getScriptProperties().getProperty('FINANCE_SPREADSHEET_ID');
  if (!id) throw new Error('FINANCE_SPREADSHEET_ID por configurar');
  return SpreadsheetApp.openById(id);
}
function fncTab_(name) {
  const sh=fncDb_().getSheetByName(name);
  if (!sh) throw new Error('Separador não encontrado: '+name);
  return sh;
}
function fncData_(name) {
  const values=fncTab_(name).getDataRange().getValues();
  if (!values.length) return [];
  const keys=values[0].map(String);
  return values.slice(1).map((row,i)=>({row:i+2,...Object.fromEntries(keys.map((key,j)=>[key,row[j]]))}))
    .filter(r=>Object.values(r).some(v=>v!=='' && v!=null && v!==r.row));
}
function fncUpdate_(name,row,patch) {
  const sh=fncTab_(name), headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  Object.entries(patch).forEach(([col,v])=>{
    const i=headers.indexOf(col);
    if (i<0) throw new Error('Coluna inexistente '+col);
    sh.getRange(row,i+1).setValue(v);
  });
}
function fncAppend_(name,object) {
  const sh=fncTab_(name), headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  sh.appendRow(headers.map(h => object[h]===undefined?'':object[h]));
}
function fncLock_(fn) {
  const lock=LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function fncNow_(){return new Date();}
function fncNumber_(v){
  const s=String(v).trim().replace(/\s/g,'').replace(',','.');
  const x=Number(s);
  if(s==='' || !Number.isFinite(x)) throw new Error('Número inválido');
  return x;
}
function fncRequired_(val,label){if (!String(val??'').trim())throw new Error(label+' obrigatório');return String(val).trim();}
function fncAudit_(id,action,before,after,actor,note) {
  fncAppend_(FNC.sheet.history,{'Data':fncNow_(),'ID documento':id,
    'Ação':action,'Antes':JSON.stringify(before),'Depois':JSON.stringify(after),
    'Utilizador':actor.email,'Notas':note||'','Origem':'App DC4'});
}
function fncSettings_(){
  const rows=fncData_(FNC.sheet.settings);
  return Object.fromEntries(rows.map(r=>[String(r['Chave']),r['Valor']]));
}
function fncBootstrap_(){
  const s=fncSettings_(),discount=s.DESCONTO_FIXOS;
  return {ok:true,settings:{discount:discount===''?null:Number(discount),
    notifications:String(s.NOTIFICACOES_ATIVAS).toUpperCase()==='SIM'},
    fractions:fncData_(FNC.sheet.fractions).map(r=>({
      id:r['ID fração'],name:r['Fração'],status:r['Estado'],email:r['Email autorizado'],
      startDate:r['Data adesão'],initialReading:r['Leitura inicial (kWh)'],
      initialDate:r['Data leitura inicial'],endDate:r['Data desativação'],row:r.row
    })),
    invoices:fncData_(FNC.sheet.invoices).map(r=>({
      id:r.ID, supplier:r['Fornecedor'],description:r['Descrição'],due:r['Vencimento'],
      value:r['Total (€)'],validation:r['Validação'],payment:r['Pagamento'],
      documentType:r['Tipo'],url:r['PDF original']
    })),
    readings:fncData_(FNC.sheet.readings).map(r=>({
      id:r['ID leitura'],fractionId:r['Fração'],cycle:r['Ciclo EDP'],date:r['Data leitura'],
      kwh:r['Leitura acumulada kWh'],state:r['Estado'],row:r.row
    })),
    payments:fncData_(FNC.sheet.payments).map(r=>({
      id:r['ID apuramento'],fractionId:r['Fração'],cycle:r['Ciclo EDP'],
      total:r['Total a pagar €'],state:r['Estado'],row:r.row
    }))};
}
function fncSaveDiscount_(req,actor){
  return fncLock_(()=>{
    const pct=fncNumber_(req.percent);
    if(pct<0||pct>100)throw new Error('Desconto entre 0 e 100%');
    const rows=fncData_(FNC.sheet.settings), record=rows.find(r=>r.Chave==='DESCONTO_FIXOS');
    if(!record)throw new Error('Configuração do desconto inexistente');
    const old=record.Valor;
    const effective=fncRequired_(req.effectiveFrom,'Data de vigência');
    const oldDiscounts=fncData_(FNC.sheet.discounts);
    // Vigência global: nunca alterar/apagar descontos que sustentam apuramentos fechados.
    if(oldDiscounts.length){
      const last=oldDiscounts[oldDiscounts.length-1];
      if(last['Fim vigência']==='')fncUpdate_(FNC.sheet.discounts,last.row,{'Fim vigência':effective});
    }
    fncAppend_(FNC.sheet.discounts,{'ID':'DESC-'+Utilities.getUuid(),
      'Início vigência':effective,'Desconto %':pct,'Motivo':String(req.reason||'Alteração administrativa'),
      'Definido por':actor.email,'Registado em':fncNow_(),'Notas':'Aplicar apenas a ciclos futuros ainda não fechados'});
    fncUpdate_(FNC.sheet.settings,record.row,{'Valor':pct,'Início vigência':effective,
      'Atualizado em':fncNow_(),'Atualizado por':actor.email});
    fncAudit_('CONFIG_WALLBOX','Desconto',old,pct,actor);
    return {ok:true,discount:pct};
  });
}
function fncUpsertFraction_(req,actor){
  return fncLock_(()=>{
    const id=fncRequired_(req.fractionId,'Fração').toUpperCase().replace(/[^0-9A-Z]/g,'');
    const label=fncRequired_(req.label,'Descrição');
    const email=fncRequired_(req.email,'Email autorizado').toLowerCase();
    if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email))throw new Error('Email inválido');
    const initial=fncNumber_(req.initialReading);
    if(initial<0)throw new Error('Leitura inicial inválida');
    const start=fncRequired_(req.initialDate,'Data da leitura inicial');
    const fraction=fncData_(FNC.sheet.fractions).find(r=>String(r['ID fração']).toUpperCase()===id);
    const data={'ID fração':id,'Fração':label,'Estado':'Ativa',
      'Email autorizado':email,'Data adesão':start,'Leitura inicial (kWh)':initial,
      'Data leitura inicial':start,'Foto inicial Drive ID':req.initialPhotoId||'',
      'Atualizado em':fncNow_(),'Observações':String(req.note||'')};
    if(fraction){
      if(String(fraction['Estado'])==='Ativa')throw new Error('Fração já ativa — desative primeiro');
      fncUpdate_(FNC.sheet.fractions,fraction.row,data);
    } else fncAppend_(FNC.sheet.fractions,data);
    fncAudit_(id,'Ativar wallbox',fraction?.Estado||'',data,actor);
    return {ok:true};
  });
}
function fncDeactivateFraction_(req,actor){
  return fncLock_(()=>{
    const id=fncRequired_(req.fractionId,'Fração');
    const fraction=fncData_(FNC.sheet.fractions).find(r=>String(r['ID fração'])===id);
    if(!fraction || String(fraction['Estado'])!=='Ativa')throw new Error('Fração não está ativa');
    const date=fncRequired_(req.endDate,'Data fim de utilização');
    const final=fncNumber_(req.finalReading);
    if(final<Number(fraction['Leitura inicial (kWh)']))throw new Error('Leitura final inferior à inicial');
    fncUpdate_(FNC.sheet.fractions,fraction.row,{'Estado':'Inativa','Data desativação':date,
      'Leitura final (kWh)':final,'Foto final Drive ID':req.finalPhotoId||'', 'Atualizado em':fncNow_()});
    fncAudit_(id,'Desativar wallbox','Ativa','Inativa',actor,'Histórico preservado; regularizar valores pendentes.');
    return {ok:true};
  });
}
function fncValidateInvoice_(req,actor){
  return fncLock_(()=>{
    const id=fncRequired_(req.id,'Fatura');
    const invoice=fncData_(FNC.sheet.invoices).find(r=>String(r.ID)===id);
    if(!invoice)throw new Error('Fatura não encontrada');
    if(String(invoice['Tipo']).toLowerCase().includes('orçamento'))throw new Error('Orçamento não é fatura');
    if(String(invoice['Validação'])==='Validada')return {ok:true,alreadyValidated:true};
    fncUpdate_(FNC.sheet.invoices,invoice.row,{'Validação':'Validada','Validado por':actor.email});
    fncAudit_(id,'Validar fatura',invoice['Validação'],'Validada',actor);
    if(String(invoice['Fornecedor']).toLowerCase().includes('edp')){
      // Ponto de entrada do ciclo. O email é intencionalmente separado:
      // só pode ser enviado após tarifário e período terem sido conferidos.
      return {ok:true,edpCycleReady:true,message:'Fatura validada; verificar tarifas e ativar notificações após publicação do backend.'};
    }
    return {ok:true};
  });
}
function fncMarkInvoicePaid_(req,actor){
  return fncLock_(()=>{
    const id=fncRequired_(req.id,'Fatura'), date=fncRequired_(req.paidAt,'Data pagamento');
    const row=fncData_(FNC.sheet.invoices).find(r=>String(r.ID)===id);
    if(!row)throw new Error('Fatura não encontrada');
    if(String(row['Validação'])!=='Validada')throw new Error('Validar fatura primeiro');
    if(String(row['Pagamento'])==='Pago')return {ok:true,alreadyPaid:true};
    if(String(row['Tipo']).toLowerCase().includes('orçamento'))throw new Error('Orçamento não pagável');
    fncUpdate_(FNC.sheet.invoices,row.row,{'Pagamento':'Pago','Data pag.':date});
    fncAudit_(id,'Registar pagamento',row['Pagamento'],'Pago',actor,'Confirmado pelo administrador. Não executa transferência.');
    return {ok:true};
  });
}
function fncResidentFraction_(actor){
  const id=String(actor.fractionId||'').toUpperCase().replace(/[^0-9A-Z]/g,'');
  if(!id)throw new Error('Fração não associada à sessão');
  const row=fncData_(FNC.sheet.fractions).find(r=>String(r['ID fração'])===id);
  if(!row||row['Estado']!=='Ativa')throw new Error('Serviço de wallbox não ativo para a fração');
  if(String(row['Email autorizado']).toLowerCase()!==String(actor.email).toLowerCase())
    throw new Error('Morador não autorizado nesta fração');
  return row;
}
function fncResidentStatus_(actor){
  const fraction=fncResidentFraction_(actor);
  const readings=fncData_(FNC.sheet.readings).filter(r=>r['Fração']===fraction['ID fração']);
  const payments=fncData_(FNC.sheet.payments).filter(r=>r['Fração']===fraction['ID fração']);
  return {ok:true,fraction:fraction['Fração'],
    readings:readings.map(r=>({id:r['ID leitura'],date:r['Data leitura'],kwh:r['Leitura acumulada kWh'],state:r['Estado']})),
    payments:payments.map(r=>({cycle:r['Ciclo EDP'],total:r['Total a pagar €'],state:r['Estado']}))};
}
function fncSubmitReading_(req,actor){
  return fncLock_(()=>{
    const fraction=fncResidentFraction_(actor);
    const kwh=fncNumber_(req.kwh), date=fncRequired_(req.readingDate,'Data leitura');
    if(kwh<0)throw new Error('Leitura inválida');
    const cycle=fncRequired_(req.cycleId,'Ciclo');
    const all=fncData_(FNC.sheet.readings).filter(r=>r['Fração']===fraction['ID fração']);
    if(all.some(r=>r['Ciclo EDP']===cycle && r['Estado']!=='Rejeitada'))throw new Error('Já existe leitura neste ciclo');
    const last=all.filter(r=>r['Estado']==='Validada').slice(-1)[0];
    const base=last?Number(last['Leitura acumulada kWh']):Number(fraction['Leitura inicial (kWh)']);
    if(kwh<base)throw new Error('Leitura inferior à última validada');
    // O endpoint existente deve aceitar imagens apenas após auth. Nunca guardar base64 no Sheet.
    const photoId=fncStorePhoto_(req.photoBase64,actor,cycle);
    const rid='LR-'+Utilities.getUuid();
    fncAppend_(FNC.sheet.readings,{'ID leitura':rid,'Fração':fraction['ID fração'],
      'Ciclo EDP':cycle,'Data leitura':date,'Leitura acumulada kWh':kwh,
      'Foto Drive ID':photoId,'Estado':'Submetida','Submetido por':actor.email,'Submetido em':fncNow_()});
    return {ok:true,readingId:rid,message:'Leitura submetida para validação'};
  });
}
function fncStorePhoto_(b64,actor,cycle){
  const folderId=PropertiesService.getScriptProperties().getProperty('FINANCE_PHOTOS_FOLDER_ID');
  if(!folderId)throw new Error('Pasta privada das fotografias por configurar');
  const data=fncRequired_(b64,'Fotografia').replace(/^data:image\/(?:jpeg|png);base64,/,'');
  if(data.length>4200000 || !/^[A-Za-z0-9+/=]+$/.test(data))throw new Error('Imagem inválida ou demasiado grande');
  const bytes=Utilities.base64Decode(data);
  return DriveApp.getFolderById(folderId)
    .createFile(Utilities.newBlob(bytes,'image/jpeg','wallbox-'+cycle+'-'+Utilities.getUuid()+'.jpg')).getId();
}
function fncValidateReading_(req,actor){
  return fncLock_(()=>{
    const rid=fncRequired_(req.id,'Leitura');
    const r=fncData_(FNC.sheet.readings).find(x=>x['ID leitura']===rid);
    if(!r||r['Estado']!=='Submetida')throw new Error('Leitura não está pendente');
    // Nesta fase o cálculo monetário está bloqueado até integração e teste
    // da fórmula do simulador por ciclo e validação dos descontos em vigor.
    fncUpdate_(FNC.sheet.readings,r.row,{'Estado':'Validada','Validado por':actor.email,'Validado em':fncNow_()});
    fncAudit_(rid,'Validar leitura','Submetida','Validada',actor);
    return {ok:true,calculationPending:true};
  });
}
function fncMarkWallboxPaid_(req,actor){
  return fncLock_(()=>{
    const id=fncRequired_(req.id,'Apuramento');
    const row=fncData_(FNC.sheet.payments).find(r=>r['ID apuramento']===id);
    if(!row || !['A pagar','Pago'].includes(row.Estado))throw new Error('Apuramento não está a pagamento');
    if(row.Estado==='Pago')return {ok:true,alreadyPaid:true};
    fncUpdate_(FNC.sheet.payments,row.row,{'Estado':'Pago','Pago em':fncRequired_(req.paidAt,'Data recebimento'),'Recebido por':actor.email});
    fncAudit_(id,'Recebimento wallbox','A pagar','Pago',actor);
    return {ok:true};
  });
}
/**
 * Cálculo isolado e testável do simulador original.
 * O total só pode ser lançado depois de ler tarifas da fatura e desconto aprovado.
 * inputs: {energyExVat,energyVat,iecExVat,iecVat,powerExVat,powerVat,
 * dgegExVat,dgegVat,avExVat,avVat,invoiceStart,invoiceEnd,buildingKwh,
 * readingStart,readingEnd,readingKwhStart,readingKwhEnd,discountPct}
 */
function financeCalculateWallbox_(i){
  function date_(str){const d=new Date(String(str).slice(0,10)+'T12:00:00Z');if(isNaN(d.getTime()))throw new Error('Data inválida');return d;}
  const p=x=>fncNumber_(i[x]), first=date_(i.invoiceStart), last=date_(i.invoiceEnd);
  const r0=date_(i.readingStart),r1=date_(i.readingEnd);
  const invoiceDays=Math.round((last-first)/86400000)+1,days=Math.round((r1-r0)/86400000);
  const kwh=p('readingKwhEnd')-p('readingKwhStart');
  if(invoiceDays<=0||days<=0||kwh<0||p('buildingKwh')<=0)throw new Error('Períodos ou consumos inválidos');
  const dim=d=>new Date(Date.UTC(d.getUTCFullYear(),d.getUTCMonth()+1,0)).getUTCDate();
  const months=12*(last.getUTCFullYear()-first.getUTCFullYear())+
    (last.getUTCMonth()-first.getUTCMonth())+
    last.getUTCDate()/dim(last)-(first.getUTCDate()-1)/dim(first);
  const withVat=(a,b)=>p(a)*(1+p(b));
  const fixedDaily=withVat('powerExVat','powerVat')+
    (withVat('dgegExVat','dgegVat')+withVat('avExVat','avVat'))*months/invoiceDays;
  const share=Math.min(1,(kwh/days)/(p('buildingKwh')/invoiceDays));
  const rawFixed=fixedDaily*days*share;
  const discount=p('discountPct');
  if(discount<0||discount>100)throw new Error('Desconto inválido');
  const fixed=rawFixed*(1-discount/100);
  const energy=kwh*(withVat('energyExVat','energyVat')+withVat('iecExVat','iecVat'));
  const total=Math.round((energy+fixed+Number.EPSILON)*100)/100;
  return {invoiceDays,readingDays:days,kwh,share,
    energy:Math.round(energy*100)/100,fixedBeforeDiscount:Math.round(rawFixed*100)/100,
    discountPct:discount,fixedAfterDiscount:Math.round(fixed*100)/100,total};
}
