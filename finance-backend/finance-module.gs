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
    cycles:'WALLBOX_CICLOS', rates:'WALLBOX_TARIFAS_CICLO', notifications:'WALLBOX_NOTIFICACOES'
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
    'finance.admin.notifyCycle': () => fncNotifyCycle_(req,actor),
    'finance.admin.setNotifications': () => fncSetNotifications_(req,actor),
    'finance.admin.upsertFraction': () => fncUpsertFraction_(req, actor),
    'finance.admin.deactivateFraction': () => fncDeactivateFraction_(req, actor),
    'finance.admin.validateInvoice': () => {
      const result=fncValidateInvoice_(req,actor);
      if(result.cycleId && String(fncSettings_().NOTIFICACOES_ATIVAS).toUpperCase()==='SIM')
        return {...result,notifications:fncNotifyCycle_({cycleId:result.cycleId},actor)};
      return result;
    },
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
      documentType:r['Tipo'],iban:r['IBAN fornecedor'],entity:r['Entidade MB'],
      reference:r['Referência MB'],url:r['PDF original']
    })),
    readings:fncData_(FNC.sheet.readings).map(r=>({
      id:r['ID leitura'],fractionId:r['Fração'],cycle:r['Ciclo EDP'],date:r['Data leitura'],
      kwh:r['Leitura acumulada kWh'],state:r['Estado'],row:r.row
    })),
    cycles:fncData_(FNC.sheet.cycles).map(r=>({
      id:r['ID ciclo'],invoiceId:r['Fatura ID'],state:r['Estado'],notifiedAt:r['Notificado em']
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
    const fraction=fncData_(FNC.sheet.fractions).filter(r=>String(r['ID fração']).toUpperCase()===id).slice(-1)[0];
    const data={'ID fração':id,'Fração':label,'Estado':'Ativa',
      'Email autorizado':email,'Data adesão':start,'Leitura inicial (kWh)':initial,
      'Data leitura inicial':start,'Foto inicial Drive ID':fncStorePhoto_(req.initialPhotoBase64,actor,'inicio-'+id),
      'Atualizado em':fncNow_(),'Observações':String(req.note||'')};
    if(fraction && String(fraction['Estado'])==='Ativa')throw new Error('Fração já ativa — desative primeiro');
    // Cada nova adesão cria uma linha própria: conserva a base/histórico da adesão anterior.
    fncAppend_(FNC.sheet.fractions,data);
    fncAudit_(id,'Ativar wallbox',fraction?.Estado||'',data,actor);
    return {ok:true};
  });
}
function fncDeactivateFraction_(req,actor){
  return fncLock_(()=>{
    const id=fncRequired_(req.fractionId,'Fração');
    const fraction=fncData_(FNC.sheet.fractions).find(r=>String(r['ID fração'])===id && r['Estado']==='Ativa');
    if(!fraction)throw new Error('Fração não está ativa');
    const date=fncRequired_(req.endDate,'Data fim de utilização');
    const final=fncNumber_(req.finalReading);
    const last=fncData_(FNC.sheet.readings).filter(r=>r['Fração']===id&&r['Estado']==='Validada')
      .filter(r=>String(r['Data leitura'])>=String(fraction['Data adesão'])).slice(-1)[0];
    const minReading=last?Number(last['Leitura acumulada kWh']):Number(fraction['Leitura inicial (kWh)']);
    if(final<minReading)throw new Error('Leitura final inferior à última leitura válida');
    const finalPhotoId=fncStorePhoto_(req.finalPhotoBase64,actor,'fim-'+id);
    fncUpdate_(FNC.sheet.fractions,fraction.row,{'Estado':'Inativa','Data desativação':date,
      'Leitura final (kWh)':final,'Foto final Drive ID':finalPhotoId, 'Atualizado em':fncNow_()});
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
    const isEdp=String(invoice['Fornecedor']).toLowerCase().includes('edp');
    const cycle=isEdp?fncOpenCycle_(invoice,actor):null;
    fncUpdate_(FNC.sheet.invoices,invoice.row,{'Validação':'Validada','Validado por':actor.email});
    fncAudit_(id,'Validar fatura',invoice['Validação'],'Validada',actor);
    if(isEdp)return {ok:true,edpCycleReady:true,cycleId:cycle,
      message:'Ciclo aberto. Notificação dos moradores depende de configuração e confirmação administrativa.'};
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
  const row=fncData_(FNC.sheet.fractions).find(r=>String(r['ID fração'])===id&&r['Estado']==='Ativa');
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
    initialReading:fraction['Leitura inicial (kWh)'],initialDate:fraction['Data leitura inicial'],
    openCycle:(fncData_(FNC.sheet.cycles).filter(c=>c['Estado']==='Aberto').slice(-1)[0]||{})['ID ciclo']||null,
    readings:readings.map(r=>({id:r['ID leitura'],cycle:r['Ciclo EDP'],date:r['Data leitura'],kwh:r['Leitura acumulada kWh'],state:r['Estado']})),
    payments:payments.map(r=>({cycle:r['Ciclo EDP'],total:r['Total a pagar €'],state:r['Estado']}))};
}
function fncSubmitReading_(req,actor){
  return fncLock_(()=>{
    const fraction=fncResidentFraction_(actor);
    const kwh=fncNumber_(req.kwh), date=fncRequired_(req.readingDate,'Data leitura');
    if(kwh<0)throw new Error('Leitura inválida');
    const cycle=fncRequired_(req.cycleId,'Ciclo');
    const activeCycle=fncData_(FNC.sheet.cycles).find(c=>c['ID ciclo']===cycle && c['Estado']==='Aberto');
    if(!activeCycle)throw new Error('Ciclo não autorizado ou já encerrado');
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
  if(data.length>4300000 || !/^[A-Za-z0-9+/=]+$/.test(data))throw new Error('Imagem inválida ou demasiado grande');
  const isPng=/^data:image\/png;base64,/.test(String(b64));
  const bytes=Utilities.base64Decode(data), mime=isPng?'image/png':'image/jpeg';
  const ext=isPng?'.png':'.jpg';
  return DriveApp.getFolderById(folderId)
    .createFile(Utilities.newBlob(bytes,mime,'wallbox-'+cycle+'-'+Utilities.getUuid()+ext)).getId();
}
function fncValidateReading_(req,actor){
  return fncLock_(()=>{
    const rid=fncRequired_(req.id,'Leitura');
    const r=fncData_(FNC.sheet.readings).find(x=>x['ID leitura']===rid);
    if(!r||r['Estado']!=='Submetida')throw new Error('Leitura não está pendente');
    const cycle=fncData_(FNC.sheet.cycles).find(x=>x['ID ciclo']===r['Ciclo EDP']);
    const tariffs=fncData_(FNC.sheet.rates).find(x=>x['ID ciclo']===r['Ciclo EDP']);
    if(!cycle||!tariffs)throw new Error('Ciclo e tarifas ainda não configurados');
    const settings=fncSettings_(), pct=settings.DESCONTO_FIXOS;
    if(pct===''||pct===null||pct===undefined)throw new Error('Definir desconto antes de validar a leitura');
    const fraction=fncData_(FNC.sheet.fractions).find(x=>x['ID fração']===r['Fração']&&x['Estado']==='Ativa');
    if(!fraction)throw new Error('Fração não encontrada');
    const prev=fncData_(FNC.sheet.readings).filter(x=>x['Fração']===r['Fração']&&x['Estado']==='Validada')
      .filter(x=>String(x['Data leitura'])>=String(fraction['Data adesão']))
      .sort((a,b)=>String(a['Data leitura']).localeCompare(String(b['Data leitura']))).slice(-1)[0];
    const fromDate=prev?prev['Data leitura']:fraction['Data leitura inicial'];
    const fromKwh=prev?prev['Leitura acumulada kWh']:fraction['Leitura inicial (kWh)'];
    if(fromDate===''||fromKwh==='')throw new Error('Leitura inicial não configurada');
    const inputs={
      invoiceStart:fncIsoDate_(tariffs['Início faturação']),
      invoiceEnd:fncIsoDate_(tariffs['Fim faturação']),
      buildingKwh:tariffs['Consumo prédio kWh'],
      energyExVat:tariffs['Energia sem IVA'],energyVat:tariffs['IVA energia'],
      iecExVat:tariffs['IEC sem IVA'],iecVat:tariffs['IVA IEC'],
      powerExVat:tariffs['Potência sem IVA'],powerVat:tariffs['IVA potência'],
      dgegExVat:tariffs['DGEG sem IVA'],dgegVat:tariffs['IVA DGEG'],
      avExVat:tariffs['Audiovisual sem IVA'],avVat:tariffs['IVA audiovisual'],
      readingStart:fncIsoDate_(fromDate),readingEnd:fncIsoDate_(r['Data leitura']),
      readingKwhStart:fromKwh,readingKwhEnd:r['Leitura acumulada kWh'],discountPct:pct
    };
    const calc=financeCalculateWallbox_(inputs);
    const paid=fncData_(FNC.sheet.payments).find(x=>x['Fração']===r['Fração']&&x['Ciclo EDP']===r['Ciclo EDP']);
    if(paid)throw new Error('Já existe apuramento para este ciclo e fração');
    const pid='WP-'+Utilities.getUuid();
    fncAppend_(FNC.sheet.payments,{'ID apuramento':pid,'Fração':r['Fração'],
      'Ciclo EDP':r['Ciclo EDP'],'Data inicial':inputs.readingStart,'Data final':inputs.readingEnd,
      'Leitura inicial kWh':fromKwh,'Leitura final kWh':r['Leitura acumulada kWh'],
      'Consumo kWh':calc.kwh,'Energia + IEC €':calc.energy,'Custos fixos antes desc. €':calc.fixedBeforeDiscount,
      'Desconto %':pct,'Parcela custos fixos €':calc.fixedAfterDiscount,'Total a pagar €':calc.total,
      'Estado':'A pagar','Emissão':fncNow_(),'Validado por':actor.email,
      'Foto leitura Drive ID':r['Foto Drive ID']});
    fncUpdate_(FNC.sheet.readings,r.row,{'Estado':'Validada','Validado por':actor.email,'Validado em':fncNow_()});
    fncAudit_(rid,'Validar leitura e apurar valor','Submetida','Validada',actor,'Apuramento '+pid);
    return {ok:true,paymentId:pid,amount:calc.total};
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

/** Fatura EDP aprovada abre um ciclo; tarifas são copiadas e imutáveis nesse ciclo. */
function fncIsoDate_(value){
  if(value instanceof Date)return Utilities.formatDate(value,'Europe/Lisbon','yyyy-MM-dd');
  const s=String(value||'').trim();if(!/^\d{4}-\d{2}-\d{2}/.test(s))throw Error('Data inválida: '+s);
  return s.slice(0,10);
}
function fncTariffMap_(){
  return Object.fromEntries(fncData_(FNC.sheet.tariff).map(x=>[String(x['Parâmetro']),x['Valor']]));
}
function fncOpenCycle_(invoice,actor){
  const id='CYCLE-'+invoice.ID, existing=fncData_(FNC.sheet.cycles).find(x=>x['ID ciclo']===id);
  if(existing)return id;
  const t=fncTariffMap_();
  const invoiceStart=fncIsoDate_(invoice['Período início']),invoiceEnd=fncIsoDate_(invoice['Período fim']);
  if(fncIsoDate_(t['Período início'])!==invoiceStart||fncIsoDate_(t['Período fim'])!==invoiceEnd)
    throw Error('Período da tabela EDP não corresponde ao documento. Rever tarifas primeiro.');
  const keys=['Consumo prédio','Energia (sem IVA)','IVA energia','IEC (sem IVA)','IVA IEC',
    'Potência (sem IVA)','IVA potência','DGEG (sem IVA)','IVA DGEG',
    'Contribuição audiovisual (sem IVA)','IVA audiovisual'];
  for(const k of keys)fncNumber_(t[k]);
  const fields={'ID ciclo':id,'Fatura ID':invoice.ID,'Início faturação':invoiceStart,
    'Fim faturação':invoiceEnd,'Consumo prédio kWh':t['Consumo prédio'],
    'Energia sem IVA':t['Energia (sem IVA)'],'IVA energia':t['IVA energia'],
    'IEC sem IVA':t['IEC (sem IVA)'],'IVA IEC':t['IVA IEC'],
    'Potência sem IVA':t['Potência (sem IVA)'],'IVA potência':t['IVA potência'],
    'DGEG sem IVA':t['DGEG (sem IVA)'],'IVA DGEG':t['IVA DGEG'],
    'Audiovisual sem IVA':t['Contribuição audiovisual (sem IVA)'],
    'IVA audiovisual':t['IVA audiovisual'],
    'Desconto %':fncSettings_().DESCONTO_FIXOS,'Criado em':fncNow_,
    'Validado por':actor.email,'Origem EDP':invoice['PDF original']};
  fields['Criado em']=fncNow_();
  fncAppend_(FNC.sheet.rates,fields);
  fncAppend_(FNC.sheet.cycles,{'ID ciclo':id,'Fatura ID':invoice.ID,
    'Início faturação':invoiceStart,'Fim faturação':invoiceEnd,'Estado':'Aberto',
    'Validado por':actor.email,'Criado em':fncNow_(),
    'Desconto aplicado %':fncSettings_().DESCONTO_FIXOS,
    'Observações':'Aguardam-se leituras das frações ativas.'});
  return id;
}
function fncNotifyCycle_(req,actor){
  const settings=fncSettings_();
  if(String(settings.NOTIFICACOES_ATIVAS).toUpperCase()!=='SIM')
    throw Error('Notificações ainda desativadas na configuração.');
  const id=fncRequired_(req.cycleId,'Ciclo');
  return fncLock_(()=>{
    const c=fncData_(FNC.sheet.cycles).find(r=>r['ID ciclo']===id);
    if(!c||c['Estado']!=='Aberto')throw Error('Ciclo não está aberto');
    const active=fncData_(FNC.sheet.fractions).filter(f=>f['Estado']==='Ativa'&&f['Email autorizado']);
    const sent=fncData_(FNC.sheet.notifications).filter(n=>n['Ciclo']===id&&n['Estado']==='Enviado');
    const sentFractions=new Set(sent.map(n=>String(n['Fração'])));
    const link='https://filiperod-byte.github.io/DomingosdaCunha4/V2/wallbox.html';
    let count=0;
    for(const f of active){
      const email=String(f['Email autorizado']).trim(),fractionId=String(f['ID fração']);
      if(sentFractions.has(fractionId))continue;
      MailApp.sendEmail({to:email,subject:'Domingos da Cunha 4 — Nova leitura da wallbox',
        body:'A fatura de eletricidade do condomínio já foi validada.\n'+
          'Pode agora submeter a leitura acumulada da sua wallbox, com fotografia, na App:\n'+link+
          '\n\nO valor a transferir será calculado após validação da leitura pela administração.'});
      fncAppend_(FNC.sheet.notifications,{'ID':'WN-'+Utilities.getUuid(),'Ciclo':id,
        'Fração':fractionId,'Email':email,'Estado':'Enviado','Enviado em':fncNow_()});
      count++;
    }
    // Sem destinatários não se considera o ciclo notificado.
    if(active.length && fncData_(FNC.sheet.notifications)
      .filter(n=>n['Ciclo']===id&&n['Estado']==='Enviado').length>=active.length){
      fncUpdate_(FNC.sheet.cycles,c.row,{'Notificado em':fncNow_()});
    }
    if(count)fncAudit_(id,'Enviar pedidos de leitura','',count+' email(s)',actor);
    return {ok:true,notified:count,eligible:active.length};
  });
}

/** Ativação/desativação das notificações no backend, auditada. */
function fncSetNotifications_(req,actor){
  return fncLock_(()=>{
    const enabled=req.enabled===true || req.enabled==='true';
    const record=fncData_(FNC.sheet.settings).find(r=>r['Chave']==='NOTIFICACOES_ATIVAS');
    if(!record)throw Error('Configuração de notificações não encontrada');
    const old=record['Valor'],v=enabled?'SIM':'NAO';
    fncUpdate_(FNC.sheet.settings,record.row,{'Valor':v,'Atualizado em':fncNow_(),'Atualizado por':actor.email});
    fncAudit_('CONFIG_WALLBOX','Notificações',old,v,actor);
    return {ok:true,enabled};
  });
}

/**
 * Descoberta automática de novos documentos na Drive /FATURAS/AAAA/MM.
 * Não tenta classificar imagens/PDF digitalizados como faturas fiáveis.
 * Apenas regista novos documentos para validação humana/extração posterior.
 * Propriedade FINANCE_FATURAS_ROOT_ID = ID da pasta FATURAS privada.
 */
function financeScanFolder_(){
  return fncLock_(()=>{
    const root=PropertiesService.getScriptProperties().getProperty('FINANCE_FATURAS_ROOT_ID');
    if(!root)throw Error('FINANCE_FATURAS_ROOT_ID por configurar');
    const seen=new Set(fncData_(FNC.sheet.invoices).map(r=>String(r['ID Drive']||'')));
    let count=0;
    const years=DriveApp.getFolderById(root).getFolders();
    while(years.hasNext()){
      const year=years.next();
      if(!/^\d{4}$/.test(year.getName()))continue;
      const months=year.getFolders();
      while(months.hasNext()){
        const month=months.next();
        if(!/^(0[1-9]|1[0-2])$/.test(month.getName()))continue;
        const files=month.getFiles();
        while(files.hasNext()){
          const file=files.next(),id=file.getId();
          if(seen.has(id))continue;
          const mimetype=file.getMimeType();
          if(!['application/pdf','image/jpeg','image/png'].includes(mimetype))continue;
          const docid='DC4-DOC-'+Utilities.getUuid();
          fncAppend_(FNC.sheet.invoices,{
            'ID':docid,'Tipo':'Documento por classificar',
            'Fornecedor':'Por identificar','Descrição':file.getName(),
            'Documento':file.getName(),'Validação':'Por validar','Pagamento':'Por confirmar',
            'PDF original':'https://drive.google.com/file/d/'+id+'/view',
            'Observações':'Importado automaticamente; verificar antes de registar pagamento.',
            'ID Drive':id
          });
          seen.add(id);count++;
        }
      }
    }
    return {ok:true,newDocuments:count};
  });
}
/** Criar o trigger uma única vez, após autorizar o projeto Apps Script. */
function financeInstallScannerTrigger_(){
  const matches=ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='financeScanFolder_');
  if(matches.length)return {ok:true,alreadyInstalled:true};
  ScriptApp.newTrigger('financeScanFolder_').timeBased().everyMinutes(15).create();
  return {ok:true,installed:true};
}
