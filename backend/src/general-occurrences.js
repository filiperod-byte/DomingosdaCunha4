// Eventos duráveis: uma linha contém o estado e o recibo do pedido, sob lock.
function createGeneralOccurrenceService_(ports) {
  function invalid(message) { const e=new Error(message); e.code='VALIDATION_ERROR'; throw e; }
  const catalog = DC4_OCCURRENCE_CATALOG;
  const publicStates = ['EM_ANALISE','AGENDADO','RESOLVIDO'];
  const text = (v,max) => String(v == null ? '' : v).trim().slice(0,max);
  const hash = v => ports.crypto.hashPin(JSON.stringify(v),'dc4-general-v1');
  const all = () => ports.general.list().map(r => ({...r, data:JSON.parse(r.DATA_JSON)}));
  function latest(events) {
    const map = new Map(); events.forEach(e => map.set(e.data.id,e.data)); return [...map.values()];
  }
  function publicItem(o) {
    return {id:o.id,scope:o.scope,floor:o.floor,category:o.category,reason:o.reason,
      location:o.publicLocation,description:o.publicDescription,status:o.status,
      reportedAt:o.reportedAt,updatedAt:o.updatedAt,scheduledFor:o.scheduledFor||'',
      resolution:o.publicResolution||'',confirmations:(o.confirmations||[]).length};
  }
  function list() { return {success:true,occurrences:latest(all()).filter(o=>publicStates.includes(o.status)).map(publicItem)}; }
  function adminList() { return {success:true,occurrences:latest(all())}; }
  function mutate(action,p,change) {
    try { return ports.lock.run(()=>{
      const requestId=text(p.requestId,100),actor=p._actor;
      if(!actor || !/^[a-zA-Z0-9_-]{16,80}$/.test(requestId))invalid('Identificador de envio inválido. Atualize a página.');
      const clean={...p};['token','action','_method','requestId','_actor','_actorName','adminEmail'].forEach(k=>delete clean[k]);
      const digest=hash(clean),events=all();
      const old=events.find(e=>e.REQUEST_ID===requestId&&e.ACTOR_ID===actor&&e.ACTION===action);
      if(old){if(old.PAYLOAD_HASH!==digest)invalid('O pedido já foi recebido com outro conteúdo. Consulte as ocorrências antes de criar um novo.');return {success:true,occurrenceId:old.data.id,replayed:true};}
      const state=change(latest(events));
      const at=ports.clock.now().toISOString();state.updatedAt=at;
      state.revision=(state.revision||0)+1;
      ports.general.append({EVENT_ID:ports.crypto.uuid(),REQUEST_ID:requestId,ACTOR_ID:actor,ACTION:action,PAYLOAD_HASH:digest,AT:at,DATA_JSON:JSON.stringify(state)});
      return {success:true,occurrenceId:state.id};
    }); } catch (e) { return {success:false,code:e.code||'RETRY_SAME_REQUEST',message:e.code?e.message:'Não foi possível confirmar a gravação. Repita o mesmo envio.'}; }
  }
  function report(p) {return mutate('report',p,()=>{
    const scope=p.scope,floor=scope==='floor'?Number(p.floor):null;
    if(!['floor','general'].includes(scope)||scope==='floor'&&(p.floor==null||String(p.floor).trim()===''||!catalog.floors.includes(floor)))invalid('Localização inválida.');
    const cat=catalog.categories.find(c=>c.id===p.category&&(c.scope==='both'||c.scope===scope));
    if(!cat||!cat.reasons.includes(p.reason))invalid('Tipo ou motivo inválido.');
    const description=text(p.description,2000),location=text(p.location,160);
    if(!description)invalid('Descreva o problema.');
    if(scope==='general'&&!location)invalid('Indique o equipamento ou local onde observou o problema.');
    const photoData=String(p.photoBase64||'');
    if(photoData.length>1600000||photoData&&!['image/jpeg','image/png','image/webp'].includes(p.photoType))invalid('Fotografia inválida ou demasiado grande.');
    const id='GER-'+ports.crypto.uuid(),at=ports.clock.now().toISOString();
    const photo=ports.files.save({base64:photoData,mimeType:p.photoType,fileName:'ocorrencia-'+id+'.jpg',folderName:'OCORRENCIAS_GERAIS_PRIVADAS',occurrenceId:id,prefix:'geral',private:true});
    return {id,scope,floor,category:cat.id,reason:p.reason,location,description,
      reportedBy:p._actorName,reporterId:p._actor,photoUrl:photo.fileUrl||'',
      status:'RECEBIDO',reportedAt:at,publicLocation:'',publicDescription:'',internalNotes:'',confirmations:[],history:[{at,status:'RECEBIDO'}]};
  });}
  function update(p){return mutate('update',p,items=>{
    const old=items.find(o=>o.id===p.occurrenceId);if(!old)invalid('Ocorrência não encontrada.');
    if(Number(p.revision)!==old.revision)invalid('Esta ocorrência foi alterada. Atualize a lista antes de guardar.');
    if(!Object.prototype.hasOwnProperty.call(catalog.statuses,p.status))invalid('Estado inválido.');
    const description=text(p.publicDescription,2000),resolution=text(p.publicResolution,2000);
    if(publicStates.includes(p.status)&&!description)invalid('Preencha a descrição para os moradores.');
    if(p.status==='RESOLVIDO'&&!resolution)invalid('Indique o que foi feito para resolver.');
    const scheduledFor=text(p.scheduledFor,10);
    if(scheduledFor&&(!/^\d{4}-\d{2}-\d{2}$/.test(scheduledFor)||isNaN(Date.parse(scheduledFor))))invalid('Data inválida.');
    if(p.status==='AGENDADO'&&!scheduledFor)invalid('Indique a data da intervenção.');
    return {...old,status:p.status,publicDescription:description,publicLocation:text(p.publicLocation,160),publicResolution:resolution,scheduledFor,
      internalNotes:text(p.internalNotes,4000),history:[...(old.history||[]),{at:ports.clock.now().toISOString(),status:p.status,by:p._actorName}].slice(-100)};
  });}
  function confirm(p){return mutate('confirm',p,items=>{
    const old=items.find(o=>o.id===p.occurrenceId);if(!old||!['EM_ANALISE','AGENDADO'].includes(old.status))invalid('Esta ocorrência já não está aberta para confirmação.');
    return {...old,confirmations:[...new Set([...(old.confirmations||[]),p._actor])]};
  });}
  return {list,adminList,report,update,confirm};
}
