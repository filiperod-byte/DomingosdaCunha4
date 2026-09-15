/* Apresentação pública partilhada pelo mapa e pelos QR Codes. */
window.dc4RenderOccurrence = function(box, o){
 box.replaceChildren();
 const title=document.createElement('strong');title.textContent='Anomalia já registada';box.appendChild(title);
 const reason=document.createElement('div');
 reason.textContent=o.reason ? 'Motivo: '+o.reason : 'O motivo ainda não está disponível nesta consulta.';
 box.appendChild(reason);
 const description=typeof o.description==='string'?o.description.trim():'';
 if(description){
  const detail=document.createElement('div');detail.style.whiteSpace='pre-wrap';detail.style.overflowWrap='anywhere';
  detail.textContent='Descrição: '+description;box.appendChild(detail);
 }else if(o.reason==='Outro'||o.reason==='Outra anomalia'){
  const missing=document.createElement('div');missing.textContent=typeof o.description==='string'?'Esta ocorrência não tem uma descrição registada.':'A descrição ainda não está disponível nesta consulta.';box.appendChild(missing);
 }
 if(o.reportedAt){
  const date=new Date(o.reportedAt);
  if(!isNaN(date.getTime())){const line=document.createElement('div');line.textContent='Registada em: '+new Intl.DateTimeFormat('pt-PT',{dateStyle:'short',timeStyle:'short',timeZone:'Europe/Lisbon'}).format(date);box.appendChild(line);}
 }
 const hint=document.createElement('div');
 hint.textContent=o.reason
  ? 'Se é a mesma situação, não precisa de a reportar novamente. Use o formulário apenas para uma anomalia diferente ou informação adicional.'
  : 'Já existe uma ocorrência aberta neste ponto. Não é possível comparar o motivo enquanto o serviço não disponibilizar esse detalhe.';
 box.appendChild(hint);
};
