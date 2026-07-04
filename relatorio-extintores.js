let CFG=null;
let ALL=[];
let FILTERED=[];
const $=id=>document.getElementById(id);

document.addEventListener('DOMContentLoaded',init);

async function init(){
  bind();
  try{
    CFG=await loadConfig();
    const email=sessionStorage.getItem('dc4_admin_email') || CFG.adminEmail || '';
    if($('reportEmail')) $('reportEmail').value=email;
    await loadData();
  }catch(e){
    console.error(e);
    $('reportList').innerHTML='<div class="empty">Não foi possível carregar o relatório.<br>'+esc(e.message||'Erro')+'</div>';
  }
}

function bind(){
  ['filterStatus','filterReason','filterFloor','filterSearch'].forEach(id=>$(id)?.addEventListener('input',applyFilters));
  $('clearFiltersBtn')?.addEventListener('click',clearFilters);
  $('refreshBtn')?.addEventListener('click',loadData);
  $('pdfBtn')?.addEventListener('click',()=>generateReport({email:false}));
  $('emailBtn')?.addEventListener('click',()=>generateReport({email:true}));
}

async function loadConfig(){
  const r=await fetch('config.json',{cache:'no-store'});
  if(!r.ok) throw new Error('config.json não encontrado.');
  return r.json();
}

async function loadData(){
  $('reportList').innerHTML='<div class="loading"><span class="spinner"></span><br>A carregar ocorrências...</div>';
  const [openRes,pendingRes]=await Promise.allSettled([apiGet('openOccurrences'),apiGet('pendingOccurrences')]);
  const open=openRes.status==='fulfilled'?asList(openRes.value).map(o=>norm(o,'open')):[];
  const pending=pendingRes.status==='fulfilled'?asList(pendingRes.value).map(o=>norm(o,'pending')):[];
  ALL=[...pending,...open].sort((a,b)=>String(b.createdAt||'').localeCompare(String(a.createdAt||'')));
  populateFilters();
  applyFilters();
}

async function apiGet(action){
  const u=new URL(CFG.backendUrl);
  u.searchParams.set('action',action);
  const r=await fetch(u,{cache:'no-store'});
  const t=await r.text();
  try{return JSON.parse(t)}catch(e){throw new Error('Resposta inválida do backend.');}
}

async function apiPost(data){
  const r=await fetch(CFG.backendUrl,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify(data)});
  const t=await r.text();
  try{return JSON.parse(t)}catch(e){throw new Error('Resposta inválida do backend.');}
}

function asList(p){
  if(!p) return [];
  if(Array.isArray(p)) return p;
  return p.occurrences || p.items || p.data || [];
}

function norm(o,status){
  const floor=o.floor??o.piso??o.FLOOR??'';
  const point=o.point??o.ponto??o.POINT??'';
  return {
    id:String(o.id||o.occurrenceId||o.OCCURRENCE_ID||`${floor}:${point}`),
    status,
    floor:Number(floor),
    floorLabel:o.floorLabel||o.pisoLabel||formatFloor(floor),
    point:String(point||''),
    location:o.location||o.localizacao||o.LOCATION||'',
    reportedBy:o.reportedBy||o.name||o.nome||o.REPORTED_BY||'',
    reason:o.reason||o.motivo||o.REASON||'Sem tipo indicado',
    notes:o.notes||o.observacao||o.NOTES||'',
    createdAt:o.createdAt||o.timestamp||o.REPORTED_AT||o.data||'',
    photoUrl:o.photoUrl||o.fotoUrl||o.PHOTO_FILE_URL||''
  };
}

function populateFilters(){
  const currentReason=$('filterReason').value;
  const currentFloor=$('filterFloor').value;
  const reasons=[...new Set(ALL.map(o=>o.reason||'Sem tipo indicado'))].sort((a,b)=>a.localeCompare(b,'pt'));
  $('filterReason').innerHTML='<option value="all">Todos os tipos</option>'+reasons.map(r=>`<option value="${attr(r)}">${esc(r)}</option>`).join('');
  if([...$('filterReason').options].some(o=>o.value===currentReason)) $('filterReason').value=currentReason;
  const floors=[...new Set(ALL.map(o=>o.floor).filter(v=>!Number.isNaN(v)))].sort((a,b)=>b-a);
  $('filterFloor').innerHTML='<option value="all">Todos os pisos</option>'+floors.map(f=>`<option value="${f}">${esc(formatFloor(f))}</option>`).join('');
  if([...$('filterFloor').options].some(o=>o.value===currentFloor)) $('filterFloor').value=currentFloor;
}

function applyFilters(){
  const st=$('filterStatus').value;
  const reason=$('filterReason').value;
  const floor=$('filterFloor').value;
  const q=String($('filterSearch').value||'').trim().toLowerCase();
  FILTERED=ALL.filter(o=>{
    if(st!=='all'&&o.status!==st) return false;
    if(reason!=='all'&&o.reason!==reason) return false;
    if(floor!=='all'&&String(o.floor)!==String(floor)) return false;
    if(q){
      const hay=[o.floorLabel,o.point,o.location,o.reportedBy,o.reason,o.notes,o.id].join(' ').toLowerCase();
      if(!hay.includes(q)) return false;
    }
    return true;
  });
  render();
}

function clearFilters(){
  $('filterStatus').value='all';
  $('filterReason').value='all';
  $('filterFloor').value='all';
  $('filterSearch').value='';
  applyFilters();
}

function render(){
  const open=ALL.filter(o=>o.status==='open').length;
  const pending=ALL.filter(o=>o.status==='pending').length;
  const types=new Set(ALL.map(o=>o.reason||'Sem tipo indicado')).size;
  $('kpiOpen').textContent=open;
  $('kpiPending').textContent=pending;
  $('kpiTypes').textContent=types;
  $('kpiShown').textContent=FILTERED.length;
  if(!FILTERED.length){
    $('reportList').innerHTML='<div class="empty">Sem resultados para os filtros selecionados.</div>';
    return;
  }
  const groups=groupBy(FILTERED,o=>o.reason||'Sem tipo indicado');
  const html=Object.keys(groups).sort((a,b)=>groups[b].length-groups[a].length||a.localeCompare(b,'pt')).map(reason=>groupHtml(reason,groups[reason])).join('');
  $('reportList').innerHTML=html;
}

function groupHtml(reason,items){
  return `<article class="group"><button class="group-head" type="button" onclick="this.closest('.group').classList.toggle('collapsed')"><div class="group-title">${esc(reason)}</div><span class="group-count">${items.length}</span></button><div class="group-body">${items.map(occHtml).join('')}</div></article>`;
}

function occHtml(o){
  const state=o.status==='pending'?'<span class="state pending">A aguardar validação</span>':'<span class="state open">Ocorrência aberta</span>';
  return `<div class="occ"><div class="occ-head"><div><div class="occ-title">${esc(o.floorLabel)} · ${esc(o.point||'Extintor')}</div><div class="occ-meta">${esc(o.location||'Sem localização indicada')}</div></div>${state}</div><div class="occ-meta"><strong>Data:</strong> ${esc(niceDate(o.createdAt)||'—')}<br><strong>Reportado por:</strong> ${esc(o.reportedBy||'—')}<br><strong>Observação:</strong> ${esc(o.notes||'—')}<br>${o.photoUrl?`<strong>Foto:</strong> <a class="photo" href="${attr(o.photoUrl)}" target="_blank" rel="noopener noreferrer">Abrir fotografia</a>`:'<strong>Foto:</strong> —'}</div></div>`;
}

function groupBy(list,fn){return list.reduce((a,x)=>{const k=fn(x);(a[k]=a[k]||[]).push(x);return a;},{});}
function formatFloor(f){const n=Number(f);if(Number.isNaN(n)) return String(f||'—');if(n<0) return `p${n}`;if(n>0) return `${n}º`;return 'Piso 0';}
function niceDate(v){if(!v) return '';const d=new Date(v);if(Number.isNaN(d.getTime())) return String(v);return new Intl.DateTimeFormat('pt-PT',{dateStyle:'short',timeStyle:'short'}).format(d);}
function esc(s){return String(s??'').replace(/[&<>\"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));}
function attr(s){return esc(s).replace(/`/g,'&#096;');}

function buildReportHtml(){
  const now=new Intl.DateTimeFormat('pt-PT',{dateStyle:'short',timeStyle:'short'}).format(new Date());
  const open=FILTERED.filter(o=>o.status==='open').length;
  const pending=FILTERED.filter(o=>o.status==='pending').length;
  const groups=groupBy(FILTERED,o=>o.reason||'Sem tipo indicado');
  const body=Object.keys(groups).sort((a,b)=>groups[b].length-groups[a].length||a.localeCompare(b,'pt')).map(reason=>`<h2>${esc(reason)} (${groups[reason].length})</h2>${groups[reason].map(o=>`<div class="occ"><strong>${esc(o.floorLabel)} · ${esc(o.point)}</strong><br>${esc(o.location||'Sem localização')}<br>Estado: ${o.status==='pending'?'A aguardar validação':'Ocorrência aberta'}<br>Data: ${esc(niceDate(o.createdAt)||'—')}<br>Reportado por: ${esc(o.reportedBy||'—')}<br>Observação: ${esc(o.notes||'—')}<br>Foto: ${o.photoUrl?`<a href="${attr(o.photoUrl)}">Abrir fotografia</a>`:'—'}</div>`).join('')}`).join('');
  return `<!doctype html><html><head><meta charset="utf-8"><title>Relatório de extintores</title><style>body{font-family:Arial,sans-serif;color:#111827;margin:28px}h1{color:#1A2F5A;margin-bottom:4px}h2{color:#1A2F5A;border-top:1px solid #ddd;padding-top:16px;margin-top:20px}.meta{color:#666;margin-bottom:18px}.summary{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin:16px 0}.box{border:1px solid #ddd;border-radius:10px;padding:12px}.n{font-size:24px;font-weight:700}.occ{border:1px solid #e5e7eb;border-radius:10px;padding:12px;margin:8px 0;line-height:1.5;break-inside:avoid}a{color:#2563EB;font-weight:700}@media print{button{display:none}}</style></head><body><h1>Relatório de extintores</h1><div class="meta">Domingos da Cunha 4 · Gerado em ${esc(now)}</div><div class="summary"><div class="box"><div class="n">${FILTERED.length}</div>Total no relatório</div><div class="box"><div class="n">${open}</div>Ocorrências abertas</div><div class="box"><div class="n">${pending}</div>A aguardar validação</div></div>${body||'<p>Sem ocorrências para apresentar.</p>'}<script>window.onload=()=>setTimeout(()=>window.print(),300)<\/script></body></html>`;
}

async function generateReport({email}){
  const msg=$('actionMsg');
  msg.innerHTML='';
  if(!FILTERED.length){msg.innerHTML='<div class="msg err">Não existem dados para gerar relatório.</div>';return;}
  if(email){
    const to=$('reportEmail').value.trim();
    if(!/^\S+@\S+\.\S+$/.test(to)){msg.innerHTML='<div class="msg err">Indica um email válido.</div>';return;}
    msg.innerHTML='<div class="msg info">A tentar enviar pelo backend...</div>';
    try{
      const r=await apiPost({action:'sendExtinguisherReport',email:to,items:JSON.stringify(FILTERED)});
      if(r.ok||r.success){msg.innerHTML='<div class="msg ok">Relatório enviado por email.</div>'; if(r.pdfUrl) showPdfLink(r.pdfUrl); return;}
      throw new Error(r.message||r.msg||'Backend ainda não tem envio de relatório.');
    }catch(e){
      const subject=encodeURIComponent('Relatório de extintores - Domingos da Cunha 4');
      const body=encodeURIComponent(buildPlainSummary());
      window.location.href=`mailto:${encodeURIComponent(to)}?subject=${subject}&body=${body}`;
      msg.innerHTML='<div class="msg info">O envio com PDF ainda depende da atualização do backend. Abri uma mensagem de email com o resumo em texto.</div>';
      return;
    }
  }
  const w=window.open('','_blank');
  if(!w){msg.innerHTML='<div class="msg err">O browser bloqueou a janela do PDF. Permite pop-ups ou tenta novamente.</div>';return;}
  w.document.write(buildReportHtml());
  w.document.close();
  msg.innerHTML='<div class="msg ok">Relatório aberto. Usa “Guardar como PDF” na janela de impressão.</div>';
}

function buildPlainSummary(){
  const groups=groupBy(FILTERED,o=>o.reason||'Sem tipo indicado');
  let out='Relatório de extintores - Domingos da Cunha 4\n\n';
  out+=`Total: ${FILTERED.length}\nOcorrências abertas: ${FILTERED.filter(o=>o.status==='open').length}\nA aguardar validação: ${FILTERED.filter(o=>o.status==='pending').length}\n\n`;
  Object.keys(groups).forEach(reason=>{
    out+=`${reason} (${groups[reason].length})\n`;
    groups[reason].forEach(o=>{out+=`- ${o.floorLabel} · ${o.point} · ${o.location||'Sem localização'} · ${o.status==='pending'?'A aguardar validação':'Aberta'}\n`;});
    out+='\n';
  });
  return out;
}

function showPdfLink(url){const a=$('pdfLink');a.href=url;a.classList.remove('hidden');}
