let CFG=null,POINT=null,FILE=null,EXISTING=null,BUSY=false,LOGIN_BUSY=false,REGISTER_BUSY=false;
const $=id=>document.getElementById(id);
document.addEventListener('DOMContentLoaded',init);
async function init(){
 try{CFG=await loadConfig();POINT=findPointFromUrl();fillReasons();renderPoint();bind();await loadExisting();}
 catch(e){console.error(e);$('pointCard').textContent='Não foi possível abrir este QR: '+e.message;$('submitBtn').disabled=true;}
}
function bind(){
 $('cameraBtn').onclick=()=>$('cameraInput').click();$('fileBtn').onclick=()=>$('fileInput').click();
 $('cameraInput').onchange=pickPhoto;$('fileInput').onchange=pickPhoto;$('reportForm').onsubmit=submit;
 $('loginForm').onsubmit=loginAndSend;$('registerForm').onsubmit=register;
 $('authBack').onclick=()=>$('authDialog').close();
 $('showRegister').onclick=showRegister;$('showLogin').onclick=showLogin;
 $('regFloor').onchange=fillFractions;
}
async function loadConfig(){const r=await dc4Fetch('config.json',{cache:'no-store'});if(!r.ok)throw new Error('Configuração não encontrada.');return r.json();}
function findPointFromUrl(){const p=new URLSearchParams(location.search);if(!p.has('floor')||p.get('floor').trim()==='')throw new Error('QR code sem piso.');const floor=Number(p.get('floor'));const point=String(p.get('point')||'').trim().toUpperCase();if(!Number.isFinite(floor)||!point)throw new Error('QR code sem piso ou extintor.');const floors=CFG?.building?.floors||[];for(const f of floors){for(const e of (f.extinguishers||[])){if(Number(f.floor)===floor&&String(e.point).toUpperCase()===point){return {floor:f.floor,floorLabel:floorLabel(f.floor,f.label),point:e.point,label:e.label||e.point,shortLabel:e.shortLabel||e.point,location:e.location||''};}}}throw new Error('O extintor indicado no QR code não foi encontrado.');}
function fillReasons(){const select=$('reason');select.innerHTML='<option value="">Selecionar motivo</option>';const reasons=CFG?.reportReasons||['Extintor em falta','Extintor danificado','Suporte vazio','Selo / verificação em falta','Acesso obstruído','Outro'];reasons.forEach(r=>{const o=document.createElement('option');o.value=r;o.textContent=r;select.appendChild(o);});}
function renderPoint(){document.title='Reportar '+POINT.floorLabel+' · '+POINT.label;$('introText').textContent='Extintor selecionado. Preencha a anomalia; o acesso de morador só é necessário ao submeter.';$('pointCard').innerHTML='<p class="point-title">'+escapeHtml(POINT.floorLabel)+' · '+escapeHtml(POINT.label)+'</p><p class="point-meta">'+escapeHtml(POINT.location||'Sem localização indicada')+'</p><span class="status pending" id="statusBadge">A consultar estado…</span><div class="existing" id="existingBox"></div>';}
async function loadExisting(){
 try{
  const result=await apiGet('status',{details:'public'});
  if(!result||result.success!==true||!Array.isArray(result.reported))throw new Error('Estado indisponível');
  EXISTING=result.reported.find(x=>makeKey(x.floor,x.point)===makeKey(POINT.floor,POINT.point))||null;
  if(EXISTING)renderExisting(EXISTING);
  else{$('statusBadge').className='status ok';$('statusBadge').textContent='Sem ocorrência validada';$('existingBox').textContent='Os reportes ainda pendentes de validação não aparecem nesta consulta pública.';}
 }catch(e){$('statusBadge').className='status pending';$('statusBadge').textContent='Estado por confirmar';$('existingBox').textContent='Não foi possível consultar o estado. Pode continuar a preparar o reporte.';}
}
function renderExisting(o){
 $('statusBadge').className='status open';$('statusBadge').textContent='Ocorrência aberta';
 const box=$('existingBox');box.replaceChildren();
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
}

function asList(p){if(!p)return[];if(Array.isArray(p))return p;return p.occurrences||p.items||p.data||[];}
function floorLabel(floor,label){const n=Number(floor);if(Number.isFinite(n)&&n>0)return n+'º';return String(label||floor||'');}
function pickPhoto(ev){const f=ev.target.files&&ev.target.files[0];if(!f)return;if(!f.type.startsWith('image/')){toast('Escolha uma imagem válida.');ev.target.value='';return;}if(f.size>18*1024*1024){toast('A fotografia é demasiado grande. Use uma imagem até 18 MB.');ev.target.value='';return;}FILE=f;$('photoMeta').textContent=f.name+' · '+formatBytes(f.size)+' · será reduzida antes do envio';}
async function submit(ev){
 ev.preventDefault();if(BUSY)return;$('msg').textContent='';
 if(!POINT||!$('reason').value){setMsg('err','Selecione o motivo.');return;}
 if(CFG?.features?.reportPhotoRequired&&!FILE){setMsg('err','A fotografia é obrigatória neste reporte.');return;}
 if(!window.dc4HasSession?.('resident')){openAuth();return;}
 await sendReport();
}
function openAuth(){showLogin();$('pin').value='';$('authMsg').textContent='';if(!$('authDialog').open)$('authDialog').showModal();$('pin').focus();}
function showLogin(){$('registerForm').classList.add('hidden');$('loginForm').classList.remove('hidden');$('authTitle').textContent='Confirmar envio';}
async function showRegister(){
 $('loginForm').classList.add('hidden');$('registerForm').classList.remove('hidden');$('authTitle').textContent='Pedir acesso de morador';
 if($('regFloor').options.length)return;
 try{
  const result=await apiGet('garage.structure');if(!Array.isArray(result.floors))throw new Error('Estrutura indisponível');
  window.QR_STRUCTURE=result.floors;$('regFloor').innerHTML='<option value="">Selecionar piso</option>';
  result.floors.forEach(f=>{const o=document.createElement('option');o.value=f.piso;o.textContent=f.label;$('regFloor').appendChild(o);});fillFractions();
 }catch(e){$('registerMsg').textContent='Não foi possível carregar os pisos. Volte a tentar.';}
}
function fillFractions(){
 $('regFraction').innerHTML='<option value="">Selecionar fração</option>';
 const f=(window.QR_STRUCTURE||[]).find(f=>String(f.piso)===$('regFloor').value);
 (f?.fraccoes||[]).forEach(v=>{const o=document.createElement('option');o.value=v;o.textContent=String(v).replaceAll('_',' ');$('regFraction').appendChild(o);});
}
async function register(ev){
 ev.preventDefault();if(REGISTER_BUSY)return;REGISTER_BUSY=true;$('registerSubmit').disabled=true;
 try{
  const result=await apiPost('garage.register',{nome:$('regName').value.trim(),email:$('regEmail').value.trim(),piso:$('regFloor').value,fracao:$('regFraction').value});
  if(!(result.ok||result.success))throw new Error(result.msg||result.message||'Não foi possível pedir acesso.');
  $('registerMsg').textContent='Pedido de acesso enviado. Aguarde a aprovação e o PIN por email. O reporte da anomalia ainda não foi enviado; o formulário e a fotografia mantêm-se enquanto esta página estiver aberta.';
 }catch(e){$('registerMsg').textContent=e.message;}
 finally{REGISTER_BUSY=false;$('registerSubmit').disabled=false;}
}
async function loginAndSend(ev){
 ev.preventDefault();if(LOGIN_BUSY||BUSY)return;
 const pin=$('pin').value.trim();if(!/^\d{6}$/.test(pin)){$('authMsg').textContent='Introduza os 6 dígitos do PIN de morador.';return;}
 LOGIN_BUSY=true;$('loginSubmit').disabled=true;$('authMsg').textContent='';
 try{
  const login=await apiPost('garage.loginPin',{pin,userAgent:navigator.userAgent});
  if(!(login.ok||login.success)||!login.token)throw new Error(login.msg||login.message||'Não foi possível validar o acesso.');
  try{sessionStorage.setItem('dc4_user',JSON.stringify({nome:login.nome,piso:login.piso,fracao:login.fracao}));}catch(e){}
  $('pin').value='';
  if(!$('authDialog').open)return; // Cancelar durante o login não envia o reporte.
  $('authDialog').close();await sendReport();
 }catch(e){$('authMsg').textContent=e.message;}
 finally{LOGIN_BUSY=false;$('loginSubmit').disabled=false;}
}
async function sendReport(){
 if(BUSY)return;BUSY=true;const btn=$('submitBtn');btn.disabled=true;
 try{
  btn.textContent=FILE?'A preparar fotografia…':'A enviar…';
  const photo=FILE?await fileToPayload(FILE):null;btn.textContent='A enviar…';
  const result=await apiPost('report',{floor:Number(POINT.floor),point:POINT.point,location:POINT.location,reason:$('reason').value,notes:$('notes').value.trim(),photoBase64:photo?.base64||'',photoDataUrl:photo?.dataUrl||'',photoName:photo?.name||'',photoType:photo?.type||'',clientTs:new Date().toISOString(),source:'github-pages-qr-direct'});
  if(!(result.success||result.ok))throw new Error(result.message||result.msg||'Não foi possível enviar o reporte.');
  $('reportForm').reset();FILE=null;$('photoMeta').textContent='Sem fotografia selecionada.';
  setMsg('ok','Reporte enviado. Aguarda validação da administração.');toast('Reporte enviado. Obrigado!');
 }catch(e){
  if(e.code==='AUTH_REQUIRED'){openAuth();$('authMsg').textContent='A sessão terminou. Entre para concluir o envio.';}
  else setMsg('err',e.message||'Não foi possível enviar. O formulário foi mantido.');
 }finally{BUSY=false;btn.disabled=false;btn.textContent='Submeter reporte';}
}
async function apiGet(action,params={}){const u=new URL(CFG.backendUrl);u.searchParams.set('action',action);Object.entries(params).forEach(([key,value])=>u.searchParams.set(key,String(value)));const r=await dc4Fetch(u,{cache:'no-store'});return parseJson(r);}
async function apiPost(action,data){const r=await dc4Fetch(CFG.backendUrl,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify(Object.assign({action},data||{}))});return parseJson(r);}
async function parseJson(r){
 const text=await r.text();let result;try{result=JSON.parse(text);}catch(e){throw new Error('O serviço devolveu uma resposta inválida. O reporte não foi confirmado.');}
 if(!r.ok)throw new Error(result.message||'Erro de ligação ao serviço.');return result;
}
async function fileToPayload(file){const originalSize=file.size;const compressed=await compressImage(file,{maxWidth:1280,maxHeight:1280,quality:.72,maxBytes:450*1024});return {...compressed,originalSize,base64:(compressed.dataUrl.split(',')[1]||'')};}
async function compressImage(file,opt){if(!/^image\/(jpeg|jpg|png|webp)$/i.test(file.type)){const dataUrl=await readFile(file);return{dataUrl,type:file.type||'image/jpeg',name:file.name||'foto.jpg',size:file.size};}const img=await loadImg(file);let w=img.width,h=img.height;const ratio=Math.min(opt.maxWidth/w,opt.maxHeight/h,1);w=Math.max(1,Math.round(w*ratio));h=Math.max(1,Math.round(h*ratio));const c=document.createElement('canvas');c.width=w;c.height=h;c.getContext('2d',{alpha:false}).drawImage(img,0,0,w,h);let q=opt.quality,d=c.toDataURL('image/jpeg',q),s=estimate(d);while(s>opt.maxBytes&&q>.45){q=Math.max(.45,q-.08);d=c.toDataURL('image/jpeg',q);s=estimate(d);}return{dataUrl:d,type:'image/jpeg',name:String(file.name||'foto').replace(/\.[^.]+$/,'')+'.jpg',size:s};}
function loadImg(file){return new Promise((res,rej)=>{const url=URL.createObjectURL(file);const img=new Image();img.onload=()=>{URL.revokeObjectURL(url);res(img)};img.onerror=()=>{URL.revokeObjectURL(url);rej(new Error('Não foi possível preparar a fotografia.'))};img.src=url;});}
function readFile(file){return new Promise((res,rej)=>{const fr=new FileReader();fr.onload=()=>res(String(fr.result));fr.onerror=()=>rej(new Error('Não foi possível ler a fotografia.'));fr.readAsDataURL(file);});}
function estimate(dataUrl){return Math.round((String(dataUrl).split(',')[1]||'').length*.75);}function makeKey(f,p){return Number(f)+':'+String(p||'').trim().toUpperCase();}function escapeHtml(s){return String(s||'').replace(/[&<>\"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));}function setMsg(type,text){$('msg').innerHTML='<div class="msg '+type+'">'+escapeHtml(text)+'</div>';}function toast(t){const e=$('toast');e.textContent=t;e.classList.add('show');clearTimeout(toast._t);toast._t=setTimeout(()=>e.classList.remove('show'),3200);}function formatBytes(bytes){const u=['B','KB','MB','GB'];let v=bytes||0,i=0;while(v>=1024&&i<u.length-1){v/=1024;i++;}return v.toFixed(v>=10||i===0?0:1)+' '+u[i];}

