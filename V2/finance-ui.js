
/* UI financeira comum: estado imediato, mensagens temporárias e chamadas autenticadas. */
(function(){
'use strict';
let timer=0, inProgress=false, basePromise=null;
const $=id=>document.getElementById(id);
const node=(tag,value,cls)=>{const n=document.createElement(tag);if(value!==undefined)n.textContent=String(value);if(cls)n.className=cls;return n};
const date=v=>{if(!v)return '—';const s=String(v).slice(0,10);if(!/^\d{4}-\d{2}-\d{2}$/.test(s))return s;return s.slice(8,10)+'/'+s.slice(5,7)+'/'+s.slice(0,4)};
const money=v=>Number.isFinite(Number(v))?Number(v).toLocaleString('pt-PT',{style:'currency',currency:'EUR'}):'—';
const iso=()=>new Date().toLocaleDateString('sv-SE');
async function base(){
 if(!basePromise)basePromise=fetch('config.json',{cache:'no-store'}).then(r=>{if(!r.ok)throw Error('Configuração indisponível');return r.json()})
 .then(v=>v.backendUrl).catch(err=>{basePromise=null;throw err});
 return basePromise;
}
async function api(action,payload={}){
 const root=await base();
 const res=await dc4Fetch(root+'?action='+encodeURIComponent(action),
  {method:'POST',body:JSON.stringify({action,...payload})});
 const json=await res.json();
 const obj=json.data&&typeof json.data==='object'?json.data:json;
 if(obj.ok===false||obj.error||obj.code==='FORBIDDEN')throw Error(obj.message||obj.error||'Ação não autorizada');
 return obj;
}
function toast(message,error=false){
 const el=$('toast');if(!el)return;
 clearTimeout(timer);el.textContent=message;el.className=error?'error':'ok';
 timer=setTimeout(()=>{el.className='';el.textContent=''},error?8000:4200);
}
async function busy(text,fn){
 if(inProgress)return null;
 inProgress=true;const el=$('progress');if(el){el.className='active';el.querySelector('[data-progress-text]').textContent=text;}
 document.querySelectorAll('main button').forEach(b=>{b.dataset.wasDisabled=b.disabled?'1':'0';b.disabled=true;});
 try{return await fn()}catch(err){toast(err.message||'Não foi possível concluir.',true);return null}
 finally{inProgress=false;if(el)el.className='';document.querySelectorAll('main button').forEach(b=>{b.disabled=b.dataset.wasDisabled==='1';delete b.dataset.wasDisabled;});}
}
async function imageData(file){
 if(!file) return '';
 if(!['image/jpeg','image/png'].includes(file.type))throw Error('Seleciona uma fotografia JPG ou PNG.');
 return new Promise((resolve,reject)=>{
   const url=URL.createObjectURL(file),im=new Image();
   im.onload=()=>{
     try{
       const scale=Math.min(1,1280/Math.max(im.width,im.height));
       const canvas=document.createElement('canvas');canvas.width=Math.max(1,Math.round(im.width*scale));
       canvas.height=Math.max(1,Math.round(im.height*scale));
       canvas.getContext('2d').drawImage(im,0,0,canvas.width,canvas.height);
       const data=canvas.toDataURL('image/jpeg',.72);
       if(data.length>4000000)throw Error('Fotografia demasiado grande.');
       resolve(data);
     }catch(e){reject(e)}finally{URL.revokeObjectURL(url)}
   };
   im.onerror=()=>{URL.revokeObjectURL(url);reject(Error('Erro ao abrir fotografia'))};im.src=url;
 });
}
window.FinUI={api,busy,toast,node,date,money,iso,imageData,$};
})();
