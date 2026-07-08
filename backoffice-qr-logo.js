// V2.2.1 - QR codes com logótipo no centro.
// Corrige a ordem de carregamento: injeta CSS logo que este ficheiro carrega e volta a renderizar os QR codes.

function getQrLogoUrl(){
  if (BO_CONFIG?.qr?.logoUrl) return BO_CONFIG.qr.logoUrl;
  return 'assets/logo-qrcode.png?v=2';
}

function buildQrImageUrl(text){
  const size = Number(BO_CONFIG?.qr?.imageSizePx || BO_CONFIG?.qr?.sizePx || 720);
  return `https://api.qrserver.com/v1/create-qr-code/?size=${size}x${size}&ecc=H&qzone=2&color=0f2747&bgcolor=FFFFFF&data=${encodeURIComponent(text)}`;
}

function injectQrLogoStyles(){
  if(document.getElementById('qr-logo-style')) return;
  const style=document.createElement('style');
  style.id='qr-logo-style';
  style.textContent=`
    .qr-card .qr-logo-wrap,
    .qr-logo-wrap{
      position:relative!important;
      display:inline-flex!important;
      align-items:center!important;
      justify-content:center!important;
      width:112px!important;
      height:112px!important;
      max-width:100%!important;
      margin:0 auto!important;
      overflow:visible!important;
    }
    .qr-logo-wrap .qr-base{
      width:100%!important;
      height:100%!important;
      max-width:none!important;
      display:block!important;
    }
    .qr-logo-wrap .qr-center-logo{
      position:absolute!important;
      left:50%!important;
      top:50%!important;
      width:31%!important;
      height:31%!important;
      transform:translate(-50%,-50%)!important;
      object-fit:contain!important;
      background:#fff!important;
      border:4px solid #fff!important;
      border-radius:10px!important;
      box-shadow:0 1px 4px rgba(15,39,71,.20)!important;
      padding:1px!important;
      z-index:3!important;
    }
    #singleQrPreview .qr-logo-wrap{width:168px!important;height:168px!important;}
    @media(min-width:768px){.qr-card .qr-logo-wrap{width:116px!important;height:116px!important;}}
  `;
  document.head.appendChild(style);
}

function qrLogoMarkup(qrUrl,label,extraClass=''){
  const logo = getQrLogoUrl();
  return `<div class="qr-logo-wrap ${extraClass}"><img class="qr-base" src="${escapeAttr(qrUrl)}" alt="QR ${escapeAttr(label)}"><img class="qr-center-logo" src="${escapeAttr(logo)}" alt="Domingos da Cunha Nº 4"></div>`;
}

(function(){
  injectQrLogoStyles();
  const previousInjectQrStyles = typeof injectQrStyles === 'function' ? injectQrStyles : null;
  injectQrStyles = function injectQrStylesWithLogo(){
    if(previousInjectQrStyles) previousInjectQrStyles();
    injectQrLogoStyles();
  };
})();

function renderSingleQrPreview(){
  injectQrLogoStyles();
  const key=bo.singleQrSelect?.value||'';
  if(!key){bo.singleQrPreview.classList.add('hidden');bo.singleQrPreview.innerHTML='';return}
  const point=getAllPoints().find(x=>makeKey(x.floor,x.point)===key);
  if(!point)return;
  const reportUrl=buildReportUrl(point);
  const imgUrl=buildQrImageUrl(reportUrl);
  const label=getQrPrintLabel(point);
  bo.singleQrPreview.classList.remove('hidden');
  bo.singleQrPreview.innerHTML=`<div style="display:flex; gap:14px; flex-wrap:wrap; align-items:center;"><div class="qr-card" style="width:190px; min-height:auto;">${qrLogoMarkup(imgUrl,label)}<div class="qr-label">${escapeHtml(label)}</div></div><div style="flex:1; min-width:240px;"><div class="kv"><div><strong>Ponto:</strong> ${escapeHtml(label)}</div><div><strong>Localização:</strong> ${escapeHtml(point.location||'—')}</div><div><strong>Logótipo:</strong> Domingos da Cunha Nº 4 no centro do QR</div><div><strong>URL:</strong></div></div><div class="code-box" style="margin-top:10px;">${escapeHtml(reportUrl)}</div></div></div>`;
}

function createQrCard(point){
  injectQrLogoStyles();
  const card=document.createElement('button');
  card.type='button';
  card.className='qr-card';
  const key=makeKey(point.floor,point.point);
  card.dataset.key=key;
  card.classList.toggle('selected',BO_QR_SELECTED.has(key));
  const url=buildReportUrl(point);
  const qr=buildQrImageUrl(url);
  const label=getQrPrintLabel(point);
  card.innerHTML=`${qrLogoMarkup(qr,label)}<div class="qr-label">${escapeHtml(label)}</div><div class="qr-url-mini">${escapeHtml(url)}</div>`;
  card.addEventListener('click',()=>toggleQrSelection(point));
  return card;
}

function openPrintWindow(points,singleMode){
  const title=singleMode?'QR Code Individual':'QR Codes - Extintores';
  const logo=getQrLogoUrl();
  const items=points.map(p=>{
    const url=buildReportUrl(p);
    const qr=buildQrImageUrl(url);
    const label=getQrPrintLabel(p);
    return `<div class="item"><div class="qr-logo-wrap"><img class="qr-base" src="${escapeAttr(qr)}" alt="QR"><img class="qr-center-logo" src="${escapeAttr(logo)}" alt="Logo"></div><div class="label">${escapeHtml(label)}</div></div>`;
  }).join('');
  const html=`<!doctype html><html><head><meta charset="utf-8"><title>${title}</title><style>@page{size:A4 portrait;margin:8mm}*{box-sizing:border-box}body{font-family:Arial,sans-serif;margin:0;color:#111827}.grid{display:grid;grid-template-columns:repeat(4,1fr);gap:5mm}.item{border:1px solid #E5E7EB;border-radius:6px;padding:4mm 3mm;text-align:center;break-inside:avoid;min-height:42mm;display:flex;flex-direction:column;align-items:center;justify-content:center}.qr-logo-wrap{position:relative;width:2.8cm;height:2.8cm;display:inline-flex;align-items:center;justify-content:center}.qr-base{width:100%;height:100%;display:block}.qr-center-logo{position:absolute;left:50%;top:50%;width:31%;height:31%;transform:translate(-50%,-50%);object-fit:contain;background:#fff;border:2px solid #fff;border-radius:4px;padding:1px;z-index:3}.label{font-weight:700;margin-top:2mm;font-size:10px;line-height:1.2}@media print{.item{page-break-inside:avoid}}</style></head><body><div class="grid">${items}</div><script>window.onload=()=>window.print()<\/script></body></html>`;
  const w=window.open('','_blank');
  if(!w)return showBackofficeToast('O browser bloqueou a janela de impressão.');
  w.document.write(html);
  w.document.close();
}

// Se o módulo foi carregado depois do primeiro desenho dos QR codes, redesenha já.
setTimeout(()=>{
  injectQrLogoStyles();
  if(typeof renderQrSection === 'function' && document.getElementById('qrGrid')) renderQrSection();
}, 250);
