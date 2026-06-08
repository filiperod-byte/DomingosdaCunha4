let BO_CONFIG = null;
let BO_OPEN_OCCURRENCES = [];
let BO_PENDING_OCCURRENCES = [];
let BO_REPORTED_SET = new Set();
let BO_CLOSE_FILE = null;
let BO_QR_SELECTED = new Set();

const bo = {};
document.addEventListener('DOMContentLoaded', initBackoffice);

async function initBackoffice(){
  cacheBackofficeElements();
  bindBackofficeEvents();
  injectQrStyles();
  try{
    BO_CONFIG = await loadBackofficeConfig();
    renderStaticConfigInfo();
    await bootPinFlow();
  }catch(error){
    console.error(error);
    showBackofficeToast('Erro ao carregar o backoffice. Verifica o config.json.');
  }
}

function cacheBackofficeElements(){
  bo.pinStatusBadge = document.getElementById('pinStatusBadge');
  bo.backendBadge = document.getElementById('backendBadge');
  bo.authSection = document.getElementById('authSection');
  bo.setupCard = document.getElementById('setupCard');
  bo.loginCard = document.getElementById('loginCard');
  bo.adminSection = document.getElementById('adminSection');
  bo.setupPin = document.getElementById('setupPin');
  bo.setupPinConfirm = document.getElementById('setupPinConfirm');
  bo.setupPinBtn = document.getElementById('setupPinBtn');
  bo.loginPin = document.getElementById('loginPin');
  bo.loginBtn = document.getElementById('loginBtn');
  bo.recoverPinBtn = document.getElementById('recoverPinBtn');
  bo.tabButtons = Array.from(document.querySelectorAll('.tab-btn'));
  bo.tabPanels = Array.from(document.querySelectorAll('.tab-panel'));
  bo.statTotal = document.getElementById('statTotal');
  bo.statOk = document.getElementById('statOk');
  bo.statAlert = document.getElementById('statAlert');
  bo.statOpen = document.getElementById('statOpen');
  bo.statPending = document.getElementById('statPending');
  bo.statusGrid = document.getElementById('statusGrid');
  bo.dashboardBackendInfo = document.getElementById('dashboardBackendInfo');
  bo.dashboardRefresh = document.getElementById('dashboardRefresh');
  bo.dashboardAdminEmail = document.getElementById('dashboardAdminEmail');
  bo.refreshDashboardBtn = document.getElementById('refreshDashboardBtn');
  bo.pendingList = document.getElementById('pendingList');
  bo.refreshPendingBtn = document.getElementById('refreshPendingBtn');
  bo.occurrencesList = document.getElementById('occurrencesList');
  bo.refreshOccurrencesBtn = document.getElementById('refreshOccurrencesBtn');
  bo.closeOccurrenceSelect = document.getElementById('closeOccurrenceSelect');
  bo.closeOccurrenceSummary = document.getElementById('closeOccurrenceSummary');
  bo.closePhotoInput = document.getElementById('closePhotoInput');
  bo.pickClosePhotoBtn = document.getElementById('pickClosePhotoBtn');
  bo.closePhotoMeta = document.getElementById('closePhotoMeta');
  bo.closeNote = document.getElementById('closeNote');
  bo.closeOccurrenceBtn = document.getElementById('closeOccurrenceBtn');
  bo.singleQrSelect = document.getElementById('singleQrSelect');
  bo.singleQrPreview = document.getElementById('singleQrPreview');
  bo.qrGrid = document.getElementById('qrGrid');
  bo.printSingleQrBtn = document.getElementById('printSingleQrBtn');
  bo.printAllQrBtn = document.getElementById('printAllQrBtn');
  bo.configBackendUrl = document.getElementById('configBackendUrl');
  bo.configAdminEmail = document.getElementById('configAdminEmail');
  bo.configPublicUrl = document.getElementById('configPublicUrl');
  bo.configQrSize = document.getElementById('configQrSize');
  bo.configRecoverPinBtn = document.getElementById('configRecoverPinBtn');
  bo.logoutBtn = document.getElementById('logoutBtn');
  bo.toast = document.getElementById('toast');
}

function bindBackofficeEvents(){
  bo.setupPinBtn?.addEventListener('click', handleSetupPin);
  bo.loginBtn?.addEventListener('click', handleLogin);
  bo.recoverPinBtn?.addEventListener('click', handleRecoverPin);
  bo.tabButtons.forEach(btn => btn.addEventListener('click', () => activateTab(btn.dataset.tab)));
  bo.refreshDashboardBtn?.addEventListener('click', refreshAllAdminData);
  bo.refreshPendingBtn?.addEventListener('click', loadPendingOccurrences);
  bo.refreshOccurrencesBtn?.addEventListener('click', loadOccurrences);
  bo.pickClosePhotoBtn?.addEventListener('click', () => bo.closePhotoInput.click());
  bo.closePhotoInput?.addEventListener('change', handleClosePhotoPicked);
  bo.closeOccurrenceSelect?.addEventListener('change', renderCloseOccurrenceSummary);
  bo.closeOccurrenceBtn?.addEventListener('click', handleCloseOccurrence);
  bo.singleQrSelect?.addEventListener('change', renderSingleQrPreview);
  bo.printSingleQrBtn?.addEventListener('click', printSelectedQrs);
  bo.printAllQrBtn?.addEventListener('click', printAllQrs);
  bo.configRecoverPinBtn?.addEventListener('click', handleRecoverPin);
  bo.logoutBtn?.addEventListener('click', logoutBackoffice);
  bo.loginPin?.addEventListener('keydown', e => { if(e.key === 'Enter') handleLogin(); });
  bo.setupPinConfirm?.addEventListener('keydown', e => { if(e.key === 'Enter') handleSetupPin(); });
}

function injectQrStyles(){
  const style = document.createElement('style');
  style.id = 'qr-select-print-style';
  style.textContent = `
    .qr-toolbar-enhanced{display:flex;flex-wrap:wrap;gap:10px;margin:0 0 14px;align-items:center}
    .qr-counter{color:var(--muted);font-weight:800;font-size:.9rem;margin-left:auto}
    .qr-grid.selectable{grid-template-columns:repeat(2,1fr)}
    .qr-card{position:relative;min-height:150px;padding:10px;cursor:pointer;transition:.16s;user-select:none}
    .qr-card:hover{transform:translateY(-1px)}
    .qr-card.selected{border-color:var(--blue);box-shadow:0 0 0 3px rgba(37,99,235,.14),var(--shadow)}
    .qr-card.selected:after{content:'✓';position:absolute;top:8px;right:8px;width:24px;height:24px;border-radius:999px;background:var(--blue);color:#fff;display:flex;align-items:center;justify-content:center;font-weight:900;font-size:.86rem}
    .qr-card img{max-width:94px;width:94px;height:94px}
    .qr-label{font-size:.76rem;line-height:1.2;margin-top:6px}
    .qr-url-mini{display:none}
    @media(min-width:768px){.qr-grid.selectable{grid-template-columns:repeat(5,1fr)}.qr-card img{max-width:104px;width:104px;height:104px}}
  `;
  document.head.appendChild(style);
}

async function loadBackofficeConfig(){
  const res = await fetch('config.json', { cache:'no-store' });
  if(!res.ok) throw new Error('Não foi possível carregar o config.json');
  return res.json();
}

function renderStaticConfigInfo(){
  const ok = !!BO_CONFIG?.backendUrl && !BO_CONFIG.backendUrl.includes('COLOCAR_AQUI');
  bo.backendBadge.innerHTML = ok ? '<span class="dot ok"></span> Backend configurado' : '<span class="dot danger"></span> Backend por configurar';
  bo.configBackendUrl.textContent = BO_CONFIG?.backendUrl || '—';
  bo.configAdminEmail.textContent = BO_CONFIG?.adminEmail || '—';
  bo.configPublicUrl.textContent = getPublicIndexUrl();
  bo.configQrSize.textContent = '2,8 cm';
  bo.dashboardBackendInfo.textContent = ok ? 'Configurado' : 'Por configurar';
  bo.dashboardAdminEmail.textContent = BO_CONFIG?.adminEmail || '—';
}

async function bootPinFlow(){
  try{
    const pinInfo = await apiGet('pinStatus');
    if(pinInfo?.pinConfigured){
      bo.pinStatusBadge.innerHTML = '<span class="dot ok"></span> PIN definido';
      showLoginCard();
      if(sessionStorage.getItem('extintores_bo_unlocked') === '1') unlockAdminArea();
    }else{
      bo.pinStatusBadge.innerHTML = '<span class="dot warn"></span> PIN ainda não definido';
      showSetupCard();
    }
  }catch(err){
    bo.pinStatusBadge.innerHTML = '<span class="dot danger"></span> Erro no PIN / backend';
    showLoginCard();
  }
}
function showSetupCard(){ bo.authSection.classList.remove('hidden'); bo.setupCard.classList.remove('hidden'); bo.loginCard.classList.add('hidden'); bo.adminSection.classList.add('hidden'); }
function showLoginCard(){ bo.authSection.classList.remove('hidden'); bo.loginCard.classList.remove('hidden'); bo.setupCard.classList.add('hidden'); bo.adminSection.classList.add('hidden'); }
async function handleSetupPin(){
  const pin = bo.setupPin.value.trim(); const confirm = bo.setupPinConfirm.value.trim();
  if(pin.length < 4){ showBackofficeToast('O PIN deve ter pelo menos 4 dígitos.'); return; }
  if(pin !== confirm){ showBackofficeToast('Os PINs não coincidem.'); return; }
  try{
    bo.setupPinBtn.disabled = true; bo.setupPinBtn.textContent = 'A guardar…';
    const r = await apiPost({ action:getAction('setPin'), pin });
    if(r?.success === false) throw new Error(r.message || 'Não foi possível guardar o PIN.');
    sessionStorage.setItem('extintores_bo_unlocked','1');
    unlockAdminArea(); showBackofficeToast('PIN definido com sucesso.');
  }catch(err){ showBackofficeToast(err.message || 'Falha ao guardar PIN.'); }
  finally{ bo.setupPinBtn.disabled = false; bo.setupPinBtn.textContent = 'Guardar PIN'; }
}
async function handleLogin(){
  const pin = bo.loginPin.value.trim();
  if(!pin){ showBackofficeToast('Introduz o PIN.'); return; }
  try{
    bo.loginBtn.disabled = true; bo.loginBtn.textContent = 'A validar…';
    const r = await apiPost({ action:getAction('validatePin'), pin });
    if(r?.success === false || r?.valid === false) throw new Error(r.message || 'PIN inválido.');
    sessionStorage.setItem('extintores_bo_unlocked','1');
    bo.loginPin.value = '';
    unlockAdminArea(); showBackofficeToast('Acesso autorizado.');
  }catch(err){ showBackofficeToast(err.message || 'PIN inválido.'); }
  finally{ bo.loginBtn.disabled = false; bo.loginBtn.textContent = 'Entrar'; }
}
async function handleRecoverPin(){
  if(!confirm('Isto envia email ao administrador e faz reset ao PIN atual. Continuar?')) return;
  const r = await apiPost({ action:getAction('resetPin'), adminEmail:BO_CONFIG?.adminEmail || '' });
  if(r?.success === false) return showBackofficeToast(r.message || 'Não foi possível recuperar/resetar PIN.');
  sessionStorage.removeItem('extintores_bo_unlocked');
  showBackofficeToast('PIN resetado.');
  showSetupCard();
}
function unlockAdminArea(){ bo.authSection.classList.add('hidden'); bo.adminSection.classList.remove('hidden'); activateTab('dashboard'); refreshAllAdminData(); }
function logoutBackoffice(){ sessionStorage.removeItem('extintores_bo_unlocked'); showBackofficeToast('Sessão terminada.'); showLoginCard(); }
function activateTab(tabName){ bo.tabButtons.forEach(btn => btn.classList.toggle('active', btn.dataset.tab === tabName)); bo.tabPanels.forEach(panel => panel.classList.toggle('active', panel.id === `tab-${tabName}`)); }
async function refreshAllAdminData(){ renderQrSection(); await Promise.all([refreshDashboard(), loadPendingOccurrences(), loadOccurrences()]); }

async function refreshDashboard(){
  try{
    const [statusResponse, pendingResponse] = await Promise.all([apiGet('status'), apiGet('pendingOccurrences')]);
    const items = extractReportedItems(statusResponse);
    BO_REPORTED_SET = new Set(items.map(normalizeStatusKey).filter(Boolean));
    BO_PENDING_OCCURRENCES = extractOccurrences(pendingResponse).map(normalizeOccurrence);
    const total = getAllPoints().length;
    const alertCount = BO_REPORTED_SET.size;
    const okCount = Math.max(total - alertCount, 0);
    bo.statTotal.textContent = String(total);
    bo.statOk.textContent = String(okCount);
    bo.statAlert.textContent = String(alertCount);
    bo.statOpen.textContent = String(BO_OPEN_OCCURRENCES.length);
    if(bo.statPending) bo.statPending.textContent = String(BO_PENDING_OCCURRENCES.length);
    renderStatusGrid();
    bo.dashboardRefresh.textContent = formatDateTime(new Date());
  }catch(err){
    console.error(err);
    bo.statusGrid.innerHTML = '<div class="empty">Não foi possível obter o estado atual.</div>';
  }
}
function renderStatusGrid(){
  const points = getAllPoints();
  bo.statusGrid.innerHTML = '';
  if(!points.length){ bo.statusGrid.innerHTML = '<div class="empty">Sem pontos configurados.</div>'; return; }
  points.forEach(point => {
    const key = makeKey(point.floor, point.point);
    const alert = BO_REPORTED_SET.has(key);
    const div = document.createElement('div');
    div.className = `status-pill ${alert ? 'alert' : ''}`;
    div.innerHTML = `<div><strong>${escapeHtml(point.floorLabel)} · ${escapeHtml(point.label)}</strong><div class="tiny muted">${escapeHtml(point.location || 'Sem localização')}</div></div><div class="badge"><span class="dot ${alert ? 'danger' : 'ok'}"></span>${alert ? 'Visível' : 'OK'}</div>`;
    bo.statusGrid.appendChild(div);
  });
}

async function loadPendingOccurrences(){
  if(!bo.pendingList) return;
  bo.pendingList.innerHTML = '<div class="loading"><div class="spinner"></div><span>A carregar reportes pendentes…</span></div>';
  try{
    const r = await apiGet('pendingOccurrences');
    BO_PENDING_OCCURRENCES = extractOccurrences(r).map(normalizeOccurrence);
    renderPendingOccurrences();
    if(bo.statPending) bo.statPending.textContent = String(BO_PENDING_OCCURRENCES.length);
  }catch(err){
    console.error(err);
    BO_PENDING_OCCURRENCES = [];
    bo.pendingList.innerHTML = '<div class="empty">Não foi possível carregar reportes pendentes.</div>';
  }
}
function renderPendingOccurrences(){
  bo.pendingList.innerHTML = '';
  if(!BO_PENDING_OCCURRENCES.length){ bo.pendingList.innerHTML = '<div class="empty">Não existem reportes pendentes de validação.</div>'; return; }
  BO_PENDING_OCCURRENCES.forEach(occ => {
    const card = document.createElement('div');
    card.className = 'list-card';
    card.innerHTML = `<div class="list-head"><div><strong>${escapeHtml(occ.floorLabel)} · ${escapeHtml(occ.point || 'Sem ponto')}</strong><div class="tiny muted">${escapeHtml(occ.location || 'Sem localização')}</div></div><span class="badge"><span class="dot warn"></span> Pendente</span></div><div class="kv"><div><strong>ID:</strong> ${escapeHtml(occ.id)}</div><div><strong>Reportado por:</strong> ${escapeHtml(occ.reportedBy || '—')}</div><div><strong>Motivo:</strong> ${escapeHtml(occ.reason || '—')}</div><div><strong>Observação:</strong> ${escapeHtml(occ.notes || '—')}</div><div><strong>Data:</strong> ${escapeHtml(occ.createdAt || '—')}</div>${occ.photoUrl ? `<div><strong>Foto:</strong> <a href="${escapeAttr(occ.photoUrl)}" target="_blank" rel="noopener noreferrer">Abrir fotografia</a></div>` : '<div><strong>Foto:</strong> —</div>'}</div><div class="stack-actions"><button class="btn btn-success" type="button" onclick="approvePendingOccurrence('${escapeAttr(occ.id)}')">Validar reporte</button><button class="btn btn-danger" type="button" onclick="rejectPendingOccurrence('${escapeAttr(occ.id)}')">Rejeitar</button></div>`;
    bo.pendingList.appendChild(card);
  });
}
async function approvePendingOccurrence(id){
  if(!confirm('Validar este reporte e torná-lo visível na consulta pública?')) return;
  const r = await apiPost({ action:'approveOccurrence', occurrenceId:id, adminEmail:BO_CONFIG?.adminEmail || '', source:'github-pages-backoffice' });
  if(r?.success === false) return showBackofficeToast(r.message || 'Não foi possível validar.');
  showBackofficeToast('Reporte validado.');
  await Promise.all([loadPendingOccurrences(), loadOccurrences(), refreshDashboard()]);
}
async function rejectPendingOccurrence(id){
  const notes = prompt('Motivo da rejeição (opcional):') || '';
  const r = await apiPost({ action:'rejectOccurrence', occurrenceId:id, notes, adminEmail:BO_CONFIG?.adminEmail || '', source:'github-pages-backoffice' });
  if(r?.success === false) return showBackofficeToast(r.message || 'Não foi possível rejeitar.');
  showBackofficeToast('Reporte rejeitado.');
  await Promise.all([loadPendingOccurrences(), refreshDashboard()]);
}

async function loadOccurrences(){
  bo.occurrencesList.innerHTML = '<div class="loading"><div class="spinner"></div><span>A carregar ocorrências…</span></div>';
  try{
    const r = await apiGet('openOccurrences');
    BO_OPEN_OCCURRENCES = extractOccurrences(r).map(normalizeOccurrence);
    renderOccurrencesList(); populateCloseOccurrenceSelect(); renderCloseOccurrenceSummary();
    bo.statOpen.textContent = String(BO_OPEN_OCCURRENCES.length);
  }catch(err){
    console.error(err); BO_OPEN_OCCURRENCES = []; renderOccurrencesList(); populateCloseOccurrenceSelect(); renderCloseOccurrenceSummary();
  }
}
function extractOccurrences(payload){ if(!payload) return []; if(Array.isArray(payload)) return payload; return payload.occurrences || payload.items || payload.data || []; }
function normalizeOccurrence(item, index){
  const floor = Number(item.floor ?? item.piso ?? item.floorNumber ?? item.level ?? 0);
  const point = String(item.point ?? item.ponto ?? item.extinguisher ?? item.extintor ?? item.code ?? '').trim();
  const id = item.id ?? item.occurrenceId ?? item.ocorrenciaId ?? `${makeKey(floor, point)}#${index + 1}`;
  return { id:String(id), floor, floorLabel:item.floorLabel || getFloorLabel(floor), point, label:item.label || point, location:item.location ?? item.localizacao ?? item.localização ?? '', reportedBy:item.reportedBy ?? item.name ?? item.nome ?? '', reason:item.reason ?? item.motivo ?? '', notes:item.notes ?? item.observacao ?? item.observação ?? '', createdAt:item.createdAt ?? item.timestamp ?? item.data ?? item.created ?? '', photoUrl:item.photoUrl ?? item.fotoUrl ?? '', status:item.status || '', raw:item };
}
function renderOccurrencesList(){
  bo.occurrencesList.innerHTML = '';
  if(!BO_OPEN_OCCURRENCES.length){ bo.occurrencesList.innerHTML = '<div class="empty">Não existem ocorrências abertas neste momento.</div>'; return; }
  BO_OPEN_OCCURRENCES.forEach(occ => {
    const card = document.createElement('div'); card.className = 'list-card';
    const photo = occ.photoUrl ? `<div><strong>Foto:</strong> <a href="${escapeAttr(occ.photoUrl)}" target="_blank" rel="noopener noreferrer">Abrir fotografia</a></div>` : '<div><strong>Foto:</strong> —</div>';
    card.innerHTML = `<div class="list-head"><div><strong>${escapeHtml(occ.floorLabel)} · ${escapeHtml(occ.point || 'Sem ponto')}</strong><div class="tiny muted">${escapeHtml(occ.location || 'Sem localização')}</div></div><span class="badge"><span class="dot danger"></span> Aberta</span></div><div class="kv"><div><strong>ID:</strong> ${escapeHtml(occ.id)}</div><div><strong>Reportado por:</strong> ${escapeHtml(occ.reportedBy || '—')}</div><div><strong>Motivo:</strong> ${escapeHtml(occ.reason || '—')}</div><div><strong>Observação:</strong> ${escapeHtml(occ.notes || '—')}</div><div><strong>Data:</strong> ${escapeHtml(occ.createdAt || '—')}</div>${photo}</div>`;
    bo.occurrencesList.appendChild(card);
  });
}
function populateCloseOccurrenceSelect(){ bo.closeOccurrenceSelect.innerHTML = '<option value="">Selecionar ocorrência aberta</option>'; BO_OPEN_OCCURRENCES.forEach(occ => { const o = document.createElement('option'); o.value = occ.id; o.textContent = `${occ.floorLabel} · ${occ.point} · ${occ.reason || 'Sem motivo'}`; bo.closeOccurrenceSelect.appendChild(o); }); }
function renderCloseOccurrenceSummary(){ const id = bo.closeOccurrenceSelect.value; const occ = BO_OPEN_OCCURRENCES.find(x => x.id === id); if(!occ){ bo.closeOccurrenceSummary.innerHTML = 'Seleciona uma ocorrência para ver o detalhe.'; return; } bo.closeOccurrenceSummary.innerHTML = `<div class="kv"><div><strong>ID:</strong> ${escapeHtml(occ.id)}</div><div><strong>Ponto:</strong> ${escapeHtml(occ.floorLabel)} · ${escapeHtml(occ.point)}</div><div><strong>Localização:</strong> ${escapeHtml(occ.location || '—')}</div><div><strong>Motivo:</strong> ${escapeHtml(occ.reason || '—')}</div><div><strong>Reportado por:</strong> ${escapeHtml(occ.reportedBy || '—')}</div><div><strong>Data:</strong> ${escapeHtml(occ.createdAt || '—')}</div>${occ.photoUrl ? `<div><strong>Foto abertura:</strong> <a href="${escapeAttr(occ.photoUrl)}" target="_blank" rel="noopener noreferrer">Abrir fotografia</a></div>` : ''}</div>`; }
function handleClosePhotoPicked(event){ const file = event.target.files?.[0]; if(!file) return; if(!file.type.startsWith('image/')){ showBackofficeToast('Seleciona uma imagem válida.'); event.target.value = ''; return; } if(file.size > 8*1024*1024){ showBackofficeToast('A fotografia é demasiado grande.'); event.target.value = ''; return; } BO_CLOSE_FILE = file; bo.closePhotoMeta.textContent = `${file.name} · ${formatBytes(file.size)}`; }
async function handleCloseOccurrence(){ const id = bo.closeOccurrenceSelect.value; const occ = BO_OPEN_OCCURRENCES.find(x => x.id === id); if(!occ){ showBackofficeToast('Seleciona uma ocorrência aberta.'); return; } if((BO_CONFIG?.features?.closePhotoRequired !== false) && !BO_CLOSE_FILE){ showBackofficeToast('A fotografia de fecho é obrigatória.'); return; } try{ bo.closeOccurrenceBtn.disabled = true; bo.closeOccurrenceBtn.textContent = 'A fechar…'; const photo = BO_CLOSE_FILE ? await fileToPayload(BO_CLOSE_FILE) : null; const r = await apiPost({ action:getAction('closeOccurrence'), occurrenceId:occ.id, floor:occ.floor, point:occ.point, closeNotes:bo.closeNote.value.trim(), closePhotoBase64:photo?.base64 || '', closePhotoDataUrl:photo?.dataUrl || '', closePhotoName:photo?.name || '', closePhotoType:photo?.type || '', clientTs:new Date().toISOString(), source:'github-pages-backoffice' }); if(r?.success === false) throw new Error(r.message || 'Falha ao fechar ocorrência.'); showBackofficeToast('Ocorrência fechada.'); bo.closeOccurrenceSelect.value = ''; bo.closeNote.value = ''; bo.closePhotoInput.value = ''; bo.closePhotoMeta.textContent = 'Nenhuma fotografia selecionada.'; BO_CLOSE_FILE = null; await Promise.all([loadOccurrences(), refreshDashboard()]); }catch(err){ showBackofficeToast(err.message || 'Não foi possível fechar.'); } finally{ bo.closeOccurrenceBtn.disabled = false; bo.closeOccurrenceBtn.textContent = 'Fechar ocorrência'; } }

function renderQrSection(){
  const points = getAllPoints();
  BO_QR_SELECTED = new Set(Array.from(BO_QR_SELECTED).filter(key => points.some(p => makeKey(p.floor,p.point) === key)));

  if(bo.singleQrSelect){
    bo.singleQrSelect.innerHTML = '<option value="">Selecionar ponto</option>';
    points.forEach(point => {
      const o = document.createElement('option');
      o.value = makeKey(point.floor, point.point);
      o.textContent = getQrPrintLabel(point);
      bo.singleQrSelect.appendChild(o);
    });
    const field = bo.singleQrSelect.closest('.field');
    if(field) field.style.display = 'none';
  }

  if(bo.printSingleQrBtn){
    bo.printSingleQrBtn.textContent = 'Imprimir selecionados';
    bo.printSingleQrBtn.classList.remove('btn-secondary');
    bo.printSingleQrBtn.classList.add('btn-primary');
  }
  if(bo.printAllQrBtn){
    bo.printAllQrBtn.textContent = 'Imprimir todos';
    bo.printAllQrBtn.classList.remove('btn-primary');
    bo.printAllQrBtn.classList.add('btn-secondary');
  }

  ensureQrToolbar();
  bo.qrGrid.classList.add('selectable');
  bo.qrGrid.innerHTML = '';

  if(!points.length){
    bo.qrGrid.innerHTML = '<div class="empty">Sem QR codes configurados.</div>';
    updateQrCounter();
    return;
  }

  points.forEach(point => bo.qrGrid.appendChild(createQrCard(point)));
  renderSingleQrPreview();
  updateQrCounter();
}
function ensureQrToolbar(){
  const tab = document.getElementById('tab-qrcodes');
  const existing = document.getElementById('qrToolbarEnhanced');
  if(existing) return;
  const toolbar = document.createElement('div');
  toolbar.id = 'qrToolbarEnhanced';
  toolbar.className = 'qr-toolbar-enhanced';
  toolbar.innerHTML = `<button id="selectAllQrBtn" class="btn btn-secondary" type="button">Selecionar todos</button><button id="clearQrSelectionBtn" class="btn btn-secondary" type="button">Limpar seleção</button><span id="qrCounter" class="qr-counter">0 selecionados</span>`;
  const helper = tab.querySelector('.helper');
  tab.insertBefore(toolbar, helper);
  document.getElementById('selectAllQrBtn').addEventListener('click', selectAllQrs);
  document.getElementById('clearQrSelectionBtn').addEventListener('click', clearQrSelection);
}
function updateQrCounter(){
  const counter = document.getElementById('qrCounter');
  if(counter) counter.textContent = `${BO_QR_SELECTED.size} selecionado${BO_QR_SELECTED.size === 1 ? '' : 's'}`;
}
function selectAllQrs(){
  BO_QR_SELECTED = new Set(getAllPoints().map(p => makeKey(p.floor,p.point)));
  renderQrCardsSelection();
  updateQrCounter();
}
function clearQrSelection(){
  BO_QR_SELECTED.clear();
  renderQrCardsSelection();
  updateQrCounter();
}
function renderQrCardsSelection(){
  bo.qrGrid.querySelectorAll('.qr-card').forEach(card => card.classList.toggle('selected', BO_QR_SELECTED.has(card.dataset.key)));
}
function toggleQrSelection(point){
  const key = makeKey(point.floor, point.point);
  if(BO_QR_SELECTED.has(key)) BO_QR_SELECTED.delete(key); else BO_QR_SELECTED.add(key);
  renderQrCardsSelection();
  updateQrCounter();
}
function renderSingleQrPreview(){
  const key = bo.singleQrSelect?.value || '';
  if(!key){ bo.singleQrPreview.classList.add('hidden'); bo.singleQrPreview.innerHTML=''; return; }
  const point = getAllPoints().find(x => makeKey(x.floor, x.point) === key);
  if(!point) return;
  const reportUrl = buildReportUrl(point);
  const imgUrl = buildQrImageUrl(reportUrl);
  bo.singleQrPreview.classList.remove('hidden');
  bo.singleQrPreview.innerHTML = `<div style="display:flex; gap:14px; flex-wrap:wrap; align-items:center;"><div class="qr-card" style="width:180px; min-height:auto;"><img src="${escapeAttr(imgUrl)}" alt="QR Code ${escapeAttr(getQrPrintLabel(point))}"><div class="qr-label">${escapeHtml(getQrPrintLabel(point))}</div></div><div style="flex:1; min-width:240px;"><div class="kv"><div><strong>Ponto:</strong> ${escapeHtml(getQrPrintLabel(point))}</div><div><strong>Localização:</strong> ${escapeHtml(point.location || '—')}</div><div><strong>URL:</strong></div></div><div class="code-box" style="margin-top:10px;">${escapeHtml(reportUrl)}</div></div></div>`;
}
function createQrCard(point){
  const card = document.createElement('button');
  card.type = 'button';
  card.className = 'qr-card';
  const key = makeKey(point.floor, point.point);
  card.dataset.key = key;
  card.classList.toggle('selected', BO_QR_SELECTED.has(key));
  const url = buildReportUrl(point);
  const qr = buildQrImageUrl(url);
  const label = getQrPrintLabel(point);
  card.innerHTML = `<img src="${escapeAttr(qr)}" alt="QR ${escapeAttr(label)}"><div class="qr-label">${escapeHtml(label)}</div><div class="qr-url-mini">${escapeHtml(url)}</div>`;
  card.addEventListener('click', () => toggleQrSelection(point));
  return card;
}
function printSelectedQrs(){
  const selected = getAllPoints().filter(p => BO_QR_SELECTED.has(makeKey(p.floor,p.point)));
  if(!selected.length) return showBackofficeToast('Seleciona pelo menos um QR code.');
  openPrintWindow(selected, false);
}
function printSingleQr(){ printSelectedQrs(); }
function printAllQrs(){ const points = getAllPoints(); if(!points.length) return showBackofficeToast('Não há QR codes.'); openPrintWindow(points, false); }
function openPrintWindow(points, singleMode){
  const title = singleMode ? 'QR Code Individual' : 'QR Codes - Extintores';
  const items = points.map(p => {
    const url = buildReportUrl(p);
    const qr = buildQrImageUrl(url);
    return `<div class="item"><img src="${escapeAttr(qr)}" alt="QR"><div class="label">${escapeHtml(getQrPrintLabel(p))}</div></div>`;
  }).join('');
  const html = `<!doctype html><html><head><meta charset="utf-8"><title>${title}</title><style>@page{size:A4 portrait;margin:8mm}*{box-sizing:border-box}body{font-family:Arial,sans-serif;margin:0;color:#111827}.grid{display:grid;grid-template-columns:repeat(4,1fr);gap:5mm}.item{border:1px solid #E5E7EB;border-radius:6px;padding:4mm 3mm;text-align:center;break-inside:avoid;min-height:42mm;display:flex;flex-direction:column;align-items:center;justify-content:center}.item img{width:2.8cm;height:2.8cm;display:block}.label{font-weight:700;margin-top:2mm;font-size:10px;line-height:1.2}@media print{.item{page-break-inside:avoid}}</style></head><body><div class="grid">${items}</div><script>window.onload=()=>window.print()<\/script></body></html>`;
  const w = window.open('', '_blank');
  if(!w) return showBackofficeToast('O browser bloqueou a janela de impressão.');
  w.document.write(html);
  w.document.close();
}
function buildReportUrl(point){
  const base = getV2EntryUrl();
  const url = new URL(base, window.location.href);
  url.searchParams.set('floor', point.floor);
  url.searchParams.set('point', point.point);
  url.searchParams.set('qr', 'extintor');
  return url.toString();
}
function getV2EntryUrl(){
  if(BO_CONFIG?.qr?.entryUrl) return BO_CONFIG.qr.entryUrl;
  const configured = BO_CONFIG?.publicBaseUrl || BO_CONFIG?.publicUrl || '';
  if(configured){
    try{
      const u = new URL(configured, window.location.href);
      u.pathname = u.pathname.replace(/\/index\.html$/,'/V2/index.html').replace(/\/$/,'/V2/index.html');
      return u.toString();
    }catch(e){}
  }
  return window.location.href.replace(/backoffice\.html.*$/, 'V2/index.html');
}
function buildQrImageUrl(text){ const size = Number(BO_CONFIG?.qr?.imageSizePx || BO_CONFIG?.qr?.sizePx || 600); return `https://api.qrserver.com/v1/create-qr-code/?size=${size}x${size}&data=${encodeURIComponent(text)}`; }
function getPublicIndexUrl(){ return BO_CONFIG?.publicBaseUrl || BO_CONFIG?.publicUrl || window.location.href.replace(/backoffice\.html.*$/, 'index.html'); }
function getQrPrintLabel(point){
  const floor = Number(point.floor);
  if(floor >= 0) return `Piso ${floor}`;
  const ext = String(point.label || point.point || '').trim();
  return `Piso ${floor} · ${ext || 'Extintor'}`;
}
function getAllPoints(){ const floors = BO_CONFIG?.building?.floors || []; return floors.flatMap(f => (f.extinguishers || []).map(e => ({ floor:f.floor, floorLabel:f.label, point:e.point, label:e.label || e.point, shortLabel:e.shortLabel || e.point, location:e.location || '' }))); }
function getFloorLabel(floor){ const match = (BO_CONFIG?.building?.floors || []).find(f => Number(f.floor) === Number(floor)); return match?.label || String(floor); }
function extractReportedItems(payload){ if(!payload) return []; if(Array.isArray(payload)) return payload; return payload.reported || payload.items || payload.data || []; }
function normalizeStatusKey(item){ if(!item) return ''; if(typeof item === 'string') return item.trim().toUpperCase(); const floor = item.floor ?? item.piso ?? item.level ?? item.floorNumber ?? item.idFloor; const point = item.point ?? item.ponto ?? item.extinguisher ?? item.extintor ?? item.code ?? item.idPoint; if(floor === undefined || floor === null || !point) return String(item.id || item.key || '').trim().toUpperCase(); return makeKey(Number(floor), String(point)); }
async function apiGet(actionName, params = {}){ const base = BO_CONFIG?.backendUrl; if(!base || base.includes('COLOCAR_AQUI')) throw new Error('backendUrl não configurado.'); const url = new URL(base); if(actionName) url.searchParams.set('action', getAction(actionName)); Object.entries(params).forEach(([k,v]) => { if(v !== undefined && v !== null && v !== '') url.searchParams.set(k,v); }); return fetchJson(url.toString()); }
async function apiPost(payload){ const base = BO_CONFIG?.backendUrl; if(!base || base.includes('COLOCAR_AQUI')) throw new Error('backendUrl não configurado.'); const fd = new FormData(); Object.entries(payload).forEach(([k,v]) => { if(v !== undefined && v !== null) fd.append(k, String(v)); }); const res = await fetch(base, { method:'POST', body:fd }); const text = await res.text(); try{ return JSON.parse(text); }catch{ return { success:res.ok, raw:text }; } }
async function fetchJson(url){ const res = await fetch(url, { cache:'no-store' }); const text = await res.text(); if(!res.ok) throw new Error(`Erro HTTP ${res.status}`); try{ return JSON.parse(text); }catch{ throw new Error('Resposta não JSON.'); } }
function getAction(name){ return BO_CONFIG?.apiActions?.[name] || name; }
function makeKey(floor, point){ return `${Number(floor)}:${String(point).trim().toUpperCase()}`; }
async function fileToPayload(file){ const dataUrl = await new Promise((resolve,reject) => { const r = new FileReader(); r.onload = () => resolve(String(r.result)); r.onerror = () => reject(new Error('Não foi possível ler a fotografia.')); r.readAsDataURL(file); }); return { name:file.name, type:file.type, size:file.size, dataUrl, base64:dataUrl.split(',')[1] || '' }; }
function showBackofficeToast(msg){ bo.toast.textContent = msg; bo.toast.classList.add('show'); clearTimeout(showBackofficeToast._t); showBackofficeToast._t = setTimeout(() => bo.toast.classList.remove('show'), 3200); }
function formatDateTime(value){ const d = value instanceof Date ? value : new Date(value); if(Number.isNaN(d.getTime())) return String(value || '—'); return new Intl.DateTimeFormat('pt-PT', { dateStyle:'short', timeStyle:'short' }).format(d); }
function formatBytes(bytes){ const units = ['B','KB','MB','GB']; let v = bytes || 0, i = 0; while(v >= 1024 && i < units.length - 1){ v /= 1024; i++; } return `${v.toFixed(v >= 10 || i === 0 ? 0 : 1)} ${units[i]}`; }
function escapeHtml(value){ return String(value ?? '').replace(/[&<>'"]/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#039;','"':'&quot;'}[ch])); }
function escapeAttr(value){ return escapeHtml(value).replace(/`/g, '&#096;'); }
