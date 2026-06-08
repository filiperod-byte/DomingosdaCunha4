// Reduz fotografias no telemóvel antes do envio para o Apps Script.
// Este ficheiro também aplica ajustes visuais e funcionais à fachada dos extintores.
window.fileToPayload = async function fileToPayload(file) {
  const originalSize = file.size;
  const compressed = await compressExtinguisherImage(file, {
    maxWidth: 1280,
    maxHeight: 1280,
    quality: 0.72,
    maxBytes: 450 * 1024
  });

  const dataUrl = compressed.dataUrl;
  const base64 = dataUrl.split(',')[1] || '';

  return {
    name: compressed.name,
    type: compressed.type,
    size: compressed.size,
    originalSize,
    dataUrl,
    base64
  };
};

async function compressExtinguisherImage(file, options) {
  if (!/^image\/(jpeg|jpg|png|webp)$/i.test(file.type)) {
    const dataUrl = await readFileAsDataURLForCompression(file);
    return {
      dataUrl,
      type: file.type || 'image/jpeg',
      name: file.name || 'foto.jpg',
      size: file.size
    };
  }

  const img = await loadImageForCompression(file);
  let width = img.width;
  let height = img.height;
  const ratio = Math.min(options.maxWidth / width, options.maxHeight / height, 1);
  width = Math.max(1, Math.round(width * ratio));
  height = Math.max(1, Math.round(height * ratio));

  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext('2d', { alpha: false });
  ctx.drawImage(img, 0, 0, width, height);

  let quality = options.quality;
  let dataUrl = canvas.toDataURL('image/jpeg', quality);
  let size = estimateDataUrlBytesForCompression(dataUrl);

  while (size > options.maxBytes && quality > 0.45) {
    quality = Math.max(0.45, quality - 0.08);
    dataUrl = canvas.toDataURL('image/jpeg', quality);
    size = estimateDataUrlBytesForCompression(dataUrl);
  }

  return {
    dataUrl,
    type: 'image/jpeg',
    name: String(file.name || 'foto').replace(/\.[^.]+$/, '') + '.jpg',
    size
  };
}

function loadImageForCompression(file) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => {
      URL.revokeObjectURL(url);
      resolve(img);
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error('Não foi possível preparar a fotografia.'));
    };
    img.src = url;
  });
}

function readFileAsDataURLForCompression(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(new Error('Não foi possível ler a fotografia.'));
    reader.readAsDataURL(file);
  });
}

function estimateDataUrlBytesForCompression(dataUrl) {
  const base64 = String(dataUrl).split(',')[1] || '';
  return Math.round(base64.length * 0.75);
}

// Fachada inspirada no prédio real + estados verde/amarelo/vermelho.
(function applyFacadeOverrides() {
  const occurrenceMap = new Map();
  const openSet = new Set();
  const pendingSet = new Set();

  const css = `
    .facade-shell:before{display:none!important}
    .facade-row.residential .facade-inner{grid-template-columns:32px 1fr auto}
    .facade-inner.no-stair{grid-template-columns:1fr auto!important}
    .facade-windows.cols-3{grid-template-columns:repeat(3,1fr)!important}
    .facade-row.residential .window{height:40px}
    .facade-row.service{padding-top:6px;padding-bottom:6px}
    .facade-row.service .floor-label{min-height:128px;background:linear-gradient(180deg,#EEF4FF,#DCE9FF);color:#1A2F5A}
    .facade-row.service .floor-body{background:linear-gradient(180deg,#F9FBFF,#EEF4FA)}
    .facade-row.service .facade-inner{min-height:128px;padding-top:12px;padding-bottom:12px}
    .storefront-zone{height:104px;display:grid;grid-template-columns:repeat(3,1fr);gap:10px;align-items:stretch}
    .storefront-window{border-radius:12px;border:1.5px solid #C4D1E3;background:linear-gradient(180deg,#FFFFFF 0%,#EFF4FA 52%,#DDE8F3 100%);position:relative;overflow:hidden;box-shadow:inset 0 1px 0 rgba(255,255,255,.9),0 1px 2px rgba(26,47,90,.05)}
    .storefront-window:before{content:"";position:absolute;left:50%;top:0;bottom:0;width:1.5px;background:#C4D1E3;transform:translateX(-50%)}
    .storefront-window:after{content:"";position:absolute;left:0;right:0;top:58%;height:1.5px;background:#C4D1E3}
    .facade-row.garage .floor-label{background:linear-gradient(180deg,#C98768,#9E563F)!important;border-color:#8E4A36!important;color:#fff!important;text-shadow:0 1px 0 rgba(0,0,0,.15)}
    .facade-row.garage .floor-body{background:linear-gradient(180deg,#B86A4C,#8F4935)!important;border-color:#9B5A43!important}
    .facade-row.garage .facade-inner{min-height:70px;grid-template-columns:1fr auto!important}
    .facade-row.garage .garage-zone{background:linear-gradient(180deg,#F4D7C4,#D99A78)!important;border:1px dashed rgba(86,45,31,.42)!important;color:#5B2F22!important;font-weight:800}
    .facade-row.garage .ext-btn{box-shadow:0 7px 16px rgba(73,34,20,.26)}
    .dot.pending{background:#F59E0B}
    .ext-btn.pending{background:linear-gradient(180deg,#FBBF24,#D97706)!important;color:#fff!important}
    .alert-box.pending{display:block;background:#FFFBEB;border-color:#FDE68A;color:#92400E}
    .alert-box.open{display:block;background:#FEF2F2;border-color:#FECACA;color:#991B1B}
    .existing-report-title{font-weight:900;margin-bottom:6px}
    .existing-report-grid{display:grid;gap:4px;font-size:.9rem;line-height:1.35}
    .existing-report-grid strong{font-weight:900}
    .existing-report-grid a{color:inherit;font-weight:900;text-decoration:underline}
    @media(max-width:420px){
      .facade-row.residential .facade-inner{grid-template-columns:24px 1fr auto}
      .facade-row.service .floor-label{min-height:112px}
      .facade-row.service .facade-inner{min-height:112px}
      .storefront-zone{height:90px;gap:7px}
      .facade-row.garage .facade-inner{grid-template-columns:1fr auto!important}
    }
    @media(min-width:768px){.facade-row.residential .facade-inner{grid-template-columns:42px 1fr auto}}
  `;

  const style = document.createElement('style');
  style.id = 'facade-overrides-style';
  style.textContent = css;
  document.head.appendChild(style);

  function escapeHtmlLocal(value) {
    return String(value || '').replace(/[&<>\"]/g, (char) => ({
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;'
    }[char]));
  }

  function displayDate(value) {
    if (!value) return '';
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return String(value);
    return new Intl.DateTimeFormat('pt-PT', {
      day: '2-digit',
      month: '2-digit',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit'
    }).format(date);
  }

  function asOccurrenceArray(payload) {
    if (!payload) return [];
    if (Array.isArray(payload)) return payload;
    if (Array.isArray(payload.occurrences)) return payload.occurrences;
    if (Array.isArray(payload.items)) return payload.items;
    if (Array.isArray(payload.data)) return payload.data;
    return [];
  }

  function clearOccurrenceState() {
    occurrenceMap.clear();
    openSet.clear();
    pendingSet.clear();
  }

  function registerOccurrence(item, type) {
    if (!item) return;
    const floor = Number(item.floor ?? item.piso ?? item.FLOOR);
    const point = String(item.point ?? item.ponto ?? item.POINT ?? '').trim();
    if (!point || Number.isNaN(floor)) return;

    const key = makeKey(floor, point);
    const detail = {
      type,
      id: item.id || item.occurrenceId || item.OCCURRENCE_ID || '',
      floor,
      point,
      floorLabel: item.floorLabel || item.FLOOR_LABEL || facadeFloorLabel(floor),
      location: item.location || item.LOCATION || '',
      reportedBy: item.reportedBy || item.REPORTED_BY || '',
      reason: item.reason || item.REASON || '',
      notes: item.notes || item.NOTES || '',
      createdAt: item.createdAt || item.reportedAt || item.REPORTED_AT || '',
      photoUrl: item.photoUrl || item.PHOTO_FILE_URL || '',
      status: item.status || item.STATUS || ''
    };

    occurrenceMap.set(key, detail);
    if (type === 'open') openSet.add(key);
    if (type === 'pending') pendingSet.add(key);
  }

  function updateLegendText() {
    const legend = document.querySelector('.legend');
    if (!legend) return;
    legend.innerHTML = `
      <span class="pill"><span class="dot ok"></span> Sem ocorrência</span>
      <span class="pill"><span class="dot pending"></span> A aguardar validação</span>
      <span class="pill"><span class="dot alert"></span> Ocorrência aberta</span>
    `;
  }

  function facadeFloorLabel(floorValue, labelValue) {
    const floorNumber = Number(floorValue);
    const label = String(labelValue ?? floorValue ?? '').trim();
    if (Number.isFinite(floorNumber) && floorNumber > 0) return `${floorNumber}º`;
    return label || String(floorValue ?? '');
  }

  function facadeKind(floor) {
    const floorNumber = Number(floor.floor);
    if (floor.type === 'garage' || floorNumber < 0) return 'garage';
    if (floor.type === 'service' || floorNumber === 0) return 'service';
    return 'residential';
  }

  function createStairTower() {
    const stair = document.createElement('div');
    stair.className = 'stair-tower';
    for (let i = 0; i < 3; i += 1) {
      stair.appendChild(document.createElement('span'));
    }
    return stair;
  }

  function createResidentialWindows(floor) {
    const floorNumber = Number(floor.floor);
    const windows = document.createElement('div');
    const totalWindows = floorNumber >= 2 && floorNumber <= 9
      ? 3
      : Math.max(2, Math.min(Number(floor.fractions || 3), 4));

    windows.className = `facade-windows cols-${totalWindows}`;
    for (let i = 0; i < totalWindows; i += 1) {
      const windowEl = document.createElement('div');
      const wide = !(floorNumber >= 2 && floorNumber <= 9) && i === 0 && totalWindows > 2;
      windowEl.className = `window facade-window ${wide ? 'wide' : ''}`;
      windows.appendChild(windowEl);
    }
    return windows;
  }

  function createStorefronts() {
    const stores = document.createElement('div');
    stores.className = 'storefront-zone';
    for (let i = 0; i < 3; i += 1) {
      const store = document.createElement('span');
      store.className = 'storefront-window';
      stores.appendChild(store);
    }
    return stores;
  }

  function createGarageZone() {
    const garage = document.createElement('div');
    garage.className = 'garage-zone facade-garage';
    garage.textContent = 'Garagem / espaço aberto';
    return garage;
  }

  loadStatuses = async function loadStatusesOverride() {
    clearOccurrenceState();
    updateLegendText();

    try {
      const [openResult, pendingResult] = await Promise.allSettled([
        apiGet('openOccurrences'),
        apiGet('pendingOccurrences')
      ]);

      if (openResult.status === 'fulfilled') {
        asOccurrenceArray(openResult.value).forEach((item) => registerOccurrence(item, 'open'));
      }
      if (pendingResult.status === 'fulfilled') {
        asOccurrenceArray(pendingResult.value).forEach((item) => registerOccurrence(item, 'pending'));
      }

      REPORTED_SET = new Set(openSet);
      els.lastRefresh.textContent = `Atualizado às ${formatTime(new Date())}`;

      if (openResult.status === 'rejected' && pendingResult.status === 'rejected') {
        throw openResult.reason || pendingResult.reason;
      }
    } catch (error) {
      console.error('Erro ao carregar estados:', error);
      clearOccurrenceState();
      REPORTED_SET = new Set();
      els.lastRefresh.textContent = 'Sem ligação ao backend';
      showToast('Não foi possível atualizar o estado. A app continua disponível para reporte.');
    }
  };

  renderBuilding = function renderBuildingOverride() {
    const floors = APP_CONFIG?.building?.floors || [];
    els.building.innerHTML = '';

    const facade = document.createElement('div');
    facade.className = 'facade-shell';

    const roof = document.createElement('div');
    roof.className = 'facade-roof';
    roof.innerHTML = '<span></span><span></span><span></span>';
    facade.appendChild(roof);

    floors.forEach((floor) => {
      const kind = facadeKind(floor);
      const displayLabel = facadeFloorLabel(floor.floor, floor.label);

      const row = document.createElement('div');
      row.className = `facade-row ${kind}`;

      const label = document.createElement('div');
      label.className = 'floor-label';
      label.textContent = displayLabel;

      const body = document.createElement('div');
      body.className = 'floor-body';

      const inner = document.createElement('div');
      inner.className = `facade-inner ${kind === 'residential' ? 'has-stair' : 'no-stair'}`;

      const facadeContent = document.createElement('div');
      facadeContent.className = `facade-content ${kind}`;

      if (kind === 'residential') {
        inner.appendChild(createStairTower());
        facadeContent.appendChild(createResidentialWindows(floor));
      } else if (kind === 'service') {
        facadeContent.appendChild(createStorefronts());
      } else {
        facadeContent.appendChild(createGarageZone());
      }

      const exts = document.createElement('div');
      exts.className = 'extinguishers';

      (floor.extinguishers || []).forEach((ext) => {
        const key = makeKey(floor.floor, ext.point);
        const detail = occurrenceMap.get(key) || null;
        const stateClass = detail?.type === 'pending' ? 'pending' : detail?.type === 'open' ? 'alert' : 'ok';
        const stateText = detail?.type === 'pending' ? ' - a aguardar validação' : detail?.type === 'open' ? ' - ocorrência aberta' : ' - sem ocorrência';

        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = `ext-btn ${stateClass}`;
        btn.dataset.floor = String(floor.floor);
        btn.dataset.point = ext.point;
        btn.dataset.location = ext.location || '';
        btn.dataset.label = ext.label || ext.point;
        btn.dataset.title = displayLabel;
        btn.setAttribute('aria-label', `${displayLabel} - ${ext.label || ext.point}${stateText}`);

        const span = document.createElement('span');
        span.textContent = ext.shortLabel || ext.label || ext.point;
        btn.appendChild(span);

        btn.addEventListener('click', () => {
          openModal({
            floor: floor.floor,
            floorLabel: displayLabel,
            point: ext.point,
            label: ext.label || ext.point,
            shortLabel: ext.shortLabel || ext.point,
            location: ext.location || '',
            isAlert: detail?.type === 'open',
            isPending: detail?.type === 'pending',
            existingOccurrence: detail
          });
        });

        exts.appendChild(btn);
      });

      inner.appendChild(facadeContent);
      inner.appendChild(exts);
      body.appendChild(inner);
      row.appendChild(label);
      row.appendChild(body);
      facade.appendChild(row);
    });

    els.building.appendChild(facade);
    els.building.classList.remove('hidden');
  };

  openModal = function openModalOverride(ext) {
    SELECTED_POINT = ext;
    SELECTED_FILE = null;
    clearPhotoInputs();

    const targetLabel = `${ext.floorLabel} - ${ext.label}`;
    els.modalTitle.textContent = 'Reportar extintor';
    els.modalSubtitle.textContent = AUTO_REPORTER_NAME
      ? `Ponto selecionado: ${targetLabel} · Reportado por: ${AUTO_REPORTER_NAME}`
      : `Ponto selecionado: ${targetLabel}`;

    const existing = ext.existingOccurrence;
    els.alreadyReportedBox.className = 'alert-box';

    if (existing) {
      const typeLabel = existing.type === 'pending'
        ? 'Já existe um reporte neste extintor a aguardar validação.'
        : 'Já existe uma ocorrência aberta neste extintor.';

      els.alreadyReportedBox.classList.add(existing.type === 'pending' ? 'pending' : 'open');
      els.alreadyReportedBox.classList.add('show');
      els.alreadyReportedBox.innerHTML = `
        <div class="existing-report-title">${escapeHtmlLocal(typeLabel)}</div>
        <div class="existing-report-grid">
          ${existing.createdAt ? `<div><strong>Data:</strong> ${escapeHtmlLocal(displayDate(existing.createdAt))}</div>` : ''}
          ${existing.reportedBy ? `<div><strong>Reportado por:</strong> ${escapeHtmlLocal(existing.reportedBy)}</div>` : ''}
          ${existing.reason ? `<div><strong>Motivo:</strong> ${escapeHtmlLocal(existing.reason)}</div>` : ''}
          ${existing.notes ? `<div><strong>Observação:</strong> ${escapeHtmlLocal(existing.notes)}</div>` : ''}
          ${existing.photoUrl ? `<div><strong>Foto:</strong> <a href="${escapeHtmlLocal(existing.photoUrl)}" target="_blank" rel="noopener noreferrer">Abrir fotografia</a></div>` : ''}
        </div>
      `;
    } else {
      els.alreadyReportedBox.classList.remove('show');
      els.alreadyReportedBox.innerHTML = 'Este extintor já tem uma ocorrência aberta. Pode submeter informação adicional, se necessário.';
    }

    els.hiddenFloor.value = String(ext.floor);
    els.hiddenPoint.value = ext.point;
    els.hiddenLocation.value = ext.location || '';
    els.reporterName.value = AUTO_REPORTER_NAME || els.reporterName.value || '';

    els.overlay.classList.add('show');
    els.overlay.setAttribute('aria-hidden', 'false');
    document.body.style.overflow = 'hidden';

    window.setTimeout(() => {
      if (AUTO_REPORTER_NAME) els.reportReason.focus();
      else els.reporterName.focus();
    }, 40);
  };

  document.addEventListener('DOMContentLoaded', updateLegendText);
})();
