// Reduz fotografias no telemóvel antes do envio para o Apps Script.
// Este ficheiro também aplica ajustes visuais à fachada dos extintores.
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

// Fachada inspirada no prédio real.
(function applyFacadeOverrides() {
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

  window.renderBuilding = function renderBuilding() {
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
        const isAlert = REPORTED_SET.has(key);

        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = `ext-btn ${isAlert ? 'alert' : 'ok'}`;
        btn.dataset.floor = String(floor.floor);
        btn.dataset.point = ext.point;
        btn.dataset.location = ext.location || '';
        btn.dataset.label = ext.label || ext.point;
        btn.dataset.title = displayLabel;
        btn.setAttribute('aria-label', `${displayLabel} - ${ext.label || ext.point}${isAlert ? ' - reportado' : ' - sem alerta'}`);

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
            isAlert
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
})();
