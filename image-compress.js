// Reduz fotografias no telemóvel antes do envio para o Apps Script.
// Este ficheiro substitui a função fileToPayload definida em app.js.
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
