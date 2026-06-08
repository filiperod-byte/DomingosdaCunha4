// Remove o ecrã intermédio do módulo de extintores.
// Ao tocar em "Controlo de extintores", abre diretamente o mapa/fachada.
window.abrirExtintores = function abrirExtintoresDireto() {
  abrirReporteExtintor();
};

// Quando a entrada vem de QR code, o fluxo é diferente:
// abre diretamente o reporte do extintor e só pede o PIN no momento de submeter.
// Assim evita a sensação de login normal + redirecionamento inesperado.
(function () {
  const params = new URLSearchParams(window.location.search);
  const floor = params.get('floor');
  const point = params.get('point');
  const qr = params.get('qr');

  if (!floor || !point) return;

  const url = new URL('../qrcode-report.html', window.location.href);
  url.searchParams.set('floor', floor);
  url.searchParams.set('point', point);
  if (qr) url.searchParams.set('qr', qr);

  window.location.replace(url.toString());
})();
