// Compatibilidade com QR codes antigos que apontam para o mapa ou para a entrada da app.
(function () {
  const root = new URL('.', document.currentScript.src);
  const params = new URLSearchParams(location.search);
  if (!params.has('floor') || !params.get('floor').trim() || !params.get('point')?.trim()) return;
  const target = new URL('qrcode-report.html', root);
  target.searchParams.set('floor', params.get('floor'));
  target.searchParams.set('point', params.get('point'));
  target.searchParams.set('qr', 'extintor');
  location.replace(target.href);
})();
