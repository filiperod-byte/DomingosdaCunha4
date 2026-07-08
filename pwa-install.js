// V3 - Instalação como app (PWA)
(function () {
  let deferredPrompt = null;

  function isStandalone() {
    return window.matchMedia('(display-mode: standalone)').matches || window.navigator.standalone === true;
  }

  function ensureInstallButton() {
    if (isStandalone() || document.getElementById('dc4-install-btn')) return;

    const btn = document.createElement('button');
    btn.id = 'dc4-install-btn';
    btn.type = 'button';
    btn.textContent = 'Instalar app';
    btn.style.cssText = [
      'position:fixed',
      'right:14px',
      'bottom:14px',
      'z-index:9999',
      'border:0',
      'border-radius:16px',
      'padding:12px 14px',
      'font:600 14px system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif',
      'background:#1A2F5A',
      'color:#fff',
      'box-shadow:0 8px 24px rgba(26,47,90,.24)',
      'cursor:pointer'
    ].join(';');

    btn.addEventListener('click', async function () {
      if (deferredPrompt) {
        deferredPrompt.prompt();
        await deferredPrompt.userChoice.catch(() => null);
        deferredPrompt = null;
        btn.remove();
      } else {
        alert('Para instalar: abra o menu do browser e escolha “Instalar app” ou “Adicionar ao ecrã principal”.');
      }
    });

    document.body.appendChild(btn);
  }

  window.addEventListener('beforeinstallprompt', function (event) {
    event.preventDefault();
    deferredPrompt = event;
    ensureInstallButton();
  });

  window.addEventListener('appinstalled', function () {
    const btn = document.getElementById('dc4-install-btn');
    if (btn) btn.remove();
  });

  if ('serviceWorker' in navigator) {
    window.addEventListener('load', function () {
      navigator.serviceWorker.register('/DomingosdaCunha4/service-worker.js', { scope: '/DomingosdaCunha4/' }).catch(console.warn);
    });
  }

  document.addEventListener('DOMContentLoaded', function () {
    setTimeout(ensureInstallButton, 1200);
  });
})();
