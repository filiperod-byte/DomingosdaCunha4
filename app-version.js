/* Versão da interface: atualizar este valor e as referências ?v= em cada publicação. */
(function () {
  'use strict';
  const VERSION = '3.6.0-rc1';
  function mount() {
    if (document.getElementById('dc4-version')) return;
    const style = document.createElement('style');
    style.textContent = '#dc4-version{position:fixed;top:calc(env(safe-area-inset-top,0px) + 18px);right:8px;z-index:10000;padding:3px 8px;border:1px solid #cbd5e1;border-radius:12px;background:#fff;color:#334155;font:600 11px/1.4 system-ui,sans-serif;cursor:pointer;box-shadow:0 1px 4px #0001}#dc4-version:focus-visible{outline:2px solid #2563eb;outline-offset:2px}body{padding-top:calc(56px + env(safe-area-inset-top,0px))!important}@media print{#dc4-version{display:none}body{padding-top:0!important}}';
    document.head.appendChild(style);
    const badge = document.createElement('button');
    badge.id = 'dc4-version';
    badge.type = 'button';
    badge.textContent = 'v' + VERSION;
    badge.title = 'Versão da interface da app';
    badge.setAttribute('aria-label', 'Informação da versão ' + VERSION);
    badge.addEventListener('click', function () {
      window.alert('App do Condomínio — v' + VERSION + '\nPublicação: 15/09/2026\n\nIndica este número quando reportares um problema.\nEste indicador identifica a interface; não verifica a versão do serviço de dados.');
    });
    document.body.appendChild(badge);
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', mount, {once:true});
  else mount();
})();
