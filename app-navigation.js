/* Navegação comum. Os destinos são explícitos: não dependem do histórico do navegador. */
(function () {
  'use strict';
  const root = new URL('.', document.currentScript.src);
  const page = location.pathname.slice(root.pathname.length) || 'index.html';
  function parentFor(page, screen, previous) {
    if (page === 'V2/index.html') {
      const parents = {
        'screen-login':'screen-start', 'screen-first-use':'screen-start',
        'screen-recover':['screen-start','screen-first-use','screen-login'].includes(previous) ? previous : 'screen-login',
        'screen-cadeado-motivo':'screen-home', 'screen-codigo':'screen-cadeado-motivo',
        'screen-extintores':'screen-home', 'screen-ok':'screen-start',
        'screen-pendente':'screen-start', 'screen-bloqueado':'screen-start'
      };
      return parents[screen] ? {screen:parents[screen]} : null;
    }
    if (page === 'V2/garagem.html') {
      const parents = {
        'screen-registo':'screen-inicio', 'screen-login':'screen-inicio',
        'screen-esqueci-pin':previous === 'screen-registo' ? previous : 'screen-login',
        'screen-codigo':'screen-motivo', 'screen-registo-ok':'screen-inicio',
        'screen-pendente':'screen-inicio', 'screen-rejeitado':'screen-inicio',
        'screen-bloqueado':'screen-inicio',
        'screen-contacto':['screen-bloqueado','screen-rejeitado'].includes(previous) ? previous : 'screen-inicio'
      };
      return parents[screen] ? {screen:parents[screen]} : {menu:true};
    }
    if (['V2/acessos-admin.html','V2/config-admin.html','V2/garagem-admin.html','backoffice.html','relatorio-extintores.html'].includes(page)) return {url:'V2/admin.html'};
    return {menu:true};
  }
  window.DC4Navigation = {parentFor};
  function mount() {
    if (document.getElementById('dc4-navigation')) return;
    const style = document.createElement('style');
    style.textContent = `
      body{padding-top:calc(56px + env(safe-area-inset-top,0px))!important}
      #dc4-navigation{position:fixed;top:0;left:0;right:0;z-index:900;box-sizing:border-box;height:calc(56px + env(safe-area-inset-top,0px));padding:calc(6px + env(safe-area-inset-top,0px)) 88px 6px max(12px,env(safe-area-inset-left,0px));background:#fff;border-bottom:1px solid #e2e8f0;display:flex;align-items:center}
      #dc4-back{display:inline-flex;gap:8px;align-items:center;justify-content:center;min-height:44px;min-width:104px;padding:8px 14px;border:1px solid #cbd5e1;border-radius:12px;background:#fff;color:#1a2f5a;font:700 14px/1.2 system-ui,sans-serif;cursor:pointer}
      #dc4-back:hover{background:#f1f5f9}#dc4-back:focus-visible{outline:3px solid #2563eb;outline-offset:2px}#dc4-back[hidden]{display:none}
      #dc4-version{top:calc(18px + env(safe-area-inset-top,0px))}
      @media print{#dc4-navigation{display:none}body{padding-top:0!important}}
    `;
    document.head.appendChild(style);
    const bar = document.createElement('nav');
    bar.id = 'dc4-navigation'; bar.setAttribute('aria-label','Navegação da app');
    const button = document.createElement('button');
    button.id = 'dc4-back'; button.type = 'button'; button.textContent = '← Voltar';
    bar.appendChild(button); document.body.prepend(bar);
    let lastScreen = document.querySelector('.screen.active')?.id || '';
    let previous = '';
    function current() { return document.querySelector('.screen.active')?.id || ''; }
    function update() {
      const screen = current();
      if (screen !== lastScreen) { previous = lastScreen; lastScreen = screen; }
      const target = page === 'qrcode-report.html' && new URLSearchParams(location.search).get('from') === 'map' ? {url:'index.html'} : parentFor(page,screen,previous);
      button.hidden = !target;
      button.title = target?.url === 'index.html' ? 'Voltar aos extintores' : target?.url ? 'Voltar à administração' : target?.menu ? 'Voltar ao menu da app' : 'Voltar ao ecrã anterior';
    }
    button.addEventListener('click', () => {
      const screen = current();
      const target = page === 'qrcode-report.html' && new URLSearchParams(location.search).get('from') === 'map' ? {url:'index.html'} : parentFor(page,screen,previous);
      if (!target) return;
      if (screen === 'screen-codigo' && typeof window.hideCode === 'function') window.hideCode();
      if (target.screen && typeof window.irPara === 'function') {
        window.irPara(target.screen); update();
        document.getElementById(target.screen)?.querySelector('input,button,a')?.focus({preventScroll:true});
      } else {
        const url = new URL(target.url || 'V2/index.html',root);
        if (target.menu) url.hash = 'menu';
        location.assign(url.href);
      }
    });
    // Regresso de um módulo ao menu, mantendo a sessão do morador.
    if (page === 'V2/index.html' && location.hash === '#menu') {
      if (window.dc4HasSession?.('resident')) {
        if (typeof window.restoreUser === 'function') window.restoreUser();
      }
      window.irPara?.('screen-home');
      history.replaceState(null,'',location.pathname + location.search);
    }
    document.querySelectorAll('.screen').forEach(screen => new MutationObserver(update).observe(screen,{attributes:true,attributeFilter:['class']}));
    update();
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded',mount,{once:true});
  else mount();
})();
