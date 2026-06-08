// Carregamento robusto das configurações de administração.
(function () {
  function valueFrom(obj, keys, fallback) {
    for (const key of keys) {
      if (obj && obj[key] !== undefined && obj[key] !== null && String(obj[key]) !== '') return obj[key];
    }
    return fallback || '';
  }

  async function safeApiGet(action) {
    const payload = await apiGet(action);
    if (payload && payload.success === false) throw new Error(payload.msg || payload.message || 'Erro no backend.');
    return payload || {};
  }

  async function getAdminConfigRobusto() {
    let admin = {};
    let publicCfg = {};
    let dashboard = {};

    try { admin = await safeApiGet('garage.adminConfig'); } catch (err) { console.warn('adminConfig falhou', err); }
    try { publicCfg = await safeApiGet('garage.publicConfig'); } catch (err) { console.warn('publicConfig falhou', err); }
    try { dashboard = await safeApiGet('garage.dashboard'); } catch (err) { console.warn('dashboard falhou', err); }

    return {
      nomeCondominio: valueFrom(admin, ['nomeCondominio', 'NOME_CONDOMINIO'], valueFrom(publicCfg, ['nomeCondominio', 'name'], '')),
      adminEmail: valueFrom(admin, ['adminEmail', 'ADMIN_EMAIL'], valueFrom(publicCfg, ['adminEmail'], '')),
      codigoAtual: valueFrom(admin, ['codigoAtual', 'CODIGO_CADEADO', 'codigo'], valueFrom(dashboard, ['codigoAtual'], '')),
      tempoVisivel: valueFrom(admin, ['tempoVisivel', 'TEMPO_VISIVEL_SEGUNDOS'], valueFrom(publicCfg, ['tempoVisivel'], '15')),
      textoAviso: valueFrom(admin, ['textoAviso', 'TEXTO_AVISO'], '')
    };
  }

  window.getAdminConfigRobusto = getAdminConfigRobusto;

  if (document.getElementById('novo-codigo')) {
    window.loadDefinicoes = async function loadDefinicoesFix() {
      const msg = document.getElementById('codigo-msg');
      if (msg) msg.innerHTML = '<div class="msg info">A carregar configurações...</div>';
      try {
        const c = await getAdminConfigRobusto();
        document.getElementById('novo-codigo').value = c.codigoAtual || '';
        document.getElementById('cfg-tempo').value = c.tempoVisivel || '15';
        document.getElementById('cfg-aviso').value = c.textoAviso || '';
        if (msg) msg.innerHTML = '';
        if (!c.codigoAtual && msg) setMsg('codigo-msg', 'err', 'Não consegui carregar o código atual. Verifica o backend.');
      } catch (err) {
        if (msg) setMsg('codigo-msg', 'err', err.message || 'Erro ao carregar definições.');
        toast('Erro ao carregar definições');
      }
    };
  }

  if (document.getElementById('cfg-nome') && !document.getElementById('novo-codigo')) {
    window.loadDefinicoes = async function loadDefinicoesGeraisFix() {
      const msg = document.getElementById('cfg-msg');
      if (msg) msg.innerHTML = '<div class="msg info">A carregar configurações...</div>';
      try {
        const c = await getAdminConfigRobusto();
        document.getElementById('cfg-nome').value = c.nomeCondominio || '';
        document.getElementById('cfg-admin-email').value = c.adminEmail || '';
        if (msg) msg.innerHTML = '';
        if ((!c.nomeCondominio || !c.adminEmail) && msg) setMsg('cfg-msg', 'err', 'Alguns dados não foram carregados. Verifica o backend.');
      } catch (err) {
        if (msg) setMsg('cfg-msg', 'err', err.message || 'Erro ao carregar configurações.');
        toast('Erro ao carregar configurações');
      }
    };
  }
})();
