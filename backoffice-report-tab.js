// V2.1 - adiciona separador Relatórios ao Backoffice dos Extintores.
(function () {
  function addReportTab() {
    const tabs = document.querySelector('.tabs');
    const adminSection = document.getElementById('adminSection');
    if (!tabs || !adminSection || document.querySelector('[data-tab="reports"]')) return;

    const btn = document.createElement('button');
    btn.className = 'tab-btn';
    btn.type = 'button';
    btn.dataset.tab = 'reports';
    btn.textContent = 'Relatórios';
    tabs.appendChild(btn);

    const panel = document.createElement('section');
    panel.id = 'tab-reports';
    panel.className = 'card tab-panel';
    panel.innerHTML = `
      <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:12px">
        <h3 class="section-title" style="margin:0">Relatório de extintores</h3>
        <a class="btn btn-primary" href="relatorio-extintores.html">Abrir relatório completo</a>
      </div>
      <div class="helper" style="margin-bottom:14px">
        Consulte as ocorrências abertas e reportes pendentes agrupados por tipo de anomalia, com filtros por piso, estado e pesquisa.
      </div>
      <div class="grid grid-5">
        <div class="stat"><div class="stat-label">Relatório</div><div class="stat-value" style="font-size:1.35rem">V2.1</div></div>
        <div class="stat"><div class="stat-label">Agrupamento</div><div class="stat-value" style="font-size:1.35rem">Tipo</div></div>
        <div class="stat"><div class="stat-label">Filtros</div><div class="stat-value" style="font-size:1.35rem">Sim</div></div>
        <div class="stat"><div class="stat-label">PDF</div><div class="stat-value" style="font-size:1.35rem">Print</div></div>
        <div class="stat"><div class="stat-label">Email</div><div class="stat-value" style="font-size:1.35rem">Texto</div></div>
      </div>
      <div class="list-card" style="margin-top:14px">
        <div class="kv">
          <div><strong>PDF:</strong> abre uma janela de impressão para guardar como PDF.</div>
          <div><strong>Email:</strong> prepara o envio com resumo em texto. O anexo automático fica para a próxima micro-versão do backend.</div>
        </div>
      </div>
    `;
    adminSection.appendChild(panel);

    btn.addEventListener('click', function () {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.toggle('active', b.dataset.tab === 'reports'));
      document.querySelectorAll('.tab-panel').forEach(p => p.classList.toggle('active', p.id === 'tab-reports'));
    });
  }

  document.addEventListener('DOMContentLoaded', function () {
    addReportTab();
    setTimeout(addReportTab, 300);
    setTimeout(addReportTab, 1000);
  });
})();
