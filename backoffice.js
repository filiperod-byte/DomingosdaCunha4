let BO_CONFIG = null;
let BO_OPEN_OCCURRENCES = [];
let BO_REPORTED_SET = new Set();
let BO_CLOSE_FILE = null;

const bo = {};

document.addEventListener("DOMContentLoaded", initBackoffice);

async function initBackoffice() {
  cacheBackofficeElements();
  bindBackofficeEvents();

  try {
    BO_CONFIG = await loadBackofficeConfig();
    renderStaticConfigInfo();
    await bootPinFlow();
  } catch (error) {
    console.error(error);
    showBackofficeToast("Erro ao carregar o back office. Verifica o config.json.");
  }
}

function cacheBackofficeElements() {
  bo.pinStatusBadge = document.getElementById("pinStatusBadge");
  bo.backendBadge = document.getElementById("backendBadge");

  bo.authSection = document.getElementById("authSection");
  bo.setupCard = document.getElementById("setupCard");
  bo.loginCard = document.getElementById("loginCard");
  bo.adminSection = document.getElementById("adminSection");

  bo.setupPin = document.getElementById("setupPin");
  bo.setupPinConfirm = document.getElementById("setupPinConfirm");
  bo.setupPinBtn = document.getElementById("setupPinBtn");

  bo.loginPin = document.getElementById("loginPin");
  bo.loginBtn = document.getElementById("loginBtn");
  bo.recoverPinBtn = document.getElementById("recoverPinBtn");

  bo.tabButtons = Array.from(document.querySelectorAll(".tab-btn"));
  bo.tabPanels = Array.from(document.querySelectorAll(".tab-panel"));

  bo.statTotal = document.getElementById("statTotal");
  bo.statOk = document.getElementById("statOk");
  bo.statAlert = document.getElementById("statAlert");
  bo.statOpen = document.getElementById("statOpen");

  bo.statusGrid = document.getElementById("statusGrid");
  bo.dashboardBackendInfo = document.getElementById("dashboardBackendInfo");
  bo.dashboardRefresh = document.getElementById("dashboardRefresh");
  bo.dashboardAdminEmail = document.getElementById("dashboardAdminEmail");
  bo.refreshDashboardBtn = document.getElementById("refreshDashboardBtn");

  bo.occurrencesList = document.getElementById("occurrencesList");
  bo.refreshOccurrencesBtn = document.getElementById("refreshOccurrencesBtn");

  bo.closeOccurrenceSelect = document.getElementById("closeOccurrenceSelect");
  bo.closeOccurrenceSummary = document.getElementById("closeOccurrenceSummary");
  bo.closePhotoInput = document.getElementById("closePhotoInput");
  bo.pickClosePhotoBtn = document.getElementById("pickClosePhotoBtn");
  bo.closePhotoMeta = document.getElementById("closePhotoMeta");
  bo.closeNote = document.getElementById("closeNote");
  bo.closeOccurrenceBtn = document.getElementById("closeOccurrenceBtn");

  bo.singleQrSelect = document.getElementById("singleQrSelect");
  bo.singleQrPreview = document.getElementById("singleQrPreview");
  bo.qrGrid = document.getElementById("qrGrid");
  bo.printSingleQrBtn = document.getElementById("printSingleQrBtn");
  bo.printAllQrBtn = document.getElementById("printAllQrBtn");

  bo.configBackendUrl = document.getElementById("configBackendUrl");
  bo.configAdminEmail = document.getElementById("configAdminEmail");
  bo.configPublicUrl = document.getElementById("configPublicUrl");
  bo.configQrSize = document.getElementById("configQrSize");
  bo.configRecoverPinBtn = document.getElementById("configRecoverPinBtn");
  bo.logoutBtn = document.getElementById("logoutBtn");

  bo.toast = document.getElementById("toast");
}

function bindBackofficeEvents() {
  bo.setupPinBtn.addEventListener("click", handleSetupPin);
  bo.loginBtn.addEventListener("click", handleLogin);
  bo.recoverPinBtn.addEventListener("click", handleRecoverPin);

  bo.tabButtons.forEach((btn) => {
    btn.addEventListener("click", () => activateTab(btn.dataset.tab));
  });

  bo.refreshDashboardBtn.addEventListener("click", refreshDashboard);
  bo.refreshOccurrencesBtn.addEventListener("click", loadOccurrences);

  bo.pickClosePhotoBtn.addEventListener("click", () => bo.closePhotoInput.click());
  bo.closePhotoInput.addEventListener("change", handleClosePhotoPicked);
  bo.closeOccurrenceSelect.addEventListener("change", renderCloseOccurrenceSummary);
  bo.closeOccurrenceBtn.addEventListener("click", handleCloseOccurrence);

  bo.singleQrSelect.addEventListener("change", renderSingleQrPreview);
  bo.printSingleQrBtn.addEventListener("click", printSingleQr);
  bo.printAllQrBtn.addEventListener("click", printAllQrs);

  bo.configRecoverPinBtn.addEventListener("click", handleRecoverPin);
  bo.logoutBtn.addEventListener("click", logoutBackoffice);

  bo.loginPin.addEventListener("keydown", (event) => {
    if (event.key === "Enter") handleLogin();
  });

  bo.setupPinConfirm.addEventListener("keydown", (event) => {
    if (event.key === "Enter") handleSetupPin();
  });
}

async function loadBackofficeConfig() {
  const res = await fetch("config.json", { cache: "no-store" });
  if (!res.ok) {
    throw new Error("Não foi possível carregar o config.json");
  }
  return res.json();
}

function renderStaticConfigInfo() {
  const backendConfigured =
    !!BO_CONFIG?.backendUrl && !BO_CONFIG.backendUrl.includes("COLOCAR_AQUI");

  bo.backendBadge.innerHTML = backendConfigured
    ? `<span class="dot ok"></span> Backend configurado`
    : `<span class="dot danger"></span> Backend por configurar`;

  bo.configBackendUrl.textContent = BO_CONFIG?.backendUrl || "—";
  bo.configAdminEmail.textContent = BO_CONFIG?.adminEmail || "—";
  bo.configPublicUrl.textContent = getPublicIndexUrl();
  bo.configQrSize.textContent = `${BO_CONFIG?.qr?.printSizeCm || 8} cm`;
  bo.dashboardBackendInfo.textContent = backendConfigured ? "Configurado" : "Por configurar";
  bo.dashboardAdminEmail.textContent = BO_CONFIG?.adminEmail || "—";
}

async function bootPinFlow() {
  sessionStorage.removeItem("extintores_bo_boot_error");

  try {
    const pinInfo = await apiGet("pinStatus");

    if (pinInfo?.pinConfigured) {
      bo.pinStatusBadge.innerHTML = `<span class="dot ok"></span> PIN definido`;
      showLoginCard();

      const unlocked = sessionStorage.getItem("extintores_bo_unlocked") === "1";
      if (unlocked) {
        unlockAdminArea();
      }
    } else {
      bo.pinStatusBadge.innerHTML = `<span class="dot warn"></span> PIN ainda não definido`;
      showSetupCard();
    }
  } catch (error) {
    console.error(error);
    bo.pinStatusBadge.innerHTML = `<span class="dot danger"></span> Erro no PIN / backend`;
    showLoginCard();
    showBackofficeToast("Não foi possível validar o estado do PIN no backend.");
  }
}

function showSetupCard() {
  bo.authSection.classList.remove("hidden");
  bo.setupCard.classList.remove("hidden");
  bo.loginCard.classList.add("hidden");
  bo.adminSection.classList.add("hidden");
}

function showLoginCard() {
  bo.authSection.classList.remove("hidden");
  bo.loginCard.classList.remove("hidden");
  bo.setupCard.classList.add("hidden");
  bo.adminSection.classList.add("hidden");
}

async function handleSetupPin() {
  const pin = bo.setupPin.value.trim();
  const confirmPin = bo.setupPinConfirm.value.trim();

  if (pin.length < 4) {
    showBackofficeToast("O PIN deve ter pelo menos 4 dígitos.");
    bo.setupPin.focus();
    return;
  }

  if (pin !== confirmPin) {
    showBackofficeToast("Os PINs não coincidem.");
    bo.setupPinConfirm.focus();
    return;
  }

  try {
    bo.setupPinBtn.disabled = true;
    bo.setupPinBtn.textContent = "A guardar…";

    const response = await apiPost({
      action: getAction("setPin"),
      pin
    });

    if (response?.success === false) {
      throw new Error(response.message || "Não foi possível guardar o PIN.");
    }

    showBackofficeToast("PIN definido com sucesso.");
    bo.pinStatusBadge.innerHTML = `<span class="dot ok"></span> PIN definido`;
    bo.setupPin.value = "";
    bo.setupPinConfirm.value = "";
    sessionStorage.setItem("extintores_bo_unlocked", "1");
    unlockAdminArea();
  } catch (error) {
    console.error(error);
    showBackofficeToast(error.message || "Falha ao guardar o PIN.");
  } finally {
    bo.setupPinBtn.disabled = false;
    bo.setupPinBtn.textContent = "Guardar PIN";
  }
}

async function handleLogin() {
  const pin = bo.loginPin.value.trim();

  if (!pin) {
    showBackofficeToast("Introduz o PIN.");
    bo.loginPin.focus();
    return;
  }

  try {
    bo.loginBtn.disabled = true;
    bo.loginBtn.textContent = "A validar…";

    const response = await apiPost({
      action: getAction("validatePin"),
      pin
    });

    if (response?.success === false || response?.valid === false) {
      throw new Error(response.message || "PIN inválido.");
    }

    sessionStorage.setItem("extintores_bo_unlocked", "1");
    bo.loginPin.value = "";
    unlockAdminArea();
    showBackofficeToast("Acesso autorizado.");
  } catch (error) {
    console.error(error);
    showBackofficeToast(error.message || "PIN inválido.");
  } finally {
    bo.loginBtn.disabled = false;
    bo.loginBtn.textContent = "Entrar";
  }
}

async function handleRecoverPin() {
  const ok = window.confirm(
    "Isto vai enviar email ao administrador e fazer reset ao PIN atual. Continuar?"
  );
  if (!ok) return;

  try {
    const response = await apiPost({
      action: getAction("resetPin"),
      adminEmail: BO_CONFIG?.adminEmail || ""
    });

    if (response?.success === false) {
      throw new Error(response.message || "Não foi possível recuperar/resetar o PIN.");
    }

    sessionStorage.removeItem("extintores_bo_unlocked");
    bo.pinStatusBadge.innerHTML = `<span class="dot warn"></span> PIN resetado`;
    showBackofficeToast("Pedido de recuperação enviado. O PIN foi resetado.");
    showSetupCard();
  } catch (error) {
    console.error(error);
    showBackofficeToast(error.message || "Falha ao recuperar/resetar PIN.");
  }
}

function unlockAdminArea() {
  bo.authSection.classList.add("hidden");
  bo.adminSection.classList.remove("hidden");
  activateTab("dashboard");
  refreshAllAdminData();
}

function logoutBackoffice() {
  sessionStorage.removeItem("extintores_bo_unlocked");
  showBackofficeToast("Sessão terminada.");
  showLoginCard();
}

function activateTab(tabName) {
  bo.tabButtons.forEach((btn) => {
    btn.classList.toggle("active", btn.dataset.tab === tabName);
  });

  bo.tabPanels.forEach((panel) => {
    panel.classList.toggle("active", panel.id === `tab-${tabName}`);
  });
}

async function refreshAllAdminData() {
  renderQrSection();
  await Promise.all([refreshDashboard(), loadOccurrences()]);
}

async function refreshDashboard() {
  try {
    const response = await apiGet("status");
    const items = extractReportedItems(response);
    BO_REPORTED_SET = new Set(items.map(normalizeStatusKey).filter(Boolean));

    const totalExtinguishers = getAllPoints().length;
    const alertCount = BO_REPORTED_SET.size;
    const okCount = Math.max(totalExtinguishers - alertCount, 0);

    bo.statTotal.textContent = String(totalExtinguishers);
    bo.statOk.textContent = String(okCount);
    bo.statAlert.textContent = String(alertCount);
    bo.statOpen.textContent = String(BO_OPEN_OCCURRENCES.length);

    renderStatusGrid();
    bo.dashboardRefresh.textContent = formatDateTime(new Date());
  } catch (error) {
    console.error(error);
    bo.statusGrid.innerHTML = `<div class="empty">Não foi possível obter o estado atual.</div>`;
    showBackofficeToast("Erro ao atualizar dashboard.");
  }
}

function renderStatusGrid() {
  const points = getAllPoints();
  bo.statusGrid.innerHTML = "";

  if (!points.length) {
    bo.statusGrid.innerHTML = `<div class="empty">Sem pontos configurados.</div>`;
    return;
  }

  points.forEach((point) => {
    const key = makeKey(point.floor, point.point);
    const alert = BO_REPORTED_SET.has(key);

    const div = document.createElement("div");
    div.className = `status-pill ${alert ? "alert" : ""}`;
    div.innerHTML = `
      <div>
        <strong>${escapeHtml(point.floorLabel)} · ${escapeHtml(point.label)}</strong>
        <div class="tiny muted">${escapeHtml(point.location || "Sem localização")}</div>
      </div>
      <div class="badge">
        <span class="dot ${alert ? "danger" : "ok"}"></span>
        ${alert ? "Reportado" : "OK"}
      </div>
    `;
    bo.statusGrid.appendChild(div);
  });
}

async function loadOccurrences() {
  bo.occurrencesList.innerHTML = `
    <div class="loading">
      <div class="spinner"></div>
      <span>A carregar ocorrências…</span>
    </div>
  `;

  try {
    const response = await apiGet("openOccurrences");
    BO_OPEN_OCCURRENCES = extractOccurrences(response).map(normalizeOccurrence);

    renderOccurrencesList();
    populateCloseOccurrenceSelect();
    renderCloseOccurrenceSummary();

    bo.statOpen.textContent = String(BO_OPEN_OCCURRENCES.length);
  } catch (error) {
    console.error(error);
    BO_OPEN_OCCURRENCES = [];
    renderOccurrencesList();
    populateCloseOccurrenceSelect();
    renderCloseOccurrenceSummary();
    showBackofficeToast("Não foi possível carregar as ocorrências.");
  }
}

function extractOccurrences(payload) {
  if (!payload) return [];
  if (Array.isArray(payload)) return payload;
  if (Array.isArray(payload.occurrences)) return payload.occurrences;
  if (Array.isArray(payload.items)) return payload.items;
  if (Array.isArray(payload.data)) return payload.data;
  return [];
}

function normalizeOccurrence(item, index) {
  const floor = Number(
    item.floor ??
      item.piso ??
      item.floorNumber ??
      item.level ??
      0
  );

  const point = String(
    item.point ??
      item.ponto ??
      item.extinguisher ??
      item.extintor ??
      item.code ??
      ""
  ).trim();

  const id =
    item.id ??
    item.occurrenceId ??
    item.ocorrenciaId ??
    `${makeKey(floor, point)}#${index + 1}`;

  return {
    id: String(id),
    floor,
    floorLabel: getFloorLabel(floor),
    point,
    label: item.label || point,
    location:
      item.location ??
      item.localizacao ??
      item.localização ??
      "",
    reportedBy:
      item.reportedBy ??
      item.name ??
      item.nome ??
      "",
    reason:
      item.reason ??
      item.motivo ??
      "",
    notes:
      item.notes ??
      item.observacao ??
      item.observação ??
      "",
    createdAt:
      item.createdAt ??
      item.timestamp ??
      item.data ??
      item.created ??
      "",
    photoUrl:
      item.photoUrl ??
      item.fotoUrl ??
      "",
    raw: item
  };
}

function renderOccurrencesList() {
  bo.occurrencesList.innerHTML = "";

  if (!BO_OPEN_OCCURRENCES.length) {
    bo.occurrencesList.innerHTML = `<div class="empty">Não existem ocorrências abertas neste momento.</div>`;
    return;
  }

  BO_OPEN_OCCURRENCES.forEach((occ) => {
    const card = document.createElement("div");
    card.className = "list-card";

    const photoLink = occ.photoUrl
      ? `<div><strong>Foto:</strong> <a href="${escapeAttr(occ.photoUrl)}" target="_blank" rel="noopener noreferrer">Abrir</a></div>`
      : `<div><strong>Foto:</strong> —</div>`;

    card.innerHTML = `
      <div class="list-head">
        <div>
          <strong>${escapeHtml(occ.floorLabel)} · ${escapeHtml(occ.point || "Sem ponto")}</strong>
          <div class="tiny muted">${escapeHtml(occ.location || "Sem localização")}</div>
        </div>
        <span class="badge"><span class="dot danger"></span> Aberta</span>
      </div>

      <div class="kv">
        <div><strong>ID:</strong> ${escapeHtml(occ.id)}</div>
        <div><strong>Reportado por:</strong> ${escapeHtml(occ.reportedBy || "—")}</div>
        <div><strong>Motivo:</strong> ${escapeHtml(occ.reason || "—")}</div>
        <div><strong>Observação:</strong> ${escapeHtml(occ.notes || "—")}</div>
        <div><strong>Data:</strong> ${escapeHtml(occ.createdAt || "—")}</div>
        ${photoLink}
      </div>
    `;

    bo.occurrencesList.appendChild(card);
  });
}

function populateCloseOccurrenceSelect() {
  bo.closeOccurrenceSelect.innerHTML = `<option value="">Selecionar ocorrência aberta</option>`;

  BO_OPEN_OCCURRENCES.forEach((occ) => {
    const option = document.createElement("option");
    option.value = occ.id;
    option.textContent = `${occ.floorLabel} · ${occ.point} · ${occ.reason || "Sem motivo"}`;
    bo.closeOccurrenceSelect.appendChild(option);
  });
}

function renderCloseOccurrenceSummary() {
  const selectedId = bo.closeOccurrenceSelect.value;
  const occ = BO_OPEN_OCCURRENCES.find((item) => item.id === selectedId);

  if (!occ) {
    bo.closeOccurrenceSummary.innerHTML = `Seleciona uma ocorrência para ver o detalhe.`;
    return;
  }

  bo.closeOccurrenceSummary.innerHTML = `
    <div class="kv">
      <div><strong>ID:</strong> ${escapeHtml(occ.id)}</div>
      <div><strong>Ponto:</strong> ${escapeHtml(occ.floorLabel)} · ${escapeHtml(occ.point)}</div>
      <div><strong>Localização:</strong> ${escapeHtml(occ.location || "—")}</div>
      <div><strong>Motivo:</strong> ${escapeHtml(occ.reason || "—")}</div>
      <div><strong>Reportado por:</strong> ${escapeHtml(occ.reportedBy || "—")}</div>
      <div><strong>Data:</strong> ${escapeHtml(occ.createdAt || "—")}</div>
    </div>
  `;
}

function handleClosePhotoPicked(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  if (!file.type.startsWith("image/")) {
    showBackofficeToast("Seleciona uma imagem válida.");
    event.target.value = "";
    return;
  }

  if (file.size > 8 * 1024 * 1024) {
    showBackofficeToast("A fotografia é demasiado grande. Usa uma imagem até 8 MB.");
    event.target.value = "";
    return;
  }

  BO_CLOSE_FILE = file;
  bo.closePhotoMeta.textContent = `${file.name} · ${formatBytes(file.size)}`;
}

async function handleCloseOccurrence() {
  const occurrenceId = bo.closeOccurrenceSelect.value;
  const occ = BO_OPEN_OCCURRENCES.find((item) => item.id === occurrenceId);

  if (!occ) {
    showBackofficeToast("Seleciona uma ocorrência aberta.");
    return;
  }

  const closePhotoRequired = BO_CONFIG?.features?.closePhotoRequired !== false;
  if (closePhotoRequired && !BO_CLOSE_FILE) {
    showBackofficeToast("A fotografia de fecho é obrigatória.");
    return;
  }

  try {
    bo.closeOccurrenceBtn.disabled = true;
    bo.closeOccurrenceBtn.textContent = "A fechar…";

    const photoPayload = BO_CLOSE_FILE ? await fileToPayload(BO_CLOSE_FILE) : null;

    const response = await apiPost({
      action: getAction("closeOccurrence"),
      occurrenceId: occ.id,
      floor: occ.floor,
      point: occ.point,
      location: occ.location,
      closeNotes: bo.closeNote.value.trim(),
      closePhotoBase64: photoPayload?.base64 || "",
      closePhotoDataUrl: photoPayload?.dataUrl || "",
      closePhotoName: photoPayload?.name || "",
      closePhotoType: photoPayload?.type || "",
      clientTs: new Date().toISOString(),
      source: "github-pages-backoffice"
    });

    if (response?.success === false) {
      throw new Error(response.message || "Falha ao fechar ocorrência.");
    }

    showBackofficeToast("Ocorrência fechada com sucesso.");
    bo.closeOccurrenceSelect.value = "";
    bo.closeNote.value = "";
    bo.closePhotoInput.value = "";
    bo.closePhotoMeta.textContent = "Nenhuma fotografia selecionada.";
    BO_CLOSE_FILE = null;

    await loadOccurrences();
    await refreshDashboard();
  } catch (error) {
    console.error(error);
    showBackofficeToast(error.message || "Não foi possível fechar a ocorrência.");
  } finally {
    bo.closeOccurrenceBtn.disabled = false;
    bo.closeOccurrenceBtn.textContent = "Fechar ocorrência";
  }
}

function renderQrSection() {
  const points = getAllPoints();

  bo.singleQrSelect.innerHTML = `<option value="">Selecionar ponto</option>`;
  bo.qrGrid.innerHTML = "";

  points.forEach((point) => {
    const option = document.createElement("option");
    option.value = makeKey(point.floor, point.point);
    option.textContent = `${point.floorLabel} · ${point.label}`;
    bo.singleQrSelect.appendChild(option);

    bo.qrGrid.appendChild(createQrCard(point));
  });

  renderSingleQrPreview();
}

function renderSingleQrPreview() {
  const selectedKey = bo.singleQrSelect.value;

  if (!selectedKey) {
    bo.singleQrPreview.classList.add("hidden");
    bo.singleQrPreview.innerHTML = "";
    return;
  }

  const point = getAllPoints().find(
    (item) => makeKey(item.floor, item.point) === selectedKey
  );

  if (!point) {
    bo.singleQrPreview.classList.add("hidden");
    bo.singleQrPreview.innerHTML = "";
    return;
  }

  const reportUrl = buildReportUrl(point);
  const imgUrl = buildQrImageUrl(reportUrl);

  bo.singleQrPreview.classList.remove("hidden");
  bo.singleQrPreview.innerHTML = `
    <div style="display:flex; gap:14px; flex-wrap:wrap; align-items:center;">
      <div class="qr-card" style="width:220px; min-height:auto;">
        <img src="${escapeAttr(imgUrl)}" alt="QR Code ${escapeAttr(point.label)}" />
        <div class="qr-label">${escapeHtml(point.floorLabel)}<br>${escapeHtml(point.label)}</div>
      </div>
      <div style="flex:1; min-width:240px;">
        <div class="kv">
          <div><strong>Ponto:</strong> ${escapeHtml(point.floorLabel)} · ${escapeHtml(point.label)}</div>
          <div><strong>Localização:</strong> ${escapeHtml(point.location || "—")}</div>
          <div><strong>URL:</strong></div>
        </div>
        <div class="code-box" style="margin-top:10px;">${escapeHtml(reportUrl)}</div>
      </div>
    </div>
  `;
}

function createQrCard(point) {
  const card = document.createElement("div");
  card.className = "qr-card";

  const reportUrl = buildReportUrl(point);
  const imgUrl = buildQrImageUrl(reportUrl);

  card.innerHTML = `
    <img src="${escapeAttr(imgUrl)}" alt="QR ${escapeAttr(point.label)}" />
    <div class="qr-label">${escapeHtml(point.floorLabel)}<br>${escapeHtml(point.label)}</div>
  `;
  return card;
}

function printSingleQr() {
  const selectedKey = bo.singleQrSelect.value;
  if (!selectedKey) {
    showBackofficeToast("Seleciona um ponto para imprimir.");
    return;
  }

  const point = getAllPoints().find(
    (item) => makeKey(item.floor, item.point) === selectedKey
  );

  if (!point) {
    showBackofficeToast("Ponto não encontrado.");
    return;
  }

  openPrintWindow([point], true);
}

function printAllQrs() {
  const points = getAllPoints();
  if (!points.length) {
    showBackofficeToast("Não há QR codes para imprimir.");
    return;
  }
  openPrintWindow(points, false);
}

function openPrintWindow(points, singleMode) {
  const sizeCm = Number(BO_CONFIG?.qr?.printSizeCm || 8);
  const title = singleMode ? "QR Code Individual" : "Folha A4 - QR Codes";

  const itemsHtml = points
    .map((point) => {
      const url = buildReportUrl(point);
      const qrUrl = buildQrImageUrl(url);
      return `
        <div class="item">
          <img src="${escapeAttr(qrUrl)}" alt="QR ${escapeAttr(point.label)}" />
          <div class="label">${escapeHtml(point.floorLabel)}<br>${escapeHtml(point.label)}</div>
        </div>
      `;
    })
    .join("");

  const html = `
    <!doctype html>
    <html lang="pt-PT">
    <head>
      <meta charset="utf-8" />
      <title>${title}</title>
      <style>
        @page { size: A4 portrait; margin: 10mm; }
        * { box-sizing: border-box; }
        body {
          margin: 0;
          font-family: Arial, sans-serif;
          color: #111827;
          background: #fff;
        }
        .sheet {
          display: grid;
          grid-template-columns: repeat(${singleMode ? 1 : 2}, 1fr);
          gap: 8mm;
          justify-items: center;
          align-items: start;
        }
        .item {
          width: ${sizeCm}cm;
          min-height: ${sizeCm + 1.1}cm;
          border: 1px solid #ddd;
          border-radius: 6mm;
          padding: 4mm;
          display: flex;
          flex-direction: column;
          align-items: center;
          justify-content: center;
          page-break-inside: avoid;
        }
        .item img {
          width: ${sizeCm - 1}cm;
          height: ${sizeCm - 1}cm;
          object-fit: contain;
          display: block;
        }
        .label {
          margin-top: 3mm;
          text-align: center;
          font-size: 9pt;
          line-height: 1.25;
          font-weight: bold;
        }
      </style>
    </head>
    <body>
      <div class="sheet">${itemsHtml}</div>
      <script>
        window.onload = function(){
          window.print();
        };
      </script>
    </body>
    </html>
  `;

  const printWindow = window.open("", "_blank");
  if (!printWindow) {
    showBackofficeToast("O browser bloqueou a janela de impressão.");
    return;
  }

  printWindow.document.open();
  printWindow.document.write(html);
  printWindow.document.close();
}

async function apiGet(actionName, params = {}) {
  const baseUrl = BO_CONFIG?.backendUrl;
  if (!baseUrl || baseUrl.includes("COLOCAR_AQUI")) {
    throw new Error("backendUrl não configurado no config.json");
  }

  const url = new URL(baseUrl);
  url.searchParams.set("action", getAction(actionName));

  Object.entries(params).forEach(([key, value]) => {
    if (value !== undefined && value !== null && value !== "") {
      url.searchParams.set(key, value);
    }
  });

  const response = await fetch(url.toString(), { cache: "no-store" });
  const text = await response.text();

  if (!response.ok) {
    throw new Error(`Erro HTTP ${response.status}`);
  }

  try {
    return JSON.parse(text);
  } catch {
    throw new Error("Resposta não JSON do backend.");
  }
}

async function apiPost(payload) {
  const baseUrl = BO_CONFIG?.backendUrl;
  if (!baseUrl || baseUrl.includes("COLOCAR_AQUI")) {
    throw new Error("backendUrl não configurado no config.json");
  }

  const formData = new FormData();

  Object.entries(payload).forEach(([key, value]) => {
    if (value === undefined || value === null) return;
    formData.append(key, String(value));
  });

  let response;
  try {
    response = await fetch(baseUrl, {
      method: "POST",
      body: formData
    });
  } catch (error) {
    throw new Error("Falha de ligação ao backend (POST bloqueado / CORS).");
  }

  const text = await response.text();

  try {
    return JSON.parse(text);
  } catch {
    if (!response.ok) {
      throw new Error(`Resposta inválida do backend (${response.status})`);
    }
    return { success: true, raw: text };
  }
}

function getAllPoints() {
  const floors = BO_CONFIG?.building?.floors || [];
  return floors.flatMap((floor) =>
    (floor.extinguishers || []).map((ext) => ({
      floor: floor.floor,
      floorLabel: floor.label,
      point: ext.point,
      label: ext.label || ext.point,
      shortLabel: ext.shortLabel || ext.point,
      location: ext.location || ""
    }))
  );
}

function getFloorLabel(floorNumber) {
  const floor = (BO_CONFIG?.building?.floors || []).find(
    (item) => Number(item.floor) === Number(floorNumber)
  );
  return floor?.label || String(floorNumber);
}

function buildReportUrl(point) {
  const publicBase = BO_CONFIG?.publicBaseUrl || getPublicIndexUrl();
  const url = new URL(publicBase);
  url.searchParams.set("floor", String(point.floor));
  url.searchParams.set("point", String(point.point));
  return url.toString();
}

function getPublicIndexUrl() {
  const url = new URL(window.location.href);
  url.pathname = url.pathname.replace(/backoffice\.html$/i, "index.html");
  url.search = "";
  url.hash = "";
  return url.toString();
}

function buildQrImageUrl(text) {
  const size = Number(BO_CONFIG?.qr?.sizePx || 220);
  return `https://quickchart.io/qr?size=${size}&text=${encodeURIComponent(text)}`;
}

function getAction(name) {
  return BO_CONFIG?.apiActions?.[name] || name;
}

function extractReportedItems(payload) {
  if (!payload) return [];
  if (Array.isArray(payload)) return payload;
  if (Array.isArray(payload.reported)) return payload.reported;
  if (Array.isArray(payload.items)) return payload.items;
  if (Array.isArray(payload.data)) return payload.data;
  return [];
}

function normalizeStatusKey(item) {
  if (!item) return "";

  if (typeof item === "string") return item.trim().toUpperCase();

  const floorRaw =
    item.floor ??
    item.piso ??
    item.level ??
    item.floorNumber ??
    item.idFloor;

  const pointRaw =
    item.point ??
    item.ponto ??
    item.extinguisher ??
    item.extintor ??
    item.code ??
    item.idPoint;

  if (floorRaw === undefined || floorRaw === null || !pointRaw) {
    const fallbackId =
      item.id ?? item.key ?? item.extinguisherId ?? item.extintorId ?? "";
    return String(fallbackId).trim().toUpperCase();
  }

  return makeKey(Number(floorRaw), String(pointRaw));
}

function makeKey(floor, point) {
  return `${Number(floor)}:${String(point).trim().toUpperCase()}`;
}

async function fileToPayload(file) {
  const dataUrl = await readFileAsDataURL(file);
  const base64 = dataUrl.split(",")[1] || "";
  return {
    name: file.name,
    type: file.type,
    size: file.size,
    dataUrl,
    base64
  };
}

function readFileAsDataURL(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(new Error("Não foi possível ler a fotografia."));
    reader.readAsDataURL(file);
  });
}

function formatDateTime(date) {
  return new Intl.DateTimeFormat("pt-PT", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit"
  }).format(date);
}

function formatBytes(bytes) {
  if (!bytes && bytes !== 0) return "";
  const units = ["B", "KB", "MB", "GB"];
  let value = bytes;
  let unitIndex = 0;

  while (value >= 1024 && unitIndex < units.length - 1) {
    value /= 1024;
    unitIndex += 1;
  }

  return `${value.toFixed(value >= 10 || unitIndex === 0 ? 0 : 1)} ${units[unitIndex]}`;
}

function showBackofficeToast(message) {
  bo.toast.textContent = message;
  bo.toast.classList.add("show");
  window.clearTimeout(showBackofficeToast._timer);
  showBackofficeToast._timer = window.setTimeout(() => {
    bo.toast.classList.remove("show");
  }, 3200);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function escapeAttr(value) {
  return escapeHtml(value);
}
