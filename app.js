let APP_CONFIG = null;
let REPORTED_SET = new Set();
let SELECTED_POINT = null;
let SELECTED_FILE = null;

const els = {};

document.addEventListener("DOMContentLoaded", initApp);

async function initApp() {
  cacheElements();
  bindEvents();

  try {
    APP_CONFIG = await loadConfig();
    fillReasonOptions();
    document.getElementById("buildingName").textContent =
      APP_CONFIG.building?.name || "Condomínio";

    await loadStatuses();
    renderBuilding();
    handleDeepLink();
  } catch (error) {
    console.error(error);
    showToast("Erro ao carregar a app. Verifica o config.json e o backend.");
  } finally {
    hideBuildingLoading();
  }
}

function cacheElements() {
  els.building = document.getElementById("building");
  els.buildingLoading = document.getElementById("buildingLoading");
  els.lastRefresh = document.getElementById("lastRefresh");

  els.overlay = document.getElementById("reportOverlay");
  els.closeModalBtn = document.getElementById("closeModalBtn");
  els.cancelBtn = document.getElementById("cancelBtn");
  els.reportForm = document.getElementById("reportForm");

  els.modalTitle = document.getElementById("modalTitle");
  els.modalSubtitle = document.getElementById("modalSubtitle");
  els.alreadyReportedBox = document.getElementById("alreadyReportedBox");

  els.hiddenFloor = document.getElementById("hiddenFloor");
  els.hiddenPoint = document.getElementById("hiddenPoint");
  els.hiddenLocation = document.getElementById("hiddenLocation");

  els.reporterName = document.getElementById("reporterName");
  els.reportReason = document.getElementById("reportReason");
  els.reportNote = document.getElementById("reportNote");
  els.submitBtn = document.getElementById("submitBtn");

  els.openCameraBtn = document.getElementById("openCameraBtn");
  els.pickFileBtn = document.getElementById("pickFileBtn");
  els.cameraInput = document.getElementById("cameraInput");
  els.fileInput = document.getElementById("fileInput");
  els.photoMeta = document.getElementById("photoMeta");

  els.toast = document.getElementById("toast");
}

function bindEvents() {
  els.closeModalBtn.addEventListener("click", closeModal);
  els.cancelBtn.addEventListener("click", closeModal);

  els.overlay.addEventListener("click", (event) => {
    if (event.target === els.overlay) closeModal();
  });

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && els.overlay.classList.contains("show")) {
      closeModal();
    }
  });

  els.openCameraBtn.addEventListener("click", () => els.cameraInput.click());
  els.pickFileBtn.addEventListener("click", () => els.fileInput.click());

  els.cameraInput.addEventListener("change", handlePhotoPicked);
  els.fileInput.addEventListener("change", handlePhotoPicked);

  els.reportForm.addEventListener("submit", submitReport);
}

async function loadConfig() {
  const res = await fetch("config.json", { cache: "no-store" });
  if (!res.ok) {
    throw new Error("Não foi possível carregar o config.json");
  }
  return res.json();
}

function fillReasonOptions() {
  const reasons = APP_CONFIG?.reportReasons || [];
  els.reportReason.innerHTML = `<option value="">Selecionar motivo</option>`;

  reasons.forEach((reason) => {
    const option = document.createElement("option");
    option.value = reason;
    option.textContent = reason;
    els.reportReason.appendChild(option);
  });
}

async function loadStatuses() {
  try {
    const response = await apiGet("status", {}, true);
    const items = extractReportedItems(response);
    REPORTED_SET = new Set(items.map(normalizeStatusKey).filter(Boolean));
    els.lastRefresh.textContent = `Atualizado às ${formatTime(new Date())}`;
  } catch (error) {
    console.error("Erro ao carregar estados:", error);
    REPORTED_SET = new Set();
    els.lastRefresh.textContent = "Sem ligação ao backend";
    showToast("Não foi possível atualizar o estado. A app continua disponível para reporte.");
  }
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

function renderBuilding() {
  const floors = APP_CONFIG?.building?.floors || [];
  els.building.innerHTML = "";

  floors.forEach((floor) => {
    const row = document.createElement("div");
    row.className = "floor-row";

    const label = document.createElement("div");
    label.className = "floor-label";
    label.textContent = floor.label;

    const body = document.createElement("div");
    body.className = "floor-body";

    const inner = document.createElement("div");
    inner.className = "floor-inner";

    const left = document.createElement("div");

    if (floor.type === "garage") {
      const garage = document.createElement("div");
      garage.className = "garage-zone";
      garage.textContent = "Garagem / espaço aberto";
      left.appendChild(garage);
    } else {
      const fractions = document.createElement("div");
      fractions.className = `fractions cols-${floor.fractions || 3}`;

      const totalFractions = floor.fractions || 3;
      for (let i = 0; i < totalFractions; i += 1) {
        const windowEl = document.createElement("div");
        windowEl.className = "window";
        fractions.appendChild(windowEl);
      }

      left.appendChild(fractions);
    }

    const exts = document.createElement("div");
    exts.className = "extinguishers";

    (floor.extinguishers || []).forEach((ext) => {
      const key = makeKey(floor.floor, ext.point);
      const isAlert = REPORTED_SET.has(key);

      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = `ext-btn ${isAlert ? "alert" : "ok"}`;
      btn.dataset.floor = String(floor.floor);
      btn.dataset.point = ext.point;
      btn.dataset.location = ext.location || "";
      btn.dataset.label = ext.label || ext.point;
      btn.dataset.title = floor.label;
      btn.setAttribute(
        "aria-label",
        `${floor.label} - ${ext.label || ext.point}${isAlert ? " - reportado" : " - sem alerta"}`
      );

      const span = document.createElement("span");
      span.textContent = ext.shortLabel || ext.label || ext.point;
      btn.appendChild(span);

      btn.addEventListener("click", () => {
        openModal({
          floor: floor.floor,
          floorLabel: floor.label,
          point: ext.point,
          label: ext.label || ext.point,
          shortLabel: ext.shortLabel || ext.point,
          location: ext.location || "",
          isAlert
        });
      });

      exts.appendChild(btn);
    });

    inner.appendChild(left);
    inner.appendChild(exts);
    body.appendChild(inner);

    row.appendChild(label);
    row.appendChild(body);

    els.building.appendChild(row);
  });

  els.building.classList.remove("hidden");
}

function openModal(ext) {
  SELECTED_POINT = ext;
  SELECTED_FILE = null;
  clearPhotoInputs();

  const targetLabel = `${ext.floorLabel} - ${ext.label}`;
  els.modalTitle.textContent = "Reportar extintor";
  els.modalSubtitle.textContent = `Ponto selecionado: ${targetLabel}`;
  els.alreadyReportedBox.classList.toggle("show", !!ext.isAlert);

  els.hiddenFloor.value = String(ext.floor);
  els.hiddenPoint.value = ext.point;
  els.hiddenLocation.value = ext.location || "";

  els.overlay.classList.add("show");
  els.overlay.setAttribute("aria-hidden", "false");
  document.body.style.overflow = "hidden";

  window.setTimeout(() => {
    els.reporterName.focus();
  }, 40);
}

function closeModal() {
  els.overlay.classList.remove("show");
  els.overlay.setAttribute("aria-hidden", "true");
  document.body.style.overflow = "";
  els.reportForm.reset();
  SELECTED_FILE = null;
  SELECTED_POINT = null;
  clearPhotoInputs();
  els.alreadyReportedBox.classList.remove("show");
}

function handlePhotoPicked(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  if (!file.type.startsWith("image/")) {
    showToast("Seleciona uma imagem válida.");
    event.target.value = "";
    return;
  }

  if (file.size > 8 * 1024 * 1024) {
    showToast("A fotografia é demasiado grande. Usa uma imagem até 8 MB.");
    event.target.value = "";
    return;
  }

  SELECTED_FILE = file;
  els.photoMeta.textContent = `${file.name} · ${formatBytes(file.size)}`;
}

function clearPhotoInputs() {
  els.cameraInput.value = "";
  els.fileInput.value = "";
  els.photoMeta.textContent = "Sem fotografia selecionada.";
}

async function submitReport(event) {
  event.preventDefault();

  if (!SELECTED_POINT) {
    showToast("Nenhum extintor selecionado.");
    return;
  }

  const name = els.reporterName.value.trim();
  const reason = els.reportReason.value;
  const notes = els.reportNote.value.trim();

  if (!name) {
    showToast("Preenche o nome.");
    els.reporterName.focus();
    return;
  }

  if (!reason) {
    showToast("Seleciona o motivo.");
    els.reportReason.focus();
    return;
  }

  const photoPayload = SELECTED_FILE ? await fileToPayload(SELECTED_FILE) : null;
  const reportPhotoRequired = !!APP_CONFIG?.features?.reportPhotoRequired;

  if (reportPhotoRequired && !photoPayload) {
    showToast("A fotografia é obrigatória neste reporte.");
    return;
  }

  els.submitBtn.disabled = true;
  els.submitBtn.textContent = "A enviar…";

  try {
    const payload = {
      action: getAction("report"),
      floor: Number(els.hiddenFloor.value),
      point: els.hiddenPoint.value,
      location: els.hiddenLocation.value,
      name,
      reason,
      notes,
      photoBase64: photoPayload?.base64 || "",
      photoDataUrl: photoPayload?.dataUrl || "",
      photoName: photoPayload?.name || "",
      photoType: photoPayload?.type || "",
      clientTs: new Date().toISOString(),
      source: "github-pages-frontoffice"
    };

    const response = await apiPost(payload);

    if (response?.success === false) {
      throw new Error(response.message || "Falha no registo do reporte.");
    }

    showToast("Reporte enviado com sucesso.");
    closeModal();
    await loadStatuses();
    renderBuilding();
  } catch (error) {
    console.error(error);
    showToast(error.message || "Não foi possível enviar o reporte.");
  } finally {
    els.submitBtn.disabled = false;
    els.submitBtn.textContent = "Enviar reporte";
  }
}

function handleDeepLink() {
  const params = new URLSearchParams(window.location.search);
  const floorParam = params.get("floor");
  const pointParam = params.get("point");

  if (!floorParam || !pointParam) return;

  const floor = Number(floorParam);
  const point = String(pointParam).trim().toUpperCase();

  const allPoints = getAllPoints();
  const match = allPoints.find(
    (item) => Number(item.floor) === floor && String(item.point).toUpperCase() === point
  );

  if (!match) {
    showToast("O ponto indicado no QR code não foi encontrado.");
    return;
  }

  window.setTimeout(() => {
    const key = makeKey(match.floor, match.point);
    openModal({
      floor: match.floor,
      floorLabel: match.floorLabel,
      point: match.point,
      label: match.label,
      shortLabel: match.shortLabel,
      location: match.location,
      isAlert: REPORTED_SET.has(key)
    });
  }, 150);
}

function getAllPoints() {
  const floors = APP_CONFIG?.building?.floors || [];
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

async function apiGet(actionName, params = {}, allowRootFallback = false) {
  const baseUrl = APP_CONFIG?.backendUrl;
  if (!baseUrl || baseUrl.includes("COLOCAR_AQUI")) {
    throw new Error("backendUrl não configurado no config.json");
  }

  const url = new URL(baseUrl);
  if (actionName) {
    url.searchParams.set("action", getAction(actionName));
  }

  Object.entries(params).forEach(([key, value]) => {
    if (value !== undefined && value !== null && value !== "") {
      url.searchParams.set(key, value);
    }
  });

  try {
    return await fetchJson(url.toString());
  } catch (error) {
    if (allowRootFallback) {
      return fetchJson(baseUrl);
    }
    throw error;
  }
}

async function apiPost(payload) {
  const baseUrl = APP_CONFIG?.backendUrl;
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

async function fetchJson(url) {
  const response = await fetch(url, { cache: "no-store" });
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

function getAction(name) {
  return APP_CONFIG?.apiActions?.[name] || name;
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

function hideBuildingLoading() {
  els.buildingLoading.classList.add("hidden");
}

function showToast(message) {
  els.toast.textContent = message;
  els.toast.classList.add("show");
  window.clearTimeout(showToast._timer);
  showToast._timer = window.setTimeout(() => {
    els.toast.classList.remove("show");
  }, 3200);
}

function formatTime(date) {
  return new Intl.DateTimeFormat("pt-PT", {
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
