// Ajustes de fachada inspirados no edifício real.
// Este ficheiro substitui apenas o render do prédio dos extintores.
(function () {
  function facadeFloorLabel(floorValue, labelValue) {
    const floorNumber = Number(floorValue);
    const label = String(labelValue ?? floorValue ?? "").trim();
    if (Number.isFinite(floorNumber) && floorNumber > 0) return `${floorNumber}º`;
    return label || String(floorValue ?? "");
  }

  function facadeKind(floor) {
    const floorNumber = Number(floor.floor);
    if (floor.type === "garage" || floorNumber < 0) return "garage";
    if (floor.type === "service" || floorNumber === 0) return "service";
    return "residential";
  }

  function createStairTower() {
    const stair = document.createElement("div");
    stair.className = "stair-tower";
    for (let i = 0; i < 3; i += 1) {
      const pane = document.createElement("span");
      stair.appendChild(pane);
    }
    return stair;
  }

  function createResidentialWindows(floor) {
    const floorNumber = Number(floor.floor);
    const windows = document.createElement("div");

    // Pisos 2 a 9: 3 janelas iguais = 3 frações.
    // Piso 10 mantém a configuração do ficheiro, porque o topo do prédio é diferente.
    const totalWindows = floorNumber >= 2 && floorNumber <= 9
      ? 3
      : Math.max(2, Math.min(Number(floor.fractions || 3), 4));

    windows.className = `facade-windows cols-${totalWindows}`;

    for (let i = 0; i < totalWindows; i += 1) {
      const windowEl = document.createElement("div");
      const shouldBeWide = !(floorNumber >= 2 && floorNumber <= 9) && i === 0 && totalWindows > 2;
      windowEl.className = `window facade-window ${shouldBeWide ? "wide" : ""}`;
      windows.appendChild(windowEl);
    }

    return windows;
  }

  function createStorefronts() {
    const stores = document.createElement("div");
    stores.className = "storefront-zone";
    for (let i = 0; i < 3; i += 1) {
      const store = document.createElement("span");
      store.className = "storefront-window";
      stores.appendChild(store);
    }
    return stores;
  }

  function createGarageZone() {
    const garage = document.createElement("div");
    garage.className = "garage-zone facade-garage";
    garage.textContent = "Garagem / espaço aberto";
    return garage;
  }

  window.renderBuilding = function renderBuilding() {
    const floors = window.APP_CONFIG?.building?.floors || APP_CONFIG?.building?.floors || [];
    els.building.innerHTML = "";

    const facade = document.createElement("div");
    facade.className = "facade-shell";

    const roof = document.createElement("div");
    roof.className = "facade-roof";
    roof.innerHTML = `<span></span><span></span><span></span>`;
    facade.appendChild(roof);

    floors.forEach((floor) => {
      const kind = facadeKind(floor);
      const displayLabel = facadeFloorLabel(floor.floor, floor.label);

      const row = document.createElement("div");
      row.className = `facade-row ${kind}`;

      const label = document.createElement("div");
      label.className = "floor-label";
      label.textContent = displayLabel;

      const body = document.createElement("div");
      body.className = "floor-body";

      const inner = document.createElement("div");
      inner.className = `facade-inner ${kind === "residential" ? "has-stair" : "no-stair"}`;

      const facadeContent = document.createElement("div");
      facadeContent.className = `facade-content ${kind}`;

      if (kind === "residential") {
        inner.appendChild(createStairTower());
        facadeContent.appendChild(createResidentialWindows(floor));
      } else if (kind === "service") {
        facadeContent.appendChild(createStorefronts());
      } else {
        facadeContent.appendChild(createGarageZone());
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
        btn.dataset.title = displayLabel;
        btn.setAttribute(
          "aria-label",
          `${displayLabel} - ${ext.label || ext.point}${isAlert ? " - reportado" : " - sem alerta"}`
        );

        const span = document.createElement("span");
        span.textContent = ext.shortLabel || ext.label || ext.point;
        btn.appendChild(span);

        btn.addEventListener("click", () => {
          openModal({
            floor: floor.floor,
            floorLabel: displayLabel,
            point: ext.point,
            label: ext.label || ext.point,
            shortLabel: ext.shortLabel || ext.point,
            location: ext.location || "",
            isAlert
          });
        });

        exts.appendChild(btn);
      });

      inner.appendChild(facadeContent);
      inner.appendChild(exts);
      body.appendChild(inner);
      row.appendChild(label);
      row.appendChild(body);
      facade.appendChild(row);
    });

    els.building.appendChild(facade);
    els.building.classList.remove("hidden");
  };
})();
