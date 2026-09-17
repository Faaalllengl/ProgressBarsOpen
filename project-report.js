(() => {
  const MODULES = [
    { key: "progress", label: "Прогресс укладки слоёв" },
    { key: "shoulders", label: "Обочины" },
    { key: "barriers", label: "Дорожные ограждения" },
    { key: "lotki", label: "Лотки" },
  ];

  const $ = (selector, root = document) => root.querySelector(selector);
  const fmt = (value) => (Math.round(Number(value) * 100) / 100).toString().replace(".", ",");
  const esc = (value) => String(value ?? "").replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;",
  }[char]));

  const REPORT_CSS = `
    .project-report-canvas { width: 1100px; box-sizing: border-box; padding: 28px; background: #fff; color: #222831; font: 14px/1.45 Arial, sans-serif; }
    .project-report-canvas * { box-sizing: border-box; }
    .project-report-cover { border-bottom: 2px solid #b3131b; padding-bottom: 18px; margin-bottom: 20px; }
    .project-report-cover h2 { font-size: 24px; margin: 0 0 4px; }
    .project-report-cover p { margin: 3px 0; color: #68717c; }
    .project-report-summary { display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; margin: 16px 0 24px; }
    .project-report-metric { border: 1px solid #d8dde3; border-radius: 7px; padding: 13px; background: #fafbfc; }
    .project-report-metric small { display: block; color: #68717c; font-size: 11px; text-transform: uppercase; }
    .project-report-metric strong { display: block; font-size: 23px; margin-top: 4px; }
    .project-report-module { margin-top: 28px; }
    .project-report-module h2 { border-bottom: 1px solid #d8dde3; padding-bottom: 8px; margin-bottom: 12px; }
    .project-report-module-meta { color: #68717c; margin-bottom: 12px; }
    .project-report-layer { border: 1px solid #d8dde3; border-radius: 7px; padding: 10px 12px; margin: 9px 0; }
    .project-report-layer-head { display: flex; justify-content: space-between; gap: 12px; margin-bottom: 7px; }
    .project-report-layer-name { font-weight: 700; }
    .project-report-layer-percent { color: #68717c; white-space: nowrap; }
    .project-report-track { position: relative; height: 24px; background: #e5e8ec; border-radius: 4px; overflow: hidden; }
    .project-report-segment { position: absolute; top: 0; height: 100%; background: #b3131b; opacity: .92; }
    .project-report-shoulder-segment { display: flex; background: transparent; }
    .project-report-shoulder-half { display: block; flex: 0 0 50%; width: 50%; background: repeating-linear-gradient(135deg, #9aa2aa 0 4px, #dfe2e5 4px 8px); }
    .project-report-shoulder-half.done { background: #b3131b; }
    .project-report-lotki-track { display: flex; }
    .project-report-lotki-segment { position: static; flex: 0 0 auto; background: #aeb5bd; }
    .project-report-lotki-segment.done { background: #4b9d69; }
    .project-report-legend { font-size: 11px; color: #68717c; margin-top: 5px; }
    .project-report-table { width: 100%; border-collapse: collapse; margin-top: 12px; font-size: 12px; }
    .project-report-table th, .project-report-table td { border: 1px solid #d8dde3; padding: 6px 7px; text-align: left; }
    .project-report-table th { background: #f3f5f7; }
    .project-report-empty { color: #68717c; padding: 12px 0; }
    .project-report-foot { border-top: 1px solid #d8dde3; margin-top: 28px; padding-top: 10px; color: #68717c; font-size: 11px; }
    .project-report-picker { border: 0; border-radius: 10px; padding: 0; max-width: 420px; width: calc(100% - 32px); box-shadow: 0 12px 40px #1f29374d; }
    .project-report-picker::backdrop { background: #1f293780; }
    .project-report-picker-form { padding: 22px; }
    .project-report-picker h2 { font-size: 20px; margin: 0 0 6px; }
    .project-report-picker p { color: #68717c; margin: 0 0 16px; }
    .project-report-picker-options { display: grid; gap: 10px; }
    .project-report-picker-option { display: flex; align-items: center; gap: 10px; border: 1px solid #d8dde3; border-radius: 6px; padding: 10px; cursor: pointer; }
    .project-report-picker-option input { width: 17px; height: 17px; }
    .project-report-picker-actions { display: flex; justify-content: flex-end; gap: 8px; margin-top: 20px; }
    .project-report-picker-error { color: #b3131b; font-size: 12px; min-height: 18px; margin-top: 8px; }
  `;

  function ensureStyle() {
    if ($("#projectReportStyles")) return;
    const style = document.createElement("style");
    style.id = "projectReportStyles";
    style.textContent = REPORT_CSS;
    document.head.appendChild(style);
  }

  function activeState(data) {
    return data?.projects?.find((project) => project.id === data.activeProjectId)?.state
      || data?.projects?.[0]?.state
      || data
      || { start: 0, total: 0, layers: [] };
  }

  function segmentsLength(segments, shoulder = false) {
    return (segments || []).reduce((sum, segment) => {
      const length = Math.max(0, Number(segment.e ?? segment.end) - Number(segment.s ?? segment.start));
      const factor = shoulder ? (Array.isArray(segment.sides) ? segment.sides.length / 2 : 0) : 1;
      return sum + length * factor;
    }, 0);
  }

  function stateStats(state, shoulder = false) {
    const route = Math.max(0, Number(state.total) - Number(state.start || 0));
    const layers = state.layers || [];
    const covered = layers.reduce((sum, layer) => sum + segmentsLength(layer.segments, shoulder), 0);
    const denominator = route * layers.length;
    return { route, covered, percent: denominator ? Math.min(100, covered / denominator * 100) : 0 };
  }

  function renderBar(layer, state, shoulder = false, unit = "km") {
    const start = Number(state.start || 0);
    const route = Math.max(0.001, Number(state.total) - start);
    return (layer.segments || []).map((segment) => {
      const from = Number(segment.s ?? segment.start);
      const to = Number(segment.e ?? segment.end);
      const left = Math.max(0, (from - start) / route * 100);
      const width = Math.max(.2, (to - from) / route * 100);
      if (!shoulder) {
        return `<span class="project-report-segment" style="left:${left}%;width:${width}%" title="${unit === "pk" ? "ПК" : "км"} ${fmt(from)}–${fmt(to)}"></span>`;
      }
      const sides = Array.isArray(segment.sides) ? segment.sides : [];
      return `<span class="project-report-segment project-report-shoulder-segment" style="left:${left}%;width:${width}%"><i class="project-report-shoulder-half ${sides.includes("left") ? "done" : ""}"></i><i class="project-report-shoulder-half ${sides.includes("right") ? "done" : ""}"></i></span>`;
    }).join("");
  }

  function sideLabel(segment) {
    const sides = Array.isArray(segment?.sides) ? [...new Set(segment.sides)] : [];
    if (sides.includes("left") && sides.includes("right")) return "слева и справа";
    if (sides.includes("left")) return "слева";
    if (sides.includes("right")) return "справа";
    return "—";
  }

  function renderLayers(state, shoulder = false, unit = "km") {
    return (state.layers || []).map((layer) => {
      const covered = segmentsLength(layer.segments, shoulder);
      const route = Math.max(0, Number(state.total) - Number(state.start || 0));
      const percent = route ? covered / route * 100 : 0;
      const unitLabel = unit === "pk" ? "ПК" : "км";
      const lengthLabel = unit === "pk" ? `${fmt(covered)} ПК (${fmt(covered * 0.1)} км)` : `${fmt(covered)} км`;
      const rows = (layer.segments || []).map((segment, index) => `<tr><td>${index + 1}</td><td>${fmt(segment.s ?? segment.start)}</td><td>${fmt(segment.e ?? segment.end)}</td>${shoulder ? `<td>${sideLabel(segment)}</td>` : ""}<td>${fmt(Math.max(0, Number(segment.e ?? segment.end) - Number(segment.s ?? segment.start)))} ${unitLabel}</td><td>${esc(segment.status || (segment.completed ? "Выполнено" : "Не выполнено"))}</td></tr>`).join("");
      const layerSides = shoulder ? [...new Set((layer.segments || []).flatMap((segment) => Array.isArray(segment.sides) ? segment.sides : []))] : [];
      const layerSideLabel = shoulder
        ? (layerSides.includes("left") && layerSides.includes("right") ? "слева и справа" : layerSides.includes("left") ? "слева" : layerSides.includes("right") ? "справа" : "")
        : "";
      return `<section class="project-report-layer"><div class="project-report-layer-head"><span class="project-report-layer-name">${esc(layer.name)}</span><span class="project-report-layer-percent">${percent.toFixed(1)}% · ${lengthLabel}${layerSideLabel ? ` · ${layerSideLabel}` : ""}</span></div><div class="project-report-track">${renderBar(layer, state, shoulder, unit)}</div><div class="project-report-legend">План: ${Number(layer.planPercent || 0).toFixed(1)}% · Выполнено: ${lengthLabel}${layerSideLabel ? ` · ${layerSideLabel}` : ""}</div>${rows ? `<table class="project-report-table"><thead><tr><th>№</th><th>От, ${unitLabel}</th><th>До, ${unitLabel}</th>${shoulder ? "<th>Сторона</th>" : ""}<th>Длина</th><th>Статус</th></tr></thead><tbody>${rows}</tbody></table>` : `<div class="project-report-empty">Участки не добавлены.</div>`}</section>`;
    }).join("");
  }

  function moduleSection(title, key, state, shoulder = false, unit = "km") {
    const stats = stateStats(state, shoulder);
    const routeLabel = unit === "pk" ? `${fmt(stats.route)} ПК (${fmt(stats.route * 0.1)} км)` : `${fmt(stats.route)} км`;
    return `<section class="project-report-module" data-report-module="${key}"><h2>${title}</h2><div class="project-report-module-meta">Протяжённость: ${routeLabel} · Выполнение модуля: <b>${stats.percent.toFixed(1)}%</b></div>${renderLayers(state, shoulder, unit) || `<div class="project-report-empty">Данные отсутствуют.</div>`}</section>`;
  }

  function lotkiSection(data) {
    const items = data?.segments || [];
    const lengthMeters = (item) => Math.max(0, item.end - item.start) * 100 * Math.max(1, (item.sides || []).length);
    const total = items.reduce((sum, item) => sum + lengthMeters(item), 0);
    const done = items.filter((item) => item.completed).reduce((sum, item) => sum + lengthMeters(item), 0);
    const rows = items.map((item, index) => `<tr><td>${index + 1}</td><td>${esc(item.startText || fmt(item.start))}</td><td>${esc(item.endText || fmt(item.end))}</td><td>${(item.sides || []).map((side) => side === "left" ? "слева" : "справа").join(", ")}</td><td>${Math.round(lengthMeters(item)).toLocaleString("ru-RU")} м</td><td>${item.completed ? "Выполнен" : "Не выполнен"}</td></tr>`).join("");
    const track = items.map((item) => `<span class="project-report-segment project-report-lotki-segment ${item.completed ? "done" : ""}" style="width:${total ? lengthMeters(item) / total * 100 : 0}%" title="${esc(item.startText)}–${esc(item.endText)}"></span>`).join("");
    return `<section class="project-report-module" data-report-module="lotki"><h2>Лотки</h2><div class="project-report-module-meta">Отрезков: ${items.length} · Выполнено: ${total ? (done / total * 100).toFixed(1) : "0.0"}% · Выполнено: ${Math.round(done).toLocaleString("ru-RU")} м из ${Math.round(total).toLocaleString("ru-RU")} м</div><div class="project-report-track project-report-lotki-track">${track || `<span class="project-report-empty">Данные отсутствуют.</span>`}</div><div class="project-report-legend"><span style="color:#4b9d69">Выполнено</span> · серый — не выполнено</div>${rows ? `<table class="project-report-table"><thead><tr><th>№</th><th>От ПК</th><th>До ПК</th><th>Сторона</th><th>Длина, м</th><th>Статус</th></tr></thead><tbody>${rows}</tbody></table>` : `<div class="project-report-empty">Данные отсутствуют.</div>`}</section>`;
  }

  function buildReport(data, selectedKeys) {
    const selected = new Set(selectedKeys);
    const main = activeState(data.main);
    const barriers = activeState(data.barriers);
    const shoulders = activeState(data.shoulders);
    const lotki = data.lotki;
    const stats = {
      progress: stateStats(main),
      shoulders: stateStats(shoulders, true),
      barriers: stateStats(barriers),
    };
    const lotkiItems = lotki?.segments || [];
    const lotkiLength = (item) => Math.max(0, item.end - item.start) * 100 * Math.max(1, (item.sides || []).length);
    const lotkiTotal = lotkiItems.reduce((sum, item) => sum + lotkiLength(item), 0);
    const lotkiDone = lotkiItems.filter((item) => item.completed).reduce((sum, item) => sum + lotkiLength(item), 0);
    const metrics = [
      ["progress", "Прогресс укладки", stats.progress.percent],
      ["shoulders", "Обочины", stats.shoulders.percent],
      ["barriers", "Ограждения", stats.barriers.percent],
      ["lotki", "Лотки", lotkiTotal ? lotkiDone / lotkiTotal * 100 : 0],
    ].filter(([key]) => selected.has(key));
    const sections = [
      selected.has("progress") ? moduleSection("Прогресс укладки слоёв", "progress", main, false, "pk") : "",
      selected.has("shoulders") ? moduleSection("Обочины", "shoulders", shoulders, true, "pk") : "",
      selected.has("barriers") ? moduleSection("Дорожные ограждения", "barriers", barriers, false, "km") : "",
      selected.has("lotki") ? lotkiSection(lotki) : "",
    ].join("");
    const report = document.createElement("div");
    report.className = "project-report-canvas";
    report.innerHTML = `<div class="project-report-cover"><h2>Общий отчёт проекта</h2><p>Дата формирования: ${new Date().toLocaleString("ru-RU")}</p><p>Отчёт объединяет выбранные модули проекта. Каждый модуль показан отдельным разделом.</p></div><div class="project-report-summary" style="grid-template-columns:repeat(${Math.max(1, Math.min(4, metrics.length))},1fr)">${metrics.map(([, label, percent]) => `<div class="project-report-metric"><small>${label}</small><strong>${percent.toFixed(1)}%</strong></div>`).join("")}</div>${sections}<div class="project-report-foot">Отчёт сформирован автоматически. Пикетаж и километры показаны в исходных единицах соответствующих модулей.</div>`;
    return report;
  }

  async function getJson(path) {
    const response = await fetch(path, { cache: "no-store" });
    if (!response.ok) throw new Error(`Не удалось загрузить ${path}`);
    return response.json();
  }

  async function exportReport(selectedKeys, sourceButton) {
    if (!window.jspdf || !window.html2canvas) {
      alert("Не удалось загрузить библиотеки PDF. Проверьте интернет и обновите страницу.");
      return;
    }
    const status = $("#statusMsg") || $("#status");
    if (sourceButton) sourceButton.dataset.exporting = "true";
    if (status) status.textContent = "формирование общего PDF…";
    try {
      const [main, barriers, shoulders, lotki] = await Promise.all([
        getJson("/api/data"),
        getJson("/api/barriers-data"),
        getJson("/api/shoulders-data"),
        getJson("/api/lotki-data"),
      ]);
      const report = buildReport({ main, barriers, shoulders, lotki }, selectedKeys);
      document.body.appendChild(report);
      try {
        const { jsPDF } = window.jspdf;
        const pdf = new jsPDF("p", "mm", "a4");
        const margin = 10;
        const pageWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();
        const imageWidth = pageWidth - margin * 2;
        const contentHeight = pageHeight - margin * 2;
        const reportWidth = Math.ceil(report.getBoundingClientRect().width);
        const modules = [...report.querySelectorAll(":scope > .project-report-module")];
        const intro = [...report.querySelectorAll(":scope > .project-report-cover, :scope > .project-report-summary")];
        const footer = report.querySelector(":scope > .project-report-foot");
        const blocks = [
          { nodes: intro, isModule: false },
          ...modules.map((module) => ({ nodes: [module], isModule: true })),
          ...(footer ? [{ nodes: [footer], isModule: false }] : []),
        ];
        let currentY = margin;
        let pageHasContent = false;

        const renderBlock = async (nodes) => {
          const frame = document.createElement("div");
          frame.className = "project-report-canvas";
          frame.style.cssText = `position:absolute;left:-100000px;top:0;width:${reportWidth}px;background:#fff;box-shadow:none;`;
          nodes.forEach((node) => frame.appendChild(node.cloneNode(true)));
          document.body.appendChild(frame);
          try {
            return await html2canvas(frame, { scale: 2, backgroundColor: "#fff", useCORS: true, logging: false });
          } finally {
            frame.remove();
          }
        };

        const addBlock = async ({ nodes, isModule }) => {
          const canvas = await renderBlock(nodes);
          const blockHeight = canvas.height * imageWidth / canvas.width;
          if (blockHeight <= contentHeight) {
            if (isModule && pageHasContent && currentY + blockHeight > pageHeight - margin) {
              pdf.addPage();
              currentY = margin;
              pageHasContent = false;
            } else if (!isModule && pageHasContent && currentY + blockHeight > pageHeight - margin) {
              pdf.addPage();
              currentY = margin;
              pageHasContent = false;
            }
            pdf.addImage(canvas, "PNG", margin, currentY, imageWidth, blockHeight);
            currentY += blockHeight;
            pageHasContent = true;
            return;
          }

          if (pageHasContent) {
            pdf.addPage();
            currentY = margin;
            pageHasContent = false;
          }
          const sliceHeight = Math.max(1, Math.floor(canvas.height * contentHeight / blockHeight));
          let offset = 0;
          let lastVisibleHeight = 0;
          while (offset < canvas.height) {
            const visibleHeight = Math.min(sliceHeight, canvas.height - offset);
            pdf.addImage(canvas, "PNG", margin, margin - offset * imageWidth / canvas.width, imageWidth, blockHeight);
            offset += visibleHeight;
            lastVisibleHeight = visibleHeight;
            if (offset < canvas.height) pdf.addPage();
          }
          currentY = margin + lastVisibleHeight * imageWidth / canvas.width;
          pageHasContent = true;
        };

        for (const block of blocks) await addBlock(block);
        pdf.save(`obshchiy-otchet-proekta-${new Date().toISOString().slice(0, 10)}.pdf`);
        if (status) status.textContent = "общий PDF сохранён";
      } finally {
        report.remove();
      }
    } catch (error) {
      console.error(error);
      if (status) status.textContent = "ошибка общего PDF";
      alert("Не удалось сформировать общий отчёт PDF.");
    } finally {
      if (sourceButton) delete sourceButton.dataset.exporting;
    }
  }

  function createPicker() {
    const dialog = document.createElement("dialog");
    dialog.className = "project-report-picker";
    dialog.innerHTML = `<form class="project-report-picker-form"><h2>Модули для общего отчёта</h2><p>Выберите модули, которые нужно включить в PDF.</p><div class="project-report-picker-options">${MODULES.map(({ key, label }) => `<label class="project-report-picker-option"><input type="checkbox" name="module" value="${key}" checked>${label}</label>`).join("")}</div><div class="project-report-picker-error" aria-live="polite"></div><div class="project-report-picker-actions"><button type="button" data-picker-cancel>Отмена</button><button type="submit">Сформировать PDF</button></div></form>`;
    document.body.appendChild(dialog);
    const form = $("form", dialog);
    const error = $(".project-report-picker-error", dialog);
    const cancel = $("[data-picker-cancel]", dialog);
    const show = (sourceButton) => {
      error.textContent = "";
      form.querySelectorAll('input[name="module"]').forEach((input) => { input.checked = true; });
      dialog._sourceButton = sourceButton;
      dialog.showModal();
    };
    cancel.addEventListener("click", () => dialog.close());
    form.addEventListener("submit", (event) => {
      event.preventDefault();
      const selected = [...form.querySelectorAll('input[name="module"]:checked')].map((input) => input.value);
      if (!selected.length) {
        error.textContent = "Выберите хотя бы один модуль.";
        return;
      }
      dialog.close();
      exportReport(selected, dialog._sourceButton);
    });
    return show;
  }

  function install() {
    const links = [...document.querySelectorAll("a.report-btn")];
    if (!links.length) return;
    ensureStyle();
    const showPicker = createPicker();
    links.forEach((link) => link.addEventListener("click", (event) => {
      event.preventDefault();
      showPicker(link);
    }));
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", install, { once: true });
  } else {
    install();
  }
})();
