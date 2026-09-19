(() => {
  const MODULES = [
    { key: "progress", label: "Прогресс укладки слоёв", page: "road-progress.html" },
    { key: "shoulders", label: "Обочины", page: "road-shoulders.html" },
    { key: "barriers", label: "Дорожные ограждения", page: "road-barriers.html" },
    { key: "lotki", label: "Лотки", page: "lotki.html" },
  ];

  const $ = (selector, root = document) => root.querySelector(selector);

  const REPORT_CSS = `
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

  function waitForFrameLoad(frame) {
    return new Promise((resolve, reject) => {
      const timeout = window.setTimeout(() => reject(new Error("Не удалось загрузить модуль отчёта.")), 30000);
      frame.addEventListener("load", () => {
        window.clearTimeout(timeout);
        resolve();
      }, { once: true });
      frame.addEventListener("error", () => {
        window.clearTimeout(timeout);
        reject(new Error("Ошибка загрузки модуля отчёта."));
      }, { once: true });
    });
  }

  async function renderModuleReport(provider, ownerWindow, ownerDocument) {
    if (!provider || typeof provider.buildSummary !== "function") {
      throw new Error("Модуль не предоставил сводку PDF.");
    }
    await provider.ready;
    const sourceReport = provider.buildSummary();
    if (!sourceReport) throw new Error("Модуль не сформировал сводку PDF.");

    const foreignDocument = ownerDocument !== document;
    const report = foreignDocument ? document.importNode(sourceReport, true) : sourceReport;
    const importedStyles = foreignDocument
      ? [...ownerDocument.querySelectorAll("style")].map((sourceStyle) => {
          const style = document.createElement("style");
          style.textContent = sourceStyle.textContent;
          document.head.appendChild(style);
          return style;
        })
      : [];
    document.body.appendChild(report);
    try {
      await new Promise((resolve) => window.requestAnimationFrame(resolve));
      return await window.html2canvas(report, {
        scale: 2,
        backgroundColor: "#ffffff",
        useCORS: true,
        logging: false,
      });
    } finally {
      report.remove();
      importedStyles.forEach((style) => style.remove());
    }
  }

  async function loadAndRenderModule(module) {
    if (window.__moduleSummaryReport?.key === module.key) {
      return renderModuleReport(window.__moduleSummaryReport, window, document);
    }

    const frame = document.createElement("iframe");
    frame.title = `Источник сводки: ${module.label}`;
    frame.style.cssText = "position:absolute;left:-100000px;top:0;width:1200px;height:1000px;border:0;";
    document.body.appendChild(frame);
    try {
      frame.src = new URL(module.page, window.location.href).href;
      await waitForFrameLoad(frame);
      return await renderModuleReport(frame.contentWindow.__moduleSummaryReport, frame.contentWindow, frame.contentDocument);
    } finally {
      frame.remove();
    }
  }

  function addCanvasPages(pdf, canvas, isFirst) {
    const margin = 8;
    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();
    const imageWidth = pageWidth - margin * 2;
    const imageHeight = canvas.height * imageWidth / canvas.width;
    const imageData = canvas.toDataURL("image/png");

    if (!isFirst) pdf.addPage();
    let heightLeft = imageHeight;
    let position = margin;
    pdf.addImage(imageData, "PNG", margin, position, imageWidth, imageHeight);
    heightLeft -= pageHeight - margin * 2;
    while (heightLeft > 0) {
      position = heightLeft - imageHeight + margin;
      pdf.addPage();
      pdf.addImage(imageData, "PNG", margin, position, imageWidth, imageHeight);
      heightLeft -= pageHeight - margin * 2;
    }
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
      const selected = MODULES.filter((module) => selectedKeys.includes(module.key));
      const pdf = new window.jspdf.jsPDF("p", "mm", "a4");
      let hasPages = false;
      for (const module of selected) {
        const canvas = await loadAndRenderModule(module);
        addCanvasPages(pdf, canvas, !hasPages);
        hasPages = true;
      }
      pdf.save(`obshchiy-otchet-proekta-${new Date().toISOString().slice(0, 10)}.pdf`);
      if (status) status.textContent = "общий PDF сохранён";
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
