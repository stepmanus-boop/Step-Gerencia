const DEFAULT_POLL_MS = 30000;

const state = {
  projects: [],
  filteredProjects: [],
  stats: null,
  meta: null,
  searchQuery: "",
  demandFilter: "",
  selectedProjectId: null,
  pollTimer: null,
};

const bodyEl = document.getElementById("projects-body");
const detailCardEl = document.getElementById("detail-card");
const sheetNameEl = document.getElementById("sheet-name");
const lastSyncEl = document.getElementById("last-sync");
const footerVersionEl = document.getElementById("footer-version");
const searchInputEl = document.getElementById("project-search");
const clearSearchEl = document.getElementById("clear-search");
const demandFilterEl = document.getElementById("demand-filter");
const searchCountEl = document.getElementById("search-count");
const tableShellEl = document.getElementById("table-shell");
const modalEl = document.getElementById("project-modal");
const modalContentEl = document.getElementById("modal-content");
const modalTitleEl = document.getElementById("modal-title");
const modalSubtitleEl = document.getElementById("modal-subtitle");
const modalCloseEl = document.getElementById("modal-close");

function formatNumber(value, fractionDigits = 0) {
  if (value == null || Number.isNaN(value)) return "—";
  return Number(value).toLocaleString("pt-BR", {
    minimumFractionDigits: fractionDigits,
    maximumFractionDigits: fractionDigits,
  });
}

function formatPercent(value) {
  if (value == null || Number.isNaN(value)) return "—";
  return `${Number(value).toLocaleString("pt-BR", {
    minimumFractionDigits: value % 1 === 0 ? 0 : 1,
    maximumFractionDigits: 1,
  })}%`;
}

function normalizeText(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "");
}

function uiStateLabel(stateValue) {
  if (stateValue === "completed") return "Finalizado";
  if (stateValue === "awaiting_shipment") return "Aguardando envio";
  if (stateValue === "in_progress") return "Em produção";
  return "Não iniciado";
}

function translateProjectStatus(projectStatus, uiState) {
  if (uiState === "completed") return "Finalizado";
  if (uiState === "awaiting_shipment") return "Aguardando envio";

  const normalized = String(projectStatus || "").trim().toUpperCase().replace(/\s+/g, " ");
  if (["ONGOING", "ON GOING", "IN PROGRESS", "EM PRODUCAO", "EM PRODUÇÃO"].includes(normalized)) {
    return "Em produção";
  }
  if (["ON HOLD", "HOLD", "PAUSED", "EM ESPERA"].includes(normalized)) {
    return uiState === "not_started" ? "Em espera" : "Em produção";
  }
  if (["COMPLETED", "DONE", "FINISHED", "CONCLUIDO", "CONCLUÍDO", "FINALIZADO"].includes(normalized)) {
    return "Finalizado";
  }
  return projectStatus || uiStateLabel(uiState);
}

function stageStatusClass(status) {
  if (status === "completed") return "completed";
  if (status === "in_progress") return "in_progress";
  if (status === "waiting") return "waiting";
  return "ignored";
}

function setClock(targetTimeId, targetDateId, locale, timeZone) {
  const now = new Date();
  const timeText = new Intl.DateTimeFormat(locale, {
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hour12: false,
    timeZone,
  }).format(now);

  const dateText = new Intl.DateTimeFormat(locale, {
    weekday: "short",
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    timeZone,
  }).format(now);

  document.getElementById(targetTimeId).textContent = timeText;
  document.getElementById(targetDateId).textContent = dateText;
}

function startClocks() {
  const tick = () => {
    setClock("clock-br-time", "clock-br-date", "pt-BR", "America/Sao_Paulo");
    setClock("clock-pt-time", "clock-pt-date", "pt-PT", "Europe/Lisbon");
  };
  tick();
  window.setInterval(tick, 1000);
}

function enrichProjects(projects) {
  return (projects || []).map((project) => {
    const searchParts = [
      project.projectDisplay,
      project.projectNumber,
      project.projectPrefix,
      project.currentStage,
      project.projectStatus,
      ...(project.spools || []).flatMap((spool) => [spool.iso, spool.description, spool.drawing]),
    ];

    return {
      ...project,
      _searchText: normalizeText(searchParts.filter(Boolean).join(" | ")),
    };
  });
}

function buildDemandOptions() {
  if (!demandFilterEl) return;
  const selected = state.demandFilter || "";
  const options = Array.from(new Set(state.projects.map((project) => project.currentStage).filter(Boolean)))
    .sort((a, b) => a.localeCompare(b, "pt-BR"));

  demandFilterEl.innerHTML = [
    '<option value="">Todas as demandas</option>',
    ...options.map((option) => `<option value="${option}">${option}</option>`),
  ].join("");

  demandFilterEl.value = options.includes(selected) ? selected : "";
  if (!options.includes(selected)) state.demandFilter = "";
}

function applyFilter() {
  const query = normalizeText(state.searchQuery).trim();
  const demand = normalizeText(state.demandFilter).trim();

  state.filteredProjects = state.projects.filter((project) => {
    const matchesQuery = !query || project._searchText.includes(query);
    const matchesDemand = !demand || normalizeText(project.currentStage).includes(demand) || normalizeText(translateProjectStatus(project.projectStatus, project.uiState)).includes(demand);
    return matchesQuery && matchesDemand;
  });

  if (!state.filteredProjects.find((project) => project.rowId === state.selectedProjectId)) {
    state.selectedProjectId = state.filteredProjects[0]?.rowId || null;
  }
}

function getSelectedProject() {
  return state.projects.find((project) => project.rowId === state.selectedProjectId) || null;
}

function renderStats() {
  if (!state.stats) return;
  document.getElementById("stat-projects").textContent = formatNumber(state.stats.totalProjects);
  document.getElementById("stat-spools").textContent = formatNumber(state.stats.totalSpools);
  document.getElementById("stat-total-weight").textContent = `${formatNumber(state.stats.totalWeightKg, 0)} kg`;
  
  document.getElementById("stat-completed").textContent = formatNumber(state.stats.completed);
  document.getElementById("stat-in-progress").textContent = formatNumber(state.stats.inProgress);
  document.getElementById("stat-not-started").textContent = formatNumber(state.stats.notStarted);
  document.getElementById("stat-average").textContent = formatPercent(state.stats.averageOverallProgress);
}

function renderTable() {
  if (!state.filteredProjects.length) {
    bodyEl.innerHTML = '<tr><td colspan="14" class="loading-cell">Nenhum projeto encontrado para a busca informada.</td></tr>';
    searchCountEl.textContent = "0 resultado(s)";
    return;
  }

  searchCountEl.textContent = `${state.filteredProjects.length} resultado(s)`;

  bodyEl.innerHTML = state.filteredProjects
    .map((project) => {
      const isActive = project.rowId === state.selectedProjectId;
      const statusText = translateProjectStatus(project.projectStatus, project.uiState);
      const rowClass = [
        ["completed", "awaiting_shipment"].includes(project.uiState) ? "completed-row" : "",
        project.uiState === "in_progress" ? "in-progress-row" : "",
        project.uiState === "not_started" ? "not-started-row" : "",
        isActive ? "active-row" : "",
      ]
        .filter(Boolean)
        .join(" ");

      const stageMap = project.stageValues || {};
      const completedSymbol = ["completed", "awaiting_shipment"].includes(project.uiState) ? "✓" : "✕";
      const statusState = ["awaiting_shipment", "completed"].includes(project.uiState) ? "completed" : project.uiState;

      return `
        <tr class="${rowClass}" data-project-id="${project.rowId}">
          <td>${project.projectDisplay}</td>
          <td>${formatNumber(project.quantitySpools)}</td>
          <td>${formatNumber(project.kilos, 2)}</td>
          <td>${formatNumber(project.m2Painting, 3)}</td>
          <td>
            <span class="stage-pill">
              <span class="stage-dot stage-dot--${stageStatusClass(project.currentStageStatus)}"></span>
              <span class="stage-text">${project.currentStage}</span>
            </span>
          </td>
          <td>${formatPercent(project.individualProgress)}</td>
          <td>${formatPercent(project.overallProgress)}</td>
          <td><span class="cell-status cell-status--${statusState}">${statusText}</span></td>
          <td>${stageMap["Fabrication Start Date"] || "—"}</td>
          <td>${stageMap["Boilermaker Finish Date"] || "—"}</td>
          <td>${stageMap["Welding Finish Date"] || "—"}</td>
          <td>${stageMap["Inspection Finish Date (QC)"] || "—"}</td>
          <td>${stageMap["TH Finish Date"] || "—"}</td>
          <td class="cell-finished cell-finished--${project.finished ? "yes" : "no"}">${completedSymbol}</td>
        </tr>
      `;
    })
    .join("");
}

function renderSelectedProjectCard() {
  const project = getSelectedProject();
  if (!project) {
    detailCardEl.innerHTML = '<div class="detail-placeholder">Selecione um projeto na tabela ou pela busca para abrir o popup.</div>';
    return;
  }

  const statusText = translateProjectStatus(project.projectStatus, project.uiState);
  const matchedSpools = project.spools?.length || 0;

  detailCardEl.innerHTML = `
    <div class="detail-hero compact">
      <div class="detail-project-title">
        <div>
          <p class="detail-project-subtitle">Projeto selecionado</p>
          <h3>${project.projectDisplay}</h3>
        </div>
        <span class="badge badge--${["awaiting_shipment", "completed"].includes(project.uiState) ? "completed" : project.uiState}">${statusText}</span>
      </div>

      <div class="detail-grid compact-grid">
        <div class="metric-chip"><span>Qtd. itens</span><strong>${formatNumber(project.quantitySpools)}</strong></div>
        <div class="metric-chip"><span>Peso total</span><strong>${formatNumber(project.kilos, 2)}</strong></div>
        <div class="metric-chip"><span>Painting</span><strong>${formatNumber(project.m2Painting, 3)}</strong></div>
        <div class="metric-chip"><span>% Individual</span><strong>${formatPercent(project.individualProgress)}</strong></div>
        <div class="metric-chip"><span>% Geral</span><strong>${formatPercent(project.overallProgress)}</strong></div>
        <div class="metric-chip"><span>Itens internos</span><strong>${matchedSpools}</strong></div>
      </div>

      <div class="current-stage-box ${project.currentStageAlert ? "alert" : ""}">
        <div class="current-stage-head">
          <span class="current-stage-label">Etapa atual</span>
          <span class="stage-progress">${formatPercent(project.currentStagePercent)}</span>
        </div>
        <div class="stage-pill">
          <span class="stage-dot stage-dot--${stageStatusClass(project.currentStageStatus)}"></span>
          <span class="stage-name">${project.currentStage}</span>
        </div>
      </div>

      <div class="detail-actions">
        <button class="primary-button" type="button" id="open-selected-project">Abrir detalhamento completo</button>
      </div>
    </div>
  `;

  const button = document.getElementById("open-selected-project");
  if (button) {
    button.addEventListener("click", () => openProjectModal(project));
  }
}

function renderModal(project) {
  const stageOrder = state.meta?.stageOrder || [];
  const milestoneList = (project.milestones || [])
    .map((item) => `<div class="milestone-chip"><span>${item.label}</span><strong>${item.value}</strong></div>`)
    .join("");

  const spoolRows = (project.spools || [])
    .map((spool) => {
      const stageColumns = stageOrder
        .map((stage) => {
          const value = spool.stageValues?.[stage.key];
          return `<td>${value == null || value === "" ? "—" : stage.type === "percent" ? formatPercent(value) : value}</td>`;
        })
        .join("");

      return `
        <tr>
          <td>${spool.iso || "—"}</td>
          <td>${spool.description || "—"}</td>
          <td>${formatNumber(spool.kilos, 2)}</td>
          <td>${formatNumber(spool.m2Painting, 3)}</td>
          <td><span class="cell-status cell-status--${["awaiting_shipment", "completed"].includes(spool.uiState) ? "completed" : spool.uiState}">${uiStateLabel(spool.uiState)}</span></td>
          <td>${spool.stage || "—"}</td>
          <td>${formatPercent(spool.individualProgress)}</td>
          <td>${formatPercent(spool.overallProgress)}</td>
          ${stageColumns}
        </tr>
      `;
    })
    .join("");

  const stageHeaders = stageOrder.map((stage) => `<th>${stage.label}</th>`).join("");
  const statusText = translateProjectStatus(project.projectStatus, project.uiState);

  modalTitleEl.textContent = project.projectDisplay;
  modalSubtitleEl.textContent = `${statusText} • ${project.spools?.length || 0} item(ns) interno(s)`;

  modalContentEl.innerHTML = `
    <section class="modal-summary-grid">
      <article class="metric-chip"><span>Qtd. itens</span><strong>${formatNumber(project.quantitySpools)}</strong></article>
      <article class="metric-chip"><span>Peso total</span><strong>${formatNumber(project.kilos, 2)}</strong></article>
      <article class="metric-chip"><span>Painting total</span><strong>${formatNumber(project.m2Painting, 3)}</strong></article>
      <article class="metric-chip"><span>% Individual</span><strong>${formatPercent(project.individualProgress)}</strong></article>
      <article class="metric-chip"><span>% Geral</span><strong>${formatPercent(project.overallProgress)}</strong></article>
      <article class="metric-chip"><span>Etapa atual</span><strong>${project.currentStage}</strong></article>
    </section>

    <section class="modal-milestones">
      ${milestoneList || '<div class="empty-inline">Nenhum marco de data disponível.</div>'}
    </section>

    <section class="modal-table-wrap">
      <table class="modal-table">
        <thead>
          <tr>
            <th>ISO</th>
            <th>Descrição</th>
            <th>Peso</th>
            <th>Painting</th>
            <th>Status</th>
            <th>Etapa atual</th>
            <th>% Individual</th>
            <th>% Geral</th>
            ${stageHeaders}
          </tr>
        </thead>
        <tbody>
          ${spoolRows || '<tr><td colspan="999" class="loading-cell">Nenhum item interno encontrado.</td></tr>'}
        </tbody>
      </table>
    </section>
  `;
}

function openProjectModal(project) {
  state.selectedProjectId = project.rowId;
  renderTable();
  renderSelectedProjectCard();
  renderModal(project);
  modalEl.classList.remove("hidden");
  modalEl.setAttribute("aria-hidden", "false");
  document.body.classList.add("modal-open");
}

function closeProjectModal() {
  modalEl.classList.add("hidden");
  modalEl.setAttribute("aria-hidden", "true");
  document.body.classList.remove("modal-open");
}

function updateMeta() {
  if (!state.meta) return;
  sheetNameEl.textContent = state.meta.sheetName || "Smartsheet";
  lastSyncEl.textContent = `Última atualização: ${new Date(state.meta.lastSync).toLocaleString("pt-BR")}`;
  footerVersionEl.textContent = `Versão da sheet: ${state.meta.version}`;
}

async function loadProjects() {
  try {
    const response = await fetch("/api/projects");
    const data = await response.json();

    if (!response.ok || !data.ok) {
      throw new Error(data.error || "Falha ao carregar projetos.");
    }

    state.projects = enrichProjects(data.projects || []);
    state.stats = data.stats || null;
    state.meta = data.meta || null;
    buildDemandOptions();

    if (!state.selectedProjectId && state.projects.length) {
      state.selectedProjectId = state.projects[0].rowId;
    }

    applyFilter();
    renderStats();
    renderTable();
    renderSelectedProjectCard();
    updateMeta();
  } catch (error) {
    bodyEl.innerHTML = `<tr><td colspan="14" class="loading-cell">${error.message}</td></tr>`;
    detailCardEl.innerHTML = `<div class="detail-placeholder">${error.message}</div>`;
  }
}

function startPolling() {
  window.clearInterval(state.pollTimer);
  state.pollTimer = window.setInterval(loadProjects, DEFAULT_POLL_MS);
}

function bindEvents() {
  searchInputEl.addEventListener("input", (event) => {
    state.searchQuery = event.target.value;
    applyFilter();
    renderTable();
    renderSelectedProjectCard();
    tableShellEl.scrollTop = 0;
  });

  clearSearchEl.addEventListener("click", () => {
    state.searchQuery = "";
    state.demandFilter = "";
    searchInputEl.value = "";
    if (demandFilterEl) demandFilterEl.value = "";
    applyFilter();
    renderTable();
    renderSelectedProjectCard();
    tableShellEl.scrollTop = 0;
    searchInputEl.focus();
  });

  if (demandFilterEl) {
    demandFilterEl.addEventListener("change", (event) => {
      state.demandFilter = event.target.value;
      applyFilter();
      renderTable();
      renderSelectedProjectCard();
      tableShellEl.scrollTop = 0;
    });
  }

  bodyEl.addEventListener("click", (event) => {
    const row = event.target.closest("tr[data-project-id]");
    if (!row) return;
    const projectId = Number(row.dataset.projectId);
    const project = state.projects.find((item) => item.rowId === projectId);
    if (!project) return;
    state.selectedProjectId = projectId;
    renderTable();
    renderSelectedProjectCard();
    openProjectModal(project);
  });

  modalEl.addEventListener("click", (event) => {
    if (event.target.matches("[data-close-modal='true']")) {
      closeProjectModal();
    }
  });

  modalCloseEl.addEventListener("click", closeProjectModal);

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") closeProjectModal();
  });
}

function init() {
  startClocks();
  bindEvents();
  loadProjects();
  startPolling();
}

init();
