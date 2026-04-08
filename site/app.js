const DEFAULT_POLL_MS = 30000;
const DEFAULT_PROJECT_ROTATE_MS = 8000;
const DEFAULT_SPOOL_ROTATE_MS = 4000;
const SPOOL_PAGE_SIZE = 8;

const state = {
  projects: [],
  stats: null,
  meta: null,
  activeIndex: 0,
  activeSpoolPage: 0,
  pollTimer: null,
  rotateTimer: null,
  spoolTimer: null,
};

const bodyEl = document.getElementById("projects-body");
const tableShellEl = document.getElementById("table-shell");
const detailCardEl = document.getElementById("detail-card");
const sheetNameEl = document.getElementById("sheet-name");
const lastSyncEl = document.getElementById("last-sync");
const footerVersionEl = document.getElementById("footer-version");

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

function uiStateLabel(stateValue) {
  if (stateValue === "completed") return "Concluído";
  if (stateValue === "in_progress") return "Em andamento";
  return "Não iniciado";
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

function renderStats() {
  if (!state.stats) return;
  document.getElementById("stat-projects").textContent = formatNumber(state.stats.totalProjects);
  document.getElementById("stat-spools").textContent = formatNumber(state.stats.totalSpools);
  document.getElementById("stat-completed").textContent = formatNumber(state.stats.completed);
  document.getElementById("stat-in-progress").textContent = formatNumber(state.stats.inProgress);
  document.getElementById("stat-not-started").textContent = formatNumber(state.stats.notStarted);
  document.getElementById("stat-average").textContent = formatPercent(state.stats.averageOverallProgress);
}

function renderTable() {
  if (!state.projects.length) {
    bodyEl.innerHTML = '<tr><td colspan="14" class="loading-cell">Nenhum projeto encontrado.</td></tr>';
    return;
  }

  bodyEl.innerHTML = state.projects
    .map((project, index) => {
      const isActive = index === state.activeIndex;
      const rowClass = [
        project.uiState === "completed" ? "completed-row" : "",
        project.uiState === "in_progress" ? "in-progress-row" : "not-started-row",
        isActive ? "active-row" : "",
      ].join(" ");

      const milestoneMap = Object.fromEntries((project.milestones || []).map((item) => [item.key, item.value]));
      const completedSymbol = project.finished ? "✓" : "✕";

      return `
        <tr class="${rowClass}" data-project-index="${index}">
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
          <td class="cell-status cell-status--${project.uiState}">${project.projectStatus || uiStateLabel(project.uiState)}</td>
          <td>${milestoneMap["Fabrication Start Date"] || "—"}</td>
          <td>${milestoneMap["Boilermaker Finish Date"] || "—"}</td>
          <td>${milestoneMap["Welding Finish Date"] || "—"}</td>
          <td>${milestoneMap["Inspection Finish Date (QC)"] || "—"}</td>
          <td>${milestoneMap["TH Finish Date"] || "—"}</td>
          <td class="cell-finished cell-finished--${project.finished ? "yes" : "no"}">${completedSymbol}</td>
        </tr>
      `;
    })
    .join("");
}

function renderMilestones(project) {
  const preferredKeys = [
    "Fabrication Start Date",
    "Boilermaker Finish Date",
    "Welding Finish Date",
    "Inspection Finish Date (QC)",
    "TH Finish Date",
    "Project Finish Date",
  ];

  const milestones = preferredKeys
    .map((key) => (project.milestones || []).find((item) => item.key === key))
    .filter(Boolean);

  if (!milestones.length) {
    return '<div class="milestone-item"><span>Marcos</span><strong>Sem datas fixadas ainda</strong></div>';
  }

  return milestones
    .map(
      (item) => `
      <div class="milestone-item">
        <span>${item.label}</span>
        <strong>${item.value}</strong>
      </div>
    `
    )
    .join("");
}

function renderSpoolRows(project) {
  const spools = project.spools || [];
  if (!spools.length) {
    return {
      rows: '<tr><td colspan="6">Sem spools detalhados para este projeto.</td></tr>',
      indicator: "0 / 0",
    };
  }

  const totalPages = Math.max(1, Math.ceil(spools.length / SPOOL_PAGE_SIZE));
  if (state.activeSpoolPage >= totalPages) {
    state.activeSpoolPage = 0;
  }

  const start = state.activeSpoolPage * SPOOL_PAGE_SIZE;
  const pageRows = spools.slice(start, start + SPOOL_PAGE_SIZE);

  const rows = pageRows
    .map((spool, index) => {
      const isActive = index === 0;
      return `
        <tr class="${isActive ? "spool-row-active" : ""}">
          <td>${spool.iso || "—"}</td>
          <td>${spool.description || "—"}</td>
          <td>${formatNumber(spool.kilos, 2)}</td>
          <td>${formatNumber(spool.m2Painting, 3)}</td>
          <td>
            <span class="stage-pill">
              <span class="stage-dot stage-dot--${stageStatusClass(spool.stageStatus)}"></span>
              <span class="stage-text">${spool.stage}</span>
            </span>
          </td>
          <td>${formatPercent(spool.overallProgress)}</td>
        </tr>
      `;
    })
    .join("");

  return {
    rows,
    indicator: `${state.activeSpoolPage + 1} / ${totalPages}`,
  };
}

function renderDetail() {
  const project = state.projects[state.activeIndex];
  if (!project) {
    detailCardEl.innerHTML = '<div class="detail-placeholder">Nenhum projeto carregado.</div>';
    return;
  }

  const spoolRows = renderSpoolRows(project);

  detailCardEl.innerHTML = `
    <div class="detail-project-title">
      <div>
        <p class="detail-project-subtitle">Projeto em rotação automática</p>
        <h3>${project.projectDisplay}</h3>
      </div>
      <span class="badge badge--${project.uiState}">${project.projectStatus || uiStateLabel(project.uiState)}</span>
    </div>

    <div class="detail-grid">
      <div class="metric-chip">
        <span>Qtd. spools</span>
        <strong>${formatNumber(project.quantitySpools)}</strong>
      </div>
      <div class="metric-chip">
        <span>Peso total</span>
        <strong>${formatNumber(project.kilos, 2)}</strong>
      </div>
      <div class="metric-chip">
        <span>Painting total</span>
        <strong>${formatNumber(project.m2Painting, 3)}</strong>
      </div>
      <div class="metric-chip">
        <span>% Individual</span>
        <strong>${formatPercent(project.individualProgress)}</strong>
      </div>
      <div class="metric-chip">
        <span>% Geral</span>
        <strong>${formatPercent(project.overallProgress)}</strong>
      </div>
      <div class="metric-chip">
        <span>Spools</span>
        <strong>${project.spoolStats.completed}/${project.spoolStats.total} concluídos</strong>
      </div>
    </div>

    <div class="current-stage-box ${project.currentStageAlert ? "alert" : ""}">
      <div class="stage-pill">
        <span class="stage-dot stage-dot--${stageStatusClass(project.currentStageStatus)}"></span>
        <span class="stage-name">${project.currentStage}</span>
      </div>
      <div class="stage-meta">Etapa atual do projeto • ${formatPercent(project.currentStagePercent)}</div>
    </div>

    <div class="milestone-list">
      ${renderMilestones(project)}
    </div>

    <div class="spool-panel">
      <div class="spool-panel-head">
        <h4>Spools / ISO</h4>
        <div class="spool-page-indicator">Página ${spoolRows.indicator}</div>
      </div>

      <table class="spool-table">
        <thead>
          <tr>
            <th>ISO</th>
            <th>Descrição</th>
            <th>Peso</th>
            <th>Painting</th>
            <th>Etapa</th>
            <th>% Geral</th>
          </tr>
        </thead>
        <tbody>
          ${spoolRows.rows}
        </tbody>
      </table>
    </div>
  `;
}

function updateMeta() {
  if (!state.meta) return;
  sheetNameEl.textContent = state.meta.sheetName || "Smartsheet";
  lastSyncEl.textContent = `Última atualização: ${new Date(state.meta.lastSync).toLocaleString("pt-BR")}`;
  footerVersionEl.textContent = `Versão da sheet: ${state.meta.version}`;
}

function scrollActiveRowIntoView() {
  const row = bodyEl.querySelector(`tr[data-project-index="${state.activeIndex}"]`);
  if (!row) return;
  row.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

function restartProjectRotation() {
  window.clearInterval(state.rotateTimer);
  if (!state.projects.length) return;

  state.rotateTimer = window.setInterval(() => {
    state.activeIndex = (state.activeIndex + 1) % state.projects.length;
    state.activeSpoolPage = 0;
    renderTable();
    renderDetail();
    scrollActiveRowIntoView();
  }, DEFAULT_PROJECT_ROTATE_MS);
}

function restartSpoolRotation() {
  window.clearInterval(state.spoolTimer);
  state.spoolTimer = window.setInterval(() => {
    const project = state.projects[state.activeIndex];
    if (!project || !project.spools || project.spools.length <= SPOOL_PAGE_SIZE) return;
    const totalPages = Math.ceil(project.spools.length / SPOOL_PAGE_SIZE);
    state.activeSpoolPage = (state.activeSpoolPage + 1) % totalPages;
    renderDetail();
  }, DEFAULT_SPOOL_ROTATE_MS);
}

async function loadProjects() {
  try {
    const response = await fetch("/api/projects");
    const data = await response.json();

    if (!response.ok || !data.ok) {
      throw new Error(data.error || "Falha ao carregar projetos.");
    }

    state.projects = data.projects || [];
    state.stats = data.stats || null;
    state.meta = data.meta || null;

    if (state.activeIndex >= state.projects.length) {
      state.activeIndex = 0;
      state.activeSpoolPage = 0;
    }

    renderStats();
    renderTable();
    renderDetail();
    updateMeta();
    scrollActiveRowIntoView();
    restartProjectRotation();
    restartSpoolRotation();
  } catch (error) {
    bodyEl.innerHTML = `<tr><td colspan="14" class="loading-cell">${error.message}</td></tr>`;
    detailCardEl.innerHTML = `<div class="detail-placeholder">${error.message}</div>`;
  }
}

function startPolling() {
  window.clearInterval(state.pollTimer);
  state.pollTimer = window.setInterval(loadProjects, DEFAULT_POLL_MS);
}

function init() {
  startClocks();
  loadProjects();
  startPolling();
}

init();
