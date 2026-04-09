const DEFAULT_POLL_MS = 30000;

const state = {
  projects: [],
  filteredProjects: [],
  stats: null,
  meta: null,
  searchQuery: "",
  demandFilter: "",
  weekFilter: "",
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
const weekFilterEl = document.getElementById("week-filter");
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


function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function getWeekAnchor(year) {
  const jan1 = new Date(Date.UTC(year, 0, 1));
  const anchor = new Date(jan1);
  anchor.setUTCDate(jan1.getUTCDate() - jan1.getUTCDay());
  return anchor;
}

function getCurrentBrazilYear() {
  return getCurrentBrazilDate().getUTCFullYear();
}

function formatProductionWeekLabel(weekNumber, weekYear) {
  return weekYear < getCurrentBrazilYear()
    ? `Semana ${weekNumber} - ${weekYear}`
    : `Semana ${weekNumber}`;
}

function getProductionWeekLabelFromDate(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
  let weekYear = date.getUTCFullYear();
  const nextAnchor = getWeekAnchor(weekYear + 1);
  if (date >= nextAnchor) {
    weekYear += 1;
  } else {
    const currentAnchor = getWeekAnchor(weekYear);
    if (date < currentAnchor) weekYear -= 1;
  }

  const anchor = getWeekAnchor(weekYear);
  const diffDays = Math.floor((date - anchor) / 86400000);
  const weekNumber = Math.floor(diffDays / 7) + 1;
  return formatProductionWeekLabel(weekNumber, weekYear);
}

function getCurrentBrazilDate() {
  const formatter = new Intl.DateTimeFormat("en-CA", {
    timeZone: "America/Sao_Paulo",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  });
  const parts = formatter.formatToParts(new Date());
  const year = Number(parts.find((item) => item.type === "year")?.value);
  const month = Number(parts.find((item) => item.type === "month")?.value);
  const day = Number(parts.find((item) => item.type === "day")?.value);
  return new Date(Date.UTC(year, month - 1, day));
}

function getCurrentProductionWeekLabel() {
  return getProductionWeekLabelFromDate(getCurrentBrazilDate());
}

function parseDateString(value) {
  const raw = String(value || "").trim();
  if (!raw) return null;
  let match = raw.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (match) {
    const day = Number(match[1]);
    const month = Number(match[2]) - 1;
    const year = Number(match[3]);
    return new Date(Date.UTC(year, month, day));
  }
  match = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) {
    const year = Number(match[1]);
    const month = Number(match[2]) - 1;
    const day = Number(match[3]);
    return new Date(Date.UTC(year, month, day));
  }
  const date = new Date(raw);
  if (!Number.isNaN(date.getTime())) {
    return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
  }
  return null;
}

function getMilestoneValue(entity, milestoneKey) {
  return (entity?.milestones || []).find((item) => item.key === milestoneKey)?.value || "";
}

function inferSpoolWeldingInfo(spool) {
  const milestoneDate = parseDateString(getMilestoneValue(spool, "Welding Finish Date"));
  const finished = spool.finished || spool.uiState === "completed" || spool.uiState === "awaiting_shipment";
  const weldedWeightKg = spool.weldedWeightKg ?? (milestoneDate || finished ? (spool.kilos || 0) : 0);
  const weldingWeek = spool.weldingWeek || (milestoneDate ? getProductionWeekLabelFromDate(milestoneDate) : "");
  return {
    ...spool,
    weldedWeightKg,
    weldingWeek,
  };
}

function buildDerivedStats(projects, upstreamStats) {
  const stats = {
    totalProjects: upstreamStats?.totalProjects ?? projects.length,
    totalSpools: upstreamStats?.totalSpools ?? 0,
    totalWeightKg: upstreamStats?.totalWeightKg ?? 0,
    totalWeldedWeightKg: upstreamStats?.totalWeldedWeightKg ?? 0,
    totalPaintingM2: upstreamStats?.totalPaintingM2 ?? 0,
    completed: upstreamStats?.completed ?? 0,
    inProgress: upstreamStats?.inProgress ?? 0,
    notStarted: upstreamStats?.notStarted ?? 0,
    averageOverallProgress: upstreamStats?.averageOverallProgress ?? 0,
  };

  let progressAccumulator = 0;
  let sawTotalWeight = Number.isFinite(upstreamStats?.totalWeightKg);
  let sawWeldedWeight = Number.isFinite(upstreamStats?.totalWeldedWeightKg);
  let sawPainting = Number.isFinite(upstreamStats?.totalPaintingM2);
  let sawTotalSpools = Number.isFinite(upstreamStats?.totalSpools);

  if (!(Number.isFinite(upstreamStats?.completed) && Number.isFinite(upstreamStats?.inProgress) && Number.isFinite(upstreamStats?.notStarted))) {
    stats.completed = 0;
    stats.inProgress = 0;
    stats.notStarted = 0;
  }

  if (!Number.isFinite(upstreamStats?.averageOverallProgress)) {
    stats.averageOverallProgress = 0;
  }

  for (const project of projects) {
    if (!sawTotalSpools) stats.totalSpools += project.quantitySpools || 0;
    if (!sawTotalWeight) stats.totalWeightKg += project.kilos || 0;
    if (!sawWeldedWeight) stats.totalWeldedWeightKg += project.weldedWeightKg || 0;
    if (!sawPainting) stats.totalPaintingM2 += project.m2Painting || 0;

    if (!(Number.isFinite(upstreamStats?.completed) && Number.isFinite(upstreamStats?.inProgress) && Number.isFinite(upstreamStats?.notStarted))) {
      if (["completed", "awaiting_shipment"].includes(project.uiState)) stats.completed += 1;
      else if (project.uiState === "in_progress") stats.inProgress += 1;
      else stats.notStarted += 1;
    }

    if (!Number.isFinite(upstreamStats?.averageOverallProgress)) {
      progressAccumulator += project.overallProgress || 0;
    }
  }

  if (!Number.isFinite(upstreamStats?.averageOverallProgress)) {
    stats.averageOverallProgress = projects.length ? progressAccumulator / projects.length : 0;
  }

  return stats;
}

function parseWeekLabel(label) {
  const text = String(label || "").trim();
  const weekMatch = text.match(/Semana\s+(\d+)/i);
  const yearMatch = text.match(/-\s*(\d{4})$/);
  return {
    week: weekMatch ? Number(weekMatch[1]) : Number.MAX_SAFE_INTEGER,
    year: yearMatch ? Number(yearMatch[1]) : getCurrentBrazilYear(),
  };
}

function getWeekNumber(label) {
  return parseWeekLabel(label).week;
}

function compareWeekLabels(a, b) {
  const left = parseWeekLabel(a);
  const right = parseWeekLabel(b);
  if (left.year !== right.year) return left.year - right.year;
  if (left.week !== right.week) return left.week - right.week;
  return String(a || "").localeCompare(String(b || ""), "pt-BR");
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
  if (uiState === "not_started") return "Não iniciado";

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


function percentStateClass(value) {
  if (value == null || Number.isNaN(value)) return "";
  if (Number(value) >= 100) return "value-complete";
  if (Number(value) > 0) return "value-progress";
  return "";
}

function tableCellClass(value, type = "percent") {
  if (type !== "percent") return "";
  return percentStateClass(value);
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
    const spools = (project.spools || []).map(inferSpoolWeldingInfo);
    const derivedWeldedWeight = spools.reduce((total, spool) => total + (spool.weldedWeightKg || 0), 0);
    const derivedWeekDates = spools
      .map((spool) => ({ label: spool.weldingWeek, date: parseDateString(getMilestoneValue(spool, "Welding Finish Date")) }))
      .filter((item) => item.label && item.date);
    derivedWeekDates.sort((a, b) => b.date - a.date);

    const projectWeldingDate = parseDateString(getMilestoneValue(project, "Welding Finish Date"));
    const searchParts = [
      project.projectDisplay,
      project.projectNumber,
      project.projectPrefix,
      project.currentStage,
      project.projectStatus,
      ...(spools || []).flatMap((spool) => [spool.iso, spool.description, spool.drawing]),
    ];

    return {
      ...project,
      spools,
      weldedWeightKg: project.weldedWeightKg ?? derivedWeldedWeight,
      weldingWeek: project.weldingWeek || (projectWeldingDate ? getProductionWeekLabelFromDate(projectWeldingDate) : derivedWeekDates[0]?.label || ""),
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

function buildWeekOptions() {
  if (!weekFilterEl) return;
  const selected = state.weekFilter || "";
  const currentWeek = getCurrentProductionWeekLabel();
  const weekLabels = Array.from(
    new Set([
      currentWeek,
      ...state.projects.map((project) => project.weldingWeek).filter(Boolean),
    ])
  ).sort(compareWeekLabels);

  const options = ['<option value="">Todas as semanas</option>'];
  for (const label of weekLabels) {
    options.push(`<option value="${label}">${label}</option>`);
  }

  weekFilterEl.innerHTML = options.join("");
  weekFilterEl.value = weekLabels.includes(selected) ? selected : "";
  if (!weekLabels.includes(selected)) state.weekFilter = "";
}

function getActiveWeekLabel() {
  return state.weekFilter || getCurrentProductionWeekLabel();
}

function getWeldedWeightForWeek(weekLabel) {
  if (!weekLabel) return 0;
  return state.projects.reduce((total, project) => {
    const spoolWeight = (project.spools || []).reduce((acc, spool) => {
      if (spool.weldingWeek !== weekLabel) return acc;
      return acc + (spool.weldedWeightKg || 0);
    }, 0);
    if (spoolWeight > 0) return total + spoolWeight;
    if (project.weldingWeek !== weekLabel) return total;
    return total + (project.weldedWeightKg || 0);
  }, 0);
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
  const activeWeek = getActiveWeekLabel();
  const weekWeight = getWeldedWeightForWeek(activeWeek);
  document.getElementById("stat-projects").textContent = formatNumber(state.stats.totalProjects);
  document.getElementById("stat-spools").textContent = `${formatNumber(weekWeight, 0)} kg`;
  document.getElementById("stat-total-weight").textContent = `${formatNumber(state.stats.totalWeightKg, 0)} kg`;
  const currentWeekEl = document.getElementById("stat-current-week");
  if (currentWeekEl) currentWeekEl.textContent = activeWeek;

  document.getElementById("stat-completed").textContent = formatNumber(state.stats.completed);
  document.getElementById("stat-in-progress").textContent = formatNumber(state.stats.inProgress);
  document.getElementById("stat-not-started").textContent = formatNumber(state.stats.notStarted);
  document.getElementById("stat-average").textContent = formatPercent(state.stats.averageOverallProgress);
}

function renderTable() {
  if (!state.filteredProjects.length) {
    bodyEl.innerHTML = '<tr><td colspan="16" class="loading-cell">Nenhum projeto encontrado para a busca informada.</td></tr>';
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
          <td>${formatNumber(project.weldedWeightKg, 0)}</td>
          <td>${project.weldingWeek || "—"}</td>
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
        <div class="metric-chip"><span>Peso total soldado</span><strong>${formatNumber(project.weldedWeightKg, 0)} kg</strong></div>
        <div class="metric-chip"><span>Semana finalizado</span><strong>${project.weldingWeek || "—"}</strong></div>
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
          const formatted = value == null || value === "" ? "—" : stage.type === "percent" ? formatPercent(value) : value;
          const cellClass = tableCellClass(value, stage.type);
          return `<td class="${cellClass}">${formatted}</td>`;
        })
        .join("");

      const observations = spool.observations ? escapeHtml(spool.observations).replace(/\n/g, "<br>") : "—";

      return `
        <tr data-modal-row="true">
          <td>${spool.iso || "—"}</td>
          <td>${spool.description || "—"}</td>
          <td class="modal-observation-cell">${observations}</td>
          <td>${formatNumber(spool.weldedWeightKg, 0)} kg</td>
          <td>${spool.weldingWeek || "—"}</td>
          <td>${formatNumber(spool.kilos, 2)}</td>
          <td>${formatNumber(spool.m2Painting, 3)}</td>
          <td><span class="cell-status cell-status--${["awaiting_shipment", "completed"].includes(spool.uiState) ? "completed" : spool.uiState}">${uiStateLabel(spool.uiState)}</span></td>
          <td class="${percentStateClass(spool.stagePercent)}">${spool.stage || "—"}</td>
          <td class="${percentStateClass(spool.individualProgress)}">${formatPercent(spool.individualProgress)}</td>
          <td class="${percentStateClass(spool.overallProgress)}">${formatPercent(spool.overallProgress)}</td>
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
      <article class="metric-chip"><span>Peso total soldado</span><strong>${formatNumber(project.weldedWeightKg, 0)} kg</strong></article>
      <article class="metric-chip"><span>Semana finalizado</span><strong>${project.weldingWeek || "—"}</strong></article>
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
            <th>Observações</th>
            <th>Peso soldado</th>
            <th>Semana finalizado</th>
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

async function fetchProjectsPayload() {
  const sources = [];
  const isGitHubPages = window.location.hostname.endsWith("github.io");

  if (!isGitHubPages) {
    sources.push("/api/projects");
  }

  sources.push("./projects.json");

  if (isGitHubPages) {
    sources.push("projects.json");
  }

  let lastError = null;

  for (const source of sources) {
    try {
      const response = await fetch(source, { cache: "no-store" });
      if (!response.ok) {
        throw new Error(`Falha ao carregar dados (${response.status})`);
      }
      const data = await response.json();
      if (!data?.ok) {
        throw new Error(data?.error || "Falha ao carregar projetos.");
      }
      return data;
    } catch (error) {
      lastError = error;
    }
  }

  throw lastError || new Error("Falha ao carregar projetos.");
}

async function loadProjects() {
  try {
    const data = await fetchProjectsPayload();
    state.projects = enrichProjects(data.projects || []);
    state.stats = buildDerivedStats(state.projects, data.stats || null);
    state.meta = data.meta || null;
    buildDemandOptions();
    buildWeekOptions();

    if (!state.selectedProjectId && state.projects.length) {
      state.selectedProjectId = state.projects[0].rowId;
    }

    applyFilter();
    renderStats();
    renderTable();
    renderSelectedProjectCard();
    updateMeta();
  } catch (error) {
    bodyEl.innerHTML = `<tr><td colspan="16" class="loading-cell">${error.message}</td></tr>`;
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
    state.weekFilter = "";
    searchInputEl.value = "";
    if (demandFilterEl) demandFilterEl.value = "";
    if (weekFilterEl) weekFilterEl.value = "";
    applyFilter();
    renderStats();
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

  if (weekFilterEl) {
    weekFilterEl.addEventListener("change", (event) => {
      state.weekFilter = event.target.value;
      renderStats();
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


  modalContentEl.addEventListener("click", (event) => {
    const row = event.target.closest("tr[data-modal-row='true']");
    if (!row) return;
    modalContentEl.querySelectorAll("tr[data-modal-row='true'].modal-row-selected").forEach((item) => {
      if (item !== row) item.classList.remove("modal-row-selected");
    });
    row.classList.toggle("modal-row-selected");
  });


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
