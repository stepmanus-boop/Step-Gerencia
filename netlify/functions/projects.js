const API_BASE = process.env.SMARTSHEET_API_BASE || "https://api.smartsheet.com/2.0";
const SHEET_NAME = process.env.SMARTSHEET_SHEET_NAME || "Progress Tracking Sheet - Piping Fabrication";
const SHEET_ID_ENV = process.env.SMARTSHEET_SHEET_ID || "";
const TOKEN = process.env.SMARTSHEET_TOKEN || "5pP36OjBaD1W2HWyxf6aoGxXasPvEl8gbqOmQ";

const cache = global.__STEP_PROGRESS_CACHE__ || {
  sheetId: null,
  sheetName: null,
  version: null,
  payload: null,
  lastSync: null,
};
global.__STEP_PROGRESS_CACHE__ = cache;

const STAGE_ORDER = [
  { key: "Drawing Execution Advance%", label: "Drawing Execution", type: "percent" },
  { key: "Procuremnt Status %", label: "Procurement Status", type: "percent" },
  { key: "Material Separation", label: "Material Separation", type: "percent" },
  { key: "Material Release to Fabrication", label: "Material Release to Fabrication", type: "percent" },
  { key: "Fabrication Start Date", label: "Fabrication Start Date", type: "date" },
  { key: "Withdrew Material", label: "Withdrew Material", type: "percent" },
  { key: "Welding Preparation", label: "Welding Preparation", type: "percent" },
  { key: "Spool Assemble and tack weld", label: "Spool Assemble and tack weld", type: "percent" },
  { key: "Boilermaker Finish Date", label: "Boilermaker Finish Date", type: "date" },
  { key: "Initial Dimensional Inspection/3D", label: "Initial Dimensional Inspection/3D", type: "percent" },
  { key: "Full welding execution", label: "Full welding execution", type: "percent" },
  { key: "Welding Finish Date", label: "Welding Finish Date", type: "date" },
  { key: "Final Dimensional Inpection/3D (QC)", label: "Final Dimensional Inpection/3D (QC)", type: "percent" },
  { key: "Non Destructive Examination (QC)", label: "Non Destructive Examination (QC)", type: "percent" },
  { key: "Inspection Finish Date (QC)", label: "Inspection Finish Date (QC)", type: "date" },
  { key: "Hydro Test Pressure (QC)", label: "Hydro Test Pressure (QC)", type: "percent" },
  { key: "TH Finish Date", label: "TH Finish Date", type: "date" },
  { key: "HDG / FBE.  (PAINT)", label: "HDG / FBE. (PAINT)", type: "percent", optional: true },
  { key: "HDG / FBE DATE SAIDA (PAINT)", label: "HDG / FBE DATE SAIDA (PAINT)", type: "date", optional: true },
  { key: "HDG / FBE DATE RETORNO (PAINT)", label: "HDG / FBE DATE RETORNO (PAINT)", type: "date", optional: true },
  { key: "Surface preparation and/or coating", label: "Surface preparation and/or coating", type: "percent" },
  { key: "Coating Finish Date", label: "Coating Finish Date", type: "date", optional: true },
  { key: "Final Inspection", label: "Final Inspection", type: "percent" },
  { key: "Package and Delivered", label: "Package and Delivered", type: "percent" },
  { key: "Project Finish Date", label: "Project Finish Date", type: "date" },
  { key: "Project Finished?", label: "Project Finished?", type: "boolean" },
];

function jsonResponse(statusCode, data) {
  return {
    statusCode,
    headers: {
      "content-type": "application/json; charset=utf-8",
      "cache-control": "no-store",
      "access-control-allow-origin": "*",
    },
    body: JSON.stringify(data),
  };
}

function getCellValue(row, key) {
  return row.values[key] || { raw: null, display: null };
}

function textValue(row, key) {
  const cell = getCellValue(row, key);
  const value = cell.display ?? cell.raw;
  return value == null ? "" : String(value).trim();
}

function parseNumberValue(input) {
  if (input == null || input === "") return null;
  if (typeof input === "number") return Number.isFinite(input) ? input : null;

  let str = String(input).trim();
  if (!str) return null;
  str = str.replace(/\s/g, "");

  const hasComma = str.includes(",");
  const hasDot = str.includes(".");

  if (hasComma && hasDot) {
    if (str.lastIndexOf(",") > str.lastIndexOf(".")) {
      str = str.replace(/\./g, "").replace(",", ".");
    } else {
      str = str.replace(/,/g, "");
    }
  } else if (hasComma) {
    str = str.replace(",", ".");
  }

  str = str.replace(/[^\d.-]/g, "");
  const num = Number(str);
  return Number.isFinite(num) ? num : null;
}

function parseNumber(row, key) {
  const cell = getCellValue(row, key);
  return parseNumberValue(cell.raw ?? cell.display);
}

function parsePercent(row, key) {
  const cell = getCellValue(row, key);
  const display = cell.display ?? "";
  const raw = cell.raw;

  if (typeof display === "string" && display.includes("%")) {
    const value = parseNumberValue(display.replace("%", ""));
    return value == null ? null : value;
  }

  const parsed = parseNumberValue(raw ?? display);
  if (parsed == null) return null;
  if (parsed >= 0 && parsed <= 1) return parsed * 100;
  return parsed;
}

function isTruthyValue(value) {
  if (value == null) return false;
  if (typeof value === "boolean") return value;
  const normalized = String(value).trim().toLowerCase();
  return ["true", "yes", "sim", "y", "1", "concluído", "concluido", "finalizado"].includes(normalized);
}

function excelSerialToDate(serial) {
  if (!Number.isFinite(serial)) return null;
  if (serial < 1 || serial > 90000) return null;
  const excelEpoch = Date.UTC(1899, 11, 30);
  const millis = excelEpoch + Math.round(serial) * 86400000;
  const date = new Date(millis);
  return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
}

function formatDateValue(value) {
  const parsedDate = parseDateObject(value);
  if (parsedDate) {
    return parsedDate.toLocaleDateString("pt-BR", { timeZone: "UTC" });
  }

  if (!value) return "";
  const raw = String(value).trim();
  if (!raw) return "";
  return raw;
}

function parseDateObject(value) {
  if (value == null || value === "") return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(Date.UTC(value.getUTCFullYear(), value.getUTCMonth(), value.getUTCDate()));
  }

  if (typeof value === "number") {
    return excelSerialToDate(value);
  }

  const raw = String(value).trim();
  if (!raw) return null;

  if (/^\d+(?:\.\d+)?$/.test(raw)) {
    const numericDate = excelSerialToDate(Number(raw));
    if (numericDate) return numericDate;
  }

  let match = raw.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (match) {
    const day = Number(match[1]);
    const month = Number(match[2]) - 1;
    const year = Number(match[3]);
    return new Date(Date.UTC(year, month, day));
  }

  match = raw.match(/^(\d{2})\/(\d{2})\/(\d{2})$/);
  if (match) {
    const day = Number(match[1]);
    const month = Number(match[2]) - 1;
    const year = Number(match[3]) >= 70 ? 1900 + Number(match[3]) : 2000 + Number(match[3]);
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

function getWeekAnchor(year) {
  const jan1 = new Date(Date.UTC(year, 0, 1));
  const anchor = new Date(jan1);
  anchor.setUTCDate(jan1.getUTCDate() - jan1.getUTCDay());
  return anchor;
}

function getCurrentBrazilDateObject() {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: "America/Sao_Paulo",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(new Date());
  const year = Number(parts.find((item) => item.type === "year")?.value || new Date().getUTCFullYear());
  const month = Number(parts.find((item) => item.type === "month")?.value || 1);
  const day = Number(parts.find((item) => item.type === "day")?.value || 1);
  return new Date(Date.UTC(year, month - 1, day));
}

function getCurrentBrazilYear() {
  return getCurrentBrazilDateObject().getUTCFullYear();
}

function formatProductionWeekLabel(weekNumber, weekYear) {
  return `Semana ${weekNumber} - ${weekYear}`;
}

function getProductionWeekLabel(value) {
  const date = parseDateObject(value);
  if (!date) return "";

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

function hasDateValue(row, key) {
  const value = textValue(row, key);
  return Boolean(value && String(value).trim());
}

function isAwaitingShipment(row) {
  const coatingPercent = parsePercent(row, "Surface preparation and/or coating") ?? 0;
  const coatingDone = coatingPercent >= 100;
  const packageDelivered = parsePercent(row, "Package and Delivered") ?? 0;
  const projectFinished = isTruthyValue(textValue(row, "Project Finished?") || getCellValue(row, "Project Finished?").raw);
  return coatingDone && packageDelivered < 100 && !projectFinished;
}

function parseProjectParts(projectText) {
  const cleaned = String(projectText || "").trim().replace(/\s+/g, " ");
  if (!cleaned) return { prefix: "", number: "", display: "" };

  const match = cleaned.match(/^(?:([A-Z]{2,5})[\s-]+)?(\d{2}-\d+(?:-\d+)*(?:-[A-Z0-9]+)?)$/i);
  if (match) {
    return { prefix: (match[1] || "").toUpperCase(), number: match[2], display: cleaned };
  }

  const loose = cleaned.match(/([A-Z]{2,5})?[\s-]*(\d{2}-\d+(?:-\d+)*(?:-[A-Z0-9]+)?)/i);
  if (loose) {
    const prefix = loose[1] ? loose[1].toUpperCase() : "";
    const number = loose[2];
    return { prefix, number, display: prefix ? `${prefix} ${number}` : number };
  }

  return { prefix: "", number: cleaned, display: cleaned };
}

function extractIsoDescription(drawingText) {
  const text = String(drawingText || "").trim();
  if (!text) return { iso: "", description: "" };
  const match = text.match(/^(.*?)\s*\((.*?)\)\s*$/);
  if (match) return { iso: match[1].trim(), description: match[2].trim() };
  return { iso: text, description: "" };
}

function stageStatusFromPercent(percent) {
  if (percent == null) return "ignored";
  if (percent >= 100) return "completed";
  if (percent > 0) return "in_progress";
  return "waiting";
}

function buildStageValues(row) {
  const stageValues = {};
  for (const stage of STAGE_ORDER) {
    if (stage.type === "percent") {
      const value = parsePercent(row, stage.key);
      stageValues[stage.key] = value == null ? null : value;
      continue;
    }
    if (stage.type === "date") {
      const value = textValue(row, stage.key);
      stageValues[stage.key] = value ? formatDateValue(value) : "";
      continue;
    }
    if (stage.type === "boolean") {
      stageValues[stage.key] = isTruthyValue(textValue(row, stage.key) || getCellValue(row, stage.key).raw) ? "Sim" : "Não";
    }
  }
  return stageValues;
}

function deriveProgress(row) {
  const milestones = [];
  const completedStages = [];
  let currentStage = null;

  for (const stage of STAGE_ORDER) {
    if (stage.type === "date") {
      const value = textValue(row, stage.key);
      if (value) {
        milestones.push({ key: stage.key, label: stage.label, value: formatDateValue(value), type: "date" });
      }
      continue;
    }

    if (stage.type === "boolean") {
      const truthy = isTruthyValue(textValue(row, stage.key) || getCellValue(row, stage.key).raw);
      milestones.push({ key: stage.key, label: stage.label, value: truthy ? "Sim" : "Não", type: "boolean" });
      if (truthy && !currentStage) {
        currentStage = { key: stage.key, label: stage.label, percent: 100, status: "completed", isAlert: false };
      }
      continue;
    }

    const percent = parsePercent(row, stage.key);
    const hasContent = percent != null || textValue(row, stage.key);
    if (!hasContent && stage.optional) continue;
    if (!hasContent) continue;

    const status = stageStatusFromPercent(percent);
    if (status === "completed") {
      completedStages.push({ key: stage.key, label: stage.label, percent: 100, status });
      continue;
    }

    if (!currentStage) {
      currentStage = {
        key: stage.key,
        label: stage.label,
        percent: percent ?? 0,
        status,
        isAlert: status === "in_progress" || status === "waiting",
      };
    }
  }

  if (!currentStage) {
    currentStage = {
      key: "Package and Delivered",
      label: "Package and Delivered",
      percent: 100,
      status: "completed",
      isAlert: false,
    };
  }

  return { currentStage, completedStages, milestones };
}

function projectUiState(projectStatus, overallProgress, finished, fabricationStartDate, awaitingShipment = false) {
  if (!fabricationStartDate) return "not_started";
  if (awaitingShipment) return "awaiting_shipment";
  if (finished || overallProgress >= 100) return "completed";
  if (overallProgress <= 0 && /^on hold$/i.test(projectStatus || "")) return "not_started";
  if (overallProgress <= 0) return "not_started";
  return "in_progress";
}

function getOperationalFlow(stageValues, fabricationStartDate, coatingPercent, finished, projectStatus) {
  const fabricationStarted = Boolean(fabricationStartDate);
  const boilermakerFinishDate = stageValues["Boilermaker Finish Date"];
  const weldingFinishDate = stageValues["Welding Finish Date"];
  const thFinishDate = stageValues["TH Finish Date"];
  const projectFinishDate = stageValues["Project Finish Date"];

  const withdrewMaterial = Number(stageValues["Withdrew Material"] || 0);
  const weldingPreparation = Number(stageValues["Welding Preparation"] || 0);
  const spoolAssemble = Number(stageValues["Spool Assemble and tack weld"] || 0);
  const fullWelding = Number(stageValues["Full welding execution"] || 0);
  const initialDimensional = Number(stageValues["Initial Dimensional Inspection/3D"] || 0);
  const finalDimensionalQc = Number(stageValues["Final Dimensional Inpection/3D (QC)"] || 0);
  const ndeQc = Number(stageValues["Non Destructive Examination (QC)"] || 0);
  const hydroTestQc = Number(stageValues["Hydro Test Pressure (QC)"] || 0);
  const packageDelivered = Number(stageValues["Package and Delivered"] || 0);
  const projectFinished = String(stageValues["Project Finished?"] || "").toLowerCase() === "sim";

  const delivered = finished || packageDelivered >= 100 || Boolean(projectFinishDate) || projectFinished;
  if (!fabricationStarted) return { state: "not_started", sector: "Geral" };
  if (delivered) return { state: "completed", sector: "Geral" };

  const calderariaComplete = Boolean(boilermakerFinishDate) && withdrewMaterial >= 100 && weldingPreparation >= 100 && spoolAssemble >= 100;
  if (!calderariaComplete) return { state: "in_production", sector: "Calderaria" };

  const soldaComplete = Boolean(weldingFinishDate) && fullWelding >= 100 && initialDimensional >= 100;
  if (!soldaComplete) return { state: "in_production", sector: "Solda" };

  const inspectionComplete = Boolean(thFinishDate) && finalDimensionalQc >= 100 && ndeQc >= 100 && hydroTestQc >= 100;
  if (!inspectionComplete) return { state: "in_inspection", sector: "Inspeção" };

  if (Number(coatingPercent || 0) >= 100) return { state: "awaiting_shipment", sector: "Envio" };
  return { state: "in_production", sector: "Pintura" };
}

function classifyStageSector(stageValue) {
  const stage = String(stageValue || '').toLowerCase();

  if (
    stage.includes('paint') ||
    stage.includes('coating') ||
    stage.includes('surface preparation') ||
    stage.includes('surface preparation and/or coating') ||
    stage.includes('hdg') ||
    stage.includes('fbe')
  ) {
    return 'Pintura';
  }

  if (
    stage.includes('inspection') ||
    stage.includes('nondestructive') ||
    stage.includes('non destructive') ||
    stage.includes('dimensional') ||
    stage.includes('hydro test') ||
    stage.includes('qc') ||
    stage.includes('th finish') ||
    stage.includes('final inspection')
  ) {
    return 'Inspeção';
  }

  if (
    stage.includes('welding') ||
    stage.includes('solda') ||
    stage.includes('spool assemble') ||
    stage.includes('tack weld')
  ) {
    return 'Solda';
  }

  if (
    stage.includes('boilermaker') ||
    stage.includes('caldeiraria') ||
    stage.includes('material release') ||
    stage.includes('material separation') ||
    stage.includes('withdrew material') ||
    stage.includes('drawing execution') ||
    stage.includes('procurement') ||
    stage.includes('fabrication')
  ) {
    return 'Calderaria';
  }

  return 'Geral';
}

function classifyAlertSector(project) {
  if (project?.operationalSector) return project.operationalSector;
  const stageValues = project?.stageValues || {};
  const flow = getOperationalFlow(
    stageValues,
    project?.fabricationStartDate,
    project?.coatingPercent,
    project?.finished,
    project?.projectStatus,
  );
  return flow.sector || 'Geral';
}

function buildAlertObservation(project, sector, diffDays) {
  const stageLabel = project?.currentStage || project?.jobProcessStatus || 'Etapa não identificada';
  const coatingPercent = Number(project?.coatingPercent || 0);
  const baseDaysText = diffDays < 0
    ? `O término planejado já venceu há ${Math.abs(diffDays)} dia(s).`
    : `Faltam ${diffDays} dia(s) para o término planejado.`;

  if (coatingPercent >= 100) {
    return {
      title: diffDays < 0 ? 'Conferência em atraso' : 'Conferência pendente',
      message: `${baseDaysText} A pintura já está em 100%. Conferir envio.`,
    };
  }

  if (sector === 'Calderaria') {
    return {
      title: diffDays < 0 ? 'Calderaria em atraso' : 'Calderaria em atenção',
      message: `${baseDaysText} O projeto ainda está na Calderaria.`,
    };
  }

  if (sector === 'Solda') {
    return {
      title: diffDays < 0 ? 'Solda em atraso' : 'Solda em atenção',
      message: `${baseDaysText} O projeto ainda está em Solda.`,
    };
  }

  if (sector === 'Inspeção') {
    return {
      title: diffDays < 0 ? 'Inspeção em atraso' : 'Inspeção em atenção',
      message: `${baseDaysText} O projeto ainda está na Inspeção, preso em ${stageLabel}.`,
    };
  }

  if (sector === 'Pintura') {
    return {
      title: diffDays < 0 ? 'Pintura em atraso' : 'Pintura em atenção',
      message: `${baseDaysText} O projeto ainda está na Pintura.`,
    };
  }

  return {
    title: diffDays < 0 ? 'Prazo vencido' : 'Prazo próximo',
    message: `${baseDaysText} O projeto segue em andamento.`,
  };
}

function isSummaryRow(row) {
  const projectText = textValue(row, "Project");
  if (!projectText) return false;
  if (row.parentId) return false;

  const quantitySpools = parseNumber(row, "Quantity Spools");
  const drawing = textValue(row, "Drawing");
  const parts = parseProjectParts(projectText);

  return Boolean(parts.prefix && parts.number && (quantitySpools != null || drawing === "ISO" || textValue(row, "Project Type")));
}

function isChildRow(row) {
  if (row.parentId) return true;
  const drawing = textValue(row, "Drawing");
  const projectText = textValue(row, "Project");
  const parts = parseProjectParts(projectText);
  return Boolean(!parts.prefix && parts.number && drawing && drawing !== "ISO");
}

function buildSpoolRow(row, parentSummary) {
  const drawingText = textValue(row, "Drawing");
  const parsedDrawing = extractIsoDescription(drawingText);
  const progress = deriveProgress(row);
  const overallProgress = parsePercent(row, "% Overall Progress") ?? parsePercent(parentSummary, "% Overall Progress") ?? 0;
  const individualProgress = parsePercent(row, "% Individual Progress") ?? overallProgress;
  const finished = isTruthyValue(getCellValue(row, "Project Finished?").raw) || overallProgress >= 100 || (parsePercent(row, "Package and Delivered") ?? 0) >= 100;
  const fabricationStartDate = textValue(row, "Fabrication Start Date");
  const stageValues = buildStageValues(row);
  const flow = getOperationalFlow(stageValues, fabricationStartDate, parsePercent(row, "Surface preparation and/or coating") ?? 0, finished, textValue(row, "PROJECT STATUS"));
  const awaitingShipment = flow.state === "awaiting_shipment";
  const uiState = flow.state === "in_inspection" ? "in_progress" : projectUiState(textValue(row, "PROJECT STATUS"), overallProgress, finished, fabricationStartDate, awaitingShipment);
  const coatingPercent = parsePercent(row, "Surface preparation and/or coating") ?? 0;
  const weldingPercent = parsePercent(row, "Full welding execution") ?? 0;
  const weldingFinishDate = textValue(row, "Welding Finish Date");
  const weldedWeightKg = (() => {
    const kilos = parseNumber(row, "Kilos");
    if (kilos == null) return null;
    if (weldingPercent >= 100) return kilos;
    if (weldingPercent > 0) return (kilos * weldingPercent) / 100;
    return 0;
  })();
  const weldingWeek = weldingPercent >= 100 && weldingFinishDate ? getProductionWeekLabel(weldingFinishDate) : "";

  return {
    rowId: row.id,
    rowNumber: row.rowNumber,
    iso: parsedDrawing.iso,
    description: parsedDrawing.description,
    drawing: drawingText,
    observations: textValue(row, "OBSERVATIONS"),
    operationalSector: flow.sector,
    operationalState: flow.state,
    plannedStartDate: formatDateValue(textValue(row, "Start Date")),
    plannedFinishDate: formatDateValue(textValue(row, "Finish Date")),
    kilos: parseNumber(row, "Kilos"),
    weldedWeightKg,
    weldingWeek,
    coatingPercent,
    m2Painting: parseNumber(row, "M2 Painting"),
    stage: progress.currentStage.label,
    stagePercent: progress.currentStage.percent,
    stageStatus: progress.currentStage.status,
    stageAlert: progress.currentStage.isAlert,
    individualProgress,
    overallProgress,
    milestones: progress.milestones,
    stageValues,
    finished: finished,
    uiState,
  };
}

function buildProject(summaryRow, childRows) {
  const projectText = textValue(summaryRow, "Project");
  const parts = parseProjectParts(projectText);
  const progress = deriveProgress(summaryRow);
  const overallProgress = parsePercent(summaryRow, "% Overall Progress") ?? 0;
  const individualProgress = parsePercent(summaryRow, "% Individual Progress") ?? overallProgress;
  const finished = isTruthyValue(getCellValue(summaryRow, "Project Finished?").raw) || overallProgress >= 100 || (parsePercent(summaryRow, "Package and Delivered") ?? 0) >= 100;
  const projectStatus = textValue(summaryRow, "PROJECT STATUS") || textValue(summaryRow, "Overall Project Status") || textValue(summaryRow, "Status");
  const coatingPercent = parsePercent(summaryRow, "Surface preparation and/or coating") ?? 0;
  const fabricationStartDate = textValue(summaryRow, "Fabrication Start Date");
  const stageValues = buildStageValues(summaryRow);
  const flow = getOperationalFlow(stageValues, fabricationStartDate, coatingPercent, finished, projectStatus);
  const awaitingShipment = flow.state === "awaiting_shipment";
  const uiState = flow.state === "in_inspection" ? "in_progress" : projectUiState(projectStatus, overallProgress, finished, fabricationStartDate, awaitingShipment);
  const weldingPercent = parsePercent(summaryRow, "Full welding execution") ?? 0;
  const weldingFinishDate = textValue(summaryRow, "Welding Finish Date");
  const spools = childRows.map((row) => buildSpoolRow(row, summaryRow));
  const summaryWeldedWeightKg = (() => {
    const kilos = parseNumber(summaryRow, "Kilos");
    if (kilos == null) return null;
    if (weldingPercent >= 100) return kilos;
    if (weldingPercent > 0) return (kilos * weldingPercent) / 100;
    return 0;
  })();
  const weldedWeightKg = spools.length
    ? spools.reduce((total, spool) => total + (spool.weldingWeek ? (spool.weldedWeightKg || 0) : 0), 0)
    : summaryWeldedWeightKg;
  const weldingWeek = weldingPercent >= 100 && weldingFinishDate ? getProductionWeekLabel(weldingFinishDate) : "";

  const spoolStats = spools.reduce((acc, spool) => {
    acc.total += 1;
    if (spool.uiState === "completed") acc.completed += 1;
    else if (spool.uiState === "in_progress") acc.inProgress += 1;
    else acc.notStarted += 1;
    return acc;
  }, { total: 0, completed: 0, inProgress: 0, notStarted: 0 });

  const operationalSector = flow.sector;

  return {
    rowId: summaryRow.id,
    rowNumber: summaryRow.rowNumber,
    projectPrefix: parts.prefix,
    projectNumber: parts.number,
    projectDisplay: parts.display || projectText,
    quantitySpools: parseNumber(summaryRow, "Quantity Spools") ?? spools.length,
    kilos: parseNumber(summaryRow, "Kilos"),
    weldedWeightKg,
    weldingWeek,
    coatingPercent,
    m2Painting: parseNumber(summaryRow, "M2 Painting"),
    currentStage: progress.currentStage.label,
    currentStagePercent: progress.currentStage.percent,
    currentStageStatus: progress.currentStage.status,
    currentStageAlert: progress.currentStage.isAlert,
    individualProgress,
    overallProgress,
    projectStatus,
    jobProcessStatus: textValue(summaryRow, "Job Process Status") || progress.currentStage.label,
    summaryDrawing: textValue(summaryRow, "Drawing"),
    projectType: textValue(summaryRow, "Project Type"),
    fabricationStartDate: formatDateValue(textValue(summaryRow, "Fabrication Start Date")),
    plannedStartDate: formatDateValue(textValue(summaryRow, "Start Date")),
    plannedFinishDate: formatDateValue(textValue(summaryRow, "Finish Date")),
    client: textValue(summaryRow, "Client"),
    vessel: textValue(summaryRow, "Vessel"),
    className: textValue(summaryRow, "Class"),
    milestones: progress.milestones,
    stageValues,
    finished: finished,
    uiState,
    operationalSector,
    operationalState: flow.state,
    spools,
    spoolStats,
  };
}

function mapApiRows(sheet) {
  const columnMap = new Map((sheet.columns || []).map((column) => [column.id, column.title]));
  return (sheet.rows || []).map((row) => {
    const values = {};
    for (const cell of row.cells || []) {
      const title = columnMap.get(cell.columnId);
      if (!title) continue;
      values[title] = { raw: cell.value ?? null, display: cell.displayValue ?? null };
    }
    return {
      id: row.id,
      rowNumber: row.rowNumber,
      parentId: row.parentId ?? null,
      siblingId: row.siblingId ?? null,
      expanded: row.expanded ?? null,
      values,
    };
  });
}

function buildProjects(rows) {
  const projects = [];
  const rowsById = new Map(rows.map((row) => [row.id, row]));
  const childrenByParent = new Map();

  for (const row of rows) {
    if (row.parentId && rowsById.has(row.parentId)) {
      if (!childrenByParent.has(row.parentId)) childrenByParent.set(row.parentId, []);
      childrenByParent.get(row.parentId).push(row);
    }
  }

  let currentSummary = null;

  for (const row of rows) {
    if (isSummaryRow(row)) {
      const directChildren = childrenByParent.get(row.id) || [];
      currentSummary = row;
      projects.push(buildProject(row, directChildren));
      continue;
    }

    if (!currentSummary) continue;
    if (!isChildRow(row)) continue;

    const currentProjectNumber = parseProjectParts(textValue(currentSummary, "Project")).number;
    const childProjectNumber = parseProjectParts(textValue(row, "Project")).number;
    if (!childProjectNumber || childProjectNumber !== currentProjectNumber) continue;

    const lastProject = projects[projects.length - 1];
    if (!lastProject) continue;
    const spool = buildSpoolRow(row, currentSummary);
    lastProject.spools.push(spool);
    lastProject.spoolStats.total += 1;
    if (spool.uiState === "completed") lastProject.spoolStats.completed += 1;
    else if (spool.uiState === "in_progress") lastProject.spoolStats.inProgress += 1;
    else lastProject.spoolStats.notStarted += 1;
  }

  for (const project of projects) {
    const unique = [];
    const seen = new Set();
    for (const spool of project.spools) {
      const key = `${spool.rowId || ""}-${spool.iso}`;
      if (seen.has(key)) continue;
      seen.add(key);
      unique.push(spool);
    }
    project.spools = unique;
    project.spoolStats = unique.reduce((acc, spool) => {
      acc.total += 1;
      if (spool.uiState === "completed") acc.completed += 1;
      else if (spool.uiState === "in_progress") acc.inProgress += 1;
      else acc.notStarted += 1;
      return acc;
    }, { total: 0, completed: 0, inProgress: 0, notStarted: 0 });
  }

  return projects;
}

function getProjectAlert(project, today = getCurrentBrazilDateObject()) {
  if (!project.fabricationStartDate) return null;
  const plannedFinish = parseDateObject(project.plannedFinishDate);
  if (!plannedFinish) return null;

  const diffDays = Math.floor((plannedFinish - today) / 86400000);
  const coatingPercent = Number(project.coatingPercent || 0);
  const sector = classifyAlertSector(project);
  const observation = buildAlertObservation(project, sector, diffDays);

  if (coatingPercent < 100 && diffDays <= 5) {
    return {
      projectDisplay: project.projectDisplay,
      projectNumber: project.projectNumber,
      projectRowId: project.rowId,
      client: project.client,
      sector,
      plannedFinishDate: project.plannedFinishDate,
      daysRemaining: diffDays,
      type: diffDays < 0 ? "overdue" : "deadline",
      title: observation.title,
      message: observation.message,
      coatingPercent,
      currentStage: project.currentStage,
    };
  }

  if (coatingPercent >= 100 && diffDays <= 3) {
    return {
      projectDisplay: project.projectDisplay,
      projectNumber: project.projectNumber,
      projectRowId: project.rowId,
      client: project.client,
      sector,
      plannedFinishDate: project.plannedFinishDate,
      daysRemaining: diffDays,
      type: diffDays < 0 ? "conference_overdue" : "conference",
      title: observation.title,
      message: observation.message,
      coatingPercent,
      currentStage: project.currentStage,
    };
  }

  return null;
}

function buildAlerts(projects) {
  const alerts = projects
    .map((project) => getProjectAlert(project))
    .filter(Boolean)
    .sort((a, b) => {
      if (a.daysRemaining !== b.daysRemaining) return a.daysRemaining - b.daysRemaining;
      return String(a.projectDisplay || "").localeCompare(String(b.projectDisplay || ""), "pt-BR");
    });

  const signature = alerts
    .map((alert) => [alert.projectDisplay, alert.type, alert.plannedFinishDate, alert.daysRemaining].join("|"))
    .join("||");

  return { alerts, signature };
}

function buildStats(projects) {
  const stats = {
    totalProjects: projects.length,
    totalSpools: 0,
    totalWeightKg: 0,
    totalWeldedWeightKg: 0,
    totalPaintingM2: 0,
    completed: 0,
    completedTags: 0,
    inProgress: 0,
    inProgressTags: 0,
    inspectionProjects: 0,
    inspectionTags: 0,
    awaitingShipment: 0,
    awaitingShipmentTags: 0,
    notStarted: 0,
    notStartedTags: 0,
    averageOverallProgress: 0,
  };

  let progressAccumulator = 0;

  for (const project of projects) {
    const tags = Number(project.quantitySpools || 0);
    stats.totalSpools += tags;
    stats.totalWeightKg += project.kilos || 0;
    stats.totalWeldedWeightKg += project.weldedWeightKg || 0;
    stats.totalPaintingM2 += project.m2Painting || 0;
    progressAccumulator += project.overallProgress || 0;

    const state = project.operationalState || project.uiState;
    if (state === "completed") {
      stats.completed += 1;
      stats.completedTags += tags;
    } else if (state === "awaiting_shipment") {
      stats.awaitingShipment += 1;
      stats.awaitingShipmentTags += tags;
      stats.completed += 1;
      stats.completedTags += tags;
    } else if (state === "in_inspection") {
      stats.inspectionProjects += 1;
      stats.inspectionTags += tags;
    } else if (state === "in_production") {
      stats.inProgress += 1;
      stats.inProgressTags += tags;
    } else {
      stats.notStarted += 1;
      stats.notStartedTags += tags;
    }
  }

  stats.averageOverallProgress = projects.length ? progressAccumulator / projects.length : 0;
  return stats;
}

async function apiFetch(path) {
  const response = await fetch(`${API_BASE}${path}`, {
    headers: {
      Authorization: `Bearer ${TOKEN}`,
      "Content-Type": "application/json",
    },
  });

  if (!response.ok) {
    const message = await response.text();
    throw new Error(`Smartsheet ${response.status}: ${message}`);
  }

  return response.json();
}

function normalizeName(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9]+/g, " ")
    .trim()
    .toLowerCase();
}

async function resolveSheetId() {
  if (cache.sheetId) return cache.sheetId;
  if (SHEET_ID_ENV) {
    cache.sheetId = SHEET_ID_ENV;
    return cache.sheetId;
  }

  const target = normalizeName(SHEET_NAME);
  let page = 1;
  let fuzzyFound = null;

  while (true) {
    const response = await apiFetch(`/sheets?page=${page}&pageSize=100`);
    const items = response.data || [];

    const exactFound = items.find((item) => normalizeName(item.name) === target);
    if (exactFound) {
      cache.sheetId = String(exactFound.id);
      cache.sheetName = exactFound.name;
      return cache.sheetId;
    }

    if (!fuzzyFound) {
      fuzzyFound = items.find((item) => normalizeName(item.name).includes(target) || target.includes(normalizeName(item.name)));
    }

    if (!items.length || page >= (response.totalPages || 1)) break;
    page += 1;
  }

  if (fuzzyFound) {
    cache.sheetId = String(fuzzyFound.id);
    cache.sheetName = fuzzyFound.name;
    return cache.sheetId;
  }

  throw new Error(`Sheet "${SHEET_NAME}" não encontrada. Defina SMARTSHEET_SHEET_ID ou confira SMARTSHEET_SHEET_NAME.`);
}

async function fetchSheetVersion(sheetId) {
  const versionData = await apiFetch(`/sheets/${sheetId}/version`);
  return versionData.version;
}

async function fetchFullSheet(sheetId) {
  return apiFetch(`/sheets/${sheetId}?includeAll=true`);
}

async function buildPayload() {
  if (!TOKEN) throw new Error("SMARTSHEET_TOKEN não configurado.");

  const sheetId = await resolveSheetId();
  const version = await fetchSheetVersion(sheetId);

  if (cache.payload && cache.version === version) {
    return cache.payload;
  }

  const sheet = await fetchFullSheet(sheetId);
  const rows = mapApiRows(sheet);
  const projects = buildProjects(rows);
  const stats = buildStats(projects);
  const alertData = buildAlerts(projects);

  const payload = {
    ok: true,
    meta: {
      sheetId,
      sheetName: sheet.name || cache.sheetName || SHEET_NAME,
      version,
      lastSync: new Date().toISOString(),
      stageOrder: STAGE_ORDER.map((stage) => ({
        key: stage.key,
        label: stage.label,
        type: stage.type,
        optional: Boolean(stage.optional),
      })),
      currentWeek: getProductionWeekLabel(getCurrentBrazilDateObject()),
      alertSignature: alertData.signature,
    },
    stats,
    alerts: alertData.alerts,
    projects,
  };

  cache.sheetId = sheetId;
  cache.sheetName = payload.meta.sheetName;
  cache.version = version;
  cache.lastSync = payload.meta.lastSync;
  cache.payload = payload;

  return payload;
}

exports.handler = async () => {
  try {
    const payload = await buildPayload();
    return jsonResponse(200, payload);
  } catch (error) {
    return jsonResponse(500, { ok: false, error: error.message });
  }
};
