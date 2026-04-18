const express = require("express");
const fs = require("fs");
const os = require("os");
const path = require("path");
const multer = require("multer");
const XLSX = require("xlsx");
const ExcelJS = require("exceljs");
const { createClient } = require("@supabase/supabase-js");

const app = express();
const PORT = Number(process.env.PORT || 3000);
const HOST = process.env.HOST || "0.0.0.0";

const IS_SERVERLESS = Boolean(
  process.env.NETLIFY || process.env.AWS_LAMBDA_FUNCTION_NAME || process.env.LAMBDA_TASK_ROOT
);

// In Netlify/AWS Lambda the code directory is read-only (/var/task).
// Use /tmp for any writes (uploads, temp files).
const DATA_DIR = path.join(__dirname, "data");
const UPLOADS_DIR = IS_SERVERLESS
  ? path.join(os.tmpdir(), "roster-uploads")
  : path.join(__dirname, "uploads");
const STORE_PATH = path.join(DATA_DIR, "store.json");
const EXPORT_TEMPLATE_CANDIDATES = [
  path.join(DATA_DIR, "export-template.xlsx"),
  path.join(__dirname, "MOVIMIENTOS PERSONAL CALIDAD 06.03.06.xlsx"),
  path.join(__dirname, "MOVIMIENTOS PERSONAL CALIDAD (versión 02).xlsx"),
  path.join(
    process.env.USERPROFILE || "",
    "Downloads",
    "MOVIMIENTOS PERSONAL CALIDAD 06.03.06.xlsx"
  ),
];

const DEFAULT_CODES = ["1", "B", "D", "S"];
const DIFFUSION_SITE = "RIO TINTO CAMP 1500";
const SUPABASE_TABLE = process.env.SUPABASE_TABLE || "roster_store";
const STORE_BACKEND = (process.env.STORE_BACKEND || "file").toLowerCase(); // "file" | "supabase"
const upload = multer({ dest: UPLOADS_DIR });

function getLanIps() {
  const result = [];
  const nets = os.networkInterfaces();
  for (const name of Object.keys(nets)) {
    for (const net of nets[name] || []) {
      if (!net || net.internal) continue;
      if (net.family !== "IPv4") continue;
      result.push(net.address);
    }
  }
  return Array.from(new Set(result));
}

function ensureDirs() {
  // DATA_DIR may be read-only in serverless; only ensure when we need file backend.
  if (!IS_SERVERLESS) {
    fs.mkdirSync(DATA_DIR, { recursive: true });
  }
  fs.mkdirSync(UPLOADS_DIR, { recursive: true });
}

function normalizeText(value) {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function cleanText(value) {
  return String(value ?? "").replace(/\s+/g, " ").trim();
}

function getHiddenEmployeeSet(store) {
  const raw = store?.meta?.hiddenEmployees;
  const items = Array.isArray(raw) ? raw : [];
  const set = new Set();
  for (const item of items) {
    const norm = normalizeText(item);
    if (norm) set.add(norm);
  }
  return set;
}

function isEmployeeHidden(store, emp) {
  const hidden = getHiddenEmployeeSet(store);
  return hidden.has(normalizeText(emp?.id)) || hidden.has(normalizeText(emp?.name));
}

function getVisibleEmployees(store) {
  const employees = Array.isArray(store?.employees) ? store.employees : [];
  const hidden = getHiddenEmployeeSet(store);
  if (!hidden.size) return employees;
  return employees.filter(
    (emp) => !hidden.has(normalizeText(emp?.id)) && !hidden.has(normalizeText(emp?.name))
  );
}

function getRoleKind(categoria) {
  const normalized = normalizeText(categoria);
  if (normalized.includes("coordinador")) return "coordinador";
  if (normalized.includes("inspector")) return "inspector";
  return "otro";
}

function toIsoDate(date) {
  return date.toISOString().slice(0, 10);
}

function addDaysToIso(isoDate, days) {
  const date = new Date(`${isoDate}T00:00:00Z`);
  date.setUTCDate(date.getUTCDate() + days);
  return toIsoDate(date);
}

function diffDaysIso(fromIso, toIso) {
  const from = new Date(`${fromIso}T00:00:00Z`);
  const to = new Date(`${toIso}T00:00:00Z`);
  return Math.round((to - from) / (1000 * 60 * 60 * 24));
}

function formatDateEs(isoDate) {
  if (!isoDate) return "-";
  return new Date(`${isoDate}T00:00:00Z`).toLocaleDateString("es-AR", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    timeZone: "UTC",
  });
}

function formatDateShort(isoDate) {
  if (!isoDate) return "-";
  return new Date(`${isoDate}T00:00:00Z`).toLocaleDateString("es-AR", {
    day: "2-digit",
    month: "2-digit",
    timeZone: "UTC",
  });
}

function normalizeCategoriaValue(value) {
  const text = cleanText(value).toUpperCase();
  const normalized = text
    .replace(/COORDIANDOR/g, "COORDINADOR")
    .replace(/COORIDNADOR/g, "COORDINADOR")
    .replace(/COORRDINADOR/g, "COORDINADOR");
  if (normalized.includes("SUPERVISOR")) {
    return "INSPECTOR";
  }
  return normalized;
}

function normalizeNameValue(value) {
  const text = cleanText(value).toUpperCase();
  if (text === "RODRIGO AGUIRRE") return "RODRIGO AGUIERRE";
  return text;
}

function sortByName(items) {
  return [...items].sort((a, b) =>
    cleanText(a.name).localeCompare(cleanText(b.name), "es", {
      sensitivity: "base",
    })
  );
}

function applyBusinessOverrides(employee) {
  const normalizedName = normalizeText(employee.name);
  employee.name = normalizeNameValue(employee.name);
  employee.categoria = normalizeCategoriaValue(employee.categoria);

  const categoryByName = {
    "rodrigo tarcaya": "INSPECTOR",
    "guadalupe chavez": "INSPECTOR",
    "silvina benitez": "INSPECTOR",
    "josefina singh": "COORDINADOR QA/QC",
  };

  if (categoryByName[normalizedName]) {
    employee.categoria = categoryByName[normalizedName];
  }
}

function normalizeDateColumns(rawColumns) {
  const ordered = [...rawColumns].sort((a, b) => a.col - b.col);
  if (ordered.length < 2) {
    return { dateColumns: ordered, gaps: [], wasFixed: false };
  }

  const gaps = [];
  for (let i = 1; i < ordered.length; i += 1) {
    const diff = diffDaysIso(ordered[i - 1].date, ordered[i].date);
    if (diff !== 1) {
      gaps.push({
        prev: ordered[i - 1].date,
        current: ordered[i].date,
        diff,
      });
    }
  }

  if (!gaps.length) {
    return { dateColumns: ordered, gaps, wasFixed: false };
  }

  const first = ordered[0].date;
  const fixed = ordered.map((col, index) => ({
    ...col,
    date: addDaysToIso(first, index),
  }));

  return { dateColumns: fixed, gaps, wasFixed: true };
}

function normalizeStoreCalendarContinuity(store) {
  const dates = sortIsoDates(store.meta?.dates || []);
  if (dates.length < 2) return false;

  let hasGap = false;
  for (let i = 1; i < dates.length; i += 1) {
    if (diffDaysIso(dates[i - 1], dates[i]) !== 1) {
      hasGap = true;
      break;
    }
  }
  if (!hasGap) return false;

  const first = dates[0];
  const remap = new Map();
  const fixedDates = dates.map((oldDate, index) => {
    const newDate = addDaysToIso(first, index);
    remap.set(oldDate, newDate);
    return newDate;
  });

  const fixedShifts = {};
  for (const [employeeId, shiftMap] of Object.entries(store.shifts || {})) {
    fixedShifts[employeeId] = {};
    for (const [oldDate, code] of Object.entries(shiftMap || {})) {
      const newDate = remap.get(oldDate) || oldDate;
      fixedShifts[employeeId][newDate] = code;
    }
  }

  store.shifts = fixedShifts;
  store.meta.dates = fixedDates;
  store.meta.dateRange = {
    from: fixedDates[0] || null,
    to: fixedDates[fixedDates.length - 1] || null,
  };
  store.audit.push({
    at: new Date().toISOString(),
    action: "calendar-gap-auto-fix",
    message: "Se corrigio un salto de fechas para mantener calendario continuo.",
  });
  store.audit = store.audit.slice(-5000);
  return true;
}

function ensureStoreHasYearDates(store, year) {
  if (!store || !store.meta) return false;
  const start = `${year}-01-01`;
  const end = `${year}-12-31`;

  const existing = new Set(Array.isArray(store.meta.dates) ? store.meta.dates : []);
  let changed = false;

  let date = start;
  while (date <= end) {
    if (!existing.has(date)) {
      existing.add(date);
      (store.meta.dates ||= []).push(date);
      changed = true;
    }
    date = addDaysToIso(date, 1);
  }

  if (!changed) return false;

  store.meta.dates = sortIsoDates(store.meta.dates || []);
  store.meta.dateRange = {
    from: store.meta.dates[0] || null,
    to: store.meta.dates[store.meta.dates.length - 1] || null,
  };
  (store.audit ||= []).push({
    at: new Date().toISOString(),
    action: "extend-calendar-year",
    message: `Se agregaron fechas faltantes para el anio ${year}.`,
  });
  store.audit = (store.audit || []).slice(-5000);
  return true;
}

function normalizeStoredEmployees(store) {
  let changed = false;
  for (const emp of store.employees || []) {
    const prev = JSON.stringify({
      name: emp.name,
      categoria: emp.categoria,
    });
    applyBusinessOverrides(emp);
    const next = JSON.stringify({
      name: emp.name,
      categoria: emp.categoria,
    });
    if (prev !== next) {
      changed = true;
    }
  }

  const sorted = sortByName(store.employees || []);
  if (
    JSON.stringify(sorted.map((item) => item.id)) !==
    JSON.stringify((store.employees || []).map((item) => item.id))
  ) {
    store.employees = sorted;
    changed = true;
  }

  if (changed) {
    store.audit.push({
      at: new Date().toISOString(),
      action: "normalize-employees",
      message: "Se normalizaron nombres/categorias y orden de empleados.",
    });
    store.audit = store.audit.slice(-5000);
  }
  return changed;
}

function parseDateValue(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return toIsoDate(value);
  }

  if (typeof value === "number" && Number.isFinite(value)) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed && parsed.y && parsed.m && parsed.d) {
      const date = new Date(Date.UTC(parsed.y, parsed.m - 1, parsed.d));
      if (!Number.isNaN(date.getTime())) {
        return toIsoDate(date);
      }
    }
  }

  const text = cleanText(value);
  if (!text) {
    return null;
  }

  const slashMatch = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (slashMatch) {
    const day = Number(slashMatch[1]);
    const month = Number(slashMatch[2]);
    const year = Number(slashMatch[3]);
    const date = new Date(Date.UTC(year, month - 1, day));
    if (
      !Number.isNaN(date.getTime()) &&
      date.getUTCFullYear() === year &&
      date.getUTCMonth() === month - 1 &&
      date.getUTCDate() === day
    ) {
      return toIsoDate(date);
    }
  }

  const isoMatch = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoMatch) {
    return text;
  }

  return null;
}

function isDateLikeString(value) {
  return /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.test(cleanText(value));
}

function toCsv(rows, delimiter = ",") {
  return rows
    .map((row) =>
      row
        .map((cell) => {
          const text = String(cell ?? "");
          if (text.includes(delimiter) || text.includes('"') || text.includes("\n")) {
            return `"${text.replace(/"/g, '""')}"`;
          }
          return text;
        })
        .join(delimiter)
    )
    .join("\n");
}

function weekdayEsShort(isoDate) {
  return new Date(`${isoDate}T00:00:00Z`).toLocaleDateString("es-AR", {
    weekday: "short",
    timeZone: "UTC",
  });
}

function excelSerialFromIso(isoDate) {
  if (!isoDate) return null;
  const utcMillis = Date.parse(`${isoDate}T00:00:00Z`);
  if (!Number.isFinite(utcMillis)) return null;
  const excelEpochUtc = Date.UTC(1899, 11, 30);
  return Math.floor((utcMillis - excelEpochUtc) / (1000 * 60 * 60 * 24));
}

function resolveExportTemplatePath() {
  for (const candidate of EXPORT_TEMPLATE_CANDIDATES) {
    if (candidate && fs.existsSync(candidate)) {
      return candidate;
    }
  }
  return null;
}

function normalizeHeaderKey(value) {
  return normalizeText(value)
    .replace(/[^\w\s/+-]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function findCellByText(ws, targetText, range) {
  const target = normalizeHeaderKey(targetText);
  for (let r = range.s.r; r <= range.e.r; r += 1) {
    for (let c = range.s.c; c <= range.e.c; c += 1) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const raw = ws[addr]?.w ?? ws[addr]?.v;
      if (normalizeHeaderKey(raw) === target) {
        return { r, c };
      }
    }
  }
  return null;
}

function setSheetCell(ws, row, col, value, explicitType) {
  const addr = XLSX.utils.encode_cell({ r: row, c: col });
  const existing = ws[addr] || {};

  if (value === null || value === undefined || value === "") {
    ws[addr] = { ...existing, t: "s", v: "" };
    delete ws[addr].w;
    return;
  }

  const inferredType = explicitType || (typeof value === "number" ? "n" : "s");
  ws[addr] = { ...existing, t: inferredType, v: value };
  delete ws[addr].w;
}

function ensureSheetRangeIncludes(ws, row, col) {
  const currentRange = XLSX.utils.decode_range(ws["!ref"] || "A1");
  const next = {
    s: {
      r: Math.min(currentRange.s.r, row),
      c: Math.min(currentRange.s.c, col),
    },
    e: {
      r: Math.max(currentRange.e.r, row),
      c: Math.max(currentRange.e.c, col),
    },
  };
  ws["!ref"] = XLSX.utils.encode_range(next);
}

function getTemplateHeaderColumns(ws, headerRow, range) {
  const columns = {};
  const expected = {
    sup: ["sup"],
    esp: ["esp"],
    name: ["apellido y nombre", "nombre"],
    dni: ["dni"],
    telefono: ["telefono"],
    ascensoDescenso: ["punto ascenso/descenso", "subida/bajada", "punto ascenso descenso"],
    categoria: ["categoria real", "categoria"],
    grupo: ["grupo"],
    tipo: ["tipo"],
    residencia: ["residencia"],
  };

  for (let c = range.s.c; c <= range.e.c; c += 1) {
    const addr = XLSX.utils.encode_cell({ r: headerRow, c });
    const header = normalizeHeaderKey(ws[addr]?.w ?? ws[addr]?.v);
    if (!header) continue;
    for (const [key, aliases] of Object.entries(expected)) {
      if (!columns[key] && aliases.includes(header)) {
        columns[key] = c;
      }
    }
  }

  return columns;
}

function getDateColumnBounds(ws, headerRow, range, fallbackStartCol) {
  let first = null;
  let last = null;

  for (let c = fallbackStartCol; c <= range.e.c; c += 1) {
    const addr = XLSX.utils.encode_cell({ r: headerRow, c });
    const parsed = parseDateValue(ws[addr]?.v ?? ws[addr]?.w);
    if (!parsed) continue;
    if (first === null) first = c;
    last = c;
  }

  if (first === null) {
    first = fallbackStartCol;
    last = fallbackStartCol;
  }

  return { first, last };
}

function getMetricRows(ws, range, dateStartCol) {
  const matches = {
    direct: ["directos en sitio por dia", "coordinadores en sitio por dia"],
    indirect: ["indirectos en sitio por dia", "inspectores en sitio por dia"],
    total: [
      "total: directos + indirectos en sitio por dia",
      "total: coordinadores + inspectores en sitio por dia",
    ],
    percentage: [
      "porcentaje de locales en sitio",
      "porcentaje coordinsp en sitio",
      "porcentaje coord+insp en sitio",
    ],
  };
  const normalizedMatches = Object.fromEntries(
    Object.entries(matches).map(([key, aliases]) => [
      key,
      aliases.map((alias) => normalizeHeaderKey(alias)),
    ])
  );
  const rows = {};
  const searchLastCol = Math.max(0, dateStartCol - 1);

  for (let r = range.s.r; r <= range.e.r; r += 1) {
    for (let c = range.s.c; c <= searchLastCol; c += 1) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const text = normalizeHeaderKey(ws[addr]?.w ?? ws[addr]?.v);
      if (!text) continue;
      for (const [key, aliases] of Object.entries(normalizedMatches)) {
        if (!rows[key] && aliases.includes(text)) {
          rows[key] = r;
        }
      }
    }
  }

  return rows;
}

function buildTemplateRosterWorkbook({ store, employees, dates, onsiteCodes }) {
  const templatePath = resolveExportTemplatePath();
  if (!templatePath) return null;

  const wb = XLSX.readFile(templatePath, {
    cellStyles: true,
    cellDates: true,
    cellNF: true,
  });
  const sheetName = wb.SheetNames[0];
  if (!sheetName) return null;
  const ws = wb.Sheets[sheetName];
  if (!ws) return null;

  const range = XLSX.utils.decode_range(ws["!ref"] || "A1");
  const nameHeader = findCellByText(ws, "APELLIDO Y NOMBRE", range);
  if (!nameHeader) return null;

  const headerRow = nameHeader.r;
  const weekdayRow = Math.max(range.s.r, headerRow - 1);
  const headerColumns = getTemplateHeaderColumns(ws, headerRow, range);
  const fallbackDateStart = (headerColumns.residencia ?? nameHeader.c) + 1;
  const dateBounds = getDateColumnBounds(ws, headerRow, range, fallbackDateStart);
  const dateStartCol = dateBounds.first;
  const dateLastRequired = dateStartCol + Math.max(0, dates.length - 1);
  const dateEndCol = Math.max(dateBounds.last, dateLastRequired);
  const metricRows = getMetricRows(ws, range, dateStartCol);
  const firstMetricRow = Math.min(
    ...Object.values(metricRows).filter((v) => Number.isFinite(v)),
    Number.POSITIVE_INFINITY
  );
  const employeeStartRow = headerRow + 1;
  const employeeEndRow = Number.isFinite(firstMetricRow)
    ? Math.max(employeeStartRow, firstMetricRow - 1)
    : employeeStartRow + Math.max(employees.length, 1) - 1;
  const fixedCols = [
    headerColumns.sup,
    headerColumns.esp,
    headerColumns.name ?? nameHeader.c,
    headerColumns.dni,
    headerColumns.telefono,
    headerColumns.ascensoDescenso,
    headerColumns.categoria,
    headerColumns.grupo,
    headerColumns.tipo,
    headerColumns.residencia,
  ].filter((c) => Number.isInteger(c));

  for (let c = dateStartCol; c <= dateEndCol; c += 1) {
    const index = c - dateStartCol;
    const dateIso = dates[index];
    if (dateIso) {
      const serial = excelSerialFromIso(dateIso);
      setSheetCell(ws, weekdayRow, c, weekdayEsShort(dateIso), "s");
      setSheetCell(ws, headerRow, c, serial, "n");
      const dateAddr = XLSX.utils.encode_cell({ r: headerRow, c });
      if (!ws[dateAddr].z) ws[dateAddr].z = "dd/mm/yyyy";
    } else {
      setSheetCell(ws, weekdayRow, c, "", "s");
      setSheetCell(ws, headerRow, c, "", "s");
    }
  }

  for (let r = employeeStartRow; r <= employeeEndRow; r += 1) {
    for (const c of fixedCols) {
      setSheetCell(ws, r, c, "", "s");
    }
    for (let c = dateStartCol; c <= dateEndCol; c += 1) {
      setSheetCell(ws, r, c, "", "s");
    }
  }

  const maxEmployees = Math.max(0, employeeEndRow - employeeStartRow + 1);
  const selectedEmployees = employees.slice(0, maxEmployees);
  selectedEmployees.forEach((emp, index) => {
    const row = employeeStartRow + index;
    if (Number.isInteger(headerColumns.sup)) setSheetCell(ws, row, headerColumns.sup, emp.sup || "", "s");
    if (Number.isInteger(headerColumns.esp)) setSheetCell(ws, row, headerColumns.esp, emp.esp || "", "s");
    if (Number.isInteger(headerColumns.name ?? nameHeader.c)) {
      setSheetCell(ws, row, headerColumns.name ?? nameHeader.c, emp.name || "", "s");
    }
    if (Number.isInteger(headerColumns.dni)) setSheetCell(ws, row, headerColumns.dni, emp.dni || "", "s");
    if (Number.isInteger(headerColumns.telefono)) {
      setSheetCell(ws, row, headerColumns.telefono, emp.telefono || "", "s");
    }
    if (Number.isInteger(headerColumns.ascensoDescenso)) {
      setSheetCell(ws, row, headerColumns.ascensoDescenso, emp.ascensoDescenso || "", "s");
    }
    if (Number.isInteger(headerColumns.categoria)) {
      setSheetCell(ws, row, headerColumns.categoria, emp.categoria || "", "s");
    }
    if (Number.isInteger(headerColumns.grupo)) setSheetCell(ws, row, headerColumns.grupo, emp.grupo || "", "s");
    if (Number.isInteger(headerColumns.tipo)) setSheetCell(ws, row, headerColumns.tipo, emp.tipo || "", "s");
    if (Number.isInteger(headerColumns.residencia)) {
      setSheetCell(ws, row, headerColumns.residencia, emp.residencia || "", "s");
    }

    for (let c = dateStartCol; c <= dateEndCol; c += 1) {
      const dateIso = dates[c - dateStartCol];
      if (!dateIso) {
        setSheetCell(ws, row, c, "", "s");
        continue;
      }
      const code = cleanText(store.shifts[emp.id]?.[dateIso]).toUpperCase();
      setSheetCell(ws, row, c, code || "", "s");
    }
  });

  const isCoordinador = (emp) => getRoleKind(emp.categoria) === "coordinador";
  const isInspector = (emp) => getRoleKind(emp.categoria) === "inspector";
  const coordinadorEmployees = selectedEmployees.filter(isCoordinador);
  const inspectorEmployees = selectedEmployees.filter(isInspector);
  const localsBase = coordinadorEmployees.length + inspectorEmployees.length || selectedEmployees.length;
  const countOnsite = (list, date) =>
    list.reduce((count, emp) => {
      const code = cleanText(store.shifts[emp.id]?.[date]).toUpperCase();
      return onsiteCodes.has(code) ? count + 1 : count;
    }, 0);

  const coordinadorByDate = dates.map((date) => countOnsite(coordinadorEmployees, date));
  const inspectorByDate = dates.map((date) => countOnsite(inspectorEmployees, date));
  const totalByDate = dates.map((_, idx) => coordinadorByDate[idx] + inspectorByDate[idx]);
  const percentByDate = totalByDate.map((total) => {
    if (!localsBase) return 0;
    return Number((total / localsBase).toFixed(4));
  });

  const fillMetricRow = (rowIndex, values, isPercent = false) => {
    if (!Number.isInteger(rowIndex)) return;
    for (let c = dateStartCol; c <= dateEndCol; c += 1) {
      const idx = c - dateStartCol;
      const value = idx < values.length ? values[idx] : "";
      if (value === "") {
        setSheetCell(ws, rowIndex, c, "", "s");
      } else {
        setSheetCell(ws, rowIndex, c, value, isPercent ? "n" : "n");
      }
    }
  };

  fillMetricRow(metricRows.direct, coordinadorByDate);
  fillMetricRow(metricRows.indirect, inspectorByDate);
  fillMetricRow(metricRows.total, totalByDate);
  fillMetricRow(metricRows.percentage, percentByDate, true);

  ensureSheetRangeIncludes(ws, Math.max(employeeEndRow, ...(Object.values(metricRows) || [0])), dateEndCol);
  return wb;
}

function buildSimpleRosterWorkbook({ store, employees, dates, onsiteCodes }) {
  const fixedHeaders = ["NOMBRE"];
  const headerRow = [...fixedHeaders, ...dates.map((date) => formatDateShort(date))];
  const weekdayRow = ["", ...dates.map((date) => weekdayEsShort(date))];

  const rows = [headerRow, weekdayRow];
  for (const emp of employees) {
    const fixed = [emp.name || ""];
    const shifts = dates.map((date) => store.shifts[emp.id]?.[date] || "");
    rows.push([...fixed, ...shifts]);
  }

  const isCoordinador = (emp) => normalizeText(emp.categoria).includes("coordinador");
  const isInspector = (emp) => normalizeText(emp.categoria).includes("inspector");
  const coordinadorEmployees = employees.filter(isCoordinador);
  const inspectorEmployees = employees.filter(isInspector);
  const localsBase = coordinadorEmployees.length + inspectorEmployees.length || employees.length;

  const countOnsite = (list, date) =>
    list.reduce((count, emp) => {
      const code = cleanText(store.shifts[emp.id]?.[date]).toUpperCase();
      return onsiteCodes.has(code) ? count + 1 : count;
    }, 0);

  const coordinadorByDate = dates.map((date) => countOnsite(coordinadorEmployees, date));
  const inspectorByDate = dates.map((date) => countOnsite(inspectorEmployees, date));
  const totalByDate = dates.map((_, idx) => coordinadorByDate[idx] + inspectorByDate[idx]);
  const percentByDate = totalByDate.map((total) => {
    if (!localsBase) return "0%";
    const percent = (total / localsBase) * 100;
    return `${percent.toFixed(1)}%`;
  });

  rows.push([]);
  rows.push(["COORDINADORES EN SITIO POR DIA", ...coordinadorByDate]);
  rows.push(["INSPECTORES EN SITIO POR DIA", ...inspectorByDate]);
  rows.push(["TOTAL: COORDINADORES + INSPECTORES EN SITIO POR DIA", ...totalByDate]);
  rows.push(["PORCENTAJE COORD+INSP EN SITIO", ...percentByDate]);

  const sheet = XLSX.utils.aoa_to_sheet(rows);
  const fixedCols = [30];
  const dateCols = dates.map(() => 6);
  sheet["!cols"] = [...fixedCols, ...dateCols].map((wch) => ({ wch }));

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, sheet, "ROSTER");

  const summaryRows = [
    ["RESUMEN EXPORTACION"],
    ["Generado", new Date().toLocaleString("es-AR")],
    ["Empleados", employees.length],
    ["Desde", dates[0] || "-"],
    ["Hasta", dates[dates.length - 1] || "-"],
    ["Dias", dates.length],
  ];
  const summarySheet = XLSX.utils.aoa_to_sheet(summaryRows);
  summarySheet["!cols"] = [{ wch: 22 }, { wch: 36 }];
  XLSX.utils.book_append_sheet(wb, summarySheet, "RESUMEN");

  return wb;
}

function getCodeColorStyle(code) {
  const normalized = cleanText(code).toUpperCase();
  if (normalized === "S") {
    return {
      font: { color: { argb: "FFB42318" }, bold: true },
      fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFFDECEC" } },
    };
  }
  if (normalized === "B") {
    return {
      font: { color: { argb: "FF1D4ED8" }, bold: true },
      fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFEAF1FF" } },
    };
  }
  if (normalized === "D") {
    return {
      font: { color: { argb: "FF15803D" }, bold: true },
      fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFECFDF3" } },
    };
  }
  if (normalized === "1") {
    return {
      font: { color: { argb: "FF111827" }, bold: true },
      fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFF2F4F7" } },
    };
  }
  return {
    font: { color: { argb: "FF667085" } },
    fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } },
  };
}

async function buildRrhhMinimalWorkbookBuffer({ store, employees, dates }) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("ROSTER", {
    views: [{ state: "frozen", xSplit: 1, ySplit: 2 }],
  });

  ws.columns = [{ width: 34 }, ...dates.map(() => ({ width: 6 }))];

  const headerValues = ["APELLIDO Y NOMBRE", ...dates.map((date) => formatDateShort(date))];
  const weekdayValues = ["", ...dates.map((date) => weekdayEsShort(date))];
  ws.addRow(headerValues);
  ws.addRow(weekdayValues);

  const borderThin = {
    top: { style: "thin", color: { argb: "FF6B7280" } },
    left: { style: "thin", color: { argb: "FF6B7280" } },
    bottom: { style: "thin", color: { argb: "FF6B7280" } },
    right: { style: "thin", color: { argb: "FF6B7280" } },
  };

  const headerFill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFD1D5DB" },
  };

  const firstHeaderCell = ws.getCell(1, 1);
  firstHeaderCell.font = { name: "Calibri", size: 12, bold: true, color: { argb: "FF111827" } };
  firstHeaderCell.alignment = { horizontal: "left", vertical: "middle" };
  firstHeaderCell.fill = headerFill;
  firstHeaderCell.border = borderThin;

  for (let col = 2; col <= dates.length + 1; col += 1) {
    const headerCell = ws.getCell(1, col);
    headerCell.font = { name: "Calibri", size: 9, bold: true, color: { argb: "FF111827" } };
    headerCell.alignment = { horizontal: "center", vertical: "middle", textRotation: 90 };
    headerCell.fill = headerFill;
    headerCell.border = borderThin;

    const weekdayCell = ws.getCell(2, col);
    weekdayCell.font = { name: "Calibri", size: 8, bold: true, color: { argb: "FF111827" } };
    weekdayCell.alignment = { horizontal: "center", vertical: "middle" };
    weekdayCell.fill = headerFill;
    weekdayCell.border = borderThin;
  }

  ws.getCell(2, 1).fill = headerFill;
  ws.getCell(2, 1).border = borderThin;

  employees.forEach((emp, idx) => {
    const rowIndex = idx + 3;
    const row = ws.getRow(rowIndex);
    row.getCell(1).value = emp.name || "";
    row.getCell(1).font = { name: "Calibri", size: 12, color: { argb: "FF111827" } };
    row.getCell(1).alignment = { horizontal: "left", vertical: "middle" };
    row.getCell(1).border = borderThin;

    dates.forEach((dateIso, dateIdx) => {
      const cell = row.getCell(dateIdx + 2);
      const code = cleanText(store.shifts[emp.id]?.[dateIso]).toUpperCase();
      const style = getCodeColorStyle(code);
      cell.value = code || "";
      cell.font = { name: "Calibri", size: 11, ...(style.font || {}) };
      cell.fill = style.fill;
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = borderThin;
    });
  });

  return wb.xlsx.writeBuffer();
}

function sortIsoDates(dates) {
  return [...dates].sort((a, b) => a.localeCompare(b));
}

function dateRange(from, to, allDates) {
  const sorted = sortIsoDates(allDates);
  if (!sorted.length) {
    return [];
  }

  const start = from && /^\d{4}-\d{2}-\d{2}$/.test(from) ? from : sorted[0];
  const end =
    to && /^\d{4}-\d{2}-\d{2}$/.test(to) ? to : sorted[sorted.length - 1];

  const safeStart = start <= end ? start : end;
  const safeEnd = start <= end ? end : start;

  return sorted.filter((date) => date >= safeStart && date <= safeEnd);
}

function getInitialStore() {
  return {
    meta: {
      allowedCodes: DEFAULT_CODES,
      discoveredCodes: [],
      onsiteCodes: ["1"],
      hiddenEmployees: [],
      dates: [],
      dateRange: { from: null, to: null },
      lastImportAt: null,
    },
    employees: [],
    shifts: {},
    audit: [],
  };
}

function normalizeStore(parsed) {
  const base = getInitialStore();
  return {
    ...base,
    ...(parsed && typeof parsed === "object" ? parsed : {}),
    meta: {
      ...base.meta,
      ...((parsed && parsed.meta) || {}),
    },
    employees: Array.isArray(parsed?.employees) ? parsed.employees : [],
    shifts: parsed?.shifts && typeof parsed.shifts === "object" ? parsed.shifts : {},
    audit: Array.isArray(parsed?.audit) ? parsed.audit : [],
  };
}

function readStoreSync() {
  ensureDirs();
  if (!fs.existsSync(STORE_PATH)) {
    const store = getInitialStore();
    fs.writeFileSync(STORE_PATH, JSON.stringify(store, null, 2), "utf8");
    return store;
  }

  try {
    const data = fs.readFileSync(STORE_PATH, "utf8");
    const parsed = JSON.parse(data);
    return normalizeStore(parsed);
  } catch (_error) {
    const fallback = getInitialStore();
    fs.writeFileSync(STORE_PATH, JSON.stringify(fallback, null, 2), "utf8");
    return fallback;
  }
}

function writeStoreSync(store) {
  fs.writeFileSync(STORE_PATH, JSON.stringify(store, null, 2), "utf8");
}

let _supabase;
function cleanEnvValue(value) {
  const trimmed = cleanText(value);
  // Remove accidental wrapping quotes from copy/paste
  if (
    (trimmed.startsWith("\"") && trimmed.endsWith("\"")) ||
    (trimmed.startsWith("'") && trimmed.endsWith("'"))
  ) {
    return trimmed.slice(1, -1).trim();
  }
  return trimmed;
}

function getSupabase() {
  const url = cleanEnvValue(process.env.SUPABASE_URL);
  const key =
    cleanEnvValue(process.env.SUPABASE_SERVICE_ROLE_KEY) ||
    cleanEnvValue(process.env.SUPABASE_ANON_KEY) ||
    cleanEnvValue(process.env.SUPABASE_KEY);
  if (!url || !key) {
    throw new Error(
      "Faltan variables de entorno de Supabase (SUPABASE_URL y SUPABASE_SERVICE_ROLE_KEY/ANON_KEY)."
    );
  }
  if (!_supabase) {
    _supabase = createClient(url, key, {
      auth: { persistSession: false },
    });
  }
  return _supabase;
}

async function readStoreFromSupabase() {
  const supabase = getSupabase();
  const { data, error } = await supabase
    .from(SUPABASE_TABLE)
    .select("data")
    .eq("id", 1)
    .maybeSingle();
  if (error) throw error;
  if (!data || !data.data) {
    const store = getInitialStore();
    await writeStoreToSupabase(store);
    return store;
  }
  return normalizeStore(data.data);
}

async function writeStoreToSupabase(store) {
  const supabase = getSupabase();
  const payload = {
    id: 1,
    data: store,
    updated_at: new Date().toISOString(),
  };
  const { error } = await supabase.from(SUPABASE_TABLE).upsert(payload, { onConflict: "id" });
  if (error) throw error;
}

async function readStore() {
  if (STORE_BACKEND === "supabase") return await readStoreFromSupabase();
  return readStoreSync();
}

async function writeStore(store) {
  if (STORE_BACKEND === "supabase") return await writeStoreToSupabase(store);
  return writeStoreSync(store);
}

function findHeaderRow(rows) {
  for (let r = 0; r < Math.min(rows.length, 40); r += 1) {
    const row = rows[r] || [];
    for (let c = 0; c < row.length; c += 1) {
      if (normalizeText(row[c]) === "apellido y nombre") {
        return r;
      }
    }
  }
  return -1;
}

function findColumnIndexes(headerRow) {
  const indexes = {};
  const aliases = {
    sup: ["sup"],
    esp: ["esp"],
    name: ["apellido y nombre"],
    dni: ["dni"],
    telefono: ["telefono"],
    ascenso: ["punto ascenso/descenso"],
    categoria: ["categoria real"],
    grupo: ["grupo"],
    tipo: ["tipo"],
    residencia: ["residencia"],
  };

  for (let c = 0; c < headerRow.length; c += 1) {
    const normalized = normalizeText(headerRow[c]);
    for (const [key, values] of Object.entries(aliases)) {
      if (indexes[key] !== undefined) {
        continue;
      }
      if (values.includes(normalized)) {
        indexes[key] = c;
      }
    }
  }

  return indexes;
}

function employeeIdentity(row, cols) {
  const dni = cleanText(row[cols.dni]);
  if (dni) {
    return `dni:${dni}`;
  }

  const name = normalizeText(row[cols.name]);
  const residencia = normalizeText(row[cols.residencia]);
  const categoria = normalizeText(row[cols.categoria]);
  return `name:${name}|res:${residencia}|cat:${categoria}`;
}

function identityToId(identity, usedIds) {
  const base = identity
    .replace(/[^a-zA-Z0-9|:_-]+/g, "-")
    .replace(/[|:]+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "")
    .slice(0, 80)
    .toLowerCase();

  let candidate = base || `emp-${Date.now()}`;
  let index = 2;
  while (usedIds.has(candidate)) {
    candidate = `${base}-${index}`;
    index += 1;
  }
  usedIds.add(candidate);
  return candidate;
}

function parseExcel(filePath) {
  const workbook = XLSX.readFile(filePath, { cellDates: true });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: true,
    defval: "",
  });

  const headerRowIndex = findHeaderRow(rows);
  if (headerRowIndex < 0) {
    throw new Error(
      "No se encontro la fila de encabezado (APELLIDO Y NOMBRE) en el Excel."
    );
  }

  const headerRow = rows[headerRowIndex] || [];
  const cols = findColumnIndexes(headerRow);
  if (cols.name === undefined) {
    throw new Error("No se encontro la columna APELLIDO Y NOMBRE.");
  }

  const metadataEnd = Math.max(
    cols.residencia ?? 10,
    cols.tipo ?? 9,
    cols.grupo ?? 8,
    cols.categoria ?? 7
  );

  const rawDateColumns = [];
  for (let c = metadataEnd + 1; c < headerRow.length; c += 1) {
    const date = parseDateValue(headerRow[c]);
    if (date) {
      rawDateColumns.push({ col: c, date });
    }
  }
  const normalizedCalendar = normalizeDateColumns(rawDateColumns);
  const dateColumns = normalizedCalendar.dateColumns;

  const employees = [];
  const shifts = {};
  const discoveredCodes = new Set();
  const usedIds = new Set();
  const identityCount = new Map();

  for (let r = headerRowIndex + 1; r < rows.length; r += 1) {
    const row = rows[r] || [];
    const name = normalizeNameValue(row[cols.name]);
    if (!name) {
      continue;
    }
    if (normalizeText(name) === "apellido y nombre") {
      continue;
    }

    const identity = employeeIdentity(row, cols);
    const count = (identityCount.get(identity) || 0) + 1;
    identityCount.set(identity, count);
    const uniqueIdentity = count > 1 ? `${identity}|row:${count}` : identity;
    const id = identityToId(uniqueIdentity, usedIds);

    const employee = {
      id,
      sup: cleanText(row[cols.sup]),
      esp: cleanText(row[cols.esp]),
      name,
      dni: cleanText(row[cols.dni]),
      telefono: cleanText(row[cols.telefono]),
      ascensoDescenso: cleanText(row[cols.ascenso]),
      categoria: normalizeCategoriaValue(row[cols.categoria]),
      grupo: cleanText(row[cols.grupo]),
      tipo: cleanText(row[cols.tipo]),
      residencia: cleanText(row[cols.residencia]),
    };
    applyBusinessOverrides(employee);

    const employeeShifts = {};
    for (const { col, date } of dateColumns) {
      const raw = row[col];
      const code = cleanText(raw).toUpperCase();
      if (!code) {
        continue;
      }
      if (isDateLikeString(code)) {
        continue;
      }
      employeeShifts[date] = code;
      discoveredCodes.add(code);
    }

    const hasMetadata =
      employee.categoria ||
      employee.grupo ||
      employee.tipo ||
      employee.residencia ||
      employee.dni;
    const hasShifts = Object.keys(employeeShifts).length > 0;

    if (!hasMetadata && !hasShifts) {
      continue;
    }

    employees.push(employee);
    shifts[id] = employeeShifts;
  }

  const sortedDates = sortIsoDates(dateColumns.map((d) => d.date));
  const sortedEmployees = sortByName(employees);

  return {
    sheetName,
    employees: sortedEmployees,
    shifts,
    discoveredCodes: sortIsoDates([...discoveredCodes]),
    dates: sortedDates,
    dateRange: {
      from: sortedDates[0] || null,
      to: sortedDates[sortedDates.length - 1] || null,
    },
    calendarFix: {
      wasFixed: normalizedCalendar.wasFixed,
      gaps: normalizedCalendar.gaps,
    },
  };
}

function getFilterOptions(employees) {
  const categorias = new Set();
  const residencias = new Set();
  const grupos = new Set();
  const tipos = new Set();
  const sites = new Set();

  for (const emp of employees) {
    if (emp.categoria) categorias.add(emp.categoria);
    if (emp.residencia) residencias.add(emp.residencia);
    if (emp.grupo) grupos.add(emp.grupo);
    if (emp.tipo) tipos.add(emp.tipo);
    if (emp.ascensoDescenso) sites.add(emp.ascensoDescenso);
  }

  return {
    categorias: [...categorias].sort(),
    residencias: [...residencias].sort(),
    grupos: [...grupos].sort(),
    tipos: [...tipos].sort(),
    sites: [...sites].sort((a, b) => a.localeCompare(b, "es")),
  };
}

function filterEmployees(employees, query) {
  const employeeId = cleanText(query.employeeId);
  const employeeIds = cleanText(query.employeeIds)
    .split(",")
    .map((id) => cleanText(id))
    .filter(Boolean);
  if (employeeId) {
    employeeIds.push(employeeId);
  }
  const selectedEmployeeIds = new Set(employeeIds);
  const search = normalizeText(query.search);
  const categoria = normalizeText(query.categoria);
  const residencia = normalizeText(query.residencia);
  const grupo = normalizeText(query.grupo);
  const tipo = normalizeText(query.tipo);

  const filtered = employees.filter((emp) => {
    if (selectedEmployeeIds.size && !selectedEmployeeIds.has(emp.id)) return false;
    if (categoria && normalizeText(emp.categoria) !== categoria) return false;
    if (residencia && normalizeText(emp.residencia) !== residencia) return false;
    if (grupo && normalizeText(emp.grupo) !== grupo) return false;
    if (tipo && normalizeText(emp.tipo) !== tipo) return false;

    if (!search) return true;

    const searchable = normalizeText(
      `${emp.name} ${emp.dni} ${emp.categoria} ${emp.residencia} ${emp.grupo} ${emp.tipo} ${emp.ascensoDescenso}`
    );
    return searchable.includes(search);
  });

  return sortByName(filtered);
}

function buildConflicts(store, dates, minCoverage) {
  const conflicts = [];
  const allowed = new Set([
    ...(store.meta.allowedCodes || []),
    ...(store.meta.discoveredCodes || []),
  ]);
  const onsiteCodes = new Set(
    (store.meta.onsiteCodes && store.meta.onsiteCodes.length ? store.meta.onsiteCodes : ["1"]).map(
      (c) => cleanText(c).toUpperCase()
    )
  );

  for (const emp of store.employees) {
    const empShifts = store.shifts[emp.id] || {};
    let assigned = 0;
    let consecutiveD = 0;
    let consecutiveB = 0;

    for (const date of dates) {
      const code = cleanText(empShifts[date]).toUpperCase();
      if (code) {
        assigned += 1;
      }

      if (code && !allowed.has(code)) {
        conflicts.push({
          severity: "high",
          type: "codigo-invalido",
          employeeId: emp.id,
          employeeName: emp.name,
          date,
          message: `${emp.name}: codigo "${code}" no permitido (${date}).`,
        });
      }

      if (code === "D") {
        consecutiveD += 1;
        if (consecutiveD === 15) {
          conflicts.push({
            severity: "medium",
            type: "exceso-d-consecutivos",
            employeeId: emp.id,
            employeeName: emp.name,
            date,
            message: `${emp.name}: supera 14 dias consecutivos con codigo D.`,
          });
        }
      } else {
        consecutiveD = 0;
      }

      if (code === "B") {
        consecutiveB += 1;
        if (consecutiveB === 8) {
          conflicts.push({
            severity: "medium",
            type: "exceso-b-consecutivos",
            employeeId: emp.id,
            employeeName: emp.name,
            date,
            message: `${emp.name}: supera 7 dias consecutivos con codigo B.`,
          });
        }
      } else {
        consecutiveB = 0;
      }
    }

    if (assigned === 0 && dates.length > 0) {
      conflicts.push({
        severity: "low",
        type: "sin-turnos",
        employeeId: emp.id,
        employeeName: emp.name,
        message: `${emp.name}: sin turnos cargados en el rango seleccionado.`,
      });
    }
  }

  for (const date of dates) {
    let active = 0;
    for (const emp of store.employees) {
      const code = cleanText(store.shifts[emp.id]?.[date]).toUpperCase();
      if (onsiteCodes.has(code)) active += 1;
    }

    if (active < minCoverage) {
      conflicts.push({
        severity: "medium",
        type: "cobertura-baja",
        date,
        message: `${date}: cobertura en sitio=${active} (codigos: ${[...onsiteCodes].join(
          ", "
        )}), minimo requerido=${minCoverage}.`,
      });
    }
  }

  const severityRank = { high: 0, medium: 1, low: 2 };
  conflicts.sort(
    (a, b) =>
      severityRank[a.severity] - severityRank[b.severity] ||
      (a.date || "").localeCompare(b.date || "")
  );

  return conflicts;
}

function buildSiteGroups(store, dates, requestedSite, requestedCodes) {
  const siteFilter = normalizeText(requestedSite);
  const defaultCodes = (store.meta.onsiteCodes || ["1"]).map((c) => c.toUpperCase());
  const codes = (requestedCodes.length ? requestedCodes : defaultCodes)
    .map((c) => cleanText(c).toUpperCase())
    .filter(Boolean);
  const onsiteCodes = new Set(codes);

  const employeesById = new Map(store.employees.map((emp) => [emp.id, emp]));
  const perSite = new Map();

  for (const emp of store.employees) {
    const site = cleanText(emp.ascensoDescenso) || "SIN SITIO DEFINIDO";
    if (siteFilter && !normalizeText(site).includes(siteFilter)) {
      continue;
    }

    if (!perSite.has(site)) {
      perSite.set(site, {
        site,
        participantsByDate: new Map(),
      });
    }

    const shifts = store.shifts[emp.id] || {};
    for (const date of dates) {
      const code = cleanText(shifts[date]).toUpperCase();
      if (!onsiteCodes.has(code)) {
        continue;
      }
      const participants = perSite.get(site).participantsByDate.get(date) || new Set();
      participants.add(emp.id);
      perSite.get(site).participantsByDate.set(date, participants);
    }
  }

  const siteSummaries = [];

  for (const siteData of perSite.values()) {
    const memberIds = new Set();
    for (const peopleSet of siteData.participantsByDate.values()) {
      for (const empId of peopleSet) {
        memberIds.add(empId);
      }
    }

    if (!memberIds.size) {
      continue;
    }

    const siteRoleMembers = { coordinadores: [], inspectores: [], otros: [] };
    for (const memberId of memberIds) {
      const employee = employeesById.get(memberId);
      if (!employee) continue;
      const role = getRoleKind(employee.categoria);
      if (role === "coordinador") siteRoleMembers.coordinadores.push(employee.name);
      else if (role === "inspector") siteRoleMembers.inspectores.push(employee.name);
      else siteRoleMembers.otros.push(employee.name);
    }
    siteRoleMembers.coordinadores.sort((a, b) => a.localeCompare(b, "es"));
    siteRoleMembers.inspectores.sort((a, b) => a.localeCompare(b, "es"));
    siteRoleMembers.otros.sort((a, b) => a.localeCompare(b, "es"));

    let peakTotal = 0;
    let peakDate = null;
    let sumTotal = 0;
    for (const date of dates) {
      const peopleSet = siteData.participantsByDate.get(date);
      const total = peopleSet ? peopleSet.size : 0;
      sumTotal += total;
      if (total > peakTotal) {
        peakTotal = total;
        peakDate = date;
      }
    }
    const avgTotal = dates.length ? sumTotal / dates.length : 0;

    const graph = new Map();
    for (const memberId of memberIds) {
      graph.set(memberId, new Set());
    }

    for (const peopleSet of siteData.participantsByDate.values()) {
      const ids = [...peopleSet];
      for (let i = 0; i < ids.length; i += 1) {
        for (let j = i + 1; j < ids.length; j += 1) {
          graph.get(ids[i]).add(ids[j]);
          graph.get(ids[j]).add(ids[i]);
        }
      }
    }

    const visited = new Set();
    const groups = [];

    for (const startId of graph.keys()) {
      if (visited.has(startId)) continue;
      const queue = [startId];
      const component = [];
      visited.add(startId);

      while (queue.length) {
        const current = queue.shift();
        component.push(current);
        for (const neighbor of graph.get(current)) {
          if (!visited.has(neighbor)) {
            visited.add(neighbor);
            queue.push(neighbor);
          }
        }
      }

      const activeDates = [];
      let coincideDays = 0;
      for (const date of dates) {
        const setForDay = siteData.participantsByDate.get(date);
        if (!setForDay) continue;
        let presentCount = 0;
        for (const memberId of component) {
          if (setForDay.has(memberId)) {
            presentCount += 1;
          }
        }
        if (presentCount > 0) activeDates.push(date);
        if (presentCount >= 2) coincideDays += 1;
      }

      const memberEmployees = component
        .map((id) => employeesById.get(id))
        .filter(Boolean)
        .sort((a, b) => cleanText(a.name).localeCompare(cleanText(b.name), "es"));
      const members = memberEmployees.map((emp) => emp.name || emp.id);
      const roleMembers = { coordinadores: [], inspectores: [], otros: [] };
      for (const employee of memberEmployees) {
        const role = getRoleKind(employee.categoria);
        if (role === "coordinador") roleMembers.coordinadores.push(employee.name);
        else if (role === "inspector") roleMembers.inspectores.push(employee.name);
        else roleMembers.otros.push(employee.name);
      }
      groups.push({
        members,
        roleMembers,
        size: members.length,
        from: activeDates[0] || null,
        to: activeDates[activeDates.length - 1] || null,
        activeDays: activeDates.length,
        coincideDays,
      });
    }

    groups.sort((a, b) => b.size - a.size || b.coincideDays - a.coincideDays);
    siteSummaries.push({
      site: siteData.site,
      groups,
      totalPeople: memberIds.size,
      roleTotals: {
        coordinadores: siteRoleMembers.coordinadores.length,
        inspectores: siteRoleMembers.inspectores.length,
        otros: siteRoleMembers.otros.length,
      },
      roleMembers: siteRoleMembers,
      peakTotal,
      peakDate,
      avgTotal: Number(avgTotal.toFixed(2)),
    });
  }

  siteSummaries.sort((a, b) => a.site.localeCompare(b.site, "es"));
  return { siteSummaries, codes: [...onsiteCodes] };
}

function computeSegmentsForCodes(dates, shifts, codeSet) {
  const segments = [];
  let current = null;
  for (const date of dates) {
    const code = cleanText(shifts?.[date]).toUpperCase();
    const active = codeSet.has(code);
    if (active) {
      if (!current) {
        current = { from: date, to: date, days: 1 };
      } else {
        current.to = date;
        current.days += 1;
      }
    } else if (current) {
      segments.push(current);
      current = null;
    }
  }
  if (current) segments.push(current);
  return segments;
}

function buildPeriodsByRole({ store, dates, requestedSite, onsiteCodes, restCodes }) {
  const siteFilter = normalizeText(requestedSite);
  const byRole = {
    coordinadores: [],
    inspectores: [],
    otros: [],
  };

  for (const emp of store.employees || []) {
    const site = cleanText(emp.ascensoDescenso) || "SIN SITIO DEFINIDO";
    if (siteFilter && !normalizeText(site).includes(siteFilter)) {
      continue;
    }
    const shifts = store.shifts?.[emp.id] || {};
    const onsiteSegments = computeSegmentsForCodes(dates, shifts, onsiteCodes);
    const restSegments = computeSegmentsForCodes(dates, shifts, restCodes);
    if (!onsiteSegments.length && !restSegments.length) continue;

    const role = getRoleKind(emp.categoria);
    const bucket =
      role === "coordinador"
        ? byRole.coordinadores
        : role === "inspector"
          ? byRole.inspectores
          : byRole.otros;

    bucket.push({
      name: emp.name,
      site,
      onsite: onsiteSegments,
      descanso: restSegments,
    });
  }

  for (const key of Object.keys(byRole)) {
    byRole[key].sort((a, b) => cleanText(a.name).localeCompare(cleanText(b.name), "es"));
  }

  return byRole;
}

function buildAlertMessage({ from, to, codes, siteSummaries, requestedSite, periods }) {
  // Backward compatible wrapper: keep the function name used by /api/alerts/preview,
  // but generate a cleaner message (obra vs descanso) that is easier to read.
  const safePeriods = periods || { coordinadores: [], inspectores: [], otros: [] };

  const fmtSeg = (seg) => `${formatDateEs(seg.from)} al ${formatDateEs(seg.to)}`;
  const fmtList = (segments) => (segments?.length ? segments.map(fmtSeg).join("; ") : "-");
  const pick = (list, key) => (list || []).filter((p) => (p[key] || []).length);

  const lines = [];
  lines.push("COMUNICACION INTERNA - PERSONAL EN OBRA");
  lines.push(`Periodo: ${formatDateEs(from)} al ${formatDateEs(to)}`);
  lines.push("Leyenda: 1=DIA EN OBRA | S=SUBIDA | B=BAJADA | D=DESCANSO");
  if (cleanText(requestedSite)) {
    lines.push(`Sitio: ${cleanText(requestedSite)}`);
  }
  lines.push("");

  const obraCoord = pick(safePeriods.coordinadores, "onsite");
  const obraInsp = pick(safePeriods.inspectores, "onsite");
  const obraOtros = pick(safePeriods.otros, "onsite");

  const renderBlock = (title, coord, insp, otros, key) => {
    lines.push(title);
    lines.push(`Coordinadores (${coord.length}):`);
    if (!coord.length) lines.push("  - -");
    for (const person of coord) {
      lines.push(`  - ${person.name}: ${fmtList(person[key])}`);
    }
    lines.push(`Inspectores (${insp.length}):`);
    if (!insp.length) lines.push("  - -");
    for (const person of insp) {
      lines.push(`  - ${person.name}: ${fmtList(person[key])}`);
    }
    if (otros.length) {
      lines.push(`Otros (${otros.length}):`);
      for (const person of otros) {
        lines.push(`  - ${person.name}: ${fmtList(person[key])}`);
      }
    }
    lines.push("");
  };

  renderBlock("EN OBRA", obraCoord, obraInsp, obraOtros, "onsite");

  const hasAny = obraCoord.length || obraInsp.length || obraOtros.length;
  if (!hasAny) {
    return [
      "COMUNICACION INTERNA - PERSONAL EN OBRA",
      `Periodo: ${formatDateEs(from)} al ${formatDateEs(to)}`,
      cleanText(requestedSite) ? `Sitio: ${cleanText(requestedSite)}` : null,
      "Sin personal en obra para el periodo seleccionado.",
    ]
      .filter(Boolean)
      .join("\n");
  }

  return lines.join("\n").trim();
}

app.use(express.json({ limit: "5mb" }));
app.use(express.urlencoded({ extended: true }));
app.use(
  express.static(path.join(__dirname, "public"), {
    etag: false,
    lastModified: false,
    setHeaders(res) {
      // Evita que el browser se quede con HTML/CSS/JS viejos durante iteraciones.
      res.setHeader("Cache-Control", "no-store");
    },
  })
);

app.get("/api/summary", async (_req, res) => {
  const store = await readStore();
  const visibleEmployees = getVisibleEmployees(store);
  const allDates = store.meta.dates || [];
  let assignments = 0;

  for (const empId of Object.keys(store.shifts)) {
    assignments += Object.keys(store.shifts[empId] || {}).length;
  }

  const byCategoria = {};
  const byResidencia = {};
  for (const emp of visibleEmployees) {
    if (emp.categoria) {
      byCategoria[emp.categoria] = (byCategoria[emp.categoria] || 0) + 1;
    }
    if (emp.residencia) {
      byResidencia[emp.residencia] = (byResidencia[emp.residencia] || 0) + 1;
    }
  }

  const rangeDates = dateRange(
    store.meta.dateRange.from,
    store.meta.dateRange.to,
    allDates
  );
  const conflicts = buildConflicts(store, rangeDates, 2);

  res.json({
    employees: visibleEmployees.length,
    assignments,
    dates: allDates.length,
    dateRange: store.meta.dateRange,
    lastImportAt: store.meta.lastImportAt,
    onsiteCodes: store.meta.onsiteCodes && store.meta.onsiteCodes.length ? store.meta.onsiteCodes : ["1"],
    byCategoria,
    byResidencia,
    conflicts: conflicts.length,
  });
});

app.get("/api/health", (_req, res) => {
  const url = cleanEnvValue(process.env.SUPABASE_URL);
  const key = cleanEnvValue(process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.SUPABASE_ANON_KEY || process.env.SUPABASE_KEY);
  const keyType = key.startsWith("sb_") ? "sb_*" : key.startsWith("eyJ") ? "jwt(eyJ...)" : key ? "other" : "missing";
  res.json({
    ok: true,
    storeBackend: STORE_BACKEND,
    serverless: IS_SERVERLESS,
    supabaseUrlSet: Boolean(url),
    supabaseUrlHost: url ? url.replace(/^https?:\/\//, "").split("/")[0] : null,
    supabaseKeySet: Boolean(key),
    supabaseKeyType: keyType,
    supabaseKeyLen: key ? key.length : 0,
  });
});

app.get("/api/codes", async (_req, res) => {
  const store = await readStore();
  const codes = Array.from(
    new Set([...(store.meta.allowedCodes || []), ...(store.meta.discoveredCodes || [])])
  ).sort();
  res.json({ codes });
});

app.get("/api/employees", async (_req, res) => {
  const store = await readStore();
  const employees = sortByName(getVisibleEmployees(store)).map((emp) => ({
    id: emp.id,
    name: emp.name,
    categoria: emp.categoria,
    residencia: emp.residencia,
    grupo: emp.grupo,
    tipo: emp.tipo,
    dni: emp.dni,
    telefono: emp.telefono,
    ascensoDescenso: emp.ascensoDescenso,
    sup: emp.sup,
    esp: emp.esp,
  }));
  res.json({ employees });
});

app.get("/api/months", async (_req, res) => {
  const store = await readStore();
  const dates = sortIsoDates(store.meta.dates || []);
  const monthMap = new Map();

  for (const date of dates) {
    const month = date.slice(0, 7);
    if (!monthMap.has(month)) {
      monthMap.set(month, { month, from: date, to: date, days: 1 });
    } else {
      const current = monthMap.get(month);
      current.to = date;
      current.days += 1;
    }
  }

  const months = [...monthMap.values()];
  res.json({ months });
});

app.get("/api/alerts/preview", async (req, res) => {
  const store = await readStore();
  const visibleEmployees = getVisibleEmployees(store);
  const dates = dateRange(req.query.from, req.query.to, store.meta.dates || []);
  const from = dates[0] || req.query.from || store.meta.dateRange.from;
  const to = dates[dates.length - 1] || req.query.to || store.meta.dateRange.to;

  const site = DIFFUSION_SITE;
  const onsiteCodes = new Set(["1"]);

  const onsiteCoordinators = [];
  const onsiteInspectors = [];
  const onsiteOthers = [];

  for (const emp of visibleEmployees) {
    const shifts = store.shifts?.[emp.id] || {};
    let hasOnsite = false;
    for (const date of dates) {
      const code = cleanText(shifts[date]).toUpperCase();
      if (onsiteCodes.has(code)) hasOnsite = true;
      if (hasOnsite) break;
    }

    if (!hasOnsite) continue;

    const role = getRoleKind(emp.categoria);
    if (role === "coordinador") onsiteCoordinators.push(emp.name);
    else if (role === "inspector") onsiteInspectors.push(emp.name);
    else onsiteOthers.push(emp.name);
  }

  const sortNames = (arr) =>
    arr.sort((a, b) => cleanText(a).localeCompare(cleanText(b), "es"));
  sortNames(onsiteCoordinators);
  sortNames(onsiteInspectors);
  sortNames(onsiteOthers);

  const messageLines = [
    "COMUNICACION INTERNA - PERSONAL EN OBRA",
    `Periodo: ${formatDateEs(from)} al ${formatDateEs(to)}`,
    `Sitio: ${site}`,
    `Coordinadores en obra: ${
      onsiteCoordinators.length ? onsiteCoordinators.join(", ") : "-"
    }`,
    `Inspectores en obra: ${onsiteInspectors.length ? onsiteInspectors.join(", ") : "-"}`,
  ];
  if (onsiteOthers.length) {
    messageLines.push(`Otros en obra: ${onsiteOthers.join(", ")}`);
  }
  const message = messageLines.join("\n");
  const subject = `Comunicacion interna personal en obra ${formatDateEs(from)}-${formatDateEs(to)}`;
  const mailtoUrl = `mailto:?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(
    message
  )}`;

  res.json({
    from,
    to,
    site,
    message,
    mailtoUrl,
  });
});

app.get("/api/roster", async (req, res) => {
  const store = await readStore();
  const employees = filterEmployees(getVisibleEmployees(store), req.query);
  const allDates = store.meta.dates || [];
  const dates = dateRange(req.query.from, req.query.to, allDates);

  const rows = employees.map((emp) => ({
    ...emp,
    shifts: dates.reduce((acc, date) => {
      const code = store.shifts[emp.id]?.[date];
      if (code) acc[date] = code;
      return acc;
    }, {}),
  }));

  const codes = Array.from(
    new Set([...(store.meta.allowedCodes || []), ...(store.meta.discoveredCodes || [])])
  ).sort();

  res.json({
    dates,
    codes,
    employees: rows,
    filterOptions: getFilterOptions(getVisibleEmployees(store)),
    totalEmployees: getVisibleEmployees(store).length,
    filteredEmployees: rows.length,
  });
});

app.get("/api/conflicts", async (req, res) => {
  const store = await readStore();
  const minCoverage = Math.max(0, Number(req.query.minCoverage || 2));
  const dates = dateRange(req.query.from, req.query.to, store.meta.dates || []);
  const conflicts = buildConflicts(store, dates, minCoverage);
  res.json({ conflicts, count: conflicts.length });
});

app.get("/api/export/roster.xlsx", async (req, res) => {
  const store = await readStore();
  const employees = filterEmployees(getVisibleEmployees(store), req.query);
  const dates = dateRange(req.query.from, req.query.to, store.meta.dates || []);
  try {
    const buffer = await buildRrhhMinimalWorkbookBuffer({ store, employees, dates });
    const fileName = `roster-${new Date().toISOString().slice(0, 10)}.xlsx`;
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader("Content-Disposition", `attachment; filename="${fileName}"`);
    res.send(Buffer.from(buffer));
  } catch (error) {
    res.status(500).json({ error: "No se pudo generar el Excel de exportacion." });
  }
});

app.get("/api/export/roster.csv", async (req, res) => {
  const store = await readStore();
  const employees = filterEmployees(getVisibleEmployees(store), req.query);
  const dates = dateRange(req.query.from, req.query.to, store.meta.dates || []);

  const headers = [
    "APELLIDO Y NOMBRE",
    "DNI",
    "CATEGORIA REAL",
    "GRUPO",
    "TIPO",
    "RESIDENCIA",
    ...dates,
  ];

  const rows = [headers];
  for (const emp of employees) {
    const fixed = [
      emp.name,
      emp.dni,
      emp.categoria,
      emp.grupo,
      emp.tipo,
      emp.residencia,
    ];
    const shiftValues = dates.map((date) => store.shifts[emp.id]?.[date] || "");
    rows.push([...fixed, ...shiftValues]);
  }

  const delimiter = ";";
  const csv = toCsv(rows, delimiter);
  const fileName = `roster-${new Date().toISOString().slice(0, 10)}.csv`;

  res.setHeader("Content-Type", "text/csv; charset=utf-8");
  res.setHeader("Content-Disposition", `attachment; filename="${fileName}"`);
  res.send(`\ufeffsep=${delimiter}\n${csv}`);
});

app.post("/api/shifts", async (req, res) => {
  const store = await readStore();
  const employeeId = cleanText(req.body.employeeId);
  const date = cleanText(req.body.date);
  const code = cleanText(req.body.code).toUpperCase();
  const user = cleanText(req.body.user) || "coordinador";

  if (!employeeId || !date) {
    return res.status(400).json({ error: "employeeId y date son obligatorios." });
  }
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
    return res.status(400).json({ error: "date debe estar en formato YYYY-MM-DD." });
  }

  const employee = store.employees.find((emp) => emp.id === employeeId);
  if (!employee) {
    return res.status(404).json({ error: "Empleado no encontrado." });
  }

  if (!store.shifts[employeeId]) {
    store.shifts[employeeId] = {};
  }

  const previous = store.shifts[employeeId][date] || "";
  if (code) {
    store.shifts[employeeId][date] = code;
    if (!store.meta.discoveredCodes.includes(code)) {
      store.meta.discoveredCodes.push(code);
      store.meta.discoveredCodes.sort();
    }
  } else {
    delete store.shifts[employeeId][date];
  }

  if (!store.meta.dates.includes(date)) {
    store.meta.dates.push(date);
    store.meta.dates.sort();
    store.meta.dateRange = {
      from: store.meta.dates[0] || null,
      to: store.meta.dates[store.meta.dates.length - 1] || null,
    };
  }

  store.audit.push({
    at: new Date().toISOString(),
    action: "update-shift",
    employeeId,
    employeeName: employee.name,
    date,
    from: previous,
    to: code,
    user,
  });
  store.audit = store.audit.slice(-5000);

  await writeStore(store);
  return res.json({ ok: true });
});

app.post("/api/config/onsite-codes", async (req, res) => {
  const store = await readStore();
  const raw = cleanText(req.body.codes);
  const codes = String(raw || "")
    .split(/[,\s]+/)
    .map((c) => cleanText(c).toUpperCase())
    .filter(Boolean);

  store.meta.onsiteCodes = codes.length ? codes : ["1"];

  for (const code of store.meta.onsiteCodes) {
    if (!store.meta.discoveredCodes.includes(code)) {
      store.meta.discoveredCodes.push(code);
    }
  }
  store.meta.discoveredCodes.sort();

  store.audit.push({
    at: new Date().toISOString(),
    action: "update-onsite-codes",
    codes: store.meta.onsiteCodes,
  });
  store.audit = store.audit.slice(-5000);

  await writeStore(store);
  return res.json({ ok: true, codes: store.meta.onsiteCodes });
});

app.post("/api/import-excel", upload.single("file"), async (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: "Debes subir un archivo Excel." });
  }

  try {
    const parsed = parseExcel(req.file.path);
    const store = await readStore();

    store.employees = parsed.employees;
    store.shifts = parsed.shifts;
    store.meta.dates = parsed.dates;
    store.meta.dateRange = parsed.dateRange;
    store.meta.lastImportAt = new Date().toISOString();
    store.meta.onsiteCodes = (store.meta.onsiteCodes || ["1"]).map((c) =>
      cleanText(c).toUpperCase()
    );
    ensureStoreHasYearDates(store, 2026);
    store.meta.discoveredCodes = Array.from(
      new Set([...(store.meta.discoveredCodes || []), ...parsed.discoveredCodes])
    ).sort();

    store.audit.push({
      at: store.meta.lastImportAt,
      action: "import-excel",
      sheetName: parsed.sheetName,
      employees: parsed.employees.length,
      dates: parsed.dates.length,
      fileName: req.file.originalname,
      calendarFix: parsed.calendarFix,
    });
    store.audit = store.audit.slice(-5000);

    await writeStore(store);

    res.json({
      ok: true,
      employees: parsed.employees.length,
      dates: parsed.dates.length,
      dateRange: parsed.dateRange,
      codes: store.meta.discoveredCodes,
      calendarFix: parsed.calendarFix,
    });
  } catch (error) {
    res.status(400).json({ error: error.message || "Error procesando el Excel." });
  } finally {
    fs.rm(req.file.path, { force: true }, () => {});
  }
});

// JSON error handler (important for Netlify Functions: avoids HTML 500 responses)
app.use((err, _req, res, _next) => {
  try {
    console.error(err);
  } catch (_) {
    // ignore
  }
  res.status(500).json({ error: err?.message || "Error interno del servidor." });
});

app.get("/{*all}", (_req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

if (require.main === module) {
  (async () => {
    ensureDirs();

    const bootStore = await readStore();

    const fixedCalendar = normalizeStoreCalendarContinuity(bootStore);
    const fixedEmployees = normalizeStoredEmployees(bootStore);
    const fixedYear = ensureStoreHasYearDates(bootStore, 2026);
    if (fixedCalendar || fixedEmployees || fixedYear) {
      await writeStore(bootStore);
    }

    // Auto-import solo tiene sentido en modo archivo (local), no en serverless/Supabase.
    if (STORE_BACKEND === "file" && !bootStore.meta.lastImportAt) {
      const candidate = fs
        .readdirSync(__dirname)
        .find((file) => file.toLowerCase().endsWith(".xlsx"));

      if (candidate) {
        try {
          const parsed = parseExcel(path.join(__dirname, candidate));
          bootStore.employees = parsed.employees;
          bootStore.shifts = parsed.shifts;
          bootStore.meta.dates = parsed.dates;
          bootStore.meta.dateRange = parsed.dateRange;
          bootStore.meta.lastImportAt = new Date().toISOString();
          bootStore.meta.onsiteCodes = ["1"];
          bootStore.meta.discoveredCodes = Array.from(
            new Set([...DEFAULT_CODES, ...parsed.discoveredCodes])
          ).sort();
          bootStore.audit.push({
            at: bootStore.meta.lastImportAt,
            action: "auto-import-first-run",
            fileName: candidate,
            employees: parsed.employees.length,
            dates: parsed.dates.length,
            calendarFix: parsed.calendarFix,
          });
          await writeStore(bootStore);
        } catch (_error) {
          // La app sigue funcionando aunque falle el auto-import.
        }
      }
    }

    app.listen(PORT, HOST, () => {
      console.log(`Roster web disponible en http://localhost:${PORT}`);
      const ips = getLanIps();
      if (ips.length) {
        console.log("Acceso desde otra PC (misma red):");
        for (const ip of ips) {
          console.log(`  http://${ip}:${PORT}`);
        }
      }
    });
  })().catch((err) => {
    // Fallo fatal de arranque
    console.error(err);
    process.exitCode = 1;
  });
}

module.exports = app;
