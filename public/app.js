const state = {
  codes: [],
  employees: [],
  selectedProfileEmployeeId: "",
  months: [],
  selectedMonth: "",
  summary: null,
  roster: null,
  kpiRoster: null,
  conflicts: [],
  filters: {
    employeeIds: [],
    search: "",
    from: "",
    to: "",
    categoria: "",
    residencia: "",
    grupo: "",
  },
  minCoverage: 2,
  alertConfig: {
    from: "",
    to: "",
    autoRange: true,
  },
  alertPreview: null,
  fitAllDays: true,
  calendarOnly: true,
  selectedUpDate: "",
  onsiteCode: "1",
};

const ui = {
  importForm: document.getElementById("importForm"),
  excelInput: document.getElementById("excelInput"),
  importMsg: document.getElementById("importMsg"),
  monthInput: document.getElementById("monthInput"),
  calendarEmployeeInput: document.getElementById("calendarEmployeeInput"),
  clearCalendarEmployeeFilterBtn: document.getElementById("clearCalendarEmployeeFilterBtn"),
  prevMonthBtn: document.getElementById("prevMonthBtn"),
  nextMonthBtn: document.getElementById("nextMonthBtn"),
  searchInput: document.getElementById("searchInput"),
  fromInput: document.getElementById("fromInput"),
  toInput: document.getElementById("toInput"),
  categoriaInput: document.getElementById("categoriaInput"),
  residenciaInput: document.getElementById("residenciaInput"),
  grupoInput: document.getElementById("grupoInput"),
  coverageInput: document.getElementById("coverageInput"),
  refreshBtn: document.getElementById("refreshBtn"),
  refreshConflictsBtn: document.getElementById("refreshConflictsBtn"),
  summaryCards: document.getElementById("summaryCards"),
  codeLegend: document.getElementById("codeLegend"),
  conflictsList: document.getElementById("conflictsList"),
  rosterTableWrap: document.getElementById("rosterTableWrap"),
  tableMeta: document.getElementById("tableMeta"),
  downloadBtn: document.getElementById("downloadBtn"),
  alertFromInput: document.getElementById("alertFromInput"),
  alertToInput: document.getElementById("alertToInput"),
  alertPreview: document.getElementById("alertPreview"),
  emailLink: document.getElementById("emailLink"),
  fitDaysInput: document.getElementById("fitDaysInput"),
  profileEmployeeInput: document.getElementById("profileEmployeeInput"),
  profileCard: document.getElementById("profileCard"),
  upDateInput: document.getElementById("upDateInput"),
  upPeopleMeta: document.getElementById("upPeopleMeta"),
  upPeopleList: document.getElementById("upPeopleList"),
  onsiteMetrics: document.getElementById("onsiteMetrics"),
  onsiteCodesInput: document.getElementById("onsiteCodesInput"),
  areaTabs: [...document.querySelectorAll("[data-area-tab]")],
  areaSections: [...document.querySelectorAll("[data-area-section]")],
  codeMeaning: document.getElementById("codeMeaning"),
};

const AUTO_REFRESH_INTERVAL_MS = 60 * 1000;
const DEFAULT_ALERT_SITE = "RIO TINTO CAMP 1500";
const dragFill = {
  active: false,
  applying: false,
  skipNextClick: false,
  anchorEmployeeId: "",
  anchorDate: "",
  employeeId: "",
  code: "",
  moved: false,
  keys: new Set(),
  buttons: new Map(),
  originalCodes: new Map(),
};

const fillHandle = document.createElement("div");
fillHandle.className = "fill-handle";
fillHandle.title = "Arrastrar para copiar codigo";
fillHandle.hidden = true;
ui.rosterTableWrap.appendChild(fillHandle);

const UNDO_LIMIT = 80;
const undoStack = [];

function toIsoDate(date) {
  return date.toISOString().slice(0, 10);
}

function fmtDate(isoDate) {
  if (!isoDate) return "-";
  return new Date(`${isoDate}T00:00:00Z`).toLocaleDateString("es-AR", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    timeZone: "UTC",
  });
}

function fmtHeader(isoDate) {
  const date = new Date(`${isoDate}T00:00:00Z`);
  return {
    dayMonth: date.toLocaleDateString("es-AR", {
      day: "2-digit",
      month: "2-digit",
      timeZone: "UTC",
    }),
    weekday: date.toLocaleDateString("es-AR", {
      weekday: "short",
      timeZone: "UTC",
    }),
  };
}

function setActiveArea(area) {
  for (const tab of ui.areaTabs) {
    tab.classList.toggle("is-active", tab.dataset.areaTab === area);
  }

  for (const section of ui.areaSections) {
    section.hidden = section.dataset.areaSection !== area;
  }

  if (area === "planificacion") {
    requestAnimationFrame(() => {
      renderRoster();
    });
  }
}

function fmtMonthLabelFromIso(isoDate) {
  const date = new Date(`${isoDate}T00:00:00Z`);
  const month = date.toLocaleDateString("es-AR", {
    month: "long",
    timeZone: "UTC",
  });
  const year = date.toLocaleDateString("es-AR", {
    year: "numeric",
    timeZone: "UTC",
  });
  return `${month} ${year}`.toUpperCase();
}

function normalizeTextLocal(value) {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function getRoleKind(categoria) {
  const normalized = normalizeTextLocal(categoria);
  if (normalized.includes("coordinador")) return "coordinador";
  if (normalized.includes("inspector")) return "inspector";
  return "otro";
}

const KNOWN_FIRST_NAMES = new Set([
  "ALEJANDRA",
  "CAROLINA",
  "GABRIEL",
  "GABRIELA",
  "GUADALUPE",
  "JOSEFINA",
  "JUAN",
  "JULIAN",
  "RODRIGO",
  "SAUL",
  "SILVINA",
]);

function formatDisplayName(rawName) {
  const name = String(rawName || "").trim();
  if (!name) return "-";
  const parts = name.split(/\s+/);
  if (parts.length < 2) return name;

  const first = parts[0].toUpperCase();
  const second = parts[1].toUpperCase();
  const firstLooksName = KNOWN_FIRST_NAMES.has(first);
  const secondLooksName = KNOWN_FIRST_NAMES.has(second);

  if (!firstLooksName && secondLooksName) {
    return [parts[1], parts[0], ...parts.slice(2)].join(" ");
  }
  return parts.join(" ");
}

function monthRangeFromKey(monthKey) {
  if (!/^\d{4}-\d{2}$/.test(monthKey)) return null;
  const [year, month] = monthKey.split("-").map(Number);
  const from = `${monthKey}-01`;
  const last = new Date(Date.UTC(year, month, 0));
  const to = toIsoDate(last);
  return { from, to };
}

function setImportMessage(text, className = "muted") {
  if (!ui.importMsg) return;
  const safeText = String(text ?? "").trim();
  if (!safeText) {
    ui.importMsg.textContent = "";
    ui.importMsg.className = "muted";
    ui.importMsg.hidden = true;
    return;
  }
  ui.importMsg.hidden = false;
  ui.importMsg.textContent = safeText;
  ui.importMsg.className = className;
}

function getCacheKey(url) {
  return `roster-cache:${url}`;
}

function shouldCacheGet(url) {
  // Cache only slow/stable reference data. Anything user-facing like roster/alerts
  // must be fresh to avoid showing stale messages during presentations.
  return (
    url === "/api/summary" ||
    url === "/api/codes" ||
    url === "/api/months" ||
    url === "/api/employees"
  );
}

function toUserErrorMessage(error) {
  const raw = String(error?.message || "");
  if (raw.toLowerCase().includes("failed to fetch")) {
    let origin =
      typeof window !== "undefined" && window.location ? String(window.location.origin || "") : "";
    if (!origin || origin === "null" || origin.startsWith("file:")) {
      origin = "http://localhost:3000";
    }
    return `No hay conexion con el servidor. Ejecuta 'npm start' y abri ${origin}`;
  }
  return raw || "Error inesperado.";
}

async function api(url, options = {}) {
  const method = (options.method || "GET").toUpperCase();
  try {
    const response = await fetch(url, options);
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(payload.error || "Error de API");
    }
    if (method === "GET" && shouldCacheGet(url)) {
      localStorage.setItem(getCacheKey(url), JSON.stringify(payload));
    }
    return payload;
  } catch (error) {
    if (method === "GET" && shouldCacheGet(url)) {
      const cached = localStorage.getItem(getCacheKey(url));
      if (cached) return JSON.parse(cached);
    }
    throw error;
  }
}

function buildQuery() {
  const params = new URLSearchParams();
  for (const [key, value] of Object.entries(state.filters)) {
    if (Array.isArray(value)) {
      if (value.length) params.set(key, value.join(","));
      continue;
    }
    if (value) params.set(key, value);
  }
  return params.toString();
}

function fillSelect(selectEl, values, currentValue) {
  const firstOption = selectEl.options[0];
  selectEl.innerHTML = "";
  selectEl.appendChild(firstOption);
  for (const value of values) {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    selectEl.appendChild(option);
  }
  selectEl.value = currentValue || "";
}

function getRolePriority(roleKind) {
  const rolePriority = { coordinador: 0, inspector: 1, otro: 2 };
  return rolePriority[roleKind] ?? 9;
}

function compareByRoleThenName(a, b) {
  const roleDiff = getRolePriority(getRoleKind(a?.categoria)) - getRolePriority(getRoleKind(b?.categoria));
  if (roleDiff !== 0) return roleDiff;
  return formatDisplayName(a?.name).localeCompare(formatDisplayName(b?.name), "es", {
    sensitivity: "base",
  });
}

function fillMonthSelect() {
  ui.monthInput.innerHTML = "";
  const customOption = document.createElement("option");
  customOption.value = "";
  customOption.textContent = "Rango personalizado";
  ui.monthInput.appendChild(customOption);

  const groups = new Map();
  for (const month of state.months) {
    const year = String(month.month || "").slice(0, 4) || "OTROS";
    if (!groups.has(year)) groups.set(year, []);
    groups.get(year).push(month);
  }

  for (const [year, months] of groups.entries()) {
    const optgroup = document.createElement("optgroup");
    optgroup.label = year;
    for (const month of months) {
      const option = document.createElement("option");
      option.value = month.month;
      option.textContent = month.label;
      optgroup.appendChild(option);
    }
    ui.monthInput.appendChild(optgroup);
  }

  ui.monthInput.value = state.selectedMonth || "";
}

function fillEmployeeSelect(selectEl, selectedId, includeAllLabel) {
  const sortedEmployees = [...state.employees].sort(compareByRoleThenName);
  selectEl.innerHTML = "";

  if (selectEl.multiple) {
    const selectedSet = new Set(Array.isArray(selectedId) ? selectedId : []);
    for (const employee of sortedEmployees) {
      const option = document.createElement("option");
      option.value = employee.id;
      option.textContent = formatDisplayName(employee.name);
      option.selected = selectedSet.has(employee.id);
      selectEl.appendChild(option);
    }
    return;
  }

  const firstOption = document.createElement("option");
  firstOption.value = "";
  firstOption.textContent = includeAllLabel;
  selectEl.appendChild(firstOption);

  for (const employee of sortedEmployees) {
    const option = document.createElement("option");
    option.value = employee.id;
    option.textContent = formatDisplayName(employee.name);
    selectEl.appendChild(option);
  }
  selectEl.value = selectedId || "";
}

function fillUpDateSelect(dates) {
  if (!ui.upDateInput) return;

  ui.upDateInput.innerHTML = "";
  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = "Seleccionar fecha";
  ui.upDateInput.appendChild(placeholder);

  if (!dates.length) {
    ui.upDateInput.disabled = true;
    state.selectedUpDate = "";
    ui.upDateInput.value = "";
    return;
  }

  for (const dateIso of dates) {
    const option = document.createElement("option");
    option.value = dateIso;
    option.textContent = `${fmtMonthLabelFromIso(dateIso)} - ${fmtDate(dateIso)}`;
    ui.upDateInput.appendChild(option);
  }

  const todayIso = toIsoDate(new Date());
  if (!state.selectedUpDate || !dates.includes(state.selectedUpDate)) {
    state.selectedUpDate = dates.includes(todayIso) ? todayIso : dates[0];
  }

  ui.upDateInput.disabled = false;
  ui.upDateInput.value = state.selectedUpDate;
}

function renderUpPeoplePanel() {
  if (!ui.upPeopleMeta || !ui.upPeopleList || !ui.upDateInput || !ui.onsiteMetrics) return;

  const kpiSource = state.kpiRoster || state.roster;
  const dates = kpiSource?.dates || [];
  const employees = kpiSource?.employees || [];
  fillUpDateSelect(dates);
  ui.upPeopleList.innerHTML = "";

  if (!dates.length || !employees.length || !state.selectedUpDate) {
    ui.upPeopleMeta.textContent = "Sin datos para mostrar.";
    ui.onsiteMetrics.innerHTML = "";
    const item = document.createElement("li");
    item.className = "empty";
    item.textContent = "No hay personal para el rango actual.";
    ui.upPeopleList.appendChild(item);
    return;
  }

  const onsiteCodes = String(state.onsiteCode || "")
    .split(/[,\s]+/)
    .map((code) => code.trim().toUpperCase())
    .filter(Boolean);
  const onsiteCodeSet = new Set(onsiteCodes.length ? onsiteCodes : ["1"]);
  const coordinadorEmployees = employees.filter((employee) =>
    normalizeTextLocal(employee.categoria).includes("coordinador")
  );
  const inspectorEmployees = employees.filter((employee) =>
    normalizeTextLocal(employee.categoria).includes("inspector")
  );

  const countOnsiteFor = (list) =>
    list.reduce((count, employee) => {
      const code = (employee.shifts[state.selectedUpDate] || "").toUpperCase();
      return onsiteCodeSet.has(code) ? count + 1 : count;
    }, 0);

  const coordinadorOnsite = countOnsiteFor(coordinadorEmployees);
  const inspectorOnsite = countOnsiteFor(inspectorEmployees);
  const totalOnsite = coordinadorOnsite + inspectorOnsite;

  const metrics = [
    { label: "Coordinadores en sitio", value: coordinadorOnsite, className: "metric-coordinador" },
    { label: "Inspectores en sitio", value: inspectorOnsite, className: "metric-inspector" },
    { label: "Total en sitio", value: totalOnsite, className: "metric-total" },
  ];
  ui.onsiteMetrics.innerHTML = metrics
    .map(
      (metric) =>
        `<article class="onsite-metric-card ${metric.className}"><p>${metric.label}</p><strong>${metric.value}</strong></article>`
    )
    .join("");

  const people = employees
    .filter((employee) => onsiteCodeSet.has((employee.shifts[state.selectedUpDate] || "").toUpperCase()))
    .map((employee) => ({
      name: formatDisplayName(employee.name),
      categoria: employee.categoria || "-",
      roleKind: getRoleKind(employee.categoria),
    }))
    .sort((a, b) => {
      const roleDiff = getRolePriority(a.roleKind) - getRolePriority(b.roleKind);
      if (roleDiff !== 0) return roleDiff;
      return a.name.localeCompare(b.name, "es", { sensitivity: "base" });
    });

  const plural = people.length === 1 ? "persona" : "personas";
  ui.upPeopleMeta.textContent = `${fmtMonthLabelFromIso(
    state.selectedUpDate
  )} | ${fmtDate(state.selectedUpDate)} | Codigos ${[...onsiteCodeSet].join(
    ", "
  )} | ${people.length} ${plural} arriba`;

  if (!people.length) {
    const item = document.createElement("li");
    item.className = "empty";
    item.textContent = "No hay personal arriba ese dia.";
    ui.upPeopleList.appendChild(item);
    return;
  }

  for (const person of people) {
    const item = document.createElement("li");
    item.className = `up-person-item role-${person.roleKind}`;
    const nameEl = document.createElement("strong");
    nameEl.textContent = person.name;
    const roleEl = document.createElement("span");
    roleEl.textContent = person.categoria;
    item.appendChild(nameEl);
    item.appendChild(roleEl);
    ui.upPeopleList.appendChild(item);
  }
}

function renderProfileCard() {
  if (!state.selectedProfileEmployeeId) {
    ui.profileCard.textContent = "Selecciona una persona para ver su ficha.";
    return;
  }

  const employee = state.employees.find((item) => item.id === state.selectedProfileEmployeeId);
  if (!employee) {
    ui.profileCard.textContent = "No se encontro la persona seleccionada.";
    return;
  }

  ui.profileCard.innerHTML = `
    <p><strong>Nombre:</strong> ${formatDisplayName(employee.name)}</p>
    <p><strong>Categoria:</strong> ${employee.categoria || "-"}</p>
    <p><strong>Residencia:</strong> ${employee.residencia || "-"}</p>
    <p><strong>Subida/Bajada:</strong> ${employee.ascensoDescenso || "-"}</p>
    <p><strong>DNI:</strong> ${employee.dni || "-"}</p>
    <p><strong>Telefono:</strong> ${employee.telefono || "-"}</p>
    <p><strong>Grupo:</strong> ${employee.grupo || "-"}</p>
    <p><strong>Tipo:</strong> ${employee.tipo || "-"}</p>
  `;
}

const CODE_MEANINGS = {
  S: "SUBIDA",
  D: "DESCANSO",
  "1": "DIA EN OBRA",
  B: "BAJADA",
};

function renderCodeMeaning() {
  if (!ui.codeMeaning) return;
  const order = ["S", "D", "1", "B"];
  const existing = new Set((state.codes || []).map((c) => String(c || "").toUpperCase()));
  const toShow = order.filter((code) => CODE_MEANINGS[code] && (!existing.size || existing.has(code)));

  ui.codeMeaning.innerHTML = toShow
    .map((code) => {
      const cls = `code-pill code-${String(code).toLowerCase()}`;
      return `<span class="${cls}"><span class="code">${code}</span>${CODE_MEANINGS[code]}</span>`;
    })
    .join("");
}

function syncMonthStateFromDateRange() {
  const matched = state.months.find(
    (month) => month.from === state.filters.from && month.to === state.filters.to
  );
  state.selectedMonth = matched ? matched.month : "";
  ui.monthInput.value = state.selectedMonth;
}

async function applySelectedMonth({ refresh = true } = {}) {
  if (!state.selectedMonth) return;
  const month = state.months.find((item) => item.month === state.selectedMonth);
  if (!month) return;

  state.filters.from = month.from;
  state.filters.to = month.to;
  ui.fromInput.value = month.from;
  ui.toInput.value = month.to;
  ui.monthInput.value = month.month;

  if (state.alertConfig.autoRange) {
    state.alertConfig.from = month.from;
    state.alertConfig.to = month.to;
    if (ui.alertFromInput) ui.alertFromInput.value = month.from;
    if (ui.alertToInput) ui.alertToInput.value = month.to;
  }

  if (refresh) {
    await refreshRosterRelated();
  }
}

async function moveMonth(delta) {
  if (!state.months.length) return;
  let index = state.months.findIndex((month) => month.month === state.selectedMonth);
  if (index < 0) index = 0;
  const nextIndex = Math.min(state.months.length - 1, Math.max(0, index + delta));
  state.selectedMonth = state.months[nextIndex].month;
  await applySelectedMonth({ refresh: true });
}

function renderSummary() {
  if (!state.summary) {
    ui.summaryCards.innerHTML = "<p class='muted'>Sin datos.</p>";
    return;
  }

  const cards = [
    ["Personal", state.summary.employees],
    ["Asignaciones", state.summary.assignments],
    ["Dias", state.summary.dates],
    ["Alertas", state.summary.conflicts],
    ["Desde", fmtDate(state.summary.dateRange?.from)],
    ["Hasta", fmtDate(state.summary.dateRange?.to)],
  ];

  ui.summaryCards.innerHTML = cards
    .map(
      ([label, value]) =>
        `<article class="stat-card"><p>${label}</p><strong>${value ?? "-"}</strong></article>`
    )
    .join("");

  ui.codeLegend.textContent = state.codes.length
    ? state.codes.map((code) => ` ${code} `).join(" | ")
    : "Sin codigos";
}

function renderRoster() {
  if (!state.roster || !state.roster.employees.length || !state.roster.dates.length) {
    ui.rosterTableWrap.innerHTML =
      "<div class='empty'>No hay datos para el rango/filtro seleccionado.</div>";
    ui.tableMeta.textContent = "";
    renderUpPeoplePanel();
    return;
  }

  const { dates, totalEmployees } = state.roster;
  const selectedSet = new Set(state.filters.employeeIds || []);
  const employees = [...state.roster.employees]
    .filter((employee) => !selectedSet.size || selectedSet.has(employee.id))
    .sort(compareByRoleThenName);
  const filteredEmployees = employees.length;
  const selectedMonthLabel =
    state.months.find((month) => month.month === state.selectedMonth)?.label ||
    "Rango personalizado";
  const leftColumns = [
    { key: "name", label: "", stickyClass: "sticky-left-1", width: 280 },
  ];

  ui.tableMeta.textContent = `Mostrando ${filteredEmployees} de ${totalEmployees} personas | Mes: ${selectedMonthLabel} | Rango visible: ${fmtDate(
    dates[0]
  )} a ${fmtDate(dates[dates.length - 1])}`;

  if (!employees.length) {
    ui.rosterTableWrap.innerHTML =
      "<div class='empty'>No hay personas con el filtro seleccionado.</div>";
    renderUpPeoplePanel();
    return;
  }

  const table = document.createElement("table");
  table.className = "roster-table";

  const thead = document.createElement("thead");
  const monthRow = document.createElement("tr");
  monthRow.className = "month-row";
  leftColumns.forEach((col) => {
    const th = document.createElement("th");
    th.textContent = col.label;
    th.rowSpan = 2;
    th.classList.add("sticky", "rowspan-head", col.stickyClass);
    monthRow.appendChild(th);
  });

  let i = 0;
  while (i < dates.length) {
    const monthKey = dates[i].slice(0, 7);
    let span = 1;
    while (i + span < dates.length && dates[i + span].slice(0, 7) === monthKey) {
      span += 1;
    }
    const th = document.createElement("th");
    th.colSpan = span;
    th.className = "month-group";
    th.textContent = fmtMonthLabelFromIso(dates[i]);
    monthRow.appendChild(th);
    i += span;
  }
  thead.appendChild(monthRow);

  const dayRow = document.createElement("tr");
  dayRow.className = "day-row";
  dates.forEach((dateIso) => {
    const th = document.createElement("th");
    th.classList.add("day-head");
    const header = fmtHeader(dateIso);
    th.innerHTML = `<span class="date-main">${header.dayMonth}</span><span class="date-sub">${header.weekday}</span>`;
    th.title = fmtDate(dateIso);
    if (dateIso.endsWith("-01")) th.classList.add("month-start");
    dayRow.appendChild(th);
  });
  thead.appendChild(dayRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  for (const employee of employees) {
    const row = document.createElement("tr");
    leftColumns.forEach((col) => {
      const td = document.createElement("td");
      if (col.key === "name") {
        const displayName = formatDisplayName(employee.name);
        td.innerHTML = `
          <div class="person-cell">
            <span class="person-name">${displayName}</span>
            <span class="person-role">${employee.categoria || "-"}</span>
          </div>
        `;
      } else {
        const value = employee[col.key];
        td.textContent = value || "-";
      }
      td.className = "sticky";
      td.classList.add(col.stickyClass);
      row.appendChild(td);
    });

    for (const date of dates) {
      const td = document.createElement("td");
      td.classList.add("day-cell");
      const code = employee.shifts[date] || "";
      if (date.endsWith("-01")) td.classList.add("month-start-cell");
      td.innerHTML = `<button class="code-cell ${code ? `code-${code.toLowerCase()}` : "code-empty"}" data-emp="${employee.id}" data-date="${date}" data-code="${code}">${code || "."}</button>`;
      row.appendChild(td);
    }

    tbody.appendChild(row);
  }

  table.appendChild(tbody);
  const fixedColsWidth = leftColumns.reduce((acc, col) => acc + col.width, 0);
  applyFitDaysLayout(table, dates.length, fixedColsWidth);
  ui.rosterTableWrap.innerHTML = "";
  ui.rosterTableWrap.appendChild(table);
  ui.rosterTableWrap.appendChild(fillHandle);
  refreshFillHandleAnchor();
  renderUpPeoplePanel();
}

function applyFitDaysLayout(table, dayCount, fixedColsWidth) {
  if (!table || !dayCount) return;
  // Even if the user disables "fit all days", we still auto-fit when the month
  // would overflow the viewport. Otherwise 31-day months look "broken".
  const wrapWidth = Math.max(ui.rosterTableWrap.clientWidth - 24, 360);
  const naturalDayWidth = 46; // matches default CSS width for day columns
  const naturalTotal = fixedColsWidth + dayCount * naturalDayWidth;
  const mustFit = state.fitAllDays || naturalTotal > wrapWidth;

  if (!mustFit) {
    table.classList.remove("fit-days");
    table.style.removeProperty("--day-col-width");
    return;
  }

  table.classList.add("fit-days");

  const availableForDays = Math.max(140, wrapWidth - fixedColsWidth);
  const computed = Math.floor(availableForDays / dayCount);
  const dayWidth = Math.max(16, Math.min(46, computed));
  table.style.setProperty("--day-col-width", `${dayWidth}px`);
}

function renderConflicts() {
  if (!state.conflicts.length) {
    ui.conflictsList.innerHTML = "<li class='empty'>Sin alertas en el rango actual.</li>";
    return;
  }
  ui.conflictsList.innerHTML = state.conflicts
    .slice(0, 120)
    .map((conflict) => `<li class="sev-${conflict.severity || "low"}">${conflict.message}</li>`)
    .join("");
}

function renderAlertPreview() {
  if (!state.alertPreview) {
    ui.alertPreview.textContent = "Genera una vista previa para enviar.";
    ui.emailLink.href = "#";
    return;
  }
  ui.alertPreview.textContent = state.alertPreview.message || "Sin contenido.";
  ui.emailLink.href = state.alertPreview.mailtoUrl || "#";
}

function refreshDownloadLink() {
  ui.downloadBtn.href = `/api/export/roster.xlsx?${buildQuery()}`;
}

function nextCode(currentCode) {
  const chain = ["", ...state.codes];
  const idx = chain.indexOf(currentCode || "");
  return chain[(idx + 1) % chain.length];
}

function setCodeCellVisual(button, code) {
  const normalized = String(code || "").toUpperCase();
  button.dataset.code = normalized;
  button.textContent = normalized || ".";

  [...button.classList]
    .filter((cls) => cls.startsWith("code-"))
    .forEach((cls) => button.classList.remove(cls));

  button.classList.add(normalized ? `code-${normalized.toLowerCase()}` : "code-empty");
}

function findCellButton(employeeId, date) {
  const buttons = ui.rosterTableWrap.querySelectorAll(".code-cell");
  for (const button of buttons) {
    if (button.dataset.emp === employeeId && button.dataset.date === date) {
      return button;
    }
  }
  return null;
}

function positionFillHandle(button) {
  const wrapRect = ui.rosterTableWrap.getBoundingClientRect();
  const buttonRect = button.getBoundingClientRect();
  const size = fillHandle.offsetWidth || 16;
  const inset = 2;
  const rawLeft = buttonRect.right - wrapRect.left - size - inset;
  const rawTop = buttonRect.bottom - wrapRect.top - size - inset;
  const maxLeft = Math.max(0, ui.rosterTableWrap.clientWidth - size - inset);
  const maxTop = Math.max(0, ui.rosterTableWrap.clientHeight - size - inset);
  const left = Math.max(0, Math.min(rawLeft, maxLeft));
  const top = Math.max(0, Math.min(rawTop, maxTop));
  fillHandle.style.left = `${left}px`;
  fillHandle.style.top = `${top}px`;
  fillHandle.hidden = false;
}

function setFillHandleAnchor(button) {
  if (!button || dragFill.active || dragFill.applying) return;
  dragFill.anchorEmployeeId = button.dataset.emp || "";
  dragFill.anchorDate = button.dataset.date || "";
  positionFillHandle(button);
}

function refreshFillHandleAnchor() {
  if (dragFill.active || dragFill.applying) {
    fillHandle.hidden = true;
    return;
  }
  if (!dragFill.anchorEmployeeId || !dragFill.anchorDate) {
    fillHandle.hidden = true;
    return;
  }
  const button = findCellButton(dragFill.anchorEmployeeId, dragFill.anchorDate);
  if (!button) {
    fillHandle.hidden = true;
    return;
  }
  positionFillHandle(button);
}

function startDragFromButton(button) {
  dragFill.active = true;
  dragFill.applying = false;
  dragFill.employeeId = button.dataset.emp || "";
  dragFill.code = (button.dataset.code || "").toUpperCase();
  dragFill.moved = false;
  dragFill.keys.clear();
  dragFill.buttons.clear();
  dragFill.originalCodes.clear();
  fillHandle.hidden = true;
  addDragTarget(button);
}

function addDragTarget(button) {
  if (!dragFill.active) return;
  const employeeId = button.dataset.emp;
  const date = button.dataset.date;
  if (!employeeId || !date) return;
  if (employeeId !== dragFill.employeeId) return;

  const key = `${employeeId}|${date}`;
  if (dragFill.keys.has(key)) return;

  dragFill.keys.add(key);
  dragFill.buttons.set(key, button);
  dragFill.originalCodes.set(key, button.dataset.code || "");
  button.classList.add("drag-target");
  setCodeCellVisual(button, dragFill.code);
  if (dragFill.keys.size > 1) dragFill.moved = true;
}

function clearDragPreview({ restoreOriginal = false } = {}) {
  for (const [key, button] of dragFill.buttons.entries()) {
    button.classList.remove("drag-target");
    if (restoreOriginal) {
      const original = dragFill.originalCodes.get(key) || "";
      setCodeCellVisual(button, original);
    }
  }

  dragFill.active = false;
  dragFill.employeeId = "";
  dragFill.code = "";
  dragFill.moved = false;
  dragFill.keys.clear();
  dragFill.buttons.clear();
  dragFill.originalCodes.clear();
  refreshFillHandleAnchor();
}

function pushUndoEntry(changes) {
  const entry = changes
    .filter((change) => String(change.fromCode || "") !== String(change.toCode || ""))
    .map((change) => ({
      employeeId: change.employeeId,
      date: change.date,
      fromCode: String(change.fromCode || "").toUpperCase(),
      toCode: String(change.toCode || "").toUpperCase(),
    }));

  if (!entry.length) return;
  undoStack.push(entry);
  if (undoStack.length > UNDO_LIMIT) {
    undoStack.shift();
  }
}

async function applyShiftUpdates(changes, { recordUndo = true, statusMessage = "" } = {}) {
  const normalized = changes
    .map((change) => ({
      employeeId: change.employeeId,
      date: change.date,
      fromCode: String(change.fromCode || "").toUpperCase(),
      toCode: String(change.toCode || "").toUpperCase(),
    }))
    .filter((change) => change.employeeId && change.date);

  const effective = normalized.filter((change) => change.fromCode !== change.toCode);
  if (!effective.length) return;

  if (statusMessage) setImportMessage(statusMessage, "muted");

  await Promise.all(
    effective.map((change) =>
      api("/api/shifts", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          employeeId: change.employeeId,
          date: change.date,
          code: change.toCode,
        }),
      })
    )
  );

  for (const change of effective) {
    const target = state.roster?.employees?.find((emp) => emp.id === change.employeeId);
    if (!target) continue;
    if (change.toCode) target.shifts[change.date] = change.toCode;
    else delete target.shifts[change.date];
  }
  for (const change of effective) {
    const target = state.kpiRoster?.employees?.find((emp) => emp.id === change.employeeId);
    if (!target) continue;
    if (change.toCode) target.shifts[change.date] = change.toCode;
    else delete target.shifts[change.date];
  }

  if (recordUndo) pushUndoEntry(effective);

  if (state.calendarOnly) {
    await Promise.all([loadSummary(), loadCodes(), loadAlertPreview(), loadKpiRoster()]);
  } else {
    await Promise.all([loadSummary(), loadCodes(), loadConflicts(), loadAlertPreview()]);
    renderSummary();
    renderConflicts();
    renderAlertPreview();
  }

  renderRoster();
  renderAlertPreview();

  setImportMessage("Guardado.", "ok");
  window.clearTimeout(applyShiftUpdates._clearTimer);
  applyShiftUpdates._clearTimer = window.setTimeout(() => {
    setImportMessage("");
  }, 1200);
}

async function applyDragFill() {
  if (!dragFill.active || dragFill.applying) return;

  if (!dragFill.moved) {
    clearDragPreview({ restoreOriginal: false });
    return;
  }

  dragFill.applying = true;
  dragFill.skipNextClick = true;

  const updates = [];
  for (const key of dragFill.keys) {
    const [employeeId, date] = key.split("|");
    const fromCode = (dragFill.originalCodes.get(key) || "").toUpperCase();
    const toCode = (dragFill.code || "").toUpperCase();
    if (fromCode === toCode) continue;
    updates.push({ employeeId, date, fromCode, toCode });
  }

  if (!updates.length) {
    clearDragPreview({ restoreOriginal: false });
    dragFill.applying = false;
    return;
  }

  try {
    await applyShiftUpdates(updates, {
      recordUndo: true,
      statusMessage: "Aplicando arrastre en calendario...",
    });
  } catch (error) {
    setImportMessage(toUserErrorMessage(error), "error");
    await loadRoster();
    renderRoster();
  } finally {
    clearDragPreview({ restoreOriginal: false });
    dragFill.applying = false;
  }
}

async function loadSummary() {
  state.summary = await api("/api/summary");

  // Si el servidor trae configuracion y no hay override local, usarla.
  if (!localStorage.getItem("roster.onsiteCodes")) {
    const serverCodes = state.summary?.onsiteCodes;
    if (Array.isArray(serverCodes) && serverCodes.length) {
      const joined = serverCodes.map((c) => String(c || "").toUpperCase()).filter(Boolean).join(",");
      if (joined) {
        state.onsiteCode = joined;
        if (ui.onsiteCodesInput) ui.onsiteCodesInput.value = joined;
      }
    }
  }

  if (!state.filters.from && !state.filters.to && state.summary.dateRange?.from) {
    const dataFrom = state.summary.dateRange.from;
    const dataTo = state.summary.dateRange.to;
    const targetYearFrom = "2026-01-01";
    const targetYearTo = "2026-12-31";
    const has2026Window = dataTo >= targetYearFrom && dataFrom <= targetYearTo;

    if (has2026Window) {
      state.filters.from = dataFrom > targetYearFrom ? dataFrom : targetYearFrom;
      state.filters.to = dataTo < targetYearTo ? dataTo : targetYearTo;
    } else {
      state.filters.from = dataFrom;
      state.filters.to = dataTo;
    }
    ui.fromInput.value = state.filters.from;
    ui.toInput.value = state.filters.to;
  }
}

async function loadCodes() {
  const data = await api("/api/codes");
  state.codes = data.codes || [];
  renderCodeMeaning();
}

async function loadEmployees() {
  const data = await api("/api/employees");
  state.employees = data.employees || [];

  if (
    state.selectedProfileEmployeeId &&
    !state.employees.some((emp) => emp.id === state.selectedProfileEmployeeId)
  ) {
    state.selectedProfileEmployeeId = "";
  }

  const validEmployeeIds = new Set(state.employees.map((emp) => emp.id));
  state.filters.employeeIds = (state.filters.employeeIds || []).filter((id) =>
    validEmployeeIds.has(id)
  );

  fillEmployeeSelect(ui.calendarEmployeeInput, state.filters.employeeIds, "Todas las personas");
  fillEmployeeSelect(ui.profileEmployeeInput, state.selectedProfileEmployeeId, "Elegir persona");
  renderProfileCard();
}

async function loadMonths() {
  const data = await api("/api/months");
  state.months = (data.months || []).map((month) => ({
    ...month,
    label: fmtMonthLabelFromIso(month.from),
  }));

  if (!state.selectedMonth && state.months.length) {
    const currentMonthKey = toIsoDate(new Date()).slice(0, 7);
    const currentMonth = state.months.find((month) => month.month === currentMonthKey);
    const fromRange = state.months.find(
      (month) => month.from === state.filters.from && month.to === state.filters.to
    );
    const month2026 = state.months.find((month) => month.month.startsWith("2026-"));
    state.selectedMonth = (currentMonth || fromRange || month2026 || state.months[0]).month;
  } else if (
    state.selectedMonth &&
    !state.months.some((month) => month.month === state.selectedMonth)
  ) {
    state.selectedMonth = "";
  }

  fillMonthSelect();
}

async function loadRoster() {
  state.roster = await api(`/api/roster?${buildQuery()}`);
  fillSelect(ui.categoriaInput, state.roster.filterOptions?.categorias || [], state.filters.categoria);
  fillSelect(ui.residenciaInput, state.roster.filterOptions?.residencias || [], state.filters.residencia);
  fillSelect(ui.grupoInput, state.roster.filterOptions?.grupos || [], state.filters.grupo);
}

async function loadKpiRoster() {
  const params = new URLSearchParams();
  if (state.filters.from) params.set("from", state.filters.from);
  if (state.filters.to) params.set("to", state.filters.to);
  state.kpiRoster = await api(`/api/roster?${params.toString()}`);
}

async function loadConflicts() {
  const params = new URLSearchParams(buildQuery());
  params.set("minCoverage", String(state.minCoverage));
  const data = await api(`/api/conflicts?${params.toString()}`);
  state.conflicts = data.conflicts || [];
}

async function loadAlertPreview() {
  const params = new URLSearchParams();
  const from = state.alertConfig.from || state.filters.from;
  const to = state.alertConfig.to || state.filters.to;
  if (from) params.set("from", from);
  if (to) params.set("to", to);
  state.alertPreview = await api(`/api/alerts/preview?${params.toString()}`);
}

let fullRefreshInFlight = null;

async function fullRefresh() {
  if (fullRefreshInFlight) return fullRefreshInFlight;

  fullRefreshInFlight = (async () => {
    try {
      await Promise.all([loadSummary(), loadCodes(), loadMonths(), loadEmployees()]);
      if (state.selectedMonth) {
        await applySelectedMonth({ refresh: false });
      } else {
        syncMonthStateFromDateRange();
      }

      // Init diffusion range to current visible month unless user changes it.
      if (!state.alertConfig.from || !state.alertConfig.to || state.alertConfig.autoRange) {
        state.alertConfig.from = state.filters.from || "";
        state.alertConfig.to = state.filters.to || "";
        if (ui.alertFromInput) ui.alertFromInput.value = state.alertConfig.from;
        if (ui.alertToInput) ui.alertToInput.value = state.alertConfig.to;
      }
      await Promise.all([loadRoster(), loadKpiRoster()]);
      if (!state.calendarOnly) {
        await Promise.all([loadConflicts(), loadAlertPreview()]);
        renderSummary();
        renderConflicts();
      } else {
        await loadAlertPreview();
      }
      renderAlertPreview();
      renderRoster();
      refreshDownloadLink();
      setImportMessage("");
    } catch (error) {
      setImportMessage(toUserErrorMessage(error), "error");
    } finally {
      fullRefreshInFlight = null;
    }
  })();

  return fullRefreshInFlight;
}

async function refreshRosterRelated() {
  await Promise.all([loadRoster(), loadKpiRoster()]);
  if (!state.calendarOnly) {
    await Promise.all([loadConflicts(), loadAlertPreview()]);
  } else {
    await loadAlertPreview();
  }
  renderRoster();
  if (!state.calendarOnly) {
    renderConflicts();
  }
  renderAlertPreview();
  refreshDownloadLink();
}

if (ui.importForm && ui.excelInput) {
  ui.importForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    const file = ui.excelInput.files?.[0];
    if (!file) return;

    setImportMessage("Importando Excel...", "muted");
    try {
      const formData = new FormData();
      formData.append("file", file);
      const result = await api("/api/import-excel", { method: "POST", body: formData });
      const gapWarn = result.calendarFix?.wasFixed
        ? " Se corrigio automaticamente un salto de fechas del Excel."
        : "";
      setImportMessage(`Excel importado correctamente.${gapWarn}`, "ok");
      await fullRefresh();
    } catch (error) {
      setImportMessage(toUserErrorMessage(error), "error");
    }
  });
}

if (ui.refreshBtn) {
  ui.refreshBtn.addEventListener("click", () => {
    fullRefresh();
  });
}

ui.refreshConflictsBtn.addEventListener("click", async () => {
  state.minCoverage = Number(ui.coverageInput.value || "2");
  if (state.calendarOnly) return;
  await loadConflicts();
  renderConflicts();
});

ui.monthInput.addEventListener("change", async (event) => {
  state.selectedMonth = event.target.value;
  if (!state.selectedMonth) return;
  await applySelectedMonth({ refresh: true });
});

ui.calendarEmployeeInput.addEventListener("change", async () => {
  state.filters.employeeIds = [...ui.calendarEmployeeInput.selectedOptions]
    .map((option) => option.value)
    .filter(Boolean);
  await refreshRosterRelated();
});

ui.calendarEmployeeInput.addEventListener("mousedown", (event) => {
  const option = event.target.closest("option");
  if (!option) return;
  event.preventDefault();
  option.selected = !option.selected;
  state.filters.employeeIds = [...ui.calendarEmployeeInput.selectedOptions]
    .map((item) => item.value)
    .filter(Boolean);
  void refreshRosterRelated();
});

ui.clearCalendarEmployeeFilterBtn.addEventListener("click", async () => {
  state.filters.employeeIds = [];
  for (const option of ui.calendarEmployeeInput.options) {
    option.selected = false;
  }
  await refreshRosterRelated();
});

ui.prevMonthBtn.addEventListener("click", async () => {
  await moveMonth(-1);
});

ui.nextMonthBtn.addEventListener("click", async () => {
  await moveMonth(1);
});

ui.searchInput.addEventListener("input", async (event) => {
  state.filters.search = event.target.value.trim();
  await refreshRosterRelated();
});

ui.profileEmployeeInput.addEventListener("change", (event) => {
  state.selectedProfileEmployeeId = event.target.value;
  renderProfileCard();
});

ui.fromInput.addEventListener("change", async (event) => {
  state.filters.from = event.target.value;
  syncMonthStateFromDateRange();
  await refreshRosterRelated();
});

ui.toInput.addEventListener("change", async (event) => {
  state.filters.to = event.target.value;
  syncMonthStateFromDateRange();
  await refreshRosterRelated();
});

ui.categoriaInput.addEventListener("change", async (event) => {
  state.filters.categoria = event.target.value;
  await refreshRosterRelated();
});

ui.residenciaInput.addEventListener("change", async (event) => {
  state.filters.residencia = event.target.value;
  await refreshRosterRelated();
});

ui.grupoInput.addEventListener("change", async (event) => {
  state.filters.grupo = event.target.value;
  await refreshRosterRelated();
});

ui.rosterTableWrap.addEventListener("mouseover", (event) => {
  const button = event.target.closest(".code-cell");
  if (!button) return;
  if (dragFill.active) {
    addDragTarget(button);
  } else {
    setFillHandleAnchor(button);
  }
});

ui.rosterTableWrap.addEventListener("mouseleave", () => {
  if (!dragFill.active && !dragFill.applying) {
    fillHandle.hidden = true;
  }
});

fillHandle.addEventListener("mousedown", (event) => {
  if (event.button !== 0) return;
  event.preventDefault();
  event.stopPropagation();
  if (!dragFill.anchorEmployeeId || !dragFill.anchorDate) return;
  const button = findCellButton(dragFill.anchorEmployeeId, dragFill.anchorDate);
  if (!button) return;
  startDragFromButton(button);
});

ui.rosterTableWrap.addEventListener("click", async (event) => {
  if (dragFill.skipNextClick) {
    dragFill.skipNextClick = false;
    return;
  }
  if (dragFill.active || dragFill.applying) return;

  const button = event.target.closest(".code-cell");
  if (!button) return;

  const employeeId = button.dataset.emp;
  const date = button.dataset.date;
  const currentCode = (button.dataset.code || "").toUpperCase();
  const newCode = (event.shiftKey ? "" : nextCode(currentCode)).toUpperCase();
  dragFill.anchorEmployeeId = employeeId || "";
  dragFill.anchorDate = date || "";

  button.disabled = true;
  button.textContent = "...";

  try {
    await applyShiftUpdates(
      [{ employeeId, date, fromCode: currentCode, toCode: newCode }],
      { recordUndo: true }
    );
  } catch (error) {
    setImportMessage(toUserErrorMessage(error), "error");
    await loadRoster();
    renderRoster();
  }
});

document.addEventListener("mouseup", () => {
  void applyDragFill();
});

ui.fitDaysInput.addEventListener("change", () => {
  state.fitAllDays = ui.fitDaysInput.checked;
  localStorage.setItem("roster.fitAllDays", state.fitAllDays ? "1" : "0");
  renderRoster();
});

ui.upDateInput.addEventListener("change", (event) => {
  state.selectedUpDate = event.target.value;
  renderUpPeoplePanel();
});

  if (ui.onsiteCodesInput) {
    ui.onsiteCodesInput.addEventListener("change", async (event) => {
      const value = String(event.target.value || "").trim().toUpperCase() || "1";
      state.onsiteCode = value;
      localStorage.setItem("roster.onsiteCodes", value);

      try {
        await api("/api/config/onsite-codes", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ codes: value }),
      });
    } catch (_error) {
      // no-op (igual funciona local)
    }

    await loadAlertPreview();
    renderAlertPreview();
    renderUpPeoplePanel();
  });
}

if (ui.alertFromInput && ui.alertToInput) {
  const onAlertRangeChange = async () => {
    state.alertConfig.autoRange = false;
    state.alertConfig.from = ui.alertFromInput.value || "";
    state.alertConfig.to = ui.alertToInput.value || "";
    await loadAlertPreview();
    renderAlertPreview();
  };
  ui.alertFromInput.addEventListener("change", onAlertRangeChange);
  ui.alertToInput.addEventListener("change", onAlertRangeChange);
}

ui.rosterTableWrap.addEventListener("scroll", () => {
  refreshFillHandleAnchor();
});

for (const tab of ui.areaTabs) {
  tab.addEventListener("click", () => {
    setActiveArea(tab.dataset.areaTab || "planificacion");
  });
}

document.addEventListener("keydown", (event) => {
  const key = String(event.key || "").toLowerCase();
  const isUndo = (event.ctrlKey || event.metaKey) && !event.shiftKey && key === "z";
  if (!isUndo) return;

  const target = event.target;
  const tag = String(target?.tagName || "").toLowerCase();
  if (tag === "input" || tag === "textarea" || tag === "select" || target?.isContentEditable) {
    return;
  }

  if (dragFill.active || dragFill.applying) return;
  const lastChange = undoStack.pop();
  if (!lastChange?.length) return;

  event.preventDefault();
  const reverseChanges = lastChange.map((change) => ({
    employeeId: change.employeeId,
    date: change.date,
    fromCode: change.toCode,
    toCode: change.fromCode,
  }));

  (async () => {
    try {
      await applyShiftUpdates(reverseChanges, {
        recordUndo: false,
        statusMessage: "Deshaciendo ultimo cambio...",
      });
    } catch (error) {
      undoStack.push(lastChange);
      setImportMessage(toUserErrorMessage(error), "error");
      await loadRoster();
      renderRoster();
    }
  })();
});

window.addEventListener("resize", () => {
  renderRoster();
});

if ("serviceWorker" in navigator) {
  window.addEventListener("load", async () => {
    try {
      const registrations = await navigator.serviceWorker.getRegistrations();
      await Promise.all(registrations.map((registration) => registration.unregister()));
      if ("caches" in window) {
        const keys = await caches.keys();
        await Promise.all(keys.map((key) => caches.delete(key)));
      }
    } catch (_error) {
      // no-op
    }
  });
}

const storedFit = localStorage.getItem("roster.fitAllDays");
if (storedFit === "0") {
  state.fitAllDays = false;
  ui.fitDaysInput.checked = false;
}

const storedOnsite = localStorage.getItem("roster.onsiteCodes");
if (storedOnsite) {
  state.onsiteCode = storedOnsite;
  if (ui.onsiteCodesInput) ui.onsiteCodesInput.value = storedOnsite;
}

setActiveArea("planificacion");

renderCodeMeaning();

setInterval(() => {
  if (document.hidden) return;
  fullRefresh();
}, AUTO_REFRESH_INTERVAL_MS);

document.addEventListener("visibilitychange", () => {
  if (!document.hidden) {
    fullRefresh();
  }
});

window.addEventListener("focus", () => {
  fullRefresh();
});

fullRefresh();
