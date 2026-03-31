async function loadData() {
  const response = await fetch("./data/projects.json");
  if (!response.ok) {
    throw new Error("Failed to load projects.json");
  }
  return response.json();
}

const tableState = {
  allRows: [],
  filteredRows: [],
  currentPage: 1,
  pageSize: 25
};

function getValueByKeyContains(record, keyword) {
  const keys = Object.keys(record);
  const key = keys.find((k) => k.toLowerCase().includes(keyword));
  return key ? record[key] : "";
}

function parseDate(value) {
  if (!value) return null;
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? null : date;
}

function estimateProgressFromDates(record) {
  const plannedStart = parseDate(getValueByKeyContains(record, "planned start"));
  const plannedEnd = parseDate(getValueByKeyContains(record, "planned end"));
  if (!plannedStart || !plannedEnd) return 0;

  const now = new Date();
  const total = plannedEnd.getTime() - plannedStart.getTime();
  if (total <= 0) return 0;
  const elapsed = Math.max(0, Math.min(now.getTime() - plannedStart.getTime(), total));
  return (elapsed / total) * 100;
}

function toProgress(record) {
  const keys = Object.keys(record);
  const progressKey = keys.find((k) => {
    const n = k.toLowerCase();
    return n.includes("progress") || n.includes("completion");
  });

  if (progressKey) {
    const value = (record[progressKey] || "").toString().replace("%", "").trim();
    const parsed = Number.parseFloat(value);
    if (Number.isFinite(parsed)) return Math.max(0, Math.min(parsed, 100));
  }

  const status = (getValueByKeyContains(record, "status") || "").toString().toLowerCase();
  if (status.includes("complete")) return 100;
  if (status.includes("approved")) return 65;
  if (status.includes("planning")) return 20;

  return estimateProgressFromDates(record);
}

function toProjectName(record, index) {
  const keys = Object.keys(record);
  const nameKey = keys.find((k) => {
    const n = k.toLowerCase();
    return n.includes("project title") || n.includes("project name") || n.includes("project");
  });
  return nameKey ? record[nameKey] || `Record ${index + 1}` : `Record ${index + 1}`;
}

function renderKpis(records) {
  const total = records.length;
  const progressValues = records.map(toProgress);
  const avg = total ? progressValues.reduce((a, b) => a + b, 0) / total : 0;
  const completed = progressValues.filter(v => v >= 100).length;

  document.getElementById("totalProjects").textContent = total;
  document.getElementById("avgProgress").textContent = `${avg.toFixed(1)}%`;
  document.getElementById("completedCount").textContent = completed;
}

function renderBars(records) {
  const chart = document.getElementById("progressChart");
  chart.innerHTML = "";

  records.forEach((record, index) => {
    const progress = toProgress(record);
    const name = toProjectName(record, index);

    const row = document.createElement("div");
    row.className = "bar-row";
    row.innerHTML = `
      <div>${name}</div>
      <div class="bar-track"><div class="bar-fill" style="width:${progress}%"></div></div>
      <div>${progress.toFixed(0)}%</div>
    `;
    chart.appendChild(row);
  });
}

function renderTable(projects) {
  const table = document.getElementById("projectsTable");
  table.innerHTML = "";
  if (!projects.length) return;

  const headers = Object.keys(projects[0]);
  const thead = document.createElement("thead");
  thead.innerHTML = `<tr>${headers.map(h => `<th>${h}</th>`).join("")}</tr>`;

  const tbody = document.createElement("tbody");
  projects.forEach((project) => {
    const tr = document.createElement("tr");
    tr.innerHTML = headers.map(h => `<td>${project[h] ?? ""}</td>`).join("");
    tbody.appendChild(tr);
  });

  table.appendChild(thead);
  table.appendChild(tbody);
}

function filterRows(rows, query) {
  const q = query.trim().toLowerCase();
  if (!q) return rows;
  return rows.filter((row) =>
    Object.values(row).some((value) => (value ?? "").toString().toLowerCase().includes(q))
  );
}

function renderPagedTable() {
  const total = tableState.filteredRows.length;
  const totalPages = Math.max(1, Math.ceil(total / tableState.pageSize));
  tableState.currentPage = Math.min(tableState.currentPage, totalPages);

  const startIndex = (tableState.currentPage - 1) * tableState.pageSize;
  const pageRows = tableState.filteredRows.slice(startIndex, startIndex + tableState.pageSize);
  renderTable(pageRows);

  const from = total === 0 ? 0 : startIndex + 1;
  const to = Math.min(startIndex + tableState.pageSize, total);
  document.getElementById("tableSummary").textContent = `${from}-${to} of ${total} rows`;
  document.getElementById("pageIndicator").textContent = `Page ${tableState.currentPage} of ${totalPages}`;
  document.getElementById("prevPage").disabled = tableState.currentPage <= 1;
  document.getElementById("nextPage").disabled = tableState.currentPage >= totalPages;
}

function setupTableControls() {
  const searchInput = document.getElementById("searchInput");
  const pageSizeSelect = document.getElementById("pageSizeSelect");

  searchInput.addEventListener("input", () => {
    tableState.currentPage = 1;
    tableState.filteredRows = filterRows(tableState.allRows, searchInput.value);
    renderPagedTable();
  });

  pageSizeSelect.addEventListener("change", () => {
    tableState.pageSize = Number(pageSizeSelect.value) || 25;
    tableState.currentPage = 1;
    renderPagedTable();
  });

  document.getElementById("prevPage").addEventListener("click", () => {
    tableState.currentPage = Math.max(1, tableState.currentPage - 1);
    renderPagedTable();
  });

  document.getElementById("nextPage").addEventListener("click", () => {
    tableState.currentPage += 1;
    renderPagedTable();
  });
}

function setupSheets(payload) {
  const select = document.getElementById("sheetSelect");
  const sheets = payload.sheets || [];
  const fallbackRows = payload.projects || [];
  const searchInput = document.getElementById("searchInput");

  const normalizedSheets = sheets.length
    ? sheets
    : [{ name: "Sheet1", rows: fallbackRows }];

  select.innerHTML = normalizedSheets
    .map((sheet, i) => `<option value="${i}">${sheet.name} (${(sheet.rows || []).length})</option>`)
    .join("");

  const renderSheet = (index) => {
    const selected = normalizedSheets[index] || normalizedSheets[0];
    const rows = selected.rows || [];
    renderKpis(rows);
    renderBars(rows);
    tableState.allRows = rows;
    tableState.currentPage = 1;
    tableState.filteredRows = filterRows(rows, searchInput.value);
    renderPagedTable();
  };

  select.addEventListener("change", () => renderSheet(Number(select.value)));
  renderSheet(0);
}

(async function init() {
  try {
    const payload = await loadData();
    const generatedAt = payload.generatedAt ? new Date(payload.generatedAt) : null;
    const lastUpdated = generatedAt && !Number.isNaN(generatedAt.getTime())
      ? generatedAt.toLocaleString()
      : "Unknown";
    document.getElementById("lastUpdated").textContent = `Last updated: ${lastUpdated}`;
    setupTableControls();
    setupSheets(payload);
  } catch (error) {
    document.body.insertAdjacentHTML(
      "beforeend",
      `<p style="color:red;padding:12px;">${error.message}</p>`
    );
  }
})();
