async function loadData() {
  const response = await fetch("./data/projects.json");
  if (!response.ok) {
    throw new Error("Failed to load projects.json");
  }
  return response.json();
}

function toProgress(project) {
  const keys = Object.keys(project).map(k => k.toLowerCase());
  const progressKey = keys.find(k => k.includes("progress") || k.includes("completion"));
  if (!progressKey) return 0;

  const originalKey = Object.keys(project).find(k => k.toLowerCase() === progressKey);
  const value = (project[originalKey] || "").toString().replace("%", "");
  const parsed = Number.parseFloat(value);
  return Number.isFinite(parsed) ? Math.max(0, Math.min(parsed, 100)) : 0;
}

function toProjectName(project, index) {
  const keys = Object.keys(project);
  const nameKey = keys.find(k => k.toLowerCase().includes("project"));
  return nameKey ? project[nameKey] || `Project ${index + 1}` : `Project ${index + 1}`;
}

function renderKpis(projects) {
  const total = projects.length;
  const progressValues = projects.map(toProgress);
  const avg = total ? progressValues.reduce((a, b) => a + b, 0) / total : 0;
  const completed = progressValues.filter(v => v >= 100).length;

  document.getElementById("totalProjects").textContent = total;
  document.getElementById("avgProgress").textContent = `${avg.toFixed(1)}%`;
  document.getElementById("completedCount").textContent = completed;
}

function renderBars(projects) {
  const chart = document.getElementById("progressChart");
  chart.innerHTML = "";

  projects.forEach((project, index) => {
    const progress = toProgress(project);
    const name = toProjectName(project, index);

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

(async function init() {
  try {
    const payload = await loadData();
    const projects = payload.projects || [];
    renderKpis(projects);
    renderBars(projects);
    renderTable(projects);
  } catch (error) {
    document.body.insertAdjacentHTML(
      "beforeend",
      `<p style="color:red;padding:12px;">${error.message}</p>`
    );
  }
})();
