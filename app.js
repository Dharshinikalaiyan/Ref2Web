const workbookData = window.WORKBOOK_DATA;
const storageKey = "reference-operations-portal-v3";
const persisted = loadState();

const state = {
  activeKind: "all",
  search: "",
  denseMode: false,
  recentSheets: persisted.recentSheets || [],
};

const els = {
  globalSearch: document.getElementById("globalSearch"),
  kindFilters: document.getElementById("kindFilters"),
  sheetDirectory: document.getElementById("sheetDirectory"),
  heroTitle: document.getElementById("heroTitle"),
  heroDescription: document.getElementById("heroDescription"),
  tenantHighlights: document.getElementById("tenantHighlights"),
  globalHighlights: document.getElementById("globalHighlights"),
  sampleHighlights: document.getElementById("sampleHighlights"),
  recentSheets: document.getElementById("recentSheets"),
  toggleDenseMode: document.getElementById("toggleDenseMode"),
};

function loadState() {
  try {
    return JSON.parse(localStorage.getItem(storageKey)) || {};
  } catch {
    return {};
  }
}

function persistState() {
  localStorage.setItem(
    storageKey,
    JSON.stringify({
      recentSheets: state.recentSheets,
      overrides: persisted.overrides || {},
      uploads: persisted.uploads || {},
    }),
  );
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function labelize(value) {
  return value.replace(/(^|\s|-)\w/g, (match) => match.toUpperCase());
}

function describeKind(kind) {
  if (kind === "global") return "Shared standards";
  if (kind === "tenant") return "Customer references";
  if (kind === "sample") return "Sample extractions";
  if (kind === "placeholder") return "Reserved sheets";
  return "All sheets";
}

function describeSheet(sheet) {
  if (sheet.kind === "sample") return "Examples and extraction walkthroughs";
  if (sheet.kind === "global") return "Shared internal reference standards";
  if (sheet.kind === "placeholder") return "Sheet reserved in workbook";
  return "Tenant-specific methods, formats, and notes";
}

function loadKinds() {
  return ["all", ...new Set(workbookData.sheets.map((sheet) => sheet.kind))];
}

function countByKind(kind) {
  return kind === "all" ? workbookData.sheets.length : workbookData.sheets.filter((sheet) => sheet.kind === kind).length;
}

function getVisibleSheets() {
  return workbookData.sheets.filter((sheet) => {
    const kindMatch = state.activeKind === "all" || sheet.kind === state.activeKind;
    const searchMatch = !state.search || JSON.stringify(sheet).toLowerCase().includes(state.search.toLowerCase());
    return kindMatch && searchMatch;
  });
}

function groupVisibleSheets() {
  const groups = { global: [], tenant: [], sample: [], placeholder: [] };
  getVisibleSheets().forEach((sheet) => {
    if (!groups[sheet.kind]) groups[sheet.kind] = [];
    groups[sheet.kind].push(sheet);
  });
  return groups;
}

function activateSheet(slug) {
  state.recentSheets = [slug, ...state.recentSheets.filter((item) => item !== slug)].slice(0, 8);
  persistState();
  window.location.href = `./detail.html#${slug}`;
}

function renderKindFilters() {
  els.kindFilters.innerHTML = "";
  loadKinds().forEach((kind) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `pill ${state.activeKind === kind ? "active" : ""}`;
    button.textContent = labelize(kind);
    button.addEventListener("click", () => {
      state.activeKind = kind;
      renderAll();
    });
    els.kindFilters.appendChild(button);
  });
}

function renderDirectory() {
  renderKindFilters();
  document.body.classList.toggle("dense", state.denseMode);
  const groups = groupVisibleSheets();
  els.sheetDirectory.innerHTML = "";

  ["global", "tenant", "sample", "placeholder"].forEach((key) => {
    const sheets = groups[key];
    if (!sheets?.length) return;
    const group = document.createElement("section");
    group.className = "directory-group";
    group.innerHTML = `<div class="group-title">${describeKind(key)}</div>`;
    sheets.forEach((sheet) => {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "sheet-card";
      button.innerHTML = `<h4>${escapeHtml(sheet.title)}</h4><p>${escapeHtml(describeSheet(sheet))}</p>`;
      button.addEventListener("click", () => activateSheet(sheet.slug));
      group.appendChild(button);
    });
    els.sheetDirectory.appendChild(group);
  });
}

function renderCompactList(target, sheets, emptyText) {
  if (!sheets.length) {
    target.innerHTML = `<div class="empty-state">${emptyText}</div>`;
    return;
  }

  target.innerHTML = "";
  sheets.forEach((sheet) => {
    const item = document.createElement("button");
    item.type = "button";
    item.className = "compact-item";
    item.innerHTML = `<strong>${escapeHtml(sheet.title)}</strong><p>${escapeHtml(describeSheet(sheet))}</p>`;
    item.addEventListener("click", () => activateSheet(sheet.slug));
    target.appendChild(item);
  });
}

function renderDashboard() {
  renderCompactList(
    els.tenantHighlights,
    workbookData.sheets.filter((sheet) => sheet.kind === "tenant").slice(0, 10),
    "No tenant sheets found.",
  );
  renderCompactList(
    els.globalHighlights,
    workbookData.sheets.filter((sheet) => sheet.kind === "global"),
    "No shared references found.",
  );
  renderCompactList(
    els.sampleHighlights,
    workbookData.sheets.filter((sheet) => sheet.kind === "sample"),
    "No sample extraction sheets found.",
  );
  renderCompactList(
    els.recentSheets,
    state.recentSheets.map((slug) => workbookData.sheets.find((sheet) => sheet.slug === slug)).filter(Boolean),
    "Open a sheet and it will appear here.",
  );
}

function bindEvents() {
  els.globalSearch.addEventListener("input", (event) => {
    state.search = event.target.value;
    renderAll();
  });

  els.toggleDenseMode.addEventListener("click", () => {
    state.denseMode = !state.denseMode;
    els.toggleDenseMode.textContent = state.denseMode ? "Comfortable" : "Compact";
    renderDirectory();
  });
}

function renderAll() {
  els.heroTitle.textContent = "Reference Overview";
  els.heroDescription.textContent = "Choose a sheet from the directory or the quick lists below to open it.";
  renderDirectory();
  renderDashboard();
}

bindEvents();
renderAll();