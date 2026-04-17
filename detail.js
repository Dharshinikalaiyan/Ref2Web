const workbookData = window.WORKBOOK_DATA;
const storageKey = "reference-operations-portal-v3";
const persisted = loadState();

const state = {
  activeSheetSlug: getInitialSheetSlug(),
  sectionFilter: "all",
  editMode: false,
  overrides: persisted.overrides || {},
  uploads: persisted.uploads || {},
  recentSheets: persisted.recentSheets || [],
  modal: { sheetSlug: null, rowNumber: null, headers: [] },
  assign: { uploadId: null },
  noteEdit: { label: "" },
};

const els = {
  heroTitle: document.getElementById("heroTitle"),
  heroDescription: document.getElementById("heroDescription"),
  sheetTitle: document.getElementById("sheetTitle"),
  sheetPicker: document.getElementById("sheetPicker"),
  sectionFilter: document.getElementById("sectionFilter"),
  rawToggle: document.getElementById("rawToggle"),
  sheetNotes: document.getElementById("sheetNotes"),
  sheetBlocks: document.getElementById("sheetBlocks"),
  rawRowsContainer: document.getElementById("rawRowsContainer"),
  relatedSheets: document.getElementById("relatedSheets"),
  exampleUpload: document.getElementById("exampleUpload"),
  uploadList: document.getElementById("uploadList"),
  editModeToggle: document.getElementById("editModeToggle"),
  exportSnapshot: document.getElementById("exportSnapshot"),
  editModal: document.getElementById("editModal"),
  editForm: document.getElementById("editForm"),
  modalTitle: document.getElementById("modalTitle"),
  closeModal: document.getElementById("closeModal"),
  clearRowEdits: document.getElementById("clearRowEdits"),
  saveRowEdits: document.getElementById("saveRowEdits"),
  assignModal: document.getElementById("assignModal"),
  assignTitle: document.getElementById("assignTitle"),
  assignTable: document.getElementById("assignTable"),
  assignRow: document.getElementById("assignRow"),
  assignColumn: document.getElementById("assignColumn"),
  closeAssignModal: document.getElementById("closeAssignModal"),
  skipAssign: document.getElementById("skipAssign"),
  saveAssignment: document.getElementById("saveAssignment"),
  textModal: document.getElementById("textModal"),
  textModalTitle: document.getElementById("textModalTitle"),
  textModalBody: document.getElementById("textModalBody"),
  closeTextModal: document.getElementById("closeTextModal"),
  previewModal: document.getElementById("previewModal"),
  previewModalTitle: document.getElementById("previewModalTitle"),
  previewModalBody: document.getElementById("previewModalBody"),
  closePreviewModal: document.getElementById("closePreviewModal"),
  noteModal: document.getElementById("noteModal"),
  noteModalTitle: document.getElementById("noteModalTitle"),
  noteModalBodyInput: document.getElementById("noteModalBodyInput"),
  closeNoteModal: document.getElementById("closeNoteModal"),
  saveNoteEdits: document.getElementById("saveNoteEdits"),
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
      overrides: state.overrides,
      uploads: state.uploads,
      recentSheets: state.recentSheets,
    }),
  );
}

function getInitialSheetSlug() {
  const slug = window.location.hash.replace(/^#/, "");
  return workbookData.sheets.some((sheet) => sheet.slug === slug) ? slug : workbookData.sheets[0]?.slug || "";
}

function setHash(slug) {
  history.replaceState(null, "", `#${slug}`);
}

function getActiveSheet() {
  return workbookData.sheets.find((sheet) => sheet.slug === state.activeSheetSlug) || workbookData.sheets[0];
}

function rememberSheet(slug) {
  state.recentSheets = [slug, ...state.recentSheets.filter((item) => item !== slug)].slice(0, 8);
  persistState();
}

function getSheetOverride(slug) {
  return state.overrides[slug] || { cells: {}, notes: {} };
}

function getSheetNoteOverrides(slug) {
  const override = getSheetOverride(slug);
  if (!override.notes) override.notes = {};
  state.overrides[slug] = override;
  return override.notes;
}

function getDisplayCell(slug, rowNumber, colIndex, fallback) {
  return getSheetOverride(slug).cells?.[`${rowNumber}:${colIndex}`] ?? fallback;
}

function setCellOverride(slug, rowNumber, colIndex, value) {
  const override = getSheetOverride(slug);
  if (!override.cells) override.cells = {};
  override.cells[`${rowNumber}:${colIndex}`] = value;
  state.overrides[slug] = override;
}

function clearRowOverrides(slug, rowNumber) {
  const cells = getSheetOverride(slug).cells || {};
  Object.keys(cells).forEach((key) => {
    if (key.startsWith(`${rowNumber}:`)) delete cells[key];
  });
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function describeSheet(sheet) {
  if (sheet.kind === "sample") return "Examples and extraction walkthroughs";
  if (sheet.kind === "global") return "Shared internal reference standards";
  if (sheet.kind === "placeholder") return "Sheet reserved in workbook";
  return "Tenant-specific methods, formats, and notes";
}

function baseName(title) {
  return title
    .replace(/sample extractions/gi, "")
    .replace(/special cases/gi, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function getRelatedSheets(sheet) {
  const normalized = baseName(sheet.title);
  return workbookData.sheets.filter(
    (candidate) => candidate.slug !== sheet.slug && baseName(candidate.title) === normalized,
  );
}

function getSheetUploads(slug) {
  return state.uploads[slug] || [];
}

function getUploadById(uploadId) {
  return getSheetUploads(state.activeSheetSlug).find((item) => item.id === uploadId);
}

function renderHeader() {
  const sheet = getActiveSheet();
  els.heroTitle.textContent = sheet.title;
  els.heroDescription.textContent = `${describeSheet(sheet)} Edit mode works for both note cards and table rows.`;
  els.sheetTitle.textContent = sheet.title;
}

function renderSheetPicker() {
  els.sheetPicker.innerHTML = workbookData.sheets.map((sheet) => `<option value="${sheet.slug}">${escapeHtml(sheet.title)}</option>`).join("");
  els.sheetPicker.value = state.activeSheetSlug;
}

function renderSectionOptions() {
  const sheet = getActiveSheet();
  const options = new Set(["all"]);
  sheet.blocks.forEach((block) => block.rows.forEach((row) => row.values[0] && options.add(row.values[0])));
  els.sectionFilter.innerHTML = [...options]
    .map((value) => `<option value="${escapeHtml(value)}">${value === "all" ? "All sections" : escapeHtml(value)}</option>`)
    .join("");
  els.sectionFilter.value = state.sectionFilter;
}

function renderNotes() {
  const sheet = getActiveSheet();
  const noteOverrides = getSheetNoteOverrides(sheet.slug);
  els.sheetNotes.innerHTML = "";
  if (!sheet.notes?.length) return;

  const grouped = [];
  const catchAll = [];
  sheet.notes.forEach((note) => {
    const label = String(note.label || "").trim();
    if (!label || label.toLowerCase() === "none") {
      if (note.content) catchAll.push(note.content);
      return;
    }
    grouped.push({ label, content: note.content || "" });
  });

  const accountNotes = grouped.find((note) => note.label.toLowerCase() === "account notes:");
  if (catchAll.length) {
    if (accountNotes) accountNotes.content = [accountNotes.content, ...catchAll].filter(Boolean).join("\n");
    else grouped.unshift({ label: "Account Notes:", content: catchAll.join("\n") });
  }

  grouped.forEach((note) => {
    const article = document.createElement("article");
    article.className = `note-card ${state.editMode ? "editable-note" : ""}`;
    article.dataset.noteLabel = note.label;
    article.innerHTML = `<strong>${escapeHtml(note.label)}</strong><p>${escapeHtml(noteOverrides[note.label] ?? note.content ?? " ")}</p>`;
    els.sheetNotes.appendChild(article);
  });
}

function rowNonEmptyValues(row) {
  return row.values.filter((value) => String(value || "").trim() !== "");
}

function splitBlockIntoSegments(rows) {
  const segments = [];
  let currentTable = [];
  let index = 0;

  while (index < rows.length) {
    const row = rows[index];
    const nonEmpty = rowNonEmptyValues(row);
    const next = rows[index + 1];
    const nextNonEmpty = next ? rowNonEmptyValues(next) : [];
    const startsList = nonEmpty.length === 1 && next && nextNonEmpty.length === 1;

    if (startsList) {
      if (currentTable.length) {
        segments.push({ type: "table", rows: currentTable });
        currentTable = [];
      }
      const title = nonEmpty[0];
      const items = [];
      index += 1;
      while (index < rows.length && rowNonEmptyValues(rows[index]).length === 1) {
        items.push(rowNonEmptyValues(rows[index])[0]);
        index += 1;
      }
      segments.push({ type: "list", title, items });
      continue;
    }

    currentTable.push(row);
    index += 1;
  }

  if (currentTable.length) segments.push({ type: "table", rows: currentTable });
  return segments;
}

function compactCellText(value, maxLength = 130) {
  const text = String(value || "").replace(/\s+/g, " ").trim();
  if (text.length <= maxLength) return text;
  return `${text.slice(0, maxLength).trim()}...`;
}

function getUploadsForCell(blockIndex, rowNumber, colIndex) {
  return getSheetUploads(state.activeSheetSlug).filter(
    (file) =>
      file.assignment &&
      file.assignment.blockIndex === blockIndex &&
      file.assignment.rowNumber === rowNumber &&
      file.assignment.colIndex === colIndex,
  );
}

function renderExampleCell(existingValue, uploads) {
  const content = [];
  if (existingValue) content.push(`<div>${escapeHtml(existingValue)}</div>`);
  if (uploads.length) {
    content.push(
      `<div class="example-chip-wrap">${uploads
        .slice(0, 3)
        .map((file) => `<button type="button" class="example-chip" data-preview-upload="${file.id}">${escapeHtml(file.name)}</button>`)
        .join("")}${uploads.length > 3 ? `<span class="example-chip">+${uploads.length - 3} more</span>` : ""}</div>`,
    );
  }
  return content.join("");
}

function looksLikeExamplesColumn(header) {
  return /example|image/i.test(header || "");
}

function renderBlocks() {
  const sheet = getActiveSheet();
  els.sheetBlocks.innerHTML = "";

  if (!sheet.blocks.length) {
    els.sheetBlocks.innerHTML = `<div class="empty-state">This sheet is currently empty in the workbook.</div>`;
    return;
  }

  let shown = 0;
  sheet.blocks.forEach((block, index) => {
    const rows = block.rows.filter((row) => state.sectionFilter === "all" || row.values[0] === state.sectionFilter);
    if (!rows.length) return;

    splitBlockIntoSegments(rows).forEach((segment) => {
      shown += 1;
      if (segment.type === "list") {
        const listCard = document.createElement("article");
        listCard.className = "table-card ref-list-card";
        listCard.innerHTML = `<div class="table-head"><strong>${escapeHtml(segment.title)}</strong></div><div class="ref-list-wrap">${segment.items.map((item) => `<div class="ref-list-item">${escapeHtml(item)}</div>`).join("")}</div>`;
        els.sheetBlocks.appendChild(listCard);
        return;
      }

      const tableRows = segment.rows;
      const blockAssignments = getSheetUploads(sheet.slug).filter((file) => file.assignment && file.assignment.blockIndex === index);
      const hasAssignedExamples = blockAssignments.some((file) => tableRows.some((row) => row.rowNumber === file.assignment.rowNumber));
      const head = block.headers.map((header, idx) => `<th${idx === 0 ? ' class="section-heading-col"' : ""}>${escapeHtml(header)}</th>`);
      if (hasAssignedExamples) head.push('<th class="examples-col">Examples</th>');

      const body = tableRows
        .map((row) => {
          const cells = block.headers
            .map((header, colIndex) => {
              const value = getDisplayCell(sheet.slug, row.rowNumber, colIndex, row.values[colIndex] ?? "");
              const klass = colIndex === 0 ? "section-heading-col-cell" : "";
              if (String(value || "").length > 140) {
                return `<td class="expandable-cell ${klass}" data-cell-title="${escapeHtml(header || "Cell detail")}" data-cell-value="${escapeHtml(value)}">${escapeHtml(compactCellText(value))}</td>`;
              }
              return `<td class="${klass}">${escapeHtml(value)}</td>`;
            })
            .join("");

          const rowUploads = hasAssignedExamples ? blockAssignments.filter((file) => file.assignment.rowNumber === row.rowNumber) : [];
          const examplesCell = hasAssignedExamples ? `<td class="examples-col-cell">${renderExampleCell("", rowUploads)}</td>` : "";
          return `<tr class="${state.editMode ? "editable-row" : ""}" data-row-number="${row.rowNumber}">${cells}${examplesCell}</tr>`;
        })
        .join("");

      const card = document.createElement("article");
      card.className = "table-card";
      card.innerHTML = `<div class="table-head"><strong>Reference Table ${index + 1}</strong></div><div class="table-wrap"><table data-block-index="${index}"><thead><tr>${head.join("")}</tr></thead><tbody>${body}</tbody></table></div>`;
      els.sheetBlocks.appendChild(card);
    });
  });

  if (!shown) els.sheetBlocks.innerHTML = `<div class="empty-state">No rows match the selected section filter.</div>`;
}

function renderRawRows() {
  const sheet = getActiveSheet();
  if (!els.rawToggle.checked) {
    els.rawRowsContainer.classList.add("hidden");
    return;
  }
  els.rawRowsContainer.classList.remove("hidden");
  const maxCols = Math.max(...sheet.rawRows.map((row) => row.values.length), 0);
  const headers = [`<th class="table-row-number">Row</th>`];
  for (let i = 0; i < maxCols; i += 1) headers.push(`<th>Col ${i + 1}</th>`);
  const rows = sheet.rawRows.map((row) => `<tr><td class="table-row-number">${row.rowNumber}</td>${Array.from({ length: maxCols }, (_, idx) => `<td>${escapeHtml(row.values[idx] ?? "")}</td>`).join("")}</tr>`).join("");
  els.rawRowsContainer.innerHTML = `<article class="table-card"><div class="table-head"><strong>Raw Worksheet Rows</strong></div><div class="table-wrap"><table><thead><tr>${headers.join("")}</tr></thead><tbody>${rows}</tbody></table></div></article>`;
}

function buildAssignmentOptions() {
  const sheet = getActiveSheet();
  return sheet.blocks
    .map((block, blockIndex) => ({
      blockIndex,
      headers: block.headers,
      rows: block.rows.map((row) => ({
        rowNumber: row.rowNumber,
        label: `${row.values[0] || "No section"} - ${row.values[2] || row.values[1] || "Row " + row.rowNumber}`,
      })),
    }))
    .filter((block) => block.rows.length);
}

function openAssignModal(uploadId) {
  const file = getUploadById(uploadId);
  if (!file) return;
  state.assign.uploadId = uploadId;
  els.assignTitle.textContent = `Choose where "${file.name}" should appear`;
  const options = buildAssignmentOptions();
  els.assignTable.innerHTML = options.map((block) => `<option value="${block.blockIndex}">Reference Table ${Number(block.blockIndex) + 1}</option>`).join("");
  updateAssignRowsAndColumns();
  els.assignModal.classList.remove("hidden");
  els.assignModal.setAttribute("aria-hidden", "false");
}

function updateAssignRowsAndColumns() {
  const options = buildAssignmentOptions();
  const blockIndex = Number(els.assignTable.value || options[0]?.blockIndex || 0);
  const block = options.find((item) => item.blockIndex === blockIndex) || options[0];
  if (!block) return;
  els.assignRow.innerHTML = block.rows.map((row) => `<option value="${row.rowNumber}">${escapeHtml(row.label)}</option>`).join("");
  els.assignColumn.innerHTML = block.headers.map((header, idx) => `<option value="${idx}">${escapeHtml(header || `Column ${idx + 1}`)}</option>`).join("");
  const suggested = block.headers.findIndex((header) => looksLikeExamplesColumn(header));
  els.assignColumn.value = String(suggested >= 0 ? suggested : 0);
}

function closeAssignModal() {
  els.assignModal.classList.add("hidden");
  els.assignModal.setAttribute("aria-hidden", "true");
}

function saveAssignment() {
  const file = getUploadById(state.assign.uploadId);
  if (!file) return;
  file.assignment = {
    blockIndex: Number(els.assignTable.value),
    rowNumber: Number(els.assignRow.value),
    colIndex: Number(els.assignColumn.value),
  };
  persistState();
  closeAssignModal();
  renderBlocks();
  renderUploads();
}

function renderRelatedSheets() {
  const related = getRelatedSheets(getActiveSheet());
  if (!related.length) {
    els.relatedSheets.innerHTML = `<div class="empty-state">No directly related sheet group was found for this tab.</div>`;
    return;
  }
  els.relatedSheets.innerHTML = "";
  related.forEach((sheet) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "link-card";
    button.innerHTML = `<strong>${escapeHtml(sheet.title)}</strong><p>${escapeHtml(describeSheet(sheet))}</p>`;
    button.addEventListener("click", () => openSheet(sheet.slug));
    els.relatedSheets.appendChild(button);
  });
}

function renderUploads() {
  const uploads = getSheetUploads(state.activeSheetSlug);
  if (!uploads.length) {
    els.uploadList.innerHTML = `<div class="empty-state">No local attachments added for this sheet yet.</div>`;
    return;
  }

  els.uploadList.innerHTML = "";
  uploads.forEach((file) => {
    const card = document.createElement("div");
    card.className = "upload-card";
    card.innerHTML = `
      <strong>${escapeHtml(file.name)}</strong>
      <p>${escapeHtml(file.type || "Unknown type")} • ${formatBytes(file.size)}</p>
      <p>${file.assignment ? `Attached to row ${file.assignment.rowNumber}, column ${file.assignment.colIndex + 1}` : "Not yet attached to a field"}</p>
      <div class="badge-row">
        <button class="secondary-btn" type="button" data-preview-upload="${file.id}">Preview</button>
        <button class="secondary-btn" type="button" data-assign-upload="${file.id}">${file.assignment ? "Change Field" : "Assign Field"}</button>
        <button class="secondary-btn" type="button" data-remove-upload="${file.id}">Remove</button>
      </div>
    `;
    els.uploadList.appendChild(card);
  });

  els.uploadList.querySelectorAll("[data-preview-upload]").forEach((button) => {
    button.addEventListener("click", () => openPreviewModal(button.dataset.previewUpload));
  });
  els.uploadList.querySelectorAll("[data-assign-upload]").forEach((button) => {
    button.addEventListener("click", () => openAssignModal(button.dataset.assignUpload));
  });
  els.uploadList.querySelectorAll("[data-remove-upload]").forEach((button) => {
    button.addEventListener("click", () => {
      state.uploads[state.activeSheetSlug] = getSheetUploads(state.activeSheetSlug).filter((file) => file.id !== button.dataset.removeUpload);
      persistState();
      renderBlocks();
      renderUploads();
    });
  });
}

function formatBytes(bytes) {
  if (!bytes) return "0 B";
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

async function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function isTextLike(file) {
  return file.type.startsWith("text/") || /(json|csv|xml|javascript)/i.test(file.type) || /\.(txt|csv|json|md)$/i.test(file.name);
}

async function handleUpload(files) {
  if (!state.uploads[state.activeSheetSlug]) state.uploads[state.activeSheetSlug] = [];
  for (const file of files) {
    const entry = {
      id: `upload-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      name: file.name,
      type: file.type,
      size: file.size,
      previewText: "",
      imageDataUrl: "",
    };
    if (file.type.startsWith("image/") && file.size <= 900000) {
      entry.imageDataUrl = await fileToDataUrl(file).catch(() => "");
      if (!entry.imageDataUrl) entry.previewText = "Image uploaded. Inline preview could not be generated.";
    } else if (isTextLike(file)) {
      entry.previewText = await file.slice(0, 2000).text().then((text) => text.slice(0, 500)).catch(() => "");
    } else {
      entry.previewText = `Uploaded file: ${file.name}. Preview is not available for this file type, but it is linked to this sheet.`;
    }
    state.uploads[state.activeSheetSlug].push(entry);
  }
  persistState();
  renderBlocks();
  renderUploads();
  const lastUpload = state.uploads[state.activeSheetSlug].at(-1);
  if (lastUpload?.id) openAssignModal(lastUpload.id);
}

function findRowContext(rowNumber) {
  const sheet = getActiveSheet();
  for (const block of sheet.blocks) {
    const row = block.rows.find((entry) => entry.rowNumber === rowNumber);
    if (row) return { sheet, row, headers: block.headers };
  }
  return null;
}

function openEditModal(rowNumber) {
  const context = findRowContext(rowNumber);
  if (!context) return;
  state.modal = { sheetSlug: context.sheet.slug, rowNumber, headers: context.headers };
  els.modalTitle.textContent = `${context.sheet.title} • Row ${rowNumber}`;
  els.editForm.innerHTML = context.headers
    .map((header, idx) => {
      const value = getDisplayCell(context.sheet.slug, rowNumber, idx, context.row.values[idx] ?? "");
      return `<div class="modal-field"><label for="modal-field-${idx}">${escapeHtml(header || `Column ${idx + 1}`)}</label><textarea id="modal-field-${idx}" data-modal-col="${idx}">${escapeHtml(value)}</textarea></div>`;
    })
    .join("");
  els.editModal.classList.remove("hidden");
  els.editModal.setAttribute("aria-hidden", "false");
}

function closeModal() {
  els.editModal.classList.add("hidden");
  els.editModal.setAttribute("aria-hidden", "true");
}

function saveModalEdits() {
  els.editForm.querySelectorAll("[data-modal-col]").forEach((field) => {
    setCellOverride(state.modal.sheetSlug, state.modal.rowNumber, Number(field.dataset.modalCol), field.value.trim());
  });
  persistState();
  closeModal();
  renderBlocks();
}

function resetModalRow() {
  clearRowOverrides(state.modal.sheetSlug, state.modal.rowNumber);
  persistState();
  closeModal();
  renderBlocks();
}

function openTextModal(title, value) {
  els.textModalTitle.textContent = title || "Cell detail";
  els.textModalBody.textContent = value || "";
  els.textModal.classList.remove("hidden");
  els.textModal.setAttribute("aria-hidden", "false");
}

function closeTextModal() {
  els.textModal.classList.add("hidden");
  els.textModal.setAttribute("aria-hidden", "true");
}

function openPreviewModal(uploadId) {
  const file = getUploadById(uploadId);
  if (!file) return;
  els.previewModalTitle.textContent = file.name;
  let html = `<p>${escapeHtml(file.type || "Unknown type")} • ${formatBytes(file.size)}</p>`;
  if (file.imageDataUrl) html += `<img src="${file.imageDataUrl}" alt="${escapeHtml(file.name)}" />`;
  else if (file.previewText) html += `<div class="preview-code">${escapeHtml(file.previewText)}</div>`;
  else html += `<p>Preview unavailable for this file type.</p>`;
  els.previewModalBody.innerHTML = html;
  els.previewModal.classList.remove("hidden");
  els.previewModal.setAttribute("aria-hidden", "false");
}

function closePreviewModal() {
  els.previewModal.classList.add("hidden");
  els.previewModal.setAttribute("aria-hidden", "true");
}

function openNoteModal(label) {
  const source = getActiveSheet().notes.find((note) => note.label === label);
  els.noteModalTitle.textContent = label;
  els.noteModalBodyInput.value = getSheetNoteOverrides(state.activeSheetSlug)[label] ?? source?.content ?? "";
  state.noteEdit.label = label;
  els.noteModal.classList.remove("hidden");
  els.noteModal.setAttribute("aria-hidden", "false");
}

function closeNoteModal() {
  els.noteModal.classList.add("hidden");
  els.noteModal.setAttribute("aria-hidden", "true");
}

function saveNoteEdits() {
  if (!state.noteEdit.label) return;
  getSheetNoteOverrides(state.activeSheetSlug)[state.noteEdit.label] = els.noteModalBodyInput.value.trim();
  persistState();
  closeNoteModal();
  renderNotes();
}

function exportSnapshot() {
  const blob = new Blob(
    [JSON.stringify({ exportedAt: new Date().toISOString(), activeSheet: getActiveSheet().title, overrides: state.overrides, uploads: state.uploads }, null, 2)],
    { type: "application/json" },
  );
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${getActiveSheet().slug}-sheet-snapshot.json`;
  a.click();
  URL.revokeObjectURL(url);
}

function openSheet(slug) {
  state.activeSheetSlug = slug;
  state.sectionFilter = "all";
  rememberSheet(slug);
  setHash(slug);
  renderAll();
}

function bindEvents() {
  els.sheetPicker.addEventListener("change", (event) => openSheet(event.target.value));
  els.sectionFilter.addEventListener("change", (event) => {
    state.sectionFilter = event.target.value;
    renderBlocks();
  });
  els.rawToggle.addEventListener("change", renderRawRows);
  els.editModeToggle.addEventListener("click", () => {
    state.editMode = !state.editMode;
    els.editModeToggle.textContent = state.editMode ? "Disable Edit Mode" : "Enable Edit Mode";
    renderNotes();
    renderBlocks();
  });
  els.exportSnapshot.addEventListener("click", exportSnapshot);
  els.exampleUpload.addEventListener("change", async (event) => {
    const files = [...event.target.files];
    if (!files.length) return;
    await handleUpload(files);
    event.target.value = "";
  });
  els.sheetBlocks.addEventListener("click", (event) => {
    const uploadChip = event.target.closest("[data-preview-upload]");
    if (uploadChip) {
      openPreviewModal(uploadChip.dataset.previewUpload);
      return;
    }
    const expandable = event.target.closest(".expandable-cell");
    if (expandable) {
      openTextModal(expandable.dataset.cellTitle, expandable.dataset.cellValue);
      return;
    }
    if (!state.editMode) return;
    const row = event.target.closest("tr[data-row-number]");
    if (row) openEditModal(Number(row.dataset.rowNumber));
  });
  els.sheetNotes.addEventListener("click", (event) => {
    if (!state.editMode) return;
    const note = event.target.closest("[data-note-label]");
    if (note) openNoteModal(note.dataset.noteLabel);
  });
  els.closeModal.addEventListener("click", closeModal);
  els.saveRowEdits.addEventListener("click", saveModalEdits);
  els.clearRowEdits.addEventListener("click", resetModalRow);
  els.assignTable.addEventListener("change", updateAssignRowsAndColumns);
  els.closeAssignModal.addEventListener("click", closeAssignModal);
  els.skipAssign.addEventListener("click", closeAssignModal);
  els.saveAssignment.addEventListener("click", saveAssignment);
  els.closeTextModal.addEventListener("click", closeTextModal);
  els.closePreviewModal.addEventListener("click", closePreviewModal);
  els.closeNoteModal.addEventListener("click", closeNoteModal);
  els.saveNoteEdits.addEventListener("click", saveNoteEdits);
  els.editModal.addEventListener("click", (event) => {
    if (event.target.dataset.closeModal === "true") closeModal();
  });
  els.assignModal.addEventListener("click", (event) => {
    if (event.target.dataset.closeAssignModal === "true") closeAssignModal();
  });
  els.textModal.addEventListener("click", (event) => {
    if (event.target.dataset.closeTextModal === "true") closeTextModal();
  });
  els.previewModal.addEventListener("click", (event) => {
    if (event.target.dataset.closePreviewModal === "true") closePreviewModal();
  });
  els.noteModal.addEventListener("click", (event) => {
    if (event.target.dataset.closeNoteModal === "true") closeNoteModal();
  });
  window.addEventListener("hashchange", () => {
    const slug = window.location.hash.replace(/^#/, "");
    if (workbookData.sheets.some((sheet) => sheet.slug === slug)) openSheet(slug);
  });
}

function renderAll() {
  renderHeader();
  renderSheetPicker();
  renderSectionOptions();
  renderNotes();
  renderBlocks();
  renderRawRows();
  renderRelatedSheets();
  renderUploads();
}

rememberSheet(state.activeSheetSlug);
bindEvents();
renderAll();