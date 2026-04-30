const appState = {
  hasApiKey: false,
  selectedFiles: [],
  filePreviewUrls: new Map(),
  sheetInfo: null,
  share: null,
  isProcessing: false,
  animationTimer: null,
  previewUrl: null,
  phoneMode: false,
  phoneAutoInspectStarted: false,
};

const SUPPORTED_UPLOAD_EXTENSIONS = new Set([".pdf", ".jpg", ".jpeg", ".png", ".webp", ".heic", ".heif"]);
const GEMINI_MODEL = "gemini-3.1-flash-lite-preview";
const IMAGE_OPTIMIZE_MAX_DIMENSION = 1800;
const IMAGE_OPTIMIZE_QUALITY = 0.86;
const IMAGE_OPTIMIZE_MIN_BYTES = 900 * 1024;
const OPTIMIZABLE_IMAGE_TYPES = new Set(["image/jpeg", "image/png", "image/webp"]);
const SUPPORTED_UPLOAD_TYPES = new Set([
  "application/pdf",
  "image/jpeg",
  "image/png",
  "image/webp",
  "image/heic",
  "image/heif",
]);

const $ = (id) => {
  const el = document.getElementById(id);
  if (el) return el;
  // Safe Fallback: return a dummy element to prevent crashes
  console.warn(`Element with ID "${id}" was not found. Creating virtual fallback.`);
  return document.createElement("div");
};

let els = {};

function initEls() {
  els = {
    apiStatus: $("apiStatus"),
    settingsButton: $("settingsButton"),
    closeSettingsButton: $("closeSettingsButton"),
    settingsDialog: $("settingsDialog"),
    apiKey: $("apiKey"),
    modelName: $("modelName"),
    saveSettingsButton: $("saveSettingsButton"),
    excelPath: $("excelPath"),
    sheetSelect: $("sheetSelect"),
    connectExcelButton: $("connectExcelButton"),
    inspectButton: $("inspectButton"),
    openExcelButton: $("openExcelButton"),
    sheetSelectionArea: $("sheetSelectionArea"),
    sheetSummary: $("sheetSummary"),
    recentExcelPaths: $("recentExcelPaths"),
    recentSheets: $("recentSheets"),
    dropzone: $("dropzone"),
    fileInput: $("fileInput"),
    cameraInput: $("cameraInput"),
    browseFileButton: document.getElementById("browseFileButton"),
    cameraButton: $("cameraButton"),
    fileBadge: $("fileBadge"),
    dropTitle: $("dropTitle"),
    dropMeta: $("dropMeta"),
    runButton: $("runButton"),
    runHint: $("runHint"),
    phoneStatus: $("phoneStatus"),
    uploadTitle: $("uploadTitle"),
    dropEmptyState: $("dropEmptyState"),
    fileGrid: $("fileGrid"),
    fileItems: $("fileItems"),
    processState: $("processState"),
    jsonConsole: $("jsonConsole"),
    resultPanel: $("resultPanel"),
    resultTitle: $("resultTitle"),
    resultTable: $("resultTable"),
    historyList: $("historyList"),
    shareButton: $("shareButton"),
    sharePopover: $("sharePopover"),
    shareUrl: $("shareUrl"),
    copyShareButton: $("copyShareButton"),
    refreshShareButton: $("refreshShareButton"),
    whatsappLink: $("whatsappLink"),
    emailLink: $("emailLink"),
    hostNote: $("hostNote"),
    successOverlay: $("successOverlay"),
    successTitle: $("successTitle"),
    successText: $("successText"),
    extractionStage: $("extractionStage"),
    documentPreview: $("documentPreview"),
    previewImage: $("previewImage"),
    previewFallback: $("previewFallback"),
    dataFlightLayer: $("dataFlightLayer"),
    excelBoard: $("excelBoard"),
    excelBoardStatus: $("excelBoardStatus"),
    excelGridPreview: $("excelGridPreview"),
    commandCard: document.querySelector(".command-card"),
  };
}

function detectPhoneMode() {
  const params = new URLSearchParams(window.location.search);
  const forced = params.get("phone") === "1" || params.get("mobile") === "1";
  const likelyPhone = /Android|iPhone|iPad|iPod|Mobile/i.test(navigator.userAgent)
    && window.matchMedia("(max-width: 820px)").matches;
  return forced || likelyPhone;
}

function updatePhoneStatus(message, kind = "info") {
  if (!els.phoneStatus || !els.phoneStatus.isConnected) return;
  els.phoneStatus.textContent = message;
  els.phoneStatus.dataset.kind = kind;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function fileSize(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function fileExtension(name) {
  const index = String(name || "").lastIndexOf(".");
  return index >= 0 ? String(name).slice(index).toLowerCase() : "";
}

function isSupportedUpload(file) {
  return SUPPORTED_UPLOAD_TYPES.has(file.type) || SUPPORTED_UPLOAD_EXTENSIONS.has(fileExtension(file.name));
}

function compactPath(path) {
  const text = String(path || "");
  if (text.length <= 54) return text;
  return `${text.slice(0, 24)}...${text.slice(-24)}`;
}

function imageBitmapToBlob(canvas, type, quality) {
  return new Promise((resolve) => {
    canvas.toBlob((blob) => resolve(blob), type, quality);
  });
}

function optimizedImageName(name) {
  const base = String(name || "invoice").replace(/\.[^.]+$/, "");
  return `${base}-optimized.jpg`;
}

async function optimizeUploadFile(file) {
  if (!OPTIMIZABLE_IMAGE_TYPES.has(file.type)) return file;

  let bitmap = null;
  try {
    bitmap = await createImageBitmap(file);
    const longestSide = Math.max(bitmap.width, bitmap.height);
    const shouldResize = longestSide > IMAGE_OPTIMIZE_MAX_DIMENSION;
    const shouldCompress = file.size > IMAGE_OPTIMIZE_MIN_BYTES || file.type !== "image/jpeg";
    if (!shouldResize && !shouldCompress) return file;

    const scale = shouldResize ? IMAGE_OPTIMIZE_MAX_DIMENSION / longestSide : 1;
    const width = Math.max(1, Math.round(bitmap.width * scale));
    const height = Math.max(1, Math.round(bitmap.height * scale));
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const context = canvas.getContext("2d", {alpha: false});
    if (!context) return file;

    context.fillStyle = "#fff";
    context.fillRect(0, 0, width, height);
    context.drawImage(bitmap, 0, 0, width, height);

    const blob = await imageBitmapToBlob(canvas, "image/jpeg", IMAGE_OPTIMIZE_QUALITY);
    if (!blob || blob.size >= file.size) return file;

    return new File([blob], optimizedImageName(file.name), {
      type: "image/jpeg",
      lastModified: file.lastModified,
    });
  } catch (error) {
    console.warn("Image optimization skipped:", error);
    return file;
  } finally {
    if (bitmap?.close) bitmap.close();
  }
}

async function apiJson(url, options = {}) {
  const response = await fetch(url, options);
  const payload = await response.json().catch(() => ({}));
  if (!response.ok || payload.ok === false) {
    throw new Error(payload.error || `Request failed: ${response.status}`);
  }
  return payload;
}

async function loadState() {
  const data = await apiJson("/api/state");
  renderState(data);
}

function renderState(data) {
  const config = data.config || {};
  appState.hasApiKey = Boolean(config.hasApiKey);
  appState.share = data.share || null;

  const statusLabel = appState.hasApiKey ? "Connected" : "Key Missing";
  els.apiStatus.innerHTML = `<span class="status-dot"></span> ${statusLabel}`;
  els.apiStatus.classList.toggle("is-ok", appState.hasApiKey);
  els.apiStatus.classList.toggle("is-warn", !appState.hasApiKey);
  els.modelName.value = config.geminiModel || GEMINI_MODEL;

  if (!els.excelPath.value && config.defaultExcelPath) {
    els.excelPath.value = config.defaultExcelPath;
  }

  renderRecent(config.recentSheets || []);
  renderExcelAutocomplete(config.recentSheets || []);
  renderHistory(data.history || []);
  renderShare(data.share || {});

  if (appState.phoneMode) {
    preparePhoneMode(config);
  }

  updateRunState();
}

function renderShare(share) {
  const shareLink = share.phoneUrl || share.lanUrl || share.currentUrl || "No link";
  els.shareUrl.textContent = shareLink;
  els.whatsappLink.href = share.whatsappUrl || "#";
  els.emailLink.href = share.emailUrl || "#";

  const localOnly = ["127.0.0.1", "localhost"].includes(window.location.hostname);
  els.hostNote.textContent = localOnly
    ? "For phone uploads, start with START_HOST_SHARE.bat."
    : "This link is ready for phones on the same network.";
}

function preparePhoneMode(config) {
  if (els.browseFileButton?.isConnected) {
    els.browseFileButton.textContent = "Upload file";
  }
  if (els.cameraButton?.isConnected) {
    els.cameraButton.setAttribute("aria-label", "Take photo");
  }
  els.dropTitle.textContent = appState.selectedFiles.length ? `${appState.selectedFiles.length} files` : "Upload invoice";
  els.dropMeta.textContent = appState.selectedFiles.length
    ? `${fileSize(appState.selectedFiles.reduce((acc, f) => acc + f.size, 0))} total`
    : "Photo, image, or PDF";

  if (!appState.hasApiKey) {
    updatePhoneStatus("Add API key on the main computer.", "warn");
    return;
  }

  if (!config.defaultExcelPath) {
    updatePhoneStatus("Connect the Excel sheet on the main computer.", "warn");
    return;
  }

  if (!appState.sheetInfo && !appState.phoneAutoInspectStarted) {
    appState.phoneAutoInspectStarted = true;
    els.excelPath.value = config.defaultExcelPath;
    updatePhoneStatus("Connecting to the office Excel sheet...");
    window.setTimeout(() => inspectSheet(config.defaultSheet || ""), 0);
    return;
  }

  if (appState.sheetInfo && !appState.isProcessing && appState.selectedFiles.length === 0) {
    updatePhoneStatus("Ready. Upload a file or take a photo.");
  }
}

function renderExcelAutocomplete(items) {
  const paths = [...new Set(items.map(item => item.path))];
  els.recentExcelPaths.innerHTML = paths.map(path => `<option value="${escapeHtml(path)}">`).join("");
}

function renderRecent(items) {
  if (!items.length) {
    els.recentSheets.innerHTML = `<div class="history-item empty-state"><strong>No sheets yet</strong><span>Connected workbooks appear here.</span></div>`;
    return;
  }

  els.recentSheets.innerHTML = items.map((item, index) => `
    <button class="recent-chip" type="button" data-index="${index}" title="${escapeHtml(item.path || "")}">
      <span class="recent-chip-top">
        <strong>${escapeHtml(item.sheet || "Sheet")}</strong>
        <span>${escapeHtml(item.lastUsed || "Recent")}</span>
      </span>
      <small>${escapeHtml(compactPath(item.path || ""))}</small>
    </button>
  `).join("");

  els.recentSheets.querySelectorAll("button").forEach((button, index) => {
    button.addEventListener("click", () => {
      const item = items[index];
      els.excelPath.value = item.path || "";
      inspectSheet(item.sheet || "");
    });
  });
}

function renderHistory(items) {
  if (!items.length) {
    els.historyList.innerHTML = `<div class="history-item empty-state"><strong>No changes yet</strong><span>Rows added will appear here.</span></div>`;
    return;
  }

  els.historyList.innerHTML = items.map((item) => {
    const cells = Object.keys(item.changedCells || {}).length;
    return `
      <div class="history-item">
        <div class="history-item-top">
          <strong>Row ${escapeHtml(item.rowNumber)} in ${escapeHtml(item.sheet)}</strong>
          <span>${cells} cells</span>
        </div>
        <span class="history-file" title="${escapeHtml(item.fileName)}">${escapeHtml(item.fileName)}</span>
        <span class="history-path" title="${escapeHtml(item.excelPath)}">${escapeHtml(compactPath(item.excelPath))}</span>
        <span class="history-time">${escapeHtml(item.time)}</span>
      </div>
    `;
  }).join("");
}

function renderSheetSummary(info) {
  appState.sheetInfo = info;
  const headers = (info.validHeaders || []).slice(0, 18);
  const extra = Math.max((info.validHeaders || []).length - headers.length, 0);
  const tags = headers.map((header) => `<span>${escapeHtml(header)}</span>`).join("");
  const more = extra ? `<span>+${extra} more</span>` : "";

  els.sheetSummary.innerHTML = `
    <strong>${escapeHtml(info.fileName)}</strong>
    <div>${escapeHtml(info.sheet)} - ${escapeHtml(info.rowCount)} rows - ${escapeHtml(info.validHeaders.length)} columns</div>
  `;
  els.sheetSummary.style.display = "block";
  els.sheetSelectionArea.style.display = "block";

  els.sheetSelect.innerHTML = (info.sheetNames || [])
    .map((sheet) => `<option value="${escapeHtml(sheet)}">${escapeHtml(sheet)}</option>`)
    .join("");
  els.sheetSelect.value = info.sheet;
  els.sheetSelect.disabled = false;
  els.openExcelButton.disabled = false;
  setStep("sheet", "ready");
  renderExcelPreview(info.validHeaders || []);

  if (appState.phoneMode) {
    updatePhoneStatus(appState.selectedFiles.length ? "Files ready. Sending to Excel..." : "Ready. Upload a file or take a photo.");
    maybeAutoRunPhone();
  }
  updateRunState();
}

function renderSheetError(message) {
  appState.sheetInfo = null;
  els.sheetSummary.innerHTML = `<span class="error-text">${escapeHtml(message)}</span>`;
  els.sheetSummary.style.display = "block";
  els.sheetSelect.disabled = true;
  els.sheetSelect.innerHTML = `<option value="">Inspect first</option>`;
  els.openExcelButton.disabled = false;
  renderExcelPreview();
  if (appState.phoneMode) {
    updatePhoneStatus(`Office setup issue: ${message}`, "warn");
  }
  updateRunState();
}

function setBusy(button, busyText, isBusy) {
  if (!button) return;
  if (isBusy) {
    button.dataset.originalText = button.textContent;
    button.textContent = busyText;
    button.disabled = true;
  } else {
    button.textContent = button.dataset.originalText || button.textContent;
    button.disabled = false;
  }
}

async function browseExcel() {
  console.log("DEBUG: Browse button clicked.");
  setBusy(els.connectExcelButton, "...", true);
  els.sheetSummary.textContent = "Opening system file picker...";
  try {
    const data = await apiJson("/api/browse-excel", {method: "POST"});
    if (!data.selected) {
      els.sheetSummary.innerHTML = `<span class="summary-placeholder">No workbook selected.</span>`;
      return;
    }
    els.excelPath.value = data.path || data.sheet?.path || "";
    renderSheetSummary(data.sheet);
    await loadState();
  } catch (error) {
    renderSheetError(error.message);
  } finally {
    setBusy(els.connectExcelButton, "Browse", false);
  }
}

async function inspectSheet(forcedSheet = "") {
  const path = els.excelPath.value.trim();
  if (!path) {
    renderSheetError("Choose or enter the workbook path first.");
    return;
  }

  const requestedSheet = forcedSheet || (els.sheetSelect.disabled ? "" : els.sheetSelect.value);
  setBusy(els.inspectButton, "Inspecting", true);
  els.sheetSummary.textContent = "Reading headers and recent rows...";

  try {
    const data = await apiJson("/api/inspect-excel", {
      method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify({path, sheet: requestedSheet}),
    });
    renderSheetSummary(data.sheet);
    await loadState();
  } catch (error) {
    renderSheetError(error.message);
  } finally {
    setBusy(els.inspectButton, "Inspect", false);
  }
}

function setSelectedFiles(files) {
  if (!files || files.length === 0) return;
  const incomingFiles = Array.from(files);
  const newFiles = incomingFiles.filter(isSupportedUpload);
  const rejected = incomingFiles.length - newFiles.length;
  if (rejected > 0) {
    els.jsonConsole.textContent = JSON.stringify({
      warning: `${rejected} unsupported file${rejected === 1 ? "" : "s"} skipped`,
      supported: "PDF, JPG, PNG, WEBP, HEIC, HEIF",
    }, null, 2);
  }
  if (newFiles.length === 0) return;

  appState.selectedFiles = [...appState.selectedFiles, ...newFiles];
  renderDocumentPreview(appState.selectedFiles[0]);
  
  renderFileGrid();
  
  els.dropzone.classList.add("has-file");
  updateRunState();

  // Auto-process on phone
  if (appState.phoneMode && appState.selectedFiles.length > 0 && !appState.isProcessing) {
    if (appState.sheetInfo) {
      runExtraction();
    } else {
      updatePhoneStatus("Connecting to Excel...", "busy");
      inspectSheet().then(() => {
        if (appState.sheetInfo) runExtraction();
      });
    }
  }
}

function renderFileGrid() {
  const count = appState.selectedFiles.length;
  
  // Cleanup old URLs to prevent memory leaks
  appState.filePreviewUrls.forEach((url) => URL.revokeObjectURL(url));
  appState.filePreviewUrls.clear();

  if (count === 0) {
    els.dropEmptyState.style.display = "flex";
    els.fileGrid.style.display = "none";
    els.fileBadge.style.display = "none";
    return;
  }

  els.dropEmptyState.style.display = "none";
  els.fileGrid.style.display = "block";
  
  els.fileItems.innerHTML = appState.selectedFiles.map((file, index) => {
    const isImage = (file.type || "").startsWith("image/");
    let previewUrl = "";
    if (isImage) {
      previewUrl = URL.createObjectURL(file);
      appState.filePreviewUrls.set(index, previewUrl);
    }

    return `
      <div class="file-icon-card" style="display: flex; flex-direction: column; align-items: center; gap: 8px; position: relative;">
        <div style="width: 64px; height: 64px; background: #f8f9fa; border-radius: 12px; display: flex; align-items: center; justify-content: center; overflow: hidden; border: 1px solid rgba(0,0,0,0.08); box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
          ${isImage ? `<img src="${previewUrl}" style="width: 100%; height: 100%; object-fit: cover;">` : `
            <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="color: rgba(0,0,0,0.3);"><path d="M13 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"/><polyline points="13 2 13 9 20 9"/></svg>
          `}
        </div>
        <span style="font-size: 0.65rem; color: #444; text-align: center; width: 80px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font-weight: 500;">${escapeHtml(file.name)}</span>
        <button onclick="event.stopPropagation(); removeFile(${index})" style="position: absolute; top: -6px; right: 2px; background: #000; color: #fff; border: none; border-radius: 50%; width: 20px; height: 20px; font-size: 12px; display: flex; align-items: center; justify-content: center; cursor: pointer; box-shadow: 0 2px 8px rgba(0,0,0,0.2); transition: transform 0.1s ease;">x</button>
      </div>
    `;
  }).join("");

  // Add a "Clear All" button if many files
  if (count > 1) {
    const clearBtn = document.createElement("div");
    clearBtn.innerHTML = `<button onclick="event.stopPropagation(); clearAllFiles()" style="grid-column: 1 / -1; margin-top: 10px; padding: 8px; background: rgba(255, 77, 77, 0.05); border: 1px solid rgba(255, 77, 77, 0.1); color: #ff4d4d; border-radius: 8px; font-size: 0.75rem; font-weight: 600; cursor: pointer;">Clear all ${count} files</button>`;
    els.fileItems.appendChild(clearBtn.firstChild);
  }

  els.fileBadge.textContent = count === 1 ? fileSize(appState.selectedFiles[0].size) : `${count} files`;
  els.fileBadge.style.display = "block";
  els.excelBoardStatus.textContent = count === 1 ? "File ready" : "Batch ready";

  if (appState.phoneMode) {
    updatePhoneStatus(appState.sheetInfo ? "Files ready. Sending to Excel..." : "Files ready. Connecting to Excel sheet...");
    maybeAutoRunPhone();
  }
}

function clearAllFiles() {
  appState.selectedFiles = [];
  els.dropzone.classList.remove("has-file");
  renderDocumentPreview(null);
  renderFileGrid();
  updateRunState();
}

function removeFile(index) {
  appState.selectedFiles.splice(index, 1);
  if (appState.selectedFiles.length === 0) {
    els.dropzone.classList.remove("has-file");
  }
  renderDocumentPreview(appState.selectedFiles[0] || null);
  renderFileGrid();
  updateRunState();
}

function renderDocumentPreview(file) {
  if (appState.previewUrl) {
    URL.revokeObjectURL(appState.previewUrl);
    appState.previewUrl = null;
  }

  els.documentPreview.classList.remove("has-image");
  els.previewImage.style.display = "none";
  els.previewImage.removeAttribute("src");
  els.previewFallback.style.display = "block";

  if (!file) {
    els.previewFallback.textContent = "No image selected";
    return;
  }

  const fallbackText = els.previewFallback;
  const fileType = file.type || "";
  const fileName = String(file.name || "").toLowerCase();
  fallbackText.textContent = fileType === "application/pdf" || fileName.endsWith(".pdf")
    ? "PDF source"
    : "Document";

  if (fileType.startsWith("image/") && !/heic|heif/i.test(fileType)) {
    appState.previewUrl = URL.createObjectURL(file);
    els.previewImage.src = appState.previewUrl;
    els.previewImage.style.display = "block";
    els.previewFallback.style.display = "none";
    els.documentPreview.classList.add("has-image");
  }
}

function updateRunState() {
  const ready = appState.hasApiKey && appState.sheetInfo && appState.selectedFiles.length > 0 && !appState.isProcessing;
  els.runButton.disabled = !ready;

  if (!appState.hasApiKey) {
    els.runHint.textContent = "Please add your Gemini API key in settings.";
  } else if (!appState.sheetInfo) {
    els.runHint.textContent = "Please connect your target Excel workbook.";
  } else if (appState.selectedFiles.length === 0) {
    els.runHint.textContent = "Please upload one or more documents.";
  } else if (appState.isProcessing) {
    els.runHint.textContent = "Running";
  } else {
    els.runHint.textContent = "Ready";
  }
}

function maybeAutoRunPhone() {
  if (!appState.phoneMode) return;
  if (!appState.hasApiKey || !appState.sheetInfo || appState.selectedFiles.length === 0 || appState.isProcessing) return;
  window.setTimeout(() => {
    if (appState.phoneMode) {
      if (!appState.isProcessing && appState.selectedFiles.length > 0 && appState.sheetInfo) {
        runExtraction();
      }
    }
  }, 220);
}

function setStep(name, state) {
  document.querySelectorAll(".step").forEach((step) => {
    if (step.dataset.step !== name) return;
    step.classList.remove("is-active", "is-ready");
    if (state === "active") step.classList.add("is-active");
    if (state === "ready") step.classList.add("is-ready");
  });
}

function resetSteps() {
  document.querySelectorAll(".step").forEach((step) => {
    step.classList.remove("is-active", "is-ready");
  });
  if (appState.sheetInfo) setStep("sheet", "ready");
}

function renderFlightChips() {
  const headers = appState.sheetInfo?.validHeaders || [];
  const labels = (headers.length ? headers : ["Invoice", "Date", "Total", "Tax", "Vendor"]).slice(0, 5);
  els.dataFlightLayer.innerHTML = labels.map((label, index) => `
    <span class="data-chip" style="--chip-y:${22 + index * 12}%; --chip-delay:${index * 260}ms">
      ${escapeHtml(label)}
    </span>
  `).join("");
}

function renderExcelPreview(headers = [], values = {}, highlighted = false) {
  const fallbackHeaders = ["Vendor", "Date", "Total", "Status"];
  const columns = (headers.length ? headers : fallbackHeaders).filter(Boolean).slice(0, 4);
  while (columns.length < 4) columns.push(fallbackHeaders[columns.length]);

  const headCells = columns.map((header) => `<div class="excel-cell head">${escapeHtml(header)}</div>`).join("");
  const rowCells = columns.map((header, index) => {
    const value = values[header] ?? "";
    const className = highlighted ? "excel-cell is-new" : "excel-cell muted";
    const display = value || (highlighted ? "" : (index === 3 ? "Ready" : "-"));
    return `<div class="${className}">${escapeHtml(display)}</div>`;
  }).join("");

  els.excelGridPreview.innerHTML = headCells + rowCells;
}

function startProcessingAnimation() {
  appState.isProcessing = true;
  updateRunState();
  resetSteps();
  setStep("vision", "active");
  els.processState.textContent = "Running";
  els.excelBoardStatus.textContent = "Scanning";
  els.extractionStage.classList.add("is-running");
  els.excelBoard.classList.add("is-receiving");
  els.commandCard?.classList.add("is-running");
  els.resultTable.style.display = "none";
  els.jsonConsole.style.display = "block";
  renderFlightChips();

  if (appState.phoneMode) {
    updatePhoneStatus("Reading document and adding the row...");
  }

  const frames = [
    "{",
    '  "schema_locked": true,',
    `  "columns": ${appState.sheetInfo?.validHeaders?.length || 0},`,
    '  "vision": "zooming source",',
    '  "fields": "flying to row",',
    '  "excel": "preparing append"',
    "}",
  ];
  let index = 1;
  els.jsonConsole.textContent = frames.slice(0, index).join("\n");

  appState.animationTimer = window.setInterval(() => {
    index = index < frames.length ? index + 1 : 4;
    els.jsonConsole.textContent = frames.slice(0, index).join("\n");
  }, 520);
}

function stopProcessingAnimation() {
  window.clearInterval(appState.animationTimer);
  appState.animationTimer = null;
  els.extractionStage.classList.remove("is-running");
  els.excelBoard.classList.remove("is-receiving");
  els.commandCard?.classList.remove("is-running");
  window.setTimeout(() => {
    if (!appState.isProcessing) {
      els.dataFlightLayer.innerHTML = "";
    }
  }, 700);
}

function finishProcessing(data) {
  appState.isProcessing = false;
  stopProcessingAnimation();
  setStep("vision", "ready");
  setStep("excel", "ready");
  setStep("done", "ready");
  els.processState.textContent = "Done";
  els.resultTitle.textContent = "Extraction Success";
  els.excelBoardStatus.textContent = `Row ${data.rowNumber}`;
  els.jsonConsole.textContent = JSON.stringify({
    rowNumber: data.rowNumber,
    added: data.changedCells,
  }, null, 2);
  renderExcelPreview(Object.keys(data.rowData || data.changedCells || {}), data.rowData || data.changedCells || {}, true);

  if (appState.phoneMode) {
    updatePhoneStatus(`Saved. Row ${data.rowNumber} added.`, "ok");
    resetPhoneFilePrompt();
  }
  updateRunState();
}

function failProcessing(message) {
  appState.isProcessing = false;
  stopProcessingAnimation();
  els.processState.textContent = "Error";
  els.excelBoardStatus.textContent = "Error";
  els.resultTable.style.display = "none";
  els.jsonConsole.style.display = "block";
  els.jsonConsole.textContent = JSON.stringify({error: message}, null, 2);
  if (appState.phoneMode) {
    updatePhoneStatus(`Could not add row: ${message}`, "warn");
  }
  updateRunState();
}

function resetPhoneFilePrompt() {
  appState.selectedFiles = [];
  els.fileInput.value = "";
  els.cameraInput.value = "";
  els.dropzone.classList.remove("has-file");
  els.dropTitle.textContent = "Upload another invoice";
  els.dropMeta.textContent = "Photo, image, or PDF";
  els.fileBadge.textContent = "Ready";
  renderDocumentPreview(null);
  renderFileGrid();
}

async function runExtraction() {
  if (!appState.sheetInfo || appState.selectedFiles.length === 0 || appState.isProcessing) return;

  const filesToProcess = [...appState.selectedFiles];
  appState.isProcessing = true;
  updateRunState();

  startProcessingAnimation();

  try {
    for (let i = 0; i < filesToProcess.length; i++) {
      const file = filesToProcess[i];
      const progress = filesToProcess.length > 1 ? `[${i + 1}/${filesToProcess.length}] ` : "";
      els.processState.textContent = `${progress}Processing ${file.name}...`;
      const uploadFile = await optimizeUploadFile(file);
      const optimizedNote = uploadFile !== file ? ` (${fileSize(file.size)} -> ${fileSize(uploadFile.size)})` : "";
      els.processState.textContent = `${progress}Sending ${uploadFile.name}${optimizedNote}...`;

      const formData = new FormData();
      formData.append("excelPath", els.excelPath.value.trim());
      formData.append("sheet", els.sheetSelect.value);
      formData.append("document", uploadFile, uploadFile.name);

      const response = await fetch("/api/extract", {
        method: "POST",
        body: formData,
      });
      const data = await response.json().catch(() => ({}));
      
      if (!response.ok || data.ok === false) {
        throw new Error(data.error || `Request failed: ${response.status}`);
      }

      // Final file in batch gets the full UI treatment
      if (i === filesToProcess.length - 1) {
        finishProcessing(data);
        renderResult(data);
        showSuccess(data);
      } else {
        // Intermediate success logging
        els.jsonConsole.innerHTML += `\n\n${progress} ${file.name} -> Row ${data.rowNumber} OK`;
      }
    }

    appState.selectedFiles = [];
    els.fileInput.value = "";
    els.cameraInput.value = "";
    els.dropzone.classList.remove("has-file");
    renderDocumentPreview(null);
    renderFileGrid();
    await loadState();
    window.setTimeout(() => inspectSheet(els.sheetSelect.value), 2800);
  } catch (error) {
    failProcessing(error.message);
  } finally {
    appState.isProcessing = false;
    updateRunState();
  }
}

function renderResult(data) {
  const rows = Object.entries(data.changedCells || data.rowData || {});
  els.resultTitle.textContent = `Row ${data.rowNumber} added`;
  els.resultTable.innerHTML = rows.length
    ? rows.map(([key, value]) => `
      <tr>
        <td>${escapeHtml(key)}</td>
        <td>${escapeHtml(value)}</td>
      </tr>
    `).join("")
    : `<tr><td>Status</td><td>Row saved with empty extracted fields.</td></tr>`;
  els.jsonConsole.style.display = "none";
  els.resultTable.style.display = "table";

  els.resultPanel.classList.remove("is-highlighted");
  void els.resultPanel.offsetWidth;
  els.resultPanel.classList.add("is-highlighted");
}

function showSuccess(data) {
  els.successTitle.textContent = `Row ${data.rowNumber} added`;
  els.successText.textContent = `${Object.keys(data.changedCells || {}).length} cells written to Excel.`;
  els.successOverlay.classList.add("is-visible");
  window.setTimeout(() => {
    els.successOverlay.classList.remove("is-visible");
  }, 2600);
}

async function saveSettings() {
  setBusy(els.saveSettingsButton, "Saving", true);
  try {
    const data = await apiJson("/api/settings", {
      method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify({
        apiKey: els.apiKey.value.trim(),
        model: GEMINI_MODEL,
      }),
    });
    els.apiKey.value = "";
    renderState({config: data.config, history: [], share: appState.share || {}});
    await loadState();
    els.settingsDialog.close();
  } catch (error) {
    els.apiKey.placeholder = error.message;
  } finally {
    setBusy(els.saveSettingsButton, "Save settings", false);
  }
}

async function openExcel() {
  if (!els.excelPath.value.trim()) {
    await browseExcel();
    return;
  }
  els.openExcelButton.disabled = true;
  try {
    await apiJson("/api/open-sheet", {
      method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify({path: els.excelPath.value.trim()}),
    });
  } catch (error) {
    els.sheetSummary.innerHTML = `<span class="error-text">${escapeHtml(error.message)}</span>`;
  } finally {
    els.openExcelButton.disabled = false;
  }
}

async function copyShareLink() {
  const value = appState.share?.phoneUrl || appState.share?.lanUrl || appState.share?.currentUrl || window.location.href;
  try {
    await navigator.clipboard.writeText(value);
    els.copyShareButton.textContent = "Copied";
    window.setTimeout(() => {
      els.copyShareButton.textContent = "Copy";
    }, 1200);
  } catch (_error) {
    els.shareUrl.textContent = value;
  }
}



function bindEvents() {
  if (els.settingsButton) els.settingsButton.addEventListener("click", () => els.settingsDialog?.showModal());
  if (els.closeSettingsButton) els.closeSettingsButton.addEventListener("click", () => els.settingsDialog?.close());
  if (els.saveSettingsButton) els.saveSettingsButton.addEventListener("click", saveSettings);
  if (els.connectExcelButton) els.connectExcelButton.addEventListener("click", browseExcel);
  if (els.inspectButton) els.inspectButton.addEventListener("click", () => inspectSheet());
  if (els.openExcelButton) els.openExcelButton.addEventListener("click", openExcel);
  if (els.excelPath) {
    els.excelPath.addEventListener("input", () => {
      // Check if value exists in datalist
      const options = Array.from(els.recentExcelPaths.options).map(o => o.value);
      if (options.includes(els.excelPath.value)) {
        inspectSheet();
      }
    });
    els.excelPath.addEventListener("keypress", (e) => {
      if (e.key === "Enter") inspectSheet();
    });
  }

  if (els.sheetSelect) els.sheetSelect.addEventListener("change", () => inspectSheet(els.sheetSelect.value));
  if (els.browseFileButton) els.browseFileButton.addEventListener("click", () => els.fileInput?.click());
  if (els.cameraButton) els.cameraButton.addEventListener("click", () => els.cameraInput?.click());
  if (els.fileInput) els.fileInput.addEventListener("change", () => setSelectedFiles(els.fileInput.files));
  if (els.cameraInput) els.cameraInput.addEventListener("change", () => setSelectedFiles(els.cameraInput.files));
  if (els.runButton) els.runButton.addEventListener("click", runExtraction);
  if (els.copyShareButton) els.copyShareButton.addEventListener("click", copyShareLink);
  if (els.refreshShareButton) els.refreshShareButton.addEventListener("click", loadState);
  if (els.shareButton) els.shareButton.addEventListener("click", (event) => {
    event.stopPropagation();
    const isOpen = els.sharePopover.classList.toggle("is-open");
    els.sharePopover.setAttribute("aria-hidden", String(!isOpen));
  });
  if (els.sharePopover) els.sharePopover.addEventListener("click", (event) => event.stopPropagation());
  document.addEventListener("click", () => {
    els.sharePopover?.classList.remove("is-open");
    els.sharePopover?.setAttribute("aria-hidden", "true");
  });

  if (els.dropzone) {
    els.dropzone.addEventListener("click", () => els.fileInput?.click());

    ["dragenter", "dragover"].forEach((eventName) => {
      els.dropzone.addEventListener(eventName, (event) => {
        event.preventDefault();
        els.dropzone.classList.add("is-dragging");
      });
    });

    ["dragleave", "drop"].forEach((eventName) => {
      els.dropzone.addEventListener(eventName, (event) => {
        event.preventDefault();
        els.dropzone.classList.remove("is-dragging");
      });
    });

    els.dropzone.addEventListener("drop", (event) => {
      const files = event.dataTransfer.files;
      if (files.length > 0) setSelectedFiles(files);
    });
  }
}

async function initApp() {
  initEls();
  appState.phoneMode = detectPhoneMode();
  document.body.classList.toggle("phone-mode", appState.phoneMode);
  if (appState.phoneMode) {
    updatePhoneStatus("Preparing upload screen...");
  }
  
  bindEvents();
  renderExcelPreview();
  try {
    await loadState();
  } catch (error) {
    if (els.jsonConsole) {
      els.jsonConsole.textContent = JSON.stringify({error: error.message}, null, 2);
    }
    if (appState.phoneMode) {
      updatePhoneStatus(`Initial setup error: ${error.message}`, "warn");
    }
  }
}

document.addEventListener("DOMContentLoaded", initApp);
