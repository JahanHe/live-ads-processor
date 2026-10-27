const dropzone = document.querySelector("#dropzone");
const fileInput = document.querySelector("#fileInput");
const chooseButton = document.querySelector("#chooseButton");
const longPlanOcrInput = document.querySelector("#longPlanOcrInput");
const ocrLongPlanButton = document.querySelector("#ocrLongPlanButton");
const ocrClipboardButton = document.querySelector("#ocrClipboardButton");
const installOcrButton = document.querySelector("#installOcrButton");
const chooseScreenshotButton = document.querySelector("#chooseScreenshotButton");
const pasteScreenshotButton = document.querySelector("#pasteScreenshotButton");
const selectedFile = document.querySelector("#selectedFile");
const fileName = document.querySelector("#fileName");
const clearButton = document.querySelector("#clearButton");
const clearLongPlanButton = document.querySelector("#clearLongPlanButton");
const processButton = document.querySelector("#processButton");
const resetWorkspaceButton = document.querySelector("#resetWorkspaceButton");
const message = document.querySelector("#message");
const metrics = document.querySelector("#metrics");
const rowsMetric = document.querySelector("#rowsMetric");
const spendMetric = document.querySelector("#spendMetric");
const dealMetric = document.querySelector("#dealMetric");
const roiMetric = document.querySelector("#roiMetric");
const downloadCard = document.querySelector("#downloadCard");
const outputName = document.querySelector("#outputName");
const downloadLink = document.querySelector("#downloadLink");
const imageDownloadLink = document.querySelector("#imageDownloadLink");
const copyImageButton = document.querySelector("#copyImageButton");
const useThousands = document.querySelector("#useThousands");
const transposeSummary = document.querySelector("#transposeSummary");
const removeZeroColumns = document.querySelector("#removeZeroColumns");
const previewSection = document.querySelector("#previewSection");
const sheet2Table = document.querySelector("#sheet2Table");
const sheet3Table = document.querySelector("#sheet3Table");
const screenshotInput = document.querySelector("#screenshotInput");
const clearScreenshotButton = document.querySelector("#clearScreenshotButton");
const screenshotThumbs = document.querySelector("#screenshotThumbs");
const outputNameInput = document.querySelector("#outputNameInput");
const warningsCard = document.querySelector("#warningsCard");
const warningsList = document.querySelector("#warningsList");
const screenshotCheckCard = document.querySelector("#screenshotCheckCard");
const screenshotCheckTitle = document.querySelector("#screenshotCheckTitle");
const screenshotCheckStatus = document.querySelector("#screenshotCheckStatus");
const screenshotCheckMessage = document.querySelector("#screenshotCheckMessage");
const screenshotCheckGrid = document.querySelector("#screenshotCheckGrid");
const crossHighlight = document.querySelector("#crossHighlight");
const screenshotZone = document.querySelector("#screenshotZone");
const pastePanel = document.querySelector("#pastePanel");
const pasteTarget = document.querySelector("#pasteTarget");
const rateMetricsAsPercent = document.querySelector("#rateMetricsAsPercent");
const longPlanPanel = document.querySelector(".long-plan-panel");
const settingsPanel = document.querySelector("#settingsPanel");
const toggleLongPlan = document.querySelector("#toggleLongPlan");
const toggleSettings = document.querySelector("#toggleSettings");
const longPlanOcrStatus = document.querySelector("#longPlanOcrStatus");
const ocrPreviewPanel = document.querySelector("#ocrPreviewPanel");
const ocrPreviewImage = document.querySelector("#ocrPreviewImage");
const ocrPreviewButton = document.querySelector("#ocrPreviewButton");
const clearOcrPreviewButton = document.querySelector("#clearOcrPreviewButton");
const imageLightbox = document.querySelector("#imageLightbox");
const lightboxImage = document.querySelector("#lightboxImage");
const lightboxTitle = document.querySelector("#lightboxTitle");
const closeLightboxButton = document.querySelector("#closeLightboxButton");
const closeLightboxBackdrop = document.querySelector("#closeLightboxBackdrop");
const sheet3SelectButton = document.querySelector("#sheet3SelectButton");
const sheet3CopySelectedButton = document.querySelector("#sheet3CopySelectedButton");
const sheet3ClearSelectionButton = document.querySelector("#sheet3ClearSelectionButton");

let currentFiles = [];
let screenshotFiles = [];
let screenshotThumbUrls = [];
let ocrPreviewUrl = "";
let previewData = {};
let activeOcrButton = null;
let ocrInstallable = false;
let sheet3SelectionMode = false;
let sessionId = 0;
let processController = null;
let ocrController = null;
let lightboxTrigger = null;
const sheet3SelectedCells = new Set();

document.body.classList.toggle("is-mac-electron", /Electron\//.test(navigator.userAgent) && /Mac/.test(navigator.platform));

function setPanelOpen(panel, open) {
  panel.hidden = !open;
  const button = panel === longPlanPanel ? toggleLongPlan : toggleSettings;
  button.setAttribute("aria-expanded", String(open));
}

toggleLongPlan.addEventListener("click", () => setPanelOpen(longPlanPanel, longPlanPanel.hidden));
toggleSettings.addEventListener("click", () => setPanelOpen(settingsPanel, settingsPanel.hidden));

const moneyFormatter = new Intl.NumberFormat("zh-CN", {
  style: "currency",
  currency: "CNY",
  minimumFractionDigits: 2,
});

const numberFormatter = new Intl.NumberFormat("zh-CN");

function resetResult() {
  processController?.abort();
  processController = null;
  setProcessBusy(false);
  metrics.hidden = true;
  downloadCard.hidden = true;
  warningsCard.hidden = true;
  screenshotCheckCard.hidden = true;
  previewSection.hidden = true;
  sheet2Table.innerHTML = "";
  sheet3Table.innerHTML = "";
  previewData = {};
  setSheet3SelectionMode(false);
  outputName.textContent = "";
  downloadLink.removeAttribute("href");
  imageDownloadLink.hidden = true;
  copyImageButton.hidden = true;
  imageDownloadLink.removeAttribute("href");
  delete copyImageButton.dataset.imageUrl;
  warningsList.innerHTML = "";
  screenshotCheckGrid.innerHTML = "";
  screenshotCheckMessage.textContent = "";
  screenshotCheckStatus.textContent = "";
}

function setProcessBusy(busy) {
  processButton.disabled = busy || !currentFiles.length;
  processButton.textContent = busy ? "正在生成…" : "生成表格";
  processButton.classList.toggle("is-busy", busy);
  processButton.setAttribute("aria-busy", String(busy));
}

function resetWorkspace() {
  sessionId += 1;
  ocrController?.abort();
  ocrController = null;
  closeImageLightbox();
  fileInput.value = "";
  screenshotInput.value = "";
  longPlanOcrInput.value = "";
  outputNameInput.value = "";
  setFiles([]);
  screenshotFiles = [];
  setScreenshots();
  clearLongPlanRow();
  setLongPlanOcrState("");
  setPanelOpen(longPlanPanel, false);
  setPanelOpen(settingsPanel, false);
  [useThousands, transposeSummary, removeZeroColumns, crossHighlight, rateMetricsAsPercent].forEach((input) => {
    input.checked = input.defaultChecked;
  });
  document.querySelectorAll("input[name='decimalMode']").forEach((input) => {
    input.checked = input.defaultChecked;
  });
  pastePanel.hidden = true;
  screenshotZone.classList.remove("is-pasting");
  pasteTarget.textContent = "在这里按 Cmd+V / Ctrl+V 粘贴截图";
  window.getSelection()?.removeAllRanges();
  message.textContent = "已清空当前资料，可以重新开始。已下载的文件不会删除。";
  window.scrollTo({ top: 0, behavior: "instant" });
  chooseButton.focus({ preventScroll: true });
}

resetWorkspaceButton.addEventListener("click", resetWorkspace);

function setFiles(files) {
  currentFiles = [...files];
  dropzone.classList.toggle("has-files", currentFiles.length > 0);
  resetResult();
  if (!currentFiles.length) {
    selectedFile.hidden = true;
    fileName.textContent = "";
    processButton.disabled = true;
    message.textContent = "";
    return;
  }
  selectedFile.hidden = false;
  fileName.textContent = currentFiles.length === 1
    ? currentFiles[0].name
    : `${currentFiles.length} 个表格：${currentFiles.map((file) => file.name).join("、")}`;
  processButton.disabled = false;
  message.textContent = "文件已就绪，可以生成 XLSX。";
}

function setScreenshots() {
  resetResult();
  screenshotZone.classList.toggle("has-files", screenshotFiles.length > 0);
  clearScreenshotButton.hidden = !screenshotFiles.length;
  pastePanel.hidden = true;
  screenshotZone.classList.remove("is-pasting");
  screenshotThumbUrls.forEach((url) => URL.revokeObjectURL(url));
  screenshotThumbUrls = [];
  screenshotThumbs.replaceChildren();
  if (!screenshotFiles.length) {
    screenshotThumbs.hidden = true;
    return;
  }
  screenshotThumbUrls = screenshotFiles.map((file) => URL.createObjectURL(file));
  screenshotFiles.forEach((file, index) => {
    const card = document.createElement("div");
    card.className = "thumb-card";
    card.innerHTML = '<button class="thumb-image-button" type="button"><img></button><div class="thumb-meta"><span></span><button class="thumb-remove" data-icon="x" type="button">移除</button></div>';
    const previewButton = card.querySelector(".thumb-image-button");
    previewButton.dataset.previewScreenshot = index;
    previewButton.setAttribute("aria-label", `放大查看 ${file.name}`);
    const img = card.querySelector("img");
    img.src = screenshotThumbUrls[index];
    img.alt = file.name;
    const label = card.querySelector("span");
    label.textContent = file.name;
    label.title = file.name;
    const removeButton = card.querySelector(".thumb-remove");
    removeButton.dataset.removeScreenshot = index;
    removeButton.setAttribute("aria-label", `移除 ${file.name}`);
    screenshotThumbs.appendChild(card);
  });
  screenshotThumbs.hidden = false;
}

function showImageLightbox(src, title = "图片预览") {
  if (!src) return;
  lightboxTrigger = document.activeElement;
  lightboxImage.src = src;
  lightboxImage.alt = title;
  lightboxTitle.textContent = title;
  imageLightbox.hidden = false;
  document.querySelector(".shell").inert = true;
  document.body.style.overflow = "hidden";
  closeLightboxButton.focus();
}

function closeImageLightbox() {
  imageLightbox.hidden = true;
  lightboxImage.removeAttribute("src");
  lightboxImage.alt = "";
  document.querySelector(".shell").inert = false;
  document.body.style.overflow = "";
  if (lightboxTrigger?.isConnected) lightboxTrigger.focus({ preventScroll: true });
  lightboxTrigger = null;
}

function friendlyErrorText(error, fallback = "操作失败，请重试。") {
  const raw = String(error?.message || error || "");
  if (!raw) return fallback;
  if (/writeText|write\(|ClipboardItem/i.test(raw)) {
    return "无法写入剪切板。请检查浏览器剪切板权限，或改用下载文件。";
  }
  if (/Clipboard|clipboard|NotAllowedError|permission denied|Permission denied|Read permission denied/i.test(raw)) {
    return "无法读取剪切板。请先允许剪切板权限，或点击粘贴截图后在输入框里按 Cmd+V / Ctrl+V。";
  }
  if (/network|fetch|Failed to fetch/i.test(raw)) {
    return "连接本地处理服务失败，请确认处理面板仍在运行。";
  }
  if (/timeout/i.test(raw)) {
    return "处理超时，请换一张更清晰或裁剪范围更小的截图后重试。";
  }
  return raw.replace(/^Error:\s*/i, "") || fallback;
}

function setOcrPreview(file) {
  if (ocrPreviewUrl) URL.revokeObjectURL(ocrPreviewUrl);
  ocrPreviewUrl = URL.createObjectURL(file);
  ocrPreviewImage.src = ocrPreviewUrl;
  ocrPreviewImage.alt = file.name || "长期计划识别截图预览";
  ocrPreviewPanel.hidden = false;
  setPanelOpen(longPlanPanel, true);
}

function clearOcrPreview() {
  if (ocrPreviewUrl) URL.revokeObjectURL(ocrPreviewUrl);
  ocrPreviewUrl = "";
  ocrPreviewImage.removeAttribute("src");
  ocrPreviewImage.alt = "长期计划识别截图预览";
  ocrPreviewPanel.hidden = true;
}

function longPlanInputs() {
  return [...document.querySelectorAll("[data-long-plan]")];
}

function longPlanRow() {
  const row = {};
  longPlanInputs().forEach((input) => {
    const value = input.value.trim();
    if (value) row[input.dataset.longPlan] = value;
  });
  return row;
}

function fillLongPlanRow(fields) {
  longPlanInputs().forEach((input) => {
    input.value = "";
  });
  longPlanInputs().forEach((input) => {
    const value = fields[input.dataset.longPlan];
    if (value !== undefined && value !== null && value !== "") input.value = value;
  });
  setPanelOpen(longPlanPanel, true);
}

function clearLongPlanRow() {
  longPlanInputs().forEach((input) => {
    input.value = "";
  });
  resetResult();
  longPlanOcrStatus.textContent = "";
  clearOcrPreview();
  message.textContent = "长期计划数据行已清空。";
}

function addScreenshots(files) {
  const images = [...files].filter((file) => file.type.startsWith("image/") || /\.(avif|bmp|gif|heic|heif|jpeg|jpg|png|tif|tiff|webp)$/i.test(file.name));
  if (!images.length) return;
  screenshotFiles = [...screenshotFiles, ...images];
  setScreenshots();
}

function setLongPlanOcrState(state, text = "", button = null) {
  const isBusy = state === "busy";
  longPlanPanel.setAttribute("aria-busy", String(isBusy));
  longPlanOcrStatus.textContent = text;
  longPlanOcrStatus.className = `inline-status${state ? ` is-${state}` : ""}`;
  [ocrLongPlanButton, ocrClipboardButton, installOcrButton, clearLongPlanButton].forEach((control) => {
    control.disabled = isBusy;
  });
  if (activeOcrButton && activeOcrButton !== button) {
    activeOcrButton.textContent = activeOcrButton.dataset.idleText;
  }
  activeOcrButton = isBusy ? button : null;
  if (button) {
    button.dataset.idleText ||= button.textContent;
    button.textContent = isBusy ? "识别中..." : button.dataset.idleText;
  }
}

function setInstallOcrVisible(visible, label = "安装截图识别组件") {
  ocrInstallable = visible;
  installOcrButton.hidden = !visible;
  installOcrButton.textContent = label || "安装截图识别组件";
}

async function refreshOcrStatus(showReady = false) {
  try {
    const response = await fetch("/api/ocr-status");
    const payload = await response.json();
    if (!response.ok || !payload.ok) return;
    setInstallOcrVisible(Boolean(!payload.installed && payload.installable), payload.installLabel);
    if (payload.installed && showReady) {
      setLongPlanOcrState("done", "截图识别组件已就绪。");
    }
  } catch {
    // OCR is optional; keep the form usable even if status probing fails.
  }
}

async function recognizeLongPlanImage(file, button = null) {
  const currentSession = sessionId;
  ocrController?.abort();
  const controller = new AbortController();
  ocrController = controller;
  setOcrPreview(file);
  const formData = new FormData();
  formData.append("ocrImage", file);
  setLongPlanOcrState("busy", "正在识别截图，右侧可先查看原图。", button);
  message.textContent = "正在识别长期计划截图，识别结果会自动填入左侧表单。";
  const response = await fetch("/api/ocr-long-plan", {
    method: "POST",
    body: formData,
    signal: controller.signal,
  });
  const payload = await response.json();
  if (currentSession !== sessionId || controller.signal.aborted) return;
  if (!response.ok || !payload.ok) {
    const error = new Error(payload.error || "OCR 识别失败。");
    error.code = payload.code || "";
    error.hint = payload.hint || "";
    error.installCommand = payload.installCommand || "";
    error.canInstall = Boolean(payload.canInstall);
    error.installLabel = payload.installLabel || "";
    throw error;
  }
  fillLongPlanRow(payload.fields || {});
  resetResult();
  setLongPlanOcrState("done", "识别完成，已预填。请对照右侧截图检查并手动修正。");
  message.textContent = "OCR 已预填长期计划数据行。识别错的数字可以直接在左侧修改。";
}

function ocrErrorText(error) {
  if (error.code !== "missing_tesseract") return friendlyErrorText(error, "OCR 识别失败，请换一张更清晰的截图后重试。");
  setInstallOcrVisible(Boolean(error.canInstall), error.installLabel);
  return [
    friendlyErrorText(error, "截图识别组件不可用。"),
    error.hint,
    error.canInstall ? "点击“安装截图识别组件”即可启用自动识别。" : error.installCommand ? `解决命令：${error.installCommand}` : "",
  ].filter(Boolean).join("\n");
}

async function imageFromClipboard() {
  if (!navigator.clipboard?.read) {
    throw new Error("当前环境不支持直接读取剪切板图片。");
  }
  const items = await navigator.clipboard.read();
  for (const item of items) {
    const imageType = item.types.find((type) => type.startsWith("image/"));
    if (!imageType) continue;
    const blob = await item.getType(imageType);
    return new File([blob], `剪切板长期计划截图-${Date.now()}.png`, { type: imageType });
  }
  throw new Error("剪切板里没有图片。");
}

function formatSummary(summary) {
  rowsMetric.textContent = numberFormatter.format(summary.rows || 0);
  spendMetric.textContent = moneyFormatter.format(summary.spend_yuan || 0);
  dealMetric.textContent = moneyFormatter.format(summary.deal_amount || 0);
  roiMetric.textContent = Number(summary.deal_roi || 0).toFixed(2);
  metrics.hidden = false;
}

function formatCheckValue(value, label = "") {
  if (value === null || value === undefined || value === "") return "未识别";
  if (typeof value !== "number") return value;
  const fractionDigits = label.includes("ROI") ? 4 : value % 1 ? 2 : 0;
  return new Intl.NumberFormat("zh-CN", {
    useGrouping: true,
    minimumFractionDigits: fractionDigits,
    maximumFractionDigits: fractionDigits,
  }).format(value);
}

function renderScreenshotCheck(check) {
  if (!check || check.status === "none") {
    screenshotCheckCard.hidden = true;
    return;
  }
  const statusText = {
    matched: "一致",
    partial: "部分核对",
    mismatch: "有差异",
    unavailable: "无法自动核对",
  }[check.status] || "待核对";
  screenshotCheckTitle.textContent = "截图与表格数据核对";
  screenshotCheckStatus.textContent = statusText;
  screenshotCheckStatus.className = `check-status is-${check.status || "unknown"}`;
  screenshotCheckMessage.textContent = check.message || "";
  screenshotCheckGrid.innerHTML = (check.items || []).map((item) => {
    const pending = { label_missing: "未检出同名表头", number_missing: "数字未识别", conflict: "截图数值冲突" };
    const state = item.match === true ? "一致" : item.match === false ? "不一致" : (pending[item.reason] || "未识别");
    const stateClass = item.match === true ? "is-match" : item.match === false ? "is-mismatch" : "is-missing";
    const screenshotValue = item.screenshotValue === null || item.screenshotValue === undefined ? state : formatCheckValue(item.screenshotValue, item.label);
    return `
      <article class="check-item ${stateClass}">
        <span>${item.label}</span>
        <strong>${state}</strong>
        <small>表格：${formatCheckValue(item.tableValue, item.label)} ｜ 截图：${screenshotValue}</small>
      </article>
    `;
  }).join("");
  screenshotCheckCard.hidden = false;
}

function activeDecimalMode() {
  return document.querySelector("input[name='decimalMode']:checked").value;
}

function shouldFormatPercent(row) {
  if (!rateMetricsAsPercent.checked) return false;
  const metric = String(row?.[0] || "");
  return ["直播间进入率", "新增/场观", "评论/场观"].includes(metric);
}

function formatCell(value, row = null) {
  if (value === null || value === undefined || value === "") return "";
  if (typeof value !== "number") return value;
  const decimalMode = activeDecimalMode();
  if (shouldFormatPercent(row)) {
    return new Intl.NumberFormat("zh-CN", {
      style: "percent",
      minimumFractionDigits: decimalMode === "fixed2" ? 2 : 0,
      maximumFractionDigits: decimalMode === "fixed2" ? 2 : 10,
    }).format(value);
  }
  const hasFraction = Math.abs(value % 1) > 1e-10;
  const options = decimalMode === "fixed2" || hasFraction
    ? { minimumFractionDigits: decimalMode === "fixed2" ? 2 : 0, maximumFractionDigits: decimalMode === "fixed2" ? 2 : 10 }
    : { maximumFractionDigits: 0 };
  return new Intl.NumberFormat("zh-CN", {
    useGrouping: useThousands.checked,
    ...options,
  }).format(value);
}

function rowGroupClass(rowIndex) {
  if (rowIndex >= 1 && rowIndex <= 4) return "row-group-1";
  if (rowIndex >= 5 && rowIndex <= 10) return "row-group-2";
  if (rowIndex >= 11 && rowIndex <= 14) return "row-group-3";
  if (rowIndex >= 15 && rowIndex <= 17) return "row-group-4";
  if (rowIndex >= 18 && rowIndex <= 20) return "row-group-5";
  if (rowIndex >= 21 && rowIndex <= 24) return "row-group-6";
  return "";
}

function renderTable(table, rows, grouped = false, options = {}) {
  table.innerHTML = "";
  table.classList.toggle("is-highlightable", crossHighlight.checked);
  table.classList.toggle("is-selectable", Boolean(options.selectable));
  const thead = document.createElement("thead");
  const tbody = document.createElement("tbody");
  const headerRow = document.createElement("tr");
  for (const value of rows[0] || []) {
    const th = document.createElement("th");
    th.textContent = formatCell(value, rows[0]);
    headerRow.appendChild(th);
  }
  thead.appendChild(headerRow);

  rows.slice(1).forEach((row, index) => {
    const tr = document.createElement("tr");
    if (grouped) tr.className = rowGroupClass(index + 1);
    row.forEach((value, colIndex) => {
      const td = document.createElement("td");
      td.textContent = formatCell(value, row);
      td.dataset.row = String(index + 1);
      td.dataset.col = String(colIndex);
      td.dataset.key = `${index + 1}:${colIndex}`;
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  table.appendChild(thead);
  table.appendChild(tbody);
}

function clearSheet3Selection() {
  sheet3SelectedCells.clear();
  sheet3Table.querySelectorAll(".is-selected-cell").forEach((cell) => {
    cell.classList.remove("is-selected-cell");
  });
}

function setSheet3SelectionMode(enabled) {
  sheet3SelectionMode = enabled;
  sheet3Table.classList.toggle("is-selection-mode", enabled);
  sheet3SelectButton.classList.toggle("is-active", enabled);
  sheet3SelectButton.textContent = enabled ? "退出选择" : "选择单元格";
  if (!enabled) clearSheet3Selection();
}

function toggleSheet3Cell(cell) {
  const key = cell.dataset.key;
  if (!key) return;
  if (sheet3SelectedCells.has(key)) {
    sheet3SelectedCells.delete(key);
    cell.classList.remove("is-selected-cell");
    return;
  }
  sheet3SelectedCells.add(key);
  cell.classList.add("is-selected-cell");
}

function selectedSheet3Rows() {
  const rows = previewData.sheet3 || [];
  const selected = [...sheet3SelectedCells].map((key) => key.split(":").map(Number));
  if (!selected.length) return [];
  const rowIndexes = [...new Set(selected.map(([row]) => row))].sort((a, b) => a - b);
  const colIndexes = [...new Set(selected.map(([, col]) => col))].sort((a, b) => a - b);
  return rowIndexes.map((rowIndex) => {
    const raw = rows[rowIndex] || [];
    return colIndexes.map((colIndex) => sheet3SelectedCells.has(`${rowIndex}:${colIndex}`) ? raw[colIndex] ?? "" : "");
  });
}

function clearTableHighlight(table) {
  table.querySelectorAll(".is-row-highlight, .is-col-highlight, .is-active-cell").forEach((cell) => {
    cell.classList.remove("is-row-highlight", "is-col-highlight", "is-active-cell");
  });
}

function applyCrossHighlight(cell) {
  const table = cell.closest("table");
  if (!table || !crossHighlight.checked) return;
  clearTableHighlight(table);
  const row = cell.parentElement;
  const colIndex = cell.cellIndex;
  row.querySelectorAll("td").forEach((td) => td.classList.add("is-row-highlight"));
  table.querySelectorAll("tr").forEach((tr) => {
    const target = tr.children[colIndex];
    if (target) target.classList.add("is-col-highlight");
  });
  cell.classList.add("is-active-cell");
}

function rowsToTsv(rows, excludeLabels = false) {
  const selectedRows = excludeLabels ? rows.slice(1) : rows;
  return selectedRows.map((row) => (excludeLabels ? row.slice(1) : row)
    .map((value) => formatCell(value, row)).join("\t")).join("\n");
}

async function copyRows(key, includeHeader) {
  const rows = previewData[key] || [];
  if (rows.length <= (includeHeader ? 0 : 1)) return;
  await navigator.clipboard.writeText(rowsToTsv(rows, !includeHeader));
  message.textContent = includeHeader ? "已复制全部表格。" : "已复制数据区，不含首行表头和第一列。";
}

chooseButton.addEventListener("click", () => fileInput.click());
chooseScreenshotButton.addEventListener("click", () => screenshotInput.click());
pasteScreenshotButton.addEventListener("click", async () => {
  const currentSession = sessionId;
  pastePanel.hidden = false;
  screenshotZone.classList.add("is-pasting");
  pasteTarget.focus();
  message.textContent = "请在粘贴框内按 Cmd+V / Ctrl+V 粘贴截图。";
  if (navigator.clipboard?.read) {
    try {
      const items = await navigator.clipboard.read();
      const files = [];
      for (const item of items) {
        const imageType = item.types.find((type) => type.startsWith("image/"));
        if (!imageType) continue;
        const blob = await item.getType(imageType);
        files.push(new File([blob], `剪切板截图-${Date.now()}.png`, { type: imageType }));
      }
      if (files.length) {
        if (currentSession !== sessionId) return;
        addScreenshots(files);
        message.textContent = "已直接读取剪切板截图。";
      }
    } catch (error) {
      if (currentSession !== sessionId) return;
      message.textContent = friendlyErrorText(error, "无法直接读取剪切板，请在粘贴框里按 Cmd+V / Ctrl+V。");
    }
  }
});

fileInput.addEventListener("change", () => {
  setFiles(fileInput.files);
});

screenshotInput.addEventListener("change", () => {
  addScreenshots(screenshotInput.files);
  screenshotInput.value = "";
});

clearButton.addEventListener("click", () => {
  fileInput.value = "";
  setFiles([]);
});

clearLongPlanButton.addEventListener("click", () => {
  clearLongPlanRow();
});

clearOcrPreviewButton.addEventListener("click", () => {
  clearOcrPreview();
  message.textContent = "识别截图预览已移除，已填写的数据不会被清空。";
});

ocrPreviewButton.addEventListener("click", () => {
  showImageLightbox(ocrPreviewUrl, "长期计划识别截图");
});

longPlanInputs().forEach((input) => {
  input.addEventListener("input", () => {
    resetResult();
    if (currentFiles.length) {
      message.textContent = "长期计划数据行已更新，可以重新生成 XLSX。";
    }
  });
});

clearScreenshotButton.addEventListener("click", () => {
  screenshotInput.value = "";
  screenshotFiles = [];
  setScreenshots();
});

screenshotThumbs.addEventListener("click", (event) => {
  const button = event.target.closest("[data-remove-screenshot]");
  if (!button) {
    const previewButton = event.target.closest("[data-preview-screenshot]");
    if (!previewButton) return;
    const index = Number(previewButton.dataset.previewScreenshot);
    showImageLightbox(screenshotThumbUrls[index], screenshotFiles[index]?.name || "截图预览");
    return;
  }
  const index = Number(button.dataset.removeScreenshot);
  if (!Number.isInteger(index)) return;
  const [removed] = screenshotFiles.splice(index, 1);
  setScreenshots();
  message.textContent = removed ? `已移除截图：${removed.name}` : "截图已移除。";
});

for (const eventName of ["dragenter", "dragover"]) {
  dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropzone.classList.add("is-dragging");
  });
}

for (const eventName of ["dragleave", "drop"]) {
  dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropzone.classList.remove("is-dragging");
  });
}

dropzone.addEventListener("drop", (event) => {
  const files = [...event.dataTransfer.files];
  if (!files.length) return;
  if (files.some((file) => !/\.(csv|tsv|txt|xlsx|xlsm|xls)$/i.test(file.name))) {
    message.textContent = "拖入区域只接收表格文件；消耗截图请用右侧控件选择。";
    return;
  }
  fileInput.files = event.dataTransfer.files;
  setFiles(files);
});

for (const eventName of ["dragenter", "dragover"]) {
  screenshotZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    screenshotZone.classList.add("is-dragging");
  });
}

for (const eventName of ["dragleave", "drop"]) {
  screenshotZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    screenshotZone.classList.remove("is-dragging");
  });
}

screenshotZone.addEventListener("drop", (event) => {
  addScreenshots(event.dataTransfer.files);
});

ocrLongPlanButton.addEventListener("click", () => longPlanOcrInput.click());
document.querySelector(".long-plan-heading .button-row").addEventListener("click", (event) => {
  event.stopPropagation();
});

installOcrButton.addEventListener("click", async () => {
  if (!ocrInstallable) return;
  const currentSession = sessionId;
  const idleText = installOcrButton.textContent;
  try {
    setLongPlanOcrState("busy", "正在安装截图识别组件，请不要关闭软件。", installOcrButton);
    installOcrButton.dataset.idleText = idleText;
    const response = await fetch("/api/install-ocr", { method: "POST" });
    const payload = await response.json();
    if (currentSession !== sessionId) return;
    if (!response.ok || !payload.ok || !payload.installed) {
      throw new Error(payload.error || "截图识别组件安装失败。");
    }
    setInstallOcrVisible(false);
    setLongPlanOcrState("done", payload.message || "截图识别组件已安装，可以识别截图。");
    message.textContent = "截图识别组件已安装，可以识别长期计划截图。";
    await refreshOcrStatus();
  } catch (error) {
    if (currentSession !== sessionId) return;
    const text = friendlyErrorText(error, "截图识别组件安装失败。");
    setLongPlanOcrState("error", text);
    message.textContent = text;
  } finally {
    if (currentSession === sessionId) installOcrButton.textContent = idleText;
  }
});

longPlanOcrInput.addEventListener("change", async () => {
  const currentSession = sessionId;
  const [file] = longPlanOcrInput.files;
  longPlanOcrInput.value = "";
  if (!file) return;
  try {
    await recognizeLongPlanImage(file, ocrLongPlanButton);
  } catch (error) {
    if (currentSession !== sessionId || error.name === "AbortError") return;
    const text = ocrErrorText(error);
    setLongPlanOcrState("error", text);
    message.textContent = text;
  }
});

ocrClipboardButton.addEventListener("click", async () => {
  const currentSession = sessionId;
  try {
    setLongPlanOcrState("busy", "正在读取剪切板截图。", ocrClipboardButton);
    const file = await imageFromClipboard();
    if (currentSession !== sessionId) return;
    await recognizeLongPlanImage(file, ocrClipboardButton);
  } catch (error) {
    if (currentSession !== sessionId || error.name === "AbortError") return;
    const text = ocrErrorText(error);
    setLongPlanOcrState("error", text);
    message.textContent = text;
  }
});

function handlePaste(event) {
  if (event.defaultPrevented) return;
  const files = [...event.clipboardData.files].filter((file) => file.type.startsWith("image/"));
  if (!files.length) return;
  event.preventDefault();
  addScreenshots(files);
  pasteTarget.textContent = "截图已添加，可以继续粘贴下一张";
  message.textContent = "已从剪切板添加截图。";
}

pasteTarget.addEventListener("paste", handlePaste);
document.addEventListener("paste", (event) => {
  if (!pastePanel.hidden || screenshotZone.contains(event.target) || event.target === document.body) {
    handlePaste(event);
  }
});

processButton.addEventListener("click", async () => {
  if (!currentFiles.length) return;
  const formData = new FormData();
  currentFiles.forEach((file) => formData.append("files", file));
  screenshotFiles.forEach((file) => formData.append("screenshots", file));
  formData.append("longPlanRow", JSON.stringify(longPlanRow()));
  if (outputNameInput.value.trim()) {
    formData.append("outputName", outputNameInput.value.trim());
  }
  formData.append("useThousands", useThousands.checked ? "true" : "false");
  formData.append("transposeSummary", transposeSummary.checked ? "true" : "false");
  formData.append("removeZeroColumns", removeZeroColumns.checked ? "true" : "false");
  formData.append("rateMetricsAsPercent", rateMetricsAsPercent.checked ? "true" : "false");
  formData.append("decimalMode", activeDecimalMode());

  resetResult();
  const controller = new AbortController();
  processController = controller;
  setProcessBusy(true);
  message.textContent = "正在生成三表结构，请稍等。";

  try {
    const response = await fetch("/api/process", {
      method: "POST",
      body: formData,
      signal: controller.signal,
    });
    const payload = await response.json();
    if (controller.signal.aborted) return;
    if (!response.ok || !payload.ok) {
      throw new Error(payload.error || "处理失败。");
    }

    formatSummary(payload.summary || {});
    outputName.textContent = payload.fileName;
    downloadLink.href = payload.downloadUrl;
    if (payload.combinedImageUrl) {
      imageDownloadLink.href = payload.combinedImageUrl;
      imageDownloadLink.download = payload.combinedImageName || "数据汇总拼接图.png";
      imageDownloadLink.hidden = false;
      copyImageButton.hidden = false;
      copyImageButton.dataset.imageUrl = payload.combinedImageUrl;
    }
    downloadCard.hidden = false;
    renderScreenshotCheck(payload.screenshotCheck);
    previewData = payload.previews || {};
    renderTable(sheet2Table, previewData.sheet2 || [], true);
    renderTable(sheet3Table, previewData.sheet3 || [], false, { selectable: true });
    setSheet3SelectionMode(false);
    previewSection.hidden = false;
    const warnings = payload.warnings || [];
    if (warnings.length) {
      warningsList.innerHTML = warnings.map((warning) => `<li>${warning}</li>`).join("");
      warningsCard.hidden = false;
    }
    message.textContent = "XLSX 已生成，数据汇总已按当前面板设置处理。";
    metrics.focus({ preventScroll: true });
    metrics.scrollIntoView({
      behavior: window.matchMedia("(prefers-reduced-motion: reduce)").matches ? "instant" : "smooth",
      block: "start",
    });
  } catch (error) {
    if (controller.signal.aborted) return;
    message.textContent = friendlyErrorText(error, "处理失败，请检查导入文件后重试。");
  } finally {
    if (processController === controller) {
      processController = null;
      setProcessBusy(false);
    }
  }
});

copyImageButton.addEventListener("click", async () => {
  const imageUrl = copyImageButton.dataset.imageUrl;
  if (!imageUrl) return;
  try {
    const response = await fetch(imageUrl);
    const blob = await response.blob();
    await navigator.clipboard.write([
      new ClipboardItem({ [blob.type || "image/png"]: blob }),
    ]);
    message.textContent = "已复制拼接图到剪切板。";
  } catch (error) {
    message.textContent = friendlyErrorText(error, "当前浏览器不允许直接复制图片，请使用下载拼接图。");
  }
});

closeLightboxButton.addEventListener("click", closeImageLightbox);
closeLightboxBackdrop.addEventListener("click", closeImageLightbox);
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && !imageLightbox.hidden) {
    closeImageLightbox();
  }
  if (event.key === "Tab" && !imageLightbox.hidden) {
    event.preventDefault();
    closeLightboxButton.focus();
  }
});

document.addEventListener("click", (event) => {
  const cell = event.target.closest(".preview-section td");
  if (cell) {
    if (sheet3SelectionMode && cell.closest("#sheet3Table")) {
      toggleSheet3Cell(cell);
      return;
    }
    applyCrossHighlight(cell);
    return;
  }
  const action = event.target.closest("[data-copy]")?.dataset.copy;
  if (!action) return;
  const [sheet, scope] = action.split("-");
  copyRows(sheet, scope === "all").catch((error) => {
    message.textContent = friendlyErrorText(error, "复制失败。");
  });
});

document.querySelectorAll("#useThousands, #crossHighlight, #rateMetricsAsPercent, input[name='decimalMode']").forEach((control) => {
  control.addEventListener("change", () => {
    if (!previewSection.hidden) {
      renderTable(sheet2Table, previewData.sheet2 || [], true);
      renderTable(sheet3Table, previewData.sheet3 || [], false, { selectable: true });
      setSheet3SelectionMode(false);
    }
  });
});

sheet3SelectButton.addEventListener("click", () => {
  setSheet3SelectionMode(!sheet3SelectionMode);
  message.textContent = sheet3SelectionMode ? "已进入 Sheet3 选择模式，点击想复制的单元格。" : "已退出 Sheet3 选择模式。";
});

sheet3ClearSelectionButton.addEventListener("click", () => {
  clearSheet3Selection();
  message.textContent = "已清除 Sheet3 选中单元格。";
});

sheet3CopySelectedButton.addEventListener("click", async () => {
  try {
    const manualText = String(window.getSelection?.().toString() || "").trim();
    if (manualText) {
      await navigator.clipboard.writeText(manualText);
      message.textContent = "已复制手动选中的内容。";
      return;
    }
    const rows = selectedSheet3Rows();
    if (!rows.length) {
      message.textContent = "请先在 Sheet3 中选中要复制的单元格。";
      return;
    }
    await navigator.clipboard.writeText(rowsToTsv(rows));
    message.textContent = "已复制 Sheet3 选中单元格。";
  } catch (error) {
    message.textContent = friendlyErrorText(error, "复制选中内容失败。");
  }
});

refreshOcrStatus();
