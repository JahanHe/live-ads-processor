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
const selectedScreenshots = document.querySelector("#selectedScreenshots");
const screenshotName = document.querySelector("#screenshotName");
const clearScreenshotButton = document.querySelector("#clearScreenshotButton");
const screenshotThumbs = document.querySelector("#screenshotThumbs");
const outputNameInput = document.querySelector("#outputNameInput");
const warningsCard = document.querySelector("#warningsCard");
const warningsList = document.querySelector("#warningsList");
const crossHighlight = document.querySelector("#crossHighlight");
const screenshotZone = document.querySelector("#screenshotZone");
const pastePanel = document.querySelector("#pastePanel");
const pasteTarget = document.querySelector("#pasteTarget");
const rateMetricsAsPercent = document.querySelector("#rateMetricsAsPercent");
const longPlanPanel = document.querySelector(".long-plan-panel");
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

let currentFiles = [];
let screenshotFiles = [];
let screenshotThumbUrls = [];
let ocrPreviewUrl = "";
let previewData = {};
let activeOcrButton = null;
let ocrInstallable = false;

const moneyFormatter = new Intl.NumberFormat("zh-CN", {
  style: "currency",
  currency: "CNY",
  minimumFractionDigits: 2,
});

const numberFormatter = new Intl.NumberFormat("zh-CN");

function resetResult() {
  metrics.hidden = true;
  downloadCard.hidden = true;
  warningsCard.hidden = true;
  previewSection.hidden = true;
  sheet2Table.innerHTML = "";
  sheet3Table.innerHTML = "";
  previewData = {};
  outputName.textContent = "";
  downloadLink.removeAttribute("href");
  imageDownloadLink.hidden = true;
  copyImageButton.hidden = true;
  imageDownloadLink.removeAttribute("href");
  warningsList.innerHTML = "";
}

function setFiles(files) {
  currentFiles = [...files];
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

function setScreenshots(files) {
  if (!screenshotFiles.length) {
    selectedScreenshots.hidden = true;
    screenshotThumbs.hidden = true;
    screenshotThumbs.innerHTML = "";
    screenshotThumbUrls.forEach((url) => URL.revokeObjectURL(url));
    screenshotThumbUrls = [];
    screenshotName.textContent = "";
    return;
  }
  selectedScreenshots.hidden = false;
  screenshotName.textContent = screenshotFiles.length === 1
    ? screenshotFiles[0].name
    : `${screenshotFiles.length} 张截图：${screenshotFiles.map((file) => file.name).join("、")}`;
  screenshotThumbUrls.forEach((url) => URL.revokeObjectURL(url));
  screenshotThumbUrls = screenshotFiles.map((file) => URL.createObjectURL(file));
  screenshotThumbs.innerHTML = screenshotFiles.map((file, index) => `
    <div class="thumb-card">
      <button class="thumb-image-button" data-preview-screenshot="${index}" type="button" aria-label="放大查看 ${file.name}">
        <img src="${screenshotThumbUrls[index]}" alt="${file.name}">
      </button>
      <div class="thumb-meta">
        <span title="${file.name}">${file.name}</span>
        <button class="thumb-remove" data-remove-screenshot="${index}" type="button" aria-label="移除 ${file.name}">移除</button>
      </div>
    </div>
  `).join("");
  screenshotThumbs.hidden = false;
}

function showImageLightbox(src, title = "图片预览") {
  if (!src) return;
  lightboxImage.src = src;
  lightboxImage.alt = title;
  lightboxTitle.textContent = title;
  imageLightbox.hidden = false;
}

function closeImageLightbox() {
  imageLightbox.hidden = true;
  lightboxImage.removeAttribute("src");
  lightboxImage.alt = "";
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
  longPlanPanel.open = true;
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
  document.querySelector(".long-plan-panel").open = true;
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
  setScreenshots(screenshotFiles);
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
  setOcrPreview(file);
  const formData = new FormData();
  formData.append("ocrImage", file);
  setLongPlanOcrState("busy", "正在识别截图，右侧可先查看原图。", button);
  message.textContent = "正在识别长期计划截图，识别结果会自动填入左侧表单。";
  const response = await fetch("/api/ocr-long-plan", {
    method: "POST",
    body: formData,
  });
  const payload = await response.json();
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

function renderTable(table, rows, grouped = false) {
  table.innerHTML = "";
  table.classList.toggle("is-highlightable", crossHighlight.checked);
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
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  table.appendChild(thead);
  table.appendChild(tbody);
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

function rowsToTsv(rows) {
  return rows.map((row) => row.map((value) => formatCell(value, row)).join("\t")).join("\n");
}

async function copyRows(key, includeHeader) {
  const rows = previewData[key] || [];
  const copyRows = includeHeader ? rows : rows.slice(1);
  if (!copyRows.length) return;
  await navigator.clipboard.writeText(rowsToTsv(copyRows));
  message.textContent = includeHeader ? "已复制全部表格。" : "已复制去表头数据。";
}

chooseButton.addEventListener("click", () => fileInput.click());
chooseScreenshotButton.addEventListener("click", () => screenshotInput.click());
pasteScreenshotButton.addEventListener("click", async () => {
  pastePanel.hidden = false;
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
        addScreenshots(files);
        message.textContent = "已直接读取剪切板截图。";
      }
    } catch (error) {
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
  setScreenshots(screenshotFiles);
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
  resetResult();
  setScreenshots(screenshotFiles);
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
  const idleText = installOcrButton.textContent;
  try {
    setLongPlanOcrState("busy", "正在安装截图识别组件，请不要关闭软件。", installOcrButton);
    installOcrButton.dataset.idleText = idleText;
    const response = await fetch("/api/install-ocr", { method: "POST" });
    const payload = await response.json();
    if (!response.ok || !payload.ok || !payload.installed) {
      throw new Error(payload.error || "截图识别组件安装失败。");
    }
    setInstallOcrVisible(false);
    setLongPlanOcrState("done", payload.message || "截图识别组件已安装，可以识别截图。");
    message.textContent = "截图识别组件已安装，可以识别长期计划截图。";
    await refreshOcrStatus();
  } catch (error) {
    const text = friendlyErrorText(error, "截图识别组件安装失败。");
    setLongPlanOcrState("error", text);
    message.textContent = text;
  } finally {
    installOcrButton.textContent = idleText;
  }
});

longPlanOcrInput.addEventListener("change", async () => {
  const [file] = longPlanOcrInput.files;
  longPlanOcrInput.value = "";
  if (!file) return;
  try {
    await recognizeLongPlanImage(file, ocrLongPlanButton);
  } catch (error) {
    const text = ocrErrorText(error);
    setLongPlanOcrState("error", text);
    message.textContent = text;
  }
});

ocrClipboardButton.addEventListener("click", async () => {
  try {
    setLongPlanOcrState("busy", "正在读取剪切板截图。", ocrClipboardButton);
    await recognizeLongPlanImage(await imageFromClipboard(), ocrClipboardButton);
  } catch (error) {
    const text = ocrErrorText(error);
    setLongPlanOcrState("error", text);
    message.textContent = text;
  }
});

function handlePaste(event) {
  const files = [...event.clipboardData.files].filter((file) => file.type.startsWith("image/"));
  if (!files.length) return;
  event.preventDefault();
  addScreenshots(files);
  pasteTarget.textContent = "截图已添加，可以继续粘贴下一张";
  message.textContent = "已从剪切板添加截图。";
}

pasteTarget.addEventListener("paste", handlePaste);
document.addEventListener("paste", (event) => {
  if (!pastePanel.hidden || event.target === document.body) {
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

  processButton.disabled = true;
  resetResult();
  message.textContent = "正在生成三表结构，请稍等。";

  try {
    const response = await fetch("/api/process", {
      method: "POST",
      body: formData,
    });
    const payload = await response.json();
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
    previewData = payload.previews || {};
    renderTable(sheet2Table, previewData.sheet2 || [], true);
    renderTable(sheet3Table, previewData.sheet3 || [], false);
    previewSection.hidden = false;
    const warnings = payload.warnings || [];
    if (warnings.length) {
      warningsList.innerHTML = warnings.map((warning) => `<li>${warning}</li>`).join("");
      warningsCard.hidden = false;
    }
    message.textContent = "XLSX 已生成，数据汇总已按当前面板设置处理。";
  } catch (error) {
    message.textContent = friendlyErrorText(error, "处理失败，请检查导入文件后重试。");
  } finally {
    processButton.disabled = !currentFiles.length;
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
});

document.addEventListener("click", (event) => {
  const cell = event.target.closest(".preview-section td");
  if (cell) {
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
      renderTable(sheet3Table, previewData.sheet3 || [], false);
    }
  });
});

refreshOcrStatus();
