const dropzone = document.querySelector("#dropzone");
const fileInput = document.querySelector("#fileInput");
const chooseButton = document.querySelector("#chooseButton");
const longPlanZone = document.querySelector("#longPlanZone");
const longPlanInput = document.querySelector("#longPlanInput");
const chooseLongPlanButton = document.querySelector("#chooseLongPlanButton");
const chooseScreenshotButton = document.querySelector("#chooseScreenshotButton");
const pasteScreenshotButton = document.querySelector("#pasteScreenshotButton");
const selectedFile = document.querySelector("#selectedFile");
const fileName = document.querySelector("#fileName");
const clearButton = document.querySelector("#clearButton");
const selectedLongPlans = document.querySelector("#selectedLongPlans");
const longPlanName = document.querySelector("#longPlanName");
const clearLongPlanButton = document.querySelector("#clearLongPlanButton");
const processButton = document.querySelector("#processButton");
const statusPill = document.querySelector("#statusPill");
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

let currentFiles = [];
let longPlanFiles = [];
let screenshotFiles = [];
let screenshotThumbUrls = [];
let previewData = {};

const moneyFormatter = new Intl.NumberFormat("zh-CN", {
  style: "currency",
  currency: "CNY",
  minimumFractionDigits: 2,
});

const numberFormatter = new Intl.NumberFormat("zh-CN");

function setStatus(text, mode = "") {
  statusPill.textContent = text;
  statusPill.className = `status-pill ${mode}`.trim();
}

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
    setStatus("等待上传");
    return;
  }
  selectedFile.hidden = false;
  fileName.textContent = currentFiles.length === 1
    ? currentFiles[0].name
    : `${currentFiles.length} 个 CSV：${currentFiles.map((file) => file.name).join("、")}`;
  processButton.disabled = false;
  message.textContent = "文件已就绪，可以生成 XLSX。";
  setStatus("已选择文件");
}

function setLongPlans(files) {
  longPlanFiles = [...files];
  resetResult();
  if (!longPlanFiles.length) {
    selectedLongPlans.hidden = true;
    longPlanName.textContent = "";
    return;
  }
  selectedLongPlans.hidden = false;
  longPlanName.textContent = longPlanFiles.length === 1
    ? longPlanFiles[0].name
    : `${longPlanFiles.length} 个长期计划：${longPlanFiles.map((file) => file.name).join("、")}`;
  message.textContent = "长期计划已添加，会和订单分析 CSV 一起汇总。";
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
      <img src="${screenshotThumbUrls[index]}" alt="${file.name}">
      <span title="${file.name}">${file.name}</span>
    </div>
  `).join("");
  screenshotThumbs.hidden = false;
}

function addScreenshots(files) {
  const images = [...files].filter((file) => file.type.startsWith("image/"));
  if (!images.length) return;
  screenshotFiles = [...screenshotFiles, ...images];
  setScreenshots(screenshotFiles);
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
chooseLongPlanButton.addEventListener("click", () => longPlanInput.click());
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
    } catch {
      // Some browsers require the user to paste manually into the focused box.
    }
  }
});

fileInput.addEventListener("change", () => {
  setFiles(fileInput.files);
});

longPlanInput.addEventListener("change", () => {
  setLongPlans(longPlanInput.files);
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
  longPlanInput.value = "";
  setLongPlans([]);
});

clearScreenshotButton.addEventListener("click", () => {
  screenshotInput.value = "";
  screenshotFiles = [];
  setScreenshots(screenshotFiles);
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
  if (files.some((file) => !file.name.toLowerCase().endsWith(".csv"))) {
    setStatus("文件格式不支持", "is-error");
    message.textContent = "拖入区域只接收 CSV 文件；消耗截图请用下方控件选择。";
    return;
  }
  fileInput.files = event.dataTransfer.files;
  setFiles(files);
});

for (const eventName of ["dragenter", "dragover"]) {
  longPlanZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    longPlanZone.classList.add("is-dragging");
  });
}

for (const eventName of ["dragleave", "drop"]) {
  longPlanZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    longPlanZone.classList.remove("is-dragging");
  });
}

longPlanZone.addEventListener("drop", (event) => {
  const files = [...event.dataTransfer.files];
  const supported = [".csv", ".xlsx", ".xlsm"];
  if (files.some((file) => !supported.some((suffix) => file.name.toLowerCase().endsWith(suffix)))) {
    setStatus("文件格式不支持", "is-error");
    message.textContent = "长期计划只支持 CSV 或 XLSX 文件。";
    return;
  }
  longPlanInput.files = event.dataTransfer.files;
  setLongPlans(files);
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
  longPlanFiles.forEach((file) => formData.append("longPlans", file));
  screenshotFiles.forEach((file) => formData.append("screenshots", file));
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
  setStatus("正在处理", "is-working");
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
    setStatus("处理完成", "is-done");
    message.textContent = "XLSX 已生成，数据汇总已按当前面板设置处理。";
  } catch (error) {
    setStatus("处理失败", "is-error");
    message.textContent = error.message;
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
    message.textContent = "当前浏览器不允许直接复制图片，请使用下载拼接图。";
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
    message.textContent = error.message || "复制失败。";
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
