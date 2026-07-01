const ORDER_DEFAULTS = {
  dataSheet: "订单表",
  firstDataRow: 12,
  lastDataRow: 28,
  specOd: 2,
  tubes: 1,
  purification: "RPC",
  delivery: "液体",
};

const primerInput = document.querySelector("#primerInput");
const orderInput = document.querySelector("#orderInput");
const generateButton = document.querySelector("#generateButton");
const primerFileName = document.querySelector("#primerFileName");
const orderFileName = document.querySelector("#orderFileName");
const primerMeta = document.querySelector("#primerMeta");
const orderMeta = document.querySelector("#orderMeta");
const primerPreview = document.querySelector("#primerPreview");
const primerCount = document.querySelector("#primerCount");
const orderStatus = document.querySelector("#orderStatus");
const outputName = document.querySelector("#outputName");
const templateFact = document.querySelector("#templateFact");
const dateFact = document.querySelector("#dateFact");
const todayText = document.querySelector("#todayText");
const statusStrip = document.querySelector(".status-strip");
const statusText = document.querySelector("#statusText");
const downloadLink = document.querySelector("#downloadLink");

let primerFile = null;
let orderFile = null;
let parsedPrimers = [];
let lastDownloadUrl = null;

function normalizeSequence(rawSequence) {
  return rawSequence.replace(/\s+/g, "");
}

function parsePrimerText(text) {
  const primers = [];
  for (const [lineIndex, originalLine] of text.split(/\r?\n/).entries()) {
    const line = originalLine.trim();
    if (!line) continue;

    let name = "";
    let sequence = "";
    const tabParts = line.split(/\t+/);

    if (tabParts.length >= 2) {
      name = tabParts[0].trim();
      sequence = normalizeSequence(tabParts.slice(1).join(""));
    } else {
      const match = line.match(/^(.+?)\s{2,}([A-Za-z.\-*\s]+)$/);
      if (match) {
        name = match[1].trim();
        sequence = normalizeSequence(match[2]);
      }
    }

    if (!name || !sequence) {
      throw new Error(`第 ${lineIndex + 1} 行无法识别。`);
    }

    primers.push({ name, sequence });
  }

  if (primers.length === 0) {
    throw new Error("没有识别到引物。");
  }

  return primers;
}

function decodePrimerBuffer(buffer) {
  const bytes = new Uint8Array(buffer);
  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) {
    return new TextDecoder("utf-16le").decode(bytes.slice(2));
  }
  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) {
    const swapped = new Uint8Array(bytes.length - 2);
    for (let index = 2; index < bytes.length; index += 2) {
      swapped[index - 2] = bytes[index + 1];
      swapped[index - 1] = bytes[index];
    }
    return new TextDecoder("utf-16le").decode(swapped);
  }
  return new TextDecoder("utf-8").decode(bytes).replace(/^\uFEFF/, "");
}

function shanghaiDateParts(date = new Date()) {
  const parts = new Intl.DateTimeFormat("zh-CN", {
    timeZone: "Asia/Shanghai",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(date);
  return Object.fromEntries(parts.map((part) => [part.type, part.value]));
}

function dateTokens(date = new Date()) {
  const { year, month, day } = shanghaiDateParts(date);
  return {
    compact: `${year}${month}${day}`,
    dashed: `${year}-${month}-${day}`,
    dotted: `${year}.${month}.${day}`,
    underscored: `${year}_${month}_${day}`,
  };
}

function makeDatedOutputName(templateFileName, date = new Date()) {
  const safeName = templateFileName?.trim() || "引物合成订购单.xlsx";
  const extensionMatch = safeName.match(/(\.[^.]+)$/);
  const extension = extensionMatch?.[1] || ".xlsx";
  const baseName = extensionMatch ? safeName.slice(0, -extension.length) : safeName;
  const tokens = dateTokens(date);
  const datePattern = /20\d{2}([-_.]?)\d{2}\1\d{2}/g;
  let match;
  let lastMatch = null;

  while ((match = datePattern.exec(baseName)) !== null) {
    lastMatch = match;
  }

  if (lastMatch) {
    const replacement =
      lastMatch[1] === "-"
        ? tokens.dashed
        : lastMatch[1] === "."
          ? tokens.dotted
          : lastMatch[1] === "_"
            ? tokens.underscored
            : tokens.compact;
    return `${baseName.slice(0, lastMatch.index)}${replacement}${baseName.slice(
      lastMatch.index + lastMatch[0].length,
    )}${extension}`;
  }

  return `${baseName}-${tokens.compact}${extension}`;
}

function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

function setStatus(message, type = "") {
  statusText.textContent = message;
  statusStrip.classList.toggle("error", type === "error");
  statusStrip.classList.toggle("warn", type === "warn");
}

function updateGenerateState() {
  generateButton.disabled = !(primerFile && orderFile && parsedPrimers.length > 0);
  if (orderFile) {
    const projectedName = makeDatedOutputName(orderFile.name);
    outputName.textContent = projectedName;
    templateFact.textContent = orderFile.name;
    orderStatus.textContent = "已选择";
  } else {
    outputName.textContent = "等待订单表格";
    templateFact.textContent = "-";
    orderStatus.textContent = "未就绪";
  }
}

function renderPrimerPreview(primers) {
  primerPreview.innerHTML = "";
  const visibleRows = primers.slice(0, 80);

  for (const primer of visibleRows) {
    const row = document.createElement("tr");
    const nameCell = document.createElement("td");
    const lengthCell = document.createElement("td");
    nameCell.textContent = primer.name;
    lengthCell.textContent = String(primer.sequence.length);
    row.append(nameCell, lengthCell);
    primerPreview.append(row);
  }

  if (primers.length > visibleRows.length) {
    const row = document.createElement("tr");
    const cell = document.createElement("td");
    cell.colSpan = 2;
    cell.className = "empty-cell";
    cell.textContent = `还有 ${primers.length - visibleRows.length} 条未显示`;
    row.append(cell);
    primerPreview.append(row);
  }

  primerCount.textContent = `${primers.length} 条`;
}

function clearDownload() {
  if (lastDownloadUrl) {
    URL.revokeObjectURL(lastDownloadUrl);
    lastDownloadUrl = null;
  }
  downloadLink.hidden = true;
  downloadLink.href = "#";
}

async function handlePrimerFile(file) {
  clearDownload();
  primerFile = file;
  primerFileName.textContent = file.name;
  primerMeta.textContent = formatBytes(file.size);

  try {
    const text = decodePrimerBuffer(await file.arrayBuffer());
    parsedPrimers = parsePrimerText(text);
    renderPrimerPreview(parsedPrimers);
    setStatus("引物文件已读取。");
  } catch (error) {
    parsedPrimers = [];
    primerCount.textContent = "0 条";
    primerPreview.innerHTML = '<tr><td colspan="2" class="empty-cell">解析失败</td></tr>';
    setStatus(error.message, "error");
  }

  updateGenerateState();
}

function handleOrderFile(file) {
  clearDownload();
  orderFile = file;
  orderFileName.textContent = file.name;
  orderMeta.textContent = formatBytes(file.size);
  setStatus("订单表格已选择。");
  updateGenerateState();
}

function setupDropWindow(windowElement, inputElement, onFile) {
  windowElement.querySelector(".drop-target").addEventListener("click", () => inputElement.click());
  inputElement.addEventListener("change", () => {
    const file = inputElement.files?.[0];
    if (file) onFile(file);
  });

  for (const eventName of ["dragenter", "dragover"]) {
    windowElement.addEventListener(eventName, (event) => {
      event.preventDefault();
      windowElement.classList.add("dragging");
    });
  }

  for (const eventName of ["dragleave", "drop"]) {
    windowElement.addEventListener(eventName, (event) => {
      event.preventDefault();
      windowElement.classList.remove("dragging");
    });
  }

  windowElement.addEventListener("drop", (event) => {
    const file = event.dataTransfer.files?.[0];
    if (file) onFile(file);
  });
}

function getRequiredWorksheet(workbook, sheetName) {
  const worksheet = workbook.getWorksheet(sheetName);
  if (!worksheet) {
    throw new Error(`订单表格中没有找到工作表“${sheetName}”。`);
  }
  return worksheet;
}

function clearCell(cell) {
  cell.value = null;
}

function fillOrderWorkbook(workbook, primers, options = {}) {
  const config = { ...ORDER_DEFAULTS, ...options };
  const worksheet = getRequiredWorksheet(workbook, config.dataSheet);
  const capacity = config.lastDataRow - config.firstDataRow + 1;

  if (primers.length > capacity) {
    throw new Error(`订单表当前只有 ${capacity} 行可填，引物有 ${primers.length} 条。`);
  }

  for (let index = 0; index < capacity; index += 1) {
    const rowNumber = config.firstDataRow + index;
    const row = worksheet.getRow(rowNumber);
    const primer = primers[index];

    row.getCell(1).value = index + 1;
    row.getCell(4).value = {
      formula: `LEN(SUBSTITUTE(C${rowNumber}," ",""))`,
      result: primer ? primer.sequence.length : 0,
    };

    if (primer) {
      row.getCell(2).value = primer.name;
      row.getCell(3).value = primer.sequence;
      row.getCell(5).value = config.specOd;
      row.getCell(6).value = config.tubes;
      row.getCell(7).value = config.purification;
      clearCell(row.getCell(8));
      row.getCell(9).value = config.delivery;
      clearCell(row.getCell(10));
    } else {
      for (const column of [2, 3, 5, 6, 7, 8, 9, 10]) {
        clearCell(row.getCell(column));
      }
    }
  }
}

async function generateWorkbook() {
  if (!window.ExcelJS) {
    throw new Error("Excel 处理库未加载，请刷新页面后重试。");
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await orderFile.arrayBuffer());
  fillOrderWorkbook(workbook, parsedPrimers);
  const outputBuffer = await workbook.xlsx.writeBuffer();
  const outputFileName = makeDatedOutputName(orderFile.name);
  return {
    outputFileName,
    blob: new Blob([outputBuffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }),
  };
}

generateButton.addEventListener("click", async () => {
  if (!primerFile || !orderFile) return;

  clearDownload();
  generateButton.disabled = true;
  setStatus("正在生成订单表...");

  try {
    const { outputFileName, blob } = await generateWorkbook();
    lastDownloadUrl = URL.createObjectURL(blob);
    downloadLink.href = lastDownloadUrl;
    downloadLink.download = outputFileName;
    downloadLink.hidden = false;
    downloadLink.click();
    outputName.textContent = outputFileName;
    setStatus(`已生成 ${outputFileName}`);
  } catch (error) {
    setStatus(error.message || "生成失败。", "error");
  } finally {
    updateGenerateState();
  }
});

setupDropWindow(document.querySelector("#primerDrop"), primerInput, handlePrimerFile);
setupDropWindow(document.querySelector("#orderDrop"), orderInput, handleOrderFile);

const tokens = dateTokens();
todayText.textContent = `今日版本 ${tokens.compact}`;
dateFact.textContent = tokens.compact;
