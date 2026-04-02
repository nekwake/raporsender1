const targetPathLabel = document.getElementById("targetPathLabel");
const pickTargetBtn = document.getElementById("pickTarget");
const dropZone = document.getElementById("dropZone");
const dropCard = document.getElementById("dropCard");
const previewCard = document.getElementById("previewCard");
const previewMeta = document.getElementById("previewMeta");
const previewTableWrap = document.getElementById("previewTableWrap");
const clearNowBtn = document.getElementById("clearNow");
const sendNowBtn = document.getElementById("sendNow");
const infoNowBtn = document.getElementById("infoNow");
const counter = document.getElementById("counter");
const statusBox = document.getElementById("statusBox");
const sentList = document.getElementById("sentList");

let currentGrid = null;
let currentMeta = {
  suggestedBaseName: "paste",
  sourceType: "paste",
};

const UI_TEXT = {
  statusReady: "Hazır.",
  targetNotSelected: "Seçilmedi",
  targetChoosing: "Hedef klasör seçiliyor...",
  targetUpdated: "Hedef klasör güncellendi.",
  pasteCleaning: "Yapıştırılıyor ve temizleniyor...",
  dropReading: "Dosya okunuyor...",
  gridEmpty: "Yapıştırılan/dışarıdan gelen veri boş ya da temizlenemedi.",
  previewReady: "Önizleme hazır. Gönder'e basabilirsiniz.",
  cleared: "Temizlendi. Yeniden bir şey yapıştırın/sürükleyin.",
  sendNoData: "Göndermek için önce veri yükleyin (Ctrl+V veya dosya bırakın).",
  sending: "Gönderiliyor...",
  deduped: "Bu içerik daha önce gönderilmiş. Tekrar gönderilmedi.",
  sendSuccess: "Tamamlandı. Kaydedildi: ",
  maxFileTooLarge: "Dosya çok büyük (25MB limiti).",
  errorPrefix: "Hata: ",
};

function setStatus(text) {
  statusBox.textContent = text;
}

function formatDate(iso) {
  return new Date(iso).toLocaleString("tr-TR");
}

function emptyCellToString(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function cleanGrid(grid) {
  if (!Array.isArray(grid)) return [];

  // Normalize rows into trimmed string cells.
  const normalized = grid.map((row) =>
    Array.isArray(row) ? row.map((c) => emptyCellToString(c)) : [emptyCellToString(row)]
  );

  // Determine max column count (ragged rows are allowed).
  const maxCols = normalized.reduce((acc, row) => Math.max(acc, row.length), 0);

  const rowsWithPads = normalized.map((row) => {
    const padded = row.slice();
    while (padded.length < maxCols) padded.push("");
    return padded;
  });

  // Remove completely empty rows.
  const nonEmptyRows = rowsWithPads.filter((row) => row.some((c) => c !== ""));

  if (!nonEmptyRows.length) return [];

  // Remove completely empty columns.
  const colCount = maxCols;
  const colKeep = new Array(colCount).fill(false).map((_, colIdx) => {
    return nonEmptyRows.some((row) => row[colIdx] !== "");
  });

  const keptColIndices = colKeep
    .map((keep, idx) => (keep ? idx : -1))
    .filter((idx) => idx !== -1);

  const cleaned = nonEmptyRows.map((row) => keptColIndices.map((colIdx) => row[colIdx]));
  return cleaned;
}

function renderPreviewTable(grid) {
  previewTableWrap.innerHTML = "";

  if (!grid || !grid.length) {
    previewCard.classList.add("hidden");
    return;
  }

  previewCard.classList.remove("hidden");

  const table = document.createElement("table");
  table.className = "tablePreview";

  const maxRowsForUI = 80;
  const maxColsForUI = 20;
  const rows = grid.slice(0, maxRowsForUI);
  const colsCount = Math.min(
    maxColsForUI,
    rows.reduce((acc, row) => Math.max(acc, row.length), 0)
  );

  for (let r = 0; r < rows.length; r++) {
    const tr = document.createElement("tr");
    const row = rows[r];
    for (let c = 0; c < colsCount; c++) {
      const td = document.createElement("td");
      const val = row[c] ?? "";
      td.textContent = val;
      tr.appendChild(td);
    }
    table.appendChild(tr);
  }

  previewTableWrap.appendChild(table);
}

function updateSendEnabled() {
  const hasTarget = Boolean(targetPathLabel.dataset.hasTarget === "1");
  const hasGrid = Boolean(currentGrid && currentGrid.length);
  sendNowBtn.disabled = !(hasTarget && hasGrid);
}

function updateCounter(config) {
  counter.textContent = `Kaydedilen gönderim: ${config.sentCount}`;
}

function parseTsv(text) {
  const lines = text.split(/\r?\n/).map((l) => l.trimEnd());

  return lines.map((line) => line.split("\t"));
}

function parseHtmlTable(html) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, "text/html");
  const table = doc.querySelector("table");
  if (!table) return null;

  const rows = [];
  const htmlRows = Array.from(table.rows);
  for (const tr of htmlRows) {
    const cells = Array.from(tr.cells);
    const row = cells.map((cell) => cell.textContent.trim());
    rows.push(row);
  }

  return rows.length ? rows : null;
}

function getClipboardGridFromPasteEvent(event) {
  const html = event.clipboardData.getData("text/html");
  if (html && html.toLowerCase().includes("<table")) {
    const grid = parseHtmlTable(html);
    if (grid) return { grid, sourceType: "clipboard-excel" };
  }

  const text = event.clipboardData.getData("text/plain");
  if (text && text.trim().length > 0) {
    return { grid: parseTsv(text), sourceType: "clipboard-text" };
  }

  return { grid: null, sourceType: "unknown" };
}

function bufferToBase64(arrayBuffer) {
  const bytes = new Uint8Array(arrayBuffer);
  const chunkSize = 0x8000;
  let binary = "";
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}

async function handleParsedGrid(grid, meta) {
  const cleaned = cleanGrid(grid);
  if (!cleaned.length) {
    currentGrid = null;
    previewCard.classList.add("hidden");
    previewMeta.textContent = "";
    setStatus(UI_TEXT.gridEmpty);
    updateSendEnabled();
    return;
  }

  currentGrid = cleaned;
  currentMeta = {
    suggestedBaseName: meta?.suggestedBaseName || "paste",
    sourceType: meta?.sourceType || "paste",
  };
  previewMeta.textContent = `Boyut: ${cleaned.length} satır x ${cleaned[0].length} sütun`;
  renderPreviewTable(cleaned);
  setStatus(UI_TEXT.previewReady);
  updateSendEnabled();
}

async function parseDroppedFile(file) {
  const maxBytes = 25 * 1024 * 1024; // 25MB
  if (file.size > maxBytes) {
    throw new Error(UI_TEXT.maxFileTooLarge);
  }

  const arrayBuffer = await file.arrayBuffer();
  const bufferBase64 = bufferToBase64(arrayBuffer);

  const parsed = await window.bridgeApi.parseFileBuffer({
    fileName: file.name,
    mimeType: file.type || "",
    bufferBase64,
  });

  return parsed;
}

async function handleDropEvent(event) {
  event.preventDefault();
  dropZone.classList.remove("dragOver");

  const files = Array.from(event.dataTransfer.files || []);
  if (!files.length) return;

  setStatus(UI_TEXT.dropReading);
  try {
    const file = files[0];
    const parsed = await parseDroppedFile(file);
    await handleParsedGrid(parsed.grid, {
      suggestedBaseName: parsed.suggestedBaseName,
      sourceType: parsed.sourceType,
    });
  } catch (err) {
    setStatus(`${UI_TEXT.errorPrefix}${err.message}`);
  }
}

async function init() {
  const config = await window.bridgeApi.getConfig();
  if (config.hasTarget) {
    targetPathLabel.textContent = config.targetPath;
    targetPathLabel.dataset.hasTarget = "1";
  } else {
    targetPathLabel.textContent = UI_TEXT.targetNotSelected;
    targetPathLabel.dataset.hasTarget = "0";
  }
  updateCounter(config);
  updateSendEnabled();

  const last = await window.bridgeApi.getLastSent();
  sentList.innerHTML = "";
  if (!last.length) {
    const li = document.createElement("li");
    li.textContent = "Henüz gönderim yok.";
    sentList.appendChild(li);
  } else {
    for (const item of last) {
      const li = document.createElement("li");
      li.textContent = `${item.name} - ${formatDate(item.sentAt)}`;
      sentList.appendChild(li);
    }
  }
}

pickTargetBtn.addEventListener("click", async () => {
  try {
    setStatus(UI_TEXT.targetChoosing);
    const selected = await window.bridgeApi.chooseTargetFolder();
    if (!selected) return;
    const config = await window.bridgeApi.getConfig();
    targetPathLabel.textContent = config.targetPath;
    targetPathLabel.dataset.hasTarget = "1";
    updateCounter(config);
    updateSendEnabled();
    setStatus(UI_TEXT.targetUpdated);
  } catch (err) {
    setStatus(`${UI_TEXT.errorPrefix}${err.message}`);
  }
});

dropZone.addEventListener("dragenter", (e) => {
  e.preventDefault();
  dropZone.classList.add("dragOver");
});

dropZone.addEventListener("dragover", (e) => {
  e.preventDefault();
  dropZone.classList.add("dragOver");
});

dropZone.addEventListener("dragleave", (e) => {
  e.preventDefault();
  dropZone.classList.remove("dragOver");
});

dropZone.addEventListener("drop", handleDropEvent);

// Paste (Ctrl+V) parsing.
window.addEventListener(
  "paste",
  async (event) => {
    try {
      // We only handle paste if it looks like table-ish data.
      const { grid, sourceType } = getClipboardGridFromPasteEvent(event);
      if (!grid) return;

      // Prevent default so text does not land in an input.
      event.preventDefault();
      setStatus(UI_TEXT.pasteCleaning);

      await handleParsedGrid(grid, {
        suggestedBaseName: "paste",
        sourceType,
      });
    } catch (err) {
      setStatus(`Hata: ${err.message}`);
    }
  },
  true
);

clearNowBtn.addEventListener("click", () => {
  currentGrid = null;
  previewMeta.textContent = "";
  previewTableWrap.innerHTML = "";
  previewCard.classList.add("hidden");
  setStatus("Temizlendi. Yeniden bir şey yapıştırın/sürükleyin.");
  updateSendEnabled();
});

sendNowBtn.addEventListener("click", async () => {
  if (!currentGrid || !currentGrid.length) {
    setStatus(UI_TEXT.sendNoData);
    return;
  }
  setStatus(UI_TEXT.sending);
  try {
    const result = await window.bridgeApi.sendGrid({
      grid: currentGrid,
      suggestedBaseName: currentMeta.suggestedBaseName,
    });

    if (result.deduped) {
      setStatus(UI_TEXT.deduped);
    } else {
      setStatus(`${UI_TEXT.sendSuccess}${result.fileName}`);
    }
    const config = await window.bridgeApi.getConfig();
    updateCounter(config);
    updateSendEnabled();

    const last = await window.bridgeApi.getLastSent();
    sentList.innerHTML = "";
    for (const item of last) {
      const li = document.createElement("li");
      li.textContent = `${item.name} - ${formatDate(item.sentAt)}`;
      sentList.appendChild(li);
    }
  } catch (err) {
    setStatus(`${UI_TEXT.errorPrefix}${err.message}`);
  }
});

window.bridgeApi.onStatus((payload) => {
  if (payload?.message) setStatus(payload.message);
});

init();
