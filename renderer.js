/**
 * Assistant — renderer: home mode selection, monthly stock, audit form, generic "other" paste flow.
 */

const topbarSubtitle = document.getElementById("topbarSubtitle");
const targetBar = document.getElementById("targetBar");

const viewHome = document.getElementById("viewHome");
const viewStock = document.getElementById("viewStock");
const viewAudit = document.getElementById("viewAudit");
const viewOther = document.getElementById("viewOther");

const btnModeStock = document.getElementById("btnModeStock");
const btnModeAudit = document.getElementById("btnModeAudit");
const btnModeOther = document.getElementById("btnModeOther");
const backFromStock = document.getElementById("backFromStock");
const backFromAudit = document.getElementById("backFromAudit");
const backFromOther = document.getElementById("backFromOther");

const dropZoneStock = document.getElementById("dropZoneStock");
const stockMonthSelect = document.getElementById("stockMonthSelect");
const saveStockDesktop = document.getElementById("saveStockDesktop");
const clearStock = document.getElementById("clearStock");
const stockPreviewCard = document.getElementById("stockPreviewCard");
const stockPreviewMeta = document.getElementById("stockPreviewMeta");
const stockPreviewWrap = document.getElementById("stockPreviewWrap");
const stockFileHint = document.getElementById("stockFileHint");
const statusBoxStock = document.getElementById("statusBoxStock");

const auditFormMount = document.getElementById("auditFormMount");
const auditBranchName = document.getElementById("auditBranchName");
const auditBranchManager = document.getElementById("auditBranchManager");
const auditRegionalManager = document.getElementById("auditRegionalManager");
const auditPreparedBy = document.getElementById("auditPreparedBy");
const auditVisitDayManager = document.getElementById("auditVisitDayManager");
const auditReportDate = document.getElementById("auditReportDate");
const auditSumKalite = document.getElementById("auditSumKalite");
const auditSumServis = document.getElementById("auditSumServis");
const auditSumTemizlik = document.getElementById("auditSumTemizlik");
const auditPercent = document.getElementById("auditPercent");
const auditLetter = document.getElementById("auditLetter");
const saveAuditDesktop = document.getElementById("saveAuditDesktop");
const statusBoxAudit = document.getElementById("statusBoxAudit");

const targetPathLabel = document.getElementById("targetPathLabel");
const pickTargetBtn = document.getElementById("pickTarget");
const dropZone = document.getElementById("dropZone");
const previewCard = document.getElementById("previewCard");
const previewMeta = document.getElementById("previewMeta");
const previewTableWrap = document.getElementById("previewTableWrap");
const clearNowBtn = document.getElementById("clearNow");
const sendNowBtn = document.getElementById("sendNow");
const counter = document.getElementById("counter");
const statusBox = document.getElementById("statusBox");
const sentList = document.getElementById("sentList");

let activeView = "home";

/** @type {string[][] | null} */
let stockGrid = null;

let currentGrid = null;
let currentMeta = {
  suggestedBaseName: "paste",
  sourceType: "paste",
};

const MONTHLY_STOCK_HEADERS = [
  "STOK ADI",
  "BİRİMİ",
  "BİRİM FİYAT",
  "GÜN AÇILIŞ",
  "GİRİŞ",
  "ÇIKIŞ",
  "İDEAL AÇILIŞ",
  "SATIŞLAR",
  "İPTAL",
  "PERSONEL",
  "ÖDENMEZ",
  "ATILAN",
  "SAYIM AÇIK",
  "SAYIM FAZLA",
  "OLMASI GEREKEN",
  "GERÇEK KAPANIŞ",
  "FARK",
  "TUTAR FARK",
];

const TURKISH_MONTHS = [
  "Ocak",
  "Şubat",
  "Mart",
  "Nisan",
  "Mayıs",
  "Haziran",
  "Temmuz",
  "Ağustos",
  "Eylül",
  "Ekim",
  "Kasım",
  "Aralık",
];

const MAX_QUALITY = 46;
const MAX_SERVICE = 56;
const MAX_CLEAN = 40;
const MAX_TOTAL = MAX_QUALITY + MAX_SERVICE + MAX_CLEAN;

const UI_TEXT = {
  statusReady: "Hazır.",
  targetNotSelected: "Seçilmedi",
  targetChoosing: "Hedef klasör seçiliyor...",
  targetUpdated: "Hedef klasör güncellendi.",
  pasteCleaning: "Yapıştırılıyor...",
  dropReading: "Dosya okunuyor...",
  gridEmpty: "Yapıştırılan/dışarıdan gelen veri boş.",
  previewReady: "Önizleme hazır. Gönder'e basabilirsiniz.",
  cleared: "Temizlendi.",
  sendNoData: "Göndermek için önce veri yükleyin.",
  sending: "Gönderiliyor...",
  deduped: "Bu içerik daha önce gönderilmiş.",
  sendSuccess: "Tamamlandı. Kaydedildi: ",
  maxFileTooLarge: "Dosya çok büyük (25MB limiti).",
  errorPrefix: "Hata: ",
};

function setStatus(text) {
  statusBox.textContent = text;
}

function setStatusStock(text) {
  statusBoxStock.textContent = text;
}

function setStatusAudit(text) {
  statusBoxAudit.textContent = text;
}

function formatDate(iso) {
  return new Date(iso).toLocaleString("tr-TR");
}

function emptyCellToString(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function showView(view) {
  activeView = view;
  viewHome.classList.toggle("hidden", view !== "home");
  viewStock.classList.toggle("hidden", view !== "stock");
  viewAudit.classList.toggle("hidden", view !== "audit");
  viewOther.classList.toggle("hidden", view !== "other");

  targetBar.classList.toggle("hidden", view !== "other");

  if (view === "home") {
    topbarSubtitle.textContent = "Belge tipi seçin";
  } else if (view === "stock") {
    topbarSubtitle.textContent = "Aylık stok kapanış";
  } else if (view === "audit") {
    topbarSubtitle.textContent = "Operasyon denetim raporu";
  } else {
    topbarSubtitle.textContent = "Diğer — tablo aktarım";
  }
}

function cleanGrid(grid) {
  if (!Array.isArray(grid)) return [];

  const normalized = grid.map((row) =>
    Array.isArray(row) ? row.map((c) => emptyCellToString(c)) : [emptyCellToString(row)]
  );

  const maxCols = normalized.reduce((acc, row) => Math.max(acc, row.length), 0);
  const rowsWithPads = normalized.map((row) => {
    const padded = row.slice();
    while (padded.length < maxCols) padded.push("");
    return padded;
  });

  const nonEmptyRows = rowsWithPads.filter((row) => row.some((c) => c !== ""));
  if (!nonEmptyRows.length) return [];

  const colKeep = new Array(maxCols).fill(false).map((_, colIdx) =>
    nonEmptyRows.some((row) => row[colIdx] !== "")
  );
  const keptColIndices = colKeep
    .map((keep, idx) => (keep ? idx : -1))
    .filter((idx) => idx !== -1);

  return nonEmptyRows.map((row) => keptColIndices.map((colIdx) => row[colIdx]));
}

/** Trim and pad ragged rows; do not drop empty rows/columns */
function gridToRowsRaw(grid) {
  if (!Array.isArray(grid)) return [];
  const normalized = grid.map((row) =>
    Array.isArray(row) ? row.map((c) => emptyCellToString(c)) : [emptyCellToString(row)]
  );
  const maxCols = normalized.reduce((acc, row) => Math.max(acc, row.length), 0);
  return normalized.map((row) => {
    const padded = row.slice();
    while (padded.length < maxCols) padded.push("");
    return padded;
  });
}

/** If column 0 is empty for every row, remove it */
function removeFirstColumnIfAllEmpty(rows) {
  if (!rows.length) return rows;
  const allEmptyFirst = rows.every((row) => !row[0] || String(row[0]).trim() === "");
  if (!allEmptyFirst) return rows;
  return rows.map((row) => row.slice(1));
}

/** Excel son satırda boş satır bırakabiliyor; tamamı boş satırları at. */
function dropCompletelyEmptyRows(rows) {
  return rows.filter((row) => row.some((c) => c !== ""));
}

function prependMonthlyHeaders(rows) {
  return [MONTHLY_STOCK_HEADERS.slice(), ...rows];
}

function processMonthlyStockRaw(rawGrid) {
  const padded = gridToRowsRaw(rawGrid);
  if (!padded.length) return [];
  const noLeadingEmptyCol = removeFirstColumnIfAllEmpty(padded);
  const dataOnly = dropCompletelyEmptyRows(noLeadingEmptyCol);
  if (!dataOnly.length) return [];
  return prependMonthlyHeaders(dataOnly);
}

function renderPreviewTable(grid, container) {
  container.innerHTML = "";
  if (!grid || !grid.length) return;

  const table = document.createElement("table");
  table.className = "tablePreview";

  const maxRowsForUI = 80;
  const maxColsForUI = 24;
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
      td.textContent = row[c] ?? "";
      if (r === 0) td.style.fontWeight = "700";
      tr.appendChild(td);
    }
    table.appendChild(tr);
  }

  container.appendChild(table);
}

function updateSendEnabled() {
  const hasTarget = targetPathLabel.dataset.hasTarget === "1";
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
  for (const tr of Array.from(table.rows)) {
    const cells = Array.from(tr.cells);
    rows.push(cells.map((cell) => cell.textContent.trim()));
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
  renderPreviewTable(cleaned, previewTableWrap);
  previewCard.classList.remove("hidden");
  setStatus(UI_TEXT.previewReady);
  updateSendEnabled();
}

async function handleParsedStockGrid(rawGrid) {
  const processed = processMonthlyStockRaw(rawGrid);
  if (!processed.length || processed.length < 2) {
    stockGrid = null;
    stockPreviewCard.classList.add("hidden");
    stockPreviewMeta.textContent = "";
    setStatusStock(UI_TEXT.gridEmpty);
    return;
  }

  stockGrid = processed;
  const dataRows = processed.length - 1;
  const cols = processed[0].length;
  stockPreviewMeta.textContent = `Başlık + ${dataRows} veri satırı, ${cols} sütun`;
  renderPreviewTable(processed, stockPreviewWrap);
  stockPreviewCard.classList.remove("hidden");
  setStatusStock("Tablo hazır. Ay seçip masaüstüne kaydedebilirsiniz.");
  updateStockFileHint();
}

async function parseDroppedFile(file) {
  const maxBytes = 25 * 1024 * 1024;
  if (file.size > maxBytes) {
    throw new Error(UI_TEXT.maxFileTooLarge);
  }
  const arrayBuffer = await file.arrayBuffer();
  const bufferBase64 = bufferToBase64(arrayBuffer);
  return window.bridgeApi.parseFileBuffer({
    fileName: file.name,
    mimeType: file.type || "",
    bufferBase64,
  });
}

function getSelectedMonthUpperLabel() {
  const idx = parseInt(stockMonthSelect.value, 10);
  if (Number.isNaN(idx) || idx < 0 || idx > 11) return "RAPOR";
  return TURKISH_MONTHS[idx].toLocaleUpperCase("tr-TR");
}

function updateStockFileHint() {
  const label = getSelectedMonthUpperLabel();
  stockFileHint.textContent = `Dosya adı: ${label} STOK KAPANIŞ.xlsx`;
}

function getAuditFormDef() {
  return window.OPERATION_AUDIT_FORM;
}

function gradeFromPercent(p) {
  if (p >= 93) return "A";
  if (p >= 85) return "B";
  if (p >= 75) return "C";
  return "F";
}

function auditEarnedId(sectionKey, index) {
  return `audit-earned-${sectionKey}-${index}`;
}

function auditCommentId(sectionKey, index) {
  return `audit-comment-${sectionKey}-${index}`;
}

function collectAuditAll() {
  const form = getAuditFormDef();
  if (!form) {
    return { detailRows: [], errors: ["Form tanımı yüklenemedi."], sums: { kalite: 0, servis: 0, temizlik: 0 } };
  }

  const detailRows = [];
  const errors = [];
  const sums = { kalite: 0, servis: 0, temizlik: 0 };

  let globalNo = 1;
  for (const section of form.sections) {
    section.items.forEach((item, idx) => {
      const earnedEl = document.getElementById(auditEarnedId(section.key, idx));
      const commentEl = document.getElementById(auditCommentId(section.key, idx));
      const raw = (earnedEl && earnedEl.value !== "" ? earnedEl.value : "0").replace(",", ".");
      const earned = Number(raw);
      const comment =
        commentEl && commentEl.value ? String(commentEl.value).trim() : "";

      if (!Number.isFinite(earned)) {
        errors.push(`${section.title} #${globalNo}: geçersiz puan`);
        globalNo += 1;
        return;
      }
      if (earned < 0 || earned > item.max) {
        errors.push(
          `${section.title} #${globalNo}: puan 0–${item.max} olmalı (girilen: ${earned})`
        );
        globalNo += 1;
        return;
      }

      sums[section.key] = (sums[section.key] || 0) + earned;
      detailRows.push([section.title, globalNo, item.text, item.max, earned, comment]);
      globalNo += 1;
    });
  }

  return { detailRows, errors, sums };
}

function updateAuditSummaryFromSums(sums, errors) {
  const k = sums.kalite ?? 0;
  const s = sums.servis ?? 0;
  const t = sums.temizlik ?? 0;
  const totalEarned = k + s + t;

  auditSumKalite.textContent = `${k} / ${MAX_QUALITY}`;
  auditSumServis.textContent = `${s} / ${MAX_SERVICE}`;
  auditSumTemizlik.textContent = `${t} / ${MAX_CLEAN}`;

  if (errors.length) {
    auditPercent.textContent = `Genel yüzde: — (${errors.length} satır hatalı)`;
    auditLetter.textContent = "—";
    setStatusAudit(errors[0]);
    return null;
  }

  const percent = (totalEarned / MAX_TOTAL) * 100;
  const letter = gradeFromPercent(percent);
  auditPercent.textContent = `Genel yüzde: %${percent.toFixed(2)} (toplam ${totalEarned} / ${MAX_TOTAL})`;
  auditLetter.textContent = letter;
  setStatusAudit(`Hesaplanan not: ${letter}`);
  return { k, s, t, totalEarned, percent, letter };
}

function recalcAuditUi() {
  const { sums, errors } = collectAuditAll();
  updateAuditSummaryFromSums(sums, errors);
}

function renderOperationAuditForm() {
  const form = getAuditFormDef();
  if (!form || !auditFormMount) return;

  auditFormMount.innerHTML = "";
  for (const section of form.sections) {
    const card = document.createElement("section");
    card.className = "card auditSectionCard";
    const h2 = document.createElement("h2");
    h2.textContent = `${section.title} (max ${section.maxTotal} puan)`;
    card.appendChild(h2);

    section.items.forEach((item, idx) => {
      const row = document.createElement("div");
      row.className = "auditQuestionRow";

      const text = document.createElement("div");
      text.className = "auditQText";
      text.textContent = item.text;

      const maxEl = document.createElement("div");
      maxEl.className = "auditQMax";
      maxEl.textContent = `Maks: ${item.max}`;

      const inputs = document.createElement("div");
      inputs.className = "auditQInputs";

      const earned = document.createElement("input");
      earned.type = "number";
      earned.min = "0";
      earned.max = String(item.max);
      earned.step = "0.5";
      earned.id = auditEarnedId(section.key, idx);
      earned.placeholder = "Puan";

      const comment = document.createElement("input");
      comment.type = "text";
      comment.id = auditCommentId(section.key, idx);
      comment.placeholder = "Yorum";

      const hint = document.createElement("div");
      hint.className = "fieldHint";
      hint.textContent = `0–${item.max}`;

      inputs.appendChild(earned);
      inputs.appendChild(comment);
      inputs.appendChild(hint);

      row.appendChild(text);
      row.appendChild(maxEl);
      row.appendChild(inputs);
      card.appendChild(row);

      earned.addEventListener("input", recalcAuditUi);
    });

    auditFormMount.appendChild(card);
  }
  recalcAuditUi();
}

function initAuditMetaDefaults() {
  const d = new Date();
  const iso = d.toISOString().slice(0, 10);
  if (auditReportDate && !auditReportDate.value) {
    auditReportDate.value = iso;
  }
}

function buildAuditOzetSheet(meta, agg) {
  const form = getAuditFormDef();
  const scale = form ? form.gradeScaleNote : "";
  const pk = ((agg.k / MAX_QUALITY) * 100).toFixed(2);
  const ps = ((agg.s / MAX_SERVICE) * 100).toFixed(2);
  const pt = ((agg.t / MAX_CLEAN) * 100).toFixed(2);

  return [
    [form ? form.title : "OPERASYON DENETİM"],
    [],
    ["Şube adı", meta.branch || "—"],
    ["Şube müdürü", meta.branchManager || "—"],
    ["Bölge sorumlusu", meta.regional || "—"],
    ["Raporu hazırlayan", meta.preparedBy || "—"],
    ["Ziyaret günü müdür", meta.visitBy || "—"],
    ["Rapor tarihi", meta.reportDate || "—"],
    [],
    ["DEĞERLENDİRME ÖZETİ"],
    ["Kategori", "Maksimum", "Alınan", "Bölüm yüzdesi"],
    ["Kalite", MAX_QUALITY, agg.k, `${pk}%`],
    ["Servis", MAX_SERVICE, agg.s, `${ps}%`],
    ["Temizlik ve güvenlik", MAX_CLEAN, agg.t, `${pt}%`],
    [],
    ["Toplam maksimum", MAX_TOTAL],
    ["Toplam alınan", agg.totalEarned],
    ["Genel yüzde", Number(agg.percent.toFixed(4))],
    ["Not (skala)", agg.letter],
    [],
    [scale],
    [],
    [
      "Not: Raporda tespit edilen eksiklikler, şubenin kendini geliştirmesi için fırsatlardır; raporu inceledikten sonra bu fırsatları değerlendiriniz.",
    ],
  ];
}


async function initOtherView() {
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

/** Rapor genelde ay bittikten sonra; varsayılan = bir önceki ay (Ocak → Aralık). */
function getPreviousMonthIndex(referenceDate = new Date()) {
  const m = referenceDate.getMonth();
  return m === 0 ? 11 : m - 1;
}

function applyDefaultStockReportingMonth() {
  if (!stockMonthSelect || !stockMonthSelect.options.length) return;
  stockMonthSelect.value = String(getPreviousMonthIndex());
}

function initMonthSelect() {
  stockMonthSelect.innerHTML = "";
  const defaultIdx = getPreviousMonthIndex();
  TURKISH_MONTHS.forEach((name, idx) => {
    const opt = document.createElement("option");
    opt.value = String(idx);
    opt.textContent = name;
    if (idx === defaultIdx) opt.selected = true;
    stockMonthSelect.appendChild(opt);
  });
  stockMonthSelect.addEventListener("change", updateStockFileHint);
}

btnModeStock.addEventListener("click", () => {
  stockGrid = null;
  stockPreviewCard.classList.add("hidden");
  stockPreviewWrap.innerHTML = "";
  stockPreviewMeta.textContent = "";
  setStatusStock(UI_TEXT.statusReady);
  showView("stock");
  applyDefaultStockReportingMonth();
  updateStockFileHint();
});

btnModeAudit.addEventListener("click", () => {
  auditPercent.textContent = "Genel yüzde: —";
  auditLetter.textContent = "—";
  auditSumKalite.textContent = `0 / ${MAX_QUALITY}`;
  auditSumServis.textContent = `0 / ${MAX_SERVICE}`;
  auditSumTemizlik.textContent = `0 / ${MAX_CLEAN}`;
  setStatusAudit(UI_TEXT.statusReady);
  showView("audit");
  initAuditMetaDefaults();
  renderOperationAuditForm();
});

btnModeOther.addEventListener("click", async () => {
  showView("other");
  await initOtherView();
});

backFromStock.addEventListener("click", () => {
  showView("home");
  setStatusStock(UI_TEXT.statusReady);
});

backFromAudit.addEventListener("click", () => {
  showView("home");
});

backFromOther.addEventListener("click", () => {
  showView("home");
});

clearStock.addEventListener("click", () => {
  stockGrid = null;
  stockPreviewCard.classList.add("hidden");
  stockPreviewWrap.innerHTML = "";
  stockPreviewMeta.textContent = "";
  setStatusStock("Temizlendi. Yeniden yapıştırın.");
});

saveStockDesktop.addEventListener("click", async () => {
  if (!stockGrid || stockGrid.length < 2) {
    setStatusStock("Önce Excel tablosunu yapıştırın.");
    return;
  }
  try {
    const label = getSelectedMonthUpperLabel();
    const fileName = `${label} STOK KAPANIŞ`;
    setStatusStock("Kaydediliyor...");
    const saved = await window.bridgeApi.saveXlsxDesktop({
      grid: stockGrid,
      fileName,
      exportKind: "monthlyStock",
    });
    setStatusStock(`Masaüstüne kaydedildi: ${saved.fileName}`);
  } catch (err) {
    setStatusStock(`${UI_TEXT.errorPrefix}${err.message}`);
  }
});

function bindDropZone(el, onFiles) {
  el.addEventListener("dragenter", (e) => {
    e.preventDefault();
    el.classList.add("dragOver");
  });
  el.addEventListener("dragover", (e) => {
    e.preventDefault();
    el.classList.add("dragOver");
  });
  el.addEventListener("dragleave", (e) => {
    e.preventDefault();
    el.classList.remove("dragOver");
  });
  el.addEventListener("drop", async (e) => {
    e.preventDefault();
    el.classList.remove("dragOver");
    const files = Array.from(e.dataTransfer.files || []);
    if (!files.length) return;
    await onFiles(files[0]);
  });
}

bindDropZone(dropZoneStock, async (file) => {
  if (activeView !== "stock") return;
  setStatusStock(UI_TEXT.dropReading);
  try {
    const parsed = await parseDroppedFile(file);
    await handleParsedStockGrid(parsed.grid);
  } catch (err) {
    setStatusStock(`${UI_TEXT.errorPrefix}${err.message}`);
  }
});

bindDropZone(dropZone, async (file) => {
  if (activeView !== "other") return;
  setStatus(UI_TEXT.dropReading);
  try {
    const parsed = await parseDroppedFile(file);
    await handleParsedGrid(parsed.grid, {
      suggestedBaseName: parsed.suggestedBaseName,
      sourceType: parsed.sourceType,
    });
  } catch (err) {
    setStatus(`${UI_TEXT.errorPrefix}${err.message}`);
  }
});

saveAuditDesktop.addEventListener("click", async () => {
  const { detailRows, errors, sums } = collectAuditAll();
  if (errors.length) {
    setStatusAudit(`Kayıt yapılamadı: ${errors[0]}`);
    return;
  }

  const k = sums.kalite ?? 0;
  const s = sums.servis ?? 0;
  const t = sums.temizlik ?? 0;
  const totalEarned = k + s + t;
  const percent = (totalEarned / MAX_TOTAL) * 100;
  const letter = gradeFromPercent(percent);

  const expectedDetail =
    (getAuditFormDef()?.sections || []).reduce((acc, sec) => acc + sec.items.length, 0) || 0;
  if (expectedDetail && detailRows.length !== expectedDetail) {
    setStatusAudit("Tüm maddeler eksiksiz ve geçerli puanlarla doldurulmalıdır.");
    return;
  }

  const meta = {
    branch: (auditBranchName && auditBranchName.value) || "",
    branchManager: (auditBranchManager && auditBranchManager.value) || "",
    regional: (auditRegionalManager && auditRegionalManager.value) || "",
    preparedBy: (auditPreparedBy && auditPreparedBy.value) || "",
    visitBy: (auditVisitDayManager && auditVisitDayManager.value) || "",
    reportDate:
      (auditReportDate && auditReportDate.value) ||
      new Date().toISOString().slice(0, 10),
  };

  const ozet = buildAuditOzetSheet(meta, { k, s, t, totalEarned, percent, letter });
  const detayHeader = [["Bölüm", "No", "Soru", "Maksimum", "Puan", "Yorum"]];
  const detayData = detayHeader.concat(detailRows);

  try {
    setStatusAudit("Kaydediliyor...");
    const safeBranch =
      (auditBranchName && auditBranchName.value.trim().replace(/\s+/g, " ")) || "Denetim";
    const fname = `Operasyon_Denetim_${safeBranch.replace(/[<>:"/\\|?*]/g, "_")}`;
    await window.bridgeApi.saveXlsxDesktop({
      fileName: fname,
      sheets: [
        { name: "Özet", data: ozet },
        { name: "Detay", data: detayData },
      ],
    });
    setStatusAudit("Masaüstüne Excel kaydedildi (Özet + Detay).");
  } catch (err) {
    setStatusAudit(`${UI_TEXT.errorPrefix}${err.message}`);
  }
});

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

clearNowBtn.addEventListener("click", () => {
  currentGrid = null;
  previewMeta.textContent = "";
  previewTableWrap.innerHTML = "";
  previewCard.classList.add("hidden");
  setStatus(UI_TEXT.cleared);
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

window.addEventListener(
  "paste",
  async (event) => {
    try {
      const { grid, sourceType } = getClipboardGridFromPasteEvent(event);
      if (!grid) return;

      if (activeView === "stock") {
        event.preventDefault();
        setStatusStock(UI_TEXT.pasteCleaning);
        await handleParsedStockGrid(grid);
        return;
      }

      if (activeView === "other") {
        event.preventDefault();
        setStatus(UI_TEXT.pasteCleaning);
        await handleParsedGrid(grid, { suggestedBaseName: "paste", sourceType });
      }
    } catch (err) {
      if (activeView === "stock") {
        setStatusStock(`${UI_TEXT.errorPrefix}${err.message}`);
      } else if (activeView === "other") {
        setStatus(`${UI_TEXT.errorPrefix}${err.message}`);
      }
    }
  },
  true
);

window.bridgeApi.onStatus((payload) => {
  if (!payload?.message) return;
  if (activeView === "other") setStatus(payload.message);
});

function init() {
  initMonthSelect();
  showView("home");
  setStatus(UI_TEXT.statusReady);
  setStatusStock(UI_TEXT.statusReady);
  setStatusAudit(UI_TEXT.statusReady);
}

init();
