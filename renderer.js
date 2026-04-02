/**
 * Assistant — renderer: gün sonu özeti, aylık stok, denetim formu, diğer yapıştırma akışı.
 */

const topbarSubtitle = document.getElementById("topbarSubtitle");
const targetBar = document.getElementById("targetBar");

const viewHome = document.getElementById("viewHome");
const viewDayEnd = document.getElementById("viewDayEnd");
const viewStock = document.getElementById("viewStock");
const viewAudit = document.getElementById("viewAudit");
const viewOther = document.getElementById("viewOther");

const btnModeDayEnd = document.getElementById("btnModeDayEnd");
const btnModeStock = document.getElementById("btnModeStock");
const btnModeAudit = document.getElementById("btnModeAudit");
const btnModeOther = document.getElementById("btnModeOther");
const backFromDayEnd = document.getElementById("backFromDayEnd");
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

const dropZoneDayEnd = document.getElementById("dropZoneDayEnd");
const daySummaryDate = document.getElementById("daySummaryDate");
const saveDayEndDesktop = document.getElementById("saveDayEndDesktop");
const sendDayEnd = document.getElementById("sendDayEnd");
const clearDayEnd = document.getElementById("clearDayEnd");
const dayEndPreviewCard = document.getElementById("dayEndPreviewCard");
const dayEndPreviewMeta = document.getElementById("dayEndPreviewMeta");
const dayEndPreviewWrap = document.getElementById("dayEndPreviewWrap");
const dayEndFileHint = document.getElementById("dayEndFileHint");
const statusBoxDayEnd = document.getElementById("statusBoxDayEnd");

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

/** @type {string[][] | null} */
let dayEndGrid = null;

/** Puan yazılınca sonraki puan kutusuna atlamak için zamanlayıcılar */
const auditEarnedNavTimers = new WeakMap();

let currentGrid = null;
let currentMeta = {
  suggestedBaseName: "paste",
  sourceType: "paste",
};

const DAY_END_HEADERS = ["Açıklama", "Adet", "Tutar", "OSF"];

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

function setStatusDayEnd(text) {
  statusBoxDayEnd.textContent = text;
}

function setStatusAudit(text) {
  statusBoxAudit.textContent = text;
}

const appToast = document.getElementById("appToast");
let appToastTimer = null;

/**
 * Uygulama içi kısa bilgi kutusu; birkaç saniye sonra kapanır.
 * @param {"success"|"info"} kind
 */
function showAppToast(message, kind = "success") {
  if (!appToast) return;
  const text = String(message || "").trim().slice(0, 280);
  if (!text) return;
  appToast.textContent = text;
  appToast.classList.remove("appToast--info", "appToast--success", "appToast--visible");
  appToast.classList.add(kind === "info" ? "appToast--info" : "appToast--success");
  appToast.hidden = false;
  clearTimeout(appToastTimer);
  requestAnimationFrame(() => {
    appToast.classList.add("appToast--visible");
  });
  appToastTimer = setTimeout(() => {
    appToast.classList.remove("appToast--visible");
    setTimeout(() => {
      appToast.hidden = true;
    }, 220);
  }, 3000);
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
  viewDayEnd.classList.toggle("hidden", view !== "dayEnd");
  viewStock.classList.toggle("hidden", view !== "stock");
  viewAudit.classList.toggle("hidden", view !== "audit");
  viewOther.classList.toggle("hidden", view !== "other");

  targetBar.classList.toggle("hidden", view !== "other" && view !== "dayEnd");

  if (view === "home") {
    topbarSubtitle.textContent = "Belge tipi seçin";
  } else if (view === "dayEnd") {
    topbarSubtitle.textContent = "Gün sonu özet raporu";
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

/**
 * Sağdan tüm satırlarda boş olan sütunları kırp (kopyalanan tabloda fazladan boş
 * sütun kalırsa Excel’de dikey boş kolon oluşmasın).
 */
function trimTrailingEmptyColumns(rows) {
  if (!rows.length) return rows;
  const maxLen = rows.reduce((m, r) => Math.max(m, r.length), 0);
  let end = maxLen;
  while (end > 0) {
    const ci = end - 1;
    const allEmpty = rows.every((row) => {
      const v = row[ci];
      if (v == null || v === "") return true;
      return String(v).trim() === "";
    });
    if (!allEmpty) break;
    end -= 1;
  }
  if (end === maxLen) return rows;
  return rows.map((row) => row.slice(0, end));
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
  return trimTrailingEmptyColumns(prependMonthlyHeaders(dataOnly));
}

function prependDayEndHeaders(rows) {
  return [DAY_END_HEADERS.slice(), ...rows];
}

function processDayEndRaw(rawGrid) {
  const padded = gridToRowsRaw(rawGrid);
  if (!padded.length) return [];
  const noLeadingEmptyCol = removeFirstColumnIfAllEmpty(padded);
  const dataOnly = dropCompletelyEmptyRows(noLeadingEmptyCol);
  if (!dataOnly.length) return [];
  return trimTrailingEmptyColumns(prependDayEndHeaders(dataOnly));
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

async function handleParsedDayEndGrid(rawGrid) {
  const processed = processDayEndRaw(rawGrid);
  if (!processed.length || processed.length < 2) {
    dayEndGrid = null;
    dayEndPreviewCard.classList.add("hidden");
    dayEndPreviewMeta.textContent = "";
    setStatusDayEnd(UI_TEXT.gridEmpty);
    updateDayEndSendEnabled();
    return;
  }

  dayEndGrid = processed;
  const dataRows = processed.length - 1;
  const cols = processed[0].length;
  dayEndPreviewMeta.textContent = `Başlık + ${dataRows} veri satırı, ${cols} sütun`;
  renderPreviewTable(processed, dayEndPreviewWrap);
  dayEndPreviewCard.classList.remove("hidden");
  setStatusDayEnd(
    "Tablo hazır. Masaüstüne kaydedin veya hedef klasöre gönderin.",
  );
  updateDayEndFileHint();
  updateDayEndSendEnabled();
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

/** Ay seçimine göre dosya adındaki yıl (ör. Ocak’ta varsayılan Ar → önceki yıl). */
function yearForSelectedStockMonth(monthIdx, ref = new Date()) {
  const curM = ref.getMonth();
  const y = ref.getFullYear();
  if (monthIdx > curM) return y - 1;
  return y;
}

function getStockFileSlug() {
  const idx = parseInt(stockMonthSelect.value, 10);
  if (Number.isNaN(idx) || idx < 0 || idx > 11) return "stok-kapanis";
  const year = yearForSelectedStockMonth(idx);
  const mon = TURKISH_MONTHS[idx].toLocaleLowerCase("tr-TR");
  return `${mon}-${year}-stok-kapanis`;
}

function updateStockFileHint() {
  stockFileHint.textContent = `Dosya adı: ${getStockFileSlug()}.xlsx`;
}

function getPreviousDayISO(ref = new Date()) {
  const d = new Date(ref.getFullYear(), ref.getMonth(), ref.getDate() - 1);
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

/** date input yyyy-mm-dd → dd-mm-yyyy-gun-sonu */
function formatDayEndFileStem(isoDate) {
  const raw = String(isoDate || "").trim();
  const parts = raw.split("-");
  if (parts.length !== 3) return "gun-sonu";
  const [y, m, d] = parts;
  if (!y || !m || !d) return "gun-sonu";
  return `${d}-${m}-${y}-gun-sonu`;
}

function applyDefaultDayEndDate() {
  if (!daySummaryDate) return;
  daySummaryDate.value = getPreviousDayISO();
}

function updateDayEndFileHint() {
  if (!dayEndFileHint || !daySummaryDate) return;
  dayEndFileHint.textContent = `Dosya adı: ${formatDayEndFileStem(daySummaryDate.value)}.xlsx`;
}

function updateDayEndSendEnabled() {
  const hasTarget = targetPathLabel.dataset.hasTarget === "1";
  const hasGrid = Boolean(dayEndGrid && dayEndGrid.length >= 2);
  if (sendDayEnd) sendDayEnd.disabled = !(hasTarget && hasGrid);
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
    return {
      detailRows: [],
      errors: [{ message: "Form tanımı yüklenemedi.", earnedEl: null }],
      sums: { kalite: 0, servis: 0, temizlik: 0 },
    };
  }

  const detailRows = [];
  /** @type {{ message: string, earnedEl: HTMLElement | null }[]} */
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
        errors.push({
          message: `${section.title} #${globalNo}: geçersiz puan`,
          earnedEl,
        });
        globalNo += 1;
        return;
      }
      if (earned < 0 || earned > item.max) {
        errors.push({
          message: `${section.title} #${globalNo}: puan 0–${item.max} olmalı (girilen: ${earned})`,
          earnedEl,
        });
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

function auditErrorText(err) {
  if (typeof err === "string") return err;
  return err?.message ?? "";
}

function updateAuditSummaryFromSums(sums, errors) {
  const k = sums.kalite ?? 0;
  const s = sums.servis ?? 0;
  const t = sums.temizlik ?? 0;
  const totalEarned = k + s + t;

  auditSumKalite.textContent = `${Number(k.toFixed(2))} / ${MAX_QUALITY}`;
  auditSumServis.textContent = `${Number(s.toFixed(2))} / ${MAX_SERVICE}`;
  auditSumTemizlik.textContent = `${Number(t.toFixed(2))} / ${MAX_CLEAN}`;

  if (errors.length) {
    auditPercent.textContent = `Genel yüzde: — (${errors.length} satır hatalı)`;
    auditLetter.textContent = "—";
    setStatusAudit(auditErrorText(errors[0]));
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

/** Yorumu atla: bir sonraki sorunun puan kutusu (veya bir sonraki bölümün ilki). */
function focusNextAuditEarnedInput(currentEarned) {
  const row = currentEarned.closest(".auditQuestionRow");
  if (!row) return;
  let nextRow = row.nextElementSibling;
  while (nextRow && !nextRow.classList.contains("auditQuestionRow")) {
    nextRow = nextRow.nextElementSibling;
  }
  if (nextRow) {
    const nextEarned = nextRow.querySelector('input[type="number"]');
    if (nextEarned) {
      nextEarned.focus();
      try {
        nextEarned.select();
      } catch (_) {}
      return;
    }
  }
  const card = row.closest(".auditSectionCard");
  const nextCard = card?.nextElementSibling;
  if (nextCard && nextCard.classList.contains("auditSectionCard")) {
    const firstRow = nextCard.querySelector(".auditQuestionRow");
    const ne = firstRow?.querySelector('input[type="number"]');
    if (ne) {
      ne.focus();
      try {
        ne.select();
      } catch (_) {}
    }
  }
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
      earned.step = "any";
      earned.id = auditEarnedId(section.key, idx);
      earned.placeholder = "Puan";
      earned.inputMode = "decimal";

      const comment = document.createElement("input");
      comment.type = "text";
      comment.id = auditCommentId(section.key, idx);
      comment.placeholder = "Yorum";

      const hint = document.createElement("div");
      hint.className = "fieldHint";
      hint.textContent = `0–${item.max} (ör. 4,8)`;

      inputs.appendChild(earned);
      inputs.appendChild(comment);
      inputs.appendChild(hint);

      row.appendChild(text);
      row.appendChild(maxEl);
      row.appendChild(inputs);
      card.appendChild(row);

      const scheduleJumpToNextEarned = () => {
        const prev = auditEarnedNavTimers.get(earned);
        if (prev) clearTimeout(prev);
        const t = setTimeout(() => {
          auditEarnedNavTimers.delete(earned);
          if (
            document.activeElement === earned &&
            document.body.contains(earned)
          ) {
            focusNextAuditEarnedInput(earned);
          }
        }, 100);
        auditEarnedNavTimers.set(earned, t);
      };

      earned.addEventListener("input", () => {
        recalcAuditUi();
        scheduleJumpToNextEarned();
      });
      earned.addEventListener("blur", () => {
        const prev = auditEarnedNavTimers.get(earned);
        if (prev) clearTimeout(prev);
        auditEarnedNavTimers.delete(earned);
      });
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
    ["Kalite", MAX_QUALITY, Number(agg.k.toFixed(2)), `${pk}%`],
    ["Servis", MAX_SERVICE, Number(agg.s.toFixed(2)), `${ps}%`],
    ["Temizlik ve güvenlik", MAX_CLEAN, Number(agg.t.toFixed(2)), `${pt}%`],
    [],
    ["Toplam maksimum", MAX_TOTAL],
    ["Toplam alınan", Number(agg.totalEarned.toFixed(2))],
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

async function initDayEndView() {
  const config = await window.bridgeApi.getConfig();
  if (config.hasTarget) {
    targetPathLabel.textContent = config.targetPath;
    targetPathLabel.dataset.hasTarget = "1";
  } else {
    targetPathLabel.textContent = UI_TEXT.targetNotSelected;
    targetPathLabel.dataset.hasTarget = "0";
  }
  updateCounter(config);
  updateDayEndSendEnabled();
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

btnModeDayEnd.addEventListener("click", async () => {
  dayEndGrid = null;
  dayEndPreviewCard.classList.add("hidden");
  dayEndPreviewWrap.innerHTML = "";
  dayEndPreviewMeta.textContent = "";
  setStatusDayEnd(UI_TEXT.statusReady);
  showView("dayEnd");
  applyDefaultDayEndDate();
  updateDayEndFileHint();
  await initDayEndView();
});

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

backFromDayEnd.addEventListener("click", () => {
  showView("home");
  setStatusDayEnd(UI_TEXT.statusReady);
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

if (daySummaryDate) {
  daySummaryDate.addEventListener("change", () => {
    updateDayEndFileHint();
  });
}

saveDayEndDesktop.addEventListener("click", async () => {
  if (!dayEndGrid || dayEndGrid.length < 2) {
    setStatusDayEnd("Önce tabloyu yapıştırın.");
    return;
  }
  try {
    const fileName = formatDayEndFileStem(daySummaryDate.value);
    setStatusDayEnd("Kaydediliyor...");
    const saved = await window.bridgeApi.saveXlsxDesktop({
      grid: dayEndGrid,
      fileName,
      exportKind: "dayEndSummary",
    });
    setStatusDayEnd(`Masaüstüne kaydedildi: ${saved.fileName}`);
    showAppToast(`Masaüstüne kaydedildi: ${saved.fileName}`);
  } catch (err) {
    setStatusDayEnd(`${UI_TEXT.errorPrefix}${err.message}`);
  }
});

sendDayEnd.addEventListener("click", async () => {
  if (!dayEndGrid || dayEndGrid.length < 2) {
    setStatusDayEnd(UI_TEXT.sendNoData);
    return;
  }
  setStatusDayEnd(UI_TEXT.sending);
  try {
    const result = await window.bridgeApi.sendGrid({
      grid: dayEndGrid,
      fileName: formatDayEndFileStem(daySummaryDate.value),
      exportKind: "dayEndSummary",
    });
    if (result.deduped) {
      setStatusDayEnd(UI_TEXT.deduped);
      showAppToast("Bu tablo zaten gönderilmiş; yeni dosya yazılmadı.", "info");
    } else {
      setStatusDayEnd(`${UI_TEXT.sendSuccess}${result.fileName}`);
      showAppToast(`Hedef klasöre kaydedildi: ${result.fileName}`);
    }
    await initDayEndView();
  } catch (err) {
    setStatusDayEnd(`${UI_TEXT.errorPrefix}${err.message}`);
  }
});

clearDayEnd.addEventListener("click", () => {
  dayEndGrid = null;
  dayEndPreviewCard.classList.add("hidden");
  dayEndPreviewWrap.innerHTML = "";
  dayEndPreviewMeta.textContent = "";
  setStatusDayEnd("Temizlendi. Yeniden yapıştırın.");
  updateDayEndSendEnabled();
});

saveStockDesktop.addEventListener("click", async () => {
  if (!stockGrid || stockGrid.length < 2) {
    setStatusStock("Önce Excel tablosunu yapıştırın.");
    return;
  }
  try {
    const fileName = getStockFileSlug();
    setStatusStock("Kaydediliyor...");
    const saved = await window.bridgeApi.saveXlsxDesktop({
      grid: stockGrid,
      fileName,
      exportKind: "monthlyStock",
    });
    setStatusStock(`Masaüstüne kaydedildi: ${saved.fileName}`);
    showAppToast(`Masaüstüne kaydedildi: ${saved.fileName}`);
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

bindDropZone(dropZoneDayEnd, async (file) => {
  if (activeView !== "dayEnd") return;
  setStatusDayEnd(UI_TEXT.dropReading);
  try {
    const parsed = await parseDroppedFile(file);
    await handleParsedDayEndGrid(parsed.grid);
  } catch (err) {
    setStatusDayEnd(`${UI_TEXT.errorPrefix}${err.message}`);
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
    const msg = auditErrorText(errors[0]);
    setStatusAudit(`Kayıt yapılamadı: ${msg}`);
    const firstBad = errors.find((e) => e && typeof e === "object" && e.earnedEl);
    if (firstBad?.earnedEl) {
      firstBad.earnedEl.scrollIntoView({ behavior: "smooth", block: "center" });
      firstBad.earnedEl.focus();
      try {
        firstBad.earnedEl.select();
      } catch (_) {}
    }
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
  const eksikRows = detailRows.filter(
    (row) => Number(row[4]) < Number(row[3]) - 1e-9,
  );
  const eksikWithHeader = detayHeader.concat(eksikRows);

  try {
    setStatusAudit("Kaydediliyor...");
    const safeBranch =
      (auditBranchName && auditBranchName.value.trim().replace(/\s+/g, " ")) || "Denetim";
    const fname = `Operasyon_Denetim_${safeBranch.replace(/[<>:"/\\|?*]/g, "_")}`;
    await window.bridgeApi.saveXlsxDesktop({
      fileName: fname,
      exportKind: "auditWorkbook",
      auditBook: {
        ozet,
        detayWithHeader: detayData,
        eksikWithHeader,
      },
    });
    setStatusAudit(
      "Masaüstüne kaydedildi: Özet, Detay ve Tam puan alınmayan sayfaları.",
    );
    showAppToast("Denetim Excel’i kaydedildi (3 sayfa).");
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
    updateDayEndSendEnabled();
    setStatus(UI_TEXT.targetUpdated);
    if (activeView === "dayEnd") {
      setStatusDayEnd(UI_TEXT.targetUpdated);
    }
  } catch (err) {
    setStatus(`${UI_TEXT.errorPrefix}${err.message}`);
    if (activeView === "dayEnd") {
      setStatusDayEnd(`${UI_TEXT.errorPrefix}${err.message}`);
    }
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
      showAppToast("Bu içerik zaten gönderilmiş; yeni dosya yazılmadı.", "info");
    } else {
      setStatus(`${UI_TEXT.sendSuccess}${result.fileName}`);
      showAppToast(`Hedef klasöre kaydedildi: ${result.fileName}`);
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

      if (activeView === "dayEnd") {
        event.preventDefault();
        setStatusDayEnd(UI_TEXT.pasteCleaning);
        await handleParsedDayEndGrid(grid);
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
      } else if (activeView === "dayEnd") {
        setStatusDayEnd(`${UI_TEXT.errorPrefix}${err.message}`);
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
  if (activeView === "dayEnd") setStatusDayEnd(payload.message);
});

function init() {
  initMonthSelect();
  showView("home");
  setStatus(UI_TEXT.statusReady);
  setStatusStock(UI_TEXT.statusReady);
  setStatusDayEnd(UI_TEXT.statusReady);
  setStatusAudit(UI_TEXT.statusReady);
}

init();
