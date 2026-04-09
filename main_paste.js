const { app, BrowserWindow, dialog, ipcMain, Menu } = require("electron");
const path = require("path");
const fs = require("fs-extra");
const { autoUpdater } = require("electron-updater");

const xlsx = require("xlsx");
const ExcelJS = require("exceljs");
const pdfParse = require("pdf-parse");

const APP_DIR = app.getPath("userData");
const STATE_FILE = path.join(APP_DIR, "state.json");

/** Arayüzde varsayılan; gönderimde pathname boşsa /ingest ile POST edilir. */
const DEFAULT_CLOUD_WORKER_URL = "https://restcloud.gokberktanis.workers.dev";

if (process.platform === "win32") {
  app.setAppUserModelId("com.nekwake.assistant");
}

let mainWindow = null;
let state = {
  /** Tam ingest URL (örn. https://assistant-xxx.workers.dev/ingest) */
  cloudWorkerUrl: "",
  /** Şube anahtarı; Worker X-Branch-Key ile doğrular */
  branchKey: "",
  sentHistory: [],
};

function sendStatus(message, level = "info") {
  if (!mainWindow) return;
  mainWindow.webContents.send("status", { message, level, at: Date.now() });
}

function sanitizeFileBaseName(baseName) {
  const name = String(baseName || "rapor");
  // Keep letters, numbers, underscore, dash. Replace others.
  return name.replace(/[^a-zA-Z0-9_-]+/g, "_").replace(/^_+|_+$/g, "") || "rapor";
}

/** Allow Turkish and most printable chars; strip Windows-forbidden path characters */
function sanitizeDesktopFileName(name) {
  const base = String(name || "dosya").trim();
  return base.replace(/[<>:"/\\|?*\x00-\x1f]/g, " ").replace(/\s+/g, " ").trim() || "dosya";
}

function fileExists(p) {
  try {
    return fs.existsSync(p);
  } catch {
    return false;
  }
}

function normalizeGridForHash(grid) {
  const normalized = (grid || []).map((row) => {
    if (!Array.isArray(row)) return [String(row ?? "").trim()];
    return row.map((c) => String(c ?? "").trim());
  });
  // Ensure stable output by trimming fully empty trailing columns per row.
  const maxCols = normalized.reduce((acc, r) => Math.max(acc, r.length), 0);
  const padded = normalized.map((r) => {
    const rr = r.slice();
    while (rr.length < maxCols) rr.push("");
    return rr;
  });
  // Remove empty rows again (renderer should already do it).
  const nonEmptyRows = padded.filter((r) => r.some((c) => c !== ""));
  return nonEmptyRows;
}

function textToGridBestEffort(text) {
  const lines = String(text || "")
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter((l) => l.length > 0);

  const grid = [];
  const maxLines = 250;
  for (const line of lines.slice(0, maxLines)) {
    // If OCR-style text: multiple spaces often separate columns.
    if (line.includes("\t")) {
      grid.push(line.split("\t").map((c) => c.trim()));
      continue;
    }

    const parts = line.split(/\s{2,}/);
    if (parts.length > 1) {
      grid.push(parts.map((c) => c.trim()));
    } else {
      grid.push([line]);
    }
  }

  // Guarantee at least one row.
  if (!grid.length) grid.push([""]);
  return grid;
}

async function loadState() {
  await fs.ensureDir(APP_DIR);
  if (await fs.pathExists(STATE_FILE)) {
    const fileState = await fs.readJson(STATE_FILE);
    state = {
      ...state,
      ...fileState,
      cloudWorkerUrl: fileState.cloudWorkerUrl || "",
      branchKey: fileState.branchKey || "",
      sentHistory: fileState.sentHistory || [],
    };
  }
  delete state.targetPath;
  delete state.sentHashes;
}

async function persistState() {
  await fs.writeJson(STATE_FILE, state, { spaces: 2 });
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 960,
    height: 700,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
    },
    icon: path.join(__dirname, "assets", "icon.png"),
  });

  mainWindow.loadFile(path.join(__dirname, "index.html"));
}

/** GitHub Releases üzerinden; yalnızca paketlenmiş kurulumda çalışır. */
function setupAutoUpdater() {
  if (!app.isPackaged) return;

  autoUpdater.autoDownload = true;
  autoUpdater.autoInstallOnAppQuit = true;

  autoUpdater.on("update-available", (info) => {
    sendStatus(`Yeni sürüm: v${info.version} indiriliyor…`, "info");
  });

  autoUpdater.on("update-downloaded", (info) => {
    sendStatus(
      `Sürüm ${info.version} hazır. Kapatıp açınca güncellenir.`,
      "success",
    );
    const win = mainWindow && !mainWindow.isDestroyed() ? mainWindow : null;
    const dlgOpts = {
      type: "info",
      title: "Assistant — güncelleme",
      message: `Sürüm ${info.version} indirildi.`,
      detail:
        "Güncelleme, uygulamayı kapattığınızda uygulanır; tekrar açtığınızda yeni sürüm çalışır.",
      buttons: ["Tamam"],
      defaultId: 0,
    };
    (win ? dialog.showMessageBox(win, dlgOpts) : dialog.showMessageBox(dlgOpts)).catch(
      () => {},
    );
  });

  autoUpdater.on("error", (err) => {
    console.error("[autoUpdater]", err?.message || err);
  });

  autoUpdater.checkForUpdates().catch((err) => {
    console.error("[checkForUpdates]", err?.message || err);
  });

  const sixHoursMs = 6 * 60 * 60 * 1000;
  setInterval(() => {
    autoUpdater.checkForUpdates().catch(() => {});
  }, sixHoursMs);
}

function effectiveDisplayCloudUrl() {
  const s = String(state.cloudWorkerUrl || "").trim();
  return s || DEFAULT_CLOUD_WORKER_URL;
}

/**
 * Worker örneği / ve /ingest kabul eder; bazı kurulumlarda kök POST 500 verebiliyor.
 * Path yoksa veya / ise POST için .../ingest kullanılır.
 */
function resolveIngestPostUrl(input) {
  const raw = String(input || "").trim() || DEFAULT_CLOUD_WORKER_URL;
  try {
    const u = new URL(raw);
    const path = (u.pathname || "").replace(/\/+$/, "") || "/";
    if (path === "" || path === "/") {
      return `${u.origin}/ingest`;
    }
    return u.toString();
  } catch {
    return raw;
  }
}

/** HTML / Cloudflare hata sayfası gövdesini kullanıcıya kısa metne indirger; cf-ray ekler. */
function summarizeCloudErrorResponse(res, bodyText) {
  const ray = res.headers.get("cf-ray") || res.headers.get("CF-Ray") || "";
  const snippet = String(bodyText || "").trim();
  const isHtml =
    /^<!DOCTYPE/i.test(snippet) ||
    /^<\s*html/i.test(snippet) ||
    /<\s*html[\s>]/i.test(snippet);
  if (isHtml) {
    let msg =
      "HTML hata sayfası döndü (Worker istisnası veya edge). Cloudflare Dashboard → Workers → bu Worker → Logs ile aynı zaman dilimindeki hatayı açın.";
    if (ray) msg += ` cf-ray: ${ray}`;
    return msg;
  }
  if (!snippet) {
    return ray ? `Yanıt gövdesi boş. cf-ray: ${ray}` : "";
  }
  return snippet.length > 240 ? `${snippet.slice(0, 237)}…` : snippet;
}

function hasCloudDeliveryConfigured() {
  return Boolean(
    effectiveDisplayCloudUrl() && String(state.branchKey || "").trim(),
  );
}

ipcMain.handle("get-config", async () => {
  const hasCloud = hasCloudDeliveryConfigured();
  return {
    cloudWorkerUrl: effectiveDisplayCloudUrl(),
    hasCloud,
    cloudConfigLocked: hasCloud,
    branchKey: hasCloud ? String(state.branchKey || "") : "",
    branchKeyConfigured: Boolean(String(state.branchKey || "").trim()),
    sentCount: state.sentHistory.length,
  };
});

/**
 * Manuel güncelleme kontrolü. Paketlenmiş uygulamada electron-updater kullanılır;
 * açılışta ve periyodik kontrol zaten setupAutoUpdater içinde.
 */
ipcMain.handle("check-for-updates", async () => {
  try {
    if (!app.isPackaged) {
      return {
        ok: true,
        kind: "development",
        message:
          "Güncelleme kontrolü yalnızca kurulu uygulamada çalışır. Şu an geliştirme modundasınız (npm start).",
      };
    }
    const result = await autoUpdater.checkForUpdates();
    if (result?.isUpdateAvailable && result.updateInfo?.version) {
      return {
        ok: true,
        kind: "available",
        version: result.updateInfo.version,
        message: `Yeni sürüm v${result.updateInfo.version} bulundu. İndirme otomatik başlar; hazır olunca kapatıp açmanız istenir.`,
      };
    }
    return {
      ok: true,
      kind: "uptodate",
      version: app.getVersion(),
      message: `Yeni sürüm yok. Kurulu sürüm: v${app.getVersion()}.`,
    };
  } catch (err) {
    const msg = err && typeof err === "object" && "message" in err ? err.message : String(err);
    return {
      ok: false,
      message: explainGithubUpdateError(msg),
    };
  }
});

/** GitHub private repo / eksik token durumunda İngilizce ham hatayı Türkçe açıklamaya bağlar. */
function explainGithubUpdateError(raw) {
  const s = String(raw || "").trim();
  const low = s.toLowerCase();
  if (
    low.includes("auth token") ||
    low.includes("double check") ||
    low.includes("bad credentials") ||
    low.includes("could not retrieve publish") ||
    (low.includes("401") && low.includes("github")) ||
    (low.includes("403") && low.includes("github"))
  ) {
    return [
      "Otomatik güncelleme bu kurulumda GitHub’dan sürüm dosyası okuyamıyor.",
      "Bunun en sık nedeni: depo private olduğu için Releases API’si dışarıya kapalıdır (token olmadan electron-updater erişemez).",
      "Seçenekler: repoyu public yapmak; yalnızca kurulum dosyalarını içeren ayrı bir public repoda release yayınlamak; veya güncellemeleri kendi sunucunuzda (generic provider) barındırmak.",
      s ? `Detay: ${s.slice(0, 200)}${s.length > 200 ? "…" : ""}` : "",
    ]
      .filter(Boolean)
      .join(" ");
  }
  return s || "Güncelleme sunucusuna ulaşılamadı (ağ veya GitHub Releases).";
}

ipcMain.handle("set-cloud-config", async (_event, payload) => {
  let workerUrl = String(payload?.workerUrl ?? "").trim();
  if (!workerUrl) workerUrl = DEFAULT_CLOUD_WORKER_URL;
  const branchKeyRaw = String(payload?.branchKey ?? "").trim();
  if (!branchKeyRaw) {
    throw new Error("Şube anahtarı gerekli. Kaydetmeden önce şube kodunu girin.");
  }
  state.cloudWorkerUrl = workerUrl;
  state.branchKey = branchKeyRaw;
  await persistState();
  return { ok: true };
});

ipcMain.handle("reset-cloud-config", async () => {
  state.cloudWorkerUrl = "";
  state.branchKey = "";
  await persistState();
  return { ok: true };
});

ipcMain.handle("parse-file-buffer", async (_event, payload) => {
  const { fileName, mimeType, bufferBase64 } = payload || {};
  if (!fileName || !bufferBase64) {
    throw new Error("Dosya verisi eksik.");
  }

  const buffer = Buffer.from(bufferBase64, "base64");
  const ext = path.extname(fileName).toLowerCase();
  const suggestedBaseName = sanitizeFileBaseName(path.parse(fileName).name);

  if ([".xlsx", ".xls", ".xlsm", ".xlsb", ".csv"].includes(ext)) {
    // Read as spreadsheet.
    const workbook = xlsx.read(buffer, { type: "buffer" });
    const firstSheet = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheet];
    const grid = xlsx.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
    return { grid, sourceType: "file-excel", suggestedBaseName };
  }

  if (ext === ".pdf" || String(mimeType || "").toLowerCase().includes("pdf")) {
    const pdfData = await pdfParse(buffer);
    const grid = textToGridBestEffort(pdfData.text);
    return { grid, sourceType: "file-pdf", suggestedBaseName };
  }

  throw new Error(
    "Desteklenmeyen dosya türü. Sadece Excel/Csv ve PDF desteklenir."
  );
});

/** Renderer’dan gelen reportDate (YYYY-MM-DD); geçersizse yerel bugün. */
function coerceReportDateISO(maybe) {
  const s = String(maybe || "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  }
  const [ys, ms, ds] = s.split("-");
  const y = Number(ys);
  const m = Number(ms);
  const day = Number(ds);
  if (!Number.isFinite(y) || m < 1 || m > 12 || day < 1 || day > 31) {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  }
  return s;
}

ipcMain.handle("send-grid", async (_event, payload) => {
  const {
    grid,
    suggestedBaseName,
    fileName: preferredStem,
    exportKind,
    reportDate: reportDateRaw,
  } = payload || {};

  if (!hasCloudDeliveryConfigured()) {
    sendStatus("Bulut gönderimi ayarlı değil; önce şube anahtarını kaydedin.", "warn");
    throw new Error("Bulut gönderimi ayarlı değil; önce şube anahtarını kaydedin.");
  }

  if (!Array.isArray(grid) || !grid.length) {
    throw new Error("Gönderilecek veri yok.");
  }

  const normalizedGrid = normalizeGridForHash(grid);

  let outBuffer;
  if (exportKind === "dayEndSummary") {
    outBuffer = await writeStyledReportXlsxBuffer(
      normalizedGrid,
      "Gün Sonu Özet",
    );
  } else if (exportKind === "monthlyStock") {
    outBuffer = await writeStyledReportXlsxBuffer(
      normalizedGrid,
      "Stok Kapanış",
    );
  } else {
    const ws = xlsx.utils.aoa_to_sheet(normalizedGrid);
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, "Sheet1");
    outBuffer = xlsx.write(wb, { type: "buffer", bookType: "xlsx" });
  }

  let fileName;
  const stemRaw = preferredStem && String(preferredStem).trim();
  if (stemRaw) {
    let stem = sanitizeDesktopFileName(stemRaw.replace(/\.xlsx$/i, ""));
    if (!stem) stem = "rapor";
    fileName = stem.toLowerCase().endsWith(".xlsx") ? stem : `${stem}.xlsx`;
  } else {
    const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
    const base = sanitizeFileBaseName(suggestedBaseName || "rapor");
    fileName = `${base}_${timestamp}.xlsx`;
  }

  let serverMessage = "";
  const ingestUrl = resolveIngestPostUrl(effectiveDisplayCloudUrl());
  const reportDate = coerceReportDateISO(reportDateRaw);
  const body = {
    version: 1,
    exportKind: exportKind || "generic",
    fileName,
    fileBase64: outBuffer.toString("base64"),
    grid: normalizedGrid,
    sentAt: new Date().toISOString(),
    reportDate,
  };
  let res;
  try {
    res = await fetch(ingestUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "X-Branch-Key": String(state.branchKey).trim(),
      },
      body: JSON.stringify(body),
    });
  } catch (err) {
    const msg = err?.message || String(err);
    throw new Error(`Buluta bağlanılamadı: ${msg}`);
  }
  if (!res.ok) {
    let detail = "";
    const t = await res.text().catch(() => "");
    try {
      const j = JSON.parse(t);
      const human =
        j && typeof j.message === "string" && String(j.message).trim()
          ? String(j.message).trim()
          : "";
      const code =
        j && typeof j.error === "string" && String(j.error).trim()
          ? String(j.error).trim()
          : "";
      if (human && code && human !== code) detail = ` (${human}) [${code}]`;
      else if (human) detail = ` (${human})`;
      else if (code) detail = ` (${code})`;
    } catch (_) {
      const summary = summarizeCloudErrorResponse(res, t);
      if (summary) detail = ` — ${summary}`;
      else if (t) detail = ` — ${t.slice(0, 120)}`;
    }
    throw new Error(`Bulut yanıtı ${res.status}${detail || ` ${res.statusText}`}`);
  }
  try {
    const ct = res.headers.get("content-type") || "";
    if (ct.includes("application/json")) {
      const j = await res.json();
      if (j && typeof j.message === "string") {
        serverMessage = j.message;
      }
    }
  } catch (_) {}

  const sentAt = new Date().toISOString();
  state.sentHistory.unshift({
    name: fileName,
    sentAt,
  });
  state.sentHistory = state.sentHistory.slice(0, 30);

  await persistState();

  sendStatus(`Buluta gönderildi: ${fileName}`, "success");
  return {
    fileName,
    viaCloud: true,
    serverMessage: serverMessage || undefined,
  };
});

ipcMain.handle("send-orders-raw", async (_event, payload) => {
  const raw = payload?.raw_data ?? payload?.rawData;
  const text = String(raw ?? "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  const trimmed = text.trim();

  if (!hasCloudDeliveryConfigured()) {
    sendStatus("Bulut gönderimi ayarlı değil; önce şube anahtarını kaydedin.", "warn");
    throw new Error("Bulut gönderimi ayarlı değil; önce şube anahtarını kaydedin.");
  }

  if (!trimmed) {
    throw new Error("Gönderilecek metin yok.");
  }

  const maxChars = 12 * 1024 * 1024;
  if (trimmed.length > maxChars) {
    throw new Error(`Metin çok uzun (en fazla ${Math.floor(maxChars / (1024 * 1024))} MB).`);
  }

  const ingestUrl = resolveIngestPostUrl(effectiveDisplayCloudUrl());
  const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
  const fileName = `siparis_ham_${timestamp}.txt`;

  const body = {
    version: 1,
    type: "orders",
    raw_data: trimmed,
    sentAt: new Date().toISOString(),
    fileName,
  };

  let res;
  try {
    res = await fetch(ingestUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "X-Branch-Key": String(state.branchKey).trim(),
      },
      body: JSON.stringify(body),
    });
  } catch (err) {
    const msg = err?.message || String(err);
    throw new Error(`Buluta bağlanılamadı: ${msg}`);
  }

  let serverMessage = "";
  if (!res.ok) {
    let detail = "";
    const t = await res.text().catch(() => "");
    try {
      const j = JSON.parse(t);
      const human =
        j && typeof j.message === "string" && String(j.message).trim()
          ? String(j.message).trim()
          : "";
      const code =
        j && typeof j.error === "string" && String(j.error).trim()
          ? String(j.error).trim()
          : "";
      if (human && code && human !== code) detail = ` (${human}) [${code}]`;
      else if (human) detail = ` (${human})`;
      else if (code) detail = ` (${code})`;
    } catch (_) {
      const summary = summarizeCloudErrorResponse(res, t);
      if (summary) detail = ` — ${summary}`;
      else if (t) detail = ` — ${t.slice(0, 120)}`;
    }
    throw new Error(`Bulut yanıtı ${res.status}${detail || ` ${res.statusText}`}`);
  }

  try {
    const ct = res.headers.get("content-type") || "";
    if (ct.includes("application/json")) {
      const j = await res.json();
      if (j && typeof j.message === "string") {
        serverMessage = j.message;
      }
    }
  } catch (_) {}

  const sentAt = new Date().toISOString();
  state.sentHistory.unshift({
    name: fileName,
    sentAt,
  });
  state.sentHistory = state.sentHistory.slice(0, 30);
  await persistState();

  sendStatus(`Sipariş listesi gönderildi: ${fileName}`, "success");
  return {
    fileName,
    viaCloud: true,
    serverMessage: serverMessage || undefined,
  };
});

ipcMain.handle("last-sent", async () => {
  return (state.sentHistory || []).slice(0, 10).map((x) => ({
    name: x.name,
    sentAt: x.sentAt,
  }));
});

function uniqueSheetName(name, usedNames) {
  let n = String(name || "Sheet")
    .replace(/[[\]:/?*\\]/g, " ")
    .trim()
    .slice(0, 31);
  if (!n) n = "Sheet";
  let candidate = n;
  let i = 1;
  while (usedNames.has(candidate)) {
    const suffix = `_${i++}`;
    candidate = `${n}`.slice(0, 31 - suffix.length) + suffix;
  }
  usedNames.add(candidate);
  return candidate;
}

/**
 * Stok / gün sonu gibi raporlar: kırmızı başlık, kenarlık, sütun genişliği (SheetJS ile yapılamaz).
 */
async function writeStyledReportXlsxBuffer(grid, worksheetName) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Assistant";
  const ws = wb.addWorksheet(worksheetName || "Rapor");

  const numRows = grid.length;
  const numCols = grid.reduce(
    (m, r) => Math.max(m, Array.isArray(r) ? r.length : 0),
    0,
  );

  const headerFill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFC62828" },
  };
  const headerFont = { bold: true, color: { argb: "FFFFFFFF" } };
  const thin = { style: "thin", color: { argb: "FF000000" } };
  const allBorders = { top: thin, left: thin, bottom: thin, right: thin };

  for (let ridx = 0; ridx < numRows; ridx++) {
    const rowArr = grid[ridx];
    const padded = [];
    for (let c = 0; c < numCols; c++) {
      const v = rowArr && rowArr[c];
      padded.push(v === "" || v == null ? "" : v);
    }
    const row = ws.addRow(padded);
    if (ridx === 0) row.height = 22;
    row.eachCell({ includeEmpty: true }, (cell) => {
      cell.border = allBorders;
      if (ridx === 0) {
        cell.fill = headerFill;
        cell.font = headerFont;
      }
      cell.alignment = { vertical: "middle", wrapText: true };
    });
  }

  for (let c = 1; c <= numCols; c++) {
    let maxLen = 8;
    for (let ridx = 0; ridx < numRows; ridx++) {
      const rowArr = grid[ridx];
      const v = rowArr && rowArr[c - 1];
      const s = v == null ? "" : String(v);
      if (s.length > maxLen) maxLen = s.length;
    }
    const w = Math.min(Math.max(maxLen * 0.92 + 2.2, 12), 48);
    ws.getColumn(c).width = w;
  }

  const buf = await wb.xlsx.writeBuffer();
  return Buffer.from(buf);
}

const AUDIT_HEADER_FILL = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FFC62828" },
};
const AUDIT_HEADER_FONT = { bold: true, color: { argb: "FFFFFFFF" } };
const AUDIT_THIN = { style: "thin", color: { argb: "FF000000" } };
const AUDIT_ALL_BORDERS = {
  top: AUDIT_THIN,
  left: AUDIT_THIN,
  bottom: AUDIT_THIN,
  right: AUDIT_THIN,
};
const AUDIT_SECTION_FILL = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FFE0E7FF" },
};

function detectAuditOzetRowKind(row) {
  const cells = row.map((c) => String(c ?? "").trim());
  const nonEmpty = cells.filter((x) => x.length > 0);
  if (nonEmpty.length === 0) return "blank";

  if (cells[0] === "DEĞERLENDİRME ÖZETİ") return "section";
  if (
    cells[0] === "Kategori" &&
    cells[1] === "Maksimum" &&
    cells[2] === "Alınan"
  ) {
    return "sumHeader";
  }
  if (cells[0]?.startsWith("Skala:")) return "scale";
  if (cells[0]?.startsWith("Not:")) return "footnote";

  if (nonEmpty.length === 1 && cells[0].length > 30) return "docTitle";

  if (
    cells[0] &&
    cells[1] !== undefined &&
    String(cells[1]).length > 0 &&
    (!cells[2] || String(cells[2]).trim() === "") &&
    cells.length <= 3
  ) {
    const labels = [
      "Şube adı",
      "Şube müdürü",
      "Bölge sorumlusu",
      "Raporu hazırlayan",
      "Ziyaret günü müdür",
      "Rapor tarihi",
    ];
    if (labels.includes(cells[0])) return "metaPair";
  }

  if (
    cells.length >= 4 &&
    ["Kalite", "Servis", "Temizlik ve güvenlik"].some((p) =>
      cells[0]?.startsWith(p),
    )
  ) {
    return "sumRow";
  }

  const totals = [
    "Toplam maksimum",
    "Toplam alınan",
    "Genel yüzde",
    "Not (skala)",
  ];
  if (totals.some((t) => cells[0]?.startsWith(t))) return "totalRow";

  if (nonEmpty.length === 1) return "singleLine";

  return "body";
}

async function fillAuditOzetWorksheet(ws, grid) {
  const numCols = Math.max(6, grid.reduce((m, r) => Math.max(m, r.length), 0));

  for (let ridx = 0; ridx < grid.length; ridx++) {
    const rowArr = grid[ridx] ?? [];
    const padded = [];
    for (let c = 0; c < numCols; c++) {
      const v = rowArr[c];
      padded.push(v === "" || v == null ? "" : v);
    }

    const kind = detectAuditOzetRowKind(rowArr);
    const row = ws.addRow(padded);

    if (kind === "blank") {
      row.height = 6;
      continue;
    }

    if (kind === "docTitle") {
      row.height = 28;
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        if (colNumber <= numCols) {
          cell.border = AUDIT_ALL_BORDERS;
          cell.font = { bold: true, size: 14, color: { argb: "FF1e3a8a" } };
          cell.alignment = { vertical: "middle", wrapText: true };
        }
      });
      try {
        ws.mergeCells(ridx + 1, 1, ridx + 1, numCols);
      } catch {
        /* */
      }
      continue;
    }

    if (kind === "section") {
      row.height = 20;
      row.eachCell({ includeEmpty: true }, (cell) => {
        cell.border = AUDIT_ALL_BORDERS;
        cell.fill = AUDIT_SECTION_FILL;
        cell.font = { bold: true, size: 12 };
        cell.alignment = { vertical: "middle", wrapText: true };
      });
      try {
        ws.mergeCells(ridx + 1, 1, ridx + 1, numCols);
      } catch {
        /* */
      }
      continue;
    }

    if (kind === "sumHeader") {
      row.height = 22;
      row.eachCell({ includeEmpty: true }, (cell) => {
        cell.border = AUDIT_ALL_BORDERS;
        cell.fill = AUDIT_HEADER_FILL;
        cell.font = AUDIT_HEADER_FONT;
        cell.alignment = { vertical: "middle", wrapText: true };
      });
      continue;
    }

    if (kind === "sumRow" || kind === "totalRow" || kind === "metaPair") {
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        cell.border = AUDIT_ALL_BORDERS;
        cell.alignment = { vertical: "middle", wrapText: true };
        if (kind === "metaPair") {
          cell.font =
            colNumber === 1
              ? { bold: true, color: { argb: "FF334155" } }
              : { color: { argb: "FF0f172a" } };
        } else if (kind === "sumRow" && colNumber === 1) {
          cell.font = { bold: true };
        } else if (kind === "totalRow" && colNumber === 1) {
          cell.font = { bold: true, color: { argb: "FF1e3a8a" } };
        }
      });
      continue;
    }

    if (kind === "scale" || kind === "footnote" || kind === "singleLine") {
      row.eachCell({ includeEmpty: true }, (cell) => {
        cell.border = AUDIT_ALL_BORDERS;
        cell.alignment = { vertical: "middle", wrapText: true };
      });
      if (kind === "scale") {
        row.getCell(1).font = { italic: true, color: { argb: "FF475569" } };
      }
      if (kind === "footnote" || kind === "singleLine") {
        try {
          ws.mergeCells(ridx + 1, 1, ridx + 1, numCols);
        } catch {
          /* */
        }
      }
      continue;
    }

    row.eachCell({ includeEmpty: true }, (cell) => {
      cell.border = AUDIT_ALL_BORDERS;
      cell.alignment = { vertical: "middle", wrapText: true };
    });
  }

  for (let c = 1; c <= numCols; c++) {
    let maxLen = 10;
    for (const r of grid) {
      const v = r[c - 1];
      const s = v == null ? "" : String(v);
      if (s.length > maxLen) maxLen = Math.min(s.length, 120);
    }
    const w = Math.min(Math.max(maxLen * 0.09 + 1.5, 12), 52);
    ws.getColumn(c).width = w;
  }
}

async function fillAuditTabularSheet(ws, grid, opts) {
  const headerRows = opts?.headerRows ?? new Set([0]);
  const numRows = grid.length;
  const numCols = grid.reduce(
    (m, r) => Math.max(m, Array.isArray(r) ? r.length : 0),
    0,
  );

  for (let ridx = 0; ridx < numRows; ridx++) {
    const rowArr = grid[ridx];
    const padded = [];
    for (let c = 0; c < numCols; c++) {
      const v = rowArr && rowArr[c];
      padded.push(v === "" || v == null ? "" : v);
    }
    const row = ws.addRow(padded);
    if (headerRows.has(ridx)) row.height = 22;
    row.eachCell({ includeEmpty: true }, (cell) => {
      cell.border = AUDIT_ALL_BORDERS;
      if (headerRows.has(ridx)) {
        cell.fill = AUDIT_HEADER_FILL;
        cell.font = AUDIT_HEADER_FONT;
      }
      cell.alignment = { vertical: "middle", wrapText: true };
    });
  }

  for (let c = 1; c <= numCols; c++) {
    let maxLen = 8;
    for (let ridx = 0; ridx < numRows; ridx++) {
      const rowArr = grid[ridx];
      const v = rowArr && rowArr[c - 1];
      const s = v == null ? "" : String(v);
      if (s.length > maxLen) maxLen = s.length;
    }
    const cap = c >= 3 ? 56 : 22;
    const w = Math.min(Math.max(maxLen * 0.09 + 1.8, 11), cap);
    ws.getColumn(c).width = w;
  }
}

async function writeAuditWorkbookBuffer({ ozet, detayWithHeader, eksikWithHeader }) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Assistant";
  const used = new Set();

  const nmOzet = uniqueSheetName("Özet", used);
  const wsOzet = wb.addWorksheet(nmOzet);
  await fillAuditOzetWorksheet(wsOzet, ozet);

  const nmDetay = uniqueSheetName("Detay", used);
  const wsDetay = wb.addWorksheet(nmDetay);
  await fillAuditTabularSheet(wsDetay, detayWithHeader, {
    headerRows: new Set([0]),
  });

  const nmEksik = uniqueSheetName("Tam puan alınmayan", used);
  const wsEksik = wb.addWorksheet(nmEksik);
  await fillAuditTabularSheet(wsEksik, eksikWithHeader, {
    headerRows: new Set([0]),
  });

  const buf = await wb.xlsx.writeBuffer();
  return Buffer.from(buf);
}

ipcMain.handle("save-xlsx-desktop", async (_event, payload) => {
  const { grid, fileName, sheets, exportKind, auditBook } = payload || {};

  let outBuffer;

  if (exportKind === "monthlyStock") {
    if (!Array.isArray(grid) || !grid.length) {
      throw new Error("Kaydedilecek tablo boş.");
    }
    outBuffer = await writeStyledReportXlsxBuffer(grid, "Stok Kapanış");
  } else if (exportKind === "dayEndSummary") {
    if (!Array.isArray(grid) || !grid.length) {
      throw new Error("Kaydedilecek tablo boş.");
    }
    outBuffer = await writeStyledReportXlsxBuffer(grid, "Gün Sonu Özet");
  } else if (exportKind === "auditWorkbook") {
    const ozet = auditBook?.ozet;
    const detayWithHeader = auditBook?.detayWithHeader;
    const eksikWithHeader = auditBook?.eksikWithHeader;
    if (
      !Array.isArray(ozet) ||
      !Array.isArray(detayWithHeader) ||
      !Array.isArray(eksikWithHeader)
    ) {
      throw new Error("Denetim kitabı verisi eksik.");
    }
    outBuffer = await writeAuditWorkbookBuffer({
      ozet,
      detayWithHeader,
      eksikWithHeader,
    });
  } else {
    const wb = xlsx.utils.book_new();
    let hasData = false;

    if (Array.isArray(sheets) && sheets.length > 0) {
      const used = new Set();
      for (let i = 0; i < sheets.length; i++) {
        const s = sheets[i];
        const data = s && Array.isArray(s.data) ? s.data : [];
        if (!data.length) continue;
        hasData = true;
        const nm = uniqueSheetName(s.name || `Sayfa${i + 1}`, used);
        const ws = xlsx.utils.aoa_to_sheet(data);
        xlsx.utils.book_append_sheet(wb, ws, nm);
      }
    }

    if (!hasData && Array.isArray(grid) && grid.length) {
      const ws = xlsx.utils.aoa_to_sheet(grid);
      xlsx.utils.book_append_sheet(wb, ws, "Sheet1");
      hasData = true;
    }

    if (!hasData) {
      throw new Error("Kaydedilecek tablo boş.");
    }

    outBuffer = xlsx.write(wb, { type: "buffer", bookType: "xlsx" });
  }

  let base = sanitizeDesktopFileName(fileName || "cikti");
  if (!base.toLowerCase().endsWith(".xlsx")) {
    base = `${base}.xlsx`;
  }

  const desktop = app.getPath("desktop");
  await fs.ensureDir(desktop);

  let outPath = path.join(desktop, base);
  let attempt = 1;
  const stem = base.replace(/\.xlsx$/i, "");
  while (fileExists(outPath) && attempt < 500) {
    outPath = path.join(desktop, `${stem}_${attempt}.xlsx`);
    attempt += 1;
  }

  await fs.writeFile(outPath, outBuffer);

  const writtenName = path.basename(outPath);
  sendStatus(`Masaüstüne kaydedildi: ${writtenName}`, "success");
  return { filePath: outPath, fileName: writtenName };
});

app.whenReady().then(async () => {
  // Windows/Linux: default menü çubuğu (File / View / DevTools) kaldırılır.
  if (process.platform !== "darwin") {
    Menu.setApplicationMenu(null);
  }

  await loadState();
  createWindow();
  setupAutoUpdater();

  // Optional usability improvement:
  // If state is empty and default target exists, keep it empty until user picks,
  // but we can auto-suggest. For now: do not auto-set to follow "first setup required".
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});

app.on("before-quit", async () => {
  // Nothing to cleanup.
});

