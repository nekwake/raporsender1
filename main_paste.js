const { app, BrowserWindow, dialog, ipcMain, Menu } = require("electron");
const path = require("path");
const os = require("os");
const fs = require("fs-extra");
const crypto = require("crypto");
const { autoUpdater } = require("electron-updater");

const xlsx = require("xlsx");
const ExcelJS = require("exceljs");
const pdfParse = require("pdf-parse");

const APP_DIR = app.getPath("userData");
const STATE_FILE = path.join(APP_DIR, "state.json");

const DEFAULT_TARGET_PATH = path.join(os.homedir(), "Desktop", "merkeze-gider");

if (process.platform === "win32") {
  app.setAppUserModelId("com.nekwake.assistant");
}

let mainWindow = null;
let state = {
  targetPath: "",
  sentHashes: {},
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

function sha256(text) {
  return crypto.createHash("sha256").update(text).digest("hex");
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
      sentHashes: fileState.sentHashes || {},
      sentHistory: fileState.sentHistory || [],
    };
  }

  if (state.targetPath) {
    const exists = await fs.pathExists(state.targetPath);
    if (!exists) state.targetPath = "";
  }
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

ipcMain.handle("get-config", async () => {
  const hasTarget = Boolean(state.targetPath);
  return {
    targetPath: state.targetPath,
    hasTarget,
    sentCount: state.sentHistory.length,
  };
});

ipcMain.handle("choose-target-folder", async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ["openDirectory", "createDirectory"],
  });
  if (result.canceled || !result.filePaths[0]) return null;

  state.targetPath = result.filePaths[0];
  await fs.ensureDir(state.targetPath);
  await persistState();
  return state.targetPath;
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

ipcMain.handle("send-grid", async (_event, payload) => {
  const { grid, suggestedBaseName, fileName: preferredStem, exportKind } =
    payload || {};

  if (!state.targetPath) {
    sendStatus("Hedef klasör seçilmedi.", "warn");
    throw new Error("Hedef klasör seçilmedi.");
  }

  if (!Array.isArray(grid) || !grid.length) {
    throw new Error("Gönderilecek veri yok.");
  }

  const normalizedGrid = normalizeGridForHash(grid);
  const comparable = JSON.stringify(normalizedGrid);
  const hash = sha256(comparable);

  const already = state.sentHashes[hash];
  if (already) {
    return { deduped: true, fileName: already.fileName, hash };
  }

  let outBuffer;
  if (exportKind === "dayEndSummary") {
    outBuffer = await writeStyledReportXlsxBuffer(
      normalizedGrid,
      "Gün Sonu Özet",
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

  let attempt = 1;
  const stemForDup = fileName.replace(/\.xlsx$/i, "");
  while (fileExists(path.join(state.targetPath, fileName)) && attempt < 1000) {
    fileName = `${stemForDup}_${attempt}.xlsx`;
    attempt += 1;
  }

  const outPath = path.join(state.targetPath, fileName);
  await fs.writeFile(outPath, outBuffer);

  state.sentHashes[hash] = {
    fileName,
    sentAt: new Date().toISOString(),
  };
  state.sentHistory.unshift({
    name: fileName,
    sentAt: state.sentHashes[hash].sentAt,
    hash,
  });
  state.sentHistory = state.sentHistory.slice(0, 30);

  await persistState();

  sendStatus(`Kaydedildi: ${fileName}`, "success");
  return { deduped: false, fileName, hash };
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

