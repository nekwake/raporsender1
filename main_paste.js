const { app, BrowserWindow, dialog, ipcMain } = require("electron");
const path = require("path");
const os = require("os");
const fs = require("fs-extra");
const crypto = require("crypto");

const xlsx = require("xlsx");
const pdfParse = require("pdf-parse");

const APP_DIR = app.getPath("userData");
const STATE_FILE = path.join(APP_DIR, "state.json");

const DEFAULT_TARGET_PATH = path.join(os.homedir(), "Desktop", "merkeze-gider");

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
  const { grid, suggestedBaseName } = payload || {};

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

  // Export normalizedGrid to a single XLSX file.
  const ws = xlsx.utils.aoa_to_sheet(normalizedGrid);
  const wb = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(wb, ws, "Sheet1");
  const outBuffer = xlsx.write(wb, { type: "buffer", bookType: "xlsx" });

  const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
  const base = sanitizeFileBaseName(suggestedBaseName || "rapor");

  let fileName = `${base}_${timestamp}.xlsx`;
  let attempt = 1;
  while (fileExists(path.join(state.targetPath, fileName)) && attempt < 1000) {
    fileName = `${base}_${timestamp}_${attempt}.xlsx`;
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

app.whenReady().then(async () => {
  await loadState();
  createWindow();

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

