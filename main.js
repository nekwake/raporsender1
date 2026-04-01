const { app, BrowserWindow, dialog, ipcMain } = require("electron");
const path = require("path");
const os = require("os");
const crypto = require("crypto");
const fs = require("fs-extra");
const chokidar = require("chokidar");

const APP_DIR = app.getPath("userData");
const STATE_FILE = path.join(APP_DIR, "state.json");
const defaultWatchPath = path.join(os.homedir(), "Desktop", "raporgonder");
const defaultTargetPath = path.join(os.homedir(), "Desktop", "gidecek-rapor");
const ALLOWED_EXTENSIONS = new Set([".xls", ".xlsx", ".xlsm", ".xlsb", ".csv"]);

let mainWindow = null;
let watcher = null;
let state = {
  watchPath: defaultWatchPath,
  drivePath: defaultTargetPath,
  sentFiles: {},
};

async function loadState() {
  await fs.ensureDir(APP_DIR);
  if (await fs.pathExists(STATE_FILE)) {
    const fileState = await fs.readJson(STATE_FILE);
    state = {
      ...state,
      ...fileState,
      sentFiles: fileState.sentFiles || {},
    };
  }
  await fs.ensureDir(state.watchPath);
  await fs.ensureDir(state.drivePath);
  await persistState();
}

async function persistState() {
  await fs.writeJson(STATE_FILE, state, { spaces: 2 });
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 760,
    height: 560,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  mainWindow.loadFile(path.join(__dirname, "index.html"));
}

function fileHash(filePath, stat) {
  return crypto
    .createHash("sha256")
    .update(`${filePath}|${stat.size}|${stat.mtimeMs}`)
    .digest("hex");
}

function sendStatus(message, level = "info") {
  if (!mainWindow) return;
  mainWindow.webContents.send("status", { message, level, at: Date.now() });
}

function isEligibleFile(fileName) {
  if (!fileName || fileName.startsWith(".")) return false;
  if (fileName.startsWith("~$")) return false;
  const ext = path.extname(fileName).toLowerCase();
  return ALLOWED_EXTENSIONS.has(ext);
}

function normalizeForCompare(folderPath) {
  const normalized = path.resolve(folderPath);
  if (process.platform === "win32") return normalized.toLowerCase();
  return normalized;
}

function foldersAreSame(sourcePath, targetPath) {
  return normalizeForCompare(sourcePath) === normalizeForCompare(targetPath);
}

async function sendPendingFiles() {
  if (!state.drivePath) {
    sendStatus("Hedef klasör seçili değil.", "warn");
    return { sent: 0, skipped: 0, failed: 0, ignored: 0, scanned: 0 };
  }

  if (foldersAreSame(state.watchPath, state.drivePath)) {
    sendStatus("Kaynak ve hedef klasör aynı olamaz.", "error");
    return { sent: 0, skipped: 0, failed: 1, ignored: 0, scanned: 0 };
  }

  await fs.ensureDir(state.watchPath);
  await fs.ensureDir(state.drivePath);

  const entries = await fs.readdir(state.watchPath);
  let sent = 0;
  let skipped = 0;
  let failed = 0;
  let ignored = 0;
  let scanned = 0;

  for (const entry of entries) {
    const source = path.join(state.watchPath, entry);
    const stat = await fs.stat(source);
    if (!stat.isFile()) continue;
    scanned += 1;

    if (!isEligibleFile(entry)) {
      ignored += 1;
      continue;
    }

    const key = fileHash(source, stat);
    if (state.sentFiles[key]) {
      skipped += 1;
      continue;
    }

    const target = path.join(state.drivePath, entry);
    try {
      if (await fs.pathExists(target)) {
        skipped += 1;
        continue;
      }

      await fs.copy(source, target, { overwrite: false, errorOnExist: true });
      state.sentFiles[key] = {
        name: entry,
        sentAt: new Date().toISOString(),
        source,
      };
      sent += 1;
      sendStatus(`Gönderildi: ${entry}`, "success");
    } catch (err) {
      failed += 1;
      sendStatus(`Hata: ${entry} - ${err.message}`, "error");
    }
  }

  if (scanned === 0) {
    sendStatus("Kaynak klasörde dosya bulunamadı.", "info");
  }

  await persistState();
  return { sent, skipped, failed, ignored, scanned };
}

function startWatcher() {
  if (watcher) {
    watcher.close();
  }

  watcher = chokidar.watch(state.watchPath, {
    ignoreInitial: true,
    awaitWriteFinish: { stabilityThreshold: 1200, pollInterval: 200 },
  });

  watcher.on("add", async () => {
    const result = await sendPendingFiles();
    sendStatus(
      `Otomatik kontrol bitti. Yeni: ${result.sent}, Atlanan: ${result.skipped}, Hata: ${result.failed}`,
      "info"
    );
  });

  watcher.on("error", (err) => {
    sendStatus(`İzleyici hatası: ${err.message}`, "error");
  });
}

ipcMain.handle("get-config", async () => {
  return {
    watchPath: state.watchPath,
    drivePath: state.drivePath,
    sentCount: Object.keys(state.sentFiles).length,
  };
});

ipcMain.handle("choose-watch-folder", async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ["openDirectory", "createDirectory"],
  });
  if (result.canceled || !result.filePaths[0]) return null;
  if (foldersAreSame(result.filePaths[0], state.drivePath)) {
    sendStatus("Kaynak ve hedef klasör aynı seçilemez.", "error");
    return null;
  }

  state.watchPath = result.filePaths[0];
  await fs.ensureDir(state.watchPath);
  await persistState();
  startWatcher();
  return state.watchPath;
});

ipcMain.handle("choose-drive-folder", async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ["openDirectory", "createDirectory"],
  });
  if (result.canceled || !result.filePaths[0]) return null;
  if (foldersAreSame(state.watchPath, result.filePaths[0])) {
    sendStatus("Hedef ve kaynak klasör aynı seçilemez.", "error");
    return null;
  }

  state.drivePath = result.filePaths[0];
  await fs.ensureDir(state.drivePath);
  await persistState();
  return state.drivePath;
});

ipcMain.handle("send-now", async () => {
  const result = await sendPendingFiles();
  sendStatus(
    `Manuel gönderim bitti. Yeni: ${result.sent}, Atlanan: ${result.skipped}, Yoksayılan: ${result.ignored}, Hata: ${result.failed}`,
    "info"
  );
  return result;
});

ipcMain.handle("last-sent", async () => {
  const latest = Object.values(state.sentFiles)
    .sort((a, b) => new Date(b.sentAt).getTime() - new Date(a.sentAt).getTime())
    .slice(0, 10);
  return latest;
});

app.whenReady().then(async () => {
  await loadState();
  createWindow();
  startWatcher();
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});

app.on("before-quit", async () => {
  if (watcher) await watcher.close();
});
