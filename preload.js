const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("bridgeApi", {
  getConfig: () => ipcRenderer.invoke("get-config"),
  checkForUpdates: () => ipcRenderer.invoke("check-for-updates"),
  setCloudConfig: (payload) => ipcRenderer.invoke("set-cloud-config", payload),
  resetCloudConfig: () => ipcRenderer.invoke("reset-cloud-config"),
  parseFileBuffer: (payload) => ipcRenderer.invoke("parse-file-buffer", payload),
  sendGrid: (payload) => ipcRenderer.invoke("send-grid", payload),
  saveXlsxDesktop: (payload) => ipcRenderer.invoke("save-xlsx-desktop", payload),
  getLastSent: () => ipcRenderer.invoke("last-sent"),
  onStatus: (callback) => {
    ipcRenderer.on("status", (_event, payload) => callback(payload));
  },
});
