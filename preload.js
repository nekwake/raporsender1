const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("bridgeApi", {
  getConfig: () => ipcRenderer.invoke("get-config"),
  chooseTargetFolder: () => ipcRenderer.invoke("choose-target-folder"),
  parseFileBuffer: (payload) => ipcRenderer.invoke("parse-file-buffer", payload),
  sendGrid: (payload) => ipcRenderer.invoke("send-grid", payload),
  getLastSent: () => ipcRenderer.invoke("last-sent"),
  onStatus: (callback) => {
    ipcRenderer.on("status", (_event, payload) => callback(payload));
  },
});
