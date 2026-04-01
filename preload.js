const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("bridgeApi", {
  getConfig: () => ipcRenderer.invoke("get-config"),
  chooseWatchFolder: () => ipcRenderer.invoke("choose-watch-folder"),
  chooseDriveFolder: () => ipcRenderer.invoke("choose-drive-folder"),
  sendNow: () => ipcRenderer.invoke("send-now"),
  getLastSent: () => ipcRenderer.invoke("last-sent"),
  onStatus: (callback) => {
    ipcRenderer.on("status", (_event, payload) => callback(payload));
  },
});
