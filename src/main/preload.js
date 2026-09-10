const { contextBridge, ipcRenderer, webUtils } = require("electron");

contextBridge.exposeInMainWorld("acaciaFlow", {
  chooseOutputDir: () => ipcRenderer.invoke("choose-output-dir"),
  getSettings: () => ipcRenderer.invoke("get-settings"),
  saveSettings: settings => ipcRenderer.invoke("save-settings", settings),
  resetSettings: () => ipcRenderer.invoke("reset-settings"),
  convertFile: payload => ipcRenderer.invoke("convert-file", payload),
  openOutputDir: dir => ipcRenderer.invoke("open-output-dir", dir),
  openPath: targetPath => ipcRenderer.invoke("open-path", targetPath),
  chooseMoveSource: () => ipcRenderer.invoke("choose-move-source"),
  chooseMoveDestination: () => ipcRenderer.invoke("choose-move-destination"),
  moveFolder: payload => ipcRenderer.invoke("move-folder", payload),
  getFilePath: file => webUtils.getPathForFile(file),
  windowControl: action => ipcRenderer.send("window-control", action)
});
