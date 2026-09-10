const { app, BrowserWindow, Menu, dialog, ipcMain, Tray, shell } = require("electron");
const fs = require("fs");
const path = require("path");
const { convertFile } = require("./converter");
const { moveFolder } = require("./folder-mover");

const appName = "Acacia Flow";
const appRoot = app.getAppPath();
const gotSingleInstanceLock = app.requestSingleInstanceLock();
let outputDir = null;
let mainWindow = null;
let tray = null;
let isQuitting = false;
let settingsPath = null;
const defaultSettings = {
  autoLaunch: true,
  closeBehavior: "tray",
  defaultOutputDir: "",
  wallpaperEnabled: true,
  animationsEnabled: true,
  confirmBeforeConvert: false,
  showDetailedErrors: true,
  rememberLastTool: true,
  advanced: {
    conflictMode: "rename",
    openOutputFolder: false,
    clearCompleted: false,
    openOutputFile: false
  }
};

let settings = { ...defaultSettings };

function mergeSettings(nextSettings = {}) {
  return {
    ...defaultSettings,
    ...nextSettings,
    advanced: {
      ...defaultSettings.advanced,
      ...(nextSettings.advanced || {})
    }
  };
}

function loadSettings() {
  settingsPath = path.join(app.getPath("userData"), "settings.json");
  try {
    settings = mergeSettings(JSON.parse(fs.readFileSync(settingsPath, "utf8")));
  } catch (_err) {
    settings = mergeSettings();
  }
  outputDir = settings.defaultOutputDir || null;
  applyAutoLaunch();
}

function saveSettings(nextSettings) {
  settings = mergeSettings(nextSettings);
  outputDir = settings.defaultOutputDir || null;
  fs.mkdirSync(path.dirname(settingsPath), { recursive: true });
  fs.writeFileSync(settingsPath, JSON.stringify(settings, null, 2), "utf8");
  applyAutoLaunch();
  syncTrayWithSettings();
  return settings;
}

function applyAutoLaunch() {
  app.setLoginItemSettings({
    openAtLogin: Boolean(settings.autoLaunch),
    path: process.execPath
  });
}

function createTray() {
  if (tray) return tray;
  tray = new Tray(path.join(appRoot, "assets", "acacia_flow.ico"));
  tray.setToolTip(appName);
  tray.setContextMenu(Menu.buildFromTemplate([
    { label: "打开 Acacia Flow", click: () => showMainWindow() },
    { type: "separator" },
    {
      label: "退出",
      click: () => {
        isQuitting = true;
        destroyTray();
        app.quit();
      }
    }
  ]));
  tray.on("double-click", () => showMainWindow());
  return tray;
}

function destroyTray() {
  if (!tray) return;
  tray.destroy();
  tray = null;
}

function syncTrayWithSettings() {
  if (!app.isReady()) return;
  if (settings.closeBehavior === "tray" && !isQuitting) {
    createTray();
    return;
  }
  destroyTray();
}

function showMainWindow() {
  if (!mainWindow) {
    createWindow();
    return;
  }
  mainWindow.show();
  if (mainWindow.isMinimized()) mainWindow.restore();
  mainWindow.focus();
}

function createWindow() {
  const win = new BrowserWindow({
    width: 1120,
    height: 760,
    minWidth: 900,
    minHeight: 620,
    title: appName,
    frame: false,
    icon: path.join(appRoot, "assets", "acacia_flow.ico"),
    webPreferences: {
      preload: path.join(appRoot, "src", "main", "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false
    }
  });
  mainWindow = win;

  Menu.setApplicationMenu(null);
  win.webContents.session.on("will-download", (_event, item) => {
    if (!outputDir) return;
    item.setSavePath(path.join(outputDir, item.getFilename()));
  });
  win.on("close", event => {
    if (settings.closeBehavior === "tray" && !isQuitting) {
      event.preventDefault();
      createTray();
      win.hide();
    }
  });
  win.on("closed", () => {
    if (mainWindow === win) mainWindow = null;
  });
  win.loadFile(path.join(appRoot, "web", "index.html"));
}

ipcMain.handle("choose-output-dir", async () => {
  const result = await dialog.showOpenDialog({
    title: "选择输出目录",
    properties: ["openDirectory", "createDirectory"]
  });
  if (result.canceled || !result.filePaths.length) return null;
  outputDir = result.filePaths[0];
  return outputDir;
});

ipcMain.handle("get-settings", async () => settings);

ipcMain.handle("save-settings", async (_event, nextSettings) => saveSettings(nextSettings));

ipcMain.handle("reset-settings", async () => saveSettings(defaultSettings));

ipcMain.handle("convert-file", async (_event, payload) => {
  return convertFile({ ...payload, outputDir: payload.outputDir || outputDir });
});

ipcMain.handle("open-output-dir", async (_event, dirPath) => {
  if (!dirPath) return "missing";
  return shell.openPath(dirPath);
});

ipcMain.handle("open-path", async (_event, targetPath) => {
  if (!targetPath) return "missing";
  return shell.openPath(targetPath);
});

ipcMain.handle("choose-move-source", async () => {
  const result = await dialog.showOpenDialog({
    title: "选择要移动的文件夹",
    properties: ["openDirectory"]
  });
  return result.canceled || !result.filePaths.length ? null : result.filePaths[0];
});

ipcMain.handle("choose-move-destination", async () => {
  const result = await dialog.showOpenDialog({
    title: "选择新的存储位置",
    properties: ["openDirectory", "createDirectory"]
  });
  return result.canceled || !result.filePaths.length ? null : result.filePaths[0];
});

ipcMain.handle("move-folder", async (_event, payload) => moveFolder(payload || {}));

ipcMain.on("window-control", (event, action) => {
  const win = BrowserWindow.fromWebContents(event.sender);
  if (!win) return;
  if (action === "minimize") win.minimize();
  if (action === "maximize") {
    if (win.isMaximized()) win.unmaximize();
    else win.maximize();
  }
  if (action === "close") {
    if (settings.closeBehavior === "exit") isQuitting = true;
    win.close();
  }
});

if (!gotSingleInstanceLock) {
  app.quit();
} else {
  app.on("second-instance", () => {
    showMainWindow();
  });

  app.on("before-quit", () => {
    isQuitting = true;
    destroyTray();
  });

  app.whenReady().then(() => {
    loadSettings();
    syncTrayWithSettings();
    createWindow();
  });

  app.on("window-all-closed", () => {
    if (process.platform !== "darwin" && (settings.closeBehavior === "exit" || isQuitting)) app.quit();
  });

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
}
