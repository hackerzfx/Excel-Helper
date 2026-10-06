/**
 * Excel Helper - Electron main process.
 *
 * Starts the embedded API server on a free 127.0.0.1 port, opens the app
 * window, and provides the native Open/Save dialogs through IPC.
 */

"use strict";

const { app, BrowserWindow, ipcMain, dialog } = require("electron");
const crypto = require("crypto");
const http = require("http");
const net = require("net");
const path = require("path");
const fs = require("fs");

const SMOKE = process.env.EXCELHELPER_SMOKE === "1";

let mainWindow = null;

function getFreePort() {
  return new Promise((resolve, reject) => {
    const probe = net.createServer();
    probe.once("error", reject);
    probe.listen(0, "127.0.0.1", () => {
      const { port } = probe.address();
      probe.close(() => resolve(port));
    });
  });
}

function registerIpc(store) {
  ipcMain.handle("open-file", async () => {
    if (!mainWindow) {
      return { canceled: false, error: "Main window not ready" };
    }
    let result;
    try {
      result = await dialog.showOpenDialog(mainWindow, {
        title: "Open an Excel file",
        properties: ["openFile"],
        filters: [{ name: "Excel Workbook", extensions: ["xlsx"] }],
      });
    } catch (err) {
      return { canceled: false, error: err.message };
    }
    if (result.canceled || !result.filePaths.length) return { canceled: true };
    const filePath = result.filePaths[0];
    try {
      const data = await fs.promises.readFile(filePath);
      return {
        name: path.basename(filePath),
        path: filePath,
        data: data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength),
      };
    } catch (err) {
      return { canceled: false, error: `Could not read the file: ${err.message}` };
    }
  });

  ipcMain.handle("save-current", async () => {
    if (!store.hasFile()) return { saved: false, error: "Nothing to save yet." };
    const result = await dialog.showSaveDialog(mainWindow, {
      title: "Save your Excel file",
      defaultPath: store.filename || "workbook.xlsx",
      filters: [{ name: "Excel Workbook", extensions: ["xlsx"] }],
    });
    if (result.canceled || !result.filePath) return { saved: false };
    try {
      let target = result.filePath;
      if (!target.toLowerCase().endsWith(".xlsx")) target += ".xlsx";
      await fs.promises.copyFile(store.currentFile, target);
      return { saved: true, path: target };
    } catch (err) {
      // Very plausible for beginners: saving over the same file they still
      // have open in Excel. The UI must survive it.
      return {
        saved: false,
        error: "Could not save there - the file may be open in Excel. Try a different name or folder.",
      };
    }
  });
}

async function start() {
  const port = await getFreePort();
  const token = crypto.randomBytes(16).toString("hex");
  const { createApp } = require("./server");
  const { WorkbookStore } = require("./excel/workbook_store");

  const store = new WorkbookStore(__dirname);
  await store.init();

  const server = createApp(store, { token }).listen(port, "127.0.0.1");
  registerIpc(store);

  if (SMOKE) {
    // Headless boot check: is the API alive? Then quit without a window.
    await new Promise((resolve, reject) => {
      server.once("listening", resolve);
      server.once("error", reject);
    });
    http
      .get({ host: "127.0.0.1", port, path: `/api/state?token=${token}` }, (res) => {
        let body = "";
        res.on("data", (chunk) => (body += chunk));
        res.on("end", () => {
          console.log(`SMOKE_OK port=${port} state=${body.slice(0, 60)}`);
          app.quit();
        });
      })
      .on("error", (err) => {
        console.error("SMOKE_FAIL", err.message);
        app.exit(1);
      });
    return;
  }

  mainWindow = new BrowserWindow({
    width: 1280,
    height: 840,
    minWidth: 1000,
    minHeight: 640,
    title: "Excel Helper",
    backgroundColor: "#f2f5f3",
    autoHideMenuBar: true,
    show: false,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      additionalArguments: [`--excelhelper-port=${port}`, `--excelhelper-token=${token}`],
      spellcheck: false,
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  });
  mainWindow.setMenuBarVisibility(false);
  // The renderer only ever shows local content - keep it that way.
  mainWindow.webContents.setWindowOpenHandler(() => ({ action: "deny" }));
  mainWindow.webContents.on("will-navigate", (event) => event.preventDefault());
  mainWindow.once("ready-to-show", () => mainWindow.show());
  await mainWindow.loadFile(path.join(__dirname, "index.html"));
}

app.whenReady().then(start).catch((err) => {
  console.error(err);
  app.exit(1);
});

app.on("window-all-closed", () => app.quit());
