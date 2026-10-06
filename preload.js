/**
 * Preload bridge for Excel Helper.
 *
 * Exposes the minimum the UI needs: the embedded server's address (port +
 * one-time token, so fetch calls go to the right place securely) and the
 * native Open/Save dialogs. No Node APIs leak into the page.
 */

"use strict";

const { contextBridge, ipcRenderer } = require("electron");

const argValue = (prefix) => {
  const arg = (process.argv || []).find((a) => String(a).startsWith(`${prefix}=`));
  return arg ? arg.split("=").slice(1).join("=") : null;
};

const port = Number(argValue("--excelhelper-port"));
const token = argValue("--excelhelper-token");

contextBridge.exposeInMainWorld("ExcelHelperAPI", {
  port: Number.isFinite(port) && port > 0 ? port : null,
  token: token || null,
  openFile: () => ipcRenderer.invoke("open-file"),
  saveCurrent: () => ipcRenderer.invoke("save-current"),
});
