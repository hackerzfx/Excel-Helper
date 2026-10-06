/**
 * WorkbookStore - holds the "current file" for Excel Helper.
 *
 * Responsibilities:
 * - keep the working copy (.work/current.xlsx) plus metadata (original name,
 *   source path on disk, active sheet for the preview)
 * - snapshot before every change so Undo can go back (last 5 steps)
 * - build the JSON preview state the UI renders
 * - watch the user's original file: if it changes on disk (they edited it in
 *   Excel meanwhile), re-import and broadcast the fresh preview (live view)
 */

"use strict";

const fs = require("fs");
const fsp = fs.promises;
const path = require("path");

const ExcelJS = require("exceljs");
const practiceData = require("./practice_data");
const { resolveValue, colToLetter, effectiveCellValue } = require("./formula_eval");

const PREVIEW_ROWS = 150;
const PREVIEW_COLS = 40;
const MAX_HISTORY = 5;

class WorkbookStore {
  constructor(baseDir) {
    this.workDir = path.join(baseDir, ".work");
    this.currentFile = path.join(this.workDir, "current.xlsx");
    this.historyDir = path.join(this.workDir, "history");
    this.sourcePath = null;
    this.filename = null;
    this.activeSheetName = null;
    this.sourceWatcher = null;
    this.watchDebounce = null;
    this.listeners = new Set();
  }

  async init() {
    await fsp.mkdir(this.historyDir, { recursive: true });
    // The working file SURVIVES restarts so users continue where they left
    // off; only the undo history is cleared.
    for (const name of await fsp.readdir(this.historyDir)) {
      await fsp.rm(path.join(this.historyDir, name), { force: true }).catch(() => {});
    }
    try {
      const saved = JSON.parse(await fsp.readFile(path.join(this.workDir, "state.json"), "utf8"));
      if (saved.filename) this.filename = saved.filename;
      if (saved.activeSheetName) this.activeSheetName = saved.activeSheetName;
      if (saved.sourcePath) this.sourcePath = saved.sourcePath;
    } catch (err) {
      /* first run - nothing to restore */
    }
    if (this.sourcePath) this.startWatching();
  }

  async saveMeta() {
    try {
      await fsp.writeFile(
        path.join(this.workDir, "state.json"),
        JSON.stringify({ filename: this.filename, activeSheetName: this.activeSheetName, sourcePath: this.sourcePath }, null, 1),
        "utf8",
      );
    } catch (err) {
      /* metadata persistence is best-effort */
    }
  }

  hasFile() {
    return fs.existsSync(this.currentFile);
  }

  subscribe(fn) {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  }

  broadcast(payload) {
    for (const fn of this.listeners) {
      try {
        fn(payload);
      } catch (err) {
        /* a dead SSE connection must never break the app */
      }
    }
  }

  /* ---------------- snapshots / undo ---------------- */

  snapshotMetaPath() {
    return path.join(this.historyDir, "meta.json");
  }

  async readSnapshotMeta() {
    try {
      return JSON.parse(await fsp.readFile(this.snapshotMetaPath(), "utf8"));
    } catch (err) {
      return {};
    }
  }

  async writeSnapshotMeta(meta) {
    await fsp.writeFile(this.snapshotMetaPath(), JSON.stringify(meta, null, 1), "utf8").catch(() => {});
  }

  async snapshot() {
    if (!this.hasFile()) return null;
    const stamp = Date.now();
    const name = `snap_${stamp}.xlsx`;
    await fsp.copyFile(this.currentFile, path.join(this.historyDir, name));
    const meta = await this.readSnapshotMeta();
    meta[name] = {
      filename: this.filename,
      activeSheetName: this.activeSheetName,
      sourcePath: this.sourcePath,
    };
    await this.writeSnapshotMeta(meta);
    const snaps = (await fsp.readdir(this.historyDir)).filter((n) => n.startsWith("snap_")).sort();
    while (snaps.length > MAX_HISTORY) {
      const oldest = snaps.shift();
      delete meta[oldest];
      await fsp.rm(path.join(this.historyDir, oldest), { force: true }).catch(() => {});
    }
    await this.writeSnapshotMeta(meta);
    return name;
  }

  async deleteSnapshot(name) {
    if (!name) return;
    const meta = await this.readSnapshotMeta();
    delete meta[name];
    await this.writeSnapshotMeta(meta);
    await fsp.rm(path.join(this.historyDir, name), { force: true }).catch(() => {});
  }

  async undo() {
    const snaps = (await fsp.readdir(this.historyDir)).filter((n) => n.startsWith("snap_")).sort();
    if (!snaps.length) return false;
    const latest = snaps[snaps.length - 1];
    await fsp.copyFile(path.join(this.historyDir, latest), this.currentFile);
    await fsp.rm(path.join(this.historyDir, latest), { force: true }).catch(() => {});
    const meta = await this.readSnapshotMeta();
    const saved = meta[latest];
    delete meta[latest];
    await this.writeSnapshotMeta(meta);
    if (saved) {
      if (saved.filename) this.filename = saved.filename;
      if (saved.activeSheetName) this.activeSheetName = saved.activeSheetName;
      this.sourcePath = saved.sourcePath || null;
      this.startWatching();
      await this.saveMeta();
    }
    return true;
  }

  /* ---------------- import / practice ---------------- */

  async importBuffer(buffer, filename, sourcePath = null) {
    // Validate before touching the current file.
    const probe = new ExcelJS.Workbook();
    try {
      await probe.xlsx.load(buffer);
    } catch (err) {
      const e = new Error(
        "That file could not be opened as an Excel workbook. Is it a real .xlsx file?"
      );
      e.status = 400;
      throw e;
    }
    await this.stopWatching();
    await this.snapshot();
    await fsp.writeFile(this.currentFile, Buffer.from(buffer));
    this.filename = path.basename(filename || "workbook.xlsx");
    this.sourcePath = sourcePath || null;
    this.activeSheetName = probe.worksheets[0] ? probe.worksheets[0].name : null;
    this.startWatching();
    await this.saveMeta();
  }

  async loadPractice(kind) {
    const wb = kind === "ledger" ? practiceData.makeLedgerWorkbook() : practiceData.makeInvoiceWorkbook();
    const buffer = await wb.xlsx.writeBuffer();
    const name = kind === "ledger" ? "Practice Ledger.xlsx" : "Practice Invoice.xlsx";
    await this.importBuffer(Buffer.from(buffer), name, null);
  }

  /* ---------------- mutation (operations) ---------------- */

  /**
   * Load the current workbook, let `mutator(workbook, activeWorksheet)` change
   * it, then save. A snapshot of the pre-change file is taken first so Undo
   * always works. Returns mutator's return value.
   */
  async withWorkbook(mutator) {
    if (!this.hasFile()) {
      const e = new Error("Open an Excel file or load practice data first (buttons at the top).");
      e.status = 400;
      throw e;
    }
    const snapName = await this.snapshot();
    try {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(await fsp.readFile(this.currentFile));
      let ws = (this.activeSheetName && wb.getWorksheet(this.activeSheetName)) || wb.worksheets[0];
      if (!ws) {
        const e = new Error("This workbook has no sheets.");
        e.status = 400;
        throw e;
      }
      this.activeSheetName = ws.name;
      const result = await mutator(wb, ws);
      const buffer = await wb.xlsx.writeBuffer();
      await fsp.writeFile(this.currentFile, Buffer.from(buffer));
      return result;
    } catch (err) {
      // The change never landed - drop the snapshot so Undo stays meaningful.
      await this.deleteSnapshot(snapName);
      throw err;
    }
  }

  /**
   * Switch the sheet shown in the preview. Pure UI state: no snapshot, no
   * file rewrite, no Undo step.
   */
  async selectSheet(name) {
    if (!this.hasFile()) {
      const e = new Error("No file open.");
      e.status = 400;
      throw e;
    }
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(await fsp.readFile(this.currentFile));
    const ws = wb.getWorksheet(String(name));
    if (!ws) {
      const e = new Error("No such sheet.");
      e.status = 400;
      throw e;
    }
    this.activeSheetName = ws.name;
    await this.saveMeta();
    return `Switched to sheet '${ws.name}'.`;
  }

  /* ---------------- preview state ---------------- */

  async buildState() {
    if (!this.hasFile()) return { has_file: false };
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(await fsp.readFile(this.currentFile));
    const ws = (this.activeSheetName && wb.getWorksheet(this.activeSheetName)) || wb.worksheets[0];
    if (!ws) return { has_file: false };

    this.activeSheetName = ws.name;
    const maxRow = Math.min(ws.rowCount, PREVIEW_ROWS);
    const maxCol = Math.min(ws.columnCount, PREVIEW_COLS);

    const grid = [];
    let formulaCount = 0;
    for (let row = 1; row <= maxRow; row += 1) {
      const rowCells = [];
      for (let col = 1; col <= maxCol; col += 1) {
        const raw = ws.getCell(row, col).value;
        const isFormula = raw !== null && typeof raw === "object" &&
          (raw.formula !== undefined || raw.sharedFormula !== undefined);
        if (isFormula) {
          formulaCount += 1;
          // Show the computed value in the preview (like Excel does); the
          // formula itself rides along as a hover tooltip.
          let text;
          let fm;
          if (raw.formula !== undefined) fm = `=${raw.formula}`;
          try {
            text = previewText(effectiveCellValue(ws, row, col));
          } catch (err) {
            text = fm !== undefined ? fm : "";
          }
          rowCells.push({ v: text, f: true, fm });
        } else {
          rowCells.push({
            v: safePreview(resolveValue, ws, raw),
            f: false,
          });
        }
      }
      grid.push(rowCells);
    }
    await this.saveMeta();

    const snaps = fs.existsSync(this.historyDir)
      ? (await fsp.readdir(this.historyDir)).filter((n) => n.startsWith("snap_"))
      : [];

    const header = [];
    const colLetters = [];
    for (let col = 1; col <= maxCol; col += 1) {
      header.push(safePreview(resolveValue, ws, ws.getCell(1, col).value));
      colLetters.push(colToLetter(col));
    }

    return {
      has_file: true,
      filename: this.filename || "workbook.xlsx",
      sheets: wb.worksheets.map((s) => s.name),
      active_sheet: ws.name,
      header,
      colLetters,
      maxRow: ws.rowCount,
      maxCol: ws.columnCount,
      truncatedRows: ws.rowCount > PREVIEW_ROWS,
      truncatedCols: ws.columnCount > PREVIEW_COLS,
      formulaCount,
      canUndo: snaps.length > 0,
      headerStyled: _headerStyled(ws),
      grid,
    };
  }

  /* ---------------- live watching of the source file ---------------- */

  startWatching() {
    if (!this.sourcePath || !fs.existsSync(this.sourcePath)) return;
    try {
      this.sourceWatcher = fs.watch(this.sourcePath, () => this.onSourceChanged());
    } catch (err) {
      this.sourceWatcher = null;
    }
  }

  stopWatching() {
    if (this.sourceWatcher) {
      try {
        this.sourceWatcher.close();
      } catch (err) {
        /* already closed */
      }
      this.sourceWatcher = null;
    }
    if (this.watchDebounce) {
      clearTimeout(this.watchDebounce);
      this.watchDebounce = null;
    }
  }

  onSourceChanged() {
    if (this.watchDebounce) clearTimeout(this.watchDebounce);
    this.watchDebounce = setTimeout(async () => {
      this.watchDebounce = null;
      if (!this.sourcePath || !fs.existsSync(this.sourcePath)) return;
      try {
        const buffer = await fsp.readFile(this.sourcePath);
        // Re-import silently: snapshot current, swap bytes, keep name/source.
        const previousSource = this.sourcePath;
        const previousName = this.filename;
        await this.importBuffer(buffer, previousName, previousSource);
        const state = await this.buildState();
        this.broadcast({
          state,
          notice: { type: "success", text: "The file changed on disk - preview refreshed automatically." },
        });
      } catch (err) {
        this.broadcast({
          notice: {
            type: "error",
            text: "The file changed on disk, but it is locked right now (probably open in Excel). Close it there and the preview will refresh.",
          },
        });
      }
    }, 800);
  }
}

function safePreview(resolveValue, ws, raw) {
  try {
    return previewText(resolveValue(ws, raw, 6));
  } catch (err) {
    return raw !== null && typeof raw === "object" && raw.error !== undefined
      ? `#${raw.error}`
      : "";
  }
}

function previewText(value) {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) {
    const dd = String(value.getDate()).padStart(2, "0");
    const mm = String(value.getMonth() + 1).padStart(2, "0");
    return `${dd}-${mm}-${value.getFullYear()}`;
  }
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "object") {
    if (Array.isArray(value.richText)) return value.richText.map((p) => p.text).join("");
    if (value.text !== undefined) return String(value.text);
    return "";
  }
  if (typeof value === "number") return Number.isInteger(value) ? String(value) : String(value);
  return String(value);
}

/** True when row 1 looks styled by makeHeaderRow (bold + pattern fill). */
function _headerStyled(ws) {
  for (let col = 1; col <= Math.min(ws.columnCount, 20); col += 1) {
    const cell = ws.getCell(1, col);
    if (cell.font && cell.font.bold && cell.fill && cell.fill.type === "pattern") {
      return true;
    }
  }
  return false;
}

module.exports = { WorkbookStore, PREVIEW_ROWS, PREVIEW_COLS };
