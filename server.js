/**
 * The embedded API server for Excel Helper.
 *
 * Runs inside the Electron main process on 127.0.0.1:<random port> and also
 * standalone (`node server.js`) for testing. The browser UI talks to it over
 * plain fetch; the contract is:
 *
 *   GET  /api/state      -> {has_file:false} | full preview state
 *   POST /api/upload     -> raw .xlsx bytes, header X-Filename
 *   POST /api/practice   -> {type:"invoice"|"ledger"}
 *   POST /api/operation  -> {op, params}
 *   POST /api/undo
 *   POST /api/sheet      -> {name}
 *   GET  /api/download   -> the current .xlsx
 *   GET  /api/events     -> Server-Sent Events: {state, notice?} pushes
 *
 * Every mutating response carries the refreshed state so the UI updates the
 * preview instantly after every click.
 */

"use strict";

const fs = require("fs");
const path = require("path");

const express = require("express");

const ops = require("./excel/operations");
const { FormulaEvalError } = require("./excel/formula_eval");

const PRACTICE_HINTS = {
  invoice:
    "Practice Invoice created! A good first try: Formulas tab > 'QTY x RATE' with columns B and C " +
    "writing into column D, then 'Totals & stats' > SUM on column D, then Basics tab > 'Make the header row'.",
  ledger:
    "Practice Ledger created! A good first try: Accounts tab > 'Add Balance column' with Debit as " +
    "Amount and Credit as Paid (running balance ON), then Formulas tab > 'Totals & stats' > SUM on the Balance column.",
};

const OPERATIONS = {
  make_header: (p, ws) => ops.makeHeaderRow(ws, p.color || "217346"),
  add_borders: (p, ws) => ops.addBorders(ws, p.range || null),
  number_format: (p, ws) => ops.applyNumberFormat(ws, p.col, p.fmt, p.symbol || "₹"),
  align_column: (p, ws) => ops.alignColumn(ws, p.col, p.mode || "center"),
  autofit: (_p, ws) => ops.autofitColumns(ws),
  style_column: (p, ws) => ops.styleColumn(ws, p.col, !!p.bold, p.color || null),
  multiply: (p, ws) => ops.multiplyColumns(ws, p.col_a, p.col_b, p.result_col, p.header || "Amount"),
  stat: (p, ws) => ops.addStat(ws, p.col, p.stat || "sum"),
  freeze_values: (p, ws) => ops.freezeFormulasAsValues(ws, p.col || null),
  fill_series: (p, ws) => ops.fillSeries(ws, p.col),
  decimals: (p, ws) => ops.adjustDecimals(ws, p.col, Number(p.delta) >= 0 ? 1 : -1),
  balance: (p, ws) => ops.addBalanceColumn(ws, p.amount_col, p.paid_col, p.running !== false),
  sort: (p, ws) => ops.sortTable(ws, p.col, !!p.desc),
  filter: (_p, ws) => ops.addFilterDropdowns(ws),
  set_cell: (p, ws) => ops.setCellValue(ws, p.cell, p.value === undefined ? "" : p.value),
  delete_rows: (p, ws) => ops.deleteRows(ws, p.row_start, p.row_end),
  style_range: (p, ws) => ops.styleRange(ws, p.range, {
    bold: p.bold === undefined ? undefined : !!p.bold,
    italic: p.italic === undefined ? undefined : !!p.italic,
    colorHex: p.color || null,
    fillColorHex: p.fill || null,
  }),
  align_range: (p, ws) => ops.alignRange(ws, p.range, p.mode || "center"),
  number_format_range: (p, ws) => ops.numberFormatRange(ws, p.range, p.fmt, p.symbol || "₹"),
  clear_range: (p, ws) => ops.clearRange(ws, p.range),
  fill_formula: (p, ws) => ops.fillFormula(ws, p.range, p.formula),
};

function fail(res, err, fallbackMessage) {
  const status = err && err.status ? err.status : 500;
  if (status >= 500) console.error(err);
  res.status(status).json({
    ok: false,
    error: (err && err.message) || fallbackMessage || "Something went wrong.",
  });
}

function createApp(store, { token = null } = {}) {
  const app = express();
  app.disable("x-powered-by");

  // The UI is loaded from file:// in Electron, so every response must allow
  // that origin.
  app.use((req, res, next) => {
    res.setHeader("Access-Control-Allow-Origin", "*");
    res.setHeader("Access-Control-Allow-Headers", "Content-Type, X-Filename, X-Source-Path, X-ExcelHelper-Token");
    res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
    if (req.method === "OPTIONS") {
      res.sendStatus(204);
      return;
    }
    next();
  });

  // When Electron provides a one-time random token, every /api call must
  // carry it (header, or ?token= for EventSource/download links). This stops
  // any other local process or website from poking the API.
  if (token) {
    app.use("/api", (req, res, next) => {
      const provided = req.headers["x-excelhelper-token"] || req.query.token;
      if (provided === token) return next();
      res.status(403).json({ ok: false, error: "Access denied." });
    });
  }

  app.use(express.json({ limit: "1mb" }));

  // Serve the UI too, so the app can also be used from a normal browser
  // during development (`node server.js`).
  app.get("/", (_req, res) => res.sendFile(path.join(__dirname, "index.html")));
  app.use("/static", express.static(path.join(__dirname, "static")));

  app.get("/api/state", async (_req, res) => {
    try {
      res.json(await store.buildState());
    } catch (err) {
      fail(res, err);
    }
  });

  app.post("/api/upload",
    express.raw({ type: () => true, limit: "64mb" }),
    async (req, res) => {
      try {
        let name = "workbook.xlsx";
        try {
          name = decodeURIComponent(req.headers["x-filename"] || name);
        } catch (err) {
          const e = new Error("The file name could not be read - please try again.");
          e.status = 400;
          throw e;
        }
        if (!req.body || !req.body.length) {
          const e = new Error("No file received - please try again.");
          e.status = 400;
          throw e;
        }
        if (!name.toLowerCase().endsWith(".xlsx")) {
          const e = new Error("Excel Helper works with .xlsx files. Please save your file as .xlsx first (File > Save As in Excel).");
          e.status = 400;
          throw e;
        }
        // Remember where the file lives on disk so the preview can refresh
        // live when it changes there (optional, set by the desktop app).
        let sourcePath = null;
        try {
          sourcePath = req.headers["x-source-path"]
            ? path.resolve(decodeURIComponent(req.headers["x-source-path"]))
            : null;
        } catch (err) {
          sourcePath = null;
        }
        await store.importBuffer(req.body, name, sourcePath);
        res.json({
          ok: true,
          message: `Opened '${name}'. Pick a module on the left and click a button!`,
          state: await store.buildState(),
        });
      } catch (err) {
        fail(res, err);
      }
    });

  app.post("/api/practice", async (req, res) => {
    try {
      const kind = (req.body && req.body.type) === "ledger" ? "ledger" : "invoice";
      await store.loadPractice(kind);
      res.json({
        ok: true,
        message: PRACTICE_HINTS[kind],
        state: await store.buildState(),
      });
    } catch (err) {
      fail(res, err);
    }
  });

  app.post("/api/operation", async (req, res) => {
    const { op, params } = req.body || {};
    const handler = OPERATIONS[op];
    if (!handler) {
      const e = new Error("Unknown operation.");
      e.status = 400;
      return fail(res, e);
    }
    try {
      const message = await store.withWorkbook((wb, ws) => handler(params || {}, ws));
      const state = await store.buildState();
      store.broadcast({ state });
      res.json({ ok: true, message, state });
    } catch (err) {
      if (err instanceof FormulaEvalError) {
        err.status = 400;
        err.message = `Could not calculate a formula: ${err.message}`;
      } else if (!(err instanceof ops.OpError) && !(err.status === 400)) {
        // Unexpected failure - the file on disk is untouched (the snapshot
        // was taken before the change), so the user can simply retry or undo.
        err.status = 500;
        err.message = "Something went wrong applying that operation. Your file is unchanged - press Undo if needed.";
      }
      fail(res, err);
    }
  });

  app.post("/api/undo", async (_req, res) => {
    try {
      const done = await store.undo();
      if (!done) {
        const e = new Error("Nothing to undo.");
        e.status = 400;
        throw e;
      }
      res.json({
        ok: true,
        message: "Undone - went back one step.",
        state: await store.buildState(),
      });
    } catch (err) {
      fail(res, err);
    }
  });

  app.post("/api/sheet", async (req, res) => {
    try {
      const name = req.body && req.body.name;
      const message = await store.selectSheet(name);
      res.json({ ok: true, message, state: await store.buildState() });
    } catch (err) {
      fail(res, err);
    }
  });

  app.get("/api/download", async (_req, res) => {
    try {
      if (!store.hasFile()) {
        const e = new Error("Nothing to download yet.");
        e.status = 400;
        throw e;
      }
      const base = path.parse(store.filename || "workbook").name;
      const safe = base.replace(/[\\/:*?"<>|]/g, "").trim() || "workbook";
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Content-Disposition", `attachment; filename="${encodeURIComponent(`${safe} - Excel Helper.xlsx`)}"`);
      res.send(await fs.promises.readFile(store.currentFile));
    } catch (err) {
      fail(res, err);
    }
  });

  app.get("/api/events", async (req, res) => {
    res.set({
      "Content-Type": "text/event-stream",
      "Cache-Control": "no-cache",
      Connection: "keep-alive",
    });
    res.flushHeaders();
    const send = (payload) => {
      try {
        res.write(`data: ${JSON.stringify(payload)}\n\n`);
      } catch (err) {
        /* connection gone */
      }
    };
    try {
      send({ state: await store.buildState() });
    } catch (err) {
      /* ignore initial push failure */
    }
    const unsubscribe = store.subscribe(send);
    const heartbeat = setInterval(() => {
      try {
        res.write(": ping\n\n");
      } catch (err) {
        /* connection gone */
      }
    }, 25000);
    req.on("close", () => {
      clearInterval(heartbeat);
      unsubscribe();
    });
  });

  return app;
}

module.exports = { createApp };

/* Standalone mode: `node server.js` - handy for testing without Electron. */
if (require.main === module) {
  const { WorkbookStore } = require("./excel/workbook_store");
  const port = Number(process.env.PORT) || 5199;
  const store = new WorkbookStore(__dirname);
  store
    .init()
    .then(() => {
      createApp(store).listen(port, "127.0.0.1", () => {
        console.log(`Excel Helper server (standalone) on http://127.0.0.1:${port}`);
      });
    })
    .catch((err) => {
      console.error(err);
      process.exit(1);
    });
}
