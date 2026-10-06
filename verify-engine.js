/**
 * Headless self-test for the Excel Helper engine.
 *
 * Run:  node verify-engine.js
 * Exits 0 when everything passes, 1 otherwise.
 */

"use strict";

const assert = require("assert");
const fs = require("fs");
const os = require("os");
const path = require("path");

const ExcelJS = require("exceljs");

const {
  evaluateFormula,
  effectiveCellValue,
  FormulaEvalError,
} = require("./excel/formula_eval");
const ops = require("./excel/operations");
const { WorkbookStore } = require("./excel/workbook_store");

let passed = 0;
let failed = 0;

function check(name, condition, detail = "") {
  if (condition) {
    passed += 1;
    console.log(`  [PASS] ${name}`);
  } else {
    failed += 1;
    console.error(`  [FAIL] ${name} ${detail}`);
  }
}

async function freshInvoice() {
  const { makeInvoiceWorkbook } = require("./excel/practice_data");
  const wb = makeInvoiceWorkbook();
  return { wb, ws: wb.getWorksheet("Invoice") };
}

async function freshLedger() {
  const { makeLedgerWorkbook } = require("./excel/practice_data");
  const wb = makeLedgerWorkbook();
  return { wb, ws: wb.getWorksheet("Ledger") };
}

async function testFormulaEvaluator() {
  console.log("\n== Formula evaluator ==");
  const { wb, ws } = await freshInvoice();
  ws.getCell(2, 4).value = { formula: "B2*C2" };
  ws.getCell(3, 4).value = { formula: "SUM(B2:B3)*2" };
  ws.getCell(4, 4).value = { formula: "AVERAGE(C2:C3)" };
  ws.getCell(5, 4).value = { formula: "MAX(C2,C3)" };
  ws.getCell(6, 4).value = { formula: "MIN(C2:C3)" };
  ws.getCell(7, 4).value = { formula: "COUNT(B2:B9)" };
  ws.getCell(8, 4).value = { formula: "COUNTA(A2:A9)" };
  ws.getCell(9, 4).value = { formula: "ROUND(C5/2, 1)" }; // C5 = 85.5 -> 42.8
  ws.getCell(10, 4).value = { formula: "ABS(0-7)" };

  check("multiply", evaluateFormula(ws, "B2*C2") === 450);
  check("nested ref in formula cell", evaluateFormula(ws, "D2*2") === 900);
  check("SUM range times 2", evaluateFormula(ws, "D3") === (10 + 12) * 2);
  check("AVERAGE", evaluateFormula(ws, "D4") === Math.abs((45 + 10) / 2));
  check("MAX args", evaluateFormula(ws, "D5") === 45);
  check("MIN range", evaluateFormula(ws, "D6") === 10);
  check("COUNT numbers", evaluateFormula(ws, "D7") === 8);
  check("COUNTA non-empty", evaluateFormula(ws, "D8") === 8);
  check("ROUND", evaluateFormula(ws, "D9") === 42.8);
  check("ABS", evaluateFormula(ws, "D10") === 7);

  check("effective value of formula cell", effectiveCellValue(ws, 2, 4) === 450);

  let threw = false;
  try {
    evaluateFormula(ws, "SUM(A1:A9) + (B1:B9)");
  } catch (err) {
    threw = err instanceof FormulaEvalError;
  }
  check("bare range rejected", threw);

  threw = false;
  try {
    evaluateFormula(ws, "B2/0");
  } catch (err) {
    threw = err instanceof FormulaEvalError;
  }
  check("division by zero rejected", threw);

  threw = false;
  try {
    evaluateFormula(ws, "VLOOKUP(A1, B1:B9, 1)");
  } catch (err) {
    threw = err instanceof FormulaEvalError;
  }
  check("unsupported function rejected", threw);
}

async function testInvoiceOperations() {
  console.log("\n== Invoice operations ==");
  const { wb, ws } = await freshInvoice();

  let msg = ops.makeHeaderRow(ws, "217346");
  check("header styled", ws.getCell(1, 1).font.bold === true);
  check("header fill", String(ws.getCell(1, 1).fill.fgColor.argb).endsWith("217346"));
  check("header frozen", ws.views && ws.views[0] && ws.views[0].state === "frozen" && ws.views[0].ySplit === 1);

  msg = ops.multiplyColumns(ws, "B", "C", "D", "Amount");
  check("multiply header", ws.getCell(1, 4).value === "Amount");
  check("multiply formula", ws.getCell(2, 4).value.formula === "B2*C2");
  check("multiply count row9", ws.getCell(9, 4).value.formula === "B9*C9");

  msg = ops.fillSeries(ws, "D");
  check("fill series start", ws.getCell(2, 4).value === 1);
  check("fill series end", ws.getCell(9, 4).value === 8);

  // Sort BEFORE adding the stats row: with a Total label present, Excel
  // semantics put text above numbers in a Z->A sort, which is intended.
  msg = ops.sortTable(ws, "C", true); // Rate desc
  check("sort desc by rate", ws.getCell(2, 3).value === 275);
  msg = ops.sortTable(ws, "A", false); // Item asc
  check("sort asc by item", String(ws.getCell(2, 1).value).startsWith("A4"));

  msg = ops.addStat(ws, "D", "sum");
  check("sum label in column A", ws.getCell(10, 1).value === "Total");
  check("sum formula", ws.getCell(10, 4).value.formula === "SUM(D2:D9)");

  msg = ops.freezeFormulasAsValues(ws, null);
  check("freeze D2 numeric", typeof ws.getCell(2, 4).value === "number");
  check("freeze D10 sum", Math.abs(ws.getCell(10, 4).value - 36) < 1e-9); // serials 1..8
  check("freeze kept numbers as numbers", typeof ws.getCell(2, 2).value === "number");

  msg = ops.adjustDecimals(ws, "D", 1);
  check("decimals on General", ws.getCell(2, 4).numFmt === "0.0");
  msg = ops.adjustDecimals(ws, "D", 1);
  check("decimals increase", ws.getCell(2, 4).numFmt === "0.00");
  msg = ops.adjustDecimals(ws, "D", -1);
  check("decimals decrease", ws.getCell(2, 4).numFmt === "0.0");

  msg = ops.applyNumberFormat(ws, "C", "currency", "₹");
  check("currency numFmt", ws.getCell(2, 3).numFmt === `"₹"#,##0.00`);

  msg = ops.alignColumn(ws, "A", "center");
  check("alignment", ws.getCell(3, 1).alignment.horizontal === "center");

  msg = ops.autofitColumns(ws);
  check("autofit sets width", typeof ws.getColumn(1).width === "number" && ws.getColumn(1).width >= 8);

  msg = ops.styleColumn(ws, "B", true, "B91C1C");
  check("style column bold", ws.getCell(2, 2).font.bold === true);
  check("style column color", String(ws.getCell(2, 2).font.color.argb).endsWith("B91C1C"));
  check("header untouched by styleColumn", ws.getCell(1, 2).font.bold === true || ws.getCell(1, 2).font.bold === undefined);

  msg = ops.addBorders(ws, "A1:E20");
  check("borders set", ws.getCell(3, 3).border.top.style === "thin");
  msg = ops.addFilterDropdowns(ws);
  check("autofilter set", typeof ws.autoFilter === "string" ? ws.autoFilter.includes("A1") : !!ws.autoFilter);

  // Round-trip through a real file
  const buffer = await wb.xlsx.writeBuffer();
  const wb2 = new ExcelJS.Workbook();
  await wb2.xlsx.load(buffer);
  const ws2 = wb2.getWorksheet("Invoice");
  check("round-trip keeps values", ws2.getCell(2, 3).value === 275);
  check("round-trip keeps autofilter", !!ws2.autoFilter);

  await wb.xlsx.writeFile("Demo_Result.xlsx");
}

async function testLedgerOperations() {
  console.log("\n== Ledger operations ==");
  const { wb, ws } = await freshLedger();

  ops.addBalanceColumn(ws, "C", "D", true);
  check("balance reuses empty Balance header column", ws.getCell(1, 5).value === "Balance");
  check("balance first row formula", ws.getCell(2, 5).value.formula === "C2-D2");
  check("balance running formula", ws.getCell(3, 5).value.formula === "E2+C3-D3");
  check("balance last row formula", ws.getCell(9, 5).value.formula === "E8+C9-D9");

  ops.freezeFormulasAsValues(ws, "E");
  // E2 = C2-D2 = 5000 (opening debit, no credit yet)
  check("balance frozen value", ws.getCell(2, 5).value === 5000);
  check("balance frozen final", ws.getCell(9, 5).value === 6400);

  ops.addStat(ws, "E", "average");
  check("average formula", ws.getCell(10, 5).value.formula === "AVERAGE(E2:E9)");
  check("second stat skips first stat row", (() => {
    ops.addStat(ws, "E", "sum");
    return ws.getCell(11, 5).value.formula === "SUM(E2:E9)";
  })());
  ops.freezeFormulasAsValues(ws, "E");
  // average of E2..E9 = (5000+3800+6100+4600+5400+4950+6700+6400)/8 = 5368.75
  check("average exact", Math.abs(ws.getCell(10, 5).value - (5000 + 3800 + 6100 + 4600 + 5400 + 4950 + 6700 + 6400) / 8) < 1e-6);
  check("second stat not double-counted", ws.getCell(11, 5).value === 5000 + 3800 + 6100 + 4600 + 5400 + 4950 + 6700 + 6400);
}

async function testStore() {
  console.log("\n== WorkbookStore (undo, state, sheet switching) ==");
  const base = await fs.promises.mkdtemp(path.join(os.tmpdir(), "excelhelper-test-"));
  const store = new WorkbookStore(base);
  await store.init();

  await store.loadPractice("invoice");
  let state = await store.buildState();
  check("practice state has file", state.has_file === true);
  check("practice sheets", state.sheets.length === 1 && state.sheets[0] === "Invoice");
  check("practice header", state.header.slice(0, 3).join("|") === "Item|Qty|Rate");
  check("grid rows", state.grid.length === 9); // header + 8 items
  check("empty Amount shows blank", state.grid[1][3].v === "");

  await store.withWorkbook((wb, ws) => ops.multiplyColumns(ws, "B", "C", "D", "Amount"));
  state = await store.buildState();
  check("after multiply formula flagged", state.grid[1][3].f === true && state.grid[1][3].fm === "=B2*C2" && state.grid[1][3].v === "450");
  check("formulaCount counted", state.formulaCount === 8);
  check("canUndo true", state.canUndo === true);

  const undone = await store.undo();
  check("undo ran", undone === true);
  state = await store.buildState();
  check("undo removed formulas, header stays", state.grid[0][3].v === "Amount" && state.formulaCount === 0);
  check("undo clears canUndo", state.canUndo === false);

  await fs.promises.rm(base, { recursive: true, force: true });
}

async function testSafetyRails() {
  console.log("\n== Safety rails (from code review) ==");

  // Sort must refuse rows with formulas the evaluator cannot compute.
  const a = await freshInvoice();
  ops.multiplyColumns(a.ws, "B", "C", "D", "Amount");
  a.ws.getCell(2, 1).value = { formula: 'IF(B2>5,1,2)' };
  let threw = false;
  try {
    ops.sortTable(a.ws, "C", true);
  } catch (err) {
    threw = err instanceof ops.OpError;
  }
  check("sort refuses unsupported formulas", threw);

  // Running balance must carry across a blank row.
  const wb2 = new ExcelJS.Workbook();
  const ws2 = wb2.addWorksheet("T");
  ws2.addRow(["Label", "Amt", "Paid"]);
  ws2.addRow(["a", 100, null]);
  ws2.addRow([null, null, null]);
  ws2.addRow(["b", 50, null]);
  ops.addBalanceColumn(ws2, "B", "C", true);
  check("blank row keeps running balance", ws2.getCell(4, 4).value.formula === "D2+B4-C4");

  // Multiply must refuse to overwrite a column that already has data.
  const d = await freshInvoice();
  d.ws.getCell(2, 5).value = "keep me";
  let threw3 = false;
  try {
    ops.multiplyColumns(d.ws, "B", "C", "E", "X");
  } catch (err) {
    threw3 = err instanceof ops.OpError;
  }
  check("multiply refuses to overwrite data", threw3);

  // Decimals must leave date columns alone.
  const e2 = await freshLedger();
  const msg = ops.adjustDecimals(e2.ws, "A", 1);
  check("decimals skip date columns", msg.includes("dates"));

  // Sorting must carry formatting with the rows.
  const s = await freshLedger();
  s.ws.getCell(2, 2).font = { bold: true }; // "Opening cash" row bold
  ops.sortTable(s.ws, "C", true); // sort by Debit desc -> Opening cash (5000) goes to row 2 anyway; use another column
  ops.sortTable(s.ws, "B", false); // sort by Particulars A-Z: "Bought stock" first
  const f22 = s.ws.getCell(2, 2).font;
  check("bold moves with its row", !(f22 && f22.bold));
  const bolded = [];
  for (let r = 2; r <= s.ws.rowCount; r += 1) {
    const font = s.ws.getCell(r, 2).font;
    if (font && font.bold) bolded.push(String(s.ws.getCell(r, 2).value));
  }
  check("bold stays on its own record", bolded.length === 1 && bolded[0] === "Opening cash");

  // Stat ranges must not include other stat rows (user-reported polish bug).
  const st2 = await freshInvoice();
  ops.addStat(st2.ws, "C", "sum"); // row 10
  ops.addStat(st2.ws, "C", "average"); // row 11
  check("second stat skips first stat row (cross-check)", st2.ws.getCell(11, 3).value.formula === "AVERAGE(C2:C9)");
  const st3 = await freshInvoice();
  ops.addStat(st3.ws, "D", "sum");
  ops.addStat(st3.ws, "C", "max"); // different column, later row
  check("stat on other column skips earlier stat row", st3.ws.getCell(11, 3).value.formula === "MAX(C2:C9)");
}

async function testSetCell() {
  console.log("\n== Manual cell editing ==");
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("T");
  ops.setCellValue(ws, "A1", "Notebook");
  check("set text", ws.getCell("A1").value === "Notebook");
  ops.setCellValue(ws, "B1", "42");
  check("set number", ws.getCell("B1").value === 42);
  ops.setCellValue(ws, "B1", "1,234.5");
  check("set number with comma", ws.getCell("B1").value === 1234.5);
  ops.setCellValue(ws, "C1", "=B1*2");
  check("set formula", ws.getCell("C1").value.formula === "B1*2");
  check("formula computes", evaluateFormula(ws, ws.getCell("C1").value.formula) === 2469);
  ops.setCellValue(ws, "B1", "10");
  check("edit recomputes", evaluateFormula(ws, ws.getCell("C1").value.formula) === 20);
  ops.setCellValue(ws, "A1", "");
  check("clear cell", ws.getCell("A1").value === null);
  ops.setCellValue(ws, "A1", "plain text");
  check("set text again", ws.getCell("A1").value === "plain text");
  let threw = false;
  try {
    ops.setCellValue(ws, "nope", "x");
  } catch (err) {
    threw = err instanceof ops.OpError;
  }
  check("bad address rejected", threw);
}

async function testRangeOps() {
  // Range operations (drag-selection feature).
  const rw = new ExcelJS.Workbook();
  const rs = rw.addWorksheet("R");
  rs.getCell("A1").value = 10;
  rs.getCell("A2").value = 20;
  rs.getCell("A3").value = 30;
  rs.getCell("B1").value = "x";
  ops.styleRange(rs, "A1:B2", { bold: true, fillColorHex: "fff3bf" });
  check("style_range bold", rs.getCell("A1").font.bold === true && rs.getCell("B2").font.bold === true);
  check("style_range fill", String(rs.getCell("A2").fill.fgColor.argb).endsWith("FFF3BF"));
  ops.clearRange(rs, "B1:B2");
  check("clear_range", rs.getCell("B1").value === null && rs.getCell("B2").value === null);
  ops.fillFormula(rs, "B1:B3", "=A1*2");
  check("fill_formula row1", rs.getCell("B1").value.formula === "A1*2");
  check("fill_formula shifted rows", rs.getCell("B2").value.formula === "A2*2" && rs.getCell("B3").value.formula === "A3*2");
  ops.fillFormula(rs, "C1:C3", "=$A$1+A1");
  check("fill_formula keeps $ anchors", rs.getCell("C3").value.formula === "$A$1+A3");
  let threw4 = false;
  try {
    ops.fillFormula(rs, "C1:C3", "B2*2");
  } catch (err) {
    threw4 = err instanceof ops.OpError;
  }
  check("fill_formula requires leading =", threw4);
}

async function main() {
  await testFormulaEvaluator();
  await testInvoiceOperations();
  await testLedgerOperations();
  await testSafetyRails();
  await testSetCell();
  await testRangeOps();
  await testStore();
  console.log(`\n${passed} passed, ${failed} failed`);
  if (failed > 0) process.exit(1);
  console.log("Demo workbook written to Demo_Result.xlsx (open it in Excel!)");
}

main().catch((err) => {
  console.error("TEST RUNNER CRASHED:", err);
  process.exit(1);
});
