/**
 * Practice workbook generators for Excel Helper.
 *
 * The Amount (invoice) and Balance (ledger) columns are deliberately left
 * EMPTY: filling them in with the app's own buttons IS the practice. Hints
 * for what to try live in the server's practice messages.
 */

"use strict";

const ExcelJS = require("exceljs");

const INVOICE_ITEMS = [
  ["Notebook - 200 pages", 10, 45],
  ["Ball pen - blue", 12, 10],
  ["Ball pen - black", 8, 10],
  ["Pencil box", 5, 85.5],
  ["Geometry box", 4, 120],
  ["A4 paper ream", 2, 275],
  ["Stapler", 3, 99.99],
  ["Sticky notes", 6, 35],
];

const LEDGER_ROWS = [
  [new Date(2026, 8, 25), "Opening cash", 5000, null],
  [new Date(2026, 8, 26), "Bought stock", null, 1200],
  [new Date(2026, 8, 28), "Cash sales", 2300, null],
  [new Date(2026, 8, 30), "Paid shop rent", null, 1500],
  [new Date(2026, 9, 1), "Received from Ravi", 800, null],
  [new Date(2026, 9, 1), "Electricity bill", null, 450],
  [new Date(2026, 9, 2), "Cash sales", 1750, null],
  [new Date(2026, 9, 2), "Paid transport", null, 300],
];

function newWorkbook() {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Excel Helper";
  return wb;
}

function makeInvoiceWorkbook() {
  const wb = newWorkbook();
  const ws = wb.addWorksheet("Invoice");
  ws.addRow(["Item", "Qty", "Rate", "Amount"]);
  for (const [name, qty, rate] of INVOICE_ITEMS) {
    // Amount (column D) is left empty on purpose - the learner fills it.
    ws.addRow([name, qty, rate]);
  }
  ws.getColumn(1).width = 24;
  ws.getColumn(2).width = 8;
  ws.getColumn(3).width = 10;
  ws.getColumn(4).width = 12;
  return wb;
}

function makeLedgerWorkbook() {
  const wb = newWorkbook();
  const ws = wb.addWorksheet("Ledger");
  ws.addRow(["Date", "Particulars", "Debit (in)", "Credit (out)", "Balance"]);
  for (const [date, particulars, debit, credit] of LEDGER_ROWS) {
    // Balance (column E) is left empty on purpose - the learner fills it.
    ws.addRow([date, particulars, debit, credit]);
  }
  for (let row = 2; row <= ws.rowCount; row += 1) {
    ws.getCell(row, 1).numFmt = "DD-MM-YYYY";
  }
  ws.getColumn(1).width = 12;
  ws.getColumn(2).width = 22;
  ws.getColumn(3).width = 12;
  ws.getColumn(4).width = 13;
  ws.getColumn(5).width = 12;
  return wb;
}

module.exports = { makeInvoiceWorkbook, makeLedgerWorkbook };
