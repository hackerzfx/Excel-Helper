/**
 * All Excel operations for Excel Helper - one function per UI button.
 *
 * Every function takes an exceljs worksheet plus simple parameters, changes
 * it, and returns a human-friendly message describing what happened. Any
 * user mistake (bad column, bad range, no data) throws OpError with a
 * plain-language explanation the UI can show directly.
 */

"use strict";

const {
  FormulaEvalError,
  effectiveCellValue,
  letterToCol,
  colToLetter,
} = require("./formula_eval");

class OpError extends Error {}

const BORDER_COLOR = "FF9CA3AF";
const THIN = { style: "thin", color: { argb: BORDER_COLOR } };
const BORDER_ALL = { top: THIN, left: THIN, bottom: THIN, right: THIN };

const NUMBER_FORMATS = {
  comma: "#,##0.00",
  percent: "0.00%",
  date: "DD-MM-YYYY",
};

const STAT_LABELS = {
  sum: { label: "Total", fn: "SUM" },
  average: { label: "Average", fn: "AVERAGE" },
  max: { label: "Maximum", fn: "MAX" },
  min: { label: "Minimum", fn: "MIN" },
  count: { label: "Count", fn: "COUNT" },
};

/* ------------------------------------------------------------------ */
/* Small helpers                                                       */
/* ------------------------------------------------------------------ */

function checkSheet(ws) {
  const empty =
    ws.rowCount < 1 ||
    (ws.rowCount === 1 && ws.columnCount === 1 &&
      (ws.getCell(1, 1).value === null || ws.getCell(1, 1).value === ""));
  if (empty) {
    throw new OpError("This sheet is empty - there is nothing to work on yet.");
  }
}

function resolveCol(ws, letter) {
  let col;
  try {
    col = letterToCol(String(letter || "").trim());
  } catch (err) {
    throw new OpError(`'${letter}' is not a valid column letter like A, B, C.`);
  }
  if (col < 1 || col > Math.max(ws.columnCount, 1)) {
    throw new OpError(
      `Column ${String(letter).toUpperCase()} is outside the table ` +
      `(the last used column is ${colToLetter(Math.max(ws.columnCount, 1))}).`
    );
  }
  return col;
}

function headerName(ws, col) {
  const v = ws.getCell(1, col).value;
  if (v === null || v === undefined || v === "") return colToLetter(col);
  if (typeof v === "object") {
    if (Array.isArray(v.richText)) return v.richText.map((p) => p.text).join("");
    if (v.formula !== undefined) return colToLetter(col);
    if (v.text !== undefined) return String(v.text);
  }
  return String(v);
}

/** Copy the style of header cell A1 onto another header cell (best effort). */
function copyHeaderStyle(ws, targetCell) {
  const src = ws.getCell(1, 1);
  const clone = (style) => (style ? JSON.parse(JSON.stringify(style)) : undefined);
  if (src.font) targetCell.font = clone(src.font);
  if (src.fill && src.fill.type === "pattern") targetCell.fill = clone(src.fill);
  if (src.alignment) targetCell.alignment = clone(src.alignment);
  if (src.border) targetCell.border = clone(src.border);
}

function safeHex(colorHex) {
  const hex = String(colorHex || "").trim().replace(/^#/, "");
  if (!/^[0-9A-Fa-f]{6}$/.test(hex)) {
    throw new OpError("Please pick a normal color (like 217346).");
  }
  return `FF${hex.toUpperCase()}`;
}

function displayText(value) {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) {
    const dd = String(value.getDate()).padStart(2, "0");
    const mm = String(value.getMonth() + 1).padStart(2, "0");
    return `${dd}-${mm}-${value.getFullYear()}`;
  }
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "number") return Number.isInteger(value) ? String(value) : String(value);
  if (typeof value === "object") {
    if (Array.isArray(value.richText)) return value.richText.map((p) => p.text).join("");
    if (value.formula !== undefined) return `=${value.formula}`;
    if (value.sharedFormula !== undefined) return value.result !== undefined && value.result !== null ? String(value.result) : "";
    if (value.text !== undefined) return String(value.text);
    return "";
  }
  return String(value);
}

function isEmptyValue(v) {
  return v === null || v === undefined || v === "";
}

function cloneCellValue(v) {
  if (v === null || typeof v !== "object") return v;
  if (v instanceof Date) return new Date(v.getTime());
  return JSON.parse(JSON.stringify(v));
}

function isFormulaValue(v) {
  return v !== null && typeof v === "object" && (v.formula !== undefined || v.sharedFormula !== undefined);
}

/** Parse "A1:E20" or "A1" into {top, left, bottom, right} (1-based numbers). */
function parseRange(rangeStr) {
  const text = String(rangeStr || "").trim().toUpperCase();
  const match = /^([A-Z]{1,3})(\d{1,7})(?::([A-Z]{1,3})(\d{1,7}))?$/.exec(text);
  if (!match) {
    throw new OpError(`'${rangeStr}' is not a valid range. Use a form like A1:E20.`);
  }
  const c1 = letterToCol(match[1]);
  const r1 = parseInt(match[2], 10);
  const c2 = match[3] ? letterToCol(match[3]) : c1;
  const r2 = match[3] ? parseInt(match[4], 10) : r1;
  return { top: Math.min(r1, r2), bottom: Math.max(r1, r2), left: Math.min(c1, c2), right: Math.max(c1, c2) };
}

/* ------------------------------------------------------------------ */
/* 1. Basics - headers, borders, formatting                            */
/* ------------------------------------------------------------------ */

function makeHeaderRow(ws, colorHex = "217346") {
  checkSheet(ws);
  const argb = safeHex(colorHex);
  for (let col = 1; col <= ws.columnCount; col += 1) {
    const cell = ws.getCell(1, col);
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb } };
    cell.alignment = { horizontal: "center", vertical: "middle" };
  }
  ws.views = [{ state: "frozen", ySplit: 1 }];
  return `Row 1 is now a proper header: bold white text on a colored background, centered, and frozen ` +
    `so it stays visible while you scroll (${ws.columnCount} columns styled).`;
}

function addBorders(ws, rangeStr = null) {
  checkSheet(ws);
  const text = (rangeStr || "").trim().toUpperCase() ||
    `A1:${colToLetter(ws.columnCount)}${ws.rowCount}`;
  const { top, bottom, left, right } = parseRange(text);
  for (let row = top; row <= bottom; row += 1) {
    for (let col = left; col <= right; col += 1) {
      ws.getCell(row, col).border = BORDER_ALL;
    }
  }
  const count = (bottom - top + 1) * (right - left + 1);
  return `Thin borders added to ${text.toUpperCase()} (${count} cells).`;
}

function applyNumberFormat(ws, letter, fmt, symbol = "₹") {
  checkSheet(ws);
  const col = resolveCol(ws, letter);
  let numFmt;
  let label;
  if (fmt === "currency") {
    const sym = (symbol || "₹").trim() || "₹";
    numFmt = `"${sym}"#,##0.00`;
    label = `currency (${sym})`;
  } else if (NUMBER_FORMATS[fmt]) {
    numFmt = NUMBER_FORMATS[fmt];
    label = fmt === "comma" ? "1,234.56" : fmt;
  } else {
    throw new OpError("Unknown number format.");
  }
  for (let row = 2; row <= ws.rowCount; row += 1) {
    ws.getCell(row, col).numFmt = numFmt;
  }
  return `Column ${letter.toUpperCase()} (${headerName(ws, col)}) now shows ${label} formatting in rows 2 onwards.`;
}

function alignColumn(ws, letter, mode = "center") {
  checkSheet(ws);
  const col = resolveCol(ws, letter);
  if (!["left", "center", "right"].includes(mode)) {
    throw new OpError("Alignment must be left, center or right.");
  }
  for (let row = 1; row <= ws.rowCount; row += 1) {
    ws.getCell(row, col).alignment = { horizontal: mode, vertical: "middle" };
  }
  return `Column ${letter.toUpperCase()} (${headerName(ws, col)}) is now aligned ${mode}.`;
}

function autofitColumns(ws) {
  checkSheet(ws);
  for (let col = 1; col <= ws.columnCount; col += 1) {
    let best = 8;
    for (let row = 1; row <= ws.rowCount; row += 1) {
      const value = ws.getCell(row, col).value;
      if (isFormulaValue(value)) continue; // formula text is not what shows
      best = Math.max(best, displayText(value).length);
    }
    ws.getColumn(col).width = Math.min(best + 3, 45);
  }
  return `Column widths adjusted to fit their content (${ws.columnCount} columns).`;
}

function styleColumn(ws, letter, bold = false, colorHex = null) {
  checkSheet(ws);
  const col = resolveCol(ws, letter);
  const font = {};
  if (bold) font.bold = true;
  if (colorHex) font.color = { argb: safeHex(colorHex) };
  if (!Object.keys(font).length) {
    throw new OpError("Choose bold, a color, or both.");
  }
  for (let row = 2; row <= ws.rowCount; row += 1) {
    ws.getCell(row, col).font = font;
  }
  const what = [];
  if (bold) what.push("bold");
  if (colorHex) what.push("colored");
  return `Column ${letter.toUpperCase()} (${headerName(ws, col)}) data cells are now ` +
    `${what.join(" and ")} (header row untouched).`;
}

/* ------------------------------------------------------------------ */
/* 2. Formulas                                                         */
/* ------------------------------------------------------------------ */

function multiplyColumns(ws, colA, colB, resultCol = null, headerNameText = "Amount") {
  checkSheet(ws);
  const a = resolveCol(ws, colA);
  const b = resolveCol(ws, colB);
  const aL = colToLetter(a);
  const bL = colToLetter(b);

  let target;
  const given = String(resultCol || "").trim().toUpperCase();
  if (given && given !== "AUTO" && given !== "__AUTO__") {
    try {
      target = letterToCol(given);
    } catch (err) {
      throw new OpError(`'${resultCol}' is not a valid result column.`);
    }
    if (target > ws.columnCount + 1) {
      throw new OpError(`Result column ${given} is too far right - the table ends at column ${colToLetter(ws.columnCount)}.`);
    }
  } else {
    target = ws.columnCount + 1;
  }
  const tL = colToLetter(target);

  // Never silently destroy data: refuse to overwrite an existing column that
  // already holds values (the header row alone is fine - it gets renamed).
  if (target !== a && target !== b) {
    for (let row = 2; row <= ws.rowCount; row += 1) {
      if (!isEmptyValue(ws.getCell(row, target).value)) {
        throw new OpError(
          `Column ${tL} (${headerName(ws, target)}) already has data. ` +
          `Pick the 'New column (automatic)' option or an empty column instead.`
        );
      }
    }
  }

  const header = (headerNameText || "Amount").trim() || "Amount";
  const headerCell = ws.getCell(1, target);
  headerCell.value = header;
  copyHeaderStyle(ws, headerCell);

  let count = 0;
  for (let row = 2; row <= ws.rowCount; row += 1) {
    const va = ws.getCell(row, a).value;
    const vb = ws.getCell(row, b).value;
    if (isEmptyValue(va) && isEmptyValue(vb)) continue;
    ws.getCell(row, target).value = { formula: `${aL}${row}*${bL}${row}` };
    count += 1;
  }
  if (count === 0) {
    throw new OpError("No data rows found to multiply.");
  }
  return `Added '${header}' in column ${tL}: ${count} formulas like =${aL}2*${bL}2. ` +
    `Each row multiplies ${aL} (${headerName(ws, a)}) by ${bL} (${headerName(ws, b)}).`;
}

/** True when a row holds a statistic written by addStat (label or formula).
 *  The label/formula may sit in any column, so scan the whole row. */
function isStatRow(ws, row) {
  for (let c = 1; c <= ws.columnCount; c += 1) {
    const v = ws.getCell(row, c).value;
    if (typeof v === "string" && Object.values(STAT_LABELS).some((m) => m.label === v)) {
      return true;
    }
    if (v && typeof v === "object" && typeof v.formula === "string" &&
        /^(SUM|AVERAGE|MAX|MIN|COUNT)\(/i.test(v.formula)) {
      return true;
    }
  }
  return false;
}

function addStat(ws, letter, stat = "sum") {
  checkSheet(ws);
  const col = resolveCol(ws, letter);
  const key = String(stat || "sum").toLowerCase();
  const meta = STAT_LABELS[key];
  if (!meta) {
    throw new OpError("Unknown statistic - choose sum, average, max, min or count.");
  }
  const L = colToLetter(col);
  // Do not include statistic rows from earlier clicks in the new formula.
  let dataEnd = ws.rowCount;
  while (dataEnd > 1 && isStatRow(ws, dataEnd)) dataEnd -= 1;
  if (dataEnd < 2) throw new OpError("No data rows found below the header.");
  const outRow = ws.rowCount + 1;

  const statCell = ws.getCell(outRow, col);
  statCell.value = { formula: `${meta.fn}(${L}2:${L}${dataEnd})` };
  const fmtAbove = ws.getCell(dataEnd, col).numFmt;
  if (fmtAbove) statCell.numFmt = fmtAbove;

  // Keep every statistic label in column A of its own line, so the totals
  // block reads like a proper summary. If the user already wrote something
  // there (e.g. their own "TOTAL" row), leave it untouched.
  let where = `in cell ${L}${outRow}`;
  if (col > 1 && isEmptyValue(ws.getCell(outRow, 1).value)) {
    const labelCell = ws.getCell(outRow, 1);
    labelCell.value = meta.label;
    labelCell.font = { bold: true };
    where += ` with the label '${meta.label}' in column A`;
  }
  return `${meta.label} of column ${L} written ${where} (live formula: =${meta.fn}(${L}2:${L}${dataEnd})).`;
}

function freezeFormulasAsValues(ws, letter = null) {
  checkSheet(ws);
  let converted = 0;
  let failed = 0;
  let scope;
  let cells;
  if (letter) {
    const col = resolveCol(ws, letter);
    scope = `column ${letter.toUpperCase()}`;
    cells = [];
    for (let row = 2; row <= ws.rowCount; row += 1) cells.push([row, col]);
  } else {
    scope = "the whole sheet";
    cells = [];
    for (let row = 1; row <= ws.rowCount; row += 1) {
      for (let col = 1; col <= ws.columnCount; col += 1) cells.push([row, col]);
    }
  }
  for (const [row, col] of cells) {
    const cell = ws.getCell(row, col);
    if (!isFormulaValue(cell.value)) continue;
    try {
      const result = effectiveCellValue(ws, row, col);
      if (typeof result === "number" || result instanceof Date) {
        cell.value = result;
        converted += 1;
      } else {
        failed += 1;
      }
    } catch (err) {
      failed += 1;
    }
  }
  if (converted === 0 && failed === 0) {
    return `No formulas found in ${scope} - nothing to convert.`;
  }
  let msg = `Converted ${converted} formula(s) in ${scope} into plain numbers that will never change.`;
  if (failed) {
    msg += ` ${failed} formula(s) were left alone because this tool can only calculate the ` +
      `basic functions it creates (SUM, AVERAGE, MAX, MIN, COUNT, + - * /).`;
  }
  return msg;
}

function fillSeries(ws, letter) {
  checkSheet(ws);
  const col = resolveCol(ws, letter);
  const L = colToLetter(col);
  let n = 0;
  for (let row = 2; row <= ws.rowCount; row += 1) {
    let hasContent = false;
    for (let c = 1; c <= ws.columnCount; c += 1) {
      if (!isEmptyValue(ws.getCell(row, c).value)) {
        hasContent = true;
        break;
      }
    }
    if (!hasContent) continue;
    ws.getCell(row, col).value = n + 1;
    n += 1;
  }
  if (n === 0) throw new OpError("No data rows found to number.");
  return `Filled serial numbers 1-${n} into column ${L} (${n} rows).`;
}

function shiftDecimals(fmt, delta) {
  const text = fmt || "General";
  if (text.toLowerCase() === "general") return delta > 0 ? "0.0" : "0";
  const parts = text.split(";");
  let p = parts[0];
  const match = /\.(0+)/.exec(p);
  if (delta > 0) {
    if (match) p = p.replace("." + match[1], "." + match[1] + "0");
    else if (p.endsWith("%")) p = p.slice(0, -1) + ".0%";
    else p += ".0";
  } else if (match) {
    if (match[1].length > 1) p = p.replace("." + match[1], "." + match[1].slice(0, -1));
    else p = p.replace("." + match[1], "");
  }
  parts[0] = p;
  return parts.join(";");
}

function adjustDecimals(ws, letter, delta) {
  checkSheet(ws);
  const col = resolveCol(ws, letter);
  const L = colToLetter(col);
  const current = ws.getCell(ws.rowCount, col).numFmt || "General";
  // Date/time formats must not get a decimal section appended.
  const looksLikeDate = /[dmyhs]/i.test(current) && !/\./.test(current);
  if (looksLikeDate) {
    return `Column ${L} (${headerName(ws, col)}) looks like it holds dates - decimal places were not changed.`;
  }
  const newFmt = shiftDecimals(current, delta > 0 ? 1 : -1);
  for (let row = 2; row <= ws.rowCount; row += 1) {
    ws.getCell(row, col).numFmt = newFmt;
  }
  const word = delta > 0 ? "Increased" : "Decreased";
  return `${word} decimal places in column ${L} (${headerName(ws, col)}). Format is now ${newFmt}.`;
}

/* ------------------------------------------------------------------ */
/* 3. Accounts - balance column                                        */
/* ------------------------------------------------------------------ */

function addBalanceColumn(ws, amountLetter, paidLetter, running = true) {
  checkSheet(ws);
  const amountCol = resolveCol(ws, amountLetter);
  const paidCol = resolveCol(ws, paidLetter);
  const aL = colToLetter(amountCol);
  const pL = colToLetter(paidCol);

  // If the sheet already ends with a Balance-like header and no data below
  // it (common in templates), fill that column instead of appending a
  // duplicate one.
  let target = ws.columnCount + 1;
  if (ws.columnCount > 1) {
    const lastCol = ws.columnCount;
    const header = ws.getCell(1, lastCol).value;
    const headerText = typeof header === "string" ? header.toLowerCase() : "";
    const dataEmpty = (() => {
      for (let row = 2; row <= ws.rowCount; row += 1) {
        if (!isEmptyValue(ws.getCell(row, lastCol).value)) return false;
      }
      return true;
    })();
    if (dataEmpty && (headerText.includes("balance") || headerText === "")) {
      target = lastCol;
    }
  }
  const tL = colToLetter(target);

  const headerCell = ws.getCell(1, target);
  if (isEmptyValue(headerCell.value)) {
    headerCell.value = "Balance";
    copyHeaderStyle(ws, headerCell);
  }

  let count = 0;
  let lastBalanceRow = null; // blank rows in between must not reset the carry
  for (let row = 2; row <= ws.rowCount; row += 1) {
    const amount = ws.getCell(row, amountCol).value;
    const paid = ws.getCell(row, paidCol).value;
    if (isEmptyValue(amount) && isEmptyValue(paid)) continue;
    const formula = running && lastBalanceRow
      ? `${tL}${lastBalanceRow}+${aL}${row}-${pL}${row}`
      : `${aL}${row}-${pL}${row}`;
    ws.getCell(row, target).value = { formula };
    lastBalanceRow = row;
    count += 1;
  }
  if (count === 0) {
    throw new OpError("No data rows found - nothing to calculate a balance for.");
  }
  const fmt = ws.getCell(2, amountCol).numFmt;
  if (fmt) {
    for (let row = 2; row <= ws.rowCount; row += 1) {
      ws.getCell(row, target).numFmt = fmt;
    }
  }
  const kind = running
    ? "running balance (each row adds to the one above)"
    : "simple balance (amount - paid for each row)";
  return `Added a 'Balance' column (${tL}) with ${count} live formulas using ${kind}.`;
}

/* ------------------------------------------------------------------ */
/* 4. Sort & filter                                                    */
/* ------------------------------------------------------------------ */

function sortTable(ws, letter, desc = false) {
  checkSheet(ws);
  const col = resolveCol(ws, letter);
  const lastCol = ws.columnCount;
  const lastRow = ws.rowCount;
  if (lastRow < 2) throw new OpError("There are no data rows to sort.");

  let converted = 0;
  let unconvertible = 0;
  const rows = [];
  for (let r = 2; r <= lastRow; r += 1) {
    const cells = [];
    for (let c = 1; c <= lastCol; c += 1) {
      const cell = ws.getCell(r, c);
      let v = cell.value;
      if (isFormulaValue(v)) {
        try {
          v = effectiveCellValue(ws, r, c);
          converted += 1;
        } catch (err) {
          unconvertible += 1;
        }
      }
      // Carry each cell's formatting with its row, so bold, fills and number
      // formats stay attached to the right record after sorting.
      cells.push({
        value: v,
        style: cell.style ? JSON.parse(JSON.stringify(cell.style)) : null,
      });
    }
    rows.push(cells);
  }
  // Formulas this tool cannot calculate (VLOOKUP, IF, ...) would keep their
  // old row references after the rows move, silently showing wrong numbers.
  if (unconvertible > 0) {
    throw new OpError(
      `This table has ${unconvertible} formula(s) with functions this tool cannot calculate ` +
      `(like VLOOKUP or IF). Sorting would move them to the wrong rows. ` +
      `Tip: use 'Freeze formulas as values' first, then sort.`
    );
  }

  const sortKey = (cells) => {
    const v = cells[col - 1].value;
    if (typeof v === "number" && Number.isFinite(v)) return [0, v, ""];
    if (typeof v === "boolean") return [1, 0, v ? "TRUE" : "FALSE"];
    if (v === null || v === undefined || v === "") return [2, 0, ""];
    if (v instanceof Date) return [0, v.getTime(), ""];
    return [1, 0, String(v).toLowerCase()];
  };

  // Excel semantics: rows with a value in the sort column get sorted; rows
  // that are blank in that column (e.g. a Total row) always stay last,
  // regardless of direction, in their original order.
  const withKey = [];
  const withoutKey = [];
  for (const cells of rows) {
    (isEmptyValue(cells[col - 1].value) ? withoutKey : withKey).push(cells);
  }
  withKey.sort((a, b) => {
    const ka = sortKey(a);
    const kb = sortKey(b);
    for (let i = 0; i < ka.length; i += 1) {
      if (ka[i] < kb[i]) return -1;
      if (ka[i] > kb[i]) return 1;
    }
    return 0;
  });
  if (desc) withKey.reverse();
  const ordered = withKey.concat(withoutKey);

  for (let i = 0; i < ordered.length; i += 1) {
    const newRow = i + 2;
    for (let c = 1; c <= lastCol; c += 1) {
      const cell = ws.getCell(newRow, c);
      const source = ordered[i][c - 1];
      cell.value = cloneCellValue(source.value);
      cell.style = source.style || {};
    }
  }
  const direction = desc ? "Z → A (largest first)" : "A → Z (smallest first)";
  const note = converted
    ? ` ${converted} formula(s) were turned into plain numbers so they stay correct after sorting.`
    : "";
  return `Table sorted by column ${letter.toUpperCase()} (${headerName(ws, col)}), ${direction}.${note}`;
}

function addFilterDropdowns(ws) {
  checkSheet(ws);
  ws.autoFilter = `A1:${colToLetter(ws.columnCount)}${ws.rowCount}`;
  return "Filter dropdowns added to the table. Open the saved file in Excel and you will see " +
    "small arrow buttons on each header - click one to show only the rows you want.";
}

/**
 * Manually set one cell (double-click editing in the preview).
 * Text stays text, plain numbers become numbers, "=..." becomes a real
 * formula, an empty string clears the cell.
 */
function setCellValue(ws, cellRef, rawValue) {
  const match = /^([A-Z]{1,3})([0-9]{1,7})$/i.exec(String(cellRef || "").trim());
  if (!match) {
    throw new OpError(`'${cellRef}' is not a valid cell address like B3.`);
  }
  const col = letterToCol(match[1].toUpperCase());
  const row = parseInt(match[2], 10);
  if (row < 1) throw new OpError("Row numbers start at 1.");
  const address = `${match[1].toUpperCase()}${row}`;
  const cell = ws.getCell(row, col);

  if (typeof rawValue !== "string") {
    cell.value = rawValue;
    return `Cell ${address} updated.`;
  }
  const text = rawValue.trim();
  if (text === "") {
    cell.value = null;
    return `Cell ${address} cleared.`;
  }
  if (text.startsWith("=")) {
    cell.value = { formula: text.slice(1) };
    return `Cell ${address} set to the formula ${text}.`;
  }
  if (/^-?\d+(\.\d+)?$/.test(text.replace(/,/g, ""))) {
    cell.value = Number(text.replace(/,/g, ""));
    return `Cell ${address} set to ${cell.value}.`;
  }
  cell.value = text;
  return `Cell ${address} set to '${text}'.`;
}

/** Apply simple styling (bold/italic/text color/fill) to every cell of a range. */
function styleRange(ws, rangeStr, opts = {}) {
  const { top, bottom, left, right } = parseRange(rangeStr);
  for (let row = top; row <= bottom; row += 1) {
    for (let col = left; col <= right; col += 1) {
      const cell = ws.getCell(row, col);
      if (opts.bold !== undefined || opts.italic !== undefined || opts.colorHex) {
        const font = Object.assign({}, cell.font);
        if (opts.bold !== undefined) font.bold = opts.bold;
        if (opts.italic !== undefined) font.italic = opts.italic;
        if (opts.colorHex) font.color = { argb: safeHex(opts.colorHex) };
        cell.font = font;
      }
      if (opts.fillColorHex) {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: safeHex(opts.fillColorHex) } };
      }
    }
  }
  const n = (bottom - top + 1) * (right - left + 1);
  return `Styling applied to ${n} cell(s) (${rangeStr.toUpperCase()}).`;
}

function alignRange(ws, rangeStr, mode = "center") {
  if (!["left", "center", "right"].includes(mode)) {
    throw new OpError("Alignment must be left, center or right.");
  }
  const { top, bottom, left, right } = parseRange(rangeStr);
  for (let row = top; row <= bottom; row += 1) {
    for (let col = left; col <= right; col += 1) {
      ws.getCell(row, col).alignment = { horizontal: mode, vertical: "middle" };
    }
  }
  return `Alignment '${mode}' applied to ${rangeStr.toUpperCase()}.`;
}

function numberFormatRange(ws, rangeStr, fmt, symbol = "₹") {
  const { top, bottom, left, right } = parseRange(rangeStr);
  let numFmt;
  let label;
  if (fmt === "currency") {
    const sym = (symbol || "₹").trim() || "₹";
    numFmt = `"${sym}"#,##0.00`;
    label = `currency (${sym})`;
  } else if (NUMBER_FORMATS[fmt]) {
    numFmt = NUMBER_FORMATS[fmt];
    label = fmt === "comma" ? "1,234.56" : fmt;
  } else {
    throw new OpError("Unknown number format.");
  }
  for (let row = top; row <= bottom; row += 1) {
    for (let col = left; col <= right; col += 1) {
      ws.getCell(row, col).numFmt = numFmt;
    }
  }
  return `${label} formatting applied to ${rangeStr.toUpperCase()}.`;
}

function clearRange(ws, rangeStr) {
  const { top, bottom, left, right } = parseRange(rangeStr);
  for (let row = top; row <= bottom; row += 1) {
    for (let col = left; col <= right; col += 1) {
      ws.getCell(row, col).value = null;
    }
  }
  const n = (bottom - top + 1) * (right - left + 1);
  return `Cleared ${n} cell(s) (${rangeStr.toUpperCase()}).`;
}

/** Shift a formula's row/column references by an offset (Excel fill behaviour).
 *  $-anchored parts stay fixed. */
function translateFormula(body, dr, dc) {
  return body.replace(/(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})/g, (m, colDollar, letters, rowDollar, digits) => {
    let col = letterToCol(letters);
    let row = parseInt(digits, 10);
    if (colDollar !== "$") col += dc;
    if (rowDollar !== "$") row += dr;
    if (row < 1 || col < 1) return m; // would fall off the sheet - leave as-is
    return `${colDollar}${colToLetter(col)}${rowDollar}${row}`;
  });
}

/** Write a formula into every cell of a range, shifting references relative
 *  to the top-left cell - the same thing Excel's fill handle does. */
function fillFormula(ws, rangeStr, rawFormula) {
  const { top, bottom, left, right } = parseRange(rangeStr);
  let body = String(rawFormula || "").trim();
  if (!body.startsWith("=")) {
    throw new OpError("Please start your formula with = (for example =B2*C2).");
  }
  body = body.slice(1);
  if (!body) throw new OpError("The formula is empty.");
  for (let row = top; row <= bottom; row += 1) {
    for (let col = left; col <= right; col += 1) {
      ws.getCell(row, col).value = { formula: translateFormula(body, row - top, col - left) };
    }
  }
  const n = (bottom - top + 1) * (right - left + 1);
  return `Formula written into ${n} cell(s) (${rangeStr.toUpperCase()}) with references shifted like Excel fill.`;
}

function deleteRows(ws, rowStart, rowEnd) {
  checkSheet(ws);
  const start = Math.max(2, parseInt(rowStart, 10));
  const end = Math.min(ws.rowCount, parseInt(rowEnd, 10));
  if (end < start) throw new OpError("No rows to clear.");
  let count = 0;
  for (let row = start; row <= end; row += 1) {
    let hadData = false;
    for (let col = 1; col <= ws.columnCount; col += 1) {
      const cell = ws.getCell(row, col);
      if (cell.value !== null && cell.value !== undefined && cell.value !== "") {
        hadData = true;
        cell.value = null;
      }
    }
    if (hadData) count += 1;
  }
  if (count === 0) throw new OpError("Selected rows are already empty.");
  return `Cleared ${count} row(s) (${start}-${end}). The rows are empty now - you can type new data.`;
}

module.exports = {
  OpError,
  makeHeaderRow,
  addBorders,
  applyNumberFormat,
  alignColumn,
  autofitColumns,
  styleColumn,
  multiplyColumns,
  addStat,
  freezeFormulasAsValues,
  fillSeries,
  adjustDecimals,
  addBalanceColumn,
  sortTable,
  addFilterDropdowns,
  setCellValue,
  deleteRows,
  styleRange,
  alignRange,
  numberFormatRange,
  clearRange,
  fillFormula,
  displayText,
};
