/**
 * Tiny Excel formula evaluator for Excel Helper - JavaScript port.
 *
 * Excel Helper writes simple, widely-compatible formulas (SUM, AVERAGE, MAX,
 * MIN, COUNT, B2*C2 ...). Excel itself calculates them the moment the file is
 * opened, but exceljs cannot - newly written formulas carry no cached result.
 * This module evaluates those formulas in JavaScript so operations like
 * "Freeze formulas as values" and sorting can know the real numbers.
 *
 * Supported: cell references (A1), ranges (A1:B9), + - * / ^, parentheses, and
 * the functions SUM, AVERAGE, MAX, MIN, COUNT, COUNTA, ROUND, ABS.
 * Anything else throws FormulaEvalError and the caller leaves the cell alone.
 */

"use strict";

const MAX_DEPTH = 30;

const FUNCTIONS = new Set(["SUM", "AVERAGE", "MAX", "MIN", "COUNT", "COUNTA", "ROUND", "ABS"]);

// Marker object returned by the parser for an A1:B9 style range (only valid
// as a function argument, e.g. SUM(A1:A9) - never "3 + (A1:A9)").
const RANGE = "__range__";

class FormulaEvalError extends Error {}

/* ------------------------------------------------------------------ */
/* Cell value access through exceljs value objects                     */
/* ------------------------------------------------------------------ */

/** Read the plain value of a cell, resolving formula objects recursively. */
function rawCellValue(ws, row, col, depth = 0) {
  if (depth >= MAX_DEPTH) {
    throw new FormulaEvalError("formula too deep (possible circular reference)");
  }
  let value;
  try {
    value = ws.getCell(row, col).value;
  } catch (err) {
    throw new FormulaEvalError(`bad cell reference: ${err.message}`);
  }
  return resolveValue(ws, value, depth);
}

/** Turn any exceljs cell value (including formula objects) into a plain value. */
function resolveValue(ws, value, depth) {
  if (value === null || value === undefined) return null;
  if (typeof value !== "object") return value;
  if (value instanceof Date) return value;
  if (value.formula !== undefined) {
    // Excel-cached result (file saved by Excel) is trustworthy; otherwise
    // evaluate the formula ourselves.
    if (value.result !== undefined && value.result !== null) return value.result;
    return evaluateFormula(ws, String(value.formula), depth + 1);
  }
  if (value.sharedFormula !== undefined) {
    if (value.result !== undefined && value.result !== null) return value.result;
    throw new FormulaEvalError("shared formula without a cached result");
  }
  if (Array.isArray(value.richText)) {
    return value.richText.map((part) => part.text).join("");
  }
  if (value.error !== undefined) {
    throw new FormulaEvalError(`cell error ${value.error}`);
  }
  if (value.hyperlink !== undefined) {
    return value.text !== undefined ? value.text : String(value.hyperlink);
  }
  if (value.toString && typeof value.toString === "function" && value.constructor === Object) {
    return String(value);
  }
  throw new FormulaEvalError("unsupported cell value");
}

/** Compute a cell's effective value for freeze/sort: formulas become numbers. */
function effectiveCellValue(ws, row, col) {
  const cell = ws.getCell(row, col);
  const value = cell.value;
  if (value === null || value === undefined) return null;
  if (typeof value !== "object") return value;
  if (value instanceof Date) return value;
  if (value.formula !== undefined) {
    if (value.result !== undefined && value.result !== null) return value.result;
    return evaluateFormula(ws, String(value.formula), 1);
  }
  if (value.sharedFormula !== undefined) {
    if (value.result !== undefined && value.result !== null) return value.result;
    throw new FormulaEvalError("shared formula without a cached result");
  }
  if (Array.isArray(value.richText)) {
    return value.richText.map((part) => part.text).join("");
  }
  if (value.error !== undefined) {
    throw new FormulaEvalError(`cell error ${value.error}`);
  }
  if (value.hyperlink !== undefined) {
    return value.text !== undefined ? value.text : String(value.hyperlink);
  }
  return value;
}

/* ------------------------------------------------------------------ */
/* Tokenizer + recursive descent parser                                */
/* ------------------------------------------------------------------ */

const TOKEN_RE = /\s*(?:(\d+(?:\.\d+)?)|(\$?[A-Za-z]{1,3}\$?\d{1,7})|([A-Za-z_][A-Za-z0-9_.]*)|([-+*/^(),:]))/y;

function tokenize(text) {
  const tokens = [];
  let pos = 0;
  while (pos < text.length) {
    TOKEN_RE.lastIndex = pos;
    const match = TOKEN_RE.exec(text);
    if (!match || match.index !== pos) {
      if (text.slice(pos).trim() === "") break;
      throw new FormulaEvalError(`cannot understand '${text.slice(pos).trim().slice(0, 20)}'`);
    }
    pos = TOKEN_RE.lastIndex;
    if (match[1] !== undefined) tokens.push({ kind: "num", text: match[1] });
    else if (match[2] !== undefined) tokens.push({ kind: "cell", text: match[2] });
    else if (match[3] !== undefined) tokens.push({ kind: "name", text: match[3] });
    else tokens.push({ kind: "op", text: match[4] });
  }
  return tokens;
}

function parseRef(ref) {
  const clean = ref.replace(/\$/g, "").toUpperCase();
  const match = /^([A-Z]{1,3})(\d{1,7})$/.exec(clean);
  if (!match) throw new FormulaEvalError(`bad cell reference '${ref}'`);
  return { row: parseInt(match[2], 10), col: letterToCol(match[1]) };
}

class Evaluator {
  constructor(ws, depth) {
    this.ws = ws;
    this.depth = depth;
    this.tokens = [];
    this.pos = 0;
  }

  evaluate(formula) {
    let body = String(formula).trim();
    if (body.startsWith("=")) body = body.slice(1);
    this.tokens = tokenize(body);
    this.pos = 0;
    if (this.tokens.length === 0) throw new FormulaEvalError("empty formula");
    const value = this.expr();
    if (this.pos !== this.tokens.length) {
      throw new FormulaEvalError("unexpected extra parts in formula");
    }
    return value;
  }

  peek() {
    return this.pos < this.tokens.length ? this.tokens[this.pos] : null;
  }

  expect(text) {
    const token = this.peek();
    if (!token || token.kind !== "op" || token.text !== text) {
      throw new FormulaEvalError(`expected '${text}'`);
    }
    this.pos += 1;
  }

  expr() {
    let value = this.term();
    for (;;) {
      const token = this.peek();
      if (token && token.kind === "op" && (token.text === "+" || token.text === "-")) {
        this.pos += 1;
        const right = this.term();
        value = token.text === "+" ? value + right : value - right;
      } else {
        return value;
      }
    }
  }

  term() {
    let value = this.factor();
    for (;;) {
      const token = this.peek();
      if (token && token.kind === "op" && (token.text === "*" || token.text === "/")) {
        this.pos += 1;
        const right = this.factor();
        if (token.text === "*") {
          value = value * right;
        } else {
          if (right === 0) throw new FormulaEvalError("division by zero");
          value = value / right;
        }
      } else {
        return value;
      }
    }
  }

  factor() {
    const token = this.peek();
    if (token && token.kind === "op" && (token.text === "+" || token.text === "-")) {
      this.pos += 1;
      const value = this.factor();
      return token.text === "+" ? value : -value;
    }
    let value = this.primary();
    const next = this.peek();
    if (next && next.kind === "op" && next.text === "^") {
      this.pos += 1;
      return value ** this.factor();
    }
    return value;
  }

  primary() {
    const token = this.next();
    if (token.kind === "num") return parseFloat(token.text);
    if (token.kind === "cell") {
      const { row, col } = parseRef(token.text);
      const next = this.peek();
      if (next && next.kind === "op" && next.text === ":") {
        this.pos += 1;
        const second = this.next();
        if (second.kind !== "cell") throw new FormulaEvalError("bad range");
        const end = parseRef(second.text);
        return {
          [RANGE]: true,
          row1: Math.min(row, end.row),
          col1: Math.min(col, end.col),
          row2: Math.max(row, end.row),
          col2: Math.max(col, end.col),
        };
      }
      return this.cellValue(row, col);
    }
    if (token.kind === "name") {
      const name = token.text.toUpperCase();
      if (name === "TRUE") return 1;
      if (name === "FALSE") return 0;
      if (FUNCTIONS.has(name)) {
        this.expect("(");
        const args = this.args();
        this.expect(")");
        return this.call(name, args);
      }
      throw new FormulaEvalError(`unknown name '${token.text}'`);
    }
    if (token.kind === "op" && token.text === "(") {
      const value = this.expr();
      this.expect(")");
      return value;
    }
    throw new FormulaEvalError("unexpected symbol in formula");
  }

  next() {
    if (this.pos >= this.tokens.length) throw new FormulaEvalError("formula ends unexpectedly");
    return this.tokens[this.pos++];
  }

  args() {
    const args = [this.arg()];
    for (;;) {
      const token = this.peek();
      if (token && token.kind === "op" && token.text === ",") {
        this.pos += 1;
        args.push(this.arg());
      } else {
        return args;
      }
    }
  }

  arg() {
    // A bare range only makes sense as a whole argument; primary() returns it
    // as a marker object which expr() passes through untouched.
    return this.expr();
  }

  isRange(arg) {
    return arg !== null && typeof arg === "object" && arg[RANGE] === true;
  }

  call(name, args) {
    if (name === "ROUND" || name === "ABS") {
      if (args.some((a) => this.isRange(a))) {
        throw new FormulaEvalError(`${name} does not accept a range`);
      }
    }
    const numbers = [];
    let counta = 0;
    for (const arg of args) {
      if (this.isRange(arg)) {
        for (let row = arg.row1; row <= arg.row2; row += 1) {
          for (let col = arg.col1; col <= arg.col2; col += 1) {
            const value = this.cellValue(row, col);
            if (value === null || (typeof value === "string" && value.trim() === "")) continue;
            counta += 1;
            if (typeof value === "number" && Number.isFinite(value)) numbers.push(value);
          }
        }
      } else if (typeof arg === "number" && Number.isFinite(arg)) {
        numbers.push(arg);
        counta += 1;
      }
    }
    switch (name) {
      case "SUM":
        return numbers.reduce((a, b) => a + b, 0);
      case "AVERAGE":
        return numbers.length ? sum(numbers) / numbers.length : 0;
      case "MAX":
        return numbers.length ? Math.max(...numbers) : 0;
      case "MIN":
        return numbers.length ? Math.min(...numbers) : 0;
      case "COUNT":
        return numbers.length;
      case "COUNTA":
        return counta;
      case "ROUND": {
        if (!args.length || typeof args[0] !== "number" || !Number.isFinite(args[0])) {
          throw new FormulaEvalError("ROUND needs a number");
        }
        let digits = 0;
        if (args.length > 1) {
          if (typeof args[1] !== "number" || !Number.isFinite(args[1])) {
            throw new FormulaEvalError("ROUND digits must be a number");
          }
          digits = Math.trunc(args[1]);
        }
        const factor = 10 ** digits;
        return Math.round(args[0] * factor) / factor;
      }
      case "ABS":
        if (!args.length || typeof args[0] !== "number" || !Number.isFinite(args[0])) {
          throw new FormulaEvalError("ABS needs a number");
        }
        return Math.abs(args[0]);
      default:
        throw new FormulaEvalError(`unsupported function ${name}`);
    }
  }

  cellValue(row, col) {
    const value = rawCellValue(this.ws, row, col, this.depth + 1);
    if (typeof value === "boolean") return value ? 1 : 0;
    return value;
  }
}

function sum(numbers) {
  return numbers.reduce((a, b) => a + b, 0);
}

/* ------------------------------------------------------------------ */
/* Column letter helpers (exceljs does not export these)               */
/* ------------------------------------------------------------------ */

function letterToCol(letters) {
  let n = 0;
  for (const ch of String(letters).toUpperCase()) {
    const code = ch.charCodeAt(0);
    if (code < 65 || code > 90) throw new FormulaEvalError(`bad column '${letters}'`);
    n = n * 26 + (code - 64);
  }
  return n;
}

function colToLetter(col) {
  let n = col;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

/** Evaluate an Excel formula string against an exceljs worksheet; returns a number. */
function evaluateFormula(ws, formula, depth = 0) {
  let value;
  try {
    value = new Evaluator(ws, depth).evaluate(formula);
  } catch (err) {
    if (err instanceof FormulaEvalError) throw err;
    if (err instanceof TypeError) {
      throw new FormulaEvalError(`text where a number was expected: ${err.message}`);
    }
    throw new FormulaEvalError(err.message);
  }
  if (value !== null && typeof value === "object" && value[RANGE] === true) {
    throw new FormulaEvalError("a range cannot be used on its own");
  }
  if (typeof value === "boolean") return value ? 1 : 0;
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw new FormulaEvalError("formula did not produce a number");
  }
  return Number.isInteger(value) ? value : value;
}

module.exports = {
  FormulaEvalError,
  evaluateFormula,
  effectiveCellValue,
  rawCellValue,
  resolveValue,
  letterToCol,
  colToLetter,
};
