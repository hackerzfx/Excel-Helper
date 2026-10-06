/**
 * Excel Helper — frontend logic.
 *
 * Talks to the embedded API server (Electron main process) over fetch.
 * Every operation returns the refreshed state, so the preview re-renders
 * instantly after every click; disk changes arrive over Server-Sent Events.
 */

"use strict";

const $ = (id) => document.getElementById(id);

const API =
  window.ExcelHelperAPI && window.ExcelHelperAPI.port
    ? `http://127.0.0.1:${window.ExcelHelperAPI.port}`
    : "";

// One-time token the Electron main process handed to the page (null in
// browser-dev mode, where the standalone server runs without a token).
const TOKEN = (window.ExcelHelperAPI && window.ExcelHelperAPI.token) || null;

function apiUrl(pathname) {
  if (!TOKEN) return `${API}${pathname}`;
  const joiner = pathname.includes("?") ? "&" : "?";
  return `${API}${pathname}${joiner}token=${encodeURIComponent(TOKEN)}`;
}

let state = { has_file: false };
let busy = false;
let bannerTimer = null;
let quickFind = "";
let selection = null; // {r1,c1,r2,c2} sheet coordinates, 1-based
let dragging = false;
let lastActiveSheet = null;

/* ---------------- small helpers ---------------- */

function showBanner(type, text) {
  const banner = $("banner");
  const bannerText = $("banner-text");
  banner.classList.remove("hidden", "success", "error");
  banner.classList.add(type === "error" ? "error" : "success");
  banner.setAttribute("role", type === "error" ? "alert" : "status");
  bannerText.textContent = text;
  if (bannerTimer) clearTimeout(bannerTimer);
  bannerTimer = setTimeout(hideBanner, 8000);
}

function hideBanner() {
  $("banner").classList.add("hidden");
  if (bannerTimer) {
    clearTimeout(bannerTimer);
    bannerTimer = null;
  }
}

function setBusy(value) {
  busy = value;
  for (const button of document.querySelectorAll("button")) {
    button.disabled = value;
  }
  renderLocks();
}

/* ---------------- range selection ---------------- */

function normSelection() {
  if (!selection) return null;
  return {
    r1: Math.min(selection.r1, selection.r2),
    r2: Math.max(selection.r1, selection.r2),
    c1: Math.min(selection.c1, selection.c2),
    c2: Math.max(selection.c1, selection.c2),
  };
}

function selectionRangeString() {
  const s = normSelection();
  if (!s || !state.has_file) return null;
  const letter = (n) => (state.colLetters || [])[n - 1];
  if (!letter(s.c1) || !letter(s.c2)) return null;
  const a = `${letter(s.c1)}${s.r1}`;
  const b = `${letter(s.c2)}${s.r2}`;
  return a === b ? a : `${a}:${b}`;
}

function paintSelection() {
  const s = normSelection();
  for (const cellEl of document.querySelectorAll("#preview .sheet-cell")) {
    if (!s) {
      cellEl.classList.remove("selected");
      continue;
    }
    const r = parseInt(cellEl.dataset.r, 10);
    const c = parseInt(cellEl.dataset.c, 10);
    const inside = r >= s.r1 && r <= s.r2 && c >= s.c1 && c <= s.c2;
    cellEl.classList.toggle("selected", inside);
  }
}


/* ------------------------------------------------------------------ */
/* Smart suggestions engine                                            */
/* ------------------------------------------------------------------ */

/** Look at the sheet's columns and data and suggest the most useful
 * next steps — like a teacher pointing at what to try. Always returns
 * at least the basics; a live selection refines the list further. */
function analyzeSheet(state) {
  if (!state.has_file) return [];
  const header = state.header || [];
  const colLetters = state.colLetters || [];
  const maxRow = state.maxRow || 0;
  const maxCol = state.columnCount || state.maxCol || colLetters.length;
  const grid = state.grid || [];

  // Map header names to 1-based column indices.
  const norm = (h) => String(h || "").toLowerCase().trim();
  const headerIndex = (predicate) => {
    const i = header.findIndex((h) => predicate(norm(h)));
    return i >= 0 ? i + 1 : 0;
  };
  const letter = (col) => colLetters[col - 1] || "";

  const suggestions = [];
  const seen = new Set();
  const push = (sug) => {
    if (seen.has(sug.id)) return;
    seen.add(sug.id);
    suggestions.push(sug);
  };

  // Detect column roles from headers.
  const qtyCol = headerIndex((h) => /^(qty|quantity|units|pcs)$/i.test(h) || h === "qty");
  const rateCol = headerIndex((h) => /^(rate|price|unit price)/i.test(h));
  const amountCol = headerIndex((h) => /^(amount|total|value|debit)/i.test(h));
  const paidCol = headerIndex((h) => /^(paid|credit|received)/i.test(h));
  const balanceCol = headerIndex((h) => /^balance/i.test(h));
  const hasHeader = header.some((h) => h !== null && h !== undefined && h !== "");
  const hasEnoughRows = maxRow >= 3; // header + at least 2 data rows

  // --- Accounts: QTY x RATE → Amount ---
  if (qtyCol && rateCol && !amountCol && hasEnoughRows) {
    push({
      id: "add-amount",
      icon: "✖️", primary: true,
      label: "Add Amount (QTY × Rate)",
      hint: `New column =${letter(qtyCol)}row × ${letter(rateCol)}row`,
      op: "multiply",
      params: { col_a: letter(qtyCol), col_b: letter(rateCol) },
    });
  }

  // --- Accounts: Amount + Paid → Balance ---
  if (amountCol && paidCol && !balanceCol && hasEnoughRows) {
    push({
      id: "add-balance",
      icon: "⚖️", primary: true,
      label: "Add Balance (=Amount − Paid)",
      hint: `New column =${letter(amountCol)}row − ${letter(paidCol)}row`,
      op: "balance",
      params: { amount_col: letter(amountCol), paid_col: letter(paidCol), running: false },
    });
  }

  // --- Add SUM / stats to any numeric column ---
  if (amountCol && hasEnoughRows) {
    push({
      id: "add-total",
      icon: "∑", primary: true,
      label: `Add Total (SUM of ${header[amountCol - 1] || "Amount"})`,
      hint: `=SUM(${letter(amountCol)}2:${letter(amountCol)}${maxRow})`,
      op: "stat",
      params: { col: letter(amountCol), stat: "sum" },
    });
  } else if (qtyCol && hasEnoughRows) {
    push({
      id: "add-total-qty",
      icon: "∑",
      label: "Add Total Quantity (SUM)",
      hint: `=SUM(${letter(qtyCol)}2:${letter(qtyCol)}${maxRow})`,
      op: "stat",
      params: { col: letter(qtyCol), stat: "sum" },
    });
  }

  // --- Header row styling (only suggest once) ---
  if (hasHeader && !state.headerStyled) {
    push({
      id: "style-header",
      icon: "🖌️",
      label: "Style the header row",
      hint: "Make row 1 bold, colored & frozen",
      op: "make_header",
      params: { color: "217346" },
    });
  }

  // --- Borders ---
  if (hasEnoughRows) {
    push({
      id: "add-borders",
      icon: "🔲",
      label: "Add borders to the table",
      hint: "Draw thin borders around every cell",
      op: "add_borders",
      params: { range: "" },
    });
  }

  // --- Sort & filter ---
  if (hasHeader && hasEnoughRows) {
    push({
      id: "add-filter",
      icon: "🔽",
      label: "Add filter dropdowns",
      hint: "Filter arrows on each header in Excel",
      op: "filter",
      params: {},
    });
  }

  // --- Format currency on money columns ---
  if ((rateCol || amountCol) && hasEnoughRows) {
    const target = amountCol || rateCol;
    push({
      id: "fmt-money",
      icon: "₹",
      label: `Format "${header[target - 1] || "column"}" as money`,
      hint: "Add ₹ symbol & thousands commas",
      op: "number_format",
      params: { col: letter(target), fmt: "currency", symbol: "₹" },
    });
  }

  // --- Fill serial numbers if first column looks empty/names ---
  if (hasEnoughRows && qtyCol !== 1 && rateCol !== 1 && amountCol !== 1) {
    const firstColData = [];
    for (let r = 2; r <= Math.min(maxRow, 6); r += 1) {
      const v = ((grid[r - 1] || [])[0] || {}).v;
      firstColData.push(v);
    }
    const firstColLooksLikeSerial = firstColData.every((v) => v === null || v === undefined || v === "");
    if (firstColLooksLikeSerial) {
      push({
        id: "fill-serial",
        icon: "🔢",
        label: "Add serial numbers (S.No)",
        hint: "Fill 1, 2, 3… down column A",
        op: "fill_series",
        params: { col: letter(1) },
      });
    }
  }

  return suggestions;
}

/** Refine suggestions when a specific range is selected. */
function analyzeSelection(state, sel) {
  if (!sel || !state.has_file) return analyzeSheet(state);
  const s = normSelection();
  if (!s) return analyzeSheet(state);

  const header = state.header || [];
  const colLetters = state.colLetters || [];
  const maxRow = state.maxRow || 0;
  const grid = state.grid || [];
  const letter = (c) => colLetters[c - 1] || "";

  const selectedColCount = s.c2 - s.c1 + 1;
  const selectedRowCount = s.r2 - s.r1 + 1;
  const firstColLetter = letter(s.c1);
  const isHeaderOnly = s.r1 === 1 && s.r2 === 1;

  // Gather selected cells' raw values.
  let numericCount = 0;
  let textCount = 0;
  let formulaCount = 0;
  let emptyCount = 0;
  for (let r = s.r1; r <= s.r2; r += 1) {
    for (let c = s.c1; c <= s.c2; c += 1) {
      const cell = (grid[r - 1] || [])[c - 1] || {};
      const v = cell.v;
      if (v === null || v === undefined || v === "") emptyCount += 1;
      else if (cell.f) formulaCount += 1;
      else if (typeof v === "number") numericCount += 1;
      else textCount += 1;
    }
  }

  const allNumeric = numericCount > 0 && textCount === 0 && formulaCount === 0;
  const allText = textCount > 0 && numericCount === 0;
  const anyFormula = formulaCount > 0;
  const headerName = normSelectionHeaderName(state, s);
  const suggestions = [];
  const seen = new Set();
  const push = (sug) => {
    if (seen.has(sug.id)) return;
    seen.add(sug.id);
    suggestions.push(sug);
  };

  // Two columns selected → multiply?
  if (selectedColCount === 2 && !isHeaderOnly && selectedRowCount >= 2) {
    push({
      id: "multiply-selected",
      icon: "✖️", primary: true,
      label: "Multiply these two columns",
      hint: `Add a new column =col1 × col2`,
      op: "multiply",
      params: { col_a: firstColLetter, col_b: letter(s.c2) },
    });
  }

  // Single numeric column → stats
  if (selectedColCount === 1 && allNumeric && selectedRowCount >= 2) {
    push({
      id: "stat-sum",
      icon: "∑", primary: true,
      label: "Add SUM",
      hint: `Total of the selected numbers`,
      op: "stat",
      params: { col: firstColLetter, stat: "sum" },
    });
    push({
      id: "stat-avg",
      icon: "Ø",
      label: "Add AVERAGE",
      hint: "Mean of the selected numbers",
      op: "stat",
      params: { col: firstColLetter, stat: "average" },
    });
    push({
      id: "stat-max",
      icon: "⬆",
      label: "Highest (MAX)",
      hint: "Largest value",
      op: "stat",
      params: { col: firstColLetter, stat: "max" },
    });
    push({
      id: "stat-min",
      icon: "⬇",
      label: "Lowest (MIN)",
      hint: "Smallest value",
      op: "stat",
      params: { col: firstColLetter, stat: "min" },
    });
  }

  // Numeric → format as money
  if (numericCount > 0 && !isHeaderOnly) {
    push({
      id: "fmt-money-sel",
      icon: "₹",
      label: "Format as money",
      hint: "Add ₹ & commas to selected cells",
      op: "number_format_range",
      params: { range: selectionRangeString(), fmt: "currency", symbol: "₹" },
    });
  }

  // Text column → sort
  if (selectedColCount === 1 && allText && selectedRowCount >= 2) {
    push({
      id: "sort-asc",
      icon: "A→Z", primary: true,
      label: `Sort "${headerName}" A→Z`,
      hint: "Sort the table by this column",
      op: "sort",
      params: { col: firstColLetter, desc: false },
    });
    push({
      id: "sort-desc",
      icon: "Z→A",
      label: `Sort "${headerName}" Z→A`,
      hint: "Sort the table in reverse",
      op: "sort",
      params: { col: firstColLetter, desc: true },
    });
  }

  // Any data → borders
  if (!isHeaderOnly && selectedRowCount >= 1) {
    push({
      id: "borders-sel",
      icon: "🔲",
      label: "Add borders",
      hint: "Border around selected cells",
      op: "add_borders",
      params: { range: selectionRangeString() },
    });
  }

  // Formulas → freeze
  if (anyFormula) {
    push({
      id: "freeze-sel",
      icon: "❄️",
      label: "Freeze formulas as values",
      hint: "Replace formulas with plain numbers",
      op: "freeze_values",
      params: { col: "" },
    });
  }

  // If nothing specific matched, fall back to sheet-level suggestions.
  if (suggestions.length === 0) return analyzeSheet(state).slice(0, 4);

  return suggestions.slice(0, 6);
}

function normSelectionHeaderName(state, s) {
  const header = state.header || [];
  const name = header[s.c1 - 1];
  return name ? String(name) : "";
}

function updateSelectionStatus() {
  const el = $("sel-status");
  if (!el) return;
  if (!state.has_file) {
    el.textContent = "";
    renderSuggestions(null);
    return;
  }
  const rs = selectionRangeString();
  if (!rs) {
    el.textContent = "Tip: drag across cells to select a range — the Smart Suggestions will refine to match your selection.";
    renderSuggestions(analyzeSheet(state));
    return;
  }
  const s = normSelection();
  const n = (s.r2 - s.r1 + 1) * (s.c2 - s.c1 + 1);
  el.textContent = `Selected: ${rs} (${n} cell${n > 1 ? "s" : ""})`;
  renderSuggestions(analyzeSelection(state, selection));
}

function renderSuggestions(suggestions) {
  const container = $("smart-suggestions");
  if (!container) return;
  container.innerHTML = "";
  if (!suggestions || suggestions.length === 0) {
    container.hidden = true;
    return;
  }
  container.hidden = false;

  const head = document.createElement("div");
  head.className = "suggestions-head";
  head.innerHTML =
    `<span class="suggestions-title">💡 Suggested next steps</span>` +
    `<span class="suggestions-hint">Click a suggestion to apply it</span>`;
  container.appendChild(head);

  const row = document.createElement("div");
  row.className = "suggestions-row";
  for (const suggestion of suggestions) {
    const chip = document.createElement("button");
    chip.className = "suggestion-chip" + (suggestion.primary ? " primary" : "");
    chip.innerHTML =
      `<span class="sugg-icon">${suggestion.icon}</span>` +
      `<span class="sugg-text"><span class="sugg-label">${suggestion.label}</span>` +
      `<span class="sugg-hint">${suggestion.hint}</span></span>`;
    chip.addEventListener("click", () => runOp(suggestion.op, suggestion.params));
    row.appendChild(chip);
  }
  container.appendChild(row);
}

/* ------------------------------------------------------------------ */
/* Right-click context menu on the preview table                       */
/* ------------------------------------------------------------------ */

/** Build context menu items for where the user right-clicked. */
function buildContextMenu(target, clickType) {
  if (!state.has_file) return [];
  const items = [];
  const add = (icon, label, action, danger = false) =>
    items.push({ icon, label, action, danger });
  const addSep = () => { if (items.length && !items[items.length - 1].sep) items.push({ sep: true }); };

  const grid = state.grid || [];
  const header = state.header || [];
  const colLetters = state.colLetters || [];
  const letter = (c) => colLetters[c - 1] || "";

  if (clickType === "column-header") {
    const col = parseInt(target.dataset.c || "0", 10);
    const colLetter = letter(col);
    const headerName = header[col - 1] || colLetter;
    add("A→Z", `Sort "${headerName}" A→Z`, { op: "sort", params: { col: colLetter, desc: false } });
    add("Z→A", `Sort "${headerName}" Z→A`, { op: "sort", params: { col: colLetter, desc: true } });
    addSep();
    add("∑", "Add SUM below", { op: "stat", params: { col: colLetter, stat: "sum" } });
    add("Ø", "Add AVERAGE", { op: "stat", params: { col: colLetter, stat: "average" } });
    add("⬆", "Add MAX", { op: "stat", params: { col: colLetter, stat: "max" } });
    add("⬇", "Add MIN", { op: "stat", params: { col: colLetter, stat: "min" } });
    add("#", "Add COUNT", { op: "stat", params: { col: colLetter, stat: "count" } });
    addSep();
    if (typeof headerName === "string" && /amount|rate|price|total|debit|value/i.test(headerName)) {
      add("₹", "Format as money", { op: "number_format", params: { col: colLetter, fmt: "currency", symbol: "₹" } });
    } else {
      add("🔢", "Format as numbers", { op: "number_format", params: { col: colLetter, fmt: "comma" } });
    }
    add("🖌️", "Bold + color column", { op: "style_column", params: { col: colLetter, bold: true, color: "217346" } });
    addSep();
    add("🔢", "Fill serial numbers", { op: "fill_series", params: { col: colLetter } });
    add("❄️", "Freeze formulas as values", { op: "freeze_values", params: { col: colLetter } });
    return items;
  }

  if (clickType === "row-number") {
    const row = parseInt(target.dataset.r || "0", 10);
    add("🗑️", `Clear row ${row}`, { op: "delete_rows", params: { row_start: row, row_end: row } }, true);
    return items;
  }

  // data cell — use current selection or the single clicked cell
  const sel = selection || { r1: 0, c1: 0, r2: 0, c2: 0 };
  const r1 = sel.r1 || parseInt(target.dataset.r || "0", 10);
  const c1 = sel.c1 || parseInt(target.dataset.c || "0", 10);
  const r2 = sel.r2 || r1;
  const c2 = sel.c2 || c1;

  if (r1 >= 2) {
    add("🗑️", `Clear row${r1 === r2 ? " " + r1 : "s " + r1 + "-" + r2}`, { op: "delete_rows", params: { row_start: r1, row_end: r2 } }, true);
  }
  addSep();
  if ((c2 - c1 + 1) === 2) {
    add("✖️", "Multiply columns", { op: "multiply", params: { col_a: letter(c1), col_b: letter(c2) } });
  }
  add("∑", "Add SUM", { op: "stat", params: { col: letter(c1), stat: "sum" } });
  add("Ø", "Add AVERAGE", { op: "stat", params: { col: letter(c1), stat: "average" } });
  addSep();
  add("₹", "Format as money", { op: "number_format_range", params: { range: rangeString(r1, c1, r2, c2), fmt: "currency", symbol: "₹" } });
  add("🔲", "Add borders", { op: "add_borders", params: { range: rangeString(r1, c1, r2, c2) } });
  addSep();
  add("❄️", "Freeze formulas as values", { op: "freeze_values", params: { col: "" } });
  add("✏️", "Edit cell", { op: "__edit__", params: { r: r1, c: c1 } });
  return items;
}

function rangeString(r1, c1, r2, c2) {
  const colLetters = state.colLetters || [];
  const a = (colLetters[c1 - 1] || "") + r1;
  const b = (colLetters[c2 - 1] || "") + r2;
  return a === b ? a : `${a}:${b}`;
}

function showContextMenu(event, target, clickType) {
  const menu = $("context-menu");
  if (!menu) return;
  hideContextMenu();
  const items = buildContextMenu(target, clickType);
  if (!items.length) return;

  menu.innerHTML = "";
  for (const item of items) {
    if (item.sep) {
      const sep = document.createElement("div");
      sep.className = "ctx-sep";
      menu.appendChild(sep);
      continue;
    }
    const button = document.createElement("button");
    button.className = "ctx-item" + (item.danger ? " danger" : "");
    button.innerHTML = `<span class="ctx-icon">${item.icon}</span>${item.label}`;
    button.addEventListener("click", () => {
      hideContextMenu();
      if (item.action.op === "__edit__") {
        const cell = document.querySelector(`#preview td[data-r="${item.action.params.r}"][data-c="${item.action.params.c}"]`);
        if (cell) startCellEdit(cell, `${(state.colLetters || [])[item.action.params.c - 1] || ""}${item.action.params.r}`, (state.grid[item.action.params.r - 1] || [])[item.action.params.c - 1] || {});
      } else {
        runOp(item.action.op, item.action.params);
      }
    });
    menu.appendChild(button);
  }
  menu.hidden = false;
  menu.style.left = "0px";
  menu.style.top = "0px";
  const rect = menu.getBoundingClientRect();
  let x = event.clientX;
  let y = event.clientY;
  if (x + rect.width > window.innerWidth) x = window.innerWidth - rect.width - 8;
  if (y + rect.height > window.innerHeight) y = window.innerHeight - rect.height - 8;
  menu.style.left = `${x}px`;
  menu.style.top = `${y}px`;
}

function hideContextMenu() {
  const menu = $("context-menu");
  if (menu) menu.hidden = true;
}

function postJson(url, body) {
  return fetchJson(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      ...(TOKEN ? { "X-ExcelHelper-Token": TOKEN } : {}),
    },
    body: JSON.stringify(body || {}),
  });
}

function fetchJson(url, options) {
  return (async () => {
    let response;
    try {
      response = await fetch(url, options);
    } catch (err) {
      showBanner("error", "Can't reach the Excel Helper server. Try restarting the app.");
      return null;
    }
    try {
      return await response.json();
    } catch (err) {
      showBanner("error", `Server error (${response.status}).`);
      return null;
    }
  })();
}

async function applyResponse(data) {
  if (!data) return;
  if (data.ok) {
    if (data.state) state = data.state;
    if (data.message) showBanner("success", data.message);
    render();
    flashPreview();
  } else {
    showBanner("error", data.error || "Something went wrong.");
  }
}

function flashPreview() {
  const wrap = $("preview-wrap");
  if (!wrap || wrap.hidden) return;
  wrap.classList.remove("updating");
  void wrap.offsetWidth;
  wrap.classList.add("updating");
}

/* Every operation button flashes when clicked so it feels alive. */
document.addEventListener("click", (event) => {
  const button = event.target.closest(".btn");
  if (!button || button.disabled) return;
  button.classList.remove("flash");
  void button.offsetWidth;
  button.classList.add("flash");
});

async function runOp(op, params) {
  if (busy) return;
  setBusy(true);
  const data = await postJson(apiUrl("/api/operation"), { op, params });
  await applyResponse(data);
  setBusy(false);
}

/* ---------------- upload (native dialog + fallback) ---------------- */

async function uploadBytes(name, buffer, sourcePath) {
  if (busy) return;
  setBusy(true);
  try {
    const data = await fetchJson(apiUrl("/api/upload"), {
      method: "POST",
      headers: {
        "Content-Type": "application/octet-stream",
        "X-Filename": encodeURIComponent(name),
        ...(TOKEN ? { "X-ExcelHelper-Token": TOKEN } : {}),
        ...(sourcePath ? { "X-Source-Path": encodeURIComponent(sourcePath) } : {}),
      },
      body: buffer,
    });
    await applyResponse(data);
  } finally {
    setBusy(false);
  }
}

async function openExcelFile() {
  if (window.ExcelHelperAPI && window.ExcelHelperAPI.openFile) {
    let result;
    try {
      result = await Promise.race([
        window.ExcelHelperAPI.openFile(),
        new Promise((_, reject) => setTimeout(() => reject(new Error("IPC timeout (5s)")), 5000)),
      ]);
    } catch (err) {
      showBanner("error", "Could not open the file dialog: " + err.message);
      return;
    }
    if (!result || result.canceled) return;
    if (result.error) {
      showBanner("error", result.error);
      return;
    }
    await uploadBytes(result.name, result.data, result.path || null);
  }
  $("file-input").click();
}

/* ---------------- rendering ---------------- */

function populateColSelects() {
  const letters = state.colLetters || [];
  const header = state.header || [];
  for (const select of document.querySelectorAll("select.colsel")) {
    const previous = select.value;
    select.innerHTML = "";
    if (select.dataset.firstValue !== undefined) {
      const option = document.createElement("option");
      option.value = select.dataset.firstValue;
      option.textContent = select.dataset.firstLabel;
      select.appendChild(option);
    }
    letters.forEach((letter, index) => {
      const option = document.createElement("option");
      option.value = letter;
      const label = header[index] ? header[index] : "";
      option.textContent = label ? `${letter} — ${label}` : letter;
      select.appendChild(option);
    });
    if (previous && [...select.options].some((o) => o.value === previous)) {
      select.value = previous;
    }
  }
}

function renderSheets() {
  const strip = $("sheet-strip");
  strip.innerHTML = "";
  const sheets = state.sheets || [];
  if (!state.has_file || sheets.length <= 1) {
    strip.hidden = true;
    return;
  }
  strip.hidden = false;
  for (const name of sheets) {
    const pill = document.createElement("button");
    pill.className = "sheet-pill" + (name === state.active_sheet ? " active" : "");
    pill.textContent = name;
    pill.addEventListener("click", async () => {
      if (busy) return;
      setBusy(true);
      await applyResponse(await postJson(apiUrl("/api/sheet"), { name }));
      setBusy(false);
    });
    strip.appendChild(pill);
  }
}

function renderTable() {
  const table = $("preview");
  table.innerHTML = "";
  if (!state.has_file) return;

  const letters = state.colLetters || [];
  const header = state.header || [];
  const grid = state.grid || [];

  const thead = document.createElement("thead");

  const letterRow = document.createElement("tr");
  const corner = document.createElement("th");
  corner.className = "gutter";
  corner.textContent = "";
  letterRow.appendChild(corner);
  for (const letter of letters) {
    const th = document.createElement("th");
    th.textContent = letter;
    letterRow.appendChild(th);
  }
  thead.appendChild(letterRow);

  const headerRow = document.createElement("tr");
  headerRow.className = "head-row";
  const headGutter = document.createElement("th");
  headGutter.className = "gutter";
  headGutter.textContent = "1";
  headerRow.appendChild(headGutter);
  header.forEach((text, i) => {
    const th = document.createElement("th");
    th.textContent = text;
    th.dataset.r = "1";
    th.dataset.c = String(i + 1);
    th.classList.add("sheet-cell");
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  for (let rowIndex = 1; rowIndex < grid.length; rowIndex += 1) {
    const tr = document.createElement("tr");
    tr.dataset.rowIndex = String(rowIndex + 1);

    const gutter = document.createElement("th");
    gutter.className = "gutter";
    gutter.textContent = String(rowIndex + 1);
    tr.appendChild(gutter);

    for (let colIndex = 0; colIndex < grid[rowIndex].length; colIndex += 1) {
      const cell = grid[rowIndex][colIndex];
      const td = document.createElement("td");
      td.textContent = cell.v === undefined || cell.v === null ? "" : cell.v;
      td.dataset.r = String(rowIndex + 1);
      td.dataset.c = String(colIndex + 1);
      td.classList.add("sheet-cell");
      if (cell.f) {
        td.classList.add("formula");
        if (cell.fm) {
          td.title = cell.fm;
          td.classList.add("has-formula");
        }
      }
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
  table.appendChild(tbody);

  const note = $("trunc-note");
  if (state.truncatedRows || state.truncatedCols) {
    note.textContent =
      `Preview shows the first ${grid.length} rows × ${letters.length} columns — ` +
      "the full file is saved, don't worry.";
    note.classList.remove("hidden");
  } else {
    note.classList.add("hidden");
  }
}

function renderLocks() {
  const hasFile = !!state.has_file;
  $("sidebar").classList.toggle("disabled", !hasFile);
  $("btn-download").disabled = busy || !hasFile;
  $("btn-undo").disabled = busy || !state.has_file || !state.canUndo;
  $("empty-state").hidden = hasFile;
  $("preview-wrap").hidden = !hasFile;
}

function applyQuickFind() {
  const query = quickFind.trim().toLowerCase();
  for (const tr of document.querySelectorAll("#preview tbody tr")) {
    let match = false;
    if (query) {
      for (const td of tr.querySelectorAll("td")) {
        if (td.textContent.toLowerCase().includes(query)) {
          match = true;
          break;
        }
      }
    }
    tr.classList.toggle("match", match);
  }
}

/* ---------------- manual cell editing ---------------- */

function startCellEdit(td, address, cell) {
  if (td.querySelector("input")) return;
  const original = cell.fm || (cell.v === undefined || cell.v === null ? "" : String(cell.v));
  const input = document.createElement("input");
  input.className = "cell-editor";
  input.value = original;
  input.setAttribute("aria-label", `Edit cell ${address}`);
  td.textContent = "";
  td.appendChild(input);
  input.focus();
  input.select();
  let finished = false;
  const finish = (commit) => {
    if (finished) return;
    finished = true;
    const value = input.value;
    if (commit && value !== original) {
      runOp("set_cell", { cell: address, value });
    } else {
      render(); // put the original content back
    }
  };
  input.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      finish(true);
    } else if (event.key === "Escape") {
      event.preventDefault();
      finish(false);
    }
  });
  input.addEventListener("blur", () => finish(true));
}

function render() {
  if (state.active_sheet !== lastActiveSheet) {
    selection = null; // a new sheet has no selection
    lastActiveSheet = state.active_sheet;
  }
  renderLocks();
  renderSheets();
  renderTable();
  populateColSelects();
  applyQuickFind();
  paintSelection();
  // Show sheet-level suggestions whenever a file is open; selection refines them.
  const suggestions = state.has_file
    ? (selection ? analyzeSelection(state, selection) : analyzeSheet(state))
    : null;
  renderSuggestions(suggestions);
}

/* ---------------- wiring ---------------- */

function wireTabs() {
  for (const tab of document.querySelectorAll(".tab")) {
    tab.addEventListener("click", () => {
      for (const t of document.querySelectorAll(".tab")) {
        t.classList.toggle("active", t === tab);
        t.setAttribute("aria-selected", t === tab ? "true" : "false");
      }
      for (const page of document.querySelectorAll(".tabpage")) {
        const active = page.id === `tab-${tab.dataset.tab}`;
        page.classList.toggle("active", active);
        page.hidden = !active;
      }
    });
  }
}

function wireOperations() {
  /* ---- dark mode toggle ---- */
  const applyTheme = (theme) => {
    document.documentElement.dataset.theme = theme;
    localStorage.setItem("excel-helper-theme", theme);
    $("btn-theme").textContent = theme === "dark" ? "☀️" : "🌙";
  };
  const savedTheme = localStorage.getItem("excel-helper-theme");
  const prefersDark = window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches;
  applyTheme(savedTheme || (prefersDark ? "dark" : "light"));
  $("btn-theme").addEventListener("click", () => {
    applyTheme(document.documentElement.dataset.theme === "dark" ? "light" : "dark");
  });

  // Double-click any cell (header or data) to edit it manually.
  $("preview").addEventListener("dblclick", (event) => {
    const cellEl = event.target.closest(".sheet-cell");
    if (!cellEl || busy || !state.has_file) return;
    const tr = cellEl.parentElement;
    const colIndex = Array.prototype.indexOf.call(tr.children, cellEl);
    const rowIndex = parseInt(cellEl.dataset.r, 10);
    const letter = (state.colLetters || [])[colIndex - 1];
    if (!letter || !rowIndex) return;
    const cell = (state.grid[rowIndex - 1] || [])[colIndex - 1] || {};
    startCellEdit(cellEl, `${letter}${rowIndex}`, cell);
  });

  // Drag across cells to select a range for the Selection tools.
  $("preview").addEventListener("mousedown", (event) => {
    const cellEl = event.target.closest(".sheet-cell");
    if (!cellEl || busy || !state.has_file || cellEl.querySelector("input")) return;
    event.preventDefault();
    dragging = true;
    const r = parseInt(cellEl.dataset.r, 10);
    const c = parseInt(cellEl.dataset.c, 10);
    selection = { r1: r, c1: c, r2: r, c2: c };
    paintSelection();
    updateSelectionStatus();
  });
  $("preview").addEventListener("mouseover", (event) => {
    if (!dragging) return;
    const cellEl = event.target.closest(".sheet-cell");
    if (!cellEl || !cellEl.dataset.r) return;
    const r = parseInt(cellEl.dataset.r, 10);
    const c = parseInt(cellEl.dataset.c, 10);
    if (!selection || (selection.r2 === r && selection.c2 === c)) return;
    selection = { r1: selection.r1, c1: selection.c1, r2: r, c2: c };
    paintSelection();
    updateSelectionStatus();
  });
  document.addEventListener("mouseup", () => {
    if (dragging) {
      dragging = false;
      updateSelectionStatus();
    }
  });
  document.addEventListener("keydown", (event) => {
    if (!selection || busy || !state.has_file) return;
    const tag = (event.target.tagName || "").toLowerCase();
    if (tag === "input" || tag === "select" || tag === "textarea") return;
    if (event.key === "Delete" || event.key === "Backspace") {
      event.preventDefault();
      const range = selectionRangeString();
      if (range) runOp("clear_range", { range });
    }
  });


  /* Context menu on right-click in the preview table */
  $("preview").addEventListener("contextmenu", (event) => {
    const th = event.target.closest("thead th");
    const rowNum = event.target.closest("tbody th.gutter");
    const td = event.target.closest("tbody td");
    event.preventDefault();
    if (th && th.dataset.c) {
      showContextMenu(event, th, "column-header");
    } else if (rowNum && rowNum.dataset.rowIndex) {
      showContextMenu(event, rowNum, "row-number");
    } else if (td && td.dataset.r) {
      showContextMenu(event, td, "cell");
    }
  });

  /* Dismiss context menu on any left click or Escape */
  document.addEventListener("click", (event) => {
    if (!event.target.closest("#context-menu")) hideContextMenu();
  });
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") hideContextMenu();
  });

  /* Sheet switching via sheet pills */
  async function switchSheet(name) {
    if (busy) return;
    setBusy(true);
    await applyResponse(await postJson(apiUrl("/api/sheet"), { name }));
    setBusy(false);
  }

  $("file-input").addEventListener("change", async (event) => {
    const file = event.target.files[0];
    if (file) await uploadBytes(file.name, await file.arrayBuffer());
    event.target.value = "";
  });

  $("btn-open").addEventListener("click", async () => {
    if (busy) return;
    setBusy(true);
    try {
      await openExcelFile();
    } finally {
      setBusy(false);
    }
  });

  $("btn-invoice").addEventListener("click", async () => {
    if (busy) return;
    setBusy(true);
    await applyResponse(await postJson(apiUrl("/api/practice"), { type: "invoice" }));
    setBusy(false);
  });

  $("btn-ledger").addEventListener("click", async () => {
    if (busy) return;
    setBusy(true);
    await applyResponse(await postJson(apiUrl("/api/practice"), { type: "ledger" }));
    setBusy(false);
  });

  $("btn-undo").addEventListener("click", async () => {
    if (busy) return;
    setBusy(true);
    await applyResponse(await postJson(apiUrl("/api/undo")));
    setBusy(false);
  });

  $("btn-download").addEventListener("click", async () => {
    if (busy || !state.has_file) return;
    if (window.ExcelHelperAPI && window.ExcelHelperAPI.saveCurrent) {
      setBusy(true);
      try {
        const result = await window.ExcelHelperAPI.saveCurrent();
        if (result && result.saved) {
          showBanner("success", `Saved to ${result.path}`);
        } else if (result && result.error) {
          showBanner("error", result.error);
        }
      } catch (err) {
        showBanner("error", "Could not open the save dialog. Please try again.");
      } finally {
        setBusy(false);
      }
      return;
    }
    window.location.assign(apiUrl("/api/download"));
  });

  $("banner-close").addEventListener("click", hideBanner);

  /* Tab 1 · Basics */
  $("btn-make-header").addEventListener("click", () =>
    runOp("make_header", { color: $("hdr-color").value.replace("#", "") }));
  $("btn-borders").addEventListener("click", () =>
    runOp("add_borders", { range: selectionRangeString() || $("borders-range").value }));
  $("btn-nf").addEventListener("click", () => {
    const range = selectionRangeString();
    if (range) {
      runOp("number_format_range", { range, fmt: $("nf-type").value, symbol: $("nf-symbol").value });
    } else {
      runOp("number_format", { col: $("nf-col").value, fmt: $("nf-type").value, symbol: $("nf-symbol").value });
    }
  });
  $("nf-type").addEventListener("change", () => {
    $("nf-symbol").style.display = $("nf-type").value === "currency" ? "" : "none";
  });
  const alignTo = (mode) => {
    const range = selectionRangeString();
    if (range) runOp("align_range", { range, mode });
    else runOp("align_column", { col: $("al-col").value, mode });
  };
  $("btn-align-left").addEventListener("click", () => alignTo("left"));
  $("btn-align-center").addEventListener("click", () => alignTo("center"));
  $("btn-align-right").addEventListener("click", () => alignTo("right"));
  $("btn-autofit").addEventListener("click", () => runOp("autofit", {}));
  $("btn-style-col").addEventListener("click", () =>
    runOp("style_column", {
      col: $("st-col").value,
      bold: $("st-bold").checked,
      color: $("st-color").value.replace("#", ""),
    }));

  /* Tab 2 · Formulas */
  $("btn-multiply").addEventListener("click", () =>
    runOp("multiply", {
      col_a: $("mu-a").value,
      col_b: $("mu-b").value,
      result_col: $("mu-result").value,
      header: $("mu-header").value,
    }));
  $("btn-stat").addEventListener("click", () =>
    runOp("stat", { col: $("st2-col").value, stat: $("st2-stat").value }));
  $("btn-freeze").addEventListener("click", () =>
    runOp("freeze_values", { col: $("fz-col").value === "__all__" ? "" : $("fz-col").value }));
  $("btn-series").addEventListener("click", () =>
    runOp("fill_series", { col: $("fs-col").value }));
  $("btn-dec-up").addEventListener("click", () =>
    runOp("decimals", { col: $("dc-col").value, delta: 1 }));
  $("btn-dec-down").addEventListener("click", () =>
    runOp("decimals", { col: $("dc-col").value, delta: -1 }));

  /* Tab 3 · Accounts */
  $("btn-balance").addEventListener("click", () =>
    runOp("balance", {
      amount_col: $("ba-amount").value,
      paid_col: $("ba-paid").value,
      running: $("ba-running").checked,
    }));

  /* Tab 4 · Sort & Filter */
  $("btn-sort-asc").addEventListener("click", () =>
    runOp("sort", { col: $("so-col").value, desc: false }));
  $("btn-sort-desc").addEventListener("click", () =>
    runOp("sort", { col: $("so-col").value, desc: true }));
  $("btn-filter").addEventListener("click", () => runOp("filter", {}));
  $("qf-input").addEventListener("input", (event) => {
    quickFind = event.target.value;
    applyQuickFind();
  });
  $("btn-qf-clear").addEventListener("click", () => {
    $("qf-input").value = "";
    quickFind = "";
    applyQuickFind();
  });
}

function listenForLiveChanges() {
  const source = new EventSource(apiUrl("/api/events"));
  source.onmessage = (event) => {
    let message;
    try {
      message = JSON.parse(event.data);
    } catch (err) {
      return;
    }
    if (message.state && !busy) {
      state = message.state;
      render();
    }
    if (message.notice) {
      showBanner(message.notice.type === "error" ? "error" : "success", message.notice.text);
    }
  };
}

async function init() {
  wireTabs();
  wireOperations();
  render();
  const data = await fetchJson(apiUrl("/api/state"));
  if (data) {
    state = data;
    render();
  }
  listenForLiveChanges();
}

init();
