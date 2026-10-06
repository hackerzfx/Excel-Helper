<div align="center">

# 📊 Excel Helper

**A friendly desktop app for everyday Excel tasks — built for people who have never used Excel formulas.**

Big buttons · Plain English · Live preview · Undo that always has your back

![Platform](https://img.shields.io/badge/platform-Windows-blue)
![Electron](https://img.shields.io/badge/built%20with-Electron-47848F?logo=electron&logoColor=white)
![Node.js](https://img.shields.io/badge/Node.js-required-339933?logo=node.js&logoColor=white)
![Output](https://img.shields.io/badge/output-real%20.xlsx-217346?logo=microsoft-excel&logoColor=white)

</div>

---

## ✨ Features

- 🖱️ **Point-and-click** — no formulas to memorize; every button does the work for you
- 👀 **Live preview** — the table updates instantly after every click
- ↩️ **Undo** — reverses the last change (last 5 steps kept)
- 📄 **Real Excel output** — real formulas and real formats that work in Excel, WPS and Google Sheets
- 🛡️ **Safe by design** — your original file is never modified; you work on a copy
- 🎓 **Practice data included** — no file? Load a sample invoice or ledger and learn

## 🚀 Quick Start

1. **Install [Node.js](https://nodejs.org)** (one-time, only needed on a new PC)
2. **Clone or download** this repository
   ```bash
   git clone https://github.com/hackerzfx/Excel-Helper.git
   cd Excel-Helper
   ```
3. **Double-click `Start Excel Helper.bat`**
4. The window opens (first start takes a few seconds while dependencies install)

## 🧭 The App in 30 Seconds

| Button | What it does |
| --- | --- |
| 📂 **Open Excel File** | Pick any `.xlsx` file on your computer |
| 🧾 **Practice Invoice** / 📒 **Practice Ledger** | Load ready-made data to learn on |
| ⬇️ **Download Result** | Save your finished file with a normal *Save As* dialog |
| ↩️ **Undo** | Reverse the last change |

The big table in the middle is a **live preview**. Click a **sheet tab** to view other sheets. If the file changes on disk (e.g. you edit it in Excel), the preview refreshes automatically.

## 🧰 Modules

### 1 · Basics — headers, borders, formatting
- **Make the header row** — bold white text on a colored background, centered and frozen while scrolling
- **Add borders** — around the whole table, or a range like `A1:E20`
- **Number format** — money (₹ $ € £), thousands commas (`1,234.56`), percent, or dates
- **Align a column** · **Auto-fit column widths** · **Style a column**

### 2 · Formulas
- **QTY × RATE** — adds a new column with real formulas like `=B2*C2`
- **Totals & stats** — SUM / AVERAGE / MAX / MIN / COUNT written below the column as a live formula
- **Freeze formulas as values** — same as *Paste Special → Values*
- **Fill serial numbers** — 1, 2, 3… down a column
- **Decimal places** — add or remove one decimal place

### 3 · Accounts Practice
- **Add Balance column** — `balance = amount − paid` as real formulas. Tick *running balance* to carry each row forward — perfect for cash books and ledgers.

**Try it:**
- *Practice Invoice* → **QTY × RATE** → **SUM** → **Make the header row**
- *Practice Ledger* → **Add Balance column** (Debit as amount, Credit as paid)

### 4 · Sort & Filter
- **Sort the table** — A→Z or Z→A by any column. Formulas in rows are converted to plain numbers while sorting so nothing breaks. Rows with an empty cell in the sort column (like a Total row) stay at the bottom, just like Excel.
- **Add filter dropdowns** — arrow buttons on every header when opened in Excel
- **Quick find** — highlights matching rows (preview only)

## 💡 Good to Know

- The preview shows the first **150 rows × 40 columns**, but *Download* always saves the **complete file**.
- Everything written is genuine Excel — no proprietary formats.

## 📦 Build a Single `.exe` (Optional)

```bat
npm run dist
```

Builds a portable `Excel Helper.exe`. The packager is downloaded once (~200 MB).

## 🔧 Troubleshooting

| Problem | Fix |
| --- | --- |
| "Node.js is required" | Install [Node.js](https://nodejs.org) and run the `.bat` again |
| Window doesn't open | Wait for the first-run install to finish, then run the `.bat` again |
| Preview says file is locked | The file is open in Excel — close it there; the preview refreshes automatically |
| Button shows an error | Read the red message — it says exactly what's missing. *Undo* restores the last step |

## 🗂 Project Structure

```text
Start Excel Helper.bat     double-click launcher
main.js                    Electron main process (window + dialogs)
preload.js                 safe bridge: port + native dialogs
server.js                  embedded API server (127.0.0.1)
excel/
├── operations.js          all 14 operations
├── formula_eval.js        tiny Excel formula calculator
├── practice_data.js       practice invoice & ledger
└── workbook_store.js      current file, undo, live preview state
verify-engine.js           self-test
```

Run the self-test:

```bash
node verify-engine.js
```

## 🤝 Contributing

Issues and pull requests are welcome. Please run `node verify-engine.js` before submitting a PR.

## 📄 License

Add your license here (e.g. [MIT](https://choosealicense.com/licenses/mit/)).
