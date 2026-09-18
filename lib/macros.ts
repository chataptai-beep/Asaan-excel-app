import { CellValue, COL, MacroOp, MacroResult, RowData } from "./types";

// ---- Pure transforms ----

export function excelClean(value: CellValue): CellValue {
  if (typeof value !== "string") return value;
  return value.split("").filter((ch) => ch.charCodeAt(0) >= 32).join("");
}

export function excelProper(value: CellValue): CellValue {
  if (typeof value !== "string") return value;
  let result = "";
  let prevIsLetter = false;
  for (const ch of value) {
    if (/[a-zA-Z]/.test(ch)) {
      result += prevIsLetter ? ch.toLowerCase() : ch.toUpperCase();
      prevIsLetter = true;
    } else {
      result += ch;
      prevIsLetter = false;
    }
  }
  return result;
}

function toNumber(value: CellValue): number | null {
  if (typeof value === "number") return value;
  if (typeof value === "string") {
    const s = value.trim().replace(/,/g, "");
    if (/^[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?$/.test(s)) return parseFloat(s);
  }
  return null;
}

export function formatPhone(value: CellValue): string {
  if (value === null || value === undefined) return "";
  const num = toNumber(value);
  if (num === null) return String(value);
  const n = Math.round(num);
  const digits = n === 0 ? "" : String(Math.abs(n));
  const pattern = "(###) ###-####";
  const slots: number[] = [];
  for (let i = 0; i < pattern.length; i++) if (pattern[i] === "#") slots.push(i);
  const out = pattern.split("");
  let p = digits.length;
  for (let k = slots.length - 1; k >= 0; k--) {
    const idx = slots[k];
    if (k === 0) { out[idx] = digits.slice(0, p); }
    else { out[idx] = p > 0 ? digits[p - 1] : ""; p = Math.max(0, p - 1); }
  }
  return (n < 0 ? "-" : "") + out.join("");
}

// ---- Helpers ----

function data(rows: RowData[]): RowData[] {
  return rows.filter((r) => !r.isDivider);
}

function withCells(row: RowData, cells: CellValue[]): RowData {
  return { ...row, cells };
}

function dedupeKey(cells: CellValue[]): string {
  return JSON.stringify(
    cells.map((v) =>
      v === null || v === undefined || v === "" ? null :
      typeof v === "string" ? v.toLowerCase() : v
    )
  );
}

// ---- Operations ----

export function sortRows(rows: RowData[], descending: boolean): { rows: RowData[]; changed: number } {
  const col = COL.E;
  const sorted = [...rows].sort((a, b) => {
    if (a.isDivider !== b.isDivider) return a.isDivider ? -1 : 1;
    if (a.isDivider && b.isDivider) return 0;
    const av = a.cells[col] ?? "";
    const bv = b.cells[col] ?? "";
    const an = toNumber(av); const bn = toNumber(bv);
    let cmp = (an !== null && bn !== null) ? an - bn : String(av).localeCompare(String(bv));
    return descending ? -cmp : cmp;
  });
  return { rows: sorted, changed: data(rows).length };
}

export function cleanRep(rows: RowData[]): { rows: RowData[]; changed: number } {
  let changed = 0;
  const result = rows.map((row) => {
    if (row.isDivider) return row;
    const cells = [...row.cells];
    const old = cells[COL.G];
    const cleaned = excelClean(old);
    if (cleaned !== old) { cells[COL.G] = cleaned; changed++; }
    return withCells(row, cells);
  });
  return { rows: result, changed };
}

export function cleanProper(rows: RowData[]): { rows: RowData[]; changed: number } {
  const cleanCols = [COL.B, COL.E, COL.I, COL.J];
  const properCols = [COL.D, COL.F];
  let changed = 0;
  const result = rows.map((row) => {
    if (row.isDivider) return row;
    const cells = [...row.cells];
    for (const c of cleanCols) {
      const old = cells[c]; const cleaned = excelClean(old);
      if (cleaned !== old) { cells[c] = cleaned; changed++; }
    }
    for (const c of properCols) {
      const old = cells[c]; const properd = excelProper(excelClean(old));
      if (properd !== old) { cells[c] = properd; changed++; }
    }
    return withCells(row, cells);
  });
  return { rows: result, changed };
}

export function formatPhoneCol(rows: RowData[]): { rows: RowData[]; changed: number } {
  let changed = 0;
  const result = rows.map((row) => {
    if (row.isDivider) return row;
    const cells = [...row.cells];
    const old = cells[COL.B];
    if (old !== null && old !== undefined && toNumber(old) !== null) {
      const formatted = formatPhone(old);
      if (formatted !== String(old)) { cells[COL.B] = formatted; changed++; }
    }
    return withCells(row, cells);
  });
  return { rows: result, changed };
}

export function capitalStates(rows: RowData[], stateCode: string): { rows: RowData[]; changed: number } {
  const code = stateCode.toUpperCase();
  const find = ` ${code} `;
  const findLower = find.toLowerCase();
  let changed = 0;
  const result = rows.map((row) => {
    if (row.isDivider) return row;
    const cells = [...row.cells];
    let rowChanged = false;
    for (let c = 0; c < cells.length; c++) {
      const v = cells[c];
      if (typeof v === "string" && v.toLowerCase().includes(findLower) && !v.includes(find)) {
        const esc = find.replace(/[-[\]{}()*+?.,\\^$|#\s]/g, "\\$&");
        cells[c] = v.replace(new RegExp(esc, "gi"), find);
        rowChanged = true;
      }
    }
    if (rowChanged) changed++;
    return withCells(row, cells);
  });
  return { rows: result, changed };
}

export function dedupeRows(rows: RowData[]): { rows: RowData[]; changed: number } {
  const seen = new Set<string>();
  const result: RowData[] = [];
  let removed = 0;
  for (const row of rows) {
    if (row.isDivider) { result.push(row); continue; }
    const allBlank = row.cells.every((v) => v === null || v === undefined || v === "");
    if (allBlank) { result.push(row); continue; }
    const key = dedupeKey(row.cells);
    if (seen.has(key)) { removed++; }
    else { seen.add(key); result.push(row); }
  }
  return { rows: result, changed: removed };
}

export function removeOutliers(rows: RowData[]): { rows: RowData[]; changed: number } {
  // Drop rows where Customer Name (col D) has no run of 3+ consecutive letters.
  // Real names always have a word ≥3 chars ("Joe", "William", etc.).
  // Phone numbers — even with embedded state codes like "(626) 862 CA 9254" — never do.
  let removed = 0;
  const result = rows.filter((row) => {
    if (row.isDivider) return true;
    const name = row.cells?.[COL.D];
    if (name === null || name === undefined || name === "") return true;
    const s = String(name).trim();
    if (s && !/[a-zA-Z]{3}/.test(s)) { removed++; return false; }
    return true;
  });
  return { rows: result, changed: removed };
}

export function moveRedToTop(rows: RowData[]): { rows: RowData[]; changed: number } {
  // Split into sections by divider rows, move red rows to top of each section
  const result: RowData[] = [];
  let section: RowData[] = [];
  let moved = 0;

  function flushSection() {
    const redInSec = section.filter((r) => r.isRed);
    const normalInSec = section.filter((r) => !r.isRed);
    moved += redInSec.length;
    result.push(...redInSec, ...normalInSec);
    section = [];
  }

  for (const row of rows) {
    if (row.isDivider) {
      flushSection();
      result.push(row);
    } else {
      section.push(row);
    }
  }
  flushSection();

  return { rows: result, changed: moved };
}

// ---- Dispatcher ----

export function applyMacro(
  rows: RowData[],
  op: MacroOp,
  sheetName: string
): { rows: RowData[]; result: MacroResult } {
  let out: { rows: RowData[]; changed: number };

  switch (op) {
    case "sort-asc":
      out = sortRows(rows, false);
      return { rows: out.rows, result: { op, changed: out.changed, message: `Sorted ${data(rows).length} rows ascending by Lead ID` } };
    case "sort-desc":
      out = sortRows(rows, true);
      return { rows: out.rows, result: { op, changed: out.changed, message: `Sorted ${data(rows).length} rows descending by Lead ID` } };
    case "clean-rep":
      out = cleanRep(rows);
      return { rows: out.rows, result: { op, changed: out.changed, message: `Clean Rep: ${out.changed} cell(s) changed` } };
    case "clean-proper":
      out = cleanProper(rows);
      return { rows: out.rows, result: { op, changed: out.changed, message: `Clean+Proper: ${out.changed} cell(s) changed` } };
    case "format-phone":
      out = formatPhoneCol(rows);
      return { rows: out.rows, result: { op, changed: out.changed, message: `Format Phone: ${out.changed} cell(s) changed` } };
    case "capital-states":
      out = capitalStates(rows, sheetName);
      return { rows: out.rows, result: { op, changed: out.changed, message: `Capital States: ${out.changed} row(s) updated` } };
    case "dedupe":
      out = dedupeRows(rows);
      return { rows: out.rows, result: { op, changed: out.changed, message: `Dedupe: ${out.changed} duplicate(s) removed` } };
    case "remove-outliers":
      out = removeOutliers(rows);
      return { rows: out.rows, result: { op, changed: out.changed, message: `Remove Outliers: ${out.changed} row(s) removed` } };
    case "move-red-top":
      out = moveRedToTop(rows);
      return { rows: out.rows, result: { op, changed: out.changed, message: `Red rows moved to top: ${out.changed} row(s)` } };
    case "run-all": {
      // Run the full pipeline: clean → dedupe → format phone → capital states → move red to top
      let r = rows;
      let totalChanged = 0;
      const steps: MacroOp[] = ["clean-proper", "clean-rep", "remove-outliers", "capital-states", "dedupe", "format-phone", "move-red-top"];
      for (const step of steps) {
        const res = applyMacro(r, step, sheetName);
        r = res.rows;
        totalChanged += res.result.changed;
      }
      return { rows: r, result: { op, changed: totalChanged, message: `Run All complete: ${totalChanged} total change(s)` } };
    }
  }
}
