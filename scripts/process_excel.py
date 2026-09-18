"""
process_excel.py  --input <path> --output <path>
Pipeline: clean-rep, clean-proper, capital-states, dedupe, format-phone.
All formatting preserved (openpyxl read/write mode).
Optimised: values_only=True scans, batch grouped row-deletes.
"""

import argparse
import re
import shutil
from decimal import Decimal, ROUND_HALF_UP
from typing import Optional
import openpyxl

STATES   = ["CA", "MA", "NJ", "IL", "CT", "FL", "TX", "NY", "MD", "RI", "NH"]
NO_RI_NH = ["CA", "MA", "NJ", "IL", "CT", "FL", "TX", "NY", "MD"]

COL_REP   = 7          # G  (1-based)
COL_PHONE = 2          # B
CLEAN_COLS  = [2, 5, 9, 10]   # B E I J
PROPER_COLS = [4, 6]           # D F
DATA_START  = 2                # row 1 = header

_CLEAN_BAD = {chr(c) for c in range(32)}


def excel_clean(v):
    if not isinstance(v, str):
        return v
    cleaned = "".join(ch for ch in v if ch not in _CLEAN_BAD)
    return cleaned


def excel_proper(v):
    if not isinstance(v, str):
        return v
    out, prev = [], False
    for ch in v:
        if ch.isalpha():
            out.append(ch.lower() if prev else ch.upper())
            prev = True
        else:
            out.append(ch)
            prev = False
    return "".join(out)


_NUM_RE = re.compile(r"^[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?$")


def excel_number(v) -> Optional[Decimal]:
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, (int, float)):
        return Decimal(str(v))
    if isinstance(v, str):
        s = v.strip().replace(",", "")
        if _NUM_RE.match(s):
            return Decimal(s)
    return None


def format_phone(v, pattern="(###) ###-####"):
    num = excel_number(v)
    if num is None:
        return ("" if v is None else str(v)), False
    n = int(num.quantize(Decimal(1), rounding=ROUND_HALF_UP))
    digits = str(abs(n)) if n else ""
    slots = [i for i, ch in enumerate(pattern) if ch == "#"]
    out = list(pattern)
    pos = len(digits)
    for k, i in enumerate(reversed(slots)):
        last = k == len(slots) - 1
        take = digits[:pos] if last else digits[max(pos - 1, 0):pos]
        pos -= len(take)
        out[i] = take
    return ("-" if n < 0 else "") + "".join(out), True


def dedupe_key(vals):
    return tuple(
        None if (v == "" or v is None)
        else (v.lower() if isinstance(v, str) else v)
        for v in vals
    )


def _group_runs(rows):
    """[2,3,4,7,9,10] -> [(2,3),(7,1),(9,2)] (start, count) for batch delete."""
    runs = []
    for r in sorted(rows, reverse=True):   # descending so deletes don't shift unprocessed
        if runs and r == runs[-1][0] - 1:
            runs[-1] = (r, runs[-1][1] + 1)
        else:
            runs.append((r, 1))
    return runs                             # already in reverse order


# ── Operations ──────────────────────────────────────────────────────────────

def apply_clean_rep(ws):
    """Scan with values_only, write only changed cells."""
    changes = {}
    for r_i, row in enumerate(ws.iter_rows(min_row=DATA_START, values_only=True)):
        v = row[COL_REP - 1] if len(row) >= COL_REP else None
        if isinstance(v, str):
            new = excel_clean(v)
            if new != v:
                changes[DATA_START + r_i] = new
    for r, new in changes.items():
        ws.cell(row=r, column=COL_REP).value = new
    return len(changes)


def apply_clean_proper(ws):
    changes = {}   # (r, c) -> new_value
    for r_i, row in enumerate(ws.iter_rows(min_row=DATA_START, values_only=True)):
        r = DATA_START + r_i
        for ci in CLEAN_COLS:
            v = row[ci - 1] if len(row) >= ci else None
            if isinstance(v, str):
                new = excel_clean(v)
                if new != v:
                    changes[(r, ci)] = new
        for ci in PROPER_COLS:
            v = row[ci - 1] if len(row) >= ci else None
            if isinstance(v, str):
                new = excel_proper(excel_clean(v))
                if new != v:
                    changes[(r, ci)] = new
    for (r, c), new in changes.items():
        ws.cell(row=r, column=c).value = new
    return len(changes)


def apply_capital_states(ws, state_code):
    code = state_code.upper()
    find = f" {code} "
    find_lower = find.lower()
    changes = {}
    for r_i, row in enumerate(ws.iter_rows(min_row=DATA_START, values_only=True)):
        r = DATA_START + r_i
        for c_i, v in enumerate(row):
            if isinstance(v, str) and find_lower in v.lower() and find not in v:
                changes[(r, c_i + 1)] = re.sub(re.escape(find), find, v, flags=re.IGNORECASE)
    for (r, c), new in changes.items():
        ws.cell(row=r, column=c).value = new
    return len(changes)


def apply_dedupe(ws):
    seen = set()
    to_delete = []
    for r_i, row in enumerate(ws.iter_rows(min_row=DATA_START, values_only=True)):
        r = DATA_START + r_i
        if all(v is None or v == "" for v in row):
            continue
        key = dedupe_key(row)
        if key in seen:
            to_delete.append(r)
        else:
            seen.add(key)
    # Batch delete in reverse-order groups (avoids O(n²) single-row deletes)
    for start, count in _group_runs(to_delete):
        ws.delete_rows(start, count)
    return len(to_delete)


def apply_format_phone(ws):
    changes = {}
    for r_i, row in enumerate(ws.iter_rows(min_row=DATA_START, values_only=True)):
        v = row[COL_PHONE - 1] if len(row) >= COL_PHONE else None
        if v is None:
            continue
        new, converted = format_phone(v)
        if converted and new != str(v):
            changes[DATA_START + r_i] = new
    for r, new in changes.items():
        ws.cell(row=r, column=COL_PHONE).value = new
    return len(changes)


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input",  required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()

    import sys
    print("Copying…", flush=True)
    shutil.copy2(args.input, args.output)
    print("Loading workbook…", flush=True)
    wb = openpyxl.load_workbook(args.output)

    for label, states, fn in [
        ("clean-rep",       STATES,   apply_clean_rep),
        ("clean-proper",    STATES,   apply_clean_proper),
        ("capital-states",  STATES,   apply_capital_states),
        ("dedupe",          NO_RI_NH, apply_dedupe),
        ("format-phone",    NO_RI_NH, apply_format_phone),
    ]:
        print(f"{label}", flush=True)
        for s in states:
            if s not in wb.sheetnames:
                continue
            ws = wb[s]
            if label == "capital-states":
                n = fn(ws, s)
            else:
                n = fn(ws)
            print(f"  {s}: {n}", flush=True)

    print("Saving…", flush=True)
    wb.save(args.output)
    print("done", flush=True)


if __name__ == "__main__":
    main()
