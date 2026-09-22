"use client";

import { useCallback, useState, useRef, useMemo } from "react";
import { parseWorkbook } from "@/lib/excel-utils";
import { applyMacro, moveRedToTop } from "@/lib/macros";
import {
  MacroOp, MacroResult, RowData, SheetData,
  VISIBLE_COLS, COLUMN_LABELS, WorkbookData, KNOWN_STATES,
} from "@/lib/types";

const MACROS: { op: MacroOp; label: string; desc: string; color: string }[] = [
  { op: "sort-asc",       label: "Sort ↑",         desc: "Macro 2 – Sort ascending by Lead ID",       color: "#2563eb" },
  { op: "sort-desc",      label: "Sort ↓",         desc: "Macro 6/7 – Sort descending by Lead ID",    color: "#7c3aed" },
  { op: "clean-rep",      label: "Clean Rep",      desc: "Macro 5 – CLEAN() on Rep column",           color: "#0891b2" },
  { op: "clean-proper",   label: "Clean+Proper",   desc: "Macro 8/9 – CLEAN + PROPER on name/addr",  color: "#0369a1" },
  { op: "format-phone",   label: "Format Phone",   desc: "Macro 11 – (###) ###-#### formatting",      color: "#059669" },
  { op: "capital-states", label: "Capital States", desc: "Macro 10 – ' Ny ' → ' NY '",               color: "#b45309" },
  { op: "dedupe",         label: "Remove Dupes",      desc: "Macro 12 – Remove duplicate rows",                        color: "#dc2626" },
  { op: "remove-outliers", label: "Remove Outliers", desc: "Remove rows where Customer Name has no letters (phone/number noise)", color: "#7e22ce" },
  { op: "move-red-top",   label: "🔴 Red to Top",  desc: "Move red-highlighted rows to top of section",             color: "#991b1b" },
];

const ROW_H = 28; // px per data row
const PAGE_SIZE = 100; // virtual rows to render

export default function Home() {
  const [workbook, setWorkbook] = useState<WorkbookData | null>(null);
  const [sheets, setSheets] = useState<SheetData[]>([]);
  const [activeSheet, setActiveSheet] = useState<string>("");
  const [log, setLog] = useState<MacroResult[]>([]);
  const [dragOver, setDragOver] = useState(false);
  const [loading, setLoading] = useState(false);
  const [loadMsg, setLoadMsg] = useState("Parsing file…");
  const [search, setSearch] = useState("");
  const [sortCol, setSortCol] = useState<number | null>(null);
  const [sortDir, setSortDir] = useState<"asc" | "desc">("asc");
  const [scrollTop, setScrollTop] = useState(0);
  const [exporting, setExporting] = useState(false);
  const [exportMsg, setExportMsg] = useState("⬇ Export Processed");
  const inputRef = useRef<HTMLInputElement>(null);
  const tableContainerRef = useRef<HTMLDivElement>(null);
  const rawFileRef = useRef<File | null>(null);

  const exportProcessed = () => {
    if (!rawFileRef.current) return;
    setExporting(true);
    setExportMsg("Reading file…");

    rawFileRef.current.arrayBuffer().then((buf) => {
      const worker = new Worker("/export.worker.js");
      worker.onmessage = (e) => {
        const { type, msg, buffer, message } = e.data;
        if (type === "progress") {
          setExportMsg(msg);
        } else if (type === "done") {
          worker.terminate();
          const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
          const url = URL.createObjectURL(blob);
          const a = document.createElement("a");
          a.href = url;
          a.download = rawFileRef.current!.name.replace(/\.[^/.]+$/, "") + "_clean.xlsx";
          a.click();
          URL.revokeObjectURL(url);
          setExporting(false);
          setExportMsg("⬇ Export Processed");
        } else if (type === "error") {
          worker.terminate();
          alert("Export failed: " + message);
          setExporting(false);
          setExportMsg("⬇ Export Processed");
        }
      };
      worker.onerror = (err) => {
        worker.terminate();
        alert("Worker error: " + err.message);
        setExporting(false);
        setExportMsg("⬇ Export Processed");
      };
      // Pass red-row numbers derived from the already-loaded display data.
      // excel.worker.js already detected them correctly; re-using that result
      // avoids duplicating (and potentially mismatching) CF / fill logic.
      const redBySheet: Record<string, number[]> = {};
      const dividerBySheet: Record<string, number[]> = {};
      for (const sheet of sheets) {
        const redRows = sheet.rows.filter(r => r.isRed).map(r => r.originalRow);
        if (redRows.length > 0) redBySheet[sheet.name] = redRows;
        const dividerRows = sheet.rows.filter(r => r.isDivider).map(r => r.originalRow);
        if (dividerRows.length > 0) dividerBySheet[sheet.name] = dividerRows;
      }
      worker.postMessage({ buffer: buf, redBySheet, dividerBySheet }, [buf]);
    }).catch((e) => {
      alert("Failed to read file: " + String(e));
      setExporting(false);
      setExportMsg("⬇ Export Processed");
    });
  };

  const loadFile = useCallback(async (file: File) => {
    if (!file.name.match(/\.(xlsx|xls|xlsm|csv)$/i)) {
      alert("Please upload an Excel file (.xlsx, .xls, .xlsm) or CSV.");
      return;
    }
    rawFileRef.current = file;
    setLoading(true);
    setLoadMsg("Reading file…");
    try {
      const wb = await parseWorkbook(file, (msg) => setLoadMsg(msg));
      setWorkbook(wb);
      setSheets(wb.sheets);
      const first = wb.sheets.find((s) => s.isStateSheet) ?? wb.sheets[0];
      setActiveSheet(first?.name ?? "");
      setLog([]);
      setSortCol(null);
      setSearch("");
      setScrollTop(0);
    } catch (e) {
      alert("Failed to parse: " + String(e));
    } finally {
      setLoading(false);
    }
  }, []);

  const onFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) loadFile(file);
    e.target.value = "";
  };

  const onDrop = (e: React.DragEvent) => {
    e.preventDefault(); setDragOver(false);
    const file = e.dataTransfer.files?.[0];
    if (file) loadFile(file);
  };

  const runMacro = (op: MacroOp) => {
    const idx = sheets.findIndex((s) => s.name === activeSheet);
    if (idx === -1) return;
    const { rows, result } = applyMacro(sheets[idx].rows, op, sheets[idx].name);
    setSheets((prev) => prev.map((s, i) => i === idx ? { ...s, rows } : s));
    setLog((prev) => [result, ...prev]);
    setSortCol(null);
  };

  const runOnAll = (op: MacroOp) => {
    let total = 0;
    setSheets((prev) => prev.map((s) => {
      if (!s.isStateSheet) return s;
      const { rows, result } = applyMacro(s.rows, op, s.name);
      total += result.changed;
      return { ...s, rows };
    }));
    setLog((prev) => [{
      op, changed: total,
      message: `[All states] ${MACROS.find((m) => m.op === op)?.label}: ${total} change(s)`,
    }, ...prev]);
    setSortCol(null);
  };

  const runAll = () => {
    setSheets((prev) => {
      let total = 0;
      const next = prev.map((s) => {
        if (!s.isStateSheet) return s;
        const { rows, result } = applyMacro(s.rows, "run-all", s.name);
        total += result.changed;
        return { ...s, rows };
      });
      setLog((l) => [{ op: "run-all", changed: total, message: `Run All on all state sheets: ${total} total change(s)` }, ...l]);
      return next;
    });
    setSortCol(null);
  };

  const currentSheet = sheets.find((s) => s.name === activeSheet);
  const stateSheets = sheets.filter((s) => s.isStateSheet);
  const otherSheets = sheets.filter((s) => !s.isStateSheet);

  // Filtered + display-sorted rows
  const displayRows = useMemo(() => {
    let rows = currentSheet?.rows ?? [];
    if (search.trim()) {
      const q = search.trim().toLowerCase();
      rows = rows.filter((r) => r.isDivider || r.cells?.some((v) => v !== null && String(v).toLowerCase().includes(q)));
    }
    if (sortCol !== null) {
      rows = [...rows].sort((a, b) => {
        if (a.isDivider !== b.isDivider) return a.isDivider ? -1 : 1;
        if (a.isDivider && b.isDivider) return 0;
        const av = a.cells?.[sortCol] ?? ""; const bv = b.cells?.[sortCol] ?? "";
        const cmp = String(av).localeCompare(String(bv), undefined, { numeric: true });
        return sortDir === "asc" ? cmp : -cmp;
      });
    } else {
      // Default view mirrors the export exactly: per divider section, red rows on
      // top then the rest, both groups sorted by the column-E lead number desc.
      rows = moveRedToTop(rows).rows;
    }
    return rows;
  }, [currentSheet, search, sortCol, sortDir]);

  const toggleSort = (col: number) => {
    if (sortCol === col) setSortDir((d) => d === "asc" ? "desc" : "asc");
    else { setSortCol(col); setSortDir("asc"); }
  };

  const totalData = currentSheet?.rows.filter((r) => !r.isDivider).length ?? 0;
  const redCount = currentSheet?.rows.filter((r) => r.isRed).length ?? 0;
  const dividerCount = currentSheet?.rows.filter((r) => r.isDivider).length ?? 0;

  // ---------------------------------------------------------------- Upload screen
  if (!workbook && !loading) {
    return (
      <div style={{ minHeight: "100vh", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 28, padding: 32, background: "#f8fafc" }}>
        <div style={{ textAlign: "center" }}>
          <h1 style={{ fontSize: 28, fontWeight: 800, margin: "0 0 6px", color: "#1e293b" }}>Asaan Excel – Lead Processor</h1>
          <p style={{ color: "#64748b", margin: 0, fontSize: 14 }}>Upload your SR Leads workbook · clean, sort, highlight and export</p>
        </div>
        <div
          onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
          onDragLeave={() => setDragOver(false)}
          onDrop={onDrop}
          onClick={() => inputRef.current?.click()}
          style={{ width: 440, maxWidth: "90vw", border: `2px dashed ${dragOver ? "#2563eb" : "#cbd5e1"}`, borderRadius: 16, padding: "48px 32px", textAlign: "center", cursor: "pointer", background: dragOver ? "#eff6ff" : "#fff", transition: "all .15s" }}
        >
          <div style={{ fontSize: 48, marginBottom: 12 }}>📊</div>
          <p style={{ fontWeight: 600, fontSize: 16, margin: "0 0 6px", color: "#1e293b" }}>Drop Excel file here</p>
          <p style={{ color: "#64748b", margin: "0 0 20px", fontSize: 13 }}>or click to browse · .xlsx, .xls, .xlsm</p>
          <button style={{ background: "#2563eb", color: "#fff", border: "none", borderRadius: 8, padding: "10px 24px", fontWeight: 700, cursor: "pointer", fontSize: 14 }}>Choose File</button>
          <input ref={inputRef} type="file" accept=".xlsx,.xls,.xlsm,.csv" style={{ display: "none" }} onChange={onFileChange} />
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(190px,1fr))", gap: 10, maxWidth: 720, width: "100%" }}>
          {MACROS.map((m) => (
            <div key={m.op} style={{ background: "#fff", border: "1px solid #e2e8f0", borderRadius: 10, padding: "10px 14px" }}>
              <div style={{ fontWeight: 700, fontSize: 12, color: m.color, marginBottom: 3 }}>{m.label}</div>
              <div style={{ fontSize: 11, color: "#64748b" }}>{m.desc}</div>
            </div>
          ))}
        </div>
      </div>
    );
  }

  if (loading) {
    return (
      <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", flexDirection: "column", gap: 16, background: "#f8fafc" }}>
        <div style={{ width: 48, height: 48, border: "4px solid #e2e8f0", borderTop: "4px solid #2563eb", borderRadius: "50%", animation: "spin 0.8s linear infinite" }} />
        <p style={{ color: "#64748b", margin: 0 }}>{loadMsg}</p>
        <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
      </div>
    );
  }

  // ---------------------------------------------------------------- Main app
  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100vh", overflow: "hidden", background: "#f8fafc" }}>
      {/* ── Header ── */}
      <header style={{ height: 52, background: "#0f172a", display: "flex", alignItems: "center", gap: 14, padding: "0 18px", flexShrink: 0, boxShadow: "0 1px 4px rgba(0,0,0,.4)" }}>
        <span style={{ fontWeight: 800, color: "#fff", fontSize: 15, letterSpacing: "-0.01em" }}>📊 Asaan Excel</span>
        <span style={{ color: "#64748b", fontSize: 12, flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{workbook?.fileName}</span>
        <button
          onClick={runAll}
          style={{ background: "#991b1b", color: "#fff", border: "none", borderRadius: 6, padding: "6px 14px", cursor: "pointer", fontSize: 12, fontWeight: 700 }}
          title="Run full pipeline on all state sheets: Clean → Dedupe → Format Phone → Capital States → Red to Top"
        >
          ⚡ Run All
        </button>
        <button onClick={() => inputRef.current?.click()} style={{ background: "#1e293b", color: "#94a3b8", border: "1px solid #334155", borderRadius: 6, padding: "6px 12px", cursor: "pointer", fontSize: 12 }}>Open File</button>
        <button
          onClick={exportProcessed}
          disabled={exporting || !rawFileRef.current}
          style={{ background: exporting ? "#64748b" : "#16a34a", color: "#fff", border: "none", borderRadius: 6, padding: "6px 14px", cursor: exporting ? "wait" : "pointer", fontSize: 12, fontWeight: 700, opacity: !rawFileRef.current ? 0.5 : 1, minWidth: 160 }}
          title="Run clean-rep, clean-proper, capital-states, dedupe, format-phone on the original file and download — runs entirely in your browser"
        >
          {exportMsg}
        </button>
        <input ref={inputRef} type="file" accept=".xlsx,.xls,.xlsm,.csv" style={{ display: "none" }} onChange={onFileChange} />
      </header>

      <div style={{ display: "flex", flex: 1, overflow: "hidden" }}>
        {/* ── Sidebar ── */}
        <aside style={{ width: 170, background: "#fff", borderRight: "1px solid #e2e8f0", display: "flex", flexDirection: "column", overflow: "hidden", flexShrink: 0 }}>
          <div style={{ padding: "10px 12px 4px", fontSize: 10, fontWeight: 800, color: "#94a3b8", textTransform: "uppercase", letterSpacing: ".06em" }}>State Sheets</div>
          <div style={{ flex: 1, overflowY: "auto" }}>
            {stateSheets.map((s) => {
              const rc = s.rows.filter((r) => r.isRed).length;
              const active = activeSheet === s.name;
              return (
                <button key={s.name} onClick={() => { setActiveSheet(s.name); setSortCol(null); setSearch(""); setScrollTop(0); if (tableContainerRef.current) tableContainerRef.current.scrollTop = 0; }}
                  style={{ display: "flex", width: "100%", textAlign: "left", padding: "7px 12px", border: "none", background: active ? "#eff6ff" : "transparent", borderLeft: `3px solid ${active ? "#2563eb" : "transparent"}`, fontWeight: active ? 700 : 400, color: active ? "#1d4ed8" : "#374151", cursor: "pointer", fontSize: 13, alignItems: "center", gap: 4 }}>
                  <span style={{ flex: 1 }}>{s.name}</span>
                  {rc > 0 && <span style={{ background: "#fee2e2", color: "#dc2626", borderRadius: 4, padding: "1px 5px", fontSize: 10, fontWeight: 700 }}>{rc}</span>}
                  <span style={{ color: "#94a3b8", fontSize: 10 }}>{s.rows.filter(r => !r.isDivider).length}</span>
                </button>
              );
            })}
            {otherSheets.length > 0 && (
              <>
                <div style={{ padding: "10px 12px 4px", fontSize: 10, fontWeight: 800, color: "#94a3b8", textTransform: "uppercase", letterSpacing: ".06em" }}>Other</div>
                {otherSheets.map((s) => {
                  const active = activeSheet === s.name;
                  return (
                    <button key={s.name} onClick={() => { setActiveSheet(s.name); setSortCol(null); setSearch(""); setScrollTop(0); if (tableContainerRef.current) tableContainerRef.current.scrollTop = 0; }}
                      style={{ display: "block", width: "100%", textAlign: "left", padding: "7px 12px", border: "none", background: active ? "#f5f3ff" : "transparent", borderLeft: `3px solid ${active ? "#7c3aed" : "transparent"}`, fontWeight: active ? 700 : 400, color: active ? "#5b21b6" : "#374151", cursor: "pointer", fontSize: 13 }}>
                      {s.name} <span style={{ float: "right", color: "#94a3b8", fontSize: 10 }}>{s.rows.filter(r => !r.isDivider).length}</span>
                    </button>
                  );
                })}
              </>
            )}
          </div>
        </aside>

        {/* ── Main ── */}
        <main style={{ flex: 1, display: "flex", flexDirection: "column", overflow: "hidden" }}>
          {/* Macro toolbar */}
          <div style={{ padding: "8px 14px", background: "#fff", borderBottom: "1px solid #e2e8f0", display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center", flexShrink: 0 }}>
            <span style={{ fontSize: 12, fontWeight: 700, color: "#374151", marginRight: 2 }}>{activeSheet}:</span>
            {MACROS.map((m) => (
              <div key={m.op} style={{ display: "flex", gap: 0 }} title={m.desc}>
                <button onClick={() => runMacro(m.op)}
                  style={{ background: m.color, color: "#fff", border: "none", borderRadius: "5px 0 0 5px", padding: "5px 9px", cursor: "pointer", fontSize: 11, fontWeight: 700, whiteSpace: "nowrap" }}>
                  {m.label}
                </button>
                <button onClick={() => runOnAll(m.op)} title={m.desc + " (all state sheets)"}
                  style={{ background: m.color, color: "#fff", border: "none", borderLeft: "1px solid rgba(255,255,255,.25)", borderRadius: "0 5px 5px 0", padding: "5px 6px", cursor: "pointer", fontSize: 10, opacity: .85 }}>
                  All
                </button>
              </div>
            ))}
            <div style={{ flex: 1 }} />
            <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search…"
              style={{ padding: "4px 10px", border: "1px solid #cbd5e1", borderRadius: 5, fontSize: 12, width: 140, outline: "none" }} />
          </div>

          {/* Legend */}
          <div style={{ padding: "4px 14px", background: "#fafafa", borderBottom: "1px solid #f1f5f9", display: "flex", gap: 16, alignItems: "center", flexShrink: 0 }}>
            <span style={{ fontSize: 11, color: "#64748b" }}>
              <span style={{ display: "inline-block", width: 10, height: 10, background: "#fee2e2", border: "1px solid #fca5a5", borderRadius: 2, marginRight: 4, verticalAlign: "middle" }} />
              Red rows: {redCount}
            </span>
            <span style={{ fontSize: 11, color: "#64748b" }}>
              <span style={{ display: "inline-block", width: 10, height: 10, background: "#f1f5f9", border: "1px solid #cbd5e1", borderRadius: 2, marginRight: 4, verticalAlign: "middle" }} />
              Sections: {dividerCount}
            </span>
            <span style={{ fontSize: 11, color: "#64748b" }}>Data rows: {totalData}</span>
          </div>

          {/* Table – virtual scroll: only render rows visible in the viewport */}
          {(() => {
            const containerH = tableContainerRef.current?.clientHeight ?? 700;
            const BUFFER = 15;
            const vsStart = Math.max(0, Math.floor(scrollTop / ROW_H) - BUFFER);
            const vsEnd   = Math.min(displayRows.length, Math.ceil((scrollTop + containerH) / ROW_H) + BUFFER);
            const padTop  = vsStart * ROW_H;
            const padBot  = Math.max(0, (displayRows.length - vsEnd) * ROW_H);
            const visible = displayRows.slice(vsStart, vsEnd);
            return (
              <div
                ref={tableContainerRef}
                style={{ flex: 1, overflow: "auto" }}
                onScroll={(e) => setScrollTop(e.currentTarget.scrollTop)}
              >
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                  <thead style={{ position: "sticky", top: 0, zIndex: 2 }}>
                    <tr style={{ background: "#0f172a" }}>
                      <th style={{ padding: "7px 10px", color: "#475569", fontWeight: 500, textAlign: "left", width: 36, fontSize: 11 }}>#</th>
                      {VISIBLE_COLS.map((col) => (
                        <th key={col} onClick={() => toggleSort(col)}
                          style={{ padding: "7px 10px", color: sortCol === col ? "#93c5fd" : "#64748b", fontWeight: 700, textAlign: "left", cursor: "pointer", userSelect: "none", whiteSpace: "nowrap", fontSize: 11, borderLeft: "1px solid #1e293b" }}>
                          {COLUMN_LABELS[col]}{sortCol === col ? (sortDir === "asc" ? " ↑" : " ↓") : " ⇅"}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {displayRows.length === 0 ? (
                      <tr><td colSpan={VISIBLE_COLS.length + 1} style={{ padding: 40, textAlign: "center", color: "#94a3b8" }}>No rows{search ? " match your search" : ""}</td></tr>
                    ) : (
                      <>
                        {padTop > 0 && <tr style={{ height: padTop }}><td colSpan={VISIBLE_COLS.length + 1} /></tr>}
                        {visible.map((row, i) => <TableRow key={`${activeSheet}-${vsStart + i}`} row={row} idx={vsStart + i} />)}
                        {padBot > 0 && <tr style={{ height: padBot }}><td colSpan={VISIBLE_COLS.length + 1} /></tr>}
                      </>
                    )}
                  </tbody>
                </table>
              </div>
            );
          })()}

          {/* Status bar */}
          <div style={{ height: 30, background: "#fff", borderTop: "1px solid #e2e8f0", display: "flex", alignItems: "center", padding: "0 14px", gap: 16, flexShrink: 0 }}>
            <span style={{ fontSize: 11, color: "#64748b" }}>
              {displayRows.filter(r => !r.isDivider).length} row{displayRows.filter(r => !r.isDivider).length !== 1 ? "s" : ""}
              {search ? ` filtered` : ""}
            </span>
            {log.length > 0 && (
              <span style={{ fontSize: 11, color: "#16a34a", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                ✓ {log[0].message}
              </span>
            )}
          </div>
        </main>

        {/* ── Log panel ── */}
        {log.length > 0 && (
          <aside style={{ width: 240, background: "#fff", borderLeft: "1px solid #e2e8f0", display: "flex", flexDirection: "column", flexShrink: 0, overflow: "hidden" }}>
            <div style={{ padding: "9px 12px", borderBottom: "1px solid #e2e8f0", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <span style={{ fontSize: 11, fontWeight: 800, color: "#374151", textTransform: "uppercase", letterSpacing: ".04em" }}>Log</span>
              <button onClick={() => setLog([])} style={{ background: "none", border: "none", color: "#94a3b8", cursor: "pointer", fontSize: 11 }}>Clear</button>
            </div>
            <div style={{ flex: 1, overflowY: "auto", padding: 6 }}>
              {log.map((entry, i) => (
                <div key={i} style={{ padding: "7px 9px", borderRadius: 5, marginBottom: 3, background: entry.changed > 0 ? "#f0fdf4" : "#f8fafc", borderLeft: `3px solid ${entry.changed > 0 ? "#16a34a" : "#94a3b8"}` }}>
                  <div style={{ fontSize: 10, fontWeight: 800, color: "#374151", marginBottom: 2, textTransform: "uppercase", letterSpacing: ".03em" }}>
                    {MACROS.find((m) => m.op === entry.op)?.label ?? entry.op}
                  </div>
                  <div style={{ fontSize: 11, color: "#64748b" }}>{entry.message}</div>
                </div>
              ))}
            </div>
          </aside>
        )}
      </div>
    </div>
  );
}

function TableRow({ row, idx }: { row: RowData; idx: number }) {
  if (row.isDivider) {
    const label = row.cells?.find((v) => v !== null && /\d{1,2}\.\d{1,2}\.\d{2,4}/.test(String(v)));
    return (
      <tr style={{ background: "#e2e8f0" }}>
        <td style={{ padding: "4px 10px", color: "#94a3b8", fontSize: 10, fontStyle: "italic" }}>§</td>
        <td colSpan={VISIBLE_COLS.length}
          style={{ padding: "4px 10px", fontSize: 11, fontWeight: 700, color: "#475569", letterSpacing: ".04em" }}>
          ── {label ?? "Section"} ──
        </td>
      </tr>
    );
  }

  const bg = row.isRed
    ? (idx % 2 === 0 ? "#fee2e2" : "#fecaca")
    : (idx % 2 === 0 ? "#fff" : "#f8fafc");

  const borderColor = row.isRed ? "#fca5a5" : "#f1f5f9";

  return (
    <tr style={{ borderBottom: `1px solid ${borderColor}`, background: bg }}>
      <td style={{ padding: "4px 10px", color: "#94a3b8", fontSize: 10 }}>{idx + 1}</td>
      {VISIBLE_COLS.map((col) => {
        const v = row.cells?.[col] ?? null;
        return (
          <td key={col} title={v !== null && v !== undefined ? String(v) : ""}
            style={{ padding: "4px 10px", maxWidth: 200, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", borderLeft: `1px solid ${borderColor}`, color: row.isRed ? "#7f1d1d" : "#1e293b", fontWeight: row.isRed ? 500 : 400 }}>
            {v !== null && v !== undefined ? String(v) : <span style={{ color: "#d1d5db" }}>—</span>}
          </td>
        );
      })}
    </tr>
  );
}
