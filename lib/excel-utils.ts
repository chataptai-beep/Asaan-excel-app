import * as XLSX from "xlsx";
import { SheetData, WorkbookData } from "./types";

export function parseWorkbook(
  file: File,
  onProgress?: (msg: string) => void
): Promise<WorkbookData> {
  return new Promise((resolve, reject) => {
    const worker = new Worker("/excel.worker.js");

    worker.onmessage = (e) => {
      const { type, msg, sheets, message } = e.data;
      if (type === "progress") {
        onProgress?.(msg);
      } else if (type === "done") {
        worker.terminate();
        resolve({ sheets: sheets as SheetData[], fileName: file.name });
      } else if (type === "error") {
        worker.terminate();
        reject(new Error(message));
      }
    };

    worker.onerror = (err) => {
      worker.terminate();
      reject(err);
    };

    file.arrayBuffer().then((buffer) => {
      worker.postMessage({ buffer }, [buffer]); // transfer ownership
    });
  });
}

export function exportWorkbook(sheets: SheetData[], fileName: string): void {
  const wb = XLSX.utils.book_new();
  for (const sheet of sheets) {
    const aoa = [sheet.headers, ...sheet.rows.map((r) => r.cells)];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    XLSX.utils.book_append_sheet(wb, ws, sheet.name);
  }
  const base = fileName.replace(/\.[^/.]+$/, "");
  XLSX.writeFile(wb, `${base}_processed.xlsx`);
}
