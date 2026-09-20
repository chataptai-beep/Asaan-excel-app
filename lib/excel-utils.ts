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

// NOTE: exporting is done by public/export.worker.js, which edits the workbook
// XML in place so every fill / font / border / conditional-format survives (the
// same guarantee as macros_toolkit.py). A values-only aoa_to_sheet export used
// to live here; it was removed because it silently discarded all formatting.
