/**
 * Parse red-row information directly from xlsx XML (styles.xml + sheet XMLs).
 * Uses JSZip to decompress the xlsx (which is a ZIP archive).
 *
 * Detects:
 *   - Per-cell fills with a red-ish color (R ≥ 200, G < 50, B < 50)
 *   - Row-level fills (customFormat="1" + style index) with same condition
 */

type FillColor =
  | { type: "rgb"; r: number; g: number; b: number }
  | { type: "theme"; theme: number; tint: number }
  | null;

function isRed(fill: FillColor): boolean {
  if (!fill || fill.type !== "rgb") return false;
  return fill.r >= 200 && fill.g < 50 && fill.b < 50;
}

function parseRgb(hex8: string): { r: number; g: number; b: number } | null {
  const h = hex8.replace(/^#/, "").toUpperCase();
  if (h.length === 8) {
    // ARGB
    return {
      r: parseInt(h.slice(2, 4), 16),
      g: parseInt(h.slice(4, 6), 16),
      b: parseInt(h.slice(6, 8), 16),
    };
  }
  if (h.length === 6) {
    return {
      r: parseInt(h.slice(0, 2), 16),
      g: parseInt(h.slice(2, 4), 16),
      b: parseInt(h.slice(4, 6), 16),
    };
  }
  return null;
}

function parseFillsFromStyles(xml: string): FillColor[] {
  const fills: FillColor[] = [];
  const fillsBlock = xml.match(/<fills[^>]*>([\s\S]*?)<\/fills>/);
  if (!fillsBlock) return fills;

  const fillEntries = fillsBlock[1].match(/<fill>([\s\S]*?)<\/fill>/g) ?? [];
  for (const fe of fillEntries) {
    const rgbM = fe.match(/<fgColor[^>]+rgb="([0-9A-Fa-f]{6,8})"/);
    if (rgbM) {
      const c = parseRgb(rgbM[1]);
      if (c) { fills.push({ type: "rgb", ...c }); continue; }
    }
    const themeM = fe.match(/<fgColor[^>]+theme="(\d+)"(?:[^>]+tint="([^"]+)")?/);
    if (themeM) {
      fills.push({ type: "theme", theme: parseInt(themeM[1]), tint: themeM[2] ? parseFloat(themeM[2]) : 0 });
      continue;
    }
    fills.push(null);
  }
  return fills;
}

function parseCellXfsFromStyles(xml: string): number[] {
  const cellxfs: number[] = [];
  const block = xml.match(/<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/);
  if (!block) return cellxfs;
  const xfEntries = block[1].match(/<xf\b[^>]*\/?>/g) ?? [];
  for (const xf of xfEntries) {
    const m = xf.match(/fillId="(\d+)"/);
    cellxfs.push(m ? parseInt(m[1]) : 0);
  }
  return cellxfs;
}

function detectRedInSheetXml(
  xml: string,
  fills: FillColor[],
  cellXfs: number[]
): Set<number> {
  const redRows = new Set<number>();

  const fillFor = (sIdx: number): FillColor => {
    const fid = sIdx < cellXfs.length ? cellXfs[sIdx] : 0;
    return fid < fills.length ? fills[fid] : null;
  };

  // 1. Row-level fills: <row r="N" ... customFormat="1" ... s="S">
  const rowRe1 = /<row r="(\d+)"[^>]*\bs="(\d+)"[^>]*\bcustomFormat="1"/g;
  const rowRe2 = /<row r="(\d+)"[^>]*\bcustomFormat="1"[^>]*\bs="(\d+)"/g;
  for (const re of [rowRe1, rowRe2]) {
    let m: RegExpExecArray | null;
    while ((m = re.exec(xml)) !== null) {
      const r = parseInt(m[1]);
      const s = parseInt(m[2]);
      if (!redRows.has(r) && isRed(fillFor(s))) redRows.add(r);
    }
  }

  // 2. Per-cell fills: <c r="B5" s="N">
  const cellRe = /<c r="[A-Z]+(\d+)"[^>]*\bs="(\d+)"/g;
  let cm: RegExpExecArray | null;
  while ((cm = cellRe.exec(xml)) !== null) {
    const r = parseInt(cm[1]);
    const s = parseInt(cm[2]);
    if (r >= 2 && !redRows.has(r) && isRed(fillFor(s))) redRows.add(r);
  }

  return redRows;
}

function parseSheetPaths(wbXml: string, relsXml: string): Record<string, string> {
  const relTargets: Record<string, string> = {};
  for (const m of relsXml.matchAll(/<Relationship [^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g)) {
    relTargets[m[1]] = m[2];
  }
  const out: Record<string, string> = {};
  for (const m of wbXml.matchAll(/<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"/g)) {
    const tgt = relTargets[m[2]] ?? "";
    const cleaned = tgt.replace(/^\//, "");
    out[m[1]] = cleaned.startsWith("xl/") ? cleaned : "xl/" + cleaned;
  }
  return out;
}

export async function detectRedRowsBySheet(
  file: File
): Promise<Map<string, Set<number>>> {
  const JSZip = (await import("jszip")).default;
  const buf = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(buf);

  const stylesXml = await zip.file("xl/styles.xml")!.async("string");
  const fills = parseFillsFromStyles(stylesXml);
  const cellXfs = parseCellXfsFromStyles(stylesXml);

  const wbXml = await zip.file("xl/workbook.xml")!.async("string");
  const relsXml = await zip.file("xl/_rels/workbook.xml.rels")!.async("string");
  const sheetPaths = parseSheetPaths(wbXml, relsXml);

  const result = new Map<string, Set<number>>();
  for (const [sheetName, xmlPath] of Object.entries(sheetPaths)) {
    const f = zip.file(xmlPath);
    if (!f) { result.set(sheetName, new Set()); continue; }
    const xml = await f.async("string");
    result.set(sheetName, detectRedInSheetXml(xml, fills, cellXfs));
  }
  return result;
}
