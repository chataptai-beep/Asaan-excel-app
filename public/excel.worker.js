/* Excel parsing Web Worker — runs off the main thread */
importScripts('/xlsx.min.js', '/jszip.min.js');

const KNOWN_STATES = ["CA","MA","NJ","IL","CT","FL","TX","NY","MD","RI","NH"];
const DIVIDER_DATE_RE = /^\s*\d{1,2}\.\d{1,2}\.\d{2,4}\s*$/;

function isDividerRow(cells) {
  let n = 0;
  for (const v of cells)
    if (v !== null && v !== undefined && DIVIDER_DATE_RE.test(String(v))) n++;
  return n >= 3;
}

function isRedRgb(hex) {
  if (!hex || hex.length < 6) return false;
  const h = hex.length === 8 ? hex.slice(2) : hex;  // strip alpha
  const r = parseInt(h.slice(0,2),16);
  const g = parseInt(h.slice(2,4),16);
  const b = parseInt(h.slice(4,6),16);
  return r >= 200 && g < 50 && b < 50;
}

// ---- Color detection via JSZip raw XML ----
function parseFills(xml) {
  const fills = [];
  const block = xml.match(/<fills[^>]*>([\s\S]*?)<\/fills>/);
  if (!block) return fills;
  for (const fe of (block[1].match(/<fill>([\s\S]*?)<\/fill>/g) || [])) {
    const rM = fe.match(/<fgColor[^>]+rgb="([0-9A-Fa-f]{6,8})"/);
    if (rM) { fills.push({ type:'rgb', hex: rM[1].toUpperCase() }); continue; }
    const tM = fe.match(/<fgColor[^>]+theme="(\d+)"(?:[^>]+tint="([^"]+)")?/);
    if (tM) { fills.push({ type:'theme', theme:+tM[1], tint: tM[2]?+tM[2]:0 }); continue; }
    fills.push(null);
  }
  return fills;
}

function parseCellXfs(xml) {
  const xfs = [];
  const block = xml.match(/<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/);
  if (!block) return xfs;
  for (const xf of (block[1].match(/<xf\b[^>]*\/?>/g) || [])) {
    const m = xf.match(/fillId="(\d+)"/);
    xfs.push(m ? +m[1] : 0);
  }
  return xfs;
}

function detectRedInXml(xml, fills, cellXfs) {
  const red = new Set();
  const fillFor = s => {
    const fid = s < cellXfs.length ? cellXfs[s] : 0;
    return fid < fills.length ? fills[fid] : null;
  };
  const isRed = f => f && f.type === 'rgb' && isRedRgb(f.hex);

  // Row-level
  for (const re of [
    /<row r="(\d+)"[^>]*\bs="(\d+)"[^>]*\bcustomFormat="1"/g,
    /<row r="(\d+)"[^>]*\bcustomFormat="1"[^>]*\bs="(\d+)"/g,
  ]) {
    let m;
    while ((m = re.exec(xml))) { const r=+m[1],s=+m[2]; if (!red.has(r)&&isRed(fillFor(s))) red.add(r); }
  }
  // Per-cell — only flag the row when column G (the Rep column) is red.
  // Other columns (e.g. Address in column F) sometimes carry red fills for
  // different reasons and should not trigger a row sort.
  let cm;
  const cr = /<c r="([A-Z]+)(\d+)"[^>]*\bs="(\d+)"/g;
  while ((cm = cr.exec(xml))) {
    // Convert column letters to 1-based index (A=1, G=7, …)
    let colIdx = 0;
    for (const ch of cm[1]) colIdx = colIdx * 26 + ch.charCodeAt(0) - 64;
    if (colIdx !== 7) continue; // only column G
    const r=+cm[2], s=+cm[3];
    if (r>=2 && !red.has(r) && isRed(fillFor(s))) red.add(r);
  }
  return red;
}

// ── Conditional-formatting red detection ──────────────────────────────────
// Red cells produced by CF rules have no `s` attribute → detectRedInXml misses them.
// We parse the CF rules from the sheet XML and evaluate them against the SheetJS
// cell values (already in memory), so no second XML parse of cell values is needed.

// Indexed-color slots that map to red/dark-red in the standard Excel palette.
const RED_INDEXED = new Set([2, 8, 9]);

function parseDxfRedIds(styleXml) {
  const out = new Set();
  const block = styleXml.match(/<dxfs[^>]*>([\s\S]*?)<\/dxfs>/);
  if (!block) return out;
  let i = 0;
  for (const m of (block[1].match(/<dxf>([\s\S]*?)<\/dxf>/g) || [])) {
    const f = m.match(/<fill>([\s\S]*?)<\/fill>/);
    if (f) {
      const rgb = f[1].match(/<(?:fg|bg)Color[^>]+rgb="([0-9A-Fa-f]{6,8})"/);
      const idx = f[1].match(/<(?:fg|bg)Color[^>]+indexed="(\d+)"/);
      if ((rgb && isRedRgb(rgb[1])) || (idx && RED_INDEXED.has(+idx[1]))) out.add(i);
    }
    i++;
  }
  return out;
}

// Convert a column letter like "G" to a 0-based column index (6).
function colLetterToIdx(letter) {
  return [...letter.toUpperCase()].reduce((a, c) => a * 26 + c.charCodeAt(0) - 64, 0) - 1;
}

// Parse condition objects from a single CF formula string.
// Appends to the `out` array and handles: equality, SEARCH (contains), OR(...)
function extractFormulaConds(formula, out) {
  // OR(...) — recurse on each arm
  const orM = formula.match(/^OR\(([\s\S]+)\)$/i);
  if (orM) {
    let depth = 0, inQ = false, start = 0;
    for (let i = 0; i < orM[1].length; i++) {
      const c = orM[1][i];
      if (c === '"') { inQ = !inQ; continue; }
      if (inQ) continue;
      if (c === '(') depth++;
      else if (c === ')') depth--;
      else if (c === ',' && depth === 0) {
        extractFormulaConds(orM[1].slice(start, i).trim(), out);
        start = i + 1;
      }
    }
    extractFormulaConds(orM[1].slice(start).trim(), out);
    return;
  }
  // SEARCH("text",$G1) — contains text
  const srM = formula.match(/SEARCH\(\s*"([^"]+)",\s*\$?([A-Za-z]+)\d+\s*\)/i);
  if (srM) { out.push({ col: colLetterToIdx(srM[2]), val: srM[1].toLowerCase(), contains: true }); return; }
  // $G1="Value" or G1="Value" — equality
  const eqM = formula.match(/\$?([A-Za-z]+)\d+\s*=\s*"([^"]*)"/);
  if (eqM) { out.push({ col: colLetterToIdx(eqM[1]), val: eqM[2].toLowerCase(), contains: false }); }
}

// Detect which rows (1-based) match a CF rule that applies a red DXF fill.
// Uses already-read SheetJS worksheet `ws` for cell values.
function detectCFRedRowsJS(ws, sheetXml, dxfRedIds) {
  if (!dxfRedIds.size || !ws || !ws['!ref']) return new Set();

  const conds = [];
  const blockRe = /<conditionalFormatting\b([^>]*)>([\s\S]*?)<\/conditionalFormatting>/g;
  let bM;
  while ((bM = blockRe.exec(sheetXml)) !== null) {
    const sqM = bM[1].match(/\bsqref="([^"]+)"/);
    const sqLetter = sqM ? (sqM[1].match(/^([A-Za-z]+)/) || [])[1] : null;
    const sqCol = sqLetter ? colLetterToIdx(sqLetter) : null;

    const ruleRe = /<cfRule\b([^>]*)>([\s\S]*?)<\/cfRule>/g;
    let rM;
    while ((rM = ruleRe.exec(bM[2])) !== null) {
      const dM = rM[1].match(/\bdxfId="(\d+)"/);
      if (!dM || !dxfRedIds.has(+dM[1])) continue;
      const fM = rM[2].match(/<formula[^>]*>([\s\S]*?)<\/formula>/);
      if (!fM) continue;
      const formula = fM[1].replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').trim();
      const cfType = (rM[1].match(/\btype="([^"]+)"/i) || [])[1]?.toLowerCase();

      if (cfType === 'cellis' && sqCol !== null) {
        const op = (rM[1].match(/\boperator="([^"]+)"/i) || [])[1]?.toLowerCase();
        if (op === 'equal') {
          const v = (formula.match(/^"([^"]*)"$/) || [])[1];
          if (v !== undefined) conds.push({ col: sqCol, val: v.toLowerCase(), contains: false });
        }
      } else {
        extractFormulaConds(formula, conds);
      }
    }
  }

  if (!conds.length) return new Set();

  const red = new Set();
  const range = XLSX.utils.decode_range(ws['!ref']);
  for (let r = 1; r <= range.e.r; r++) {       // r is 0-based; skip row 0 (header)
    for (const { col, val, contains } of conds) {
      const cell = ws[XLSX.utils.encode_cell({ r, c: col })];
      if (!cell) continue;
      const v = String(cell.v ?? '').toLowerCase();
      if (contains ? v.includes(val) : v === val) { red.add(r + 1); break; } // 1-based
    }
  }
  return red;
}

function parseSheetPaths(wbXml, relsXml) {
  const rel = {};
  for (const m of relsXml.matchAll(/<Relationship [^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g))
    rel[m[1]] = m[2];
  const out = {};
  for (const m of wbXml.matchAll(/<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"/g)) {
    const t = (rel[m[2]] || '').replace(/^\//,'');
    out[m[1]] = t.startsWith('xl/') ? t : 'xl/'+t;
  }
  return out;
}

self.onmessage = async (e) => {
  const { buffer } = e.data;
  try {
    self.postMessage({ type:'progress', msg:'Decompressing…' });
    const u8 = new Uint8Array(buffer);

    // 1. Parse values with SheetJS
    self.postMessage({ type:'progress', msg:'Reading cell values…' });
    const wb = XLSX.read(u8, { type:'array', cellDates:false, cellText:false });

    // 2. Detect red rows with JSZip
    self.postMessage({ type:'progress', msg:'Detecting highlighted rows…' });
    const zip = await JSZip.loadAsync(buffer);
    const stylesXml = await zip.file('xl/styles.xml').async('string');
    const fills = parseFills(stylesXml);
    const cellXfs = parseCellXfs(stylesXml);
    const dxfRedIds = parseDxfRedIds(stylesXml); // for CF detection
    const wbXml = await zip.file('xl/workbook.xml').async('string');
    const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
    const sheetPaths = parseSheetPaths(wbXml, relsXml);

    const redBySheet = {};
    for (const [name, path] of Object.entries(sheetPaths)) {
      const f = zip.file(path);
      if (!f) { redBySheet[name] = []; continue; }
      self.postMessage({ type:'progress', msg:`Scanning ${name}…` });
      const xml = await f.async('string');
      const directRed = detectRedInXml(xml, fills, cellXfs);
      const cfRed     = detectCFRedRowsJS(wb.Sheets[name], xml, dxfRedIds);
      redBySheet[name] = [...new Set([...directRed, ...cfRed])];
    }

    // 3. Build serializable result
    self.postMessage({ type:'progress', msg:'Building table data…' });
    const sheets = wb.SheetNames.map(name => {
      const ws = wb.Sheets[name];
      if (!ws || !ws['!ref']) return { name, headers:[], rows:[], isStateSheet: KNOWN_STATES.includes(name) };
      const range = XLSX.utils.decode_range(ws['!ref']);
      const maxRow = range.e.r;
      const maxCol = range.e.c;
      const redSet = new Set(redBySheet[name] || []);

      const headers = [];
      for (let c=0;c<=maxCol;c++) {
        const cell = ws[XLSX.utils.encode_cell({r:0,c})];
        headers.push(cell ? String(cell.v??'') : '');
      }

      const rows = [];
      for (let r=1;r<=maxRow;r++) {
        const cells = [];
        for (let c=0;c<=maxCol;c++) {
          const cell = ws[XLSX.utils.encode_cell({r,c})];
          if (!cell) { cells.push(null); continue; }
          if (cell.t==='n') cells.push(cell.v);
          else if (cell.t==='s') cells.push(cell.v);
          else if (cell.t==='b') cells.push(cell.v?'TRUE':'FALSE');
          else if (cell.v===undefined||cell.v===null) cells.push(null);
          else cells.push(String(cell.v));
        }
        const sourceRow = r + 1;
        const isDivider = isDividerRow(cells);
        const isRed = !isDivider && redSet.has(sourceRow);
        rows.push({ cells: cells ?? [], isRed, isDivider, originalRow: sourceRow });
      }

      return { name, headers, rows, isStateSheet: KNOWN_STATES.includes(name.trim().toUpperCase()) };
    });

    self.postMessage({ type:'done', sheets });
  } catch(err) {
    self.postMessage({ type:'error', message: String(err) });
  }
};
