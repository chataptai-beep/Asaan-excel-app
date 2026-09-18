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
  // Per-cell
  let cm;
  const cr = /<c r="[A-Z]+(\d+)"[^>]*\bs="(\d+)"/g;
  while ((cm = cr.exec(xml))) { const r=+cm[1],s=+cm[2]; if (r>=2&&!red.has(r)&&isRed(fillFor(s))) red.add(r); }
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
    const wbXml = await zip.file('xl/workbook.xml').async('string');
    const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
    const sheetPaths = parseSheetPaths(wbXml, relsXml);

    const redBySheet = {};
    for (const [name, path] of Object.entries(sheetPaths)) {
      const f = zip.file(path);
      if (!f) { redBySheet[name] = []; continue; }
      self.postMessage({ type:'progress', msg:`Scanning ${name}…` });
      const xml = await f.async('string');
      redBySheet[name] = [...detectRedInXml(xml, fills, cellXfs)];
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
