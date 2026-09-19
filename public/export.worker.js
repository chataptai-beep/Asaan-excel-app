/* Export Web Worker — applies macros_toolkit run-all pipeline directly on XLSX XML.
   No Python, no server. Pipeline: clean-rep, clean-proper, capital-states, dedupe, format-phone.
   All formatting (fills, borders, fonts, red highlights) is preserved because we only
   change cell *values*, never touching styles.xml or conditional formatting. */

importScripts('/jszip.min.js');

const STATES   = ["CA","MA","NJ","IL","CT","FL","TX","NY","MD","RI","NH"];
const NO_RI_NH = ["CA","MA","NJ","IL","CT","FL","TX","NY","MD"];
// 1-based column indices
const CLEAN_COLS  = new Set([2,5,9,10]); // B E I J
const PROPER_COLS = new Set([4,6]);       // D F
const COL_REP   = 7;  // G
const COL_PHONE = 2;  // B
const DATA_START = 2; // row 1 = header

// ── Pure transforms ────────────────────────────────────────────────────────

function xlClean(v) {
  if (typeof v !== 'string') return v;
  return v.split('').filter(ch => ch.charCodeAt(0) >= 32).join('');
}
function xlProper(v) {
  if (typeof v !== 'string') return v;
  let r = '', prev = false;
  for (const ch of v) {
    if (/[a-zA-Z]/.test(ch)) { r += prev ? ch.toLowerCase() : ch.toUpperCase(); prev = true; }
    else { r += ch; prev = false; }
  }
  return r;
}
function fmtPhone(raw) {
  const s = typeof raw === 'number' ? String(raw) : String(raw || '').trim().replace(/,/g,'');
  if (!/^[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?$/.test(s)) return null;
  const n = Math.round(parseFloat(s));
  const digits = n === 0 ? '' : String(Math.abs(n));
  const pat = "(###) ###-####";
  const slots = [];
  for (let i = 0; i < pat.length; i++) if (pat[i] === '#') slots.push(i);
  const out = pat.split('');
  let p = digits.length;
  for (let k = slots.length-1; k >= 0; k--) {
    const idx = slots[k];
    if (k === 0) { out[idx] = digits.slice(0, p); }
    else { out[idx] = p > 0 ? digits[p-1] : ''; p = Math.max(0, p-1); }
  }
  return (n < 0 ? '-' : '') + out.join('');
}

// ── XML helpers ────────────────────────────────────────────────────────────

function escXml(s) {
  return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
function unescXml(s) {
  return s.replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&apos;/g,"'");
}
function colToIdx(col) {
  let idx = 0;
  for (const ch of col.toUpperCase()) idx = idx * 26 + ch.charCodeAt(0) - 64;
  return idx;
}
function refCol(ref) {
  const m = ref.match(/^([A-Za-z]+)/);
  return m ? colToIdx(m[1]) : 0;
}

// Locate the next <row> element at or after `from`.
// Empty styled rows are written self-closing (<row r="3730" ht="21"/>), which has
// no </row>; scanning for one would swallow the following row into the same chunk
// and renumber both under one r — Excel then sees duplicate/out-of-order rows and
// demands to repair the file.
function nextRow(xml, from) {
  let open = from;
  for (;;) {
    open = xml.indexOf('<row', open);
    if (open === -1) return null;
    const c = xml[open + 4];
    if (c === '>' || c === '/' || c === ' ' || c === '\t' || c === '\n' || c === '\r') break;
    open += 4; // skip <rowBreaks and friends
  }
  const tagEnd = xml.indexOf('>', open);
  if (tagEnd === -1) return null;
  if (xml[tagEnd - 1] === '/')
    return { open, end: tagEnd + 1, bodyStart: -1, bodyEnd: -1, selfClose: true };
  const close = xml.indexOf('</row>', tagEnd);
  if (close === -1) return null;
  return { open, end: close + 6, bodyStart: tagEnd + 1, bodyEnd: close, selfClose: false };
}

// ── Shared strings ─────────────────────────────────────────────────────────

function parseSharedStrings(xml) {
  const strings = [];
  let pos = 0;
  while (true) {
    const si = xml.indexOf('<si>', pos);
    if (si === -1) break;
    const siEnd = xml.indexOf('</si>', si);
    if (siEnd === -1) break;
    const body = xml.slice(si + 4, siEnd);
    let text = '', tp = 0;
    while (true) {
      const ts = body.indexOf('<t', tp);
      if (ts === -1) break;
      const tgt = body.indexOf('>', ts);
      if (tgt === -1) break;
      const te = body.indexOf('</t>', tgt);
      if (te === -1) break;
      text += unescXml(body.slice(tgt + 1, te));
      tp = te + 4;
    }
    strings.push(text);
    pos = siEnd + 5;
  }
  return strings;
}

// Apply capital-states to ALL shared strings (global text substitution, safe).
// Fix state abbreviations in a single string: " ca " → " CA ", etc.
function capStatesStr(s) {
  if (!s) return s;
  let ns = s;
  for (const code of STATES) {
    const find = ` ${code} `;
    if (ns.toLowerCase().includes(find.toLowerCase()) && !ns.includes(find)) {
      ns = ns.replace(new RegExp(find.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'), 'gi'), find);
    }
  }
  return ns;
}

function applyCapStates(strings) {
  const changed = new Map(); // idx -> newStr
  for (let i = 0; i < strings.length; i++) {
    const s = strings[i];
    if (!s) continue;
    const ns = capStatesStr(s);
    if (ns !== s) changed.set(i, ns);
  }
  return changed;
}

function rebuildSSXml(ssXml, changed) {
  if (changed.size === 0) return ssXml;
  let i = 0;
  return ssXml.replace(/<si>([\s\S]*?)<\/si>/g, (match, body) => {
    const idx = i++;
    const newStr = changed.get(idx);
    if (newStr === undefined) return match;
    if (body.includes('<r>') || body.match(/<r\s/)) return match; // skip rich text
    const preserve = body.includes('xml:space="preserve"');
    const tTag = preserve ? `<t xml:space="preserve">${escXml(newStr)}</t>` : `<t>${escXml(newStr)}</t>`;
    return `<si>${tTag}</si>`;
  });
}

// ── Conditional-formatting red detection ──────────────────────────────────
// Rows that appear red due to CF rules (not direct cell fill) also need moving.

function parseDxfRedIds(styleXml) {
  const redIds = new Set();
  const block = styleXml.match(/<dxfs[^>]*>([\s\S]*?)<\/dxfs>/);
  if (!block) return redIds;
  let i = 0, m;
  const re = /<dxf>([\s\S]*?)<\/dxf>/g;
  while ((m = re.exec(block[1])) !== null) {
    const fillM = m[1].match(/<fill>([\s\S]*?)<\/fill>/);
    if (fillM) {
      const fgM = fillM[1].match(/<fgColor[^>]+rgb="([0-9A-Fa-f]{6,8})"/);
      const iM  = fillM[1].match(/<fgColor[^>]+indexed="(\d+)"/);
      if (fgM && isRedRgb(fgM[1])) redIds.add(i);
      if (iM  && isRedRgb(INDEXED_COLORS[+iM[1]] || '')) redIds.add(i);
    }
    i++;
  }
  return redIds;
}

// Scan a sheet's <conditionalFormatting> rules that reference a red DXF.
// Handles the two common patterns:
//   equality  : $G1="Spring Saunders"
//   containsText: NOT(ISERROR(SEARCH("text",$G1)))
function detectCFRedRows(sheetXml, redDxfIds, ss) {
  if (redDxfIds.size === 0) return new Set();

  const conditions = [];
  const cfRe = /<cfRule\b([^>]*)>([\s\S]*?)<\/cfRule>/g;
  let cfM;
  while ((cfM = cfRe.exec(sheetXml)) !== null) {
    const dxfM = cfM[1].match(/\bdxfId="(\d+)"/);
    if (!dxfM || !redDxfIds.has(+dxfM[1])) continue;
    const fM = cfM[2].match(/<formula[^>]*>([\s\S]*?)<\/formula>/);
    if (!fM) continue;
    const formula = fM[1].trim();
    // Pattern A: $G1="Value" — exact match on a column
    const eqM = formula.match(/\$([A-Za-z]+)\d+\s*=\s*"([^"]*)"/);
    if (eqM) { conditions.push({ col: colToIdx(eqM[1]), value: eqM[2].toLowerCase(), contains: false }); continue; }
    // Pattern B: SEARCH("text",$G1) — contains-text rule
    const srM = formula.match(/SEARCH\(\s*"([^"]+)",\s*\$([A-Za-z]+)\d+\s*\)/i);
    if (srM) { conditions.push({ col: colToIdx(srM[2]), value: srM[1].toLowerCase(), contains: true }); }
  }
  if (conditions.length === 0) return new Set();

  const red = new Set();
  let p = 0;
  while (true) {
    const r = nextRow(sheetXml, p);
    if (!r) break;
    p = r.end;
    if (r.selfClose) continue; // no cells to test
    const rnM = sheetXml.slice(r.open, r.bodyStart).match(/\br="(\d+)"/);
    if (!rnM || +rnM[1] < DATA_START) continue;
    const rowNum = +rnM[1];
    const cells  = parseCells(sheetXml.slice(r.bodyStart, r.bodyEnd));
    for (const cond of conditions) {
      const cell = cells.find(c => c.col === cond.col);
      if (!cell) continue;
      const val = String(getCellValue(cell, ss) || '').toLowerCase();
      if (cond.contains ? val.includes(cond.value) : val === cond.value) {
        red.add(rowNum); break;
      }
    }
  }
  return red;
}

// ── Sheet processing ───────────────────────────────────────────────────────

function processSheet(xml, ss, doPhone, doDedupe) {
  const SD_O = '<sheetData>', SD_C = '</sheetData>';
  const sdS = xml.indexOf(SD_O), sdE = xml.lastIndexOf(SD_C);
  if (sdS === -1 || sdE === -1) return xml;

  const before = xml.slice(0, sdS + SD_O.length);
  const sd     = xml.slice(sdS + SD_O.length, sdE);
  const after  = xml.slice(sdE);

  const seen = new Set();
  const parts = [];
  let p = 0;

  while (p < sd.length) {
    const r = nextRow(sd, p);
    if (!r) { parts.push(sd.slice(p)); break; }
    parts.push(sd.slice(p, r.open));

    const rowXml = sd.slice(r.open, r.end);
    p = r.end;

    const rnM = rowXml.match(/\br="(\d+)"/);
    const rowNum = rnM ? parseInt(rnM[1]) : 0;
    // Self-closing rows hold no cells — nothing to clean, dedupe or reformat.
    if (r.selfClose || rowNum < DATA_START) { parts.push(rowXml); continue; }

    const rowHdr = sd.slice(r.open, r.bodyStart);
    const rowBody = sd.slice(r.bodyStart, r.bodyEnd);

    // Parse cells
    const cells = parseCells(rowBody);

    // Build value array for dedupe
    if (doDedupe) {
      const maxC = cells.reduce((m,c) => Math.max(m, c.col), 0);
      const vals = new Array(Math.max(maxC, 10)).fill(null);
      for (const c of cells) {
        if (c.col < 1) continue;
        vals[c.col - 1] = getCellValue(c, ss);
      }
      const blank = vals.every(v => v === null || v === '');
      if (!blank) {
        const key = JSON.stringify(vals.map(v =>
          (v===null||v==='')?null:(typeof v==='string'?v.toLowerCase():v)));
        if (seen.has(key)) continue; // skip duplicate row
        seen.add(key);
      }
    }

    // Transform cell values (clean-rep, clean-proper, format-phone)
    let newBody = rowBody;
    for (const cell of cells) {
      if (cell.selfClose || cell.isFormula) continue;
      const newVal = transformCell(cell, ss, doPhone);
      if (newVal === null) continue;
      const newCell = inlineCell(cell, newVal);
      newBody = newBody.split(cell.full).join(newCell);
    }

    parts.push(rowHdr + newBody + '</row>');
  }

  return before + parts.join('') + after;
}

function parseCells(rowBody) {
  const cells = [];
  let p = 0;
  while (p < rowBody.length) {
    const cs = rowBody.indexOf('<c', p);
    if (cs === -1) break;
    const tEnd = rowBody.indexOf('>', cs);
    if (tEnd === -1) break;
    const self = rowBody[tEnd - 1] === '/';
    let ce;
    if (self) { ce = tEnd + 1; }
    else {
      const ci = rowBody.indexOf('</c>', tEnd);
      if (ci === -1) break;
      ce = ci + 4;
    }
    const full = rowBody.slice(cs, ce);
    const attrStr = rowBody.slice(cs + 2, self ? tEnd - 1 : tEnd);
    const refM = attrStr.match(/\br="([^"]+)"/);
    const tM   = attrStr.match(/\bt="([^"]+)"/);
    const ref = refM ? refM[1] : '';
    const t   = tM   ? tM[1]  : '';
    const col = refCol(ref);
    const inner = self ? '' : rowBody.slice(tEnd + 1, ce - 4);
    const isFormula = inner.includes('<f');
    cells.push({ full, attrStr, ref, t, col, inner, selfClose: self, isFormula });
    p = ce;
  }
  return cells;
}

function getCellValue(cell, ss) {
  const vM = cell.inner.match(/<v[^>]*>([^<]*)<\/v>/);
  if (!vM) return null;
  if (cell.t === 's') {
    const idx = parseInt(vM[1]);
    return isNaN(idx) ? null : (ss[idx] ?? null);
  }
  return vM[1] || null;
}

function transformCell(cell, ss, doPhone) {
  const col = cell.col, t = cell.t;

  if (t === 's') {
    const vM = cell.inner.match(/<v[^>]*>(\d+)<\/v>/);
    if (!vM) return null;
    const idx = parseInt(vM[1]);
    if (isNaN(idx) || idx >= ss.length) return null;
    const old = ss[idx];

    if (col === COL_REP) {
      const nv = xlClean(old); return nv !== old ? nv : null;
    }
    if (CLEAN_COLS.has(col) || PROPER_COLS.has(col)) {
      let nv = old;
      if (CLEAN_COLS.has(col)) nv = xlClean(nv);
      if (PROPER_COLS.has(col)) {
        nv = xlProper(xlClean(nv));
        // xlProper lowercases 2nd+ letters of every word, turning "CA" → "Ca".
        // Re-apply state capitalization so the address keeps the correct abbreviation.
        nv = capStatesStr(nv);
      }
      return nv !== old ? nv : null;
    }
  }

  // format-phone: numeric col B
  if (doPhone && col === COL_PHONE && t !== 's' && t !== 'b' && t !== 'e') {
    const vM = cell.inner.match(/<v[^>]*>([^<]*)<\/v>/);
    if (vM && vM[1]) return fmtPhone(vM[1]);
  }

  return null;
}

function inlineCell(cell, newVal) {
  // Strip existing t="..." attr, add t="inlineStr"
  const attrs = cell.attrStr.replace(/\bt="[^"]*"\s*/g,'').trim();
  const ws = (newVal !== newVal.trim()) ? ' xml:space="preserve"' : '';
  return `<c ${attrs} t="inlineStr"><is><t${ws}>${escXml(newVal)}</t></is></c>`;
}

// ── Red-row detection (mirrors excel.worker.js logic exactly) ─────────────

function isRedRgb(hex) {
  if (!hex || hex.length < 6) return false;
  const h = hex.length === 8 ? hex.slice(2) : hex;
  const r = parseInt(h.slice(0,2),16);
  const g = parseInt(h.slice(2,4),16);
  const b = parseInt(h.slice(4,6),16);
  return r >= 200 && g < 50 && b < 50;
}

// Standard Excel indexed color palette — index 2 = Red, index 8 = Dark Red.
// Excel files created by older tools or macros often use indexed colors instead of RGB.
const INDEXED_COLORS = [
  '000000','FFFFFF','FF0000','00FF00','0000FF','FFFF00','FF00FF','00FFFF',
  '800000','008000','000080','808000','800080','008080','C0C0C0','808080',
  '9999FF','993366','FFFFCC','CCFFFF','660066','FF8080','0066CC','CCCCFF',
  '000080','FF00FF','FFFF00','00FFFF','800080','800000','008080','0000FF',
  '00CCFF','CCFFFF','CCFFCC','FFFF99','99CCFF','FF99CC','CC99FF','FFCC99',
  '3366FF','33CCCC','99CC00','FFCC00','FF9900','FF6600','666699','969696',
  '003366','339966','003300','333300','993300','993366','333399','333333',
];

function parseFills(xml) {
  const fills = [];
  const block = xml.match(/<fills[^>]*>([\s\S]*?)<\/fills>/);
  if (!block) return fills;
  for (const fe of (block[1].match(/<fill>([\s\S]*?)<\/fill>/g) || [])) {
    const rM = fe.match(/<fgColor[^>]+rgb="([0-9A-Fa-f]{6,8})"/);
    if (rM) { fills.push({ type:'rgb', hex: rM[1].toUpperCase() }); continue; }
    const iM = fe.match(/<fgColor[^>]+indexed="(\d+)"/);
    if (iM) {
      const hex = INDEXED_COLORS[+iM[1]];
      fills.push(hex ? { type:'rgb', hex: hex.toUpperCase() } : null);
      continue;
    }
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

function detectRedRows(sheetXml, fills, cellXfs) {
  const red = new Set();
  const fillFor = s => {
    const fid = s < cellXfs.length ? cellXfs[s] : 0;
    return fid < fills.length ? fills[fid] : null;
  };
  const isRed = f => f && f.type === 'rgb' && isRedRgb(f.hex);
  // Row-level custom format
  for (const re of [
    /<row r="(\d+)"[^>]*\bs="(\d+)"[^>]*\bcustomFormat="1"/g,
    /<row r="(\d+)"[^>]*\bcustomFormat="1"[^>]*\bs="(\d+)"/g,
  ]) {
    let m;
    while ((m = re.exec(sheetXml))) {
      const r=+m[1], s=+m[2];
      if (!red.has(r) && isRed(fillFor(s))) red.add(r);
    }
  }
  // Per-cell — only flag column G (Rep); other columns (e.g. red Address) must not sort the row.
  const cr = /<c r="([A-Z]+)(\d+)"[^>]*\bs="(\d+)"/g;
  let cm;
  while ((cm = cr.exec(sheetXml))) {
    let colIdx = 0;
    for (const ch of cm[1]) colIdx = colIdx * 26 + ch.charCodeAt(0) - 64;
    if (colIdx !== 7) continue;
    const r=+cm[2], s=+cm[3];
    if (r >= 2 && !red.has(r) && isRed(fillFor(s))) red.add(r);
  }
  return red; // set of 1-based row numbers
}

// Renumber a single row's r attributes (row tag + all cell refs) from oldNum → newNum.
function renumberRow(rowXml, oldNum, newNum) {
  if (oldNum === newNum) return rowXml;
  const o = String(oldNum), n = String(newNum);
  // Update <row r="N"> (first plain-number r= in the tag, never has letters)
  let xml = rowXml.replace(`r="${o}"`, `r="${n}"`);
  // Update every <c r="XN"> cell reference (column letters + old row number)
  xml = xml.replace(new RegExp(`r="([A-Z]+)${o}"`, 'g'), `r="$1${n}"`);
  // An array formula carries its own range: <f t="array" ref="D29083">. Left at the
  // old row it no longer covers its host cell, which Excel reports as damage.
  xml = xml.replace(/\bref="([^"]+)"/g, (m, ref) =>
    `ref="${ref.replace(new RegExp(`([A-Z]+)${o}(?![0-9])`, 'g'), `$1${n}`)}"`);
  return xml;
}

// Move red rows to the top of sheetData (after header), renumbering rows so
// Excel respects the new order (the r attribute controls physical placement).
function moveRedToTop(sheetXml, redRows) {
  if (redRows.size === 0) return sheetXml;

  const SD_O = '<sheetData>', SD_C = '</sheetData>';
  const sdS = sheetXml.indexOf(SD_O), sdE = sheetXml.lastIndexOf(SD_C);
  if (sdS === -1 || sdE === -1) return sheetXml;

  const before = sheetXml.slice(0, sdS + SD_O.length);
  const sd     = sheetXml.slice(sdS + SD_O.length, sdE);
  const after  = sheetXml.slice(sdE);

  const headers = [], reds = [], normals = [];
  let p = 0;
  while (p < sd.length) {
    const r = nextRow(sd, p);
    if (!r) break;
    const rowXml = sd.slice(r.open, r.end);
    p = r.end;
    const rnM = rowXml.match(/\br="(\d+)"/);
    const rowNum = rnM ? +rnM[1] : 0;
    if (rowNum < DATA_START) headers.push([rowXml, rowNum]);
    else if (redRows.has(rowNum)) reds.push([rowXml, rowNum]);
    else normals.push([rowXml, rowNum]);
  }

  // Renumber every row in its new position so Excel places it correctly.
  const out = [];
  let newNum = 1;
  for (const [xml, oldNum] of [...headers, ...reds, ...normals]) {
    out.push(renumberRow(xml, oldNum, newNum++));
  }
  const lastRow = newNum - 1;

  // Update <dimension ref="A1:XN"/> to reflect the new last row so Excel does
  // not flag a stale range reference on open.
  const updatedBefore = before.replace(/<dimension ref="([^"]*)"/g, (match, ref) => {
    const colM = ref.match(/:([A-Z]+)\d+/);
    const lastCol = colM ? colM[1] : 'Z';
    return `<dimension ref="A1:${lastCol}${lastRow}"`;
  });

  return updatedBefore + out.join('') + after;
}

// ── Post-process sheet XML to fix stale references ─────────────────────────
// Called after processSheet + moveRedToTop. Handles every element in the
// sheet XML whose row/range reference can become stale after dedup/renumber.
async function fixSheetMeta(zip, sheetPath, xml) {
  // Find the actual last row by scanning sheetData.
  // moveRedToTop already updates <dimension> when red rows exist, but when it
  // returns early (no red rows) the dimension is stale after dedup.
  const SD_C = '</sheetData>';
  const sdE = xml.lastIndexOf(SD_C);
  let lastRow = 0, lastCol = 'A';

  if (sdE !== -1) {
    const sd = xml.slice(0, sdE);
    let m;
    const re = /<row\b[^>]*\br="(\d+)"/g;
    while ((m = re.exec(sd)) !== null) {
      const n = +m[1];
      if (n > lastRow) lastRow = n;
    }
  }

  // Read the last column from the existing (possibly stale) dimension element.
  const dimM = xml.match(/<dimension ref="[^"]*:([A-Z]+)\d+"/);
  if (dimM) lastCol = dimM[1];

  if (lastRow === 0) return xml;

  let result = xml
    // 1. Fix <dimension ref="..."> (handles the no-red-rows case)
    .replace(/<dimension ref="([^"]*)"/g, `<dimension ref="A1:${lastCol}${lastRow}"`)
    // 2. Fix <autoFilter ref="..."> — stale after dedup/renumber
    .replace(/(<autoFilter\b[^>]*)\bref="[^"]*"/g, `$1ref="A1:${lastCol}${lastRow}"`)
    // 3. Strip page breaks — row numbers are invalid after renumber
    .replace(/<rowBreaks\b[^>]*>[\s\S]*?<\/rowBreaks>/g, '')
    .replace(/<colBreaks\b[^>]*>[\s\S]*?<\/colBreaks>/g, '');

  // 4. Fix Excel table refs (xl/tables/tableN.xml) if this sheet has any.
  // A table whose ref="A1:J83516" is wrong after renumber → Excel flags it.
  const tableIds = [];
  for (const m of result.matchAll(/<tablePart\s[^>]*r:id="([^"]+)"/g)) tableIds.push(m[1]);

  if (tableIds.length > 0) {
    // Derive the _rels path for this sheet.
    const parts = sheetPath.split('/');
    const relsPath = parts.slice(0, -1).join('/') + '/_rels/' + parts[parts.length - 1] + '.rels';
    const relsF = zip.file(relsPath);
    if (relsF) {
      const relsContent = await relsF.async('string');
      for (const id of tableIds) {
        const relM = relsContent.match(new RegExp(`Id="${id}"[^>]*Target="([^"]+)"`));
        if (!relM) continue;
        const target = relM[1]; // e.g., "../tables/table1.xml"
        const base = sheetPath.slice(0, sheetPath.lastIndexOf('/') + 1);
        const tablePath = (target.startsWith('../'))
          ? base.replace(/[^/]+\/$/, '') + target.slice(3)
          : base + target;
        const tableF = zip.file(tablePath);
        if (!tableF) continue;
        let tableXml = await tableF.async('string');
        // Update the table's ref and its autoFilter ref to the new range.
        tableXml = tableXml
          .replace(/(<table\b[^>]*)\bref="[^"]*"/, `$1ref="A1:${lastCol}${lastRow}"`)
          .replace(/(<autoFilter\b[^>]*)\bref="[^"]*"/, `$1ref="A1:${lastCol}${lastRow}"`);
        zip.file(tablePath, tableXml);
      }
    }
  }

  return result;
}

// ── Sheet path lookup ──────────────────────────────────────────────────────

function sheetPaths(wbXml, relsXml) {
  const rel = {};
  for (const m of relsXml.matchAll(/Id="([^"]+)"[^>]*Target="([^"]+)"/g))
    rel[m[1]] = m[2];
  const out = {};
  for (const m of wbXml.matchAll(/<sheet[^>]+name="([^"]+)"[^>]+r:id="([^"]+)"/g)) {
    const t = (rel[m[2]]||'').replace(/^\//,'');
    out[m[1]] = t.startsWith('xl/') ? t : 'xl/' + t;
  }
  return out;
}

// ── Main ───────────────────────────────────────────────────────────────────

self.onmessage = async (e) => {
  const { buffer, redBySheet = {} } = e.data;
  try {
    self.postMessage({ type:'progress', msg:'Loading file…' });
    const zip = await JSZip.loadAsync(buffer);

    // Remove calc chain — its cell references become invalid after row renumbering.
    // Excel recalculates on open when this file is absent. Must also remove it
    // from [Content_Types].xml and xl/_rels/workbook.xml.rels so Excel does not
    // flag a broken relationship pointing to a missing part.
    zip.remove('xl/calcChain.xml');
    const ctFile = zip.file('[Content_Types].xml');
    if (ctFile) {
      let ct = await ctFile.async('string');
      ct = ct.replace(/<Override[^>]*calcChain[^>]*\/>/g, '');
      zip.file('[Content_Types].xml', ct);
    }
    const wbRelsFile = zip.file('xl/_rels/workbook.xml.rels');
    if (wbRelsFile) {
      let rels = await wbRelsFile.async('string');
      rels = rels.replace(/<Relationship[^>]*calcChain[^>]*\/>/g, '');
      zip.file('xl/_rels/workbook.xml.rels', rels);
    }

    // ── Shared strings ──────────────────────────────────────
    const ssFile = zip.file('xl/sharedStrings.xml');
    if (!ssFile) { self.postMessage({ type:'error', message:'No shared strings in this file' }); return; }

    self.postMessage({ type:'progress', msg:'Applying capital states…' });
    let ssXml = await ssFile.async('string');
    const ss = parseSharedStrings(ssXml);
    const ssChanged = applyCapStates(ss);
    for (const [i, v] of ssChanged) ss[i] = v;
    if (ssChanged.size > 0) zip.file('xl/sharedStrings.xml', rebuildSSXml(ssXml, ssChanged));

    // ── Styles (for red-row detection) ─────────────────────
    const stylesXml = await zip.file('xl/styles.xml').async('string');
    const fills     = parseFills(stylesXml);
    const cellXfs   = parseCellXfs(stylesXml);
    const redDxfIds = parseDxfRedIds(stylesXml); // for conditional-formatting detection

    // ── Sheet paths ─────────────────────────────────────────
    const wbXml   = await zip.file('xl/workbook.xml').async('string');
    const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
    const paths   = sheetPaths(wbXml, relsXml);

    // ── Process each state sheet ────────────────────────────
    for (const state of STATES) {
      const path = paths[state];
      if (!path) continue;
      const f = zip.file(path);
      if (!f) continue;
      self.postMessage({ type:'progress', msg:`Processing ${state}…` });
      const xml = await f.async('string');
      // Prefer the red-row data already detected by excel.worker.js (the display
      // worker); it is more reliable because it runs before any modifications and
      // handles every fill / CF pattern the file actually uses.  Fall back to
      // local detection only when the caller did not supply pre-detected data.
      const passed = redBySheet[state];
      const redRows = (passed && passed.length > 0)
        ? new Set(passed)
        : new Set([...detectRedRows(xml, fills, cellXfs), ...detectCFRedRows(xml, redDxfIds, ss)]);
      const doP = NO_RI_NH.includes(state);
      // Clean/dedupe/phone transforms, then move red rows to top with renumbering.
      let out = processSheet(xml, ss, doP, doP);
      out = moveRedToTop(out, redRows);
      // Fix every stale range reference created by the above transforms.
      out = await fixSheetMeta(zip, path, out);
      zip.file(path, out);
    }

    // ── Strip stale named ranges from workbook.xml ──────────
    // Print areas and other defined names reference row numbers that are no
    // longer valid after dedup + renumber. Excel flags them as corrupt.
    const wbXmlClean = wbXml.replace(/<definedNames\b[^>]*>[\s\S]*?<\/definedNames>/g, '<definedNames/>');
    if (wbXmlClean !== wbXml) zip.file('xl/workbook.xml', wbXmlClean);

    // ── Generate output ─────────────────────────────────────
    self.postMessage({ type:'progress', msg:'Building output file…' });
    const out = await zip.generateAsync({
      type: 'arraybuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
    });

    self.postMessage({ type:'done', buffer: out }, [out]);
  } catch (err) {
    self.postMessage({ type:'error', message: String(err) });
  }
};
