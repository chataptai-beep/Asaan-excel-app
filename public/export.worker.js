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
function applyCapStates(strings) {
  const changed = new Map(); // idx -> newStr
  for (let i = 0; i < strings.length; i++) {
    let s = strings[i];
    if (!s) continue;
    let ns = s;
    for (const code of STATES) {
      const find = ` ${code} `;
      if (ns.toLowerCase().includes(find.toLowerCase()) && !ns.includes(find)) {
        ns = ns.replace(new RegExp(find.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'), 'gi'), find);
      }
    }
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
    const rOpen = sd.indexOf('<row', p);
    if (rOpen === -1) { parts.push(sd.slice(p)); break; }
    parts.push(sd.slice(p, rOpen));

    const rClose = sd.indexOf('</row>', rOpen);
    if (rClose === -1) { parts.push(sd.slice(rOpen)); p = sd.length; break; }
    const rEnd = rClose + 6;
    const rowXml = sd.slice(rOpen, rEnd);
    p = rEnd;

    const rnM = rowXml.match(/\br="(\d+)"/);
    const rowNum = rnM ? parseInt(rnM[1]) : 0;
    if (rowNum < DATA_START) { parts.push(rowXml); continue; }

    // Split row: header tag + content + </row>
    const hdrEnd = rowXml.indexOf('>') + 1;
    const rowHdr = rowXml.slice(0, hdrEnd);
    const rowBody = rowXml.slice(hdrEnd, rowXml.length - 6);

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
      if (PROPER_COLS.has(col)) nv = xlProper(xlClean(nv));
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
  const { buffer } = e.data;
  try {
    self.postMessage({ type:'progress', msg:'Loading file…' });
    const zip = await JSZip.loadAsync(buffer);

    // ── Shared strings ──────────────────────────────────────
    const ssFile = zip.file('xl/sharedStrings.xml');
    if (!ssFile) { self.postMessage({ type:'error', message:'No shared strings in this file' }); return; }

    self.postMessage({ type:'progress', msg:'Applying capital states…' });
    let ssXml = await ssFile.async('string');
    const ss = parseSharedStrings(ssXml);
    const ssChanged = applyCapStates(ss);
    for (const [i, v] of ssChanged) ss[i] = v;
    if (ssChanged.size > 0) zip.file('xl/sharedStrings.xml', rebuildSSXml(ssXml, ssChanged));

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
      const doP = NO_RI_NH.includes(state);
      zip.file(path, processSheet(xml, ss, doP, doP));
    }

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
