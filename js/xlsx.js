// Fichiers Excel.
// - Modèle du client : on modifie directement le XML du .xlsx pour ne toucher qu'aux cellules
//   choisies ; mise en forme, logos, formules et validations restent intacts.
// - Tableau standard : classeur sobre généré de zéro.

import { getJSZip, loadScript, CDN } from './lib.js';
import { fieldDef, lineValue, singleValue, lineColDef, EUR_EQUIVALENT, LINE_FIELDS, DEFAULT_SIMPLE_COLUMNS } from './fields.js';
import { parseISO, blobToBase64 } from './ui.js';

export const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_XML = 'http://www.w3.org/XML/1998/namespace';
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

/* ---------- Références de cellules ---------- */

export const colToNum = s => [...s.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
export function numToCol(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
export function splitRef(ref) {
  const m = /^\$?([A-Z]{1,3})\$?(\d+)$/i.exec(ref || '');
  return m ? { col: m[1].toUpperCase(), c: colToNum(m[1]), r: +m[2] } : null;
}

/* ---------- Petits outils XML ---------- */

const parseXml = s => new DOMParser().parseFromString(s, 'application/xml');
const serialize = doc => {
  const s = new XMLSerializer().serializeToString(doc);
  return s.startsWith('<?xml') ? s : XML_DECL + s;
};
const kids = (el, name) => [...el.children].filter(c => c.localName === name);
const kid = (el, name) => [...el.children].find(c => c.localName === name) || null;
const all = (root, name) => [...root.getElementsByTagNameNS('*', name)];

function textOf(el) {
  let s = '';
  for (const t of all(el, 't')) if (t.parentElement?.localName !== 'rPh') s += t.textContent;
  return s;
}

function resolvePart(target, base = 'xl') {
  if (target.startsWith('/')) return target.slice(1);
  const out = [];
  for (const p of (`${base}/${target}`).split('/')) {
    if (p === '..') out.pop(); else if (p && p !== '.') out.push(p);
  }
  return out.join('/');
}

const BUILTIN_FMTS = { 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00', 14: 'dd/mm/yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 22: 'dd/mm/yyyy hh:mm', 49: '@' };
function formatCode(styles, s) {
  if (!styles) return '';
  const id = styles.xfs[s] ?? 0;
  return styles.numFmts[id] || BUILTIN_FMTS[id] || '';
}

const escXml = s => String(s).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '')
  .replace(/[&<>"]/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[ch]);

/* ---------- Dates Excel ---------- */

const EPOCH = Date.UTC(1899, 11, 30);
export function isoToSerial(iso) {
  const d = parseISO(iso);
  return d ? Math.round((Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) - EPOCH) / 86400000) : null;
}
function serialToDisplay(n) {
  const d = new Date(EPOCH + Math.round(n) * 86400000);
  return d.toLocaleDateString('fr-BE', { timeZone: 'UTC' });
}

const BUILTIN_DATE_FMTS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58]);
function isDateFormat(id, numFmts) {
  if (BUILTIN_DATE_FMTS.has(id)) return true;
  const code = numFmts[id];
  if (!code) return false;
  const stripped = code.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '').replace(/\\./g, '');
  return /[dmy]/i.test(stripped) && !/^general$/i.test(stripped.trim());
}

/* ---------- Lecture d'un classeur ---------- */

export async function openWorkbook(data) {
  const JSZip = await getJSZip();
  let zip;
  try {
    zip = await JSZip.loadAsync(data, typeof data === 'string' ? { base64: true } : undefined);
  } catch {
    throw new Error("Ce fichier n'est pas un classeur Excel .xlsx valide. Les anciens fichiers .xls doivent d'abord être enregistrés au format .xlsx.");
  }
  const wbFile = zip.file('xl/workbook.xml');
  if (!wbFile) throw new Error("Ce fichier n'est pas un classeur Excel .xlsx.");
  const wb = parseXml(await wbFile.async('string'));
  const relsDoc = parseXml(await zip.file('xl/_rels/workbook.xml.rels').async('string'));
  const rels = {};
  for (const r of all(relsDoc, 'Relationship')) rels[r.getAttribute('Id')] = r.getAttribute('Target');
  const sheets = all(wb, 'sheet').map(s => {
    const rid = s.getAttributeNS(NS_REL, 'id') || s.getAttribute('r:id');
    return { name: s.getAttribute('name'), path: rels[rid] ? resolvePart(rels[rid]) : '', hidden: !!s.getAttribute('state') && s.getAttribute('state') !== 'visible' };
  }).filter(s => s.path && zip.file(s.path));
  if (!sheets.length) throw new Error('Aucune feuille de calcul trouvée dans ce classeur.');

  const sstFile = zip.file('xl/sharedStrings.xml');
  const shared = sstFile ? all(parseXml(await sstFile.async('string')), 'si').map(textOf) : [];
  const styles = await readStyles(zip);
  return { zip, wb, sheets, shared, styles };
}

async function readStyles(zip) {
  const file = zip.file('xl/styles.xml');
  if (!file) return null;
  const doc = parseXml(await file.async('string'));
  const numFmts = {};
  for (const nf of all(doc, 'numFmt')) numFmts[+nf.getAttribute('numFmtId')] = nf.getAttribute('formatCode');
  const cellXfs = kid(doc.documentElement, 'cellXfs');
  const xfs = cellXfs ? kids(cellXfs, 'xf').map(x => +x.getAttribute('numFmtId') || 0) : [];
  return { doc, numFmts, cellXfs, xfs, dirty: false, cache: new Map() };
}

// Contenu d'une feuille, pour l'aperçu et la détection de la taille du tableau.
export async function readSheet(book, path) {
  const doc = parseXml(await book.zip.file(path).async('string'));
  const cells = new Map();
  const sharedFormulas = {};
  let maxRow = 0;
  let maxCol = 0;
  for (const row of all(doc, 'row')) {
    for (const c of kids(row, 'c')) {
      const p = splitRef(c.getAttribute('r'));
      if (!p) continue;
      const t = c.getAttribute('t');
      const s = +c.getAttribute('s') || 0;
      const f = kid(c, 'f');
      const v = kid(c, 'v');
      let val = null;
      if (t === 's' && v) val = book.shared[+v.textContent] ?? '';
      else if (t === 'inlineStr') val = textOf(kid(c, 'is') || c);
      else if (t === 'b' && v) val = v.textContent === '1' ? 'VRAI' : 'FAUX';
      else if (v) val = t === 'str' || t === 'e' ? v.textContent : Number(v.textContent);
      let formula = null;
      if (f) {
        formula = f.textContent || '';
        if (f.getAttribute('t') === 'shared') {
          const si = f.getAttribute('si');
          if (formula) sharedFormulas[si] = formula; else formula = sharedFormulas[si] || '=';
        }
      }
      const isDate = typeof val === 'number' && book.styles && isDateFormat(book.styles.xfs[s] || 0, book.styles.numFmts);
      const display = val == null ? '' : typeof val === 'number' ? (isDate ? serialToDisplay(val) : val.toLocaleString('fr-BE', { maximumFractionDigits: 2 })) : String(val);
      cells.set(`${p.col}${p.r}`, { ref: `${p.col}${p.r}`, r: p.r, c: p.c, val, display, formula, isDate, fmt: formatCode(book.styles, s) });
      if (display !== '' || formula) { maxRow = Math.max(maxRow, p.r); maxCol = Math.max(maxCol, p.c); }
    }
  }
  const merges = all(doc, 'mergeCell').map(m => {
    const [a, b] = (m.getAttribute('ref') || '').split(':');
    const pa = splitRef(a);
    const pb = splitRef(b || a);
    return pa && pb ? { r1: pa.r, c1: pa.c, r2: pb.r, c2: pb.c } : null;
  }).filter(Boolean);
  for (const m of merges) { maxRow = Math.max(maxRow, m.r2); maxCol = Math.max(maxCol, m.c2); }
  const cols = all(doc, 'col').map(c => ({ min: +c.getAttribute('min'), max: +c.getAttribute('max'), width: +c.getAttribute('width') || 0, hidden: c.getAttribute('hidden') === '1' }));
  return { cells, merges, cols, maxRow, maxCol };
}

/* ---------- Taille du tableau des dépenses ---------- */

// Nombre de lignes disponibles avant la ligne « Total » (ou toute autre cellule déjà remplie).
export function templateCapacity(tpl, sheet) {
  if (tpl.capacity) return tpl.capacity;
  const lineCols = Object.keys(tpl.lineCols || {});
  if (!tpl.lineRow || !lineCols.length) return Infinity;
  const colNums = new Set(lineCols.map(colToNum));
  let boundary = Infinity;
  for (const cell of sheet.cells.values()) {
    if (cell.r <= tpl.lineRow || !colNums.has(cell.c)) continue;
    const isTotalFormula = cell.formula && /SUM|SUBTOTAL|AGGREGATE|:/i.test(cell.formula);
    const isLabel = typeof cell.val === 'string' && cell.val.trim();
    if (isTotalFormula || isLabel) boundary = Math.min(boundary, cell.r);
  }
  for (const ref of Object.keys(tpl.cells || {})) {
    const p = splitRef(ref);
    if (p && p.r > tpl.lineRow && colNums.has(p.c)) boundary = Math.min(boundary, p.r);
  }
  for (const m of sheet.merges) {
    if (m.r1 <= tpl.lineRow) continue;
    for (const c of colNums) if (c >= m.c1 && c <= m.c2) boundary = Math.min(boundary, m.r1);
  }
  return boundary === Infinity ? Infinity : boundary - tpl.lineRow;
}

/* ---------- Écriture dans un modèle ---------- */

class SheetWriter {
  constructor(doc, styles, shared = []) {
    this.doc = doc;
    this.shared = shared;
    this.ns = doc.documentElement.namespaceURI || NS_MAIN;
    this.styles = styles;
    this.sheetData = all(doc, 'sheetData')[0];
    this.rows = new Map(kids(this.sheetData, 'row').map(r => [+r.getAttribute('r'), r]));
    this.colStyles = [];
    for (const c of all(doc, 'col')) {
      if (!c.getAttribute('style')) continue;
      for (let i = +c.getAttribute('min'); i <= +c.getAttribute('max'); i++) this.colStyles[i] = +c.getAttribute('style');
    }
    this.brokenShared = new Set();
  }

  el(name) { return this.doc.createElementNS(this.ns, name); }

  row(r) {
    let row = this.rows.get(r);
    if (row) return row;
    row = this.el('row');
    row.setAttribute('r', String(r));
    const next = [...this.rows.keys()].filter(k => k > r).sort((a, b) => a - b)[0];
    if (next != null) this.sheetData.insertBefore(row, this.rows.get(next)); else this.sheetData.append(row);
    this.rows.set(r, row);
    return row;
  }

  cell(ref) {
    const p = splitRef(ref);
    const row = this.row(p.r);
    row.removeAttribute('spans');
    for (const c of kids(row, 'c')) {
      const q = splitRef(c.getAttribute('r'));
      if (!q) continue;
      if (q.c === p.c) return c;
      if (q.c > p.c) {
        const n = this.el('c');
        n.setAttribute('r', ref);
        row.insertBefore(n, c);
        return n;
      }
    }
    const n = this.el('c');
    n.setAttribute('r', ref);
    row.append(n);
    return n;
  }

  existingStyle(ref) {
    const p = splitRef(ref);
    const row = this.rows.get(p.r);
    const c = row && kids(row, 'c').find(x => x.getAttribute('r') === ref);
    if (c?.hasAttribute('s')) return +c.getAttribute('s');
    if (row?.getAttribute('customFormat') === '1' && row.hasAttribute('s')) return +row.getAttribute('s');
    return this.colStyles[p.c] ?? null;
  }

  ensureFormat(s, accept, numFmtId) {
    const st = this.styles;
    if (!st?.cellXfs) return s;
    if (accept(st.xfs[s] ?? 0)) return s;
    const key = `${s}:${numFmtId}`;
    if (st.cache.has(key)) return st.cache.get(key);
    const source = kids(st.cellXfs, 'xf')[s] || kids(st.cellXfs, 'xf')[0];
    const clone = source ? source.cloneNode(true) : st.doc.createElementNS(st.doc.documentElement.namespaceURI, 'xf');
    clone.setAttribute('numFmtId', String(numFmtId));
    clone.setAttribute('applyNumberFormat', '1');
    st.cellXfs.append(clone);
    st.xfs.push(numFmtId);
    const index = st.xfs.length - 1;
    st.cellXfs.setAttribute('count', String(st.xfs.length));
    st.cache.set(key, index);
    st.dirty = true;
    return index;
  }

  set(ref, value, type = 'text', fallbackStyle = null) {
    const c = this.cell(ref);
    let s = c.hasAttribute('s') ? +c.getAttribute('s') : (fallbackStyle ?? this.existingStyle(ref) ?? 0);
    const f = kid(c, 'f');
    if (f?.getAttribute('t') === 'shared' && f.hasAttribute('ref')) this.brokenShared.add(f.getAttribute('si'));
    while (c.firstChild) c.firstChild.remove();
    c.removeAttribute('t');

    const writeNumber = n => { const v = this.el('v'); v.textContent = String(n); c.append(v); };
    let done = value === '' || value == null;
    if (!done && type === 'date') {
      const serial = isoToSerial(value);
      if (serial != null) {
        s = this.ensureFormat(s, id => isDateFormat(id, this.styles?.numFmts || {}), 14);
        writeNumber(serial);
        done = true;
      }
    }
    if (!done && (type === 'money' || type === 'int' || type === 'number')) {
      const n = Number(value);
      if (isFinite(n)) {
        if (type === 'money') s = this.ensureFormat(s, id => id !== 0, 4);
        writeNumber(n);
        done = true;
      }
    }
    if (!done) {
      c.setAttribute('t', 'inlineStr');
      const is = this.el('is');
      const t = this.el('t');
      const text = String(value).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '');
      t.textContent = text;
      if (/^\s|\s$|\n/.test(text)) t.setAttributeNS(NS_XML, 'xml:space', 'preserve');
      is.append(t);
      c.append(is);
    }
    if (s) c.setAttribute('s', String(s)); else c.removeAttribute('s');
  }

  finish() {
    // Une formule partagée dont la cellule « maître » a été écrasée : on retire les copies orphelines.
    if (this.brokenShared.size) {
      for (const f of all(this.doc, 'f')) {
        if (f.getAttribute('t') === 'shared' && this.brokenShared.has(f.getAttribute('si'))) {
          const c = f.parentElement;
          kid(c, 'v')?.remove();
          f.remove();
          c.removeAttribute('t');
        }
      }
    }
    // Résultats des formules écrits dans le fichier : les totaux s'affichent même sans recalcul
    // (mode protégé d'Excel, aperçus Outlook / iPhone). Excel recalcule de toute façon à l'ouverture.
    computeFormulas(this.doc, this.shared);
    kid(this.doc.documentElement, 'dimension')?.remove();
    return serialize(this.doc);
  }
}

/* ---------- Calcul des formules ---------- */

const FAIL = Symbol('fail');

// Décale les références relatives d'une formule (formules partagées).
function offsetFormula(formula, dr, dc) {
  return formula.split(/("(?:[^"]|"")*")/).map((part, i) => (i % 2 ? part : part.replace(REF_RX,
    (m, prefix, d1, c1, d2, r1, d3, c2, d4, r2) => {
      const one = (d, c, dd, r) => `${d}${d ? c : numToCol(colToNum(c) + dc)}${dd}${dd ? r : +r + dr}`;
      return `${prefix || ''}${one(d1, c1, d2, r1)}${c2 ? ':' + one(d3, c2, d4, r2) : ''}`;
    }))).join('');
}

function tokenize(s) {
  const tokens = [];
  let i = 0;
  while (i < s.length) {
    const rest = s.slice(i);
    if (/^\s/.test(rest)) { i++; continue; }
    if (rest[0] === '"') {
      let j = i + 1;
      let str = '';
      while (j < s.length && !(s[j] === '"' && s[j + 1] !== '"')) { str += s[j]; j += s[j] === '"' ? 2 : 1; }
      tokens.push({ t: 'str', v: str });
      i = j + 1;
      continue;
    }
    if (/^(?:'[^']+'|[A-Za-z_][\w.]*)!/.test(rest)) throw FAIL; // autre feuille : non géré
    let m = /^\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?(?![\w(])/.exec(rest);
    if (m) {
      tokens.push(m[3] ? { t: 'range', c1: colToNum(m[1]), r1: +m[2], c2: colToNum(m[3]), r2: +m[4] } : { t: 'ref', ref: `${m[1]}${m[2]}` });
      i += m[0].length;
      continue;
    }
    if ((m = /^\d+(?:\.\d+)?(?:E[+-]?\d+)?/i.exec(rest))) { tokens.push({ t: 'num', v: parseFloat(m[0]) }); i += m[0].length; continue; }
    if ((m = /^(TRUE|FALSE)(?![\w(])/i.exec(rest))) { tokens.push({ t: 'bool', v: m[1].toUpperCase() === 'TRUE' }); i += m[0].length; continue; }
    if ((m = /^([A-Z][A-Z0-9.]*)\(/i.exec(rest))) { tokens.push({ t: 'fn', v: m[1].toUpperCase() }); i += m[0].length; continue; }
    if ((m = /^(<>|<=|>=|[-+*/^&=<>(),;%])/.exec(rest))) { tokens.push({ t: 'op', v: m[1] }); i += m[1].length; continue; }
    throw FAIL;
  }
  return tokens;
}

function evaluate(formula, valueOf, rangeOf) {
  const tk = tokenize(formula);
  let pos = 0;
  const peek = () => tk[pos];
  const next = () => tk[pos++];
  const isOp = v => peek()?.t === 'op' && peek().v === v;
  const num = x => {
    if (x == null || x === '') return 0;
    if (typeof x === 'boolean') return x ? 1 : 0;
    if (typeof x === 'number') return x;
    const n = Number(x);
    if (isFinite(n)) return n;
    throw FAIL;
  };
  const numbers = args => args.flatMap(a => (a?.range ? a.range.filter(v => typeof v === 'number') : [num(a)]));
  const round = (x, n, mode) => {
    const k = 10 ** num(n);
    const v = Math.abs(num(x)) * k;
    const r = mode === 'up' ? Math.ceil(v - 1e-9) : mode === 'down' ? Math.floor(v + 1e-9) : Math.round(v);
    return Math.sign(num(x)) * r / k;
  };
  const call = (name, args) => {
    switch (name) {
      case 'SUM': return numbers(args).reduce((a, b) => a + b, 0);
      case 'AVERAGE': { const n = numbers(args); if (!n.length) throw FAIL; return n.reduce((a, b) => a + b, 0) / n.length; }
      case 'MIN': { const n = numbers(args); return n.length ? Math.min(...n) : 0; }
      case 'MAX': { const n = numbers(args); return n.length ? Math.max(...n) : 0; }
      case 'COUNT': return numbers(args).length;
      case 'ROUND': return round(args[0], args[1] ?? 0);
      case 'ROUNDUP': return round(args[0], args[1] ?? 0, 'up');
      case 'ROUNDDOWN': return round(args[0], args[1] ?? 0, 'down');
      case 'ABS': return Math.abs(num(args[0]));
      case 'IF': return num(args[0]) ? (args[1] ?? true) : (args[2] ?? false);
      case 'AND': return args.every(a => num(a));
      case 'OR': return args.some(a => num(a));
      default: throw FAIL;
    }
  };
  const compare = (a, b, op) => {
    const [x, y] = typeof a === 'string' || typeof b === 'string' ? [String(a ?? ''), String(b ?? '')] : [num(a), num(b)];
    return { '=': x === y, '<>': x !== y, '<': x < y, '>': x > y, '<=': x <= y, '>=': x >= y }[op];
  };
  function primary() {
    const t = next();
    if (!t) throw FAIL;
    if (t.t === 'num' || t.t === 'str' || t.t === 'bool') return t.v;
    if (t.t === 'ref') { const v = valueOf(t.ref); if (v === FAIL) throw FAIL; return v; }
    if (t.t === 'op' && t.v === '(') { const v = comparison(); if (!isOp(')')) throw FAIL; next(); return v; }
    if (t.t === 'fn') {
      const args = [];
      if (!isOp(')')) {
        do { args.push(peek()?.t === 'range' && [',', ';', ')'].includes(tk[pos + 1]?.v) ? { range: rangeOf(next()) } : comparison()); }
        while ((isOp(',') || isOp(';')) && next());
      }
      if (!isOp(')')) throw FAIL;
      next();
      return call(t.v, args);
    }
    throw FAIL;
  }
  function unary() {
    if (isOp('-')) { next(); return -num(unary()); }
    if (isOp('+')) { next(); return unary(); }
    let v = primary();
    while (isOp('%')) { next(); v = num(v) / 100; }
    return v;
  }
  function power() { let a = unary(); while (isOp('^')) { next(); a = num(a) ** num(unary()); } return a; }
  function product() {
    let a = power();
    while (isOp('*') || isOp('/')) {
      const op = next().v;
      const b = num(power());
      if (op === '/' && b === 0) throw FAIL;
      a = op === '*' ? num(a) * b : num(a) / b;
    }
    return a;
  }
  function sum() {
    let a = product();
    while (isOp('+') || isOp('-')) { const op = next().v; const b = product(); a = op === '+' ? num(a) + num(b) : num(a) - num(b); }
    return a;
  }
  function concat() { let a = sum(); while (isOp('&')) { next(); a = String(a ?? '') + String(sum() ?? ''); } return a; }
  function comparison() {
    let a = concat();
    while (peek()?.t === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(peek().v)) { const op = next().v; a = compare(a, concat(), op); }
    return a;
  }
  const out = comparison();
  if (pos !== tk.length) throw FAIL;
  return out;
}

function computeFormulas(doc, shared = []) {
  const ns = doc.documentElement.namespaceURI || NS_MAIN;
  const cells = new Map();
  for (const c of all(doc, 'c')) {
    const p = splitRef(c.getAttribute('r'));
    if (!p) continue;
    const f = kid(c, 'f');
    const v = kid(c, 'v');
    const t = c.getAttribute('t');
    let value = null;
    if (!f) {
      if (t === 's' && v) value = shared[+v.textContent] ?? '';
      else if (t === 'inlineStr') value = textOf(kid(c, 'is') || c);
      else if (t === 'str' && v) value = v.textContent;
      else if (t === 'b' && v) value = v.textContent === '1';
      else if (t === 'e') value = FAIL;
      else if (v) value = Number(v.textContent);
    }
    cells.set(`${p.col}${p.r}`, { el: c, f, p, value, formula: null });
  }
  const masters = {};
  for (const cell of cells.values()) {
    if (cell.f?.getAttribute('t') === 'shared' && cell.f.textContent) masters[cell.f.getAttribute('si')] = { text: cell.f.textContent, p: cell.p };
  }
  for (const cell of cells.values()) {
    const f = cell.f;
    if (!f || f.getAttribute('t') === 'array') continue;
    if (f.getAttribute('t') === 'shared' && !f.textContent) {
      const m = masters[f.getAttribute('si')];
      cell.formula = m ? offsetFormula(m.text, cell.p.r - m.p.r, cell.p.c - m.p.c) : null;
    } else {
      cell.formula = f.textContent;
    }
  }
  const memo = new Map();
  const visiting = new Set();
  const valueOf = ref => {
    const cell = cells.get(ref);
    if (!cell) return null;
    if (!cell.f) return cell.value;
    if (memo.has(ref)) return memo.get(ref);
    if (visiting.has(ref) || !cell.formula) return FAIL;
    visiting.add(ref);
    let out;
    try { out = evaluate(cell.formula, valueOf, rangeOf); } catch { out = FAIL; }
    visiting.delete(ref);
    memo.set(ref, out);
    return out;
  };
  const rangeOf = r => {
    const out = [];
    for (const [ref, cell] of cells) {
      if (cell.p.r < r.r1 || cell.p.r > r.r2 || cell.p.c < r.c1 || cell.p.c > r.c2) continue;
      const v = valueOf(ref);
      if (v === FAIL) throw FAIL;
      out.push(v);
    }
    return out;
  };
  for (const [ref, cell] of cells) {
    if (!cell.f) continue;
    const res = valueOf(ref);
    kid(cell.el, 'v')?.remove();
    cell.el.removeAttribute('t');
    if (res === FAIL || res == null) continue;
    const v = doc.createElementNS(ns, 'v');
    if (typeof res === 'number') {
      if (!isFinite(res)) continue;
      v.textContent = String(Math.round(res * 1e10) / 1e10);
    } else if (typeof res === 'boolean') {
      cell.el.setAttribute('t', 'b');
      v.textContent = res ? '1' : '0';
    } else {
      cell.el.setAttribute('t', 'str');
      v.textContent = String(res);
    }
    cell.el.append(v);
  }
}

const CALC_PR_AFTER = ['oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst'];

async function forceRecalc(book) {
  const { zip, wb } = book;
  const root = wb.documentElement;
  let calcPr = kid(root, 'calcPr');
  if (!calcPr) {
    calcPr = wb.createElementNS(root.namespaceURI, 'calcPr');
    const before = [...root.children].find(ch => CALC_PR_AFTER.includes(ch.localName));
    if (before) root.insertBefore(calcPr, before); else root.append(calcPr);
  }
  calcPr.setAttribute('fullCalcOnLoad', '1');
  zip.file('xl/workbook.xml', serialize(wb));

  if (zip.file('xl/calcChain.xml')) {
    zip.remove('xl/calcChain.xml');
    const ct = parseXml(await zip.file('[Content_Types].xml').async('string'));
    for (const o of all(ct, 'Override')) if (/calcChain/i.test(o.getAttribute('PartName'))) o.remove();
    zip.file('[Content_Types].xml', serialize(ct));
    const relsDoc = parseXml(await zip.file('xl/_rels/workbook.xml.rels').async('string'));
    for (const r of all(relsDoc, 'Relationship')) if (/calcChain/i.test(r.getAttribute('Target'))) r.remove();
    zip.file('xl/_rels/workbook.xml.rels', serialize(relsDoc));
  }
}

/* ---------- Colonnes conditionnelles (type de dépense, devise) ---------- */

const SKIP = Symbol('skip');

function lineColumns(tpl) {
  const cols = Object.entries(tpl.lineCols || {}).map(([col, v]) => ({ col, ...lineColDef(v) }));
  return {
    cols,
    covered: new Set(cols.flatMap(c => c.categories || [])),
    dedicated: new Set(cols.flatMap(c => c.currencies || [])),
  };
}

// Valeur d'une colonne pour un justificatif, ou SKIP si la colonne ne le concerne pas.
function lineCellValue(c, item, covered, dedicated) {
  if (c.categories?.length || c.others) {
    const cat = item.category || 'Autre';
    if (!(c.categories?.includes(cat) || (c.others && !covered.has(cat)))) return SKIP;
  }
  if (c.currencies?.length) {
    const cur = item.currency || 'EUR';
    if (c.currencies.includes(cur)) return lineValue(c.field, item);
    // Devise sans colonne propre : montant converti dans la colonne en euros.
    if (c.currencies.includes('EUR') && !dedicated.has(cur)) return lineValue(EUR_EQUIVALENT[c.field] || c.field, item);
    return SKIP;
  }
  return lineValue(c.field, item);
}

function placementWarnings(tpl, report) {
  const { cols, covered, dedicated } = lineColumns(tpl);
  const money = cols.filter(c => fieldDef(c.field)?.type === 'money');
  if (!money.length) return [];
  const warnings = [];
  for (const item of report.items) {
    if (money.some(c => lineCellValue(c, item, covered, dedicated) !== SKIP)) continue;
    warnings.push(`« ${item.merchant || 'Justificatif ' + item.index} » (${item.category || 'sans catégorie'}, ${item.currency || 'EUR'}) n'a pas de colonne de montant dans ce modèle.`);
  }
  for (const item of report.items) {
    if (item.currency && item.currency !== 'EUR' && !item.rate) warnings.push(`Taux de change manquant pour « ${item.merchant || 'Justificatif ' + item.index} » (${item.currency}).`);
  }
  return warnings;
}

/* ---------- Insertion de lignes (le tableau du modèle est trop court) ---------- */

const REF_RX = /(?<![\w.À-ɏ])((?:'(?:[^']|'')+'|[A-Za-z_À-ɏ][\w.À-ɏ]*)!)?(\$?)([A-Z]{1,3})(\$?)(\d+)(?::(\$?)([A-Z]{1,3})(\$?)(\d+))?(?![\w(!])/g;
const unquoteSheet = p => p.slice(0, -1).replace(/^'(.*)'$/, '$1').replace(/''/g, "'");

// Applique `fn(ligne, absolu)` à chaque référence de cellule d'une formule (hors chaînes de texte).
function mapRefs(formula, appliesTo, fn) {
  return formula.split(/("(?:[^"]|"")*")/).map((part, i) => (i % 2 ? part : part.replace(REF_RX,
    (m, prefix, d1, c1, d2, r1, d3, c2, d4, r2) => {
      if (!appliesTo(prefix)) return m;
      let out = `${prefix || ''}${d1}${c1}${d2}${fn(+r1, !!d2)}`;
      if (c2) out += `:${d3}${c2}${d4}${fn(+r2, !!d4)}`;
      return out;
    }))).join('');
}

async function insertRows(book, doc, sheet, at, n) {
  const shift = r => (r >= at ? r + n : r);
  const same = prefix => !prefix || unquoteSheet(prefix) === sheet.name;
  const other = prefix => !!prefix && unquoteSheet(prefix) === sheet.name;
  const shiftText = (text, applies = same) => mapRefs(text, applies, r => shift(r));
  const shiftAttr = (el, attr) => { if (el.hasAttribute(attr)) el.setAttribute(attr, shiftText(el.getAttribute(attr))); };

  const sheetData = all(doc, 'sheetData')[0];
  const rows = kids(sheetData, 'row');
  // Modèle des nouvelles lignes : la ligne du tableau juste au-dessus du point d'insertion.
  const source = rows.find(r => +r.getAttribute('r') === at - 1) || rows.find(r => +r.getAttribute('r') === at);
  const sourceR = source ? +source.getAttribute('r') : at;
  const template = source ? source.cloneNode(true) : null;

  for (const row of rows) {
    const r = +row.getAttribute('r');
    if (r < at) continue;
    row.setAttribute('r', String(r + n));
    row.removeAttribute('spans');
    for (const c of kids(row, 'c')) {
      const p = splitRef(c.getAttribute('r'));
      if (p) c.setAttribute('r', `${p.col}${p.r + n}`);
    }
  }
  for (const f of all(doc, 'f')) {
    if (f.textContent) f.textContent = shiftText(f.textContent);
    shiftAttr(f, 'ref');
  }
  for (const el of all(doc, 'mergeCell')) shiftAttr(el, 'ref');
  for (const el of all(doc, 'hyperlink')) shiftAttr(el, 'ref');
  for (const el of all(doc, 'autoFilter')) shiftAttr(el, 'ref');
  for (const el of all(doc, 'selection')) { shiftAttr(el, 'sqref'); shiftAttr(el, 'activeCell'); }
  for (const el of all(doc, 'pane')) shiftAttr(el, 'topLeftCell');
  for (const el of [...all(doc, 'conditionalFormatting'), ...all(doc, 'dataValidation')]) {
    shiftAttr(el, 'sqref');
    for (const f of [...all(el, 'formula'), ...all(el, 'formula1'), ...all(el, 'formula2')]) f.textContent = shiftText(f.textContent);
  }

  // Nouvelles lignes : même hauteur, mêmes styles, mêmes formules ligne à ligne.
  const ns = doc.documentElement.namespaceURI || NS_MAIN;
  const before = kids(sheetData, 'row').find(r => +r.getAttribute('r') >= at + n) || null;
  for (let i = 0; i < n; i++) {
    const newR = at + i;
    const row = doc.createElementNS(ns, 'row');
    row.setAttribute('r', String(newR));
    for (const a of ['ht', 'customHeight', 's', 'customFormat']) if (template?.hasAttribute(a)) row.setAttribute(a, template.getAttribute(a));
    for (const src of template ? kids(template, 'c') : []) {
      const p = splitRef(src.getAttribute('r'));
      if (!p) continue;
      const c = doc.createElementNS(ns, 'c');
      c.setAttribute('r', `${p.col}${newR}`);
      if (src.hasAttribute('s')) c.setAttribute('s', src.getAttribute('s'));
      const f = kid(src, 'f');
      if (f?.textContent && f.getAttribute('t') !== 'shared' && f.getAttribute('t') !== 'array') {
        const nf = doc.createElementNS(ns, 'f');
        nf.textContent = mapRefs(f.textContent, () => true, (r, abs) => (abs ? r : r + (newR - sourceR)));
        c.append(nf);
      }
      row.append(c);
    }
    sheetData.insertBefore(row, before);
  }

  // Autres feuilles et noms définis qui pointent vers cette feuille.
  for (const s of book.sheets) {
    if (s.path === sheet.path) continue;
    const xml = await book.zip.file(s.path).async('string');
    if (!xml.includes('<f') || !xml.includes(sheet.name)) continue;
    const d = parseXml(xml);
    for (const f of all(d, 'f')) if (f.textContent) f.textContent = shiftText(f.textContent, other);
    book.zip.file(s.path, serialize(d));
  }
  for (const dn of all(book.wb, 'definedName')) dn.textContent = shiftText(dn.textContent, other);

  // Images et tableaux ancrés sous le point d'insertion.
  const relsPath = sheet.path.replace(/([^/]+)$/, '_rels/$1.rels');
  const relsFile = book.zip.file(relsPath);
  if (relsFile) {
    const rels = parseXml(await relsFile.async('string'));
    for (const rel of all(rels, 'Relationship')) {
      const type = rel.getAttribute('Type') || '';
      const target = resolvePart(rel.getAttribute('Target') || '', sheet.path.replace(/\/[^/]+$/, ''));
      const file = book.zip.file(target);
      if (!file) continue;
      if (/\/drawing$/.test(type)) {
        const d = parseXml(await file.async('string'));
        for (const anchor of [...all(d, 'from'), ...all(d, 'to')]) {
          const rowEl = kid(anchor, 'row');
          if (rowEl && +rowEl.textContent >= at - 1) rowEl.textContent = String(+rowEl.textContent + n);
        }
        book.zip.file(target, serialize(d));
      } else if (/\/table$/.test(type)) {
        const d = parseXml(await file.async('string'));
        shiftAttr(d.documentElement, 'ref');
        for (const el of all(d, 'autoFilter')) shiftAttr(el, 'ref');
        book.zip.file(target, serialize(d));
      }
    }
  }
  kid(doc.documentElement, 'dimension')?.remove();
}

/* ---------- Remplissage ---------- */

async function writeTemplate(book, sheet, tpl, report, capacity) {
  const doc = parseXml(await book.zip.file(sheet.path).async('string'));
  const shiftRef = (ref, at, n) => { const p = splitRef(ref); return p && p.r >= at ? `${p.col}${p.r + n}` : ref; };
  let at = Infinity;
  let extra = 0;
  if (tpl.lineRow && isFinite(capacity) && report.items.length > capacity) {
    // On agrandit le tableau avant sa dernière ligne : les sommes du type SUM(D9:D18) s'étendent d'elles-mêmes.
    extra = report.items.length - capacity;
    at = capacity >= 2 ? tpl.lineRow + capacity - 1 : tpl.lineRow + 1;
    await insertRows(book, doc, sheet, at, extra);
  }
  const w = new SheetWriter(doc, book.styles, book.shared);
  for (const [ref, m] of Object.entries(tpl.cells || {})) {
    const def = fieldDef(m.field);
    if (def) w.set(shiftRef(ref, at, extra), singleValue(m.field, report, m.value, m.currency), def.type);
  }
  if (tpl.lineRow) {
    const { cols, covered, dedicated } = lineColumns(tpl);
    for (const c of cols) {
      const def = fieldDef(c.field);
      if (!def) continue;
      const firstStyle = w.existingStyle(`${c.col}${tpl.lineRow}`);
      report.items.forEach((item, i) => {
        const v = lineCellValue(c, item, covered, dedicated);
        if (v !== SKIP) w.set(`${c.col}${tpl.lineRow + i}`, v, def.type, firstStyle);
      });
    }
  }
  book.zip.file(sheet.path, w.finish());
  if (book.styles?.dirty) book.zip.file('xl/styles.xml', serialize(book.styles.doc));
  await forceRecalc(book);
  return extra;
}

export async function fillTemplate(tpl, report, baseName) {
  const book = await openWorkbook(tpl.data);
  const sheet = book.sheets.find(s => s.name === tpl.sheet) || book.sheets[0];
  const content = await readSheet(book, sheet.path);
  const capacity = templateCapacity(tpl, content);
  const warnings = placementWarnings(tpl, report);
  const added = await writeTemplate(book, sheet, tpl, report, capacity);
  const blob = await book.zip.generateAsync({ type: 'blob', mimeType: XLSX_MIME, compression: 'DEFLATE' });
  return { files: [{ name: `${baseName}.xlsx`, blob, mime: XLSX_MIME }], capacity, added, warnings, content };
}

/* ---------- Anciens fichiers .xls ---------- */

// Conversion .xls → .xlsx dans le navigateur : valeurs, formules, fusions, largeurs et formats de nombres.
// Les couleurs et bordures ne sont pas reprises : Excel (Enregistrer sous… .xlsx) garde tout.
export async function convertXlsToXlsx(file) {
  await loadScript(CDN.sheetjs);
  const XLSX = window.XLSX;
  const wb = XLSX.read(await file.arrayBuffer(), { type: 'array', cellStyles: true, cellNF: true, cellFormula: true });
  const out = XLSX.write(wb, { type: 'array', bookType: 'xlsx', cellStyles: true, compression: true });
  return blobToBase64(new Blob([out], { type: XLSX_MIME }));
}

/* ---------- Description d'une feuille pour l'analyse par l'IA ---------- */

export function describeSheet(content, { maxRows = 150, maxCols = 40 } = {}) {
  const byRow = new Map();
  for (const c of content.cells.values()) {
    if (c.r > maxRows || c.c > maxCols || (c.display === '' && !c.formula)) continue;
    if (!byRow.has(c.r)) byRow.set(c.r, []);
    byRow.get(c.r).push(c);
  }
  const lines = [...byRow.keys()].sort((a, b) => a - b).map(r => {
    const cells = byRow.get(r).sort((a, b) => a.c - b.c).map(c => {
      const value = typeof c.val === 'number' && !c.isDate ? c.val : c.display;
      let s = `${c.ref}: ${c.formula ? '=' + c.formula : JSON.stringify(value)}`;
      if (c.fmt) s += ` [format ${c.fmt}]`;
      return s;
    });
    return `Ligne ${r} → ${cells.join(' | ')}`;
  });
  const merges = content.merges.map(m => `${numToCol(m.c1)}${m.r1}:${numToCol(m.c2)}${m.r2}`);
  const widths = content.cols.filter(c => c.width && c.min <= content.maxCol + 1)
    .map(c => `${numToCol(c.min)}${Math.min(c.max, content.maxCol + 1) > c.min ? '-' + numToCol(Math.min(c.max, content.maxCol + 1)) : ''}=${Math.round(c.width)}`);
  return [
    `Dernière ligne utilisée : ${content.maxRow}, dernière colonne : ${numToCol(content.maxCol)}.`,
    merges.length ? `Cellules fusionnées : ${merges.join(', ')}.` : 'Aucune cellule fusionnée.',
    widths.length ? `Largeurs de colonnes : ${widths.join(', ')}.` : '',
    'Contenu (cellules non vides) :',
    ...lines,
  ].filter(Boolean).join('\n');
}

/* ---------- Tableau standard ---------- */

const STYLES_XML = (currencySymbol) => `${XML_DECL}<styleSheet xmlns="${NS_MAIN}">
<numFmts count="3"><numFmt numFmtId="164" formatCode="#,##0.00\\ &quot;${escXml(currencySymbol)}&quot;"/><numFmt numFmtId="165" formatCode="dd/mm/yyyy"/><numFmt numFmtId="166" formatCode="0.0000"/></numFmts>
<fonts count="4">
<font><sz val="11"/><color rgb="FF1D1D1F"/><name val="Calibri"/><family val="2"/></font>
<font><b/><sz val="11"/><color rgb="FF1D1D1F"/><name val="Calibri"/><family val="2"/></font>
<font><b/><sz val="18"/><color rgb="FF1D1D1F"/><name val="Calibri"/><family val="2"/></font>
<font><sz val="11"/><color rgb="FF6E6E73"/><name val="Calibri"/><family val="2"/></font>
</fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F7"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="3">
<border><left/><right/><top/><bottom/><diagonal/></border>
<border><left/><right/><top/><bottom style="thin"><color rgb="FFD8D8DC"/></bottom><diagonal/></border>
<border><left/><right/><top style="medium"><color rgb="FF1D1D1F"/></top><bottom/><diagonal/></border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="14">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="top"/></xf>
<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment vertical="top"/></xf>
<xf numFmtId="1" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="top"/></xf>
<xf numFmtId="0" fontId="1" fillId="0" borderId="2" xfId="0" applyFont="1" applyBorder="1"/>
<xf numFmtId="164" fontId="1" fillId="0" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="2" xfId="0" applyBorder="1"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
<xf numFmtId="4" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment vertical="top"/></xf>
<xf numFmtId="166" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment vertical="top"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

const S = { title: 1, info: 2, header: 3, text: 4, date: 5, money: 6, int: 7, totalLabel: 8, totalMoney: 9, totalEmpty: 10, headerRight: 11, moneyPlain: 12, number: 13 };
const SYMBOLS = { EUR: '€', USD: '$', GBP: '£', CHF: 'CHF', CAD: 'CA$' };

function strCell(ref, text, s) {
  return `<c r="${ref}" s="${s}" t="inlineStr"><is><t xml:space="preserve">${escXml(text)}</t></is></c>`;
}
function numCell(ref, n, s) {
  return `<c r="${ref}" s="${s}"><v>${n}</v></c>`;
}

export async function buildSimple(cfg, report, baseName) {
  const JSZip = await getJSZip();
  let columns = (cfg?.columns?.length ? cfg.columns : DEFAULT_SIMPLE_COLUMNS).filter(c => fieldDef(c.field));
  // Plusieurs devises : montants d'origine sans symbole, colonnes de conversion ajoutées, totaux en euros.
  const foreign = report.items.some(i => i.currency && i.currency !== 'EUR');
  if (foreign) {
    const has = f => columns.some(c => c.field === f);
    columns = [
      ...columns,
      ...(has('currency') ? [] : [{ field: 'currency', label: 'Devise' }]),
      ...(has('rate') ? [] : [{ field: 'rate', label: 'Taux €' }]),
      ...(has('totalEur') ? [] : [{ field: 'totalEur', label: 'TVAC en €' }]),
    ];
  }
  const inEuros = key => !foreign || /Eur$/.test(key);
  const defs = columns.map(c => ({ ...fieldDef(c.field), label: c.label || fieldDef(c.field).label }));
  const rows = [];
  let r = 1;
  const addRow = (cells, height) => { rows.push(`<row r="${r}"${height ? ` ht="${height}" customHeight="1"` : ''}>${cells.join('')}</row>`); r++; };

  addRow([strCell(`A${r}`, cfg?.title || 'Note de frais', S.title)], 28);
  const info = [
    report.profile?.name && `Nom : ${report.profile.name}${report.profile.company ? ' – ' + report.profile.company : ''}`,
    report.client?.name && `Destinataire : ${report.client.name}${report.client.company ? ' – ' + report.client.company : ''}`,
    `Période : ${report.period}`,
    `Date : ${parseISO(report.today).toLocaleDateString('fr-BE')}`,
    report.profile?.iban && `IBAN : ${report.profile.iban}`,
  ].filter(Boolean);
  for (const line of info) addRow([strCell(`A${r}`, line, S.info)]);
  r++;

  const headerRow = r;
  addRow(defs.map((d, i) => strCell(`${numToCol(i + 1)}${r}`, d.label, d.type === 'money' ? S.headerRight : S.header)), 22);
  const firstLine = r;
  for (const item of report.items) {
    addRow(defs.map((d, i) => {
      const ref = `${numToCol(i + 1)}${r}`;
      const v = lineValue(d.key, item);
      if (v === '' || v == null) return `<c r="${ref}" s="${d.type === 'money' ? S.money : S.text}"/>`;
      if (d.type === 'date') { const n = isoToSerial(v); return n != null ? numCell(ref, n, S.date) : strCell(ref, v, S.text); }
      if (d.type === 'money') return numCell(ref, Number(v), inEuros(d.key) ? S.money : S.moneyPlain);
      if (d.type === 'int') return numCell(ref, Number(v), S.int);
      if (d.type === 'number') return numCell(ref, Number(v), S.number);
      return strCell(ref, v, S.text);
    }));
  }
  const lastLine = r - 1;

  const firstMoney = defs.findIndex(d => d.type === 'money');
  if (firstMoney >= 0) {
    const totals = { net: report.net, vat: report.vat, total: report.total, totalEur: report.totalEur };
    addRow(defs.map((d, i) => {
      const col = numToCol(i + 1);
      const ref = `${col}${r}`;
      if (d.type === 'money' && inEuros(d.key)) {
        const cached = totals[d.key] ?? Math.round(report.items.reduce((s, it) => s + (Number(it[d.key]) || 0), 0) * 100) / 100;
        return report.items.length
          ? `<c r="${ref}" s="${S.totalMoney}"><f>SUM(${col}${firstLine}:${col}${lastLine})</f><v>${cached}</v></c>`
          : numCell(ref, 0, S.totalMoney);
      }
      if (i === Math.max(0, firstMoney - 1)) return strCell(ref, 'Total', S.totalLabel);
      return `<c r="${ref}" s="${S.totalEmpty}"/>`;
    }), 22);
  }

  const colsXml = defs.map((d, i) => `<col min="${i + 1}" max="${i + 1}" width="${d.width || 16}" customWidth="1"/>`).join('');
  const sheetXml = `${XML_DECL}<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_REL}">
<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>
<sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane ySplit="${headerRow}" topLeftCell="A${headerRow + 1}" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
<sheetFormatPr defaultRowHeight="16"/>
<cols>${colsXml}</cols>
<sheetData>${rows.join('')}</sheetData>
<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/>
<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/>
</worksheet>`;

  const sheetName = escXml((cfg?.title || 'Note de frais').replace(/[[\]:*?/\\]/g, ' ').slice(0, 31).trim() || 'Note de frais');
  const zip = new JSZip();
  zip.file('[Content_Types].xml', `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`);
  zip.file('_rels/.rels', `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file('xl/workbook.xml', `${XML_DECL}<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_REL}"><bookViews><workbookView/></bookViews><sheets><sheet name="${sheetName}" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
  zip.file('xl/styles.xml', STYLES_XML(SYMBOLS[report.currency] || report.currency || '€'));
  zip.file('xl/worksheets/sheet1.xml', sheetXml);
  const blob = await zip.generateAsync({ type: 'blob', mimeType: XLSX_MIME, compression: 'DEFLATE' });
  return { files: [{ name: `${baseName}.xlsx`, blob, mime: XLSX_MIME }], split: false };
}

export const lineFieldOptions = () => LINE_FIELDS;
