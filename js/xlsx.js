// Fichiers Excel.
// - Modèle du client : on modifie directement le XML du .xlsx pour ne toucher qu'aux cellules
//   choisies ; mise en forme, logos, formules et validations restent intacts.
// - Tableau standard : classeur sobre généré de zéro.

import { getJSZip } from './lib.js';
import { fieldDef, lineValue, singleValue, summarize, LINE_FIELDS, DEFAULT_SIMPLE_COLUMNS } from './fields.js';
import { parseISO } from './ui.js';

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

function resolvePart(target) {
  if (target.startsWith('/')) return target.slice(1);
  const out = [];
  for (const p of ('xl/' + target).split('/')) {
    if (p === '..') out.pop(); else if (p && p !== '.') out.push(p);
  }
  return out.join('/');
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
      cells.set(`${p.col}${p.r}`, { ref: `${p.col}${p.r}`, r: p.r, c: p.c, val, display, formula });
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
  constructor(doc, styles) {
    this.doc = doc;
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
    if (!done && (type === 'money' || type === 'int')) {
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
    // Valeurs de formules en cache obsolètes : Excel recalcule à l'ouverture.
    for (const f of all(this.doc, 'f')) {
      const c = f.parentElement;
      kid(c, 'v')?.remove();
      c.removeAttribute('t');
    }
    kid(this.doc.documentElement, 'dimension')?.remove();
    return serialize(this.doc);
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

async function writeTemplate(book, sheetPath, tpl, report) {
  const doc = parseXml(await book.zip.file(sheetPath).async('string'));
  const w = new SheetWriter(doc, book.styles);
  for (const [ref, m] of Object.entries(tpl.cells || {})) {
    const def = fieldDef(m.field);
    if (def) w.set(ref, singleValue(m.field, report, m.value), def.type);
  }
  if (tpl.lineRow) {
    for (const [col, field] of Object.entries(tpl.lineCols || {})) {
      const def = fieldDef(field);
      if (!def) continue;
      const firstStyle = w.existingStyle(`${col}${tpl.lineRow}`);
      report.items.forEach((item, i) => w.set(`${col}${tpl.lineRow + i}`, lineValue(field, item), def.type, firstStyle));
    }
  }
  book.zip.file(sheetPath, w.finish());
  if (book.styles?.dirty) book.zip.file('xl/styles.xml', serialize(book.styles.doc));
  await forceRecalc(book);
}

function chunk(list, size) {
  const out = [];
  for (let i = 0; i < list.length; i += size) out.push(list.slice(i, i + size));
  return out.length ? out : [[]];
}

export async function fillTemplate(tpl, report, baseName) {
  const first = await openWorkbook(tpl.data);
  const sheet = first.sheets.find(s => s.name === tpl.sheet) || first.sheets[0];
  const content = await readSheet(first, sheet.path);
  const capacity = templateCapacity(tpl, content);
  const hasLines = tpl.lineRow && Object.keys(tpl.lineCols || {}).length;
  const parts = hasLines && report.items.length > capacity ? chunk(report.items, capacity) : [report.items];
  const files = [];
  for (let i = 0; i < parts.length; i++) {
    const book = i === 0 ? first : await openWorkbook(tpl.data);
    const part = parts.length > 1 ? { ...report, ...summarize(parts[i]) } : report;
    await writeTemplate(book, sheet.path, tpl, part);
    const blob = await book.zip.generateAsync({ type: 'blob', mimeType: XLSX_MIME, compression: 'DEFLATE' });
    files.push({ name: `${baseName}${parts.length > 1 ? ` (${i + 1} sur ${parts.length})` : ''}.xlsx`, blob, mime: XLSX_MIME });
  }
  return { files, capacity, split: parts.length > 1 };
}

/* ---------- Tableau standard ---------- */

const STYLES_XML = (currencySymbol) => `${XML_DECL}<styleSheet xmlns="${NS_MAIN}">
<numFmts count="2"><numFmt numFmtId="164" formatCode="#,##0.00\\ &quot;${escXml(currencySymbol)}&quot;"/><numFmt numFmtId="165" formatCode="dd/mm/yyyy"/></numFmts>
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
<cellXfs count="12">
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
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

const S = { title: 1, info: 2, header: 3, text: 4, date: 5, money: 6, int: 7, totalLabel: 8, totalMoney: 9, totalEmpty: 10, headerRight: 11 };
const SYMBOLS = { EUR: '€', USD: '$', GBP: '£', CHF: 'CHF', CAD: 'CA$' };

function strCell(ref, text, s) {
  return `<c r="${ref}" s="${s}" t="inlineStr"><is><t xml:space="preserve">${escXml(text)}</t></is></c>`;
}
function numCell(ref, n, s) {
  return `<c r="${ref}" s="${s}"><v>${n}</v></c>`;
}

export async function buildSimple(cfg, report, baseName) {
  const JSZip = await getJSZip();
  const columns = (cfg?.columns?.length ? cfg.columns : DEFAULT_SIMPLE_COLUMNS).filter(c => fieldDef(c.field));
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
      if (d.type === 'money') return numCell(ref, Number(v), S.money);
      if (d.type === 'int') return numCell(ref, Number(v), S.int);
      return strCell(ref, v, S.text);
    }));
  }
  const lastLine = r - 1;

  const firstMoney = defs.findIndex(d => d.type === 'money');
  if (firstMoney >= 0) {
    const totals = { net: report.net, vat: report.vat, total: report.total };
    addRow(defs.map((d, i) => {
      const col = numToCol(i + 1);
      const ref = `${col}${r}`;
      if (d.type === 'money') {
        const cached = totals[d.key] ?? 0;
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
