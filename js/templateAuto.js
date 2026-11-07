// Lecture automatique (sans IA) de la structure d'un modèle Excel de note de frais :
// en-têtes de colonnes, colonnes par type de dépense (repas, transport, logement, divers…),
// sous-colonnes par devise (€, CHF, £…) et cellules de taux de change.
// Sert quand Claude n'est pas disponible, et de garde-fou quand son analyse oublie la ventilation.

import { CATEGORIES, lineColDef } from './fields.js';
import { numToCol } from './xlsx.js';

const norm = s => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();

const TRANSPORT = ['Carburant', 'Parking & péage', 'Taxi & VTC', 'Train & transports en commun', 'Avion', 'Location de véhicule'];
const GROUPS = [
  { rx: /\b(food|repas|restaurants?|restauration|meals?|nourriture|dejeuners?|diners?|horeca)\b/, categories: ['Restaurant & repas', 'Représentation & cadeaux'] },
  { rx: /\b(transports?|deplacements?|travel|voyages?|trains?|avions?|taxis?|carburants?|fuel|mobilite|kilometres?)\b/, categories: TRANSPORT },
  { rx: /\b(lodging|hebergements?|hotels?|logements?|accommodation|nuitees?|sejours?)\b/, categories: ['Hébergement'] },
  { rx: /\b(fournitures?|materiel|supplies|equipements?)\b/, categories: ['Fournitures & matériel'] },
  { rx: /\b(telecom|telephonie|logiciels?|abonnements?|software)\b/, categories: ['Abonnements & logiciels'] },
  { rx: /\b(representations?|cadeaux?|gifts?|entertainment)\b/, categories: ['Représentation & cadeaux'] },
  { rx: /\b(parkings?|peages?)\b/, categories: ['Parking & péage'] },
  { rx: /\b(divers|autres?|others?|misc|miscellaneous|sundries|varia)\b/, others: true },
];

const CURRENCIES = [
  [/(^|[^a-z])(eur|euros?)([^a-z]|$)|€/, 'EUR'],
  [/(^|[^a-z])(chf|sfr|francs? suisses?|swiss)([^a-z]|$)/, 'CHF'],
  [/(^|[^a-z])(gbp|livres?|pounds?|sterling)([^a-z]|$)|£/, 'GBP'],
  [/(^|[^a-z])(usd|dollars?|us\$)([^a-z]|$)|\$/, 'USD'],
];
const currencyOf = text => CURRENCIES.find(([rx]) => rx.test(norm(text)))?.[1] || '';

// Ordre important : « montant TVA » avant « montant », « HTVA » avant « TVA ».
const FIELDS = [
  [/^(date|jour|day)\b|\bdate\b/, 'date'],
  [/\bjustificatif|\b(n[°o]|num(ero)?|no\.?|#)\s*(de\s*)?(piece|ticket|facture|recu)?$/, 'index'],
  [/\b(htva|hors tva|h\.t\.?|ht|net|excl)\b/, 'net'],
  [/^(tva|btw|vat)\b|\bmontant (de )?(la )?tva\b/, 'vat'],
  [/\b(fournisseur|commercant|etablissement|magasin|enseigne|vendor|merchant|supplier)\b/, 'merchant'],
  [/\b(designation|libelle|description|objet|motif|lieu|details?)\b/, 'label'],
  [/\b(categorie|category|type de depense|nature)\b/, 'category'],
  [/\b(devise|currency|monnaie)\b/, 'currency'],
  [/\b(taux|rate|cours)\b/, 'rate'],
  [/\b(paiement|payment)\b/, 'payment'],
  [/\b(remarques?|commentaires?|notes?|observations?)\b/, 'notes'],
  [/\b(tvac|ttc|incl|montant|amount|totale?|prix|cout)\b/, 'total'],
];
const fieldOf = text => FIELDS.find(([rx]) => rx.test(norm(text)))?.[1] || '';
const groupOf = text => (norm(text).length <= 40 ? GROUPS.find(g => g.rx.test(norm(text))) : null);

const SINGLES = [
  [/^(nom|name|collaborateur|employe|employee|demandeur|beneficiaire|consultant)\b/, 'profile.name'],
  [/^(periode|mois|month|period)\b/, 'report.period'],
  [/^(societe|departement|service|company)\b/, 'profile.company'],
  [/^(iban|compte bancaire|n° de compte)\b/, 'profile.iban'],
  [/^(e-?mail|courriel)\b/, 'profile.email'],
  [/^date\b/, 'report.today'],
];

const isText = c => c && typeof c.val === 'string' && c.val.trim() && !c.formula;

export function detectStructure(content) {
  const cells = content.cells;
  const at = (r, c) => cells.get(`${numToCol(c)}${r}`);
  const firstFormulaRow = Math.min(...[...cells.values()].filter(c => c.formula).map(c => c.r), Infinity);
  const top = Math.min(40, isFinite(firstFormulaRow) ? firstFormulaRow - 1 : content.maxRow);

  // 1. Table des taux : un code de devise suivi d'un nombre saisi à la main (ex. « CHF | 0,9145 »),
  //    juste à côté ou sous un en-tête « Taux ».
  const rateHeader = (r, c) => {
    for (let k = r - 1; k >= Math.max(1, r - 3); k--) {
      const h = at(k, c);
      if (isText(h)) return /\b(taux|rate|cours|change)\b/.test(norm(h.val));
    }
    return false;
  };
  const fx = {};
  const rateCols = new Set();
  for (const c of cells.values()) {
    if (c.r > top || !isText(c) || c.val.trim().length > 8) continue;
    const cur = currencyOf(c.val);
    if (!cur || cur === 'EUR') continue;
    for (let dc = 1; dc <= 2; dc++) {
      const n = at(c.r, c.c + dc);
      if (isText(n)) break;
      const isRate = n && typeof n.val === 'number' && !n.formula && !n.isDate && n.val > 0.001 && n.val < 1000;
      if (isRate && (rateHeader(n.r, n.c) || (dc === 1 && n.val !== 1))) {
        fx[n.ref] = cur;
        rateCols.add(c.c);
        rateCols.add(n.c);
        break;
      }
    }
  }

  // 2. Ligne d'en-têtes : celle qui contient le plus de libellés reconnus.
  const scoreRow = r => [...cells.values()].filter(c => c.r === r && isText(c) && !rateCols.has(c.c) && (fieldOf(c.val) || groupOf(c.val))).length;
  let headerRow = 0;
  let best = 1;
  for (let r = 1; r <= top; r++) { const s = scoreRow(r); if (s > best) { best = s; headerRow = r; } }
  if (!headerRow) return null;
  const rowCells = r => [...cells.values()].filter(c => c.r === r && isText(c) && !rateCols.has(c.c)).sort((a, b) => a.c - b.c);
  const subRow = rowCells(headerRow + 1).filter(c => currencyOf(c.val) && c.val.trim().length <= 12).length >= 2 ? headerRow + 1 : 0;
  const mergeAt = (r, c) => content.merges.find(m => m.r1 === r && m.c1 === c);

  const lineCols = {};
  const groups = [];
  for (const cell of rowCells(headerRow)) {
    const g = groupOf(cell.val);
    if (g && !fieldOf(cell.val).match(/^(date|index|label|merchant)$/)) {
      const m = mergeAt(cell.r, cell.c);
      let c2 = m ? m.c2 : cell.c;
      if (!m && subRow) while (!at(headerRow, c2 + 1) && at(subRow, c2 + 1) && currencyOf(at(subRow, c2 + 1).val)) c2++;
      const filter = g.others ? { others: true } : { categories: [...g.categories] };
      const cols = [];
      for (let c = cell.c; c <= c2; c++) {
        const cur = subRow ? currencyOf(at(subRow, c)?.val) : currencyOf(cell.val);
        if (subRow && !cur && c2 > cell.c) continue;
        lineCols[numToCol(c)] = { field: 'total', ...filter, ...(cur ? { currencies: [cur] } : {}) };
        cols.push(numToCol(c));
      }
      groups.push({ name: cell.val.trim(), cols, filter });
      continue;
    }
    const field = fieldOf(cell.val);
    if (!field) continue;
    const def = { field };
    const cur = currencyOf(cell.val) || (subRow ? currencyOf(at(subRow, cell.c)?.val) : '');
    if (['total', 'net', 'vat'].includes(field) && cur) def.currencies = [cur];
    lineCols[numToCol(cell.c)] = def;
  }
  if (!Object.keys(lineCols).length) return null;

  // 3. Avec des colonnes par type, une colonne « total » générique (ex. « TOTAUX ») ne reçoit rien.
  if (groups.length) {
    for (const [col, d] of Object.entries(lineCols)) if (d.field === 'total' && !d.categories && !d.others) delete lineCols[col];
    // Une catégorie ne va que dans une colonne : la plus spécifique.
    for (const cat of CATEGORIES) {
      const owners = groups.filter(g => g.filter.categories?.includes(cat));
      if (owners.length < 2) continue;
      const keep = owners.reduce((a, b) => (a.filter.categories.length <= b.filter.categories.length ? a : b));
      for (const g of owners) if (g !== keep) g.filter.categories = g.filter.categories.filter(x => x !== cat);
      for (const g of owners) for (const col of g.cols) if (lineCols[col]) lineCols[col].categories = [...g.filter.categories];
    }
  }
  const hasMerchant = Object.values(lineCols).some(d => d.field === 'merchant');
  for (const d of Object.values(lineCols)) if (d.field === 'label' && hasMerchant) d.field = 'description';

  // 4. Première ligne du tableau : sous les en-têtes, ou au numéro « 1 » d'une colonne numérotée.
  let lineRow = (subRow || headerRow) + 1;
  const indexCol = Object.entries(lineCols).find(([, d]) => d.field === 'index')?.[0];
  if (indexCol) {
    for (let r = lineRow; r <= lineRow + 6; r++) {
      if (cells.get(`${indexCol}${r}`)?.val === 1) { lineRow = r; break; }
    }
  }

  // 5. Cellules uniques au-dessus du tableau : « Nom : … », « Période : … », taux de change.
  const singles = {};
  for (const [ref, cur] of Object.entries(fx)) singles[ref] = { field: 'fxRate', currency: cur };
  for (const c of cells.values()) {
    if (c.r >= headerRow || !isText(c) || rateCols.has(c.c)) continue;
    const label = norm(c.val).replace(/[:.]+$/, '').trim();
    const field = SINGLES.find(([rx]) => rx.test(label))?.[1];
    if (!field) continue;
    const m = mergeAt(c.r, c.c);
    for (let dc = (m ? m.c2 - c.c : 0) + 1; dc <= 4; dc++) {
      const target = at(c.r, c.c + dc);
      if (!target || (target.display === '' && !target.formula)) { singles[`${numToCol(c.c + dc)}${c.r}`] = { field }; break; }
      if (target.formula) break;
    }
  }

  const curs = [...new Set(Object.values(lineCols).flatMap(d => d.currencies || []))];
  const parts = [`tableau des dépenses à partir de la ligne ${lineRow}`];
  if (groups.length) parts.push(`montants répartis par type (${groups.map(g => g.name).join(', ')})${curs.length > 1 ? ` et par devise (${curs.join(', ')})` : ''}`);
  if (Object.keys(fx).length) parts.push(`taux de change en ${Object.entries(fx).map(([ref, cur]) => `${ref} (${cur})`).join(' et ')}, mis à jour au taux BCE`);
  const warnings = [];
  if (groups.length && !groups.some(g => g.filter.others)) warnings.push('Aucune colonne « divers » trouvée : les dépenses d’autres catégories n’auront pas de colonne. Indiquez-en une via les conditions d’une colonne.');

  return {
    lineRow,
    lineCols,
    cells: singles,
    groups,
    summary: `Lecture automatique : ${parts.join(' ; ')}.`,
    warnings,
  };
}

// La configuration met-elle tous les montants dans une seule colonne alors que le modèle les ventile ?
export function mappingMissesGroups(tpl, content) {
  const detected = detectStructure(content);
  if (!detected?.groups.length) return null;
  const mapped = Object.values(tpl.lineCols || {}).map(lineColDef);
  const filtered = mapped.some(d => d.categories?.length || d.others);
  return filtered ? null : detected;
}
