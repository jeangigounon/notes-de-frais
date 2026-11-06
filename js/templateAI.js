// Lecture d'un modèle Excel par Claude : il comprend la structure (formules, colonnes par type de dépense
// et par devise, taux de change…) et propose où écrire chaque information. L'utilisateur vérifie ensuite.

import { state } from './store.js';
import { CATEGORIES, LINE_FIELDS, SINGLE_FIELDS, PLACEHOLDERS } from './fields.js';
import { loadSdk, claudeClient, friendlyError } from './ai.js';
import { describeSheet, splitRef } from './xlsx.js';

const LINE_KEYS = LINE_FIELDS.map(f => f.key);
const SINGLE_KEYS = [...SINGLE_FIELDS.map(f => f.key), 'fxRate', 'fixed'];

const SCHEMA = {
  type: 'object',
  additionalProperties: false,
  properties: {
    summary: { type: 'string' },
    lineRow: { type: 'integer' },
    lineColumns: {
      type: 'array',
      items: {
        type: 'object',
        additionalProperties: false,
        properties: {
          col: { type: 'string' },
          field: { type: 'string', enum: LINE_KEYS },
          categories: { type: 'array', items: { type: 'string', enum: CATEGORIES } },
          others: { type: 'boolean' },
          currency: { type: 'string' },
        },
        required: ['col', 'field', 'categories', 'others', 'currency'],
      },
    },
    singleCells: {
      type: 'array',
      items: {
        type: 'object',
        additionalProperties: false,
        properties: {
          ref: { type: 'string' },
          field: { type: 'string', enum: SINGLE_KEYS },
          currency: { type: 'string' },
          value: { type: 'string' },
        },
        required: ['ref', 'field', 'currency', 'value'],
      },
    },
    fileName: { type: 'string' },
    warnings: { type: 'array', items: { type: 'string' } },
  },
  required: ['summary', 'lineRow', 'lineColumns', 'singleCells', 'fileName', 'warnings'],
};

const fieldList = list => list.map(f => `- ${f.key} : ${f.label}`).join('\n');

const SYSTEM = `Tu configures le remplissage automatique d'un modèle Excel de note de frais fourni par un client. L'application remplit le modèle à partir de justificatifs (tickets, factures) : une ligne du tableau des dépenses par justificatif, plus quelques cellules uniques (nom, période, totaux…). Tu reçois le contenu de la feuille et tu indiques où écrire chaque information.

Champs disponibles pour les colonnes du tableau (une valeur par justificatif) :
${fieldList(LINE_FIELDS)}

Champs disponibles pour les cellules uniques :
${fieldList(SINGLE_FIELDS)}
- fxRate : taux de change d'une devise vers l'euro (valeur en euros d'une unité de la devise) ; indique la devise dans currency.
- fixed : texte toujours identique ; indique-le dans value.

Catégories de dépenses de l'application : ${CATEGORIES.join(' ; ')}.

Règles :
- lineRow est la ligne où s'écrit le premier justificatif : la première ligne vide du tableau, sous les en-têtes et sous-en-têtes. Si une colonne de numéros pré-remplis (1, 2, 3…) existe, lineRow est la ligne du numéro 1. Les lignes suivantes sont remplies vers le bas.
- Ne place jamais un champ dans une cellule qui contient une formule, ni sur une ligne de totaux : le modèle recalcule lui-même ses totaux.
- Si les montants sont répartis en colonnes par type de dépense (repas, transport, logement, divers…), crée une colonne de montant par type, avec dans categories les catégories de l'application qui y vont. Mets others à true pour la colonne qui reçoit toutes les catégories non attribuées ailleurs (en général « divers » ou « autres ») ; il en faut une si le modèle a des colonnes par type.
- Si les montants sont répartis par devise, renseigne currency (code ISO : EUR, CHF, GBP, USD…) pour chaque colonne de montant. Des symboles comme €€€, ££ ou « CHF » désignent la devise. Sinon laisse currency vide.
- Utilise le champ total (montant TVAC dans la devise du ticket) pour les colonnes de montant ventilées par devise, et totalEur pour une colonne explicitement en euros convertis.
- Si le modèle contient des cellules de taux de change saisies à la main (pas des formules), mappe-les avec fxRate et la devise concernée. Regarde comment les formules les utilisent pour vérifier le sens du taux.
- Mappe aussi, si le modèle les prévoit, le nom du collaborateur, la période, la date, le numéro de justificatif, la description… N'invente pas de cellule : si rien n'est prévu, ne mets rien.
- summary : explique en 2 à 4 phrases, en français, comment fonctionne ce modèle (structure, formules, devises) et ce que tu as configuré.
- fileName : propose un nom de fichier (sans extension) inspiré du nom d'origine, avec ces variables si utile : ${PLACEHOLDERS.map(([k]) => k).join(', ')}. Garde les parties fixes telles quelles.
- warnings : points à vérifier par l'utilisateur (hypothèses, cas non couverts, incohérences relevées dans les formules). Liste vide si rien.`;

const COL_RX = /^[A-Z]{1,3}$/;

function sanitize(raw, content) {
  const maxCol = Math.max(content.maxCol + 2, 30);
  const lineCols = {};
  for (const c of raw.lineColumns || []) {
    const col = String(c.col || '').toUpperCase().replace(/[^A-Z]/g, '');
    if (!COL_RX.test(col) || !LINE_KEYS.includes(c.field) || splitRef(`${col}1`).c > maxCol) continue;
    const def = { field: c.field };
    const cats = (c.categories || []).filter(x => CATEGORIES.includes(x));
    if (cats.length) def.categories = cats;
    if (c.others) def.others = true;
    const cur = String(c.currency || '').toUpperCase().trim();
    if (/^[A-Z]{3}$/.test(cur)) def.currencies = [cur];
    lineCols[col] = def;
  }
  const cells = {};
  for (const s of raw.singleCells || []) {
    const p = splitRef(String(s.ref || '').toUpperCase());
    if (!p || !SINGLE_KEYS.includes(s.field)) continue;
    const ref = `${p.col}${p.r}`;
    if (content.cells.get(ref)?.formula) continue;
    if (s.field === 'fixed') { if (s.value?.trim()) cells[ref] = { field: 'fixed', value: s.value.trim() }; continue; }
    if (s.field === 'fxRate') {
      const cur = String(s.currency || '').toUpperCase().trim();
      if (/^[A-Z]{3}$/.test(cur)) cells[ref] = { field: 'fxRate', currency: cur };
      continue;
    }
    cells[ref] = { field: s.field };
  }
  const lineRow = Number.isInteger(raw.lineRow) && raw.lineRow > 0 && Object.keys(lineCols).length ? raw.lineRow : null;
  // Pas de cellule unique sur la ligne des justificatifs, dans une colonne du tableau.
  if (lineRow) for (const ref of Object.keys(cells)) { const p = splitRef(ref); if (p.r === lineRow && lineCols[p.col]) delete cells[ref]; }
  return {
    lineRow,
    lineCols: lineRow ? lineCols : {},
    cells,
    fileName: String(raw.fileName || '').trim(),
    summary: String(raw.summary || '').trim(),
    warnings: (raw.warnings || []).map(String).filter(Boolean),
  };
}

export const canAnalyzeTemplates = () => state.settings.aiMode === 'claude' && !!state.settings.anthropicKey;

export async function analyzeTemplate(content, { sheetName, fileName, otherSheets = [] }) {
  const Anthropic = await loadSdk();
  const client = claudeClient(Anthropic, state.settings.anthropicKey);
  const model = state.settings.model || 'claude-opus-5-5';
  const prompt = [
    `Nom du fichier d'origine : ${fileName}`,
    `Feuille à remplir : « ${sheetName} »${otherSheets.length ? ` (autres feuilles du classeur : ${otherSheets.join(', ')})` : ''}`,
    '',
    describeSheet(content),
  ].join('\n');
  const params = {
    model,
    max_tokens: 16000,
    system: SYSTEM,
    output_config: { effort: 'medium', format: { type: 'json_schema', schema: SCHEMA } },
    messages: [{ role: 'user', content: prompt }],
  };
  const attempts = [
    ...(model.includes('haiku') ? [] : [() => client.beta.messages.create({ ...params, betas: ['server-side-fallback-2026-07-01'], fallbacks: 'default' })]),
    () => client.messages.create(params),
    () => client.messages.create({
      ...params,
      output_config: { effort: 'medium' },
      system: `${SYSTEM}\n\nRéponds uniquement par un objet JSON avec les clés : ${SCHEMA.required.join(', ')}.`,
    }),
  ];
  let res;
  for (let i = 0; i < attempts.length && !res; i++) {
    try {
      res = await attempts[i]();
    } catch (err) {
      const detail = err?.error?.error?.message || err?.message || '';
      const retryable = err instanceof Anthropic.BadRequestError && /output_config|format|schema|fallback|beta/i.test(detail);
      if (!retryable || i === attempts.length - 1) throw friendlyError(err, Anthropic);
    }
  }
  if (res.stop_reason === 'refusal') throw new Error("Claude n'a pas pu analyser ce modèle.");
  if (res.stop_reason === 'max_tokens') throw new Error('Analyse incomplète : modèle trop volumineux.');
  const text = res.content.filter(b => b.type === 'text').map(b => b.text).join('');
  return sanitize(JSON.parse(text.slice(text.indexOf('{'), text.lastIndexOf('}') + 1)), content);
}
