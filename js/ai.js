// Reconnaissance des justificatifs : Claude (vision) ou OCR local gratuit (Tesseract) en secours.

import { state, getExpense, saveExpense } from './store.js';
import { CATEGORIES, DOC_TYPES, PAYMENTS, num } from './fields.js';
import { blobToBase64, todayISO } from './ui.js';
import { CDN, loadScript } from './lib.js';

export const MODELS = [
  { id: 'claude-opus-5-5', label: 'Claude Opus 5.5', hint: 'Le plus précis · environ 2 centimes par ticket' },
  { id: 'claude-sonnet-5-5', label: 'Claude Sonnet 5.5', hint: 'Rapide · environ 1 centime par ticket' },
  { id: 'claude-haiku-5-5', label: 'Claude Haiku 5.5', hint: 'Le plus économique · moins de 0,2 centime' },
];

const nullableNumber = { anyOf: [{ type: 'number' }, { type: 'null' }] };

const SCHEMA = {
  type: 'object',
  additionalProperties: false,
  properties: {
    isReceipt: { type: 'boolean' },
    docType: { type: 'string', enum: DOC_TYPES },
    merchant: { type: 'string' },
    date: { type: 'string', description: 'AAAA-MM-JJ, ou chaîne vide' },
    description: { type: 'string' },
    category: { type: 'string', enum: CATEGORIES },
    currency: { type: 'string', description: 'Code ISO 4217' },
    total: nullableNumber,
    net: nullableNumber,
    vat: nullableNumber,
    vatRate: { type: 'string' },
    invoiceNo: { type: 'string' },
    payment: { type: 'string', enum: [...PAYMENTS, ''] },
    confidence: { type: 'string', enum: ['haute', 'moyenne', 'faible'] },
    remark: { type: 'string' },
  },
  required: ['isReceipt', 'docType', 'merchant', 'date', 'description', 'category', 'currency', 'total', 'net', 'vat', 'vatRate', 'invoiceNo', 'payment', 'confidence', 'remark'],
};

const SYSTEM = `Tu lis des justificatifs de dépenses professionnelles (tickets de caisse, factures, notes de restaurant, reçus de carte, titres de transport), photographiés ou scannés, le plus souvent belges et rédigés en français, en néerlandais ou en anglais. Tu en extrais les informations nécessaires à une note de frais.

- total : le montant réellement payé, TVA comprise (TVAC, TTC, incl. BTW). Ignore la monnaie rendue et les pourboires non inclus.
- net : le total hors TVA (HTVA, HT, excl. BTW). vat : le montant total de TVA. S'il y a plusieurs taux, additionne les montants et liste les taux dans vatRate (ex. « 6 % + 21 % »). Si la TVA n'apparaît pas, mets null.
- date : la date de l'achat ou de la prestation (pas l'échéance), au format AAAA-MM-JJ. Les dates belges s'écrivent JJ/MM/AAAA.
- merchant : le nom commercial tel qu'on le dirait (« Delhaize », « SNCB », « Hôtel Amigo »).
- description : quelques mots utiles pour la note de frais (« Déjeuner – 2 couverts », « Plein diesel », « Train Bruxelles – Liège »).
- payment : le moyen de paiement s'il est visible, sinon une chaîne vide.
- isReceipt : false si le document n'est pas un justificatif de dépense.
- remark : signale brièvement un doute (photo floue, montant peu lisible, plusieurs tickets sur la photo), sinon une chaîne vide.

Si une information est absente ou illisible, laisse une chaîne vide (ou null pour un montant) plutôt que de deviner.`;

/* ---------- Claude ---------- */

let sdkPromise;
const loadSdk = () => (sdkPromise ||= import(CDN.anthropic).then(m => m.default));

function friendlyError(err, Anthropic) {
  const detail = err?.error?.error?.message || err?.message || String(err);
  if (err instanceof Anthropic.AuthenticationError) return new Error('Clé API Claude refusée. Vérifiez-la dans Réglages.');
  if (err instanceof Anthropic.PermissionDeniedError) return new Error("Cette clé API n'a pas accès à ce modèle.");
  if (err instanceof Anthropic.RateLimitError) return new Error('Trop de demandes en même temps. Réessayez dans un instant.');
  if (err instanceof Anthropic.BadRequestError) {
    if (/credit balance/i.test(detail)) return new Error('Crédit API Claude épuisé : rechargez-le sur console.anthropic.com.');
    return new Error(`Requête refusée : ${detail}`);
  }
  if (err instanceof Anthropic.APIConnectionError) return new Error("Impossible de joindre l'API Claude. Vérifiez la connexion.");
  if (err instanceof Anthropic.APIError) return new Error(`Erreur de l'API Claude (${err.status ?? '?'}) : ${detail}`);
  return err instanceof Error ? err : new Error(detail);
}

function claudeClient(Anthropic, apiKey) {
  return new Anthropic({ apiKey, dangerouslyAllowBrowser: true, maxRetries: 2 });
}

async function claudeExtract(exp) {
  const Anthropic = await loadSdk();
  const client = claudeClient(Anthropic, state.settings.anthropicKey);
  const model = state.settings.model || 'claude-opus-5-5';
  const data = await blobToBase64(exp.file);
  const media = exp.mime === 'application/pdf'
    ? { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data } }
    : { type: 'image', source: { type: 'base64', media_type: exp.mime, data } };

  const params = {
    model,
    max_tokens: 16000,
    system: SYSTEM,
    output_config: { effort: 'low', format: { type: 'json_schema', schema: SCHEMA } },
    messages: [{
      role: 'user',
      content: [media, { type: 'text', text: `Analyse ce justificatif. Nous sommes le ${todayISO()} (utile si l'année est abrégée).` }],
    }],
  };

  const plain = {
    ...params,
    output_config: { effort: 'low' },
    system: `${SYSTEM}\n\nRéponds uniquement par un objet JSON avec exactement ces clés : ${SCHEMA.required.join(', ')}.`,
  };
  const attempts = [
    // Opus et Sonnet : repli automatique côté serveur si un filtre de sécurité refuse la demande.
    ...(model.includes('haiku') ? [] : [() => client.beta.messages.create({ ...params, betas: ['server-side-fallback-2026-07-01'], fallbacks: 'default' })]),
    () => client.messages.create(params),
    () => client.messages.create(plain),
  ];

  let res;
  for (let i = 0; i < attempts.length && !res; i++) {
    try {
      res = await attempts[i]();
    } catch (err) {
      // Option refusée par l'API (repli, format structuré) : on retente avec une requête plus simple.
      const detail = err?.error?.error?.message || err?.message || '';
      const retryable = err instanceof Anthropic.BadRequestError && /output_config|format|schema|fallback|beta/i.test(detail);
      if (!retryable || i === attempts.length - 1) throw friendlyError(err, Anthropic);
    }
  }
  if (res.stop_reason === 'refusal') throw new Error("Claude n'a pas pu analyser ce document.");
  if (res.stop_reason === 'max_tokens') throw new Error('Réponse incomplète, réessayez.');
  const text = res.content.filter(b => b.type === 'text').map(b => b.text).join('');
  const json = text.slice(text.indexOf('{'), text.lastIndexOf('}') + 1);
  if (!json) throw new Error('Réponse illisible, réessayez.');
  return normalize(JSON.parse(json), 'claude');
}

export async function testClaudeKey(apiKey, model) {
  const Anthropic = await loadSdk();
  const client = claudeClient(Anthropic, apiKey);
  try {
    await client.models.retrieve(model);
    return true;
  } catch (err) {
    throw friendlyError(err, Anthropic);
  }
}

/* ---------- OCR local ---------- */

async function localExtract(exp) {
  if (!exp.mime.startsWith('image/') || exp.mime === 'image/heic') {
    throw new Error("L'analyse locale ne lit que les photos. Ajoutez une clé Claude dans Réglages pour les PDF.");
  }
  await loadScript(CDN.tesseract);
  const worker = await window.Tesseract.createWorker('fra');
  try {
    const { data } = await worker.recognize(exp.file);
    return parseReceiptText(data.text || '');
  } finally {
    await worker.terminate();
  }
}

const AMOUNT_RX = /(?<![\d.,])(\d{1,3}(?:[ . ]\d{3})+|\d+)\s?[.,]\s?(\d{2})(?![\d%])/g;
const amountsIn = line => [...line.matchAll(AMOUNT_RX)]
  .map(m => parseFloat(m[1].replace(/[ . ]/g, '') + '.' + m[2]))
  .filter(n => n > 0 && n < 100000);

const CATEGORY_HINTS = [
  [/restaurant|resto|brasserie|bistro|café|cafe|pizz|sushi|traiteur|taverne|couverts?|lunch|déjeuner|diner|dîner/i, 'Restaurant & repas'],
  [/carbur|diesel|essence|benzine|gasoil|euro ?95|euro ?98|\btotal ?energies\b|\bq8\b|\besso\b|\bshell\b|texaco|lukoil|gulf|dats ?24|\blitres?\b/i, 'Carburant'],
  [/parking|interparking|q-park|indigo|péage|tolhuis|stationnement/i, 'Parking & péage'],
  [/taxi|uber|bolt|heetch|vtc/i, 'Taxi & VTC'],
  [/sncb|nmbs|\bstib\b|\bmivb\b|\btec\b|de lijn|eurostar|thalys|train|métro|tram/i, 'Train & transports en commun'],
  [/brussels airlines|ryanair|easyjet|lufthansa|\bklm\b|air france|vol |flight|boarding/i, 'Avion'],
  [/h[oô]tel|nuitée|booking\.com|airbnb|chambre|overnachting/i, 'Hébergement'],
  [/hertz|avis|europcar|sixt|location de v/i, 'Location de véhicule'],
  [/fnac|mediamarkt|coolblue|papeterie|bureau|fourniture|action|brico|ikea/i, 'Fournitures & matériel'],
];

export function parseReceiptText(text) {
  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  const all = lines.flatMap(amountsIn);

  const TOTAL_RX = /(total|totaal|à payer|a payer|te betalen|net à payer|montant|bedrag|tvac|ttc|incl)/i;
  const NOT_TOTAL_RX = /(sous[- ]?total|subtotal|htva|excl|hors|tva\s*\d|btw\s*\d|rendu|change|terug)/i;
  const totals = lines.filter(l => TOTAL_RX.test(l) && !NOT_TOTAL_RX.test(l)).flatMap(amountsIn);
  const total = totals.length ? Math.max(...totals) : all.length ? Math.max(...all) : null;

  let vatRate = '';
  let vat = null;
  for (const l of lines) {
    if (!/(tva|btw|vat|t\.v\.a)/i.test(l)) continue;
    const rate = /(\d{1,2}(?:[.,]\d{1,2})?)\s?%/.exec(l);
    if (rate && !vatRate) vatRate = rate[1].replace('.', ',') + ' %';
    const nums = amountsIn(l).filter(n => total == null || n < total);
    if (nums.length && vat == null) vat = Math.min(...nums);
  }

  let date = '';
  for (const m of text.matchAll(/(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2,4})/g)) {
    let [, d, mo, y] = m.map(Number);
    if (y < 100) y += 2000;
    if (d >= 1 && d <= 31 && mo >= 1 && mo <= 12 && y >= 2000 && y <= 2100) {
      date = `${y}-${String(mo).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
      break;
    }
  }

  const merchantLine = lines.find(l => /[a-zà-ü]{3,}/i.test(l) && !/(ticket|facture|factuur|tva|btw|tél|tel\.|www|@|date|caisse|kassa)/i.test(l)) || '';
  const merchant = merchantLine.replace(/[^\p{L}\p{N}&'’ .-]/gu, '').trim().slice(0, 60);
  const category = CATEGORY_HINTS.find(([rx]) => rx.test(text))?.[1] || '';
  const payment = /bancontact|maestro|payconiq|debit|débit|carte|card|\bcb\b/i.test(text) ? 'Carte bancaire'
    : /visa|mastercard|amex|american express/i.test(text) ? 'Carte de crédit'
      : /esp[eè]ces|cash|contant/i.test(text) ? 'Espèces' : '';

  return normalize({
    isReceipt: true,
    docType: /facture|factuur|invoice/i.test(text) ? 'Facture' : 'Ticket de caisse',
    merchant, date, description: '', category, currency: /\bUSD\b|\$/.test(text) ? 'USD' : 'EUR',
    total, net: null, vat, vatRate, invoiceNo: '', payment,
    confidence: 'faible',
    remark: 'Lecture locale approximative : vérifiez les montants.',
  }, 'local');
}

/* ---------- Normalisation ---------- */

function normalize(r, engine) {
  const str = v => (typeof v === 'string' ? v.trim() : '');
  const date = /^\d{4}-\d{2}-\d{2}$/.test(str(r.date)) ? str(r.date) : '';
  const currency = /^[A-Z]{3}$/.test(str(r.currency).toUpperCase()) ? str(r.currency).toUpperCase() : 'EUR';
  return {
    isReceipt: r.isReceipt !== false,
    docType: DOC_TYPES.includes(r.docType) ? r.docType : '',
    merchant: str(r.merchant),
    date,
    description: str(r.description),
    category: CATEGORIES.includes(r.category) ? r.category : '',
    currency,
    total: num(r.total),
    net: num(r.net),
    vat: num(r.vat),
    vatRate: str(r.vatRate),
    invoiceNo: str(r.invoiceNo),
    payment: PAYMENTS.includes(r.payment) ? r.payment : '',
    confidence: r.confidence || 'moyenne',
    remark: str(r.remark),
    engine,
  };
}

/* ---------- File d'attente d'analyse ---------- */

const queue = [];
let running = false;

export function analyzeExpense(id, { overwrite = false } = {}) {
  if (!queue.some(job => job.id === id)) queue.push({ id, overwrite });
  const exp = getExpense(id);
  if (exp && exp.aiStatus !== 'pending') {
    exp.aiStatus = 'pending';
    exp.aiError = '';
    saveExpense(exp);
  }
  pump();
}

async function pump() {
  if (running) return;
  running = true;
  while (queue.length) await runJob(queue.shift());
  running = false;
}

async function runJob({ id, overwrite }) {
  const exp = getExpense(id);
  if (!exp) return;
  const useClaude = state.settings.aiMode === 'claude' && !!state.settings.anthropicKey;
  let result = null;
  let error = '';
  try {
    result = useClaude ? await claudeExtract(exp) : await localExtract(exp);
  } catch (err) {
    console.error(err);
    error = err.message || String(err);
  }
  const fresh = getExpense(id);
  if (!fresh) return;
  if (result) {
    for (const [k, v] of Object.entries(result)) {
      if (!(k in fresh.data) || v === '' || v == null) continue;
      const cur = fresh.data[k];
      const isEmpty = cur == null || cur === '' || (k === 'currency' && cur === 'EUR');
      if (overwrite || isEmpty) fresh.data[k] = v;
    }
    fresh.ai = {
      engine: result.engine,
      model: useClaude ? state.settings.model : 'tesseract',
      confidence: result.confidence,
      remark: result.remark,
      isReceipt: result.isReceipt,
      at: Date.now(),
    };
    fresh.aiStatus = 'done';
    fresh.aiError = '';
  } else {
    fresh.aiStatus = 'error';
    fresh.aiError = error;
  }
  await saveExpense(fresh);
}

// Reprend les analyses interrompues (application fermée pendant l'analyse).
export function resumePending() {
  for (const exp of state.expenses) if (exp.aiStatus === 'pending') analyzeExpense(exp.id);
}
