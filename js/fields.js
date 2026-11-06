// Vocabulaire métier : catégories, champs disponibles pour les fichiers Excel, calcul d'une note de frais.

import { todayISO, parseISO, safeFileName } from './ui.js';

export const CATEGORIES = [
  'Restaurant & repas',
  'Carburant',
  'Parking & péage',
  'Taxi & VTC',
  'Train & transports en commun',
  'Avion',
  'Hébergement',
  'Location de véhicule',
  'Fournitures & matériel',
  'Abonnements & logiciels',
  'Représentation & cadeaux',
  'Autre',
];

export const DOC_TYPES = ['Ticket de caisse', 'Facture', 'Note de restaurant', 'Reçu de paiement par carte', 'Titre de transport', 'Autre'];
export const PAYMENTS = ['Carte bancaire', 'Carte de crédit', 'Carte société', 'Espèces', 'Virement', 'Autre'];
export const CURRENCIES = ['EUR', 'USD', 'GBP', 'CHF', 'CAD'];

export const CLIENT_COLORS = {
  blue: '#0a84ff',
  indigo: '#5e5ce6',
  purple: '#bf5af2',
  pink: '#ff375f',
  red: '#ff453a',
  orange: '#ff9f0a',
  green: '#30d158',
  teal: '#40c8e0',
  graphite: '#8e8e93',
};

// Champs répétés : une ligne du tableau par justificatif.
export const LINE_FIELDS = [
  { key: 'index', label: 'N°', type: 'int', width: 6 },
  { key: 'date', label: 'Date', type: 'date', width: 12 },
  { key: 'merchant', label: 'Commerçant / fournisseur', type: 'text', width: 28 },
  { key: 'description', label: 'Description', type: 'text', width: 36 },
  { key: 'category', label: 'Catégorie', type: 'text', width: 24 },
  { key: 'docType', label: 'Type de justificatif', type: 'text', width: 20 },
  { key: 'net', label: 'Montant HTVA', type: 'money', width: 13 },
  { key: 'vat', label: 'Montant TVA', type: 'money', width: 13 },
  { key: 'vatRate', label: 'Taux de TVA', type: 'text', width: 11 },
  { key: 'total', label: 'Montant TVAC', type: 'money', width: 14 },
  { key: 'currency', label: 'Devise', type: 'text', width: 8 },
  { key: 'invoiceNo', label: 'N° de facture / ticket', type: 'text', width: 18 },
  { key: 'payment', label: 'Moyen de paiement', type: 'text', width: 18 },
  { key: 'notes', label: 'Remarque', type: 'text', width: 30 },
  { key: 'fileName', label: 'Nom du justificatif joint', type: 'text', width: 40 },
];

// Champs uniques : une seule cellule dans le fichier.
export const SINGLE_FIELDS = [
  { key: 'profile.name', label: 'Mon nom', type: 'text', group: 'Moi' },
  { key: 'profile.email', label: 'Mon e-mail', type: 'text', group: 'Moi' },
  { key: 'profile.company', label: 'Ma société / service', type: 'text', group: 'Moi' },
  { key: 'profile.iban', label: 'Mon IBAN', type: 'text', group: 'Moi' },
  { key: 'client.name', label: 'Nom du destinataire', type: 'text', group: 'Destinataire' },
  { key: 'client.company', label: 'Société du destinataire', type: 'text', group: 'Destinataire' },
  { key: 'report.today', label: 'Date du jour', type: 'date', group: 'Note de frais' },
  { key: 'report.period', label: 'Période (ex. novembre 2026)', type: 'text', group: 'Note de frais' },
  { key: 'report.month', label: 'Mois (ex. 11/2026)', type: 'text', group: 'Note de frais' },
  { key: 'report.from', label: 'Date de la 1re dépense', type: 'date', group: 'Note de frais' },
  { key: 'report.to', label: 'Date de la dernière dépense', type: 'date', group: 'Note de frais' },
  { key: 'report.count', label: 'Nombre de justificatifs', type: 'int', group: 'Note de frais' },
  { key: 'report.net', label: 'Total HTVA', type: 'money', group: 'Note de frais' },
  { key: 'report.vat', label: 'Total TVA', type: 'money', group: 'Note de frais' },
  { key: 'report.total', label: 'Total TVAC', type: 'money', group: 'Note de frais' },
];

export const FIXED_FIELD = { key: 'fixed', label: 'Texte fixe', type: 'text' };

export const DEFAULT_SIMPLE_COLUMNS = [
  { field: 'index', label: 'N°' },
  { field: 'date', label: 'Date' },
  { field: 'merchant', label: 'Commerçant' },
  { field: 'description', label: 'Description' },
  { field: 'category', label: 'Catégorie' },
  { field: 'net', label: 'HTVA' },
  { field: 'vat', label: 'TVA' },
  { field: 'total', label: 'TVAC' },
];

export const PLACEHOLDERS = [
  ['{nom}', 'votre nom'],
  ['{periode}', 'la période couverte'],
  ['{total}', 'le total TVAC'],
  ['{nb}', 'le nombre de justificatifs'],
  ['{client}', 'le nom du destinataire'],
  ['{date}', 'la date du jour'],
];

export function fieldDef(key) {
  if (key === 'fixed') return FIXED_FIELD;
  return LINE_FIELDS.find(f => f.key === key) || SINGLE_FIELDS.find(f => f.key === key) || null;
}

export function emptyData() {
  return {
    docType: '', merchant: '', date: '', description: '', category: '', currency: 'EUR',
    total: null, net: null, vat: null, vatRate: '', invoiceNo: '', payment: '', notes: '',
  };
}

/* ---------- Montants ---------- */

export const round2 = n => Math.round(n * 100) / 100;

export function num(v) {
  if (v === '' || v == null) return null;
  if (typeof v === 'number') return isFinite(v) ? v : null;
  const s = String(v).replace(/[\s €$£]/g, '');
  // « 1.234,56 » ou « 1,234.56 » ou « 12,5 »
  const normalized = /,\d{1,2}$/.test(s) ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
  const n = parseFloat(normalized);
  return isFinite(n) ? n : null;
}

// Complète HTVA / TVA / TVAC quand deux des trois sont connus.
export function amounts(d) {
  let total = num(d.total);
  let net = num(d.net);
  let vat = num(d.vat);
  if (total == null && net != null && vat != null) total = round2(net + vat);
  if (vat == null && total != null && net != null) vat = round2(total - net);
  if (net == null && total != null && vat != null) net = round2(total - vat);
  return { total, net, vat };
}

/* ---------- Note de frais ---------- */

const monthLong = d => d.toLocaleDateString('fr-BE', { month: 'long', year: 'numeric' });

export function periodLabel(dates) {
  if (!dates.length) return monthLong(new Date());
  const a = parseISO(dates[0]);
  const b = parseISO(dates[dates.length - 1]);
  if (a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth()) return monthLong(a);
  if (a.getFullYear() === b.getFullYear()) return `${a.toLocaleDateString('fr-BE', { month: 'long' })} – ${monthLong(b)}`;
  return `${monthLong(a)} – ${monthLong(b)}`;
}

function extFor(mime, name) {
  if (mime === 'application/pdf') return 'pdf';
  if (mime === 'image/png') return 'png';
  if (mime === 'image/jpeg') return 'jpg';
  const m = /\.([a-z0-9]{2,5})$/i.exec(name || '');
  return m ? m[1].toLowerCase() : 'jpg';
}

export function summarize(items) {
  const sum = k => round2(items.reduce((s, it) => s + (it[k] || 0), 0));
  const dates = items.map(i => i.date).filter(Boolean).sort();
  const first = parseISO(dates[0] || todayISO());
  return {
    items,
    count: items.length,
    total: sum('total'),
    net: sum('net'),
    vat: sum('vat'),
    from: dates[0] || '',
    to: dates[dates.length - 1] || '',
    period: periodLabel(dates),
    month: `${String(first.getMonth() + 1).padStart(2, '0')}/${first.getFullYear()}`,
  };
}

export function buildReport(client, expenses, profile) {
  const sorted = [...expenses].sort((a, b) => (a.data.date || '9999').localeCompare(b.data.date || '9999') || a.createdAt - b.createdAt);
  const items = sorted.map((exp, i) => {
    const n = String(i + 1).padStart(2, '0');
    const label = [exp.data.merchant, exp.data.date].filter(Boolean).join(' - ');
    return {
      ...exp.data,
      ...amounts(exp.data),
      exp,
      index: i + 1,
      fileName: safeFileName(`Justificatif ${n}${label ? ' - ' + label : ''}`) + '.' + extFor(exp.mime, exp.fileName),
    };
  });
  return {
    profile,
    client,
    currency: items.find(i => i.currency)?.currency || 'EUR',
    today: todayISO(),
    ...summarize(items),
  };
}

export function lineValue(key, item) {
  const v = item[key];
  return v == null ? '' : v;
}

export function singleValue(key, report, fixedValue = '') {
  if (key === 'fixed') return fixedValue;
  const [scope, prop] = key.split('.');
  if (scope === 'profile') return report.profile?.[prop] ?? '';
  if (scope === 'client') return report.client?.[prop] ?? '';
  if (scope === 'report') return report[prop] ?? '';
  return '';
}

export function fillPlaceholders(str, report) {
  const fmt = n => new Intl.NumberFormat('fr-BE', { style: 'currency', currency: report.currency || 'EUR' }).format(n || 0);
  return String(str || '')
    .replaceAll('{nom}', report.profile?.name || '')
    .replaceAll('{periode}', report.period || '')
    .replaceAll('{total}', fmt(report.total))
    .replaceAll('{nb}', String(report.count))
    .replaceAll('{client}', report.client?.name || '')
    .replaceAll('{mois}', report.month || '')
    .replaceAll('{date}', parseISO(report.today).toLocaleDateString('fr-BE'));
}

// Données fictives pour tester le format d'un client.
export function sampleExpenses() {
  const d = new Date();
  const iso = day => todayISO(new Date(d.getFullYear(), d.getMonth(), day));
  const mk = (i, data) => ({ id: 'ex' + i, createdAt: i, mime: 'image/jpeg', fileName: 'exemple.jpg', data: { ...emptyData(), ...data } });
  return [
    mk(1, { date: iso(3), merchant: 'Brasserie de la Gare', description: 'Déjeuner client – 2 couverts', category: 'Restaurant & repas', docType: 'Note de restaurant', total: 68.5, vat: 7.95, vatRate: '12 %', payment: 'Carte bancaire' }),
    mk(2, { date: iso(5), merchant: 'SNCB', description: 'Train Bruxelles-Central – Liège-Guillemins A/R', category: 'Train & transports en commun', docType: 'Titre de transport', total: 31.6, vat: 1.79, vatRate: '6 %', payment: 'Carte bancaire' }),
    mk(3, { date: iso(8), merchant: 'Interparking', description: 'Parking Grand-Place', category: 'Parking & péage', docType: 'Ticket de caisse', total: 18, vat: 3.12, vatRate: '21 %', payment: 'Carte bancaire' }),
  ];
}
