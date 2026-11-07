// État de l'application.
// - profil, clients et réglages : localStorage (profil + clients synchronisables via GitHub Gist)
// - justificatifs (avec leurs photos) : IndexedDB, uniquement sur l'appareil

import { uid } from './ui.js';
import { DEFAULT_SIMPLE_COLUMNS } from './fields.js';

const KEYS = { profile: 'ndf.profile', clients: 'ndf.clients', settings: 'ndf.settings' };

export const DEFAULT_SETTINGS = {
  aiMode: 'claude',
  anthropicKey: '',
  model: 'claude-opus-5-5',
  sendMethod: 'auto',
  msClientId: '',
  msTenant: '',
  msAccountType: 'work',
  msLoginHint: '',
  ghToken: '',
  gistId: '',
  lastSync: 0,
  lastClientId: '',
};

function read(key, fallback) {
  try {
    const raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : fallback;
  } catch {
    return fallback;
  }
}

function write(key, value) {
  try {
    localStorage.setItem(key, JSON.stringify(value));
  } catch (err) {
    console.error(err);
    throw new Error("Stockage plein : impossible d'enregistrer. Retirez un modèle Excel volumineux.");
  }
}

export const state = {
  profile: { name: '', email: '', company: '', iban: '', updatedAt: 0, ...read(KEYS.profile, {}) },
  clients: read(KEYS.clients, []),
  settings: { ...DEFAULT_SETTINGS, ...read(KEYS.settings, {}) },
  expenses: [],
  sync: { status: 'idle', message: '' },
};

/* ---------- Évènements ---------- */

const listeners = new Set();
export function onChange(fn) {
  listeners.add(fn);
  return () => listeners.delete(fn);
}
export function emit(what, detail = {}) {
  for (const fn of [...listeners]) {
    try { fn(what, detail); } catch (err) { console.error(err); }
  }
}

let syncHook = null;
export function setSyncHook(fn) { syncHook = fn; }

// Identifiant de cet appareil (pour savoir où un justificatif a été pris).
export function deviceId() {
  let id = localStorage.getItem('ndf.device');
  if (!id) { id = uid(); localStorage.setItem('ndf.device', id); }
  return id;
}

/* ---------- Profil & réglages ---------- */

export function saveProfile(patch) {
  state.profile = { ...state.profile, ...patch, updatedAt: Date.now() };
  write(KEYS.profile, state.profile);
  emit('profile');
  syncHook?.();
}

export function saveSettings(patch) {
  state.settings = { ...state.settings, ...patch };
  write(KEYS.settings, state.settings);
  emit('settings');
}

/* ---------- Clients ---------- */

export const activeClients = () =>
  state.clients.filter(c => !c.deleted).sort((a, b) => (a.name || '').localeCompare(b.name || '', 'fr', { sensitivity: 'base' }));

export const getClient = id => (id && state.clients.find(c => c.id === id && !c.deleted)) || null;

export function newClient(patch = {}) {
  return {
    id: uid(),
    name: '',
    company: '',
    color: 'blue',
    to: '',
    cc: '',
    subject: 'Note de frais – {nom} – {periode}',
    body: 'Bonjour,\n\nVous trouverez ci-joint ma note de frais pour {periode} : {nb} justificatif(s) pour un total de {total}.\n\nMerci d’avance,\n{nom}',
    excel: {
      mode: 'simple',
      fileName: 'Note de frais {nom} {periode}',
      simple: { title: 'Note de frais', columns: DEFAULT_SIMPLE_COLUMNS.map(c => ({ ...c })) },
      template: null,
    },
    attachReceipts: true,
    notes: '',
    createdAt: Date.now(),
    updatedAt: Date.now(),
    ...patch,
  };
}

export function saveClient(client, { source = '' } = {}) {
  client.updatedAt = Date.now();
  const i = state.clients.findIndex(c => c.id === client.id);
  if (i >= 0) state.clients[i] = client; else state.clients.push(client);
  write(KEYS.clients, state.clients);
  emit('clients', { id: client.id, source });
  syncHook?.();
}

export function deleteClient(id) {
  const i = state.clients.findIndex(c => c.id === id);
  if (i < 0) return;
  // Pierre tombale : la suppression se propage aux autres appareils.
  state.clients[i] = { id, deleted: true, updatedAt: Date.now() };
  write(KEYS.clients, state.clients);
  emit('clients', { id });
  syncHook?.();
}

export function replaceSynced({ profile, clients }) {
  if (profile) { state.profile = profile; write(KEYS.profile, profile); emit('profile', { source: 'sync' }); }
  if (clients) { state.clients = clients; write(KEYS.clients, clients); emit('clients', { source: 'sync' }); }
}

/* ---------- Justificatifs (IndexedDB) ---------- */

let dbPromise;
function db() {
  dbPromise ||= new Promise((resolve, reject) => {
    const req = indexedDB.open('notes-de-frais', 1);
    req.onupgradeneeded = () => req.result.createObjectStore('expenses', { keyPath: 'id' });
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
  return dbPromise;
}

async function idb(mode, op) {
  const d = await db();
  return new Promise((resolve, reject) => {
    const tx = d.transaction('expenses', mode);
    const req = op(tx.objectStore('expenses'));
    tx.oncomplete = () => resolve(req?.result);
    tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(tx.error || new Error('Enregistrement annulé'));
  });
}

const sortExpenses = () => state.expenses.sort((a, b) => b.createdAt - a.createdAt);

export async function loadExpenses() {
  state.expenses = (await idb('readonly', s => s.getAll())) || [];
  sortExpenses();
  updateBadge();
  emit('expenses');
}

export const getExpense = id => state.expenses.find(e => e.id === id) || null;
export const pendingExpenses = () => state.expenses.filter(e => e.status !== 'sent');

// `touch: false` : enregistrement technique (synchro) qui ne compte pas comme une modification.
export async function saveExpense(exp, { touch = true, sync = true } = {}) {
  if (touch) exp.updatedAt = Date.now();
  await idb('readwrite', s => s.put(exp));
  const i = state.expenses.findIndex(e => e.id === exp.id);
  if (i >= 0) state.expenses[i] = exp; else { state.expenses.push(exp); sortExpenses(); }
  updateBadge();
  emit('expenses', { id: exp.id });
  if (sync) syncHook?.();
}

// Justificatifs supprimés : mémorisés pour que la suppression atteigne les autres appareils.
const TOMBSTONES = 'ndf.deletedExpenses';
export const expenseTombstones = () => read(TOMBSTONES, {});
export function setExpenseTombstones(map) { write(TOMBSTONES, map); }

export async function deleteExpense(id, { sync = true } = {}) {
  const exp = getExpense(id);
  await idb('readwrite', s => s.delete(id));
  state.expenses = state.expenses.filter(e => e.id !== id);
  if (sync) {
    setExpenseTombstones({ ...expenseTombstones(), [id]: { at: Date.now(), receipt: exp?.receipt || null } });
    syncHook?.();
  }
  updateBadge();
  emit('expenses', { id, deleted: true });
  return exp;
}

function updateBadge() {
  const n = pendingExpenses().length;
  try {
    if (n) navigator.setAppBadge?.(n); else navigator.clearAppBadge?.();
  } catch { /* non supporté */ }
}
