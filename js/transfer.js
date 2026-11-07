// Configurer le téléphone en scannant un QR code affiché sur le PC, et sauvegarde complète.

import { state, saveSettings, saveProfile, saveExpense, replaceSynced } from './store.js';
import { CDN, loadScript } from './lib.js';
import { blobToBase64, base64ToBlob, confirmDialog, download, todayISO, toast } from './ui.js';

const SETTING_KEYS = ['aiMode', 'anthropicKey', 'model', 'sendMethod', 'msClientId', 'msTenant', 'msAccountType', 'msLoginHint', 'ghToken', 'gistId'];

const b64urlEncode = s => btoa(unescape(encodeURIComponent(s))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
const b64urlDecode = s => decodeURIComponent(escape(atob(s.replace(/-/g, '+').replace(/_/g, '/'))));

export function setupLink() {
  const s = {};
  for (const k of SETTING_KEYS) if (state.settings[k]) s[k] = state.settings[k];
  const { name, email, company, iban } = state.profile;
  const payload = { v: 1, s, p: { name, email, company, iban } };
  return `${location.origin}${location.pathname}#import=${b64urlEncode(JSON.stringify(payload))}`;
}

export async function qrSvg(text) {
  await loadScript(CDN.qrcode);
  const qr = window.qrcode(0, 'M');
  qr.addData(text);
  qr.make();
  return qr.createSvgTag({ cellSize: 4, margin: 0, scalable: true });
}

// Appelé au démarrage : #import=… dans l'adresse.
export async function handleImportLink() {
  if (!location.hash.startsWith('#import=')) return false;
  const raw = location.hash.slice('#import='.length);
  history.replaceState(null, '', location.pathname + location.search + '#/');
  let payload;
  try { payload = JSON.parse(b64urlDecode(raw)); } catch { toast('Lien de configuration illisible', { type: 'error' }); return true; }
  const ok = await confirmDialog({
    title: 'Configurer cet appareil ?',
    message: 'Les réglages de votre PC (clé Claude, synchronisation, Microsoft 365 et profil) vont être copiés ici.',
    confirm: 'Configurer',
  });
  if (!ok) return true;
  const s = {};
  for (const k of SETTING_KEYS) if (payload.s?.[k] != null) s[k] = payload.s[k];
  saveSettings(s);
  if (payload.p) saveProfile({ ...payload.p, updatedAt: 0 });
  toast('Appareil configuré', { type: 'success' });
  const { syncNow } = await import('./sync.js');
  syncNow();
  return true;
}

/* ---------- Sauvegarde ---------- */

export async function exportBackup() {
  const expenses = await Promise.all(state.expenses.map(async e => ({ ...e, file: e.file ? await blobToBase64(e.file) : null })));
  const data = { app: 'notes-de-frais', version: 1, exportedAt: Date.now(), profile: state.profile, clients: state.clients, expenses };
  download(new Blob([JSON.stringify(data)], { type: 'application/json' }), `notes-de-frais-sauvegarde-${todayISO()}.json`);
}

export async function importBackup(file) {
  const data = JSON.parse(await file.text());
  if (data.app !== 'notes-de-frais') throw new Error("Ce fichier n'est pas une sauvegarde de Notes de frais.");
  const known = new Map(state.clients.map(c => [c.id, c]));
  for (const c of data.clients || []) {
    const cur = known.get(c.id);
    if (!cur || (c.updatedAt || 0) > (cur.updatedAt || 0)) known.set(c.id, c);
  }
  replaceSynced({ clients: [...known.values()], profile: (data.profile?.updatedAt || 0) > (state.profile.updatedAt || 0) ? data.profile : null });
  let added = 0;
  for (const e of data.expenses || []) {
    if (state.expenses.some(x => x.id === e.id)) continue;
    await saveExpense({ ...e, file: e.file ? base64ToBlob(e.file, e.mime) : null });
    added++;
  }
  return { clients: (data.clients || []).length, expenses: added };
}
