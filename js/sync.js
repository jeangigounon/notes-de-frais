// Synchronisation PC ↔ téléphone via GitHub :
// - un Gist secret « notes-de-frais.json » : profil, clients (modèles Excel compris) et liste des justificatifs ;
// - un Gist secret par photo / PDF de justificatif, lu directement par son lien brut (léger, jamais tout re-téléchargé).

import {
  state, saveSettings, replaceSynced, emit, setSyncHook,
  getExpense, saveExpense, deleteExpense, expenseTombstones, setExpenseTombstones,
} from './store.js';
import { debounce, blobToBase64, base64ToBlob } from './ui.js';

const FILE = 'notes-de-frais.json';
const API = 'https://api.github.com';
const DAY = 86400000;
const KEEP_SENT_DAYS = 60;     // un justificatif envoyé reste synchronisé 60 jours
const KEEP_PHOTO_DAYS = 30;    // sa photo reste en ligne 30 jours après l'envoi
const MAX_PHOTO = 7_000_000;   // au-delà, la photo reste sur l'appareil

function setStatus(status, message = '') {
  state.sync = { status, message };
  emit('sync');
}

async function gh(path, { method = 'GET', body } = {}) {
  const res = await fetch(API + path, {
    method,
    headers: {
      Accept: 'application/vnd.github+json',
      Authorization: `Bearer ${state.settings.ghToken}`,
      'X-GitHub-Api-Version': '2022-11-28',
      ...(body ? { 'Content-Type': 'application/json' } : {}),
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  if (res.status === 401) throw new Error('Jeton GitHub refusé : vérifiez-le dans Réglages.');
  if (res.status === 404) throw Object.assign(new Error('Gist introuvable.'), { notFound: true });
  if (res.status === 403 || res.status === 429) throw new Error('GitHub limite temporairement les envois, nouvel essai dans quelques minutes.');
  if (!res.ok) throw new Error(`GitHub : erreur ${res.status}`);
  return res.status === 204 ? null : res.json();
}

async function readGist(gist) {
  const file = gist.files?.[FILE];
  if (!file) return null;
  const content = file.truncated ? await (await fetch(file.raw_url)).text() : file.content;
  return JSON.parse(content || '{}');
}

async function findGist() {
  for (let page = 1; page <= 10; page++) {
    const list = await gh(`/gists?per_page=100&page=${page}`);
    const hit = list.find(g => g.files?.[FILE]);
    if (hit) return hit.id;
    if (list.length < 100) break;
  }
  return null;
}

/* ---------- Justificatifs ---------- */

const age = ts => Date.now() - (ts || 0);
const isSyncable = e => e.status !== 'sent' || age(e.sentAt || e.updatedAt) < KEEP_SENT_DAYS * DAY;

// Ce qui part dans le Gist : tout sauf le fichier lui-même (et une vignette réduite).
function expenseMeta(e) {
  const { file, thumb, syncThumb, ...rest } = e;
  return { ...rest, thumb: syncThumb || null };
}

async function shrinkThumb(dataUrl, side = 160) {
  const img = new Image();
  img.src = dataUrl;
  await img.decode();
  const scale = Math.min(1, side / Math.max(img.naturalWidth, img.naturalHeight));
  const canvas = document.createElement('canvas');
  canvas.width = Math.round(img.naturalWidth * scale);
  canvas.height = Math.round(img.naturalHeight * scale);
  canvas.getContext('2d').drawImage(img, 0, 0, canvas.width, canvas.height);
  return canvas.toDataURL('image/jpeg', 0.6);
}

function mergeExpenses(localList, remoteList, localTomb, remoteTomb) {
  const tombstones = { ...(remoteTomb || {}) };
  for (const [id, t] of Object.entries(localTomb)) if (!tombstones[id] || t.at > tombstones[id].at) tombstones[id] = t;
  for (const [id, t] of Object.entries(tombstones)) if (age(t.at) > KEEP_SENT_DAYS * DAY) delete tombstones[id];
  const byId = new Map();
  for (const e of remoteList || []) byId.set(e.id, e);
  for (const e of localList) {
    const r = byId.get(e.id);
    if (!r || (e.updatedAt || 0) > (r.updatedAt || 0)) byId.set(e.id, { ...e, receipt: e.receipt || r?.receipt || null });
    else if (!r.receipt && e.receipt) byId.set(e.id, { ...r, receipt: e.receipt });
  }
  for (const [id, t] of Object.entries(tombstones)) {
    const e = byId.get(id);
    if (e && (e.updatedAt || 0) <= t.at) byId.delete(id);
  }
  return { expenses: [...byId.values()].filter(isSyncable), tombstones };
}

// Applique la version distante d'un justificatif en gardant le fichier et la vignette locaux.
async function applyRemote(meta) {
  const local = getExpense(meta.id);
  if (local && (local.updatedAt || 0) >= (meta.updatedAt || 0) && (local.receipt || !meta.receipt)) return;
  const { thumb: remoteThumb, ...rest } = meta;
  const base = local && (local.updatedAt || 0) >= (meta.updatedAt || 0) ? { ...local, receipt: meta.receipt } : { ...(local || {}), ...rest };
  await saveExpense({
    ...base,
    file: local?.file || null,
    thumb: local?.thumb || remoteThumb || null,
    syncThumb: remoteThumb || local?.syncThumb || null,
  }, { touch: false, sync: false });
}

async function uploadReceipt(exp) {
  const name = `justificatif-${exp.id}.b64`;
  const gist = await gh('/gists', {
    method: 'POST',
    body: { description: 'Notes de frais – justificatif synchronisé', public: false, files: { [name]: { content: await blobToBase64(exp.file) } } },
  });
  return { gist: gist.id, url: gist.files[name].raw_url, size: exp.file.size };
}

async function removeReceipt(receipt) {
  try { await gh(`/gists/${receipt.gist}`, { method: 'DELETE' }); } catch (err) { if (!err.notFound) throw err; }
}

const downloading = new Map();

// Récupère la photo d'un justificatif ajouté sur un autre appareil.
export function ensureReceiptFile(exp) {
  if (exp.file || !exp.receipt?.url) return Promise.resolve(exp);
  if (!downloading.has(exp.id)) {
    downloading.set(exp.id, (async () => {
      try {
        const res = await fetch(exp.receipt.url);
        if (!res.ok) throw new Error('Photo introuvable en ligne');
        const blob = base64ToBlob((await res.text()).trim(), exp.mime || 'image/jpeg');
        const fresh = getExpense(exp.id) || exp;
        fresh.file = blob;
        await saveExpense(fresh, { touch: false, sync: false });
        return fresh;
      } finally {
        downloading.delete(exp.id);
      }
    })());
  }
  return downloading.get(exp.id);
}

async function syncExpenses(remote) {
  // Vignettes réduites pour la liste partagée.
  for (const e of state.expenses) {
    if (e.thumb && !e.syncThumb && isSyncable(e)) {
      try { e.syncThumb = await shrinkThumb(e.thumb); await saveExpense(e, { touch: false, sync: false }); } catch { /* sans vignette */ }
    }
  }
  const localTomb = expenseTombstones();
  let merged = mergeExpenses(state.expenses.filter(isSyncable).map(expenseMeta), remote?.expenses, localTomb, remote?.deletedExpenses);

  for (const meta of merged.expenses) await applyRemote(meta);
  for (const [id, t] of Object.entries(merged.tombstones)) {
    const e = getExpense(id);
    if (e && (e.updatedAt || 0) <= t.at) await deleteExpense(id, { sync: false });
  }

  // Photos à mettre en ligne (justificatifs pris sur cet appareil).
  let uploaded = false;
  for (const e of state.expenses) {
    if (!isSyncable(e) || e.receipt || !e.file || e.file.size > MAX_PHOTO) continue;
    e.receipt = await uploadReceipt(e);
    await saveExpense(e, { sync: false });
    uploaded = true;
  }

  // Photos qui n'ont plus à rester en ligne : justificatifs supprimés, ou envoyés depuis longtemps.
  const tomb = { ...merged.tombstones };
  for (const [id, t] of Object.entries(tomb)) {
    if (t.receipt && !t.receiptRemoved) { await removeReceipt(t.receipt); tomb[id] = { ...t, receiptRemoved: true }; }
  }
  for (const e of state.expenses) {
    if (e.receipt && e.status === 'sent' && age(e.sentAt) > KEEP_PHOTO_DAYS * DAY && e.file) {
      await removeReceipt(e.receipt);
      e.receipt = null;
      await saveExpense(e, { sync: false });
      uploaded = true;
    }
  }
  setExpenseTombstones(Object.fromEntries(Object.entries(tomb).filter(([id]) => localTomb[id] || !getExpense(id))));
  if (uploaded) merged = mergeExpenses(state.expenses.filter(isSyncable).map(expenseMeta), merged.expenses, tomb, {});
  return { expenses: merged.expenses, deletedExpenses: tomb };
}

/* ---------- Synchronisation complète ---------- */

function mergeConfig(remote) {
  const byId = new Map();
  for (const c of remote?.clients || []) byId.set(c.id, c);
  for (const c of state.clients) {
    const r = byId.get(c.id);
    if (!r || (c.updatedAt || 0) > (r.updatedAt || 0)) byId.set(c.id, c);
  }
  const rp = remote?.profile;
  const profile = rp && (rp.updatedAt || 0) > (state.profile.updatedAt || 0) ? rp : state.profile;
  return { profile, clients: [...byId.values()] };
}

const fingerprint = d => JSON.stringify({
  p: d?.profile?.updatedAt || 0,
  c: (d?.clients || []).map(c => [c.id, c.updatedAt || 0]).sort(),
  e: (d?.expenses || []).map(e => [e.id, e.updatedAt || 0, e.receipt?.gist || '']).sort(),
  d: Object.keys(d?.deletedExpenses || {}).sort(),
});

let running = null;
let again = false;

export function syncNow() {
  if (!state.settings.ghToken) return Promise.resolve(false);
  if (running) { again = true; return running; }
  running = (async () => {
    setStatus('syncing');
    try {
      let gistId = state.settings.gistId;
      if (!gistId) gistId = await findGist();
      let remote = null;
      if (gistId) {
        try {
          remote = await readGist(await gh(`/gists/${gistId}`));
        } catch (err) {
          if (!err.notFound) throw err;
          gistId = await findGist();
          remote = gistId ? await readGist(await gh(`/gists/${gistId}`)) : null;
        }
      }
      const config = mergeConfig(remote);
      if (fingerprint(config) !== fingerprint({ profile: state.profile, clients: state.clients })) replaceSynced(config);
      const expenses = await syncExpenses(remote);
      const payload = { app: 'notes-de-frais', version: 2, ...config, ...expenses };
      if (!gistId) {
        const created = await gh('/gists', {
          method: 'POST',
          body: { description: 'Notes de frais – configuration synchronisée (ne pas supprimer)', public: false, files: { [FILE]: { content: JSON.stringify(payload) } } },
        });
        gistId = created.id;
      } else if (fingerprint(payload) !== fingerprint(remote)) {
        await gh(`/gists/${gistId}`, { method: 'PATCH', body: { files: { [FILE]: { content: JSON.stringify(payload) } } } });
      }
      saveSettings({ gistId, lastSync: Date.now() });
      setStatus('ok');
      // Photos venues de l'autre appareil : téléchargées en arrière-plan.
      for (const e of state.expenses.filter(x => !x.file && x.receipt)) ensureReceiptFile(e).catch(() => {});
      return true;
    } catch (err) {
      console.error(err);
      setStatus('error', err.message);
      return false;
    } finally {
      running = null;
      if (again) { again = false; setTimeout(syncNow, 0); }
    }
  })();
  return running;
}

const scheduled = debounce(() => syncNow(), 2000);

export function initSync() {
  setSyncHook(() => { if (state.settings.ghToken) scheduled(); });
  if (state.settings.ghToken) syncNow();
  const stale = ms => state.settings.ghToken && Date.now() - state.settings.lastSync > ms;
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'visible' && stale(15000)) syncNow();
  });
  // Application ouverte : on récupère régulièrement ce qui a été ajouté ailleurs.
  setInterval(() => { if (document.visibilityState === 'visible' && stale(120000)) syncNow(); }, 30000);
}
