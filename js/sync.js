// Synchronisation PC ↔ téléphone du profil et des clients (modèles Excel compris)
// via un Gist GitHub secret. Les justificatifs restent sur l'appareil qui les a photographiés.

import { state, saveSettings, replaceSynced, emit, setSyncHook } from './store.js';
import { debounce } from './ui.js';

const FILE = 'notes-de-frais.json';
const API = 'https://api.github.com';

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
  if (!res.ok) throw new Error(`GitHub : erreur ${res.status}`);
  return res.json();
}

const localPayload = () => ({ app: 'notes-de-frais', version: 1, profile: state.profile, clients: state.clients });

async function readGist(gist) {
  const file = gist.files?.[FILE];
  if (!file) return null;
  const content = file.truncated ? await (await fetch(file.raw_url)).text() : file.content;
  return JSON.parse(content || '{}');
}

async function findGist() {
  for (let page = 1; page <= 5; page++) {
    const list = await gh(`/gists?per_page=100&page=${page}`);
    const hit = list.find(g => g.files?.[FILE]);
    if (hit) return hit.id;
    if (list.length < 100) break;
  }
  return null;
}

function merge(local, remote) {
  const byId = new Map();
  for (const c of remote?.clients || []) byId.set(c.id, c);
  for (const c of local.clients) {
    const r = byId.get(c.id);
    if (!r || (c.updatedAt || 0) > (r.updatedAt || 0)) byId.set(c.id, c);
  }
  const rp = remote?.profile;
  const profile = rp && (rp.updatedAt || 0) > (local.profile.updatedAt || 0) ? rp : local.profile;
  return { ...local, profile, clients: [...byId.values()] };
}

const fingerprint = d => JSON.stringify({ p: d?.profile?.updatedAt || 0, c: (d?.clients || []).map(c => [c.id, c.updatedAt || 0]).sort() });

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
      const local = localPayload();
      const merged = merge(local, remote);
      if (fingerprint(merged) !== fingerprint(local)) replaceSynced(merged);
      if (!gistId) {
        const created = await gh('/gists', {
          method: 'POST',
          body: { description: 'Notes de frais – configuration synchronisée (ne pas supprimer)', public: false, files: { [FILE]: { content: JSON.stringify(merged) } } },
        });
        gistId = created.id;
      } else if (fingerprint(merged) !== fingerprint(remote)) {
        await gh(`/gists/${gistId}`, { method: 'PATCH', body: { files: { [FILE]: { content: JSON.stringify(merged) } } } });
      }
      saveSettings({ gistId, lastSync: Date.now() });
      setStatus('ok');
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
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'visible' && state.settings.ghToken && Date.now() - state.settings.lastSync > 60000) syncNow();
  });
}
