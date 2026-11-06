// Démarrage, navigation et coquille de l'application (barre latérale sur PC, barre d'onglets sur téléphone).

import { state, onChange, loadExpenses, pendingExpenses } from './store.js';
import { $, icon, esc, toast, timeAgo, spinner } from './ui.js';
import { initCapture, openCamera } from './capture.js';
import { resumePending } from './ai.js';
import { initSync } from './sync.js';
import { msHandleRedirect } from './mail.js';
import { handleImportLink } from './transfer.js';
import { openSendSheet } from './views/sendSheet.js';
import * as home from './views/home.js';
import * as expenses from './views/expenses.js';
import * as clients from './views/clients.js';
import * as settings from './views/settings.js';

const ROUTES = {
  '': { view: home, nav: 'home', title: 'Accueil' },
  depenses: { view: expenses, nav: 'expenses', title: 'Dépenses' },
  clients: { view: clients, nav: 'clients', title: 'Clients' },
  reglages: { view: settings, nav: 'settings', title: 'Réglages' },
};

const NAV = [
  { id: 'home', label: 'Accueil', icon: 'home', href: '#/' },
  { id: 'expenses', label: 'Dépenses', icon: 'receipt', href: '#/depenses' },
  { id: 'clients', label: 'Clients', icon: 'users', href: '#/clients' },
  { id: 'settings', label: 'Réglages', icon: 'sliders', href: '#/reglages' },
];

let current = null;

function parseRoute() {
  const [path = '', ...params] = location.hash.replace(/^#\/?/, '').split('/').map(decodeURIComponent);
  return { path, params };
}

function syncLine() {
  if (!state.settings.ghToken) return '';
  if (state.sync.status === 'syncing') return `${spinner()}Synchronisation…`;
  if (state.sync.status === 'error') return `<span class="status-dot err"></span>Synchronisation en échec`;
  return `${icon('cloud')}Synchronisé ${timeAgo(state.settings.lastSync)}`;
}

function renderNav(active) {
  const count = pendingExpenses().length;
  $('#sidebar').innerHTML = `
    <a class="brand" href="#/"><img src="icons/icon.svg" alt="" width="30" height="30"><span>Notes de frais</span></a>
    <button type="button" class="btn btn-primary btn-block sidebar-cta" data-act="add">${icon('plus')}Nouveau justificatif</button>
    <nav class="nav">
      ${NAV.map(n => `<a class="nav-item ${n.id === active ? 'on' : ''}" href="${n.href}">${icon(n.icon)}<span>${n.label}</span>${n.id === 'expenses' && count ? `<span class="count">${count}</span>` : ''}</a>`).join('')}
    </nav>
    <div class="sidebar-foot" data-sync-line>${syncLine()}</div>`;

  const tab = n => `<a class="tab ${n.id === active ? 'on' : ''}" href="${n.href}">${icon(n.icon)}<span>${n.label}</span>${n.id === 'expenses' && count ? `<b class="badge">${count}</b>` : ''}</a>`;
  $('#tabbar').innerHTML = `
    ${tab(NAV[0])}${tab(NAV[1])}
    <button type="button" class="tab tab-cam" data-act="camera" aria-label="Photographier un ticket"><span class="cam">${icon('camera')}</span></button>
    ${tab(NAV[2])}${tab(NAV[3])}`;
}

function render({ keepScroll = false } = {}) {
  const { path, params } = parseRoute();
  const route = ROUTES[path] || ROUTES[''];
  const changed = current?.route !== route || current?.params.join('/') !== params.join('/');
  const y = window.scrollY;
  renderNav(route.nav);
  const main = $('#main');
  const page = document.createElement('div');
  page.className = `page page-${route.nav}`;
  main.replaceChildren(page);
  route.view.render(page, params);
  current = { route, params, page };
  document.title = route === ROUTES[''] ? 'Notes de frais' : `${route.title} · Notes de frais`;
  if (keepScroll && !changed) window.scrollTo(0, y); else window.scrollTo(0, 0);
}

function onStore(what, detail) {
  if (what === 'sync' || what === 'settings') {
    const line = document.querySelector('[data-sync-line]');
    if (line) line.innerHTML = syncLine();
  }
  if (!current) return;
  const view = current.route.view;
  if (!view.deps?.includes(what)) {
    if (what === 'expenses') renderNav(current.route.nav);
    return;
  }
  if (view.onStoreChange?.(what, detail, current.page, current.params)) {
    if (what === 'expenses') renderNav(current.route.nav);
    return;
  }
  render({ keepScroll: true });
}

async function boot() {
  if ('serviceWorker' in navigator && location.protocol === 'https:') {
    navigator.serviceWorker.register('sw.js').catch(err => console.warn('Service worker', err));
  }

  // Retour de la connexion Microsoft : reprendre l'envoi en cours.
  let resumeSend = null;
  try {
    const res = await msHandleRedirect();
    if (res.handled) {
      toast('Connecté à Microsoft 365', { type: 'success' });
      if (res.returnState?.kind === 'send') resumeSend = res.returnState;
    }
  } catch (err) {
    toast(err.message, { type: 'error', duration: 7000 });
  }

  try {
    await loadExpenses();
  } catch (err) {
    console.error(err);
    toast('Stockage local indisponible (navigation privée ?) : les justificatifs ne seront pas conservés.', { type: 'error', duration: 7000 });
  }

  initCapture();
  document.addEventListener('click', e => {
    const act = e.target.closest('[data-act]')?.dataset.act;
    if (!e.target.closest('#sidebar, #tabbar')) return;
    if (act === 'camera') openCamera();
    if (act === 'add') $('#file-input').click();
  });

  window.addEventListener('hashchange', () => {
    if (location.hash.startsWith('#import=')) { handleImportLink().then(() => render()); return; }
    render();
  });
  onChange(onStore);

  const importing = location.hash.startsWith('#import=');
  render();
  if (importing) await handleImportLink();

  initSync();
  resumePending();
  setInterval(() => { const l = document.querySelector('[data-sync-line]'); if (l) l.innerHTML = syncLine(); }, 60000);

  if (resumeSend) openSendSheet({ clientId: resumeSend.clientId, preset: resumeSend });

  // Outils de diagnostic dans la console.
  window.ndf = { state, render, esc };
}

boot();
