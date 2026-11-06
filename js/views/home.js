// Accueil : capturer un justificatif en un geste, voir ce qui reste à envoyer.

import { state, activeClients, getClient, pendingExpenses } from '../store.js';
import { esc, icon, money, fmtDateShort, todayISO, isTouch } from '../ui.js';
import { amounts, CLIENT_COLORS } from '../fields.js';
import { openCamera, openFilePicker } from '../capture.js';
import { openSendSheet } from './sendSheet.js';
import { openExpenseSheet } from './expenseSheet.js';
import { expenseRow } from './expenses.js';

export const deps = ['expenses', 'clients', 'profile'];

export const initials = name => (name || '?').split(/\s+/).filter(Boolean).slice(0, 2).map(w => w[0].toUpperCase()).join('') || '?';

export function render(page) {
  const pending = pendingExpenses();
  const clients = activeClients();
  const groups = new Map();
  for (const e of pending) {
    const key = getClient(e.clientId) ? e.clientId : '';
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(e);
  }
  const oldest = pending.reduce((m, e) => Math.min(m, e.createdAt), Infinity);
  const oldDays = isFinite(oldest) ? Math.floor((Date.now() - oldest) / 86400000) : 0;
  const totalPending = pending.reduce((s, e) => s + (amounts(e.data).total || 0), 0);
  const today = new Date().toLocaleDateString('fr-BE', { weekday: 'long', day: 'numeric', month: 'long' });
  const recent = state.expenses.slice(0, 5);
  const firstName = (state.profile.name || '').split(' ')[0];

  const clientCards = [...groups.entries()].filter(([k]) => k).map(([id, items]) => {
    const c = getClient(id);
    const total = items.reduce((s, e) => s + (amounts(e.data).total || 0), 0);
    const since = Math.min(...items.map(e => e.createdAt));
    return `
      <article class="send-card">
        <div class="send-card-head">
          <span class="avatar" style="background:${CLIENT_COLORS[c.color] || CLIENT_COLORS.blue}">${esc(initials(c.name))}</span>
          <div class="row-main"><span class="row-title">${esc(c.name || 'Sans nom')}</span><span class="row-sub">${items.length} justificatif${items.length > 1 ? 's' : ''} · depuis le ${fmtDateShort(todayISO(new Date(since)))}</span></div>
        </div>
        <div class="send-card-foot">
          <span class="big-amount">${money(total)}</span>
          <button type="button" class="btn btn-primary" data-send="${id}">${icon('send')}Envoyer</button>
        </div>
      </article>`;
  }).join('');

  const orphans = groups.get('') || [];

  page.innerHTML = `
    <header class="page-head">
      <div><p class="eyebrow">${esc(today)}</p><h1 class="large-title">${firstName ? `Bonjour ${esc(firstName)}` : 'Notes de frais'}</h1></div>
    </header>

    <section class="card capture" data-capture>
      <div class="capture-icon">${icon('camera')}</div>
      <div class="capture-text">
        <h2>Nouveau justificatif</h2>
        <p>Photographiez un ticket ou importez une facture : l'IA en lit le contenu pour vous.</p>
      </div>
      <div class="btn-row">
        ${isTouch() ? `
          <button type="button" class="btn btn-primary btn-lg" data-act="camera">${icon('camera')}Prendre une photo</button>
          <button type="button" class="btn btn-secondary btn-lg" data-act="file">${icon('upload')}Choisir un fichier</button>` : `
          <button type="button" class="btn btn-primary btn-lg" data-act="file">${icon('upload')}Importer un ticket ou une facture</button>`}
      </div>
      <p class="drop-hint hide-compact">ou glissez-déposez vos tickets et factures n'importe où sur cette page, ou collez-les avec Ctrl+V</p>
    </section>

    ${!state.profile.name ? `
      <a class="banner info" href="#/reglages">${icon('user')}<span><strong>Indiquez votre nom</strong> pour qu'il apparaisse sur vos notes de frais.</span>${icon('chevronRight', 'chev')}</a>` : ''}
    ${!clients.length ? `
      <a class="banner info" href="#/clients">${icon('users')}<span><strong>Ajoutez vos clients</strong> : adresse e-mail et format Excel attendu. Plus confortable depuis votre PC.</span>${icon('chevronRight', 'chev')}</a>` : ''}

    ${pending.length ? `
      <section class="section">
        <h2 class="section-title">À envoyer <span class="pill">${pending.length}</span><span class="section-aside">${money(totalPending)}</span></h2>
        ${oldDays >= 7 ? `<div class="banner">${icon('clock')}<span>Le plus ancien justificatif attend depuis <strong>${oldDays} jours</strong>. Pensez à l'envoyer !</span></div>` : ''}
        ${clientCards ? `<div class="send-grid">${clientCards}</div>` : ''}
        ${orphans.length ? `
          <h3 class="group-label">Sans destinataire</h3>
          <div class="list inset-thumb">${orphans.map(e => expenseRow(e, { showClient: false })).join('')}</div>
          <p class="group-foot">Ouvrez un justificatif pour choisir à qui l'envoyer.</p>` : ''}
      </section>` : `
      <section class="section">
        <div class="all-done">${icon('check')}<div><strong>Tout est envoyé.</strong><span>Aucun justificatif en attente.</span></div></div>
      </section>`}

    ${recent.length ? `
      <section class="section">
        <h2 class="section-title">Récents <a class="section-link" href="#/depenses">Tout voir</a></h2>
        <div class="list inset-thumb">${recent.map(e => expenseRow(e)).join('')}</div>
      </section>` : ''}
  `;

  page.onclick = e => {
    const act = e.target.closest('[data-act]')?.dataset.act;
    if (act === 'camera') return openCamera();
    if (act === 'file') return openFilePicker();
    const send = e.target.closest('[data-send]');
    if (send) return openSendSheet({ clientId: send.dataset.send });
    const open = e.target.closest('[data-open]');
    if (open) openExpenseSheet(open.dataset.open);
  };
}
