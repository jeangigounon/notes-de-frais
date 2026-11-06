// Accueil : capturer un justificatif en un geste, voir ce qui reste à envoyer.

import { state, activeClients, getClient, pendingExpenses } from '../store.js';
import { esc, icon, money, fmtDateShort, todayISO, isTouch } from '../ui.js';
import { CLIENT_COLORS } from '../fields.js';
import { openCamera, openFilePicker } from '../capture.js';
import { openSendSheet } from './sendSheet.js';
import { openExpenseSheet } from './expenseSheet.js';
import { expenseRow, thumb } from './expenses.js';
import { makeReport } from '../mail.js';

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
  const totalPending = makeReport(null, pending).displayTotal;
  const today = new Date().toLocaleDateString('fr-BE', { weekday: 'long', day: 'numeric', month: 'long' });
  const recent = state.expenses.slice(0, 5);
  const firstName = (state.profile.name || '').split(' ')[0];

  const clientCards = [...groups.entries()].filter(([k]) => k).map(([id, items]) => {
    const c = getClient(id);
    const rep = makeReport(c, items);
    const since = Math.min(...items.map(e => e.createdAt));
    return `
      <article class="send-card">
        <div class="send-card-head">
          <span class="avatar" style="background:${CLIENT_COLORS[c.color] || CLIENT_COLORS.blue}">${esc(initials(c.name))}</span>
          <div class="row-main"><span class="row-title">${esc(c.name || 'Sans nom')}</span><span class="row-sub">Note en cours · ${items.length} justificatif${items.length > 1 ? 's' : ''} depuis le ${fmtDateShort(todayISO(new Date(since)))}</span></div>
        </div>
        <div class="thumbs">${items.slice(0, 6).map(e => thumb(e)).join('')}${items.length > 6 ? `<span class="thumb more">+${items.length - 6}</span>` : ''}</div>
        <div class="send-card-foot">
          <span class="big-amount">${money(rep.displayTotal, rep.displayCurrency)}</span>
          <span class="send-card-actions">
            <button type="button" class="btn btn-secondary btn-icon-label" data-add-ticket="${id}" title="Ajouter un ticket à cette note">${icon('plus')}<span>Ticket</span></button>
            <button type="button" class="btn btn-primary" data-send="${id}">${icon('send')}Envoyer</button>
          </span>
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
        <p>Photographiez un ticket ou importez une facture PDF : l'IA en lit le contenu et l'ajoute à la note du bon client.</p>
      </div>
      <div class="btn-row">
        ${isTouch() ? `
          <button type="button" class="btn btn-primary btn-lg" data-act="camera">${icon('camera')}Prendre une photo</button>
          <button type="button" class="btn btn-secondary btn-lg" data-act="file">${icon('doc')}Facture PDF ou photo</button>` : `
          <button type="button" class="btn btn-primary btn-lg" data-act="file">${icon('upload')}Importer des tickets ou factures</button>`}
      </div>
      <p class="drop-hint hide-compact">Photos et PDF acceptés, plusieurs à la fois · glissez-déposez-les n'importe où sur cette page, ou collez-les avec Ctrl+V</p>
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
    const add = e.target.closest('[data-add-ticket]');
    if (add) return isTouch() ? openCamera({ clientId: add.dataset.addTicket }) : openFilePicker({ clientId: add.dataset.addTicket });
    const open = e.target.closest('[data-open]');
    if (open) openExpenseSheet(open.dataset.open);
  };
}
