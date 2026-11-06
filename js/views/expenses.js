// Liste des dépenses (à envoyer / envoyées / toutes).

import { state, getClient } from '../store.js';
import { esc, icon, money, fmtDate, todayISO } from '../ui.js';
import { amounts, CLIENT_COLORS } from '../fields.js';
import { openExpenseSheet } from './expenseSheet.js';
import { openFilePicker } from '../capture.js';

export const deps = ['expenses', 'clients'];

let filter = 'pending';
let query = '';

export function thumb(e) {
  const busy = e.aiStatus === 'pending' ? ' busy' : '';
  if (e.thumb) return `<span class="thumb${busy}"><img src="${e.thumb}" alt=""></span>`;
  const pdf = e.mime === 'application/pdf';
  return `<span class="thumb thumb-doc${busy}">${icon(pdf ? 'doc' : 'image')}${pdf ? '<small>PDF</small>' : ''}</span>`;
}

export function expenseRow(e, { showClient = true } = {}) {
  const c = getClient(e.clientId);
  const { total } = amounts(e.data);
  const title = e.data.merchant || (e.aiStatus === 'pending' ? 'Analyse en cours…' : 'Justificatif');
  const sub = [fmtDate(e.data.date || todayISO(new Date(e.createdAt))), e.data.category].filter(Boolean).join(' · ');
  let tag = '';
  if (e.status === 'sent') tag = `<span class="sent-tag">${icon('check')}Envoyé</span>`;
  else if (showClient) tag = c
    ? `<span class="client-tag"><i style="background:${CLIENT_COLORS[c.color] || CLIENT_COLORS.blue}"></i>${esc(c.name || 'Sans nom')}</span>`
    : '<span class="client-tag muted">Sans destinataire</span>';
  return `
    <button type="button" class="row row-btn" data-open="${e.id}">
      ${thumb(e)}
      <span class="row-main"><span class="row-title">${esc(title)}</span><span class="row-sub">${esc(sub)}</span></span>
      <span class="row-end"><span class="amount">${total != null ? money(total, e.data.currency) : '—'}</span>${tag}</span>
    </button>`;
}

function monthKey(e) {
  const iso = e.data.date || todayISO(new Date(e.createdAt));
  return iso.slice(0, 7);
}

export function render(page) {
  const counts = {
    pending: state.expenses.filter(e => e.status !== 'sent').length,
    sent: state.expenses.filter(e => e.status === 'sent').length,
  };
  const q = query.trim().toLowerCase();
  const list = state.expenses
    .filter(e => filter === 'all' || (filter === 'sent' ? e.status === 'sent' : e.status !== 'sent'))
    .filter(e => !q || [e.data.merchant, e.data.description, e.data.category, getClient(e.clientId)?.name].join(' ').toLowerCase().includes(q))
    .sort((a, b) => (b.data.date || todayISO(new Date(b.createdAt))).localeCompare(a.data.date || todayISO(new Date(a.createdAt))) || b.createdAt - a.createdAt);

  const months = new Map();
  for (const e of list) {
    const k = monthKey(e);
    if (!months.has(k)) months.set(k, []);
    months.get(k).push(e);
  }

  page.innerHTML = `
    <header class="page-head">
      <div><h1 class="large-title">Dépenses</h1></div>
      <button type="button" class="btn btn-secondary hide-compact" data-act="add">${icon('upload')}Importer</button>
    </header>
    <div class="toolbar">
      <div class="seg" role="tablist">
        <button type="button" data-filter="pending" class="${filter === 'pending' ? 'on' : ''}">À envoyer${counts.pending ? ` · ${counts.pending}` : ''}</button>
        <button type="button" data-filter="sent" class="${filter === 'sent' ? 'on' : ''}">Envoyées</button>
        <button type="button" data-filter="all" class="${filter === 'all' ? 'on' : ''}">Toutes</button>
      </div>
      <label class="search">${icon('search')}<input type="search" placeholder="Rechercher" value="${esc(query)}" data-search></label>
    </div>
    ${list.length ? [...months.entries()].map(([k, items]) => {
      const sum = items.reduce((s, e) => s + (amounts(e.data).total || 0), 0);
      const label = new Date(k + '-15').toLocaleDateString('fr-BE', { month: 'long', year: 'numeric' });
      return `
        <div class="group-label"><span>${esc(label)}</span><span>${money(sum)}</span></div>
        <div class="list inset-thumb">${items.map(e => expenseRow(e)).join('')}</div>`;
    }).join('') : `
      <div class="empty">
        ${icon(filter === 'sent' ? 'send' : 'tray')}
        <h3>${q ? 'Aucun résultat' : filter === 'sent' ? 'Rien n’a encore été envoyé' : 'Tout est envoyé'}</h3>
        <p>${q ? 'Essayez un autre mot.' : 'Les justificatifs que vous ajoutez apparaissent ici.'}</p>
      </div>`}
  `;

  page.onclick = e => {
    const open = e.target.closest('[data-open]');
    if (open) return openExpenseSheet(open.dataset.open);
    const f = e.target.closest('[data-filter]');
    if (f) { filter = f.dataset.filter; render(page); return; }
    if (e.target.closest('[data-act="add"]')) openFilePicker();
  };
  const search = page.querySelector('[data-search]');
  search.oninput = () => {
    query = search.value;
    const pos = search.selectionStart;
    render(page);
    const again = page.querySelector('[data-search]');
    again.focus();
    again.setSelectionRange(pos, pos);
  };
}
