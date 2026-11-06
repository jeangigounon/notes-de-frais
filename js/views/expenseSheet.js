// Fiche d'un justificatif : aperçu, proposition de l'IA, correction, choix du destinataire.

import { state, getExpense, saveExpense, deleteExpense, activeClients, getClient, onChange, saveSettings } from '../store.js';
import { esc, icon, openSheet, confirmDialog, toast, numStr, fmtDate, spinner } from '../ui.js';
import { CATEGORIES, DOC_TYPES, PAYMENTS, CURRENCIES, CLIENT_COLORS, num, amounts } from '../fields.js';
import { analyzeExpense } from '../ai.js';
import { openSendSheet } from './sendSheet.js';

const TEXT_FIELDS = ['merchant', 'description', 'date', 'category', 'docType', 'vatRate', 'invoiceNo', 'payment', 'currency', 'notes'];
const MONEY_FIELDS = ['total', 'net', 'vat'];

const options = (list, value, placeholder = 'Choisir…') =>
  `<option value="">${placeholder}</option>` + list.map(v => `<option ${v === value ? 'selected' : ''}>${esc(v)}</option>`).join('');

function aiStatusHtml(exp) {
  if (exp.aiStatus === 'pending') {
    const engine = state.settings.aiMode === 'claude' && state.settings.anthropicKey ? 'Claude lit le justificatif…' : 'Lecture du justificatif…';
    return `<div class="ai-status pending">${spinner()}<span>${engine}</span></div>`;
  }
  if (exp.aiStatus === 'error') {
    return `<div class="ai-status error">${icon('alert')}<span>${esc(exp.aiError || 'Analyse impossible')}</span><button type="button" class="btn btn-sm btn-secondary" data-act="reanalyze">Réessayer</button></div>`;
  }
  if (exp.aiStatus === 'done' && exp.ai) {
    const local = exp.ai.engine === 'local';
    const warn = exp.ai.isReceipt === false ? "Ce document ne ressemble pas à un justificatif. " : '';
    const remark = warn + (exp.ai.remark || '');
    const conf = { haute: 'confiance élevée', moyenne: 'confiance moyenne', faible: 'à vérifier' }[exp.ai.confidence] || '';
    return `<div class="ai-status done ${local || exp.ai.confidence === 'faible' || warn ? 'warn' : ''}">
      ${icon('sparkles')}
      <span><strong>${local ? 'Lu localement' : 'Rempli par Claude'}</strong>${conf ? ` · ${conf}` : ''}${remark ? `<br><small>${esc(remark)}</small>` : ''}
      ${local && !state.settings.anthropicKey ? '<br><small>Ajoutez une clé Claude dans Réglages pour une lecture bien plus fiable.</small>' : ''}</span>
      <button type="button" class="btn btn-sm btn-plain" data-act="reanalyze" title="Relancer l'analyse">${icon('sync')}</button>
    </div>`;
  }
  return '';
}

export function openExpenseSheet(id, { isNew = false } = {}) {
  let exp = getExpense(id);
  if (!exp) return;
  const touched = new Set();
  let removed = false;

  const previewUrl = exp.thumb && exp.file ? URL.createObjectURL(exp.file) : '';
  const sheet = openSheet({
    title: isNew ? 'Nouveau justificatif' : 'Justificatif',
    left: `<button type="button" class="btn-text" data-sheet-close>${isNew ? 'Plus tard' : 'Fermer'}</button>`,
    onClose: () => {
      unsubscribe();
      if (previewUrl) setTimeout(() => URL.revokeObjectURL(previewUrl), 1000);
      if (!removed) persist(true);
    },
  });

  const preview = exp.thumb
    ? `<img src="${previewUrl || exp.thumb}" alt="Aperçu du justificatif">`
    : `<div class="doc-tile">${icon(exp.mime === 'application/pdf' ? 'doc' : 'image')}<span>${esc(exp.fileName || 'Document')}</span></div>`;

  const clientChips = () => {
    const clients = activeClients();
    return clients.map(c => `
      <button type="button" class="chip ${exp.clientId === c.id ? 'on' : ''}" data-client="${c.id}">
        <span class="dot" style="background:${CLIENT_COLORS[c.color] || CLIENT_COLORS.blue}"></span>${esc(c.name || 'Sans nom')}
      </button>`).join('') + `<a class="chip ghost" href="#/clients" data-sheet-close>${icon('plus')}Nouveau client</a>`;
  };

  const sent = exp.status === 'sent';
  const sentTo = getClient(exp.sentTo);
  sheet.body.innerHTML = `
    <button type="button" class="exp-preview" data-act="view" aria-label="Agrandir">${preview}<span class="zoom">${icon('eye')}</span></button>
    <div data-ai>${aiStatusHtml(exp)}</div>
    ${sent ? `<div class="banner success">${icon('check')}<span>Envoyé ${exp.sentAt ? 'le ' + fmtDate(new Date(exp.sentAt).toISOString().slice(0, 10)) : ''}${sentTo ? ' à ' + esc(sentTo.name) : ''}.</span><button type="button" class="btn btn-sm btn-secondary" data-act="unsend">Remettre à envoyer</button></div>` : ''}

    <div class="amount-hero">
      <input class="amount-input" data-f="total" inputmode="decimal" autocomplete="off" placeholder="0,00" value="${esc(numStr(num(exp.data.total)))}" aria-label="Montant TVAC">
      <span class="cur" data-cur>${esc(exp.data.currency === 'EUR' || !exp.data.currency ? '€' : exp.data.currency)}</span>
    </div>
    <p class="amount-label">Montant payé, TVA comprise</p>

    <div class="list form">
      <label class="field"><span>Date</span><input type="date" data-f="date" value="${esc(exp.data.date)}"></label>
      <label class="field"><span>Commerçant</span><input data-f="merchant" autocomplete="off" placeholder="Ex. Delhaize" value="${esc(exp.data.merchant)}"></label>
      <label class="field"><span>Description</span><input data-f="description" autocomplete="off" placeholder="Ex. Déjeuner client" value="${esc(exp.data.description)}"></label>
      <label class="field"><span>Catégorie</span><select data-f="category">${options(CATEGORIES, exp.data.category)}</select></label>
    </div>

    <details class="more">
      <summary>${icon('chevronDown')}TVA et détails</summary>
      <div class="list form">
        <label class="field"><span>Montant HTVA</span><input data-f="net" inputmode="decimal" autocomplete="off" placeholder="0,00" value="${esc(numStr(num(exp.data.net)))}"></label>
        <label class="field"><span>Montant TVA</span><input data-f="vat" inputmode="decimal" autocomplete="off" placeholder="0,00" value="${esc(numStr(num(exp.data.vat)))}"></label>
        <label class="field"><span>Taux de TVA</span><input data-f="vatRate" autocomplete="off" placeholder="Ex. 21 %" value="${esc(exp.data.vatRate)}"></label>
        <label class="field"><span>Type</span><select data-f="docType">${options(DOC_TYPES, exp.data.docType)}</select></label>
        <label class="field"><span>N° de facture</span><input data-f="invoiceNo" autocomplete="off" value="${esc(exp.data.invoiceNo)}"></label>
        <label class="field"><span>Paiement</span><select data-f="payment">${options(PAYMENTS, exp.data.payment)}</select></label>
        <label class="field"><span>Devise</span><select data-f="currency">${CURRENCIES.map(c => `<option ${c === (exp.data.currency || 'EUR') ? 'selected' : ''}>${c}</option>`).join('')}</select></label>
        <label class="field stack"><span>Remarque</span><textarea data-f="notes" rows="2" placeholder="Ex. invités : M. Dupont, Mme Martin">${esc(exp.data.notes)}</textarea></label>
      </div>
    </details>

    <h3 class="group-label">Envoyer à</h3>
    <div class="chips" data-chips>${clientChips()}</div>

    <div class="sheet-actions">
      <button type="button" class="btn btn-primary btn-lg" data-act="send">${icon('send')}${sent ? 'Renvoyer' : 'Envoyer maintenant'}</button>
      <button type="button" class="btn btn-secondary btn-lg" data-act="later">${sent ? 'Fermer' : 'Garder pour plus tard'}</button>
    </div>
    <button type="button" class="btn btn-plain danger delete-btn" data-act="delete">${icon('trash')}Supprimer ce justificatif</button>
  `;

  const body = sheet.body;
  const input = f => body.querySelector(`[data-f="${f}"]`);
  const fitAmount = () => {
    const el = input('total');
    el.style.width = `${Math.max(4, (el.value || el.placeholder).length + 0.5)}ch`;
  };
  fitAmount();

  function readForm() {
    for (const f of TEXT_FIELDS) exp.data[f] = input(f).value.trim();
    for (const f of MONEY_FIELDS) exp.data[f] = num(input(f).value);
  }

  async function persist(silent = false) {
    if (removed) return;
    const fresh = getExpense(id);
    if (!fresh) return;
    exp = fresh;
    readForm();
    await saveExpense(exp);
    if (!silent) toast('Enregistré', { type: 'success' });
  }

  // Quand l'IA a fini, on remplit les champs que l'utilisateur n'a pas modifiés.
  let lastStatus = exp.aiStatus;
  const unsubscribe = onChange((what, detail) => {
    if (what !== 'expenses' || detail.id !== id || removed) return;
    const fresh = getExpense(id);
    if (!fresh) return;
    const prev = lastStatus;
    lastStatus = fresh.aiStatus;
    body.querySelector('[data-ai]').innerHTML = aiStatusHtml(fresh);
    if (prev === 'pending' && fresh.aiStatus === 'done') {
      for (const f of [...TEXT_FIELDS, ...MONEY_FIELDS]) {
        if (touched.has(f)) continue;
        const v = fresh.data[f];
        const el = input(f);
        el.value = MONEY_FIELDS.includes(f) ? numStr(num(v)) : (v ?? '');
        el.classList.add('filled');
        setTimeout(() => el.classList.remove('filled'), 1200);
      }
      fitAmount();
      updateCurrency();
      const derived = amounts(fresh.data);
      if (derived.vat != null && !input('vat').value) input('vat').placeholder = numStr(derived.vat);
      if (derived.net != null && !input('net').value) input('net').placeholder = numStr(derived.net);
    }
  });

  const updateCurrency = () => {
    const c = input('currency').value;
    body.querySelector('[data-cur]').textContent = c === 'EUR' ? '€' : c;
  };

  body.oninput = e => {
    const f = e.target.dataset?.f;
    if (!f) return;
    touched.add(f);
    if (f === 'total') fitAmount();
    if (f === 'currency') updateCurrency();
  };
  body.onchange = body.oninput;

  body.onclick = async e => {
    const chip = e.target.closest('[data-client]');
    if (chip) {
      exp.clientId = exp.clientId === chip.dataset.client ? '' : chip.dataset.client;
      if (exp.clientId) saveSettings({ lastClientId: exp.clientId });
      body.querySelector('[data-chips]').innerHTML = clientChips();
      return;
    }
    const act = e.target.closest('[data-act]')?.dataset.act;
    if (!act) return;
    if (act === 'view') return viewFile(exp);
    if (act === 'reanalyze') {
      touched.clear();
      readForm();
      await saveExpense(exp);
      analyzeExpense(id, { overwrite: true });
      return;
    }
    if (act === 'later') { await persist(true); sheet.close(); if (!sent) toast('Gardé pour plus tard', { type: 'success' }); return; }
    if (act === 'unsend') {
      exp.status = 'pending';
      exp.sentAt = null;
      await persist(true);
      sheet.close();
      toast('Remis dans les justificatifs à envoyer');
      return;
    }
    if (act === 'delete') {
      const ok = await confirmDialog({ title: 'Supprimer ce justificatif ?', message: 'La photo et les informations seront effacées de cet appareil.', confirm: 'Supprimer', destructive: true });
      if (!ok) return;
      removed = true;
      await deleteExpense(id);
      sheet.close();
      toast('Justificatif supprimé');
      return;
    }
    if (act === 'send') {
      if (!exp.clientId) {
        toast('Choisissez d’abord à qui l’envoyer');
        const chips = body.querySelector('[data-chips]');
        chips.classList.remove('shake');
        void chips.offsetWidth;
        chips.classList.add('shake');
        chips.scrollIntoView({ behavior: 'smooth', block: 'center' });
        return;
      }
      await persist(true);
      const clientId = exp.clientId;
      sheet.close();
      openSendSheet({ clientId, focusId: id, includeSent: sent });
    }
  };
}

export function viewFile(exp) {
  const url = URL.createObjectURL(exp.file);
  if (exp.mime === 'application/pdf' || !exp.mime.startsWith('image/')) {
    window.open(url, '_blank');
    setTimeout(() => URL.revokeObjectURL(url), 60000);
    return;
  }
  openSheet({
    title: exp.data.merchant || 'Justificatif',
    className: 'viewer',
    body: `<div class="viewer-img"><img src="${url}" alt="Justificatif"></div>`,
    onClose: () => URL.revokeObjectURL(url),
  });
}
