// Envoi d'une note de frais à un client : sélection des justificatifs, message, pièces jointes, envoi.

import { state, getClient, activeClients, getExpense, saveExpense } from '../store.js';
import { esc, icon, openSheet, actionSheet, confirmDialog, toast, money, fmtDate, download, spinner, isTouch, debounce, todayISO } from '../ui.js';
import { CLIENT_COLORS } from '../fields.js';
import {
  makeReport, composeMessage, buildAttachments, preferredMethod, sendViaGraph, msSignIn, shareFiles,
  canShareFiles, buildEml, buildZip, splitEmails, isEmail, ensureRates,
} from '../mail.js';
import { thumb } from './expenses.js';
import { ensureReceiptFile } from '../sync.js';

const METHOD_LABEL = {
  graph: 'Envoi direct depuis votre boîte Outlook',
  share: 'Ouvre le menu Partager : choisissez Outlook ou Mail',
  eml: 'Télécharge un brouillon à ouvrir dans Outlook, prêt à partir',
  zip: 'Télécharge une archive avec tous les fichiers',
};

export async function openSendSheet({ clientId, focusId = null, includeSent = false, preset = null } = {}) {
  let client = getClient(clientId);
  if (!client) {
    const clients = activeClients();
    if (!clients.length) { toast('Ajoutez d’abord un client'); location.hash = '#/clients'; return; }
    const picked = await actionSheet({ title: 'Envoyer à…', actions: clients.map(c => ({ label: c.name || 'Sans nom', value: c.id })) });
    if (!picked) return;
    client = getClient(picked);
  }

  const candidateIds = state.expenses.filter(e => e.clientId === client.id && e.status !== 'sent').map(e => e.id);
  if (focusId && !candidateIds.includes(focusId)) candidateIds.unshift(focusId);
  const candidates = () => candidateIds.map(getExpense).filter(Boolean);
  const selected = new Set(preset?.expenseIds || (includeSent && focusId ? [focusId] : candidateIds));

  const edited = { subject: !!preset?.subject, body: !!preset?.body };
  let msg = { to: [], cc: [], subject: preset?.subject || '', body: preset?.body || '' };
  let prepared = null;
  let preparing = null;
  let prepareError = '';
  let busy = false;

  const sheet = openSheet({
    title: `Note de frais · ${client.name || 'client'}`,
    left: '<button type="button" class="btn-text" data-sheet-close>Annuler</button>',
  });
  const body = sheet.body;

  const selection = () => candidates().filter(e => selected.has(e.id));
  const report = () => makeReport(client, selection());

  function recompose() {
    const rep = report();
    const fresh = composeMessage(client, rep);
    msg.to = preset?.to ? splitEmails(preset.to) : msg.to.length ? msg.to : fresh.to;
    msg.cc = preset?.cc != null ? splitEmails(preset.cc) : msg.cc.length ? msg.cc : fresh.cc;
    if (!edited.subject) msg.subject = fresh.subject;
    if (!edited.body) msg.body = fresh.body;
    return rep;
  }

  function render() {
    const rep = recompose();
    const items = candidates();
    const method = preferredMethod();
    body.innerHTML = `
      <div class="send-head">
        <span class="avatar lg" style="background:${CLIENT_COLORS[client.color] || CLIENT_COLORS.blue}">${esc((client.name || '?').slice(0, 1).toUpperCase())}</span>
        <div><div class="big-amount">${money(rep.displayTotal, rep.displayCurrency)}</div><div class="row-sub">${rep.count} justificatif${rep.count > 1 ? 's' : ''} · ${esc(rep.period)}${rep.displayCurrency === 'EUR' && rep.items.some(i => i.currency !== 'EUR') ? ' · converti en euros' : ''}</div></div>
      </div>

      <div class="list form">
        <label class="field"><span>À</span><input data-m="to" inputmode="email" autocomplete="off" placeholder="adresse@client.be" value="${esc(msg.to.join(', '))}"></label>
        <label class="field"><span>Cc</span><input data-m="cc" inputmode="email" autocomplete="off" value="${esc(msg.cc.join(', '))}"></label>
        <label class="field"><span>Objet</span><input data-m="subject" autocomplete="off" value="${esc(msg.subject)}"></label>
        <label class="field stack"><span>Message</span><textarea data-m="body" rows="6">${esc(msg.body)}</textarea></label>
      </div>

      <h3 class="group-label"><span>Justificatifs</span>${items.length > 1 ? `<button type="button" class="link" data-act="toggle-all">${selected.size === items.length ? 'Tout désélectionner' : 'Tout sélectionner'}</button>` : ''}</h3>
      <div class="list inset-thumb">
        ${items.map(e => `
          <button type="button" class="row row-btn" data-pick="${e.id}">
            <span class="check ${selected.has(e.id) ? 'on' : ''}">${icon('check')}</span>
            ${thumb(e)}
            <span class="row-main"><span class="row-title">${esc(e.data.merchant || 'Justificatif')}</span><span class="row-sub">${esc([fmtDate(e.data.date), e.data.category].filter(Boolean).join(' · '))}</span></span>
            <span class="amount">${money(e.data.total, e.data.currency)}</span>
          </button>`).join('') || '<div class="row"><span class="row-sub">Aucun justificatif en attente pour ce client.</span></div>'}
      </div>
      ${items.some(e => e.aiStatus === 'pending') ? `<p class="group-foot">${spinner()} Analyse en cours : attendez qu'elle se termine pour envoyer des montants complets.</p>` : ''}

      <h3 class="group-label">Pièces jointes</h3>
      <div class="list" data-attachments>${attachmentsHtml()}</div>

      <div class="sheet-actions">
        <button type="button" class="btn btn-primary btn-lg" data-act="send" ${!selected.size || !prepared || busy ? 'disabled' : ''}>${busy ? spinner('light') : icon('send')}${busy ? 'Envoi…' : 'Envoyer'}</button>
        <button type="button" class="btn btn-secondary btn-lg" data-act="more" ${!prepared ? 'disabled' : ''}>Autres options</button>
      </div>
      <p class="method-hint">${icon(method === 'graph' ? 'mail' : method === 'share' ? 'share' : 'download')}${esc(METHOD_LABEL[method])}</p>
    `;
  }

  function attachmentsHtml() {
    if (!selected.size) return '<div class="row"><span class="row-sub">Sélectionnez au moins un justificatif.</span></div>';
    if (prepareError) return `<div class="row"><span class="row-sub danger">${esc(prepareError)}</span></div>`;
    if (!prepared) return `<div class="row">${spinner()}<span class="row-sub">Préparation du fichier Excel…</span></div>`;
    const excel = prepared.excelFiles.map((f, i) => `
      <button type="button" class="row row-btn" data-preview="${i}">
        <span class="att-icon xl">${icon('table')}</span>
        <span class="row-main"><span class="row-title">${esc(f.name)}</span><span class="row-sub">Excel · ${selected.size} ligne${selected.size > 1 ? 's' : ''}${prepared.added ? ` (dont ${prepared.added} ajoutée${prepared.added > 1 ? 's' : ''} au modèle)` : ''} · toucher pour vérifier</span></span>
        ${icon('download', 'chev')}
      </button>`).join('');
    const photos = prepared.receipts.length ? `
      <div class="row">
        <span class="att-icon">${icon('image')}</span>
        <span class="row-main"><span class="row-title">${prepared.receipts.length} justificatif${prepared.receipts.length > 1 ? 's' : ''}</span><span class="row-sub">${Math.round(prepared.receipts.reduce((s, f) => s + f.blob.size, 0) / 1024)} Ko · photos et PDF</span></span>
      </div>` : '';
    const warn = prepared.warnings.map(w => `<div class="row"><span class="att-icon warn">${icon('alert')}</span><span class="row-sub">${esc(w)}</span></div>`).join('');
    return excel + photos + warn || '<div class="row"><span class="row-sub">Aucune pièce jointe</span></div>';
  }

  const refreshAttachments = () => {
    const el = body.querySelector('[data-attachments]');
    if (el) el.innerHTML = attachmentsHtml();
    body.querySelector('[data-act="send"]')?.toggleAttribute('disabled', !selected.size || !prepared || busy);
    body.querySelector('[data-act="more"]')?.toggleAttribute('disabled', !prepared);
  };

  // Les fichiers sont préparés à l'avance : le menu Partager exige d'être ouvert sans délai après le geste.
  const prepare = debounce(async () => {
    prepared = null;
    prepareError = '';
    refreshAttachments();
    if (!selected.size) return;
    const job = (async () => {
      // Photos prises sur l'autre appareil : on les récupère avant de construire l'e-mail.
      await Promise.all(selection().map(e => ensureReceiptFile(e).catch(() => null)));
      await ensureRates(selection());
      return buildAttachments(client, report());
    })();
    preparing = job;
    try {
      const result = await job;
      if (preparing === job) prepared = result;
    } catch (err) {
      console.error(err);
      if (preparing === job) prepareError = `Impossible de préparer le fichier Excel : ${err.message}`;
    }
    if (preparing === job) refreshAttachments();
  }, 150);

  render();
  prepare();

  body.oninput = e => {
    const k = e.target.dataset?.m;
    if (!k) return;
    if (k === 'to' || k === 'cc') msg[k] = splitEmails(e.target.value);
    else { msg[k] = e.target.value; edited[k] = true; }
  };

  async function markSent(via) {
    const ids = [...selected];
    for (const id of ids) {
      const e = getExpense(id);
      if (!e) continue;
      e.status = 'sent';
      e.sentAt = Date.now();
      e.sentTo = client.id;
      e.sentVia = via;
      e.sentToEmails = msg.to.join(', ');
      await saveExpense(e);
    }
    const rep = report();
    body.innerHTML = `
      <div class="success">
        <div class="success-icon">${icon('check')}</div>
        <h2>${via === 'graph' ? 'Note de frais envoyée' : 'C’est noté'}</h2>
        <p>${ids.length} justificatif${ids.length > 1 ? 's' : ''} · ${money(rep.displayTotal, rep.displayCurrency)}<br>${via === 'graph' ? `envoyés à ${esc(msg.to.join(', '))}` : 'marqués comme envoyés'}</p>
        <button type="button" class="btn btn-primary btn-lg btn-block" data-sheet-close>Terminé</button>
      </div>`;
    sheet.setLeft('');
  }

  function validRecipients() {
    if (!msg.to.length) { toast('Indiquez au moins un destinataire', { type: 'error' }); body.querySelector('[data-m="to"]')?.focus(); return false; }
    const bad = [...msg.to, ...msg.cc].find(a => !isEmail(a));
    if (bad) { toast(`Adresse invalide : ${bad}`, { type: 'error' }); return false; }
    return true;
  }

  async function askMarkSent(via) {
    const ok = await confirmDialog({ title: 'L’e-mail est-il parti ?', message: 'Marquer ces justificatifs comme envoyés pour ne plus les voir dans « À envoyer ».', confirm: 'Oui, envoyé', cancel: 'Pas encore' });
    if (ok) await markSent(via);
  }

  async function run(method) {
    if (!prepared || busy) return;
    const files = prepared.files;
    if (method === 'share') {
      if (!canShareFiles(files)) { toast('Le partage de fichiers n’est pas disponible ici : brouillon Outlook à la place.'); return run('eml'); }
      try {
        await shareFiles(msg, files);
        if (msg.to.length) toast('Adresse du destinataire copiée : collez-la dans le champ « À »');
        await askMarkSent('share');
      } catch (err) {
        if (err.name !== 'AbortError') toast(`Partage impossible : ${err.message}`, { type: 'error' });
      }
      return;
    }
    if (method === 'excel') {
      for (const f of prepared.excelFiles) download(f.blob, f.name);
      return;
    }
    if (method === 'zip') {
      download(await buildZip(files), `${prepared.excelFiles[0]?.name.replace(/\.xlsx$/, '') || 'Note de frais ' + todayISO()}.zip`);
      return askMarkSent('zip');
    }
    if (!validRecipients()) return;
    if (method === 'eml') {
      download(await buildEml(msg, files), `${msg.subject.replace(/[\\/:*?"<>|]+/g, ' ').slice(0, 80) || 'Note de frais'}.eml`);
      toast('Ouvrez le fichier téléchargé : il s’ouvre dans Outlook, prêt à envoyer.', { duration: 6000 });
      return askMarkSent('eml');
    }
    if (method === 'graph') {
      if (!state.settings.msClientId) { toast('Configurez Microsoft 365 dans Réglages pour l’envoi direct'); return run(isTouch() ? 'share' : 'eml'); }
      busy = true;
      render();
      refreshAttachments();
      try {
        const res = await sendViaGraph(msg, files);
        if (res.needsLogin) {
          busy = false;
          render();
          refreshAttachments();
          const ok = await confirmDialog({ title: 'Connexion à Outlook', message: 'Connectez-vous une fois à votre compte Microsoft pour envoyer directement depuis votre adresse.', confirm: 'Se connecter' });
          if (ok) {
            await msSignIn({ kind: 'send', clientId: client.id, expenseIds: [...selected], to: msg.to.join(', '), cc: msg.cc.join(', '), subject: msg.subject, body: msg.body });
          }
          return;
        }
        await markSent('graph');
      } catch (err) {
        busy = false;
        render();
        refreshAttachments();
        toast(err.message, { type: 'error', duration: 6000 });
      }
    }
  }

  body.onclick = async e => {
    const pick = e.target.closest('[data-pick]');
    if (pick) {
      const id = pick.dataset.pick;
      if (selected.has(id)) selected.delete(id); else selected.add(id);
      render();
      prepare();
      return;
    }
    const prev = e.target.closest('[data-preview]');
    if (prev && prepared) { const f = prepared.excelFiles[+prev.dataset.preview]; download(f.blob, f.name); return; }
    const act = e.target.closest('[data-act]')?.dataset.act;
    if (act === 'toggle-all') {
      const all = candidates();
      if (selected.size === all.length) selected.clear(); else all.forEach(x => selected.add(x.id));
      render();
      prepare();
      return;
    }
    if (act === 'send') return run(preferredMethod());
    if (act === 'more') {
      const shareable = prepared && canShareFiles(prepared.files);
      const choice = await actionSheet({
        title: 'Envoyer autrement',
        actions: [
          state.settings.msClientId && { label: 'Envoyer directement (Outlook)', value: 'graph', icon: 'mail' },
          shareable && { label: 'Partager (Outlook, Mail…)', value: 'share', icon: 'share' },
          { label: 'Brouillon Outlook (.eml)', value: 'eml', icon: 'mail' },
          { label: 'Télécharger tout (.zip)', value: 'zip', icon: 'archive' },
          prepared?.excelFiles.length && { label: 'Télécharger seulement l’Excel', value: 'excel', icon: 'table' },
          { label: 'Marquer comme envoyé', value: 'mark', icon: 'check' },
        ].filter(Boolean),
      });
      if (choice === 'mark') return markSent('manual');
      if (choice) run(choice);
    }
  };
}
