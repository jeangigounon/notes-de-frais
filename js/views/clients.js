// Clients : liste et fiche (e-mail, format Excel attendu, pièces jointes).

import { state, activeClients, getClient, newClient, saveClient, deleteClient } from '../store.js';
import { esc, icon, toast, confirmDialog, debounce, download, isCompact, blobToBase64, pickFile, openSheet, spinner } from '../ui.js';
import { CLIENT_COLORS, LINE_FIELDS, PLACEHOLDERS, DEFAULT_SIMPLE_COLUMNS, fieldDef, sampleExpenses } from '../fields.js';
import { openWorkbook, readSheet, templateCapacity, convertXlsToXlsx } from '../xlsx.js';
import { makeReport, buildAttachments } from '../mail.js';
import { analyzeTemplate, canAnalyzeTemplates } from '../templateAI.js';
import { openMapper } from './mapper.js';
import { initials } from './home.js';

export const deps = ['clients'];

export function onStoreChange(what, detail, page, params) {
  if (what !== 'clients') return false;
  // La fiche ouverte se met à jour elle-même : on ne rafraîchit que la liste.
  if (detail.source === 'editor') {
    const list = page.querySelector('[data-client-list]');
    if (list) list.innerHTML = listHtml(params[0]);
    return true;
  }
  return false;
}

function listHtml(currentId) {
  const clients = activeClients();
  if (!clients.length) return '<div class="row"><span class="row-sub">Aucun client pour l’instant.</span></div>';
  return clients.map(c => `
    <a class="row row-btn ${c.id === currentId ? 'current' : ''}" href="#/clients/${c.id}">
      <span class="avatar" style="background:${CLIENT_COLORS[c.color] || CLIENT_COLORS.blue}">${esc(initials(c.name))}</span>
      <span class="row-main"><span class="row-title">${esc(c.name || 'Sans nom')}</span><span class="row-sub">${esc(c.to || 'Pas encore d’adresse')}</span></span>
      ${icon('chevronRight', 'chev')}
    </a>`).join('');
}

export function render(page, params) {
  const id = params[0];
  const client = id ? getClient(id) : null;
  if (id && !client) { location.hash = '#/clients'; return; }
  const compact = isCompact();

  if (compact && client) {
    page.innerHTML = `<a class="back-link" href="#/clients">${icon('chevronLeft')}Clients</a><div data-editor></div>`;
    renderEditor(page.querySelector('[data-editor]'), client);
    bindNew(page);
    return;
  }

  page.innerHTML = `
    <header class="page-head">
      <div><h1 class="large-title">Clients</h1><p class="subtitle">À qui envoyer vos notes de frais, et sous quelle forme.</p></div>
      <button type="button" class="btn btn-primary" data-act="new">${icon('plus')}Nouveau client</button>
    </header>
    <div class="split">
      <div class="split-list"><div class="list" data-client-list>${listHtml(id)}</div></div>
      <div class="split-detail" data-editor>${client ? '' : compact ? '' : `
        <div class="empty">${icon('users')}<h3>${activeClients().length ? 'Choisissez un client' : 'Ajoutez votre premier client'}</h3>
        <p>Pour chaque client : son adresse e-mail, le message type et le fichier Excel qu'il attend.</p>
        ${activeClients().length ? '' : `<button type="button" class="btn btn-primary" data-act="new">${icon('plus')}Nouveau client</button>`}</div>`}
      </div>
    </div>`;
  if (client) renderEditor(page.querySelector('[data-editor]'), client);
  bindNew(page);
}

function bindNew(page) {
  page.addEventListener('click', e => {
    if (!e.target.closest('[data-act="new"]')) return;
    const c = newClient();
    saveClient(c);
    location.hash = `#/clients/${c.id}`;
    setTimeout(() => document.querySelector('[data-k="name"]')?.focus(), 80);
  });
}

/* ---------- Fiche client ---------- */

function templateSummary(tpl) {
  const single = Object.keys(tpl.cells || {}).length;
  const lines = Object.keys(tpl.lineCols || {}).length;
  const parts = [];
  if (lines) parts.push(`tableau à partir de la ligne ${tpl.lineRow} (${lines} colonne${lines > 1 ? 's' : ''})`);
  if (single) parts.push(`${single} cellule${single > 1 ? 's' : ''} unique${single > 1 ? 's' : ''}`);
  return parts.length ? parts.join(' · ') : 'Aucune cellule configurée';
}

function renderEditor(host, client) {
  const c = structuredClone(client);
  c.excel ||= { mode: 'simple', fileName: 'Note de frais {nom} {periode}', simple: { title: 'Note de frais', columns: DEFAULT_SIMPLE_COLUMNS.map(x => ({ ...x })) }, template: null };
  c.excel.simple ||= { title: 'Note de frais', columns: DEFAULT_SIMPLE_COLUMNS.map(x => ({ ...x })) };

  const save = debounce(() => {
    try {
      saveClient(structuredClone(c), { source: 'editor' });
      const ind = host.querySelector('[data-saved]');
      if (ind) { ind.classList.add('show'); clearTimeout(ind._t); ind._t = setTimeout(() => ind.classList.remove('show'), 1600); }
    } catch (err) {
      toast(err.message, { type: 'error' });
    }
  }, 350);

  // Zone de dépôt du modèle (aussi proposée, en plus discret, en mode « Tableau standard »).
  const dropzone = (compact = false) => `
    <div class="dropzone ${compact ? 'compact' : ''}" data-dropzone data-act="upload">
      ${icon('table')}
      <strong>${compact ? 'Le client a son propre fichier Excel ? Déposez-le ici' : 'Déposez ici le fichier Excel du client'}</strong>
      <span>ou cliquez pour le choisir (.xlsx ou .xls). ${canAnalyzeTemplates()
        ? 'Claude l’analyse et propose où placer chaque information ; vous n’avez plus qu’à vérifier.'
        : 'Pour que Claude l’analyse automatiquement, <a href="#/reglages">ajoutez votre clé Claude dans Réglages</a> sur cet appareil (la clé est enregistrée appareil par appareil).'}</span>
    </div>`;

  const draw = () => {
    const mode = c.excel.mode;
    const tpl = c.excel.template;
    host.innerHTML = `
      <div class="editor-head">
        <span class="avatar xl" style="background:${CLIENT_COLORS[c.color] || CLIENT_COLORS.blue}">${esc(initials(c.name))}</span>
        <div class="editor-title">
          <h2>${esc(c.name || 'Nouveau client')}</h2>
          <span class="saved" data-saved>${icon('check')}Enregistré</span>
        </div>
      </div>

      <h3 class="group-label">Client</h3>
      <div class="list form">
        <label class="field"><span>Nom</span><input data-k="name" autocomplete="off" placeholder="Ex. Comptabilité ACME" value="${esc(c.name)}"></label>
        <label class="field"><span>Société</span><input data-k="company" autocomplete="off" placeholder="Facultatif" value="${esc(c.company)}"></label>
        <div class="field"><span>Couleur</span><div class="colors">${Object.entries(CLIENT_COLORS).map(([k, v]) => `<button type="button" class="${c.color === k ? 'on' : ''}" data-color="${k}" style="background:${v};color:${v}" aria-label="${k}"></button>`).join('')}</div></div>
      </div>

      <h3 class="group-label">E-mail</h3>
      <div class="list form">
        <label class="field"><span>À</span><input data-k="to" inputmode="email" autocomplete="off" placeholder="compta@client.be" value="${esc(c.to)}"></label>
        <label class="field"><span>Cc</span><input data-k="cc" inputmode="email" autocomplete="off" placeholder="Facultatif" value="${esc(c.cc)}"></label>
        <label class="field"><span>Objet</span><input data-k="subject" autocomplete="off" value="${esc(c.subject)}"></label>
        <label class="field stack"><span>Message</span><textarea data-k="body" rows="6">${esc(c.body)}</textarea></label>
      </div>
      <p class="group-foot">Variables remplacées automatiquement : ${PLACEHOLDERS.map(([k, d]) => `<button type="button" class="var" data-var="${k}" title="${esc(d)}">${k}</button>`).join(' ')}</p>

      <h3 class="group-label">Fichier Excel demandé</h3>
      <div class="seg block" role="tablist">
        <button type="button" data-mode="template" class="${mode === 'template' ? 'on' : ''}">Son modèle</button>
        <button type="button" data-mode="simple" class="${mode === 'simple' ? 'on' : ''}">Tableau standard</button>
        <button type="button" data-mode="none" class="${mode === 'none' ? 'on' : ''}">Pas d'Excel</button>
      </div>

      ${mode === 'template' ? (tpl ? `
        <div class="list tpl-card">
          <div class="row">
            <span class="att-icon xl">${icon('table')}</span>
            <span class="row-main"><span class="row-title">${esc(tpl.name)}</span><span class="row-sub">Feuille « ${esc(tpl.sheet)} » · ${esc(templateSummary(tpl))}</span></span>
          </div>
          ${tpl.ai?.summary ? `<div class="row ai-summary"><span class="att-icon ai">${icon('sparkles')}</span><span class="row-sub">${esc(tpl.ai.summary)}</span></div>` : ''}
          ${tpl.ai?.error ? `<div class="row"><span class="att-icon warn">${icon('alert')}</span><span class="row-sub">L'analyse par Claude a échoué : ${esc(tpl.ai.error)}</span></div>` : ''}
          ${!tpl.lineRow ? `<div class="row"><span class="att-icon warn">${icon('alert')}</span><span class="row-sub">Indiquez où commence le tableau des dépenses : lancez l'analyse par Claude, ou cliquez sur une cellule de la première ligne vide et choisissez un champ « Par justificatif ».</span></div>` : ''}
          <div class="row tpl-actions">
            <button type="button" class="btn ${tpl.lineRow ? 'btn-secondary' : 'btn-primary'}" data-act="analyze">${icon('sparkles')}${tpl.ai?.summary ? 'Analyser à nouveau' : 'Analyser avec Claude'}</button>
            <button type="button" class="btn ${tpl.lineRow ? 'btn-primary' : 'btn-secondary'}" data-act="map">${icon('grid')}Vérifier les cellules</button>
            <button type="button" class="btn btn-secondary" data-act="replace">Remplacer</button>
            <button type="button" class="btn btn-plain danger" data-act="remove-tpl">Retirer</button>
          </div>
        </div>` : dropzone()) : ''}
      ${mode !== 'template' ? dropzone(true) : ''}

      ${mode === 'simple' ? `
        <div class="list form">
          <label class="field"><span>Titre</span><input data-simple="title" autocomplete="off" value="${esc(c.excel.simple.title)}"></label>
        </div>
        <h3 class="group-label">Colonnes</h3>
        <div class="list cols-list">
          ${c.excel.simple.columns.map((col, i) => `
            <div class="field col-row">
              <span class="col-field">${esc(fieldDef(col.field)?.label || col.field)}</span>
              <input data-col-label="${i}" value="${esc(col.label)}" aria-label="Titre de la colonne">
              <span class="col-tools">
                <button type="button" class="btn-icon" data-col-up="${i}" ${i === 0 ? 'disabled' : ''} aria-label="Monter">${icon('arrowUp')}</button>
                <button type="button" class="btn-icon" data-col-down="${i}" ${i === c.excel.simple.columns.length - 1 ? 'disabled' : ''} aria-label="Descendre">${icon('arrowDown')}</button>
                <button type="button" class="btn-icon danger" data-col-del="${i}" aria-label="Retirer">${icon('x')}</button>
              </span>
            </div>`).join('')}
          <label class="field"><span>Ajouter</span>
            <select data-col-add><option value="">Choisir une colonne…</option>${LINE_FIELDS.filter(f => !c.excel.simple.columns.some(x => x.field === f.key)).map(f => `<option value="${f.key}">${esc(f.label)}</option>`).join('')}</select>
          </label>
        </div>` : ''}

      ${mode !== 'none' ? `
        <div class="list form" style="margin-top:16px">
          <label class="field"><span>Nom du fichier</span><input data-excel="fileName" autocomplete="off" value="${esc(c.excel.fileName)}"></label>
        </div>
        <button type="button" class="btn btn-secondary test-btn" data-act="sample">${icon('download')}Tester avec des données d'exemple</button>` : ''}

      <h3 class="group-label">Pièces jointes</h3>
      <div class="list form">
        <label class="field"><span>Joindre les photos et PDF des justificatifs</span><span class="switch"><input type="checkbox" data-k="attachReceipts" ${c.attachReceipts !== false ? 'checked' : ''}><span></span></span></label>
      </div>

      <h3 class="group-label">Notes</h3>
      <div class="list form">
        <label class="field stack"><textarea data-k="notes" rows="3" placeholder="Ex. à envoyer avant le 5 du mois, code projet…">${esc(c.notes)}</textarea></label>
      </div>

      <button type="button" class="btn btn-plain danger delete-btn" data-act="delete">${icon('trash')}Supprimer ce client</button>
    `;
  };

  draw();

  // Claude lit le modèle et propose le placement de chaque information.
  async function runAnalysis(book, sheet) {
    const tpl = c.excel.template;
    const wait = openSheet({
      title: 'Analyse du modèle',
      left: '',
      body: `<div class="analysis">${spinner()}<strong>Claude étudie le modèle…</strong><span>Structure, formules, colonnes par type de dépense et par devise. Environ 20 secondes.</span></div>`,
    });
    try {
      const content = await readSheet(book, sheet.path);
      const res = await analyzeTemplate(content, { sheetName: sheet.name, fileName: tpl.originalName || tpl.name, otherSheets: book.sheets.filter(s => s.path !== sheet.path).map(s => s.name) });
      Object.assign(tpl, { sheet: sheet.name, lineRow: res.lineRow, lineCols: res.lineCols, cells: res.cells, capacity: null, ai: { summary: res.summary, warnings: res.warnings, at: Date.now() } });
      if (res.fileName) c.excel.fileName = res.fileName;
      save.flush();
      return true;
    } catch (err) {
      console.error(err);
      tpl.ai = { error: err.message || String(err), at: Date.now() };
      save.flush();
      toast(`Analyse impossible : ${err.message}`, { type: 'error', duration: 6000 });
      return false;
    } finally {
      wait.close();
    }
  }

  async function analyzeCurrent() {
    if (!canAnalyzeTemplates()) {
      const go = await confirmDialog({
        title: 'Clé Claude absente sur cet appareil',
        message: 'La clé Claude est enregistrée appareil par appareil. Ajoutez-la dans Réglages (ou utilisez le QR code de configuration), puis relancez l’analyse.',
        confirm: 'Ouvrir Réglages',
      });
      if (go) location.hash = '#/reglages';
      return;
    }
    try {
      const book = await openWorkbook(c.excel.template.data);
      const sheet = book.sheets.find(s => s.name === c.excel.template.sheet) || book.sheets[0];
      const ok = await runAnalysis(book, sheet);
      draw();
      if (ok) openMapper(c, () => { save.flush(); draw(); });
    } catch (err) {
      toast(err.message, { type: 'error' });
    }
  }

  async function loadTemplate(file) {
    if (!file) return;
    const isXls = /\.xls$/i.test(file.name);
    if (!isXls && !/\.xlsx$/i.test(file.name)) { toast('Choisissez un fichier Excel (.xlsx ou .xls)', { type: 'error' }); return; }
    if (file.size > 1_500_000) { toast('Ce modèle est trop lourd (1,5 Mo maximum).', { type: 'error' }); return; }
    try {
      let data;
      let name = file.name;
      if (isXls) {
        const ok = await confirmDialog({
          title: 'Ancien format Excel (.xls)',
          message: 'Il sera converti en .xlsx : formules, fusions, largeurs et formats de nombres sont conservés, mais pas les couleurs ni les bordures. Pour tout garder, ouvrez-le dans Excel, enregistrez-le en « Classeur Excel (.xlsx) » et déposez cette version.',
          confirm: 'Convertir',
        });
        if (!ok) return;
        data = await convertXlsToXlsx(file);
        name = file.name.replace(/\.xls$/i, '.xlsx');
      } else {
        data = await blobToBase64(file);
      }
      const book = await openWorkbook(data);
      const visible = book.sheets.find(s => !s.hidden) || book.sheets[0];
      const prev = c.excel.template;
      const keep = !canAnalyzeTemplates() && prev;
      c.excel.template = { name, originalName: file.name, data, sheet: visible.name, cells: keep ? prev.cells : {}, lineRow: keep ? prev.lineRow : null, lineCols: keep ? prev.lineCols : {}, capacity: null };
      c.excel.mode = 'template';
      if (!c.excel.fileName || c.excel.fileName === 'Note de frais {nom} {periode}') c.excel.fileName = file.name.replace(/\.xlsx?$/i, '') + ' {periode}';
      save.flush();
      draw();
      if (canAnalyzeTemplates()) await runAnalysis(book, visible);
      draw();
      openMapper(c, () => { save.flush(); draw(); });
    } catch (err) {
      console.error(err);
      toast(err.message, { type: 'error', duration: 5000 });
    }
  }

  host.oninput = e => {
    const t = e.target;
    if (t.dataset.k) {
      c[t.dataset.k] = t.type === 'checkbox' ? t.checked : t.value;
      if (t.dataset.k === 'name') {
        host.querySelector('.editor-title h2').textContent = t.value || 'Nouveau client';
        host.querySelector('.editor-head .avatar').textContent = initials(t.value);
      }
    } else if (t.dataset.simple) c.excel.simple[t.dataset.simple] = t.value;
    else if (t.dataset.excel) c.excel[t.dataset.excel] = t.value;
    else if (t.dataset.colLabel != null) c.excel.simple.columns[+t.dataset.colLabel].label = t.value;
    else return;
    save();
  };

  host.onchange = e => {
    const t = e.target;
    if (t.matches('[data-col-add]') && t.value) {
      const def = fieldDef(t.value);
      c.excel.simple.columns.push({ field: def.key, label: def.label });
      save.flush();
      draw();
    } else if (t.type === 'checkbox') host.oninput(e);
  };

  host.onclick = async e => {
    if (e.target.closest('a')) return;
    const t = e.target.closest('button, [data-act]');
    if (!t) return;
    if (t.dataset.color) { c.color = t.dataset.color; save.flush(); draw(); return; }
    if (t.dataset.mode) { c.excel.mode = t.dataset.mode; save.flush(); draw(); return; }
    if (t.dataset.var) {
      const area = host.querySelector('[data-k="body"]');
      const pos = area.selectionStart ?? area.value.length;
      area.value = area.value.slice(0, pos) + t.dataset.var + area.value.slice(area.selectionEnd ?? pos);
      area.focus();
      area.setSelectionRange(pos + t.dataset.var.length, pos + t.dataset.var.length);
      c.body = area.value;
      save();
      return;
    }
    const cols = c.excel.simple.columns;
    if (t.dataset.colUp != null) { const i = +t.dataset.colUp; [cols[i - 1], cols[i]] = [cols[i], cols[i - 1]]; save.flush(); draw(); return; }
    if (t.dataset.colDown != null) { const i = +t.dataset.colDown; [cols[i + 1], cols[i]] = [cols[i], cols[i + 1]]; save.flush(); draw(); return; }
    if (t.dataset.colDel != null) { cols.splice(+t.dataset.colDel, 1); save.flush(); draw(); return; }

    const act = t.dataset.act;
    if (act === 'upload' || act === 'replace') return loadTemplate(await pickFile('.xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel'));
    if (act === 'map') return openMapper(c, () => { save.flush(); draw(); });
    if (act === 'analyze') return analyzeCurrent();
    if (act === 'remove-tpl') {
      if (await confirmDialog({ title: 'Retirer le modèle ?', message: 'La configuration des cellules sera perdue.', confirm: 'Retirer', destructive: true })) {
        c.excel.template = null;
        save.flush();
        draw();
      }
      return;
    }
    if (act === 'sample') {
      try {
        if (c.excel.mode === 'template' && !c.excel.template) { toast('Ajoutez d’abord le modèle du client'); return; }
        const rep = makeReport({ ...c, attachReceipts: false }, sampleExpenses());
        const { excelFiles, warnings } = await buildAttachments({ ...c, attachReceipts: false }, rep);
        for (const f of excelFiles) download(f.blob, f.name.replace(/\.xlsx$/, ' (exemple).xlsx'));
        warnings.forEach(w => toast(w));
      } catch (err) {
        toast(err.message, { type: 'error' });
      }
      return;
    }
    if (act === 'delete') {
      if (await confirmDialog({ title: `Supprimer ${c.name || 'ce client'} ?`, message: 'Ses justificatifs restent dans vos dépenses, sans destinataire.', confirm: 'Supprimer', destructive: true })) {
        save.cancel();
        deleteClient(c.id);
        location.hash = '#/clients';
        toast('Client supprimé');
      }
    }
  };

  // Un fichier Excel déposé n'importe où sur la fiche devient le modèle du client.
  const zone = () => host.querySelector('[data-dropzone]');
  host.addEventListener('dragover', e => {
    if (![...(e.dataTransfer?.types || [])].includes('Files')) return;
    e.preventDefault();
    zone()?.classList.add('drag');
  });
  host.addEventListener('dragleave', e => { if (!host.contains(e.relatedTarget)) zone()?.classList.remove('drag'); });
  host.addEventListener('drop', e => {
    const files = [...(e.dataTransfer?.files || [])];
    if (!files.length) return;
    e.preventDefault();
    zone()?.classList.remove('drag');
    const excel = files.find(f => /\.xlsx?$/i.test(f.name));
    if (excel) loadTemplate(excel);
    else toast('Déposez ici le fichier Excel du client (.xlsx ou .xls). Les tickets s’ajoutent depuis l’accueil.');
  });
}

export { templateCapacity, readSheet };
