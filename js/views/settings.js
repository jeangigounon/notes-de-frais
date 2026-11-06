// Réglages : profil, reconnaissance, envoi des e-mails, synchronisation, configuration du téléphone, sauvegarde.

import { state, saveProfile, saveSettings } from '../store.js';
import { esc, icon, toast, debounce, openSheet, confirmDialog, timeAgo, spinner, pickFile } from '../ui.js';
import { MODELS, testClaudeKey } from '../ai.js';
import { redirectUri, msAccount, msSignIn, msSignOut } from '../mail.js';
import { syncNow } from '../sync.js';
import { setupLink, qrSvg, exportBackup, importBackup } from '../transfer.js';

export const deps = ['settings', 'profile', 'sync'];

export function onStoreChange(what, detail, page) {
  // Pas de rafraîchissement complet pendant la saisie (le champ perdrait le focus).
  if (what === 'sync') { const el = page.querySelector('[data-sync-status]'); if (el) el.innerHTML = syncStatusHtml(); return true; }
  if (page.contains(document.activeElement) && document.activeElement.matches('input, textarea, select')) return true;
  return false;
}

function syncStatusHtml() {
  const s = state.settings;
  if (!s.ghToken) return '<span class="status-dot"></span>Désactivée';
  if (state.sync.status === 'syncing') return `${spinner()}Synchronisation…`;
  if (state.sync.status === 'error') return `<span class="status-dot err"></span>${esc(state.sync.message)}`;
  return `<span class="status-dot ok"></span>Synchronisé ${timeAgo(s.lastSync)}`;
}

export function render(page) {
  const p = state.profile;
  const s = state.settings;
  const account = msAccount();
  const ghTokenUrl = 'https://github.com/settings/tokens/new?scopes=gist&description=Notes%20de%20frais';

  page.innerHTML = `
    <header class="page-head"><div><h1 class="large-title">Réglages</h1></div></header>

    <h3 class="group-label">Mon profil</h3>
    <div class="list form">
      <label class="field"><span>Nom</span><input data-p="name" autocomplete="name" placeholder="Prénom Nom" value="${esc(p.name)}"></label>
      <label class="field"><span>E-mail</span><input data-p="email" type="email" autocomplete="email" placeholder="vous@societe.be" value="${esc(p.email)}"></label>
      <label class="field"><span>Société</span><input data-p="company" autocomplete="organization" placeholder="Société ou service" value="${esc(p.company)}"></label>
      <label class="field"><span>IBAN</span><input data-p="iban" autocomplete="off" placeholder="Pour le remboursement (facultatif)" value="${esc(p.iban)}"></label>
    </div>
    <p class="group-foot">Ces informations peuvent être placées dans les fichiers Excel de vos clients.</p>

    <h3 class="group-label">Reconnaissance des tickets</h3>
    <div class="seg block">
      <button type="button" data-ai="claude" class="${s.aiMode === 'claude' ? 'on' : ''}">${icon('sparkles')}IA Claude</button>
      <button type="button" data-ai="local" class="${s.aiMode === 'local' ? 'on' : ''}">Locale (gratuite)</button>
    </div>
    ${s.aiMode === 'claude' ? `
      <div class="list form" style="margin-top:12px">
        <label class="field"><span>Clé API</span><input data-s="anthropicKey" type="password" autocomplete="off" spellcheck="false" placeholder="sk-ant-…" value="${esc(s.anthropicKey)}"></label>
        <label class="field"><span>Modèle</span><select data-s="model">${MODELS.map(m => `<option value="${m.id}" ${m.id === s.model ? 'selected' : ''}>${esc(m.label)}</option>`).join('')}</select></label>
        <div class="row"><span class="row-sub">${esc(MODELS.find(m => m.id === s.model)?.hint || '')}</span><button type="button" class="btn btn-sm btn-secondary" data-act="test-key" ${s.anthropicKey ? '' : 'disabled'}>Tester la clé</button></div>
      </div>
      <p class="group-foot">Créez une clé sur <a href="https://console.anthropic.com/settings/keys" target="_blank" rel="noopener">console.anthropic.com</a> (rubrique API Keys) et ajoutez quelques euros de crédit. La clé reste enregistrée sur cet appareil uniquement. Sans clé, la lecture locale (moins fiable) est utilisée.</p>` : `
      <p class="group-foot">La lecture locale fonctionne sans compte, directement sur l'appareil, mais elle est approximative : vérifiez toujours le montant et la date.</p>`}

    <h3 class="group-label">Envoi des e-mails</h3>
    <div class="list form">
      <label class="field"><span>Méthode</span><select data-s="sendMethod">
        <option value="auto" ${s.sendMethod === 'auto' ? 'selected' : ''}>Automatique</option>
        <option value="graph" ${s.sendMethod === 'graph' ? 'selected' : ''}>Microsoft 365 (envoi direct)</option>
        <option value="share" ${s.sendMethod === 'share' ? 'selected' : ''}>Menu Partager (téléphone)</option>
        <option value="eml" ${s.sendMethod === 'eml' ? 'selected' : ''}>Brouillon Outlook (.eml)</option>
      </select></label>
    </div>
    <p class="group-foot">Automatique : Microsoft 365 s'il est configuré, sinon le menu Partager sur téléphone et un brouillon Outlook sur PC.</p>

    <h3 class="group-label">Microsoft 365 · envoi direct depuis votre adresse</h3>
    <div class="list form">
      <label class="field"><span>ID d'application</span><input data-s="msClientId" autocomplete="off" spellcheck="false" placeholder="00000000-0000-0000-0000-000000000000" value="${esc(s.msClientId)}"></label>
      <label class="field"><span>Organisation</span><input data-s="msTenant" autocomplete="off" spellcheck="false" placeholder="${esc(p.email.split('@')[1] || 'votre-domaine.be')}" value="${esc(s.msTenant)}"></label>
      <div class="row">
        <span class="row-main">${account ? `<span class="row-title">${esc(account.name || account.email)}</span><span class="row-sub">${esc(account.email)} · connecté</span>` : '<span class="row-sub">Non connecté</span>'}</span>
        ${account ? '<button type="button" class="btn btn-sm btn-secondary" data-act="ms-out">Se déconnecter</button>' : `<button type="button" class="btn btn-sm btn-primary" data-act="ms-in" ${s.msClientId ? '' : 'disabled'}>Se connecter</button>`}
      </div>
    </div>
    <details class="help">
      <summary>${icon('info')}Comment activer l'envoi direct ?</summary>
      <ol>
        <li>Ouvrez <a href="https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade" target="_blank" rel="noopener">entra.microsoft.com → Inscriptions d'applications</a> et cliquez sur <em>Nouvelle inscription</em>.</li>
        <li>Nom : « Notes de frais ». Types de comptes : <em>Comptes dans cet annuaire d'organisation uniquement</em>.</li>
        <li>URI de redirection : plateforme <em>Application monopage (SPA)</em>, adresse :<br><code class="code" data-copy>${esc(redirectUri())}</code></li>
        <li>Copiez l'<em>ID d'application (client)</em> ci-dessus. Dans <em>Autorisations de l'API</em>, vérifiez que <em>Microsoft Graph › Mail.Send</em> (déléguée) est présente.</li>
        <li>Cliquez sur « Se connecter ». Si Microsoft demande l'accord d'un administrateur, transmettez cette page à votre service informatique, ou utilisez le menu Partager / le brouillon Outlook.</li>
      </ol>
    </details>

    <h3 class="group-label">Synchronisation PC ↔ téléphone</h3>
    <div class="list form">
      <label class="field"><span>Jeton GitHub</span><input data-s="ghToken" type="password" autocomplete="off" spellcheck="false" placeholder="ghp_…" value="${esc(s.ghToken)}"></label>
      <div class="row"><span class="row-sub sync-status" data-sync-status>${syncStatusHtml()}</span><button type="button" class="btn btn-sm btn-secondary" data-act="sync" ${s.ghToken ? '' : 'disabled'}>${icon('sync')}Synchroniser</button></div>
    </div>
    <p class="group-foot">Votre profil et vos clients (modèles Excel compris) sont copiés dans un Gist secret de votre compte GitHub. <a href="${ghTokenUrl}" target="_blank" rel="noopener">Créer un jeton</a> avec uniquement l'autorisation « gist ». Les photos des tickets restent sur l'appareil qui les a prises.</p>

    <h3 class="group-label">Configurer mon téléphone</h3>
    <div class="list">
      <button type="button" class="row row-btn" data-act="qr">
        <span class="att-icon">${icon('qr')}</span>
        <span class="row-main"><span class="row-title">Afficher le QR code de configuration</span><span class="row-sub">Scannez-le avec l'appareil photo du téléphone : clé Claude, synchronisation et profil sont copiés.</span></span>
        ${icon('chevronRight', 'chev')}
      </button>
    </div>
    <details class="help">
      <summary>${icon('phone')}Installer l'application sur l'iPhone</summary>
      <ol>
        <li>Ouvrez ce site dans Safari.</li>
        <li>Touchez ${icon('share')} <em>Partager</em>, puis <em>Sur l'écran d'accueil</em>.</li>
        <li>L'icône « Frais » s'ouvre en plein écran, comme une app. Vos tickets y sont conservés durablement.</li>
      </ol>
    </details>

    <h3 class="group-label">Sauvegarde</h3>
    <div class="list">
      <button type="button" class="row row-btn" data-act="export"><span class="att-icon">${icon('download')}</span><span class="row-main"><span class="row-title">Exporter une sauvegarde</span><span class="row-sub">Clients, profil et justificatifs avec leurs photos (.json)</span></span></button>
      <button type="button" class="row row-btn" data-act="import"><span class="att-icon">${icon('upload')}</span><span class="row-main"><span class="row-title">Importer une sauvegarde</span><span class="row-sub">Les éléments déjà présents sont conservés</span></span></button>
    </div>
    <p class="group-foot" data-storage></p>
  `;

  navigator.storage?.estimate?.().then(({ usage }) => {
    const el = page.querySelector('[data-storage]');
    if (el && usage) el.textContent = `${state.expenses.length} justificatif${state.expenses.length > 1 ? 's' : ''} · ${(usage / 1048576).toFixed(1).replace('.', ',')} Mo utilisés sur cet appareil.`;
  });

  const saveP = debounce((k, v) => saveProfile({ [k]: v }), 400);
  const saveS = debounce(patch => saveSettings(patch), 300);
  const pendingSettings = {};

  page.oninput = e => {
    const t = e.target;
    if (t.dataset.p) saveP(t.dataset.p, t.value.trim());
    if (t.dataset.s) {
      pendingSettings[t.dataset.s] = t.value.trim();
      saveS({ ...pendingSettings });
    }
  };
  page.onchange = e => {
    const t = e.target;
    if (t.dataset.s) {
      saveS.flush({ ...pendingSettings, [t.dataset.s]: t.value.trim() });
      if (t.dataset.s === 'ghToken' && t.value.trim()) syncNow().then(ok => ok && toast('Synchronisation activée', { type: 'success' }));
      if (t.dataset.s === 'model') render(page);
      // Champs texte : on active les boutons sans tout redessiner (le clic en cours serait perdu).
      const enable = { anthropicKey: 'test-key', msClientId: 'ms-in', ghToken: 'sync' }[t.dataset.s];
      if (enable) page.querySelector(`[data-act="${enable}"]`)?.toggleAttribute('disabled', !t.value.trim());
    }
  };

  page.onclick = async e => {
    const ai = e.target.closest('[data-ai]');
    if (ai) { saveSettings({ aiMode: ai.dataset.ai }); render(page); return; }
    const copy = e.target.closest('[data-copy]');
    if (copy) { navigator.clipboard?.writeText(copy.textContent).then(() => toast('Adresse copiée', { type: 'success' })); return; }
    const act = e.target.closest('[data-act]')?.dataset.act;
    if (!act) return;
    if (act === 'test-key') {
      const btn = e.target.closest('button');
      btn.disabled = true;
      btn.innerHTML = spinner();
      try {
        await testClaudeKey(state.settings.anthropicKey, state.settings.model);
        toast('Clé valide : la reconnaissance par Claude est prête', { type: 'success' });
      } catch (err) {
        toast(err.message, { type: 'error', duration: 5000 });
      }
      btn.disabled = false;
      btn.textContent = 'Tester la clé';
    }
    if (act === 'ms-in') { try { await msSignIn(null); } catch (err) { toast(err.message, { type: 'error' }); } }
    if (act === 'ms-out') { msSignOut(); render(page); toast('Déconnecté de Microsoft 365'); }
    if (act === 'sync') {
      const ok = await syncNow();
      toast(ok ? 'Synchronisé' : state.sync.message, { type: ok ? 'success' : 'error' });
    }
    if (act === 'qr') return showQr();
    if (act === 'export') {
      try { await exportBackup(); } catch (err) { toast(err.message, { type: 'error' }); }
    }
    if (act === 'import') {
      const file = await pickFile('application/json,.json');
      if (!file) return;
      try {
        const res = await importBackup(file);
        toast(`Importé : ${res.expenses} justificatif${res.expenses > 1 ? 's' : ''}, ${res.clients} client${res.clients > 1 ? 's' : ''}`, { type: 'success' });
      } catch (err) {
        toast(err.message, { type: 'error' });
      }
    }
  };
}

async function showQr() {
  const s = state.settings;
  if (!s.anthropicKey && !s.ghToken && !s.msClientId) {
    toast('Configurez d’abord la clé Claude ou la synchronisation sur ce PC');
    return;
  }
  const ok = await confirmDialog({
    title: 'Afficher le QR code ?',
    message: 'Il contient vos clés (Claude, GitHub). Ne l’affichez pas devant d’autres personnes.',
    confirm: 'Afficher',
  });
  if (!ok) return;
  const sheet = openSheet({ title: 'Configurer mon téléphone', body: `<div class="loading">${spinner()}</div>` });
  try {
    const svg = await qrSvg(setupLink());
    sheet.body.innerHTML = `
      <div class="qr-box">
        <div class="qr">${svg}</div>
        <ol class="qr-steps">
          <li>Ouvrez l'appareil photo de l'iPhone et visez ce code.</li>
          <li>Touchez le lien qui apparaît : le site s'ouvre et se configure.</li>
          <li>Ajoutez-le ensuite à l'écran d'accueil (Partager → Sur l'écran d'accueil).</li>
        </ol>
        ${s.ghToken ? '' : '<p class="group-foot">Sans jeton GitHub, vos clients ne seront pas copiés sur le téléphone. Activez la synchronisation pour les retrouver partout.</p>'}
      </div>`;
  } catch (err) {
    sheet.body.innerHTML = `<div class="empty">${icon('alert')}<h3>QR code indisponible</h3><p>${esc(err.message)}</p></div>`;
  }
}
