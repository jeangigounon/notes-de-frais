// Préparation et envoi de la note de frais :
// Microsoft 365 (envoi direct via Graph), menu Partager du téléphone, brouillon Outlook (.eml) ou archive .zip.

import { state, saveExpense } from './store.js';
import { buildReport, fillPlaceholders } from './fields.js';
import { buildSimple, fillTemplate } from './xlsx.js';
import { getJSZip } from './lib.js';
import { blobToBase64, safeFileName, todayISO } from './ui.js';
import { rateToEur } from './fx.js';

export const splitEmails = s => String(s || '').split(/[,;\s]+/).map(x => x.trim()).filter(Boolean);
export const isEmail = s => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(s);

/* ---------- Composition ---------- */

export function composeMessage(client, report) {
  return {
    to: splitEmails(client.to),
    cc: splitEmails(client.cc),
    subject: fillPlaceholders(client.subject || 'Note de frais – {nom} – {periode}', report),
    body: fillPlaceholders(client.body || '', report),
  };
}

export function makeReport(client, expenses) {
  return buildReport(client, expenses, state.profile);
}

// Taux de change BCE du jour de chaque dépense en devise (une seule fois, puis enregistré).
export async function ensureRates(expenses) {
  for (const e of expenses) {
    const cur = e.data.currency || 'EUR';
    if (cur === 'EUR' || e.data.rate) continue;
    try {
      e.data.rate = await rateToEur(cur, e.data.date);
      await saveExpense(e);
    } catch { /* hors ligne : signalé dans les avertissements */ }
  }
}

export async function buildAttachments(client, report) {
  const excel = client.excel || { mode: 'simple' };
  const baseName = safeFileName(fillPlaceholders(excel.fileName || 'Note de frais {nom} {periode}', report));
  let sheets = { files: [], warnings: [] };
  if (excel.mode === 'template' && excel.template?.data) {
    // Cellules « taux de change » d'une devise absente de la note : taux du jour.
    for (const m of Object.values(excel.template.cells || {})) {
      const cur = m.field === 'fxRate' && String(m.currency || '').toUpperCase();
      if (cur && cur !== 'EUR' && report.fx[cur] == null) {
        try { report.fx[cur] = await rateToEur(cur, report.to || todayISO()); } catch { /* laissé vide */ }
      }
    }
    sheets = await fillTemplate(excel.template, report, baseName);
  } else if (excel.mode !== 'none') {
    sheets = await buildSimple(excel.simple, report, baseName);
  }
  const warnings = [...(sheets.warnings || [])];
  const receipts = [];
  if (client.attachReceipts !== false) {
    for (const it of report.items) {
      if (it.exp.file) receipts.push({ name: it.fileName, blob: it.exp.file, mime: it.exp.mime });
      else warnings.push(`La photo de « ${it.merchant || 'Justificatif ' + it.index} » n'est pas encore arrivée sur cet appareil : elle ne sera pas jointe.`);
    }
  }
  return { files: [...sheets.files, ...receipts], excelFiles: sheets.files, receipts, warnings, added: sheets.added || 0 };
}

export function preferredMethod() {
  const m = state.settings.sendMethod;
  if (m && m !== 'auto') return m;
  if (state.settings.msClientId) return 'graph';
  if (matchMedia('(hover: none) and (pointer: coarse)').matches && navigator.canShare) return 'share';
  return 'eml';
}

/* ---------- Microsoft 365 (OAuth PKCE + Graph) ---------- */

const MS_SCOPES = 'openid profile offline_access User.Read Mail.Send';
const MS_KEY = 'ndf.ms';
const PKCE_KEY = 'ndf.pkce';

export const redirectUri = () => location.origin + location.pathname;
const msTenant = () => state.settings.msTenant || state.profile.email.split('@')[1] || 'organizations';
const msAuthority = () => `https://login.microsoftonline.com/${encodeURIComponent(msTenant())}/oauth2/v2.0`;

function readTokens() {
  try { return JSON.parse(localStorage.getItem(MS_KEY) || 'null'); } catch { return null; }
}
function saveTokens(t) {
  const prev = readTokens() || {};
  let account = prev.account || null;
  if (t.id_token) {
    try {
      const payload = JSON.parse(decodeURIComponent(escape(atob(t.id_token.split('.')[1].replace(/-/g, '+').replace(/_/g, '/')))));
      account = { name: payload.name || '', email: payload.preferred_username || payload.email || '' };
    } catch { /* jeton illisible : on garde l'ancien compte */ }
  }
  localStorage.setItem(MS_KEY, JSON.stringify({
    access_token: t.access_token,
    refresh_token: t.refresh_token || prev.refresh_token,
    expiresAt: Date.now() + (t.expires_in || 3600) * 1000,
    account,
  }));
}
export const msAccount = () => readTokens()?.account || null;
export function msSignOut() { localStorage.removeItem(MS_KEY); }

const b64url = bytes => btoa(String.fromCharCode(...new Uint8Array(bytes))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
const randomString = n => b64url(crypto.getRandomValues(new Uint8Array(n)));

export async function msSignIn(returnState = null) {
  if (!state.settings.msClientId) throw new Error("Renseignez d'abord l'identifiant de l'application Microsoft dans Réglages.");
  const verifier = randomString(48);
  const challenge = b64url(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier)));
  const st = randomString(16);
  sessionStorage.setItem(PKCE_KEY, JSON.stringify({ verifier, state: st, returnState, tenant: msTenant() }));
  const params = new URLSearchParams({
    client_id: state.settings.msClientId,
    response_type: 'code',
    redirect_uri: redirectUri(),
    response_mode: 'fragment',
    scope: MS_SCOPES,
    code_challenge: challenge,
    code_challenge_method: 'S256',
    state: st,
  });
  if (state.profile.email) params.set('login_hint', state.profile.email);
  location.assign(`${msAuthority()}/authorize?${params}`);
}

async function tokenRequest(body) {
  const res = await fetch(`${msAuthority()}/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: state.settings.msClientId, scope: MS_SCOPES, ...body }),
  });
  const json = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error((json.error_description || json.error || `Erreur ${res.status}`).split(/\r?\n/)[0]);
  return json;
}

// Retour de la page de connexion Microsoft (#code=…&state=…).
export async function msHandleRedirect() {
  const hash = location.hash.slice(1);
  if (!/(^|&)(code|error)=/.test(hash)) return { handled: false };
  const p = new URLSearchParams(hash);
  history.replaceState(null, '', location.pathname + location.search + '#/');
  const saved = JSON.parse(sessionStorage.getItem(PKCE_KEY) || 'null');
  sessionStorage.removeItem(PKCE_KEY);
  if (p.get('error')) throw new Error(`Connexion Microsoft refusée : ${(p.get('error_description') || p.get('error')).split(/\r?\n/)[0]}`);
  if (!saved || saved.state !== p.get('state')) throw new Error('Réponse de connexion inattendue, réessayez.');
  const tokens = await tokenRequest({ grant_type: 'authorization_code', code: p.get('code'), redirect_uri: redirectUri(), code_verifier: saved.verifier });
  saveTokens(tokens);
  return { handled: true, returnState: saved.returnState };
}

async function msAccessToken() {
  const t = readTokens();
  if (!t) return null;
  if (t.expiresAt - 60000 > Date.now()) return t.access_token;
  if (t.refresh_token) {
    try {
      const fresh = await tokenRequest({ grant_type: 'refresh_token', refresh_token: t.refresh_token });
      saveTokens(fresh);
      return fresh.access_token;
    } catch { /* session expirée : nouvelle connexion nécessaire */ }
  }
  return null;
}

export async function msIsSignedIn() {
  return !!(await msAccessToken());
}

export async function sendViaGraph(msg, files) {
  const token = await msAccessToken();
  if (!token) return { needsLogin: true };
  const graph = async (path, body) => {
    const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: body ? JSON.stringify(body) : undefined,
    });
    if (res.status === 401) { msSignOut(); throw Object.assign(new Error('Session Microsoft expirée.'), { needsLogin: true }); }
    if (!res.ok) {
      const json = await res.json().catch(() => ({}));
      throw new Error(`Microsoft 365 : ${json.error?.message || `erreur ${res.status}`}`);
    }
    return res.status === 202 || res.status === 204 ? null : res.json().catch(() => null);
  };
  const recipients = list => list.map(address => ({ emailAddress: { address } }));
  const message = {
    subject: msg.subject,
    body: { contentType: 'Text', content: msg.body },
    toRecipients: recipients(msg.to),
    ccRecipients: recipients(msg.cc),
  };
  const attachments = await Promise.all(files.map(async f => ({
    '@odata.type': '#microsoft.graph.fileAttachment',
    name: f.name,
    contentType: f.mime,
    contentBytes: await blobToBase64(f.blob),
  })));
  const size = attachments.reduce((s, a) => s + a.contentBytes.length, 0);
  try {
    if (size < 3_000_000) {
      await graph('/me/sendMail', { message: { ...message, attachments }, saveToSentItems: true });
    } else {
      // Gros envoi : brouillon puis pièces jointes une par une (limite de 4 Mo par requête).
      const draft = await graph('/me/messages', message);
      for (const a of attachments) {
        if (a.contentBytes.length > 3_000_000) throw new Error(`La pièce jointe « ${a.name} » est trop lourde pour un envoi direct.`);
        await graph(`/me/messages/${draft.id}/attachments`, a);
      }
      await graph(`/me/messages/${draft.id}/send`);
    }
  } catch (err) {
    if (err.needsLogin) return { needsLogin: true };
    throw err;
  }
  return { ok: true };
}

/* ---------- Menu Partager (iPhone / Android) ---------- */

export function canShareFiles(files) {
  try {
    return !!navigator.canShare && navigator.canShare({ files: files.map(f => new File([f.blob], f.name, { type: f.mime })) });
  } catch {
    return false;
  }
}

// Doit être appelé directement depuis le geste de l'utilisateur (sans attente préalable).
export function shareFiles(msg, files) {
  if (msg.to.length) navigator.clipboard?.writeText(msg.to.join(', ')).catch(() => {});
  return navigator.share({
    files: files.map(f => new File([f.blob], f.name, { type: f.mime })),
    title: msg.subject,
    text: msg.body,
  });
}

/* ---------- Brouillon Outlook (.eml) ---------- */

const utf8b64 = s => btoa(unescape(encodeURIComponent(s)));
const encWord = s => (/^[\x20-\x7e]*$/.test(s) ? s : `=?UTF-8?B?${utf8b64(s)}?=`);
const wrap76 = s => s.replace(/.{1,76}/g, '$&\r\n').trimEnd();

export async function buildEml(msg, files) {
  const boundary = `----=_NDF_${Math.random().toString(36).slice(2)}`;
  const lines = [
    'X-Unsent: 1',
    `To: ${msg.to.join(', ')}`,
    msg.cc.length ? `Cc: ${msg.cc.join(', ')}` : null,
    `Subject: ${encWord(msg.subject)}`,
    'MIME-Version: 1.0',
    `Content-Type: multipart/mixed; boundary="${boundary}"`,
    '',
    `--${boundary}`,
    'Content-Type: text/plain; charset="utf-8"',
    'Content-Transfer-Encoding: base64',
    '',
    wrap76(utf8b64(msg.body)),
  ].filter(l => l !== null);
  for (const f of files) {
    const name = encWord(f.name);
    lines.push(
      `--${boundary}`,
      `Content-Type: ${f.mime}; name="${name}"`,
      `Content-Disposition: attachment; filename="${name}"`,
      'Content-Transfer-Encoding: base64',
      '',
      wrap76(await blobToBase64(f.blob)),
    );
  }
  lines.push(`--${boundary}--`, '');
  return new Blob([lines.join('\r\n')], { type: 'message/rfc822' });
}

/* ---------- Archive ---------- */

export async function buildZip(files) {
  const JSZip = await getJSZip();
  const zip = new JSZip();
  for (const f of files) zip.file(f.name, f.blob);
  return zip.generateAsync({ type: 'blob', mimeType: 'application/zip' });
}
