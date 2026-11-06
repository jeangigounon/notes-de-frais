// Briques d'interface partagées : helpers, icônes, feuilles, alertes, toasts.

export const $ = (sel, root = document) => root.querySelector(sel);
export const $$ = (sel, root = document) => [...root.querySelectorAll(sel)];

export function esc(value) {
  return String(value ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[ch]);
}

export function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}

export function debounce(fn, ms = 300) {
  let timer;
  const run = (...args) => { clearTimeout(timer); timer = setTimeout(() => fn(...args), ms); };
  run.flush = (...args) => { clearTimeout(timer); fn(...args); };
  run.cancel = () => clearTimeout(timer);
  return run;
}

export const isCompact = () => matchMedia('(max-width: 899px)').matches;
export const isTouch = () => matchMedia('(hover: none) and (pointer: coarse)').matches;

/* ---------- Dates & montants ---------- */

const pad = n => String(n).padStart(2, '0');
export function todayISO(d = new Date()) {
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}
export function parseISO(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso || '');
  return m ? new Date(+m[1], +m[2] - 1, +m[3], 12) : null;
}
export function fmtDate(iso, opts = { day: 'numeric', month: 'short', year: 'numeric' }) {
  const d = parseISO(iso);
  return d ? d.toLocaleDateString('fr-BE', opts) : '';
}
export function fmtDateShort(iso) {
  return fmtDate(iso, { day: 'numeric', month: 'short' });
}
export function daysAgo(ts) {
  const days = Math.floor((Date.now() - ts) / 86400000);
  if (days <= 0) return "aujourd'hui";
  if (days === 1) return 'hier';
  return `il y a ${days} jours`;
}
export function timeAgo(ts) {
  if (!ts) return 'jamais';
  const s = Math.round((Date.now() - ts) / 1000);
  if (s < 45) return "à l'instant";
  if (s < 3600) return `il y a ${Math.round(s / 60)} min`;
  if (s < 86400) return `il y a ${Math.round(s / 3600)} h`;
  return daysAgo(ts);
}

export function money(n, currency = 'EUR') {
  if (n == null || !isFinite(n)) return '—';
  try {
    return new Intl.NumberFormat('fr-BE', { style: 'currency', currency: currency || 'EUR' }).format(n);
  } catch {
    return `${numStr(n)} ${currency}`;
  }
}
export function numStr(n) {
  if (n == null || !isFinite(n)) return '';
  return new Intl.NumberFormat('fr-BE', { minimumFractionDigits: 2, maximumFractionDigits: 2, useGrouping: false }).format(n);
}

/* ---------- Fichiers ---------- */

export function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result).split(',')[1] || '');
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(blob);
  });
}
export function base64ToBlob(b64, type = 'application/octet-stream') {
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return new Blob([bytes], { type });
}
export function download(blob, name) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.append(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
}
export function safeFileName(s) {
  return String(s || '').replace(/[\\/:*?"<>|\u0000-\u001f]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 120) || 'fichier';
}
export function pickFile(accept, { multiple = false } = {}) {
  return new Promise(resolve => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = accept;
    input.multiple = multiple;
    input.onchange = () => resolve(multiple ? [...input.files] : input.files[0] || null);
    input.click();
  });
}

/* ---------- Icônes (traits fins, façon SF Symbols) ---------- */

const ICONS = {
  home: '<path d="M3.5 10.5 12 3.5l8.5 7"/><path d="M5.5 9v10.5a1 1 0 0 0 1 1H10v-6h4v6h3.5a1 1 0 0 0 1-1V9"/>',
  receipt: '<path d="M6 3h12v18l-2-1.3-2 1.3-2-1.3-2 1.3-2-1.3L6 21z"/><path d="M9 8h6M9 12h6M9 16h3.5"/>',
  camera: '<path d="M14.6 4.5H9.4L7.6 7H5a2 2 0 0 0-2 2v9a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V9a2 2 0 0 0-2-2h-2.6z"/><circle cx="12" cy="13" r="3.6"/>',
  users: '<circle cx="9" cy="8.5" r="3.5"/><path d="M2.5 20c.7-3.6 3.3-5.6 6.5-5.6s5.8 2 6.5 5.6"/><path d="M15.5 5.2a3.4 3.4 0 0 1 0 6.6"/><path d="M18 14.7c1.8.8 3 2.6 3.5 5.3"/>',
  sliders: '<path d="M4 7h9M17 7h3M4 17h3M11 17h9"/><circle cx="15" cy="7" r="2.2"/><circle cx="9" cy="17" r="2.2"/>',
  plus: '<path d="M12 5v14M5 12h14"/>',
  chevronRight: '<path d="m9.5 6 6 6-6 6"/>',
  chevronLeft: '<path d="m14.5 6-6 6 6 6"/>',
  chevronDown: '<path d="m6 9.5 6 6 6-6"/>',
  x: '<path d="M6.5 6.5l11 11M17.5 6.5l-11 11"/>',
  check: '<path d="m5 12.5 4.5 4.5L19 7.5"/>',
  trash: '<path d="M4 7h16M10 11v6M14 11v6M6 7l.9 12.1A2 2 0 0 0 8.9 21h6.2a2 2 0 0 0 2-1.9L18 7M9 7V4.5A1.5 1.5 0 0 1 10.5 3h3A1.5 1.5 0 0 1 15 4.5V7"/>',
  send: '<path d="M20.5 3.5 10 14"/><path d="M20.5 3.5 14 20.5l-4-6.5-6.5-4z"/>',
  sparkles: '<path d="M11 3.5l1.7 4.3 4.3 1.7-4.3 1.7L11 15.5l-1.7-4.3L5 9.5l4.3-1.7z"/><path d="M18 14l.8 2.2L21 17l-2.2.8L18 20l-.8-2.2L15 17l2.2-.8z"/>',
  upload: '<path d="M12 15V4M7.5 8.5 12 4l4.5 4.5"/><path d="M4 15v3a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-3"/>',
  download: '<path d="M12 4v11M7.5 10.5 12 15l4.5-4.5"/><path d="M4 15v3a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-3"/>',
  table: '<rect x="3.5" y="4" width="17" height="16" rx="2"/><path d="M3.5 9h17M3.5 14.5h17M9 9v11"/>',
  mail: '<rect x="3" y="5" width="18" height="14" rx="2.5"/><path d="m4 7 8 6 8-6"/>',
  image: '<rect x="3.5" y="4" width="17" height="16" rx="2.5"/><circle cx="9" cy="9.5" r="1.6"/><path d="m20.5 15.5-4.5-4.5-9 9"/>',
  clock: '<circle cx="12" cy="12" r="8.5"/><path d="M12 7.5V12l3 2"/>',
  alert: '<path d="M10.3 4.3 2.8 17.5A2 2 0 0 0 4.5 20.5h15a2 2 0 0 0 1.7-3L13.7 4.3a2 2 0 0 0-3.4 0z"/><path d="M12 9.5v4M12 17h.01"/>',
  search: '<circle cx="11" cy="11" r="6.5"/><path d="m20 20-4.3-4.3"/>',
  sync: '<path d="M19.5 10.5A7.5 7.5 0 0 0 6.2 7.2L4.5 9"/><path d="M4.5 4.5V9H9"/><path d="M4.5 13.5a7.5 7.5 0 0 0 13.3 3.3l1.7-1.8"/><path d="M19.5 19.5V15H15"/>',
  qr: '<rect x="3.5" y="3.5" width="6.5" height="6.5" rx="1.2"/><rect x="14" y="3.5" width="6.5" height="6.5" rx="1.2"/><rect x="3.5" y="14" width="6.5" height="6.5" rx="1.2"/><path d="M14 14h2.5v2.5H14zM18 18h2.5v2.5H18zM14 18.5v2M18.5 14h2"/>',
  more: '<circle cx="5.5" cy="12" r="1.2" fill="currentColor"/><circle cx="12" cy="12" r="1.2" fill="currentColor"/><circle cx="18.5" cy="12" r="1.2" fill="currentColor"/>',
  doc: '<path d="M14 3.5H7.5a2 2 0 0 0-2 2v13a2 2 0 0 0 2 2h9a2 2 0 0 0 2-2V8z"/><path d="M14 3.5V8h4.5"/>',
  user: '<circle cx="12" cy="8.5" r="4"/><path d="M4.5 20.5c.9-3.9 3.8-6 7.5-6s6.6 2.1 7.5 6"/>',
  key: '<circle cx="8" cy="15" r="4"/><path d="m11 12 8.5-8.5M16.5 6.5l2.5 2.5M14.5 8.5l2 2"/>',
  phone: '<rect x="7" y="2.5" width="10" height="19" rx="2.6"/><path d="M11 18.5h2"/>',
  pencil: '<path d="M4 20h4L19 9a2.8 2.8 0 0 0-4-4L4 16z"/><path d="m13.5 6.5 4 4"/>',
  arrowUp: '<path d="M12 19V5M6.5 10.5 12 5l5.5 5.5"/>',
  arrowDown: '<path d="M12 5v14M6.5 13.5 12 19l5.5-5.5"/>',
  copy: '<rect x="8.5" y="8.5" width="11.5" height="11.5" rx="2.2"/><path d="M15.5 8.5V6a2 2 0 0 0-2-2H6a2 2 0 0 0-2 2v7.5a2 2 0 0 0 2 2h2.5"/>',
  cloud: '<path d="M7 18.5a4.5 4.5 0 0 1-.4-9A6 6 0 0 1 18 11a3.8 3.8 0 0 1-.5 7.5z"/>',
  tray: '<path d="M3.5 13.5h5l1.5 3h4l1.5-3h5"/><path d="M6 5h12l2.5 8.5v5a1.5 1.5 0 0 1-1.5 1.5H5a1.5 1.5 0 0 1-1.5-1.5v-5z"/>',
  share: '<path d="M12 3.5v11M8 7.5l4-4 4 4"/><path d="M8 10.5H6.5a2 2 0 0 0-2 2V19a2 2 0 0 0 2 2h11a2 2 0 0 0 2-2v-6.5a2 2 0 0 0-2-2H16"/>',
  eye: '<path d="M2.5 12S6 5.5 12 5.5 21.5 12 21.5 12 18 18.5 12 18.5 2.5 12 2.5 12z"/><circle cx="12" cy="12" r="3"/>',
  archive: '<rect x="3" y="4" width="18" height="5" rx="1.5"/><path d="M5 9v9a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V9M10 13h4"/>',
  grid: '<rect x="3.5" y="3.5" width="17" height="17" rx="2"/><path d="M3.5 9.2h17M3.5 14.8h17M9.2 3.5v17M14.8 3.5v17"/>',
  undo: '<path d="M9 14 4.5 9.5 9 5"/><path d="M4.5 9.5H15a5 5 0 0 1 0 10h-3"/>',
  info: '<circle cx="12" cy="12" r="8.5"/><path d="M12 11v5.5M12 7.8h.01"/>',
};

export function icon(name, cls = '') {
  return `<svg class="i ${cls}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICONS[name] || ''}</svg>`;
}

export function spinner(cls = '') {
  return `<span class="spinner ${cls}" role="status" aria-label="Chargement"></span>`;
}

/* ---------- Toasts ---------- */

export function toast(message, { type = 'info', duration = 3000, action } = {}) {
  const host = document.getElementById('toasts');
  const el = document.createElement('div');
  el.className = `toast toast-${type}`;
  const ic = type === 'success' ? icon('check') : type === 'error' ? icon('alert') : '';
  el.innerHTML = `${ic}<span>${esc(message)}</span>${action ? `<button type="button">${esc(action.label)}</button>` : ''}`;
  const dismiss = () => { el.classList.remove('show'); setTimeout(() => el.remove(), 300); };
  if (action) el.querySelector('button').onclick = () => { action.onClick(); dismiss(); };
  host.append(el);
  requestAnimationFrame(() => requestAnimationFrame(() => el.classList.add('show')));
  setTimeout(dismiss, duration);
}

/* ---------- Feuilles (modales) ---------- */

export function openSheet({ title = '', body = '', left, right = '', wide = false, className = '', onClose } = {}) {
  const host = document.getElementById('sheets');
  const wrap = document.createElement('div');
  wrap.className = `sheet-wrap ${wide ? 'wide' : ''} ${className}`;
  wrap.innerHTML = `
    <div class="sheet-backdrop"></div>
    <section class="sheet" role="dialog" aria-modal="true" aria-label="${esc(title)}">
      <div class="sheet-drag">
        <div class="sheet-grabber"></div>
        <header class="sheet-head">
          <div class="sheet-left">${left ?? '<button type="button" class="btn-text" data-sheet-close>Fermer</button>'}</div>
          <h2 class="sheet-title">${esc(title)}</h2>
          <div class="sheet-right">${right}</div>
        </header>
      </div>
      <div class="sheet-body"></div>
    </section>`;
  const bodyEl = wrap.querySelector('.sheet-body');
  if (typeof body === 'string') bodyEl.innerHTML = body; else if (body) bodyEl.append(body);
  host.append(wrap);
  document.body.classList.add('has-sheet');
  requestAnimationFrame(() => requestAnimationFrame(() => wrap.classList.add('open')));

  let closed = false;
  const sheetEl = wrap.querySelector('.sheet');
  const close = result => {
    if (closed) return;
    closed = true;
    document.removeEventListener('keydown', onKey);
    sheetEl.style.transform = '';
    wrap.classList.remove('open');
    wrap.classList.add('closing');
    setTimeout(() => {
      wrap.remove();
      if (!host.children.length) document.body.classList.remove('has-sheet');
    }, 320);
    onClose?.(result);
  };
  const onKey = e => {
    if (e.key === 'Escape' && host.lastElementChild === wrap && !document.querySelector('.alert-wrap, .as-wrap, .popover')) close();
  };
  document.addEventListener('keydown', onKey);
  wrap.querySelector('.sheet-backdrop').addEventListener('click', () => close());
  wrap.addEventListener('click', e => { if (e.target.closest('[data-sheet-close]')) close(); });
  enableDragToClose(wrap, close);

  return {
    el: wrap,
    body: bodyEl,
    close,
    get closed() { return closed; },
    setTitle(t) { wrap.querySelector('.sheet-title').textContent = t; },
    setLeft(html) { wrap.querySelector('.sheet-left').innerHTML = html; },
    setRight(html) { wrap.querySelector('.sheet-right').innerHTML = html; },
  };
}

function enableDragToClose(wrap, close) {
  const handle = wrap.querySelector('.sheet-drag');
  const sheet = wrap.querySelector('.sheet');
  let startY = null;
  let dy = 0;
  handle.addEventListener('touchstart', e => {
    if (!isCompact()) return;
    startY = e.touches[0].clientY;
    dy = 0;
    sheet.style.transition = 'none';
  }, { passive: true });
  handle.addEventListener('touchmove', e => {
    if (startY == null) return;
    dy = Math.max(0, e.touches[0].clientY - startY);
    sheet.style.transform = `translateY(${dy}px)`;
  }, { passive: true });
  handle.addEventListener('touchend', () => {
    if (startY == null) return;
    sheet.style.transition = '';
    if (dy > 110) close(); else sheet.style.transform = '';
    startY = null;
  });
}

/* ---------- Alertes & feuilles d'actions ---------- */

function overlay(className, html) {
  const wrap = document.createElement('div');
  wrap.className = className;
  wrap.innerHTML = html;
  document.body.append(wrap);
  requestAnimationFrame(() => requestAnimationFrame(() => wrap.classList.add('open')));
  const remove = () => { wrap.classList.remove('open'); setTimeout(() => wrap.remove(), 250); };
  return { wrap, remove };
}

export function confirmDialog({ title, message = '', confirm = 'OK', cancel = 'Annuler', destructive = false }) {
  return new Promise(resolve => {
    const { wrap, remove } = overlay('alert-wrap', `
      <div class="alert" role="alertdialog" aria-label="${esc(title)}">
        <div class="alert-text"><h3>${esc(title)}</h3>${message ? `<p>${esc(message)}</p>` : ''}</div>
        <div class="alert-actions">
          ${cancel ? `<button type="button" data-v="0">${esc(cancel)}</button>` : ''}
          <button type="button" data-v="1" class="${destructive ? 'danger' : 'strong'}">${esc(confirm)}</button>
        </div>
      </div>`);
    const done = v => { document.removeEventListener('keydown', onKey, true); remove(); resolve(v); };
    const onKey = e => {
      if (e.key === 'Escape') { e.stopPropagation(); done(false); }
      if (e.key === 'Enter') { e.preventDefault(); done(true); }
    };
    document.addEventListener('keydown', onKey, true);
    wrap.addEventListener('click', e => {
      const b = e.target.closest('button[data-v]');
      if (b) done(b.dataset.v === '1');
    });
    wrap.querySelector('button.strong, button.danger')?.focus();
  });
}

export function promptDialog({ title, message = '', value = '', placeholder = '', confirm = 'OK' }) {
  return new Promise(resolve => {
    const { wrap, remove } = overlay('alert-wrap', `
      <div class="alert" role="alertdialog" aria-label="${esc(title)}">
        <div class="alert-text"><h3>${esc(title)}</h3>${message ? `<p>${esc(message)}</p>` : ''}
          <input type="text" value="${esc(value)}" placeholder="${esc(placeholder)}">
        </div>
        <div class="alert-actions">
          <button type="button" data-v="0">Annuler</button>
          <button type="button" data-v="1" class="strong">${esc(confirm)}</button>
        </div>
      </div>`);
    const input = wrap.querySelector('input');
    const done = ok => { document.removeEventListener('keydown', onKey, true); remove(); resolve(ok ? input.value : null); };
    const onKey = e => {
      if (e.key === 'Escape') { e.stopPropagation(); done(false); }
      if (e.key === 'Enter') { e.preventDefault(); done(true); }
    };
    document.addEventListener('keydown', onKey, true);
    wrap.addEventListener('click', e => {
      const b = e.target.closest('button[data-v]');
      if (b) done(b.dataset.v === '1');
    });
    setTimeout(() => { input.focus(); input.select(); }, 60);
  });
}

export function actionSheet({ title = '', actions = [], cancel = 'Annuler' }) {
  return new Promise(resolve => {
    const { wrap, remove } = overlay('as-wrap', `
      <div class="as" role="menu">
        <div class="as-group">
          ${title ? `<div class="as-title">${esc(title)}</div>` : ''}
          ${actions.map((a, i) => `<button type="button" data-i="${i}" class="${a.destructive ? 'danger' : ''}" ${a.disabled ? 'disabled' : ''}>${a.icon ? icon(a.icon) : ''}<span>${esc(a.label)}</span></button>`).join('')}
        </div>
        <div class="as-group"><button type="button" class="cancel" data-i="-1">${esc(cancel)}</button></div>
      </div>`);
    const done = v => { document.removeEventListener('keydown', onKey, true); remove(); resolve(v); };
    const onKey = e => { if (e.key === 'Escape') { e.stopPropagation(); done(null); } };
    document.addEventListener('keydown', onKey, true);
    wrap.addEventListener('click', e => {
      const b = e.target.closest('button[data-i]');
      if (b) { const i = +b.dataset.i; done(i >= 0 ? actions[i].value ?? i : null); return; }
      if (e.target === wrap) done(null);
    });
  });
}
