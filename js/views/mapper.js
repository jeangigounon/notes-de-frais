// Configuration d'un modèle Excel : on clique sur une cellule et on choisit ce qu'elle doit contenir.

import { esc, icon, openSheet, toast, promptDialog, confirmDialog, spinner } from '../ui.js';
import { LINE_FIELDS, SINGLE_FIELDS, CATEGORIES, CURRENCIES, fieldDef, lineColDef, lineColLabel } from '../fields.js';
import { openWorkbook, readSheet, templateCapacity, splitRef, numToCol } from '../xlsx.js';

const MIN_ROWS = 30;
const MIN_COLS = 8;

function fieldLabel(m) {
  if (m.field === 'fixed') return `« ${m.value} »`;
  if (m.field === 'fxRate') return `Taux ${m.currency} → €`;
  return fieldDef(m.field)?.label || m.field;
}

// Conditions d'une colonne : seulement certaines catégories et/ou une devise.
function editConditions(col, def) {
  return new Promise(resolve => {
    const d = lineColDef(def);
    const picked = new Set(d.categories || []);
    let others = !!d.others;
    let currency = d.currencies?.[0] || '';
    let result = null;
    const sheet = openSheet({
      title: `Colonne ${col} · ${fieldDef(d.field)?.label || d.field}`,
      left: '<button type="button" class="btn-text" data-sheet-close>Annuler</button>',
      right: '<button type="button" class="btn-text strong" data-ok>OK</button>',
      onClose: () => resolve(result),
    });
    const draw = () => {
      sheet.body.innerHTML = `
        <p class="group-foot" style="margin-top:4px">Utile quand le modèle a une colonne par type de dépense ou par devise : la valeur n'est écrite que pour les justificatifs concernés.</p>
        <h3 class="group-label">Catégories</h3>
        <div class="chips">
          <button type="button" class="chip ${!picked.size && !others ? 'on' : ''}" data-all>Toutes</button>
          ${CATEGORIES.map(c => `<button type="button" class="chip ${picked.has(c) ? 'on' : ''}" data-cat="${esc(c)}">${esc(c)}</button>`).join('')}
        </div>
        <div class="list form" style="margin-top:12px">
          <label class="field"><span>Et toutes les catégories sans colonne propre (« divers »)</span><span class="switch"><input type="checkbox" data-others ${others ? 'checked' : ''}><span></span></span></label>
        </div>
        <h3 class="group-label">Devise</h3>
        <div class="list form">
          <label class="field"><span>Seulement</span><select data-cur>
            <option value="">Toutes les devises</option>
            ${CURRENCIES.map(x => `<option ${x === currency ? 'selected' : ''}>${x}</option>`).join('')}
          </select></label>
        </div>
        <p class="group-foot">Une colonne en EUR reçoit aussi, convertis en euros, les montants des devises qui n'ont pas de colonne propre dans le modèle.</p>`;
    };
    draw();
    sheet.body.onclick = e => {
      const cat = e.target.closest('[data-cat]')?.dataset.cat;
      if (cat) { if (picked.has(cat)) picked.delete(cat); else picked.add(cat); draw(); return; }
      if (e.target.closest('[data-all]')) { picked.clear(); others = false; draw(); }
    };
    sheet.body.onchange = e => {
      if (e.target.matches('[data-others]')) { others = e.target.checked; draw(); }
      if (e.target.matches('[data-cur]')) currency = e.target.value;
    };
    sheet.el.querySelector('[data-ok]').onclick = () => {
      result = { field: d.field };
      if (picked.size) result.categories = [...picked];
      if (others) result.others = true;
      if (currency) result.currencies = [currency];
      sheet.close();
    };
  });
}

export async function openMapper(client, onDone) {
  const tpl = client.excel.template;
  const work = structuredClone({ cells: tpl.cells || {}, lineRow: tpl.lineRow || null, lineCols: tpl.lineCols || {}, capacity: tpl.capacity || null, sheet: tpl.sheet });
  let discard = false;

  const sheet = openSheet({
    title: `Modèle · ${tpl.name}`,
    wide: true,
    left: '<button type="button" class="btn-text" data-act="cancel">Annuler</button>',
    right: '<button type="button" class="btn-text strong" data-act="done">Terminé</button>',
    onClose: () => {
      closePopover();
      if (discard) return;
      Object.assign(client.excel.template, work);
      onDone?.();
    },
  });
  sheet.body.innerHTML = `<div class="loading">${spinner()} Lecture du modèle…</div>`;

  let book;
  let content;
  try {
    book = await openWorkbook(tpl.data);
    const current = book.sheets.find(s => s.name === work.sheet) || book.sheets[0];
    work.sheet = current.name;
    content = await readSheet(book, current.path);
  } catch (err) {
    sheet.body.innerHTML = `<div class="empty">${icon('alert')}<h3>Modèle illisible</h3><p>${esc(err.message)}</p></div>`;
    return;
  }

  let popover = null;
  function closePopover() {
    popover?.remove();
    popover = null;
    document.removeEventListener('pointerdown', outside, true);
    document.removeEventListener('keydown', onPopKey, true);
    sheet.body.querySelector('td.sel')?.classList.remove('sel');
  }
  function outside(e) {
    if (popover && !popover.contains(e.target) && !e.target.closest('td[data-ref]')) closePopover();
  }
  function onPopKey(e) {
    if (e.key === 'Escape') { e.stopPropagation(); closePopover(); }
  }

  function autoCapacity() {
    return templateCapacity({ ...work, capacity: null }, content);
  }

  function gridHtml() {
    const cap = work.capacity || autoCapacity();
    const rows = Math.min(150, Math.max(MIN_ROWS, content.maxRow + 6, (work.lineRow || 0) + Math.min(isFinite(cap) ? cap : 5, 40) + 3));
    const cols = Math.min(40, Math.max(MIN_COLS, content.maxCol + 2));
    const widths = [];
    for (let c = 1; c <= cols; c++) {
      const def = content.cols.find(x => c >= x.min && c <= x.max);
      widths[c] = def?.hidden ? 18 : def?.width ? Math.max(36, Math.min(320, Math.round(def.width * 7 + 6))) : 72;
    }
    const covered = new Set();
    const anchors = new Map();
    for (const m of content.merges) {
      anchors.set(`${m.r1}:${m.c1}`, m);
      for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) if (r !== m.r1 || c !== m.c1) covered.add(`${r}:${c}`);
    }
    const ghostRows = work.lineRow ? Math.min(isFinite(cap) ? cap : 4, 4) : 0;

    const tableWidth = 42 + widths.slice(1).reduce((s, w) => s + w, 0);
    let html = `<table class="xl" style="width:${tableWidth}px"><colgroup><col style="width:42px">${widths.slice(1).map(w => `<col style="width:${w}px">`).join('')}</colgroup>`;
    html += `<thead><tr><th class="corner"></th>${Array.from({ length: cols }, (_, i) => `<th>${numToCol(i + 1)}</th>`).join('')}</tr></thead><tbody>`;
    for (let r = 1; r <= rows; r++) {
      html += `<tr><th class="rh ${r === work.lineRow ? 'line' : ''}">${r}</th>`;
      for (let c = 1; c <= cols; c++) {
        if (covered.has(`${r}:${c}`)) continue;
        const col = numToCol(c);
        const ref = `${col}${r}`;
        const m = anchors.get(`${r}:${c}`);
        const span = m ? ` colspan="${m.c2 - m.c1 + 1}" rowspan="${m.r2 - m.r1 + 1}"` : '';
        const cell = content.cells.get(ref);
        const single = work.cells[ref];
        const lineField = r === work.lineRow ? work.lineCols[col] : null;
        const ghost = work.lineRow && work.lineCols[col] && r > work.lineRow && r < work.lineRow + ghostRows;
        let cls = '';
        let inner = '';
        if (single) { cls = 'm-single'; inner = `<span class="tag ${single.field === 'fixed' ? 'fixed' : ''}">${esc(fieldLabel(single))}</span>`; }
        else if (lineField) { cls = 'm-line'; inner = `<span class="tag line" title="${esc(lineColLabel(lineField))}">↓ ${esc(lineColLabel(lineField, { short: true }))}</span>`; }
        else if (ghost) { cls = 'm-ghost'; inner = '⋮'; }
        else if (cell) {
          if (cell.formula && !cell.display) { cls = 'formula'; inner = 'ƒx'; }
          else { cls = cell.formula ? 'formula' : ''; inner = esc(cell.display); }
        }
        html += `<td data-ref="${ref}" class="${cls}"${span} title="${esc(ref + (cell?.formula ? ' · =' + cell.formula : cell?.display ? ' · ' + cell.display : ''))}">${inner}</td>`;
      }
      html += '</tr>';
    }
    return html + '</tbody></table>';
  }

  function sideHtml() {
    const lines = Object.entries(work.lineCols).sort(([a], [b]) => a.localeCompare(b, 'en', { numeric: true }));
    const singles = Object.entries(work.cells).sort(([a], [b]) => {
      const pa = splitRef(a); const pb = splitRef(b);
      return pa.r - pb.r || pa.c - pb.c;
    });
    const auto = autoCapacity();
    return `
      <h3 class="group-label">Par justificatif${work.lineRow ? ` · ligne ${work.lineRow}` : ''}</h3>
      <div class="list">
        ${lines.length ? lines.map(([col, f]) => {
          const d = lineColDef(f);
          const cond = [d.categories?.length ? d.categories.join(', ') : '', d.others ? '+ autres catégories' : '', d.currencies?.length ? `devise ${d.currencies.join(', ')}` : ''].filter(Boolean).join(' · ');
          return `
          <div class="row compact"><span class="tag line">Col. ${col}</span><span class="row-main"><span class="row-title">${esc(fieldDef(d.field)?.label || d.field)}</span>${cond ? `<span class="row-sub">${esc(cond)}</span>` : ''}</span>
            <button type="button" class="btn-icon" data-cond="${col}" aria-label="Conditions" title="Seulement certaines catégories ou devises">${icon('sliders')}</button>
            <button type="button" class="btn-icon" data-unline="${col}" aria-label="Retirer">${icon('x')}</button></div>`;
        }).join('')
        : '<div class="row compact"><span class="row-sub">Cliquez sur la première ligne vide du tableau des dépenses, colonne par colonne.</span></div>'}
        ${lines.length ? `
          <label class="field"><span>Lignes disponibles</span><input type="number" min="1" max="500" data-capacity placeholder="${isFinite(auto) ? auto : 'illimité'}" value="${work.capacity || ''}"></label>` : ''}
      </div>
      ${lines.length ? `<p class="group-foot">${isFinite(auto) ? `${auto} ligne${auto > 1 ? 's' : ''} avant la ligne de total (détection automatique). S'il y a plus de justificatifs, des lignes sont ajoutées au tableau et les totaux suivent : tout reste dans un seul fichier.` : 'Aucune ligne de total détectée sous le tableau : toutes les dépenses tiennent sur une seule feuille.'}</p>` : ''}

      <h3 class="group-label">Une seule fois</h3>
      <div class="list">
        ${singles.length ? singles.map(([ref, m]) => `
          <div class="row compact"><span class="tag ${m.field === 'fixed' ? 'fixed' : ''}">${ref}</span><span class="row-main"><span class="row-title">${esc(fieldLabel(m))}</span></span><button type="button" class="btn-icon" data-unset="${ref}" aria-label="Retirer">${icon('x')}</button></div>`).join('')
        : '<div class="row compact"><span class="row-sub">Votre nom, la période, le total… où le client les attend.</span></div>'}
      </div>`;
  }

  function draw() {
    const scroll = sheet.body.querySelector('.grid-wrap');
    const pos = scroll ? [scroll.scrollLeft, scroll.scrollTop] : [0, 0];
    const ai = tpl.ai;
    sheet.body.innerHTML = `
      ${ai?.summary ? `
        <div class="ai-box">
          <div class="ai-box-head">${icon('sparkles')}<strong>${ai.engine === 'auto' ? 'Lecture automatique du modèle' : 'Analyse de Claude'}</strong><span>vérifiez puis touchez « Terminé »</span></div>
          <p>${esc(ai.summary)}</p>
          ${ai.warnings?.length ? `<ul>${ai.warnings.map(w => `<li>${esc(w)}</li>`).join('')}</ul>` : ''}
        </div>` : ai?.error ? `
        <div class="ai-box error">
          <div class="ai-box-head">${icon('alert')}<strong>L'analyse du modèle a échoué</strong></div>
          <p>${esc(ai.error)} Vous pouvez configurer les cellules à la main ci-dessous, ou relancer l'analyse depuis la fiche du client.</p>
        </div>` : ''}
      <div class="mapper">
        <div class="mapper-main">
          <div class="mapper-bar">
            ${book.sheets.length > 1 ? `<label class="inline-select">Feuille <select data-sheet>${book.sheets.map(s => `<option ${s.name === work.sheet ? 'selected' : ''}>${esc(s.name)}</option>`).join('')}</select></label>` : ''}
            <span class="legend"><i class="lg-single"></i>Une seule fois<i class="lg-line"></i>Par justificatif</span>
          </div>
          <p class="mapper-help">Cliquez sur une cellule pour choisir ce qu'elle contiendra. Pour le tableau des dépenses, cliquez sur sa <strong>première ligne vide</strong> et choisissez un champ « Par justificatif » pour chaque colonne : les lignes suivantes se remplissent toutes seules.</p>
          <div class="grid-wrap">${gridHtml()}</div>
        </div>
        <aside class="mapper-side">${sideHtml()}</aside>
      </div>`;
    const wrap = sheet.body.querySelector('.grid-wrap');
    wrap.scrollLeft = pos[0];
    wrap.scrollTop = pos[1];
  }

  function assignLine(ref, field) {
    const p = splitRef(ref);
    if (work.lineRow && work.lineRow !== p.r && Object.keys(work.lineCols).length) toast(`Tableau des dépenses déplacé à la ligne ${p.r}`);
    work.lineRow = p.r;
    const prev = work.lineCols[p.col];
    work.lineCols[p.col] = prev ? { ...lineColDef(prev), field } : { field };
    for (const r of Object.keys(work.cells)) {
      const q = splitRef(r);
      if (q.r === p.r && work.lineCols[q.col]) delete work.cells[r];
    }
  }

  function clearRef(ref) {
    const p = splitRef(ref);
    delete work.cells[ref];
    if (p.r === work.lineRow) delete work.lineCols[p.col];
    if (!Object.keys(work.lineCols).length) work.lineRow = null;
  }

  function assignSingle(ref, field, value) {
    clearRef(ref);
    work.cells[ref] = field === 'fixed' ? { field, value } : field === 'fxRate' ? { field, currency: value } : { field };
  }

  function openPopover(td) {
    closePopover();
    const ref = td.dataset.ref;
    const p = splitRef(ref);
    const cell = content.cells.get(ref);
    const curSingle = work.cells[ref];
    const curLine = p.r === work.lineRow && work.lineCols[p.col] ? lineColDef(work.lineCols[p.col]).field : null;
    td.classList.add('sel');
    const groups = [...new Set(SINGLE_FIELDS.map(f => f.group))];
    popover = document.createElement('div');
    popover.className = 'popover';
    popover.innerHTML = `
      <div class="pop-head"><strong>Cellule ${ref}</strong>${cell?.display ? ` · « ${esc(cell.display.slice(0, 40))} »` : cell?.formula ? ' · formule' : ''}</div>
      <div class="pop-group">Par justificatif · une ligne par ticket</div>
      ${LINE_FIELDS.map(f => `<button type="button" class="pop-item ${curLine === f.key ? 'cur' : ''}" data-line="${f.key}"><i class="dot line"></i>${esc(f.label)}${curLine === f.key ? icon('check') : ''}</button>`).join('')}
      ${groups.map(g => `
        <div class="pop-group">Une seule fois · ${esc(g)}</div>
        ${SINGLE_FIELDS.filter(f => f.group === g).map(f => `<button type="button" class="pop-item ${curSingle?.field === f.key ? 'cur' : ''}" data-single="${f.key}"><i class="dot"></i>${esc(f.label)}${curSingle?.field === f.key ? icon('check') : ''}</button>`).join('')}`).join('')}
      <div class="pop-group">Autre</div>
      <button type="button" class="pop-item ${curSingle?.field === 'fxRate' ? 'cur' : ''}" data-fx><i class="dot"></i>Taux de change d'une devise…</button>
      <button type="button" class="pop-item ${curSingle?.field === 'fixed' ? 'cur' : ''}" data-fixed><i class="dot fixed"></i>Texte fixe…</button>
      ${curSingle || curLine ? `<button type="button" class="pop-item danger" data-clear>${icon('x')}Vider cette cellule</button>` : ''}`;
    document.body.append(popover);
    // Sous la cellule, sinon au-dessus, sinon à côté : la cellule reste visible.
    const rect = td.getBoundingClientRect();
    const pw = popover.offsetWidth;
    const ph = popover.offsetHeight;
    const clampTop = y => Math.max(12, Math.min(y, innerHeight - ph - 12));
    let left = Math.min(rect.left, innerWidth - pw - 12);
    let top;
    if (rect.bottom + 6 + ph <= innerHeight - 12) top = rect.bottom + 6;
    else if (rect.top - 6 - ph >= 12) top = rect.top - ph - 6;
    else {
      top = clampTop(rect.top - ph / 2);
      left = rect.right + 8 + pw <= innerWidth - 12 ? rect.right + 8 : Math.max(12, rect.left - pw - 8);
    }
    popover.style.left = `${Math.max(12, left)}px`;
    popover.style.top = `${top}px`;
    setTimeout(() => document.addEventListener('pointerdown', outside, true), 0);
    document.addEventListener('keydown', onPopKey, true);

    popover.onclick = async e => {
      const b = e.target.closest('button');
      if (!b) return;
      if (b.dataset.line) assignLine(ref, b.dataset.line);
      else if (b.dataset.single) assignSingle(ref, b.dataset.single);
      else if (b.hasAttribute('data-fixed')) {
        closePopover();
        const value = await promptDialog({ title: `Texte fixe en ${ref}`, message: 'Ex. un code projet ou un centre de coût, toujours identique.', value: curSingle?.value || '' });
        if (value == null || !value.trim()) return;
        assignSingle(ref, 'fixed', value.trim());
      } else if (b.hasAttribute('data-fx')) {
        closePopover();
        const cur = await promptDialog({ title: `Taux de change en ${ref}`, message: 'Code de la devise (ex. CHF, GBP, USD). La cellule recevra la valeur en euros d’une unité de cette devise (taux BCE).', value: curSingle?.currency || '', placeholder: 'CHF' });
        const code = String(cur || '').trim().toUpperCase();
        if (!/^[A-Z]{3}$/.test(code)) return;
        assignSingle(ref, 'fxRate', code);
      } else if (b.hasAttribute('data-clear')) clearRef(ref);
      closePopover();
      draw();
    };
  }

  draw();

  sheet.body.onclick = e => {
    const td = e.target.closest('td[data-ref]');
    if (td) return openPopover(td);
    const cond = e.target.closest('[data-cond]');
    if (cond) {
      const col = cond.dataset.cond;
      editConditions(col, work.lineCols[col]).then(res => { if (res) { work.lineCols[col] = res; draw(); } });
      return;
    }
    const unline = e.target.closest('[data-unline]');
    if (unline) { delete work.lineCols[unline.dataset.unline]; if (!Object.keys(work.lineCols).length) work.lineRow = null; draw(); return; }
    const unset = e.target.closest('[data-unset]');
    if (unset) { delete work.cells[unset.dataset.unset]; draw(); }
  };
  sheet.body.onchange = async e => {
    if (e.target.matches('[data-capacity]')) {
      const n = parseInt(e.target.value, 10);
      work.capacity = n > 0 ? n : null;
      draw();
    }
    if (e.target.matches('[data-sheet]')) {
      const name = e.target.value;
      if ((Object.keys(work.cells).length || Object.keys(work.lineCols).length)
        && !(await confirmDialog({ title: 'Changer de feuille ?', message: 'Les cellules configurées sur la feuille actuelle seront effacées.', confirm: 'Changer' }))) {
        e.target.value = work.sheet;
        return;
      }
      Object.assign(work, { sheet: name, cells: {}, lineCols: {}, lineRow: null, capacity: null });
      content = await readSheet(book, book.sheets.find(s => s.name === name).path);
      draw();
    }
  };
  sheet.el.querySelector('.sheet-head').addEventListener('click', e => {
    const act = e.target.closest('[data-act]')?.dataset.act;
    if (act === 'cancel') { discard = true; sheet.close(); }
    if (act === 'done') sheet.close();
  });
}
