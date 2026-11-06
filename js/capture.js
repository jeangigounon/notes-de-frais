// Ajout de justificatifs : appareil photo, sélection de fichiers, glisser-déposer, coller.

import { state, saveExpense, getClient, activeClients } from './store.js';
import { emptyData } from './fields.js';
import { prepareFile } from './image.js';
import { analyzeExpense } from './ai.js';
import { uid, toast, $ } from './ui.js';
import { openExpenseSheet } from './views/expenseSheet.js';

// Client choisi d'avance quand on ajoute un ticket depuis la note en cours d'un client.
let presetClientId = null;

export function openCamera({ clientId = null } = {}) { presetClientId = clientId; $('#capture-input').click(); }
export function openFilePicker({ clientId = null } = {}) { presetClientId = clientId; $('#file-input').click(); }

function defaultClientId() {
  if (getClient(state.settings.lastClientId)) return state.settings.lastClientId;
  const clients = activeClients();
  return clients.length === 1 ? clients[0].id : '';
}

export async function handleFiles(fileList, { clientId } = {}) {
  const files = [...(fileList || [])];
  if (!files.length) return;
  navigator.storage?.persist?.().catch(() => {});
  const created = [];
  for (const file of files) {
    try {
      const prep = await prepareFile(file);
      const exp = {
        id: uid(),
        createdAt: Date.now(),
        status: 'pending',
        clientId: clientId ?? defaultClientId(),
        file: prep.blob,
        mime: prep.mime,
        fileName: prep.name,
        thumb: prep.thumb,
        pages: prep.pages || 1,
        data: emptyData(),
        aiStatus: 'none',
      };
      await saveExpense(exp);
      created.push(exp);
    } catch (err) {
      toast(err.message || "Impossible d'ajouter ce fichier", { type: 'error' });
    }
  }
  if (!created.length) return;
  for (const exp of created) analyzeExpense(exp.id);
  if (created.length === 1) {
    openExpenseSheet(created[0].id, { isNew: true });
  } else {
    toast(`${created.length} justificatifs ajoutés · analyse en cours`, { type: 'success' });
    if (!location.hash.startsWith('#/depenses')) location.hash = '#/depenses';
  }
}

export function initCapture() {
  for (const id of ['capture-input', 'file-input']) {
    const input = $(`#${id}`);
    input.addEventListener('change', () => {
      const files = [...input.files];
      const clientId = presetClientId;
      presetClientId = null;
      input.value = '';
      handleFiles(files, clientId ? { clientId } : {});
    });
  }

  // Glisser-déposer n'importe où (sur PC).
  const overlay = $('#drop-overlay');
  let depth = 0;
  const hasFiles = e => [...(e.dataTransfer?.types || [])].includes('Files');
  // Une zone de dépôt dédiée (modèle Excel d'un client) a priorité sur l'ajout de justificatifs.
  const localZone = () => document.querySelector('[data-dropzone]');
  window.addEventListener('dragenter', e => {
    if (!hasFiles(e) || localZone() || document.body.classList.contains('has-sheet')) return;
    depth++;
    overlay.classList.add('on');
  });
  window.addEventListener('dragleave', () => { depth = Math.max(0, depth - 1); if (!depth) overlay.classList.remove('on'); });
  window.addEventListener('dragover', e => { if (hasFiles(e)) e.preventDefault(); });
  window.addEventListener('drop', e => {
    if (!hasFiles(e)) return;
    e.preventDefault();
    depth = 0;
    overlay.classList.remove('on');
    if (localZone() || document.body.classList.contains('has-sheet')) return;
    handleFiles(e.dataTransfer.files);
  });

  // Coller une capture d'écran ou un fichier (Ctrl+V).
  document.addEventListener('paste', e => {
    if (e.target.closest?.('input, textarea, [contenteditable]')) return;
    const files = [...(e.clipboardData?.files || [])].filter(f => f.type.startsWith('image/') || f.type === 'application/pdf');
    if (files.length) { e.preventDefault(); handleFiles(files); }
  });
}
