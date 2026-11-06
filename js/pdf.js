// Factures PDF : vignette de la 1re page, texte intégré, rendu en image (pour la lecture locale).

import { CDN, loadScript } from './lib.js';

async function pdfjs() {
  if (!window.pdfjsLib) await loadScript(CDN.pdfjs);
  window.pdfjsLib.GlobalWorkerOptions.workerSrc = CDN.pdfjsWorker;
  return window.pdfjsLib;
}

async function open(blob) {
  const lib = await pdfjs();
  return lib.getDocument({ data: new Uint8Array(await blob.arrayBuffer()), isEvalSupported: false }).promise;
}

async function renderPage(doc, maxSide) {
  const page = await doc.getPage(1);
  const base = page.getViewport({ scale: 1 });
  const viewport = page.getViewport({ scale: maxSide / Math.max(base.width, base.height) });
  const canvas = document.createElement('canvas');
  canvas.width = Math.round(viewport.width);
  canvas.height = Math.round(viewport.height);
  const ctx = canvas.getContext('2d');
  ctx.fillStyle = '#fff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  await page.render({ canvasContext: ctx, viewport }).promise;
  return canvas;
}

// Vignette (data URL JPEG) et nombre de pages.
export async function pdfThumb(blob, maxSide = 320) {
  const doc = await open(blob);
  try {
    return { thumb: (await renderPage(doc, maxSide)).toDataURL('image/jpeg', 0.75), pages: doc.numPages };
  } finally {
    doc.destroy();
  }
}

// Texte intégré au PDF (factures électroniques), ligne par ligne.
export async function pdfText(blob, maxPages = 5) {
  const doc = await open(blob);
  try {
    let text = '';
    for (let p = 1; p <= Math.min(doc.numPages, maxPages); p++) {
      const content = await (await doc.getPage(p)).getTextContent();
      text += content.items.map(i => i.str + (i.hasEOL ? '\n' : ' ')).join('') + '\n';
    }
    return text;
  } finally {
    doc.destroy();
  }
}

// Première page en image (PDF scanné sans texte), pour l'OCR.
export async function pdfFirstPageBlob(blob, maxSide = 2000) {
  const doc = await open(blob);
  try {
    const canvas = await renderPage(doc, maxSide);
    return await new Promise(resolve => canvas.toBlob(resolve, 'image/jpeg', 0.9));
  } finally {
    doc.destroy();
  }
}
