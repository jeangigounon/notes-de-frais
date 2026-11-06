// Préparation des photos et PDF : redimensionnement, compression JPEG et vignette.

import { pdfThumb } from './pdf.js';

const MAX_SIDE = 2000;
const THUMB_SIDE = 320;

function guessMime(name = '') {
  const ext = name.split('.').pop().toLowerCase();
  return { pdf: 'application/pdf', jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png', webp: 'image/webp', gif: 'image/gif', heic: 'image/heic', heif: 'image/heif' }[ext] || '';
}

async function decode(file) {
  if ('createImageBitmap' in window) {
    try { return await createImageBitmap(file, { imageOrientation: 'from-image' }); } catch { /* on tente <img> */ }
  }
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => { URL.revokeObjectURL(url); resolve(img); };
    img.onerror = () => { URL.revokeObjectURL(url); reject(new Error('decode')); };
    img.src = url;
  });
}

function draw(img, maxSide) {
  const w = img.naturalWidth || img.width;
  const h = img.naturalHeight || img.height;
  const scale = Math.min(1, maxSide / Math.max(w, h));
  const canvas = document.createElement('canvas');
  canvas.width = Math.round(w * scale);
  canvas.height = Math.round(h * scale);
  const ctx = canvas.getContext('2d');
  ctx.fillStyle = '#fff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
  return canvas;
}

const toBlob = (canvas, type, quality) => new Promise((resolve, reject) =>
  canvas.toBlob(b => (b ? resolve(b) : reject(new Error('Compression impossible'))), type, quality));

const renameTo = (name, ext) => (name || 'photo').replace(/\.[^.]+$/, '') + '.' + ext;

export async function prepareFile(file) {
  const mime = file.type || guessMime(file.name);
  if (mime === 'application/pdf') {
    if (file.size > 20_000_000) throw new Error(`« ${file.name} » est trop lourd (20 Mo maximum).`);
    let thumb = null;
    let pages = 1;
    try {
      ({ thumb, pages } = await pdfThumb(file));
    } catch (err) {
      console.warn('Vignette PDF', err);
    }
    return { blob: file, mime, name: file.name || 'facture.pdf', thumb, pages };
  }
  if (!mime.startsWith('image/')) {
    throw new Error(`« ${file.name} » n'est ni une photo ni un PDF.`);
  }
  let img;
  try {
    img = await decode(file);
  } catch {
    // Ex. HEIC sur un PC Windows : on garde l'original tel quel.
    return { blob: file, mime, name: file.name, thumb: null, undecodable: true };
  }
  const blob = await toBlob(draw(img, MAX_SIDE), 'image/jpeg', 0.85);
  const thumb = draw(img, THUMB_SIDE).toDataURL('image/jpeg', 0.72);
  img.close?.();
  return { blob, mime: 'image/jpeg', name: renameTo(file.name, 'jpg'), thumb };
}
