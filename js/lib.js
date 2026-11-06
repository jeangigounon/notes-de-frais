// Bibliothèques externes, chargées uniquement quand on en a besoin (versions figées).

export const CDN = {
  jszip: 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js',
  qrcode: 'https://cdnjs.cloudflare.com/ajax/libs/qrcode-generator/1.4.4/qrcode.min.js',
  tesseract: 'https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js',
  anthropic: 'https://cdn.jsdelivr.net/npm/@anthropic-ai/sdk@0.137.0/+esm',
  pdfjs: 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js',
  pdfjsWorker: 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js',
  sheetjs: 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js',
};

const loaded = new Map();

export function loadScript(src) {
  if (!loaded.has(src)) {
    loaded.set(src, new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = src;
      s.async = true;
      s.crossOrigin = 'anonymous';
      s.onload = resolve;
      s.onerror = () => {
        loaded.delete(src);
        s.remove();
        reject(new Error(`Impossible de charger ${new URL(src).hostname}. Vérifiez votre connexion.`));
      };
      document.head.append(s);
    }));
  }
  return loaded.get(src);
}

export async function getJSZip() {
  if (!window.JSZip) await loadScript(CDN.jszip);
  return window.JSZip;
}
