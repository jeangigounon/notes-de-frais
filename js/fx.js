// Taux de change de référence de la BCE (via frankfurter.dev, gratuit, sans clé).

import { todayISO } from './ui.js';

const cache = new Map();

// Valeur en euros d'une unité de la devise, au jour indiqué (ou au dernier jour ouvré précédent).
export async function rateToEur(currency, date) {
  const cur = String(currency || 'EUR').toUpperCase();
  if (cur === 'EUR') return 1;
  const day = date && date <= todayISO() ? date : 'latest';
  const key = `${cur}:${day}`;
  if (!cache.has(key)) {
    cache.set(key, fetch(`https://api.frankfurter.dev/v1/${day}?base=${encodeURIComponent(cur)}&symbols=EUR`)
      .then(res => {
        if (!res.ok) throw new Error(`Taux ${cur} indisponible`);
        return res.json();
      })
      .then(json => {
        const rate = json?.rates?.EUR;
        if (!rate) throw new Error(`Taux ${cur} indisponible`);
        return rate;
      })
      .catch(err => { cache.delete(key); throw err; }));
  }
  return cache.get(key);
}
