/**
 * Proxy di riserva (facoltativo) per l'API Open Data di Normattiva — Cloudflare Worker.
 * Serve SOLO se il "Test connessione" del plugin fallisce per CORS.
 * Inoltra esclusivamente le richieste di dettaglio atto e aggiunge le intestazioni CORS
 * per il tuo indirizzo GitHub Pages. Non registra né conserva dati.
 *
 * Configura ORIGINE con il tuo dominio GitHub Pages (senza percorso), es. https://mariorossi.github.io
 */
const ORIGINE = 'https://TUO-UTENTE.github.io';
const DESTINAZIONE = 'https://api.normattiva.it/t/normattiva.api/bff-opendata/v1';
const PERCORSI_AMMESSI = ['/api/v1/atto/dettaglio-atto-urn'];

function cors(origin) {
  return {
    'Access-Control-Allow-Origin': origin === ORIGINE ? ORIGINE : 'null',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Accept',
    'Access-Control-Max-Age': '86400',
    'Vary': 'Origin'
  };
}

export default {
  async fetch(request) {
    const url = new URL(request.url);
    const origin = request.headers.get('Origin') || '';
    if (request.method === 'OPTIONS') return new Response(null, { status: 204, headers: cors(origin) });
    if (request.method !== 'POST' || !PERCORSI_AMMESSI.includes(url.pathname)) {
      return new Response('Non ammesso', { status: 403, headers: cors(origin) });
    }
    const corpo = await request.text();
    if (corpo.length > 2000) return new Response('Richiesta troppo grande', { status: 413, headers: cors(origin) });
    const risposta = await fetch(DESTINAZIONE + url.pathname, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' },
      body: corpo
    });
    const testo = await risposta.text();
    return new Response(testo, {
      status: risposta.status,
      headers: { ...cors(origin), 'Content-Type': 'application/json; charset=utf-8' }
    });
  }
};
