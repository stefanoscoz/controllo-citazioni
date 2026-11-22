/**
 * Proxy (facoltativo) per il plugin "Controllo citazioni" — Cloudflare Worker.
 *
 * 1. /api/v1/atto/dettaglio-atto-urn  → inoltra a Normattiva (serve solo se il browser di Word blocca le chiamate dirette)
 * 2. /pertinenza                      → verifica di pertinenza con l'API Claude (livello 3)
 *
 * Configurazione (Cloudflare → Worker → Settings → Variables and Secrets):
 *   ANTHROPIC_API_KEY  (Secret, obbligatoria per /pertinenza)
 *   MODELLO            (facoltativa, es. "claude-sonnet-5-5")
 * Il Worker non registra né conserva testi. Ammette solo richieste dalla tua pagina GitHub Pages.
 */
const ORIGINE = 'https://stefanoscoz.github.io';
const NORMATTIVA = 'https://api.normattiva.it/t/normattiva.api/bff-opendata/v1';
const MODELLO_PREDEFINITO = 'claude-sonnet-5-5';

function cors(origin) {
  return {
    'Access-Control-Allow-Origin': origin === ORIGINE ? ORIGINE : 'null',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Accept',
    'Access-Control-Max-Age': '86400',
    'Vary': 'Origin'
  };
}

function json(dati, stato, origin) {
  return new Response(JSON.stringify(dati), { status: stato || 200, headers: { ...cors(origin), 'Content-Type': 'application/json; charset=utf-8' } });
}

const ISTRUZIONI = `Sei un assistente per dottori commercialisti italiani. Ricevi una frase tratta da un parere e il testo di una norma (vigente alla data di riferimento) che la frase cita.
Valuta SOLO se il testo della norma fornito sostiene ciò che la frase gli attribuisce. Non usare conoscenze esterne sulla norma: basati esclusivamente sul testo fornito.
Rispondi con un unico oggetto JSON, senza altro testo:
{"giudizio": "coerente" | "dubbia" | "non_pertinente", "motivazione": "max 2 frasi in italiano", "passaggio": "breve passaggio della norma rilevante, max 25 parole, o stringa vuota"}
- "coerente": il testo della norma sostiene l'affermazione.
- "dubbia": il collegamento è parziale, generico o dipende da altre disposizioni non fornite.
- "non_pertinente": il testo della norma riguarda altro o contraddice l'affermazione.`;

async function pertinenza(request, env, origin) {
  if (!env.ANTHROPIC_API_KEY) return json({ errore: 'Chiave API non configurata nel Worker' }, 500, origin);
  let corpo;
  try { corpo = await request.json(); } catch (e) { return json({ errore: 'JSON non valido' }, 400, origin); }
  const frase = String(corpo.frase || '').slice(0, 2500);
  const testo = String(corpo.testoNorma || '').slice(0, 8000);
  if (!frase || !testo) return json({ errore: 'Frase o testo della norma mancanti' }, 400, origin);

  const messaggio = `Data di riferimento: ${String(corpo.dataRiferimento || '').slice(0, 10)}
Citazione: ${String(corpo.citazione || '').slice(0, 300)}

<frase_del_parere>
${frase}
</frase_del_parere>

<testo_della_norma>
${testo}
</testo_della_norma>`;

  const r = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'x-api-key': env.ANTHROPIC_API_KEY, 'anthropic-version': '2023-06-01' },
    body: JSON.stringify({
      model: env.MODELLO || MODELLO_PREDEFINITO,
      max_tokens: 400,
      temperature: 0,
      system: ISTRUZIONI,
      messages: [{ role: 'user', content: messaggio }]
    })
  });
  if (!r.ok) return json({ errore: 'API Claude: HTTP ' + r.status }, 502, origin);
  const dati = await r.json();
  const testoRisposta = (dati.content || []).map((b) => b.text || '').join('');
  const m = testoRisposta.match(/\{[\s\S]*\}/);
  let esito = null;
  try { esito = m ? JSON.parse(m[0]) : null; } catch (e) { esito = null; }
  if (!esito || !['coerente', 'dubbia', 'non_pertinente'].includes(esito.giudizio)) {
    return json({ giudizio: 'dubbia', motivazione: 'Risposta del modello non interpretabile: verifica manuale.', passaggio: '' }, 200, origin);
  }
  return json({ giudizio: esito.giudizio, motivazione: String(esito.motivazione || '').slice(0, 400), passaggio: String(esito.passaggio || '').slice(0, 300), modello: dati.model || null }, 200, origin);
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const origin = request.headers.get('Origin') || '';
    if (request.method === 'OPTIONS') return new Response(null, { status: 204, headers: cors(origin) });
    if (origin !== ORIGINE) return json({ errore: 'Origine non ammessa' }, 403, origin);
    if (request.method !== 'POST') return json({ errore: 'Metodo non ammesso' }, 405, origin);

    if (url.pathname === '/pertinenza') return pertinenza(request, env, origin);

    if (url.pathname === '/api/v1/atto/dettaglio-atto-urn') {
      const corpo = await request.text();
      if (corpo.length > 2000) return json({ errore: 'Richiesta troppo grande' }, 413, origin);
      const r = await fetch(NORMATTIVA + url.pathname, { method: 'POST', headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' }, body: corpo });
      return new Response(await r.text(), { status: r.status, headers: { ...cors(origin), 'Content-Type': 'application/json; charset=utf-8' } });
    }
    return json({ errore: 'Percorso non ammesso' }, 404, origin);
  }
};
