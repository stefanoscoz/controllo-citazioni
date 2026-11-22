/**
 * Proxy (facoltativo) per il plugin "Controllo citazioni" — Cloudflare Worker.
 *
 * 1. /def                             → verifica di prassi e giurisprudenza tributaria su DEF (MEF, def.finanze.it)
 * 2. /def/massima                     → massima e intitolazione di una sentenza presente in DEF
 * 3. /pertinenza                      → verifica di pertinenza con l'API Claude (livello 3, facoltativa)
 * 4. /api/v1/atto/dettaglio-atto-urn  → inoltro a Normattiva (riserva: il plugin di norma la interroga direttamente)
 *
 * Configurazione (Cloudflare → Worker → Settings → Variables and Secrets):
 *   ANTHROPIC_API_KEY  (Secret, solo per /pertinenza; /def funziona anche senza)
 *   MODELLO            (facoltativa, es. "claude-sonnet-5-5")
 *   PONTE_DEF          (facoltativa, indirizzo del ponte Netlify, es. https://ponte-def-xxx.netlify.app)
 *   CHIAVE_PONTE       (Secret, uguale a quella impostata su Netlify)
 * Il Worker non registra né conserva testi. Ammette solo richieste dalla tua pagina GitHub Pages.
 */
const ORIGINE = 'https://stefanoscoz.github.io';
const NORMATTIVA = 'https://api.normattiva.it/t/normattiva.api/bff-opendata/v1';
const MODELLO_PREDEFINITO = 'claude-sonnet-5-5';
const DEF = 'https://def.finanze.it/DocTribFrontend/';
const DEF_HTTP = 'http://def.finanze.it/DocTribFrontend/';
const INTESTAZIONI_DEF = { 'Accept': 'text/html', 'User-Agent': 'Mozilla/5.0 (Controllo citazioni)' };

// Cloudflare rifiuta i certificati che non riesce a validare (HTTP 525/526, es. catena
// incompleta lato server). In quel caso si riprova in http, senza seguire redirect verso https.
async function scaricaDef(percorso, traccia, env) {
  let r;
  // Ponte DEF (Netlify): aggiunge l'intermedio Sectigo mancante, vedi ponte-def/
  if (env && env.PONTE_DEF) {
    const i = percorso.indexOf('?');
    const indirizzo = env.PONTE_DEF.replace(/\/+$/, '') + '/def?pagina=' + encodeURIComponent(percorso.slice(0, i < 0 ? undefined : i)) + (i < 0 ? '' : '&' + percorso.slice(i + 1));
    try {
      r = await fetch(indirizzo, { headers: { 'x-chiave-ponte': env.CHIAVE_PONTE || '' } });
      if (traccia) traccia.push('ponte: HTTP ' + r.status);
      if (r.ok) return { ok: true, testo: await r.text(), via: 'ponte' };
      return { ok: false, errore: 'ponte HTTP ' + r.status + (r.status === 401 ? ' (CHIAVE_PONTE diversa tra Worker e Netlify)' : ': ' + (await r.text()).slice(0, 200)) };
    } catch (e) {
      return { ok: false, errore: 'ponte non raggiungibile: ' + String(e && e.message || e) };
    }
  }
  try {
    r = await fetch(DEF + percorso, { headers: INTESTAZIONI_DEF });
    if (traccia) traccia.push('https: HTTP ' + r.status);
    if (r.ok) return { ok: true, testo: await r.text(), via: 'https' };
    if (r.status !== 525 && r.status !== 526) return { ok: false, errore: 'HTTP ' + r.status };
  } catch (e) {
    if (traccia) traccia.push('https: ' + String(e && e.message || e));
  }
  try {
    r = await fetch(DEF_HTTP + percorso, { headers: INTESTAZIONI_DEF, redirect: 'manual' });
    if (traccia) traccia.push('http: HTTP ' + r.status + (r.headers.get('Location') ? ' → ' + r.headers.get('Location') : ''));
    if (r.ok) return { ok: true, testo: await r.text(), via: 'http' };
    if (r.status >= 300 && r.status < 400) return { ok: false, errore: 'certificato di def.finanze.it non accettato da Cloudflare (526) e http reindirizzato su https' };
    return { ok: false, errore: 'HTTP ' + r.status };
  } catch (e) {
    return { ok: false, errore: 'http: ' + String(e && e.message || e) };
  }
}

// ---------------------------------------------------------------- DEF (MEF)
// Le pagine di DEF incorporano i risultati in una variabile JavaScript con XML ("var xmlResult = '...'").
function estraiVariabile(html, nome) {
  const m = html.match(new RegExp('var\\s+' + nome + "\\s*=\\s*'([\\s\\S]*?)';\\s*(?://|\\n|\\r|</script>)"));
  return m ? m[1].replace(/\\(.)/g, '$1') : null;
}
function testoPulito(s) {
  return String(s || '').replace(/<!\[CDATA\[|\]\]>/g, '').replace(/<!--[\s\S]*?-->/g, '').replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&#39;|&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/\s+/g, ' ').trim();
}
function analizzaRisultati(html) {
  const xml = estraiVariabile(html, 'xmlResult');
  if (xml === null) return { errore: 'Formato della pagina DEF non riconosciuto' };
  // pagina troncata con xmlResult vuoto: DEF non ha completato la ricerca (non vuol dire «0 risultati»)
  if (!xml.trim()) return { errore: 'DEF ha restituito una pagina vuota: ricerca momentaneamente non disponibile' };
  const tot = (xml.match(/<totaleProvvedimenti>(\d+)<\/totaleProvvedimenti>/) || [])[1];
  const risultati = [];
  const re = /<Provvedimento idProvvedimento="([^"]+)"[^>]*>([\s\S]*?)<\/Provvedimento>/g;
  let m;
  while ((m = re.exec(xml)) && risultati.length < 10) {
    const estremi = testoPulito((m[2].match(/<estremi[^>]*>([\s\S]*?)<\/estremi>/) || [])[1]);
    const titolo = testoPulito((m[2].match(/<titoloProvvedimento>([\s\S]*?)<\/titoloProvvedimento>/) || [])[1]).slice(0, 300);
    risultati.push({ id: m[1], estremi, titolo });
  }
  return { totale: tot ? parseInt(tot, 10) : risultati.length, risultati };
}
function analizzaDettaglio(html) {
  const xml = estraiVariabile(html, 'xmlDettaglio');
  if (xml === null) return { errore: 'Formato della pagina DEF non riconosciuto' };
  const pezzo = (tag) => testoPulito((xml.match(new RegExp('<' + tag + '[^>]*>([\\s\\S]*?)</' + tag + '>')) || [])[1]);
  return { estremi: pezzo('estremi'), intitolazione: pezzo('intitolazione'), massima: pezzo('massima').slice(0, 4000) };
}
const TIPI_AMMESSI = /^(|Circolare|Risoluzione|Interpello|Risposta|Nota|Provvedimento|Sentenza|Ordinanza|Decreto|Decisione)$/;

async function def(url, origin, env) {
  const q = url.searchParams;
  const ambito = q.get('ambito') === 'G' ? 'G' : 'P';
  const tipo = q.get('tipo') || '';
  const numero = (q.get('numero') || '').replace(/\D/g, '');
  const anno = (q.get('anno') || '').replace(/\D/g, '');
  const ente = (q.get('ente') || '').slice(0, 120);
  if (!numero || anno.length !== 4 || !TIPI_AMMESSI.test(tipo)) return json({ errore: 'Parametri non validi' }, 400, origin);
  // DEF (da novembre 2026) completa la ricerca solo se riceve TUTTI i campi del modulo, anche vuoti
  const p = new URLSearchParams(ambito === 'G'
    ? [['js_enabled', '0'], ['tipoComplessitaRicerca', 'avanzata'], ['ricercaAreaRiservata', 'false'], ['tipoRicerca', 'RA'], ['device', 'D'], ['ambitoRicerca', 'G']]
    : [['js_enabled', '0'], ['tipoComplessitaRicerca', 'avanzata'], ['tipoRicerca', 'RA'], ['device', 'D'], ['ambitoRicerca', 'P']]);
  [['parole', ''], ['ricercaNelTitolo', ''], ['tipoCriterioRicerca', '0'], ['tipo_ord', 'DATA'], ['tipoEstremi', tipo], ['numero', numero],
    ['giornoDataEmissioneDa', ''], ['meseDataEmissioneDa', ''], ['annoDataEmissioneDa', anno], ['dataEmissioneDa', ''],
    ['giornoDataEmissioneA', ''], ['meseDataEmissioneA', ''], ['annoDataEmissioneA', anno], ['dataEmissioneA', ''],
    ['ente', ente], ['superEnte', ''], ['materiaFiscale', ''], ['classificazioneArgomento', '']].forEach(([k, v]) => p.append(k, v));
  const pagina = ambito === 'G' ? 'executeAdvancedGiurisprudenzaSearch.do' : 'executeAdvancedPrassiSearch.do';
  const r = await scaricaDef(pagina + '?' + p.toString(), null, env);
  if (!r.ok) return json({ errore: 'DEF: ' + r.errore }, 502, origin);
  const esito = analizzaRisultati(r.testo);
  if (esito.errore) return json(esito, 502, origin);
  esito.ricerca = DEF + pagina + '?' + p.toString();
  return json(esito, 200, origin);
}

async function defMassima(url, origin, env) {
  const id = url.searchParams.get('id') || '';
  if (!/^\{[0-9A-F-]{36}\}$/i.test(id)) return json({ errore: 'Identificativo non valido' }, 400, origin);
  const r = await scaricaDef('getGiurisprudenzaDetail.do?id=' + encodeURIComponent(id), null, env);
  if (!r.ok) return json({ errore: 'DEF: ' + r.errore }, 502, origin);
  const d = analizzaDettaglio(r.testo);
  return json(d, d.errore ? 502 : 200, origin);
}

function cors(origin) {
  return {
    'Access-Control-Allow-Origin': origin === ORIGINE ? ORIGINE : 'null',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
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

export { analizzaRisultati, analizzaDettaglio, estraiVariabile }; // per i test

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const origin = request.headers.get('Origin') || '';
    if (request.method === 'OPTIONS') return new Response(null, { status: 204, headers: cors(origin) });
    if (request.method === 'GET' && url.pathname === '/diagnosi') {
      // apribile anche dal browser: prova la connessione a DEF e non espone dati riservati
      const traccia = [];
      const r = await scaricaDef('executeAdvancedPrassiSearch.do?js_enabled=0&tipoComplessitaRicerca=avanzata&tipoRicerca=RA&device=D&ambitoRicerca=P&parole=&ricercaNelTitolo=&tipoCriterioRicerca=0&tipo_ord=DATA&tipoEstremi=Circolare&numero=32&giornoDataEmissioneDa=&meseDataEmissioneDa=&annoDataEmissioneDa=2023&dataEmissioneDa=&giornoDataEmissioneA=&meseDataEmissioneA=&annoDataEmissioneA=2023&dataEmissioneA=&ente=Agenzia+delle+Entrate&superEnte=&materiaFiscale=&classificazioneArgomento=', traccia, env);
      const esito = r.ok ? analizzaRisultati(r.testo) : null;
      return new Response(JSON.stringify({ versione: '0.4.4', def: r.ok ? (esito.errore ? 'raggiunto via ' + r.via + ', ma ' + esito.errore : 'ok via ' + r.via) : 'errore: ' + r.errore, risultati: esito && !esito.errore ? esito.totale : null, tentativi: traccia }, null, 2), { status: 200, headers: { 'Content-Type': 'application/json; charset=utf-8' } });
    }
    if (origin !== ORIGINE) return json({ errore: 'Origine non ammessa' }, 403, origin);
    if (request.method === 'GET' && url.pathname === '/def') return def(url, origin, env);
    if (request.method === 'GET' && url.pathname === '/def/massima') return defMassima(url, origin, env);
    if (request.method === 'GET' && url.pathname === '/stato') return json({ ok: true, ia: !!env.ANTHROPIC_API_KEY, versione: '0.4.4' }, 200, origin);
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
