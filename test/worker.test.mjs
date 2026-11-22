// Test del Worker con risposte di DEF ricostruite dal formato reale (node test/worker.test.mjs)
import assert from 'node:assert';
import worker, { analizzaRisultati, analizzaDettaglio } from '../proxy/worker.js';
const pagina = `<script type="text/javascript">
		// risultati ricerca
		var xmlResult = '<?xml version=\\"1.0\\" encoding=\\"UTF-8\\"?><risultatiRicerca><sessione \\/><contatori><contatorePrassi>1<\\/contatorePrassi><\\/contatori><risultati><pagina>1<\\/pagina><totaleProvvedimenti>1<\\/totaleProvvedimenti><Provvedimento idProvvedimento=\\"{93911C48-035E-4FE4-B79D-4E5E2D1B1E43}\\"><estremi link=\\"true\\">Circolare del 05\\/12\\/2023 n. 32 - Agenzia delle Entrate - Direzione Centrale Coordinamento Normativo<\\/estremi><titoliProvvedimento><titoloProvvedimento><![CDATA[ <p>Articolo 1, comma 54, della legge 29 dicembre 2022, n. 197 <!-- -->(legge di bilancio 2023) - Modifiche al regime forfetario per l\\'anno 2023 \\- prova<\\/p> ]]><\\/titoloProvvedimento><\\/titoliProvvedimento><\\/Provvedimento><\\/risultati><\\/risultatiRicerca>'; // submit
		function submitSaveSearch(){}</script>`;
const r = analizzaRisultati(pagina);
assert.strictEqual(r.totale, 1);
assert.strictEqual(r.risultati[0].id, '{93911C48-035E-4FE4-B79D-4E5E2D1B1E43}');
assert.strictEqual(r.risultati[0].estremi, 'Circolare del 05/12/2023 n. 32 - Agenzia delle Entrate - Direzione Centrale Coordinamento Normativo');
assert.ok(r.risultati[0].titolo.includes("per l'anno 2023 - prova"));
const vuota = analizzaRisultati(`var xmlResult = '<?xml version=\\"1.0\\"?><risultatiRicerca><risultati><totaleProvvedimenti>0<\\/totaleProvvedimenti><\\/risultati><\\/risultatiRicerca>'; // x`);
assert.strictEqual(vuota.totale, 0);
assert.ok(analizzaRisultati('<html>altro</html>').errore);
const det = analizzaDettaglio(`var xmlDettaglio = '<risultatiRicerca><risultati><Provvedimento idProvvedimento=\\"{5020E9A1-0000-C41E-B7D8-67342CB2EE13}\\"><estremi link=\\"false\\">Ordinanza del 04\\/11\\/2026 n. 28437 - Corte di Cassazione - Sezione\\/Collegio 5<\\/estremi><titoliProvvedimento><titoloProvvedimento><intitolazione><![CDATA[<p>Riscossione- Cartella di pagamento<\\/p>]]><\\/intitolazione><massima><![CDATA[<p><p style=\\"text-align: justify;\\">In tema di motivazione della cartella.<\\/p>]]><\\/massima><\\/titoloProvvedimento><\\/titoliProvvedimento><\\/Provvedimento><\\/risultati><\\/risultatiRicerca>';
`);
assert.ok(det.estremi.startsWith('Ordinanza del 04/11/2026 n. 28437'));
assert.strictEqual(det.intitolazione, 'Riscossione- Cartella di pagamento');
assert.strictEqual(det.massima, 'In tema di motivazione della cartella.');
// flusso completo dello sportello /def con DEF simulato
const chiamate = [];
globalThis.fetch = async (url) => { chiamate.push(String(url)); return new Response(pagina, { status: 200, headers: { 'Content-Type': 'text/html' } }); };
const O = { Origin: 'https://stefanoscoz.github.io' };
let res = await worker.fetch(new Request('https://w.dev/def?ambito=P&tipo=Circolare&numero=32&anno=2023&ente=Agenzia%20delle%20Entrate', { headers: O }), {});
let j = await res.json();
assert.strictEqual(res.status, 200); assert.strictEqual(j.totale, 1);
assert.strictEqual(res.headers.get('Access-Control-Allow-Origin'), 'https://stefanoscoz.github.io');
assert.ok(/executeAdvancedPrassiSearch\.do\?.*tipoEstremi=Circolare&numero=32&annoDataEmissioneDa=2023&annoDataEmissioneA=2023&ente=Agenzia\+delle\+Entrate/.test(chiamate[0]), chiamate[0]);
res = await worker.fetch(new Request('https://w.dev/def?ambito=G&tipo=&numero=28437&anno=2026&ente=Corte%20di%20Cassazione', { headers: O }), {});
assert.strictEqual(res.status, 200); assert.ok(/executeAdvancedGiurisprudenzaSearch\.do\?.*ricercaAreaRiservata=false/.test(chiamate[1]));
res = await worker.fetch(new Request('https://w.dev/def?ambito=P&tipo=Circolare&numero=32&anno=2023', { headers: { Origin: 'https://altro.example' } }), {});
assert.strictEqual(res.status, 403);
res = await worker.fetch(new Request('https://w.dev/def?ambito=P&tipo=Altro;rm&numero=32&anno=2023', { headers: O }), {});
assert.strictEqual(res.status, 400);
res = await worker.fetch(new Request('https://w.dev/stato', { headers: O }), {});
j = await res.json(); assert.strictEqual(j.ok, true); assert.strictEqual(j.ia, false);
res = await worker.fetch(new Request('https://w.dev/def/massima?id=x', { headers: O }), {});
assert.strictEqual(res.status, 400);
console.log('Worker: 20 controlli superati');
