// Test della logica pura: node test/core.test.js
'use strict';
const assert = require('assert');
const CC = require('../core.js');

let ok = 0;
function t(nome, fn) {
  try { fn(); ok++; console.log('  ✓ ' + nome); }
  catch (e) { console.error('  ✗ ' + nome + '\n    ' + e.message); process.exitCode = 1; }
}

const paragrafi = [
  { testo: "Oggetto: permanenza nel regime forfetario – periodo d'imposta 2022" },
  { testo: "Il regime è riservato ai contribuenti che nell'anno precedente hanno percepito compensi, ragguagliati ad anno, non superiori a euro 65.000, ai sensi dell'art. 1, comma 54, della L. 23 dicembre 2014, n. 190. Sono inoltre previste cause ostative connesse al possesso di partecipazioni, disciplinate dall'art. 1, comma 71, della medesima legge.", sezione: '§ 2' },
  { testo: "Nei rapporti con l'Amministrazione finanziaria valgono i principi di collaborazione e buona fede richiamati dall'art. 10, comma 1, della L. 27 luglio 2000, n. 212. L'Agenzia delle Entrate ha fornito chiarimenti sul regime con la circolare n. 32/E del 5 dicembre 2023. Secondo la giurisprudenza di legittimità, il superamento della soglia comporta l'uscita dal regime dall'anno successivo (Cass., sez. V, 30 febbraio 2024, n. 12345).", sezione: '§ 3' },
  { testo: "Si vedano anche l'art. 96, comma 5, del TUIR, l'art. 2423 c.c., l'art. 13, comma 1, lett. a-bis), del D.Lgs. 472/1997 e l'art. 16-bis del D.L. 29 novembre 2008, n. 185." }
];

const cit = CC.estraiCitazioni(paragrafi);

console.log('Estrazione');
t('trova 9 citazioni', () => assert.strictEqual(cit.length, 9, cit.map(c => c.testo).join(' | ')));
t('art. 1, c. 54, L. 190/2014 con data completa', () => {
  const c = cit[0];
  assert.strictEqual(c.genere, 'norma');
  assert.strictEqual(c.articolo, '1'); assert.strictEqual(c.comma, '54');
  assert.strictEqual(c.atto.urn, 'urn:nir:stato:legge:2014-12-23;190');
  assert.strictEqual(c.testo, 'art. 1, comma 54, della L. 23 dicembre 2014, n. 190');
  assert.strictEqual(c.sezione, '§ 2');
});
t('"della medesima legge" eredita l\'atto precedente', () => {
  const c = cit[1];
  assert.strictEqual(c.comma, '71'); assert.ok(c.relativo);
  assert.strictEqual(c.atto.urn, 'urn:nir:stato:legge:2014-12-23;190');
});
t('Statuto del contribuente', () => assert.strictEqual(CC.urnArticolo(cit[2]), 'urn:nir:stato:legge:2000-07-27;212~art10'));
t('circolare riconosciuta come prassi', () => { assert.strictEqual(cit[3].genere, 'prassi'); assert.strictEqual(cit[3].numero, 32); });
t('sentenza con data impossibile -> non trovata bloccante', () => {
  const c = cit[4];
  assert.strictEqual(c.genere, 'sentenza'); assert.strictEqual(c.dataValida, false);
  const e = CC.esitoFormale(c); assert.strictEqual(e.esito, 'ko'); assert.ok(e.bloccante);
});
t('alias TUIR', () => assert.strictEqual(CC.urnArticolo(cit[5]), 'urn:nir:stato:decreto.del.presidente.della.repubblica:1986-12-22;917~art96'));
t('alias c.c. (non scambiato per "comma")', () => { assert.strictEqual(cit[6].comma, undefined); assert.strictEqual(cit[6].atto.key, 'regio.decreto:262:1942'); });
t('forma breve D.Lgs. 472/1997 completata dal dizionario + lettera a-bis', () => {
  const c = cit[7];
  assert.strictEqual(c.lettera, 'abis');
  assert.strictEqual(c.atto.urn, 'urn:nir:stato:decreto.legislativo:1997-12-18;472');
});
t('art. 16-bis di decreto-legge con data', () => assert.strictEqual(CC.urnArticolo(cit[8]), 'urn:nir:stato:decreto.legge:2008-11-29;185~art16bis'));

console.log('Data di riferimento');
t('periodo d\'imposta 2022 -> 2022-12-31', () => assert.strictEqual(CC.suggerisciDataRiferimento(paragrafi[0].testo, '2026-11-22').data, '2022-12-31'));
t('date impossibili', () => { assert.ok(!CC.dataValida(30, 2, 2024)); assert.ok(CC.dataValida(29, 2, 2024)); assert.ok(!CC.dataValida(29, 2, 2023)); });

console.log('Analisi della risposta di Normattiva (risposte simulate)');
const htmlRif = '<div><h2>Art. 1</h2><p>53. Testo del comma cinquantatre.</p><p>54. I contribuenti persone fisiche applicano il regime forfetario se nell\'anno precedente hanno conseguito ricavi ovvero hanno percepito compensi, ragguagliati ad anno, non superiori a euro 65.000.</p><p>55. Altro comma.</p></div>';
const htmlOggi = htmlRif.replace('65.000', '85.000');
t('estrae il comma 54', () => {
  const s = CC.estraiComma(CC.testoDaHtml(htmlRif), '54');
  assert.ok(s.startsWith('54.')); assert.ok(s.includes('65.000')); assert.ok(!s.includes('55.'));
});
t('testo cambiato -> Modificata con confronto', () => {
  const v = CC.valutaNorma(cit[0], { trovato: true, html: htmlRif }, { trovato: true, html: htmlOggi }, '2022-12-31');
  assert.strictEqual(v.esito, 'mod');
  assert.ok(v.confronto.some(s => s.tipo === '-' && s.testo.includes('65.000')));
  assert.ok(v.confronto.some(s => s.tipo === '+' && s.testo.includes('85.000')));
});
t('testo uguale -> Verificata', () => {
  const v = CC.valutaNorma(cit[0], { trovato: true, html: htmlRif }, { trovato: true, html: htmlRif }, '2022-12-31');
  assert.strictEqual(v.esito, 'ok');
});
t('articolo diverso restituito (trappola "primo articolo") -> Non trovata', () => {
  const v = CC.valutaNorma(cit[2], { trovato: true, html: '<h2>Art. 1</h2><p>1. Testo.</p>' }, null, '2022-12-31');
  assert.strictEqual(v.esito, 'ko');
});
t('comma assente -> Non verificabile, mai Non trovata', () => {
  const v = CC.valutaNorma(cit[1], { trovato: true, html: htmlRif }, { trovato: true, html: htmlRif }, '2022-12-31');
  assert.strictEqual(v.esito, 'nv');
});
t('errore di rete -> Non verificabile', () => {
  const v = CC.valutaNorma(cit[0], { errore: 'CORS' }, null, '2022-12-31');
  assert.strictEqual(v.esito, 'nv'); assert.ok(v.link.includes('!vig=2022-12-31'));
});
t('date API AAAAMMGG', () => { assert.strictEqual(CC.dataApi('20230101'), '2023-01-01'); assert.strictEqual(CC.dataApi('99999999'), null); });

console.log('Varianti di citazione');
function una(t) { return CC.estraiCitazioni([{ testo: t }]); }
t('artt. 2423 e 2424 c.c. -> un riferimento con due articoli', () => { const c = una('artt. 2423 e 2424 c.c.')[0]; assert.strictEqual(c.atto.key, 'regio.decreto:262:1942'); assert.deepStrictEqual(c.altriArticoli.map(a => a.articolo), ['2424']); });
t('art. 2423 e ss. c.c.', () => { const c = una('art. 2423 e ss. c.c.')[0]; assert.ok(c.seguenti); assert.strictEqual(c.atto.key, 'regio.decreto:262:1942'); });
t('art. 19, D.P.R. 633/1972 (virgola, senza "del")', () => assert.strictEqual(una('art. 19, D.P.R. 633/1972')[0].atto.key, 'decreto.presidente.repubblica:633:1972'));
t('data numerica 26/10/1972', () => assert.strictEqual(una('art. 19 D.P.R. 26/10/1972 n. 633')[0].atto.data, '1972-10-26'));
t('"n. 472 del 1997"', () => assert.strictEqual(una('art. 13 del D.Lgs. n. 472 del 1997')[0].atto.key, 'decreto.legislativo:472:1997'));
t('"comma 1 dell\'art. 19"', () => { const c = una('comma 1 dell’art. 19 del DPR 633/72')[0]; assert.strictEqual(c.comma, '1'); assert.ok(c.testo.startsWith('comma 1')); });
t('atto senza articolo: "il DPR 633/72"', () => { const c = una('il DPR 633/72 prevede')[0]; assert.strictEqual(c.genere, 'atto'); });
t('"legge di bilancio" non è una citazione', () => assert.strictEqual(una('ai sensi della legge di bilancio').length, 0));
t('alias cc senza punti', () => assert.strictEqual(una('art. 2423 cc')[0].atto.key, 'regio.decreto:262:1942'));

console.log('Abrogazioni (testi reali restituiti da Normattiva)');
t('articolo abrogato riconosciuto', () => { const a = CC.abrogazione(CC.testoDaHtml('<p>Art. 10</p><p>((ARTICOLO ABROGATO DAL D.LGS. 19 GENNAIO 2026, N. 10 ))</p>')); assert.strictEqual(a.cosa, 'articolo'); assert.ok(a.da.includes('19 GENNAIO 2026')); });
t('abrogato oggi ma vigente alla data -> Modificata', () => {
  const c = una('art. 2409 c.c.')[0];
  const v = CC.valutaNorma(c, { trovato: true, html: '<p>Art. 2409. (Denunzia al tribunale). Se vi è fondato sospetto...</p>', inizioVig: '2004-01-01', fineVig: '2026-04-28' }, { trovato: true, html: '<p>Art. 2409. ((ARTICOLO ABROGATO DAL D.LGS. 27 MARZO 2026, N. 47 ))</p>', inizioVig: '2026-04-29' }, '2025-12-31');
  assert.strictEqual(v.esito, 'mod'); assert.ok(v.abrogato); assert.ok(v.motivo.includes('29/04/2026'));
});
t('abrogato alla data -> Non trovata bloccante', () => {
  const c = una('art. 2409 c.c.')[0];
  const v = CC.valutaNorma(c, { trovato: true, html: '<p>Art. 2409. ((ARTICOLO ABROGATO DAL D.LGS. 27 MARZO 2026, N. 47 ))</p>', inizioVig: '2026-04-29' }, null, '2026-11-22');
  assert.strictEqual(v.esito, 'ko'); assert.ok(v.bloccante);
});
t('URN dei D.P.R. nel formato Normattiva', () => assert.strictEqual(CC.urnArticolo(una('art. 10 DPR 633/72')[0]), 'urn:nir:stato:decreto.del.presidente.della.repubblica:1972-10-26;633~art10'));
t('articolo non ancora esistente alla data -> Non trovata', () => { const v = CC.valutaNorma(una('art. 2409 c.c.')[0], { trovato: false, nonAncoraVigente: true }, null, '1940-01-01'); assert.strictEqual(v.esito, 'ko'); });
console.log('Struttura del testo per la lettura');
t('codice civile: commi non numerati e righe spezzate dalle modifiche', () => {
  const st = CC.strutturaNorma(CC.testoDaHtml('<div>Art. 2423.<br><br><div>(Redazione del bilancio).</div><div>Gli amministratori devono redigere il bilancio, costituito dal conto economico</div><div>((, dal rendiconto finanziario))</div><div>e dalla nota integrativa.<br><br>Il bilancio deve essere redatto con chiarezza.</div></div>'));
  assert.strictEqual(st.rubrica, 'Redazione del bilancio'); assert.strictEqual(st.commi.length, 2); assert.ok(st.commi[0].righe[0].testo.includes('((, dal rendiconto finanziario))'));
});
t('comma 2 c.c. = secondo capoverso', () => assert.strictEqual(CC.estraiComma(CC.testoDaHtml('<p>Art. 2423.</p><p>(Redazione del bilancio).</p><p>Primo capoverso.</p><p>Secondo capoverso.</p>'), '2'), 'Secondo capoverso.'));
t('elenco numerato "1)" come voci del comma (art. 10 IVA)', () => { const st = CC.strutturaNorma(CC.testoDaHtml('<h2>Art. 10</h2><div>Operazioni esenti</div><div>Sono esenti:<br>1) le prestazioni;<br>27-quater) altre;</div>')); assert.strictEqual(st.commi.length, 1); assert.strictEqual(st.commi[0].righe[2].lettera, '27-quater'); });
t('pseudonimizzazione', () => { const p = CC.pseudonimizza('RSSMRA80A01H501U, IT60X0542811101000000123456, mario@studio.it'); assert.ok(!/RSSMRA|IT60|mario@/.test(p)); });
console.log('Giurisprudenza e prassi (forme di citazione)');
t('Cassazione con sezione, tipo e data', () => { const c = una('Cass. civ., sez. trib., ord. 4 novembre 2026, n. 28437')[0]; assert.strictEqual(c.corte, 'cassazione'); assert.strictEqual(c.numero, 28437); assert.strictEqual(c.data, '2026-11-04'); assert.strictEqual(c.tipoProvv, 'Ordinanza'); assert.strictEqual(c.sezioneN, 5); });
t('Cass. n. 12345 del 2024 / Cass. 28437/2026', () => { assert.strictEqual(una('Cass. n. 12345 del 2024')[0].anno, 2024); assert.strictEqual(una('Cass. 28437/2026')[0].numero, 28437); });
t('CGT di merito con luogo e grado', () => { const c = una('CGT di primo grado di Rimini, sentenza n. 100 del 26 giugno 2024')[0]; assert.strictEqual(c.corte, 'merito'); assert.strictEqual(c.luogo, 'Rimini'); assert.strictEqual(c.grado, 1); assert.strictEqual(c.data, '2024-06-26'); });
t('CTR / CTP con luogo composto', () => { assert.strictEqual(una('CTR Lombardia n. 3456/2019')[0].grado, 2); assert.strictEqual(una('CTP Reggio Emilia, sez. 2, n. 45/2018')[0].luogo, 'Reggio Emilia'); });
t('prassi: data prima del numero e /E/anno', () => { assert.strictEqual(una('circolare dell’Agenzia delle Entrate 5 dicembre 2023, n. 32/E')[0].data, '2023-12-05'); assert.strictEqual(una('circ. AdE n. 32/E/2023')[0].anno, 2023); });
t('nessun falso positivo su «la Cassazione ha chiarito»', () => assert.strictEqual(una('la Cassazione ha chiarito').length, 0));

console.log('DEF (risposte reali dello sportello /def)');
const R_CIRC = { totale: 1, risultati: [{ id: '{93911C48-035E-4FE4-B79D-4E5E2D1B1E43}', estremi: 'Circolare del 05/12/2023 n. 32 - Agenzia delle Entrate - Direzione Centrale Coordinamento Normativo', titolo: 'Articolo 1, comma 54, della legge 29 dicembre 2022, n. 197 (legge di bilancio 2023) - Modifiche al regime forfetario' }] };
const R_CASS = { totale: 1, risultati: [{ id: '{5020E9A1-0000-C41E-B7D8-67342CB2EE13}', estremi: 'Ordinanza del 04/11/2026 n. 28437 - Corte di Cassazione - Sezione/Collegio 5', titolo: 'Riscossione- Cartella di pagamento - Pericolo riscossione' }] };
const R_100 = { totale: 2, risultati: [{ id: '{363038C2-0000-0000-0000-000000000000}', estremi: 'Sentenza del 26/06/2024 n. 100 - Corte di giustizia tributaria di primo grado di Rimini - Sezione/Collegio 1', titolo: '' }, { id: '{4FFD4A2B-0000-0000-0000-000000000000}', estremi: 'Sentenza del 30/01/2024 n. 100 - Corte di Giustizia UE - Sezione/Collegio 8', titolo: 'IVA' }] };
t('circolare trovata con data giusta -> Verificata + link al documento', () => { const v = CC.valutaDef(una('circolare n. 32/E del 5 dicembre 2023')[0], R_CIRC); assert.strictEqual(v.esito, 'ok'); assert.ok(v.linkDoc.includes('getPrassiDetail.do?id=')); assert.ok(v.motivo.includes('regime forfetario')); });
t('data diversa -> Data diversa', () => { const v = CC.valutaDef(una('circolare n. 32/E del 5 novembre 2023')[0], R_CIRC); assert.strictEqual(v.esito, 'data'); assert.ok(v.motivo.includes('05/12/2023')); });
t('Cassazione: tipo diverso segnalato come nota', () => { const v = CC.valutaDef(una('Cass., sez. V, sent. 4 novembre 2026, n. 28437')[0], R_CASS); assert.strictEqual(v.esito, 'ok'); assert.ok(v.motivo.includes('risulta come Ordinanza')); });
t('assente in DEF -> Non verificabile, mai Non trovata', () => { const v = CC.valutaDef(una('Cass. n. 6789/2024')[0], { totale: 0, risultati: [] }); assert.strictEqual(v.esito, 'nv'); assert.ok(v.link.includes('italgiure')); });
t('merito: filtra per luogo tra più risultati', () => { const v = CC.valutaDef(una('CGT I grado Rimini, sent. n. 100/2024')[0], R_100); assert.strictEqual(v.esito, 'ok'); assert.ok(v.def.estremi.includes('Rimini')); });
t('merito: luogo assente nei risultati -> Non verificabile', () => assert.strictEqual(CC.valutaDef(una('CGT I grado Milano, sent. n. 100/2024')[0], R_100).esito, 'nv'));
t('errore del Worker -> Non verificabile con link di ricerca', () => { const v = CC.valutaDef(una('circ. AdE n. 32/E/2023')[0], { errore: 'HTTP 502' }); assert.strictEqual(v.esito, 'nv'); assert.ok(v.linkDef.includes('executeAdvancedPrassiSearch.do')); });
t('link di ricerca DEF (prassi) con i parametri provati dal vivo', () => { const u = CC.urlRicercaDef(una('circolare n. 32/E del 5 dicembre 2023')[0]); assert.ok(/tipoEstremi=Circolare&numero=32&annoDataEmissioneDa=2023&annoDataEmissioneA=2023&ente=Agenzia\+delle\+Entrate/.test(u)); });
t('senza Worker: prassi e Cassazione hanno il link «Cerca su DEF»', () => { assert.ok(CC.esitoFormale(una('Cass. n. 28437/2026')[0]).linkDef.includes('ente=Corte+di+Cassazione')); assert.ok(CC.esitoFormale(una('risoluzione n. 1/E del 2024')[0]).linkDef); });
console.log('\n' + ok + ' test superati' + (process.exitCode ? ' — CI SONO ERRORI' : ''));
