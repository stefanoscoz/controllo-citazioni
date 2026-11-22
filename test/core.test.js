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
t('alias TUIR', () => assert.strictEqual(CC.urnArticolo(cit[5]), 'urn:nir:stato:decreto.presidente.repubblica:1986-12-22;917~art96'));
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

console.log('\n' + ok + ' test superati' + (process.exitCode ? ' — CI SONO ERRORI' : ''));
