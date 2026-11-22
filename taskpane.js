/* Controllo citazioni — pannello Word (Office.js). Nessun server: tutto gira nel pannello. */
/* global Office, Word, CC */
(function () {
  'use strict';

  var API_NORMATTIVA = 'https://api.normattiva.it/t/normattiva.api/bff-opendata/v1';
  var PERCORSO_DETTAGLIO = '/api/v1/atto/dettaglio-atto-urn';
  var PREFISSO_COMMENTO = '[Controllo citazioni] ';
  var COLORI = { ok: '#17694D', mod: '#8A5100', dub: '#5B3FA8', ko: '#B42318', nv: '#4B5563' };

  var stato = {
    citazioni: [], dataRif: null, ora: null, filtro: 'rivedere',
    commenti: false, cache: {}, nota: null
  };

  // ------------------------------------------------------------------ utilità
  function $(id) { return document.getElementById(id); }
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  function oggiISO() { var d = new Date(); return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2); }
  function formatoData(iso) { if (!iso) return ''; var p = iso.split('-'); return p[2] + '/' + p[1] + '/' + p[0]; }
  function messaggio(t, ms) {
    var el = $('messaggio'); el.textContent = t || '';
    if (t) setTimeout(function () { if (el.textContent === t) el.textContent = ''; }, ms || 4000);
  }
  function leggiImpostazione(k) { try { return localStorage.getItem('cc.' + k) || ''; } catch (e) { return ''; } }
  function scriviImpostazione(k, v) { try { localStorage.setItem('cc.' + k, v); } catch (e) { /* storage non disponibile */ } }
  function mostra(vista) {
    ['vistaAvvio', 'vistaRiepilogo', 'vistaDettaglio'].forEach(function (id) { $(id).hidden = id !== vista; });
  }

  // ------------------------------------------------------------- Normattiva
  function baseApi() { return (leggiImpostazione('proxy') || API_NORMATTIVA).replace(/\/+$/, ''); }

  function chiamaNormattiva(urn) {
    if (stato.cache[urn]) return Promise.resolve(stato.cache[urn]);
    var ctrl = typeof AbortController !== 'undefined' ? new AbortController() : null;
    var timer = ctrl ? setTimeout(function () { ctrl.abort(); }, 20000) : null;
    return fetch(baseApi() + PERCORSO_DETTAGLIO, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' },
      body: JSON.stringify({ urn: urn }),
      signal: ctrl ? ctrl.signal : undefined
    }).then(function (res) {
      if (!res.ok && res.status !== 404) throw new Error('HTTP ' + res.status);
      return res.json().catch(function () { return {}; });
    }).then(function (json) {
      var d = (json && json.data) || {};
      var atto = d.atto || null;
      var lista = Array.isArray(d.lista) ? d.lista : [];
      var out;
      if (!atto && lista.length > 1) out = { trovato: false, multiplo: true };
      else {
        atto = atto || lista[0] || null;
        if (!atto || !atto.articoloHtml) out = { trovato: false };
        else out = {
          trovato: true, html: atto.articoloHtml,
          inizioVig: CC.dataApi(atto.articoloDataInizioVigenza),
          fineVig: CC.dataApi(atto.articoloDataFineVigenza),
          titolo: atto.titolo || ''
        };
      }
      stato.cache[urn] = out;
      return out;
    }).catch(function (e) {
      var motivo = e && e.name === 'AbortError' ? 'tempo scaduto' : (e && /HTTP/.test(e.message) ? e.message : 'rete o CORS');
      return { errore: motivo };
    }).then(function (r) { if (timer) clearTimeout(timer); return r; });
  }

  // Testo vigente alla data; se la sintassi "!vig=" non fosse accettata, riprova senza.
  function versioneAllaData(urnArt, data) {
    return chiamaNormattiva(urnArt + '!vig=' + data).then(function (r) {
      if (r.trovato || r.errore || r.multiplo) return r;
      return chiamaNormattiva(urnArt).then(function (r2) {
        if (r2.trovato) r2 = Object.assign({}, r2, { senzaData: true });
        return r2;
      });
    });
  }

  function verificaCitazione(c) {
    var formale = CC.esitoFormale(c);
    if (formale) {
      if (c.genere === 'norma' && c.atto && !c.atto.errore) formale.link = CC.linkNormattiva(CC.urnArticolo(c), stato.dataRif);
      if (c.genere === 'sentenza') formale.link = 'https://www.italgiure.giustizia.it/sncass/';
      return Promise.resolve(formale);
    }
    var urnArt = CC.urnArticolo(c);
    return versioneAllaData(urnArt, stato.dataRif).then(function (rif) {
      if (rif.errore || !rif.trovato) return CC.valutaNorma(c, rif, null, stato.dataRif);
      return chiamaNormattiva(urnArt + '!vig=' + oggiISO()).then(function (oggi) {
        var v = CC.valutaNorma(c, rif, oggi, stato.dataRif);
        if (rif.senzaData) v.motivo += ' Nota: versione alla data non disponibile, confronto sulle date di vigenza.';
        return v;
      });
    });
  }

  // ------------------------------------------------------------------- Word
  function leggiDocumento() {
    return Word.run(function (ctx) {
      var pars = ctx.document.body.paragraphs;
      pars.load('items/text,items/styleBuiltIn');
      return ctx.sync().then(function () {
        var sezione = null, tutto = [];
        var out = pars.items.map(function (p) {
          var t = p.text || '';
          var titolo = (p.styleBuiltIn && /^Heading/i.test(p.styleBuiltIn)) || (/^\s*\d+(\.\d+)*\.?\s+\S/.test(t) && t.length < 90);
          var mm = /^\s*(\d+(?:\.\d+)*)\.?\s/.exec(t);
          if (titolo) sezione = mm ? '§ ' + mm[1] : t.slice(0, 30);
          tutto.push(t);
          return { testo: t, sezione: titolo ? null : sezione };
        });
        return { paragrafi: out, testo: tutto.join('\n') };
      });
    });
  }

  function rimuoviMarcature(ctx) {
    var ccs = ctx.document.contentControls;
    ccs.load('items/tag');
    var comm = null;
    if (stato.commenti) { comm = ctx.document.body.getComments(); comm.load('items/content'); }
    return ctx.sync().then(function () {
      ccs.items.forEach(function (cc) {
        if (!cc.tag) return;
        if (cc.tag.indexOf('cc-cit-') === 0) cc.delete(true);          // tiene il testo
        else if (cc.tag === 'cc-nota' || cc.tag === 'cc-registro') cc.delete(false); // elimina anche il contenuto
      });
      if (comm) comm.items.forEach(function (c) { if ((c.content || '').indexOf(PREFISSO_COMMENTO) === 0) c.delete(); });
      return ctx.sync();
    });
  }

  function applicaMarcature() {
    return Word.run(function (ctx) {
      return rimuoviMarcature(ctx).then(function () {
        var pars = ctx.document.body.paragraphs;
        pars.load('items');
        return ctx.sync().then(function () {
          var ricerche = [];
          var contatori = {};
          stato.citazioni.forEach(function (c) {
            if (!c.testo || c.testo.length > 250) return;
            var chiave = c.paragrafo + '|' + c.testo;
            contatori[chiave] = (contatori[chiave] || 0) + 1;
            var res = pars.items[c.paragrafo].search(c.testo, { matchCase: true });
            res.load('items');
            ricerche.push({ c: c, res: res, occorrenza: contatori[chiave] - 1 });
          });
          return ctx.sync().then(function () {
            ricerche.forEach(function (r) {
              var range = r.res.items[r.occorrenza] || r.res.items[0];
              if (!range) return;
              var v = r.c.verifica || {};
              var cc = range.insertContentControl();
              cc.tag = 'cc-cit-' + r.c.id;
              cc.title = '[' + r.c.id + '] ' + CC.ESITI[v.esito].etichetta + (v.bloccante ? ' · bloccante' : '');
              cc.color = COLORI[v.esito];
              cc.appearance = 'BoundingBox';
              if (stato.commenti && v.esito !== 'ok') {
                cc.getRange('Whole').insertComment(PREFISSO_COMMENTO + '[' + r.c.id + '] ' + CC.ESITI[v.esito].etichetta + (v.bloccante ? ' (bloccante)' : '') + ' — ' + v.motivo);
              }
            });
            return ctx.sync();
          });
        });
      });
    });
  }

  function vaiAlTesto(id) {
    return Word.run(function (ctx) {
      var cc = ctx.document.contentControls.getByTag('cc-cit-' + id).getFirstOrNullObject();
      cc.load('isNullObject');
      return ctx.sync().then(function () {
        if (cc.isNullObject) { messaggio('Marcatura non trovata: ricontrolla il documento.'); return; }
        cc.select();
        return ctx.sync();
      });
    }).catch(errore);
  }

  function conteggio() {
    var n = { ok: 0, mod: 0, dub: 0, ko: 0, nv: 0 };
    stato.citazioni.forEach(function (c) { n[c.verifica.esito]++; });
    return n;
  }

  function testoNota(validata) {
    var n = conteggio();
    var esiti = [n.ok + ' verificate', n.mod + ' modificate', n.dub + ' con pertinenza dubbia', n.nv + ' non verificabili', n.ko + ' non trovate'].join(', ');
    var riga1 = 'Nota di verifica — citazioni controllate il ' + formatoData(oggiISO()) + ' alle ' + stato.ora +
      ' con Controllo citazioni 0.1; data di riferimento ' + formatoData(stato.dataRif) + '. Esiti: ' + esiti + '.';
    var nome = leggiImpostazione('nome') || '[professionista]';
    var riga2 = validata
      ? 'Verifica del professionista: eseguita da ' + nome + ' il ' + formatoData(oggiISO()) + '.'
      : 'Verifica del professionista: in attesa.';
    var riga3 = 'Documento predisposto con il supporto di strumenti informatici e di intelligenza artificiale per attività strumentali e di supporto (art. 13 L. 132/2025).';
    return [riga1, riga2, riga3];
  }

  function inserisciNota(validata) {
    return Word.run(function (ctx) {
      var vecchie = ctx.document.contentControls.getByTag('cc-nota');
      vecchie.load('items');
      return ctx.sync().then(function () {
        vecchie.items.forEach(function (cc) { cc.delete(false); });
        var righe = testoNota(validata);
        var body = ctx.document.body;
        var p1 = body.insertParagraph(righe[0], 'End');
        body.insertParagraph(righe[1], 'End');
        var p3 = body.insertParagraph(righe[2], 'End');
        var cc = p1.getRange('Whole').expandTo(p3.getRange('Whole')).insertContentControl();
        cc.tag = 'cc-nota'; cc.title = 'Nota di verifica · Controllo citazioni'; cc.appearance = 'BoundingBox';
        cc.font.size = 9; cc.font.color = '#2A3340';
        return ctx.sync();
      });
    }).then(function () { messaggio(validata ? 'Documento segnato come verificato.' : 'Nota di verifica inserita in fondo al documento.'); });
  }

  function inserisciRegistro() {
    return Word.run(function (ctx) {
      var vecchie = ctx.document.contentControls.getByTag('cc-registro');
      vecchie.load('items');
      return ctx.sync().then(function () {
        vecchie.items.forEach(function (cc) { cc.delete(false); });
        var valori = [['#', 'Citazione', 'Esito', 'Motivo', 'Fonte']];
        stato.citazioni.forEach(function (c) {
          var v = c.verifica;
          valori.push([String(c.id), c.testo, CC.ESITI[v.esito].etichetta + (v.bloccante ? ' (bloccante)' : ''), v.motivo, v.urn || (c.genere === 'norma' ? '—' : 'controllo formale')]);
        });
        var body = ctx.document.body;
        var titolo = body.insertParagraph('Registro di verifica delle citazioni — ' + formatoData(oggiISO()) + ', ' + stato.ora + ' — data di riferimento ' + formatoData(stato.dataRif), 'End');
        titolo.font.bold = true;
        var tabella = titolo.insertTable(valori.length, 5, 'After', valori);
        tabella.font.size = 8;
        tabella.headerRowCount = 1;
        return ctx.sync().then(function () {
          try {
            var cc = titolo.getRange('Whole').expandTo(tabella.getRange('Whole')).insertContentControl();
            cc.tag = 'cc-registro'; cc.title = 'Registro · Controllo citazioni';
            return ctx.sync();
          } catch (e) { return null; }
        });
      });
    }).then(function () { messaggio('Registro inserito in fondo al documento.'); });
  }

  // --------------------------------------------------------------- controllo
  function avvia() {
    var data = $('inData').value;
    if (!data) { messaggio('Indica la data di riferimento.'); return; }
    stato.dataRif = data;
    $('btnAvvia').disabled = true;
    $('avanzamento').textContent = 'Lettura del documento…';
    leggiDocumento().then(function (doc) {
      stato.citazioni = CC.estraiCitazioni(doc.paragrafi);
      if (!stato.citazioni.length) { $('avanzamento').textContent = 'Nessuna citazione trovata nel documento.'; return null; }
      var fatte = 0;
      // in sequenza per non sovraccaricare il servizio pubblico
      return stato.citazioni.reduce(function (p, c) {
        return p.then(function () {
          $('avanzamento').textContent = 'Verifica ' + (++fatte) + ' di ' + stato.citazioni.length + ': ' + c.testo.slice(0, 60);
          return verificaCitazione(c).then(function (v) { c.verifica = v; });
        });
      }, Promise.resolve()).then(function () {
        var d = new Date(); stato.ora = ('0' + d.getHours()).slice(-2) + ':' + ('0' + d.getMinutes()).slice(-2);
        $('avanzamento').textContent = 'Inserimento delle marcature…';
        return applicaMarcature();
      }).then(function () {
        $('avanzamento').textContent = '';
        renderRiepilogo(); mostra('vistaRiepilogo');
      });
    }).catch(errore).then(function () { $('btnAvvia').disabled = false; });
  }

  // ------------------------------------------------------------------ vista
  function chip(esito) {
    var icone = {
      ok: '<path d="M20 6 9 17l-5-5"/>',
      mod: '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
      dub: '<circle cx="12" cy="12" r="9"/><path d="M9.5 9a2.5 2.5 0 1 1 3.5 2.3c-.6.3-1 .9-1 1.6V14"/>',
      ko: '<path d="M18 6 6 18M6 6l12 12"/>',
      nv: '<path d="M5 12h14"/>'
    };
    return '<span class="chip"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
      icone[esito] + '</svg>' + esc(CC.ESITI[esito].etichetta) + '</span>';
  }

  function ordinate() {
    return stato.citazioni.slice().sort(function (a, b) {
      return CC.ESITI[a.verifica.esito].gravita - CC.ESITI[b.verifica.esito].gravita || a.id - b.id;
    });
  }

  function renderRiepilogo() {
    var n = conteggio();
    var tot = stato.citazioni.length;
    var bloccanti = stato.citazioni.filter(function (c) { return c.verifica.bloccante; }).length;
    var daRivedere = tot - n.ok;
    $('ctxData').textContent = formatoData(stato.dataRif);
    $('ctxOra').textContent = 'oggi, ' + stato.ora;
    $('totale').textContent = tot + (tot === 1 ? ' citazione trovata' : ' citazioni trovate');
    $('daRivedere').textContent = daRivedere + ' da rivedere';
    $('avvisoBloccante').hidden = !bloccanti;
    $('avvisoTitolo').textContent = bloccanti + (bloccanti === 1 ? ' citazione bloccante' : ' citazioni bloccanti');
    $('fRivedere').textContent = 'Da rivedere (' + daRivedere + ')';
    $('fTutte').textContent = 'Tutte (' + tot + ')';

    var ordine = ['ko', 'dub', 'mod', 'nv', 'ok'];
    $('barra').innerHTML = ordine.map(function (k) {
      return n[k] ? '<div style="flex:' + n[k] + ';background:' + COLORI[k] + '"></div>' : '';
    }).join('');
    $('conteggi').innerHTML = ordine.map(function (k) {
      return '<span style="color:' + COLORI[k] + '">' + n[k] + ' ' + esc(CC.ESITI[k].etichetta) + '</span>';
    }).join('');

    var voci = ordinate().filter(function (c) { return stato.filtro === 'tutte' || c.verifica.esito !== 'ok'; });
    $('elenco').innerHTML = voci.length ? voci.map(function (c) {
      var v = c.verifica;
      return '<article class="voce esito-' + v.esito + (v.bloccante ? ' bloccante' : '') + '">' +
        '<div class="voce-testa"><span class="numero">' + c.id + '</span>' + chip(v.esito) +
        (v.bloccante ? '<span class="chip chip-blocco">Bloccante</span>' : '') +
        '<span class="sezione">' + esc(c.sezione || '') + '</span></div>' +
        '<div class="voce-titolo">' + esc(c.testo) + '</div>' +
        '<div class="voce-motivo">' + esc(v.motivo) + '</div>' +
        '<div class="voce-azioni">' +
        '<button type="button" data-vai="' + c.id + '">Vai al testo</button>' +
        '<button type="button" data-dettaglio="' + c.id + '">Dettaglio</button>' +
        (v.link ? '<a href="' + esc(v.link) + '" target="_blank" rel="noopener">Apri fonte</a>' : '') +
        '</div></article>';
    }).join('') : '<p class="vuoto">Nessuna citazione da rivedere.</p>';

    aggiornaPulsanteSegna();
  }

  function renderConfronto(seg) {
    if (!seg) return '';
    return seg.map(function (s) {
      if (s.tipo === '=') return esc(s.testo);
      return '<span class="' + (s.tipo === '-' ? 'del' : 'ins') + '">' + esc(s.testo) + '</span>';
    }).join(' ');
  }

  function renderDettaglio(id) {
    var c = stato.citazioni.filter(function (x) { return x.id === id; })[0];
    if (!c) return;
    var v = c.verifica;
    var h = '<article class="esito-' + v.esito + '" style="display:flex;flex-direction:column;gap:14px">' +
      '<div class="voce-testa"><span class="numero">' + c.id + '</span>' + chip(v.esito) +
      (v.bloccante ? '<span class="chip chip-blocco">Bloccante</span>' : '') +
      '<span class="sezione">' + esc(c.sezione || '') + '</span></div>' +
      '<div class="det-blocco"><h3>Nel documento</h3><p class="citato">' + esc(c.testo) + '</p></div>' +
      '<div class="det-blocco"><h3>Esito</h3><p class="aiuto" style="color:#22303F">' + esc(v.motivo) + '</p></div>';
    if (c.atto && c.atto.nome) h += '<div class="det-blocco"><h3>Atto</h3><p class="aiuto">' + esc(c.atto.nome) + '</p></div>';
    if (v.confronto) {
      h += '<div class="det-blocco"><h3>Confronto: vigente al ' + esc(formatoData(stato.dataRif)) + ' → oggi</h3><div class="estratto">' + renderConfronto(v.confronto) + '</div></div>';
    } else if (v.estrattoRif) {
      h += '<div class="det-blocco"><h3>Testo vigente al ' + esc(formatoData(stato.dataRif)) + '</h3><div class="estratto">' + esc(v.estrattoRif.slice(0, 2500)) + '</div></div>';
    }
    if (v.urn) h += '<div class="det-blocco"><h3>Identificativo</h3><a class="urn" href="' + esc(v.link) + '" target="_blank" rel="noopener">' + esc(v.urn) + '</a>' +
      '<p class="aiuto">Fonte: Normattiva (IPZS), testo non ufficiale. Consultata oggi alle ' + esc(stato.ora) + '.</p></div>';
    else if (v.link) h += '<div class="det-blocco"><a href="' + esc(v.link) + '" target="_blank" rel="noopener">Apri la fonte per verificare</a></div>';
    h += '<div class="riga-pulsanti"><button type="button" class="secondario" data-vai="' + c.id + '">Vai al testo</button></div></article>';
    $('dettaglio').innerHTML = h;
    mostra('vistaDettaglio');
  }

  function aggiornaPulsanteSegna() {
    var bloccanti = stato.citazioni.some(function (c) { return c.verifica && c.verifica.bloccante; });
    $('chkVerificato').disabled = bloccanti;
    if (bloccanti) $('chkVerificato').checked = false;
    $('btnSegna').disabled = bloccanti || !$('chkVerificato').checked;
  }

  // ------------------------------------------------------------- test API
  function testConnessione() {
    var out = $('esitoTest');
    out.textContent = 'Verifica in corso…';
    stato.cache = {};
    var prove = [
      { urn: 'urn:nir:stato:legge:2000-07-27;212~art10', atteso: '10', nome: 'Statuto del contribuente, art. 10' },
      { urn: CC.ATTI['regio.decreto:262:1942'] ? 'urn:nir:stato:regio.decreto:1942-03-16;262:2~art2423' : null, atteso: '2423', nome: 'Codice civile, art. 2423 (controllo allegato)' }
    ];
    var righe = [];
    return prove.reduce(function (p, prova) {
      return p.then(function () {
        if (!prova.urn) return;
        return chiamaNormattiva(prova.urn + '!vig=' + oggiISO()).then(function (r) {
          if (r.errore) righe.push('✗ ' + prova.nome + ': ' + r.errore + (r.errore === 'rete o CORS' ? ' → serve il proxy (README, sezione 6).' : ''));
          else if (!r.trovato) righe.push('✗ ' + prova.nome + ': nessun testo restituito.');
          else {
            var art = CC.articoloRestituito(CC.testoDaHtml(r.html));
            righe.push((art === prova.atteso ? '✓ ' : '⚠ ') + prova.nome + ': restituito art. ' + (art || '?'));
          }
        });
      });
    }, Promise.resolve()).then(function () {
      out.textContent = righe.join('\n') + '\nServizio: ' + baseApi();
    });
  }

  function errore(e) {
    var t = (e && (e.message || e.code)) || String(e);
    if (e && e.code === 'ApiNotFound') t = 'Funzione non disponibile in questa versione di Word.';
    messaggio('Errore: ' + t, 7000);
    if (window.console) console.error(e);
  }

  // ------------------------------------------------------------------ avvio
  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Word) { messaggio('Questo componente funziona solo in Word.'); return; }
    if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) {
      messaggio('Word troppo vecchio: serve WordApi 1.3 o successiva.', 10000); $('btnAvvia').disabled = true; return;
    }
    stato.commenti = Office.context.requirements.isSetSupported('WordApi', '1.4');

    $('inNome').value = leggiImpostazione('nome');
    $('inProxy').value = leggiImpostazione('proxy');
    $('inData').value = oggiISO();

    leggiDocumento().then(function (doc) {
      var s = CC.suggerisciDataRiferimento(doc.testo, oggiISO());
      $('inData').value = s.data;
      $('fonteData').textContent = 'Rilevata da: ' + s.fonte + '. Le norme sono confrontate con il testo vigente a questa data.';
    }).catch(function () { /* documento vuoto o non leggibile */ });

    $('btnAvvia').addEventListener('click', avvia);
    $('btnRicontrolla').addEventListener('click', function () { stato.cache = {}; mostra('vistaAvvio'); avvia(); });
    $('btnRimuovi').addEventListener('click', function () {
      Word.run(function (ctx) { return rimuoviMarcature(ctx); }).then(function () { messaggio('Marcature rimosse.'); mostra('vistaAvvio'); }).catch(errore);
    });
    $('btnNota').addEventListener('click', function () { inserisciNota(false).catch(errore); });
    $('btnRegistro').addEventListener('click', function () { inserisciRegistro().catch(errore); });
    $('chkVerificato').addEventListener('change', aggiornaPulsanteSegna);
    $('btnSegna').addEventListener('click', function () {
      if (!leggiImpostazione('nome')) { messaggio('Indica il tuo nome nelle impostazioni (icona in alto a destra).', 6000); return; }
      inserisciNota(true).catch(errore);
    });
    $('fRivedere').addEventListener('click', function () { stato.filtro = 'rivedere'; $('fRivedere').setAttribute('aria-pressed', 'true'); $('fTutte').setAttribute('aria-pressed', 'false'); renderRiepilogo(); });
    $('fTutte').addEventListener('click', function () { stato.filtro = 'tutte'; $('fTutte').setAttribute('aria-pressed', 'true'); $('fRivedere').setAttribute('aria-pressed', 'false'); renderRiepilogo(); });
    $('btnIndietro').addEventListener('click', function () { mostra('vistaRiepilogo'); });
    document.body.addEventListener('click', function (ev) {
      var a = ev.target.closest ? ev.target.closest('a[target=_blank]') : null;
      if (a && Office.context.ui && Office.context.ui.openBrowserWindow &&
          Office.context.requirements.isSetSupported('OpenBrowserWindowApi', '1.1')) {
        ev.preventDefault(); Office.context.ui.openBrowserWindow(a.href); return;
      }
      var t = ev.target.closest ? ev.target.closest('[data-vai],[data-dettaglio]') : null;
      if (!t) return;
      if (t.hasAttribute('data-vai')) vaiAlTesto(parseInt(t.getAttribute('data-vai'), 10));
      else renderDettaglio(parseInt(t.getAttribute('data-dettaglio'), 10));
    });
    $('btnImpostazioni').addEventListener('click', function () {
      var p = $('impostazioni'); p.hidden = !p.hidden;
      $('btnImpostazioni').setAttribute('aria-expanded', String(!p.hidden));
    });
    $('btnSalva').addEventListener('click', function () {
      scriviImpostazione('nome', $('inNome').value.trim());
      scriviImpostazione('proxy', $('inProxy').value.trim());
      stato.cache = {};
      messaggio('Impostazioni salvate.');
    });
    $('btnTest').addEventListener('click', function () { testConnessione().catch(errore); });
  });
})();
