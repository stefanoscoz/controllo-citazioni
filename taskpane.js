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
    ['vistaAvvio', 'vistaRiepilogo', 'vistaDettaglio', 'vistaConsulta'].forEach(function (id) { $(id).hidden = id !== vista; });
    var consulta = vista === 'vistaConsulta';
    $('tabConsulta').setAttribute('aria-pressed', String(consulta));
    $('tabControllo').setAttribute('aria-pressed', String(!consulta));
    if (consulta) $('suggerimento').hidden = true;
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
        else if (/NON ANCORA ESISTENTE O VIGENTE/i.test(atto.articoloHtml)) out = { trovato: false, nonAncoraVigente: true };
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
  // Sempre con la data esplicita: senza data Normattiva può restituire versioni future (es. dal 2027)
  function versioneAllaData(urnArt, data) {
    return chiamaNormattiva(urnArt + '!vig=' + data);
  }

  function verificaCitazione(c) {
    if (c.altriArticoli && c.altriArticoli.length) return verificaMultipla(c);
    if (c.genere === 'atto' && c.atto && !c.atto.errore) {
      return chiamaNormattiva(c.atto.urn + '!vig=' + stato.dataRif).then(function (r) { return CC.valutaAtto(c, r); });
    }
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
        return v;
      });
    });
  }

  // "artt. 2423 e 2424 c.c.": verifica ogni articolo e riporta l'esito peggiore
  function verificaMultipla(c) {
    var tutti = [{ articolo: c.articolo, articoloTesto: c.articoloTesto }].concat(c.altriArticoli);
    var esiti = [];
    return tutti.reduce(function (p, a) {
      return p.then(function () {
        var copia = Object.assign({}, c, { articolo: a.articolo, articoloTesto: a.articoloTesto, altriArticoli: null });
        return verificaCitazione(copia).then(function (v) { esiti.push({ a: a, v: v }); });
      });
    }, Promise.resolve()).then(function () {
      var peggiore = esiti.slice().sort(function (x, y) { return CC.ESITI[x.v.esito].gravita - CC.ESITI[y.v.esito].gravita; })[0].v;
      return Object.assign({}, peggiore, {
        motivo: esiti.map(function (e) { return 'art. ' + e.a.articoloTesto + ': ' + CC.ESITI[e.v.esito].etichetta + ' — ' + e.v.motivo; }).join(' | '),
        bloccante: esiti.some(function (e) { return e.v.bloccante; })
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
      ' con Controllo citazioni 0.2; data di riferimento ' + formatoData(stato.dataRif) + '. Esiti: ' + esiti + '.';
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


  // ------------------------------------------------------- consulta norma
  var consulta = { ultima: null, versioni: null, cit: null, timerSel: null, inCorso: false };

  function addGiorni(iso, n) {
    var d = new Date(iso + 'T12:00:00Z'); d.setUTCDate(d.getUTCDate() + n);
    return d.toISOString().slice(0, 10);
  }

  // Legge la selezione e, se possibile, la citazione completa che la contiene nel paragrafo
  function leggiSelezione() {
    return Word.run(function (ctx) {
      var sel = ctx.document.getSelection();
      sel.load('text');
      var par = sel.paragraphs.getFirst();
      par.load('text');
      return ctx.sync().then(function () {
        var st = (sel.text || '').trim();
        if (!st || st.length > 400) return null;
        var pt = par.text || '';
        var idx = pt.indexOf(st);
        if (idx >= 0) {
          var cits = CC.estraiCitazioni([{ testo: pt }]).filter(function (c) { return c.genere === 'norma' || c.genere === 'atto'; });
          var sovr = cits.filter(function (c) { return c.inizio < idx + st.length && c.fine > idx; });
          if (sovr.length) return sovr[0].testo;
        }
        var proprie = CC.estraiCitazioni([{ testo: st }]).filter(function (c) { return c.genere === 'norma' || c.genere === 'atto'; });
        return proprie.length ? proprie[0].testo : st;
      });
    });
  }

  function chiamataAllaData(urn, data) { return chiamaNormattiva(urn + '!vig=' + data); }

  // Ricostruisce le versioni dell'articolo seguendo le date di inizio/fine vigenza restituite da Normattiva
  function versioniArticolo(urnArt, data) {
    var MAX = 16;
    return chiamataAllaData(urnArt, data).then(function (v0) {
      if (v0.errore || !v0.trovato) return { errore: v0.errore, trovato: false, multiplo: v0.multiplo };
      var elenco = [v0];
      var affidabile = !(v0.inizioVig && v0.inizioVig > data);
      function avanti(cur) {
        if (!cur.fineVig || elenco.length >= MAX) return Promise.resolve();
        return chiamataAllaData(urnArt, addGiorni(cur.fineVig, 1)).then(function (nx) {
          if (!nx.trovato || nx.inizioVig === cur.inizioVig) return;
          elenco.push(nx); return avanti(nx);
        });
      }
      function indietro(cur) {
        if (!cur.inizioVig || elenco.length >= MAX) return Promise.resolve();
        var d = addGiorni(cur.inizioVig, -1);
        return chiamataAllaData(urnArt, d).then(function (pv) {
          if (!pv.trovato || pv.inizioVig === cur.inizioVig || (pv.inizioVig && pv.inizioVig > d)) return;
          elenco.unshift(pv); return indietro(pv);
        });
      }
      return avanti(v0).then(function () { return indietro(v0); }).then(function () {
        return { trovato: true, versioni: elenco, attiva: elenco.indexOf(v0), affidabile: affidabile, troncato: elenco.length >= MAX };
      });
    });
  }

  function parteDaMostrare(html, c) {
    var testo = CC.testoDaHtml(html);
    if ($('chkSoloComma').checked && c.comma) {
      var cm = CC.estraiComma(testo, c.comma);
      if (cm) return { testo: cm, comma: true };
    }
    return { testo: testo, comma: false };
  }

  function riepilogoTesto(c, r, data) {
    if (c.genere === 'atto') return [c.testo + (c.atto.nome ? ' (' + c.atto.nome + ')' : ''), 'Atto presente in Normattiva. Per la multivigenza indica un articolo (es. «art. 19 ' + c.testo + '»).', 'Fonte: Normattiva (IPZS), consultata il ' + formatoData(oggiISO()) + '; testo non ufficiale.'];
    var att = r.versioni[r.attiva];
    var ultima = r.versioni[r.versioni.length - 1];
    var righe = [];
    righe.push(c.testo + (c.atto && c.atto.nome ? ' (' + c.atto.nome + ')' : ''));
    righe.push('Versione applicabile al ' + formatoData(data) + ': in vigore ' +
      (att.inizioVig ? 'dal ' + formatoData(att.inizioVig) : '') + (att.fineVig ? ' al ' + formatoData(att.fineVig) : ' ad oggi') + '.');
    righe.push('Versioni ricostruite: ' + r.versioni.length + (r.troncato ? '+' : '') +
      (r.versioni[0].inizioVig ? ', la prima dal ' + formatoData(r.versioni[0].inizioVig) : '') + '.');
    var abrAtt = CC.abrogazione(CC.testoDaHtml(att.html));
    if (abrAtt) righe.push('Attenzione: alla data indicata l\'' + abrAtt.cosa + ' risulta abrogato (' + abrAtt.da + '). Consulta le versioni precedenti per il testo previgente.');
    var abrUlt = ultima !== att ? CC.abrogazione(CC.testoDaHtml(ultima.html)) : null;
    if (r.versioni.length > 1 && ultima.inizioVig) righe.push((abrUlt ? 'Abrogato dal ' + formatoData(ultima.inizioVig) + ' (' + abrUlt.da + ')' : 'Ultima modifica in vigore dal ' + formatoData(ultima.inizioVig)) + (ultima.inizioVig > oggiISO() ? ' — data futura.' : '.') + (r.attiva === r.versioni.length - 1 ? ' È la versione applicabile.' : ''));
    righe.push('Fonte: Normattiva (IPZS), consultata il ' + formatoData(oggiISO()) + '; testo non ufficiale.');
    return righe;
  }

  function avviaConsulta(testoCit) {
    if (consulta.inCorso) return;
    var data = $('inDataCons').value || oggiISO();
    $('inCit').value = testoCit;
    var out = $('consRisultato');
    var cits = CC.estraiCitazioni([{ testo: testoCit }]).filter(function (c) { return c.genere === 'norma' || c.genere === 'atto'; });
    var c = cits[0];
    if (!c || !c.atto || c.atto.errore) {
      out.innerHTML = '<p class="avvertenza">' + esc(c && c.atto && c.atto.errore ? c.atto.errore + '.' : 'Non riconosco la norma. Scrivi gli estremi completi, es. «art. 19 DPR 633/72» o «art. 2423 c.c.». Se il riferimento è solo «art. 5», aggiungi l\'atto.') + '</p>';
      return;
    }
    consulta.inCorso = true; consulta.cit = c; consulta.ultima = testoCit;
    out.innerHTML = '<p class="aiuto">Ricerca su Normattiva e ricostruzione delle versioni…</p>';
    var urn = c.genere === 'atto' ? c.atto.urn : CC.urnArticolo(c);
    (c.genere === 'atto' ? chiamataAllaData(urn, data).then(function (r) { return r.trovato ? { trovato: true, versioni: [r], attiva: 0, affidabile: true } : r; }) : versioniArticolo(urn, data))
      .then(function (r) { consulta.versioni = r; renderConsulta(c, r, data, urn); })
      .catch(errore)
      .then(function () { consulta.inCorso = false; });
  }

  function renderConsulta(c, r, data, urn) {
    var out = $('consRisultato');
    var link = CC.linkNormattiva(urn, data);
    if (r.errore) { out.innerHTML = '<p class="avvertenza">Normattiva non raggiungibile (' + esc(r.errore) + ').</p><a href="' + esc(link) + '" target="_blank" rel="noopener">Apri su Normattiva</a>'; return; }
    if (r.multiplo) { out.innerHTML = '<p class="avvertenza">Più atti corrispondono: indica la data completa dell\'atto.</p>'; return; }
    if (!r.trovato) { out.innerHTML = '<p class="avvertenza">Normattiva non restituisce l\'articolo indicato.</p><a href="' + esc(link) + '" target="_blank" rel="noopener">Apri su Normattiva</a>'; return; }

    var att = r.versioni[r.attiva];
    var artRest = CC.articoloRestituito(CC.testoDaHtml(att.html));
    var h = '<div style="display:flex;flex-direction:column;gap:12px">';
    h += '<div><p class="cons-titolo">' + esc(c.testo) + '</p><p class="cons-sotto">' + esc((c.atto.nome || c.atto.urn)) + '</p></div>';
    if (c.genere === 'norma' && artRest && artRest !== c.articolo) {
      h += '<p class="avvertenza">Attenzione: Normattiva restituisce l\'art. ' + esc(artRest) + ' e non l\'art. ' + esc(c.articoloTesto) + '. Il riferimento potrebbe essere errato.</p>';
    }
    if (!r.affidabile) {
      h += '<p class="avvertenza">La versione restituita è in vigore dal ' + esc(formatoData(att.inizioVig)) + ', dopo la data richiesta: l\'articolo non esisteva ancora oppure la versione storica non è disponibile.</p>';
    }
    var righe = riepilogoTesto(c, r, data);
    h += '<div class="riepilogo"><strong>Riepilogo multivigenza</strong><span>' + righe.slice(1, -1).map(esc).join('<br>') + '</span>' +
      '<div class="riga-pulsanti"><button type="button" class="secondario" id="btnCommRiep">Inserisci come commento</button>' +
      '<button type="button" class="secondario" id="btnCopiaRiep">Copia</button></div></div>';
    if (r.versioni.length > 1) {
      h += '<div class="det-blocco"><h3>Versioni dell\'articolo</h3><ul class="versioni">' + r.versioni.map(function (v, i) {
        var oggi = oggiISO();
        var futura = v.inizioVig && v.inizioVig > oggi;
        var corrente = (!v.inizioVig || v.inizioVig <= oggi) && (!v.fineVig || v.fineVig >= oggi);
        var abr = CC.abrogazione(CC.testoDaHtml(v.html));
        var periodo = (v.inizioVig ? 'dal ' + formatoData(v.inizioVig) : 'origine') + (v.fineVig ? ' al ' + formatoData(v.fineVig) : (futura ? '' : ' → oggi'));
        var badges = (i === r.attiva ? '<span class="badge">al ' + esc(formatoData(data)) + '</span>' : '') +
          (abr ? '<span class="badge" style="background:#FCE9E7;color:#B42318">abrogato</span>' : '') +
          (futura ? '<span class="badge" style="background:#E8F0FA;color:#1D4F8C">futura</span>' : (corrente && i !== r.attiva ? '<span class="badge oggi">vigente oggi</span>' : ''));
        return '<li class="' + (i === r.attiva ? 'attiva' : '') + '"><div class="ver-testa"><strong>' + (i + 1) + '</strong><span>' + esc(periodo) + '</span>' + badges + '</div>' +
          '<div class="ver-azioni"><button type="button" data-ver="' + i + '">Testo</button>' + (i > 0 ? '<button type="button" data-diff="' + i + '">Differenze con la ' + i + '</button>' : '') + '</div></li>';
      }).join('') + '</ul></div>';
    }
    var parte = parteDaMostrare(att.html, c);
    h += '<div class="det-blocco"><h3 id="titTesto">Testo vigente al ' + esc(formatoData(data)) + (parte.comma ? ' · comma ' + esc(c.commaTesto) : '') + '</h3><div class="estratto" id="testoVer">' + esc(parte.testo.slice(0, 6000)) + '</div></div>';
    h += '<a href="' + esc(link) + '" target="_blank" rel="noopener">Apri su Normattiva</a>';
    h += '<p class="aiuto">Le versioni sono ricostruite dalle date di vigenza fornite da Normattiva (max 16). Fonte: Normattiva (IPZS), testo non ufficiale.</p></div>';
    out.innerHTML = h;
  }

  function mostraVersione(i, diff) {
    var r = consulta.versioni, c = consulta.cit;
    if (!r || !r.versioni[i]) return;
    var v = r.versioni[i];
    var periodo = (v.inizioVig ? 'dal ' + formatoData(v.inizioVig) : '') + (v.fineVig ? ' al ' + formatoData(v.fineVig) : ' → oggi');
    if (!diff) {
      var p = parteDaMostrare(v.html, c);
      $('titTesto').textContent = 'Versione ' + (i + 1) + ' (' + periodo + ')' + (p.comma ? ' · comma ' + c.commaTesto : '');
      $('testoVer').textContent = p.testo.slice(0, 6000);
    } else {
      var a = parteDaMostrare(r.versioni[i - 1].html, c).testo, b = parteDaMostrare(v.html, c).testo;
      var seg = CC.confrontaParole(a, b);
      $('titTesto').textContent = 'Differenze: versione ' + i + ' → versione ' + (i + 1) + ' (' + periodo + ')';
      $('testoVer').innerHTML = seg ? renderConfronto(seg) : esc('Testo troppo lungo per il confronto: attiva «Mostra solo il comma citato».');
    }
    var tv = $('testoVer'); if (tv.scrollIntoView) tv.scrollIntoView({ block: 'nearest' });
  }

  function inserisciCommentoRiepilogo() {
    var r = consulta.versioni, c = consulta.cit;
    if (!r || !r.trovato) return;
    if (!stato.commenti) { messaggio('Questa versione di Word non consente di inserire commenti dal plugin.'); return; }
    var testo = PREFISSO_COMMENTO + riepilogoTesto(c, r, $('inDataCons').value || oggiISO()).join(' ');
    Word.run(function (ctx) {
      var sel = ctx.document.getSelection();
      sel.load('text');
      return ctx.sync().then(function () {
        if (!(sel.text || '').trim()) { messaggio('Seleziona nel documento il riferimento a cui collegare il commento.'); return; }
        sel.insertComment(testo);
        return ctx.sync().then(function () { messaggio('Riepilogo inserito come commento.'); });
      });
    }).catch(errore);
  }

  function copiaRiepilogo() {
    var r = consulta.versioni, c = consulta.cit;
    if (!r || !r.trovato) return;
    var testo = riepilogoTesto(c, r, $('inDataCons').value || oggiISO()).join('\n');
    var ta = document.createElement('textarea');
    ta.value = testo; ta.setAttribute('readonly', ''); ta.style.position = 'absolute'; ta.style.left = '-9999px';
    document.body.appendChild(ta); ta.select();
    var ok = false; try { ok = document.execCommand('copy'); } catch (e) { ok = false; }
    document.body.removeChild(ta);
    messaggio(ok ? 'Riepilogo copiato.' : 'Copia non riuscita: seleziona il testo manualmente.');
  }

  function consultaDaSelezione(forza) {
    return leggiSelezione().then(function (t) {
      if (!t) { if (forza) messaggio('Seleziona nel documento il riferimento alla norma.'); return; }
      var riconosciuta = CC.estraiCitazioni([{ testo: t }]).some(function (c) { return (c.genere === 'norma' || c.genere === 'atto') && c.atto && !c.atto.errore; });
      if (!$('vistaConsulta').hidden) {
        if (forza || (riconosciuta && t !== consulta.ultima)) avviaConsulta(t);
      } else if (riconosciuta) {
        $('suggTesto').textContent = 'Selezione: «' + t.slice(0, 80) + '»';
        $('btnSugg').setAttribute('data-cit', t);
        $('suggerimento').hidden = false;
      }
    }).catch(function () { /* selezione non leggibile */ });
  }

  function suCambioSelezione() {
    if (consulta.timerSel) clearTimeout(consulta.timerSel);
    consulta.timerSel = setTimeout(function () {
      if (!$('vistaConsulta').hidden && !$('chkSegui').checked) return;
      consultaDaSelezione(false);
    }, 700);
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
      // Le versioni storiche sono disponibili? (la versione restituita deve coprire la data richiesta)
      var urn = 'urn:nir:stato:legge:2014-12-23;190~art1', d = '2022-12-31';
      return chiamaNormattiva(urn + '!vig=' + d).then(function (r) {
        if (!r.trovato) { righe.push('? Versioni storiche: nessun testo per la prova.'); return; }
        var copre = (!r.inizioVig || r.inizioVig <= d) && (!r.fineVig || r.fineVig >= d);
        righe.push((copre ? '✓ Versioni storiche disponibili' : '⚠ Versioni storiche: la data sembra ignorata') +
          ' (L. 190/2014 art. 1 al 31/12/2022 → in vigore ' + (r.inizioVig ? 'dal ' + formatoData(r.inizioVig) : '?') + (r.fineVig ? ' al ' + formatoData(r.fineVig) : ' ad oggi') + ')');
      });
    }).then(function () {
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

    // Consulta norma
    $('inDataCons').value = oggiISO();
    $('tabControllo').addEventListener('click', function () { mostra(stato.citazioni.length && stato.citazioni[0].verifica ? 'vistaRiepilogo' : 'vistaAvvio'); });
    $('tabConsulta').addEventListener('click', function () { mostra('vistaConsulta'); consultaDaSelezione(false); });
    $('btnCerca').addEventListener('click', function () { if ($('inCit').value.trim()) { consulta.ultima = null; avviaConsulta($('inCit').value.trim()); } });
    $('inCit').addEventListener('keydown', function (e) { if (e.key === 'Enter') $('btnCerca').click(); });
    $('inDataCons').addEventListener('change', function () { if (consulta.cit) { stato.cache = stato.cache || {}; avviaConsulta($('inCit').value.trim()); } });
    $('chkSoloComma').addEventListener('change', function () { if (consulta.versioni && consulta.versioni.trovato) renderConsulta(consulta.cit, consulta.versioni, $('inDataCons').value || oggiISO(), consulta.cit.genere === 'atto' ? consulta.cit.atto.urn : CC.urnArticolo(consulta.cit)); });
    $('btnSugg').addEventListener('click', function () { mostra('vistaConsulta'); avviaConsulta($('btnSugg').getAttribute('data-cit')); });
    $('consRisultato').addEventListener('click', function (ev) {
      var b = ev.target.closest ? ev.target.closest('button') : null;
      if (!b) return;
      if (b.id === 'btnCommRiep') inserisciCommentoRiepilogo();
      else if (b.id === 'btnCopiaRiep') copiaRiepilogo();
      else if (b.hasAttribute('data-ver')) mostraVersione(parseInt(b.getAttribute('data-ver'), 10), false);
      else if (b.hasAttribute('data-diff')) mostraVersione(parseInt(b.getAttribute('data-diff'), 10), true);
    });
    try {
      Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, suCambioSelezione);
    } catch (e) { /* evento non disponibile */ }
    if (/[?&]vista=consulta/.test(window.location.search)) { mostra('vistaConsulta'); consultaDaSelezione(true); }
  });
})();
