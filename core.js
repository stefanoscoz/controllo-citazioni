/*
 * Controllo citazioni — logica pura (nessuna dipendenza da Word o dal browser).
 * Usabile nel pannello (window.CC) e nei test Node (require('./core.js')).
 *
 * Contiene:
 *  - dizionario degli atti più citati (alias -> URN NIR)
 *  - estrazione delle citazioni normative, giurisprudenziali e di prassi
 *  - costruzione degli URN Normattiva
 *  - controlli formali (date impossibili)
 *  - utilità per confrontare il testo di un comma in due versioni
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.CC = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // ---------------------------------------------------------------------------
  // 1. Dizionario degli atti (estendibile). Chiave = tipo:numero:anno
  //    Le date sono quelle di emanazione usate negli URN di Normattiva.
  // ---------------------------------------------------------------------------
  var ATTI = {
    'decreto.presidente.repubblica:917:1986': { data: '1986-12-22', nome: 'TUIR (D.P.R. 917/1986)' },
    'decreto.presidente.repubblica:633:1972': { data: '1972-10-26', nome: 'Decreto IVA (D.P.R. 633/1972)' },
    'decreto.presidente.repubblica:600:1973': { data: '1973-09-29', nome: 'D.P.R. 600/1973 (accertamento)' },
    'legge:212:2000': { data: '2000-07-27', nome: 'Statuto del contribuente (L. 212/2000)' },
    'legge:190:2014': { data: '2014-12-23', nome: 'L. 190/2014' },
    'legge:132:2025': { data: '2025-09-23', nome: 'L. 132/2025 (intelligenza artificiale)' },
    'decreto.legislativo:546:1992': { data: '1992-12-31', nome: 'D.Lgs. 546/1992 (processo tributario)' },
    'decreto.legislativo:472:1997': { data: '1997-12-18', nome: 'D.Lgs. 472/1997 (sanzioni)' },
    'decreto.legislativo:471:1997': { data: '1997-12-18', nome: 'D.Lgs. 471/1997 (sanzioni)' },
    'decreto.legislativo:241:1997': { data: '1997-07-09', nome: 'D.Lgs. 241/1997' },
    'decreto.legislativo:446:1997': { data: '1997-12-15', nome: 'D.Lgs. 446/1997 (IRAP)' },
    'decreto.legislativo:74:2000': { data: '2000-03-10', nome: 'D.Lgs. 74/2000 (reati tributari)' },
    'decreto.legislativo:231:2001': { data: '2001-06-08', nome: 'D.Lgs. 231/2001' },
    'decreto.legislativo:231:2007': { data: '2007-11-21', nome: 'D.Lgs. 231/2007 (antiriciclaggio)' },
    'decreto.legislativo:39:2010': { data: '2010-01-27', nome: 'D.Lgs. 39/2010 (revisione legale)' },
    'decreto.legislativo:14:2019': { data: '2019-01-12', nome: 'Codice della crisi (D.Lgs. 14/2019)' },
    // Codice civile: allegato del R.D. 262/1942. Il numero di allegato (":2") va collaudato
    // con il "Test connessione": se Normattiva restituisce un altro articolo, correggere qui.
    'regio.decreto:262:1942': { data: '1942-03-16', nome: 'Codice civile', allegato: '2' }
  };

  // Denominazioni abituali -> chiave del dizionario
  var ALIAS = [
    { re: /t\.?\s?u\.?\s?i\.?\s?r\.?|testo unico delle imposte sui redditi/iy, key: 'decreto.presidente.repubblica:917:1986' },
    { re: /statuto (?:dei diritti )?del contribuente/iy, key: 'legge:212:2000' },
    { re: /decreto iva/iy, key: 'decreto.presidente.repubblica:633:1972' },
    { re: /codice della crisi(?: d['’]impresa e dell['’]insolvenza)?|c\.?c\.?i\.?i\.?(?![a-z])/iy, key: 'decreto.legislativo:14:2019' },
    { re: /c\.\s?c\.|cod\.\s?civ\.|codice civile|cc(?![a-zà-ù])/iy, key: 'regio.decreto:262:1942' }
  ];

  var MESI = { gennaio: 1, febbraio: 2, marzo: 3, aprile: 4, maggio: 5, giugno: 6, luglio: 7, agosto: 8, settembre: 9, ottobre: 10, novembre: 11, dicembre: 12 };
  var MESI_RE = '(gennaio|febbraio|marzo|aprile|maggio|giugno|luglio|agosto|settembre|ottobre|novembre|dicembre)';
  var LATINI = '(?:bis|ter|quater|quinquies|sexies|septies|octies|novies|nonies|decies)';

  function pad(n) { return (n < 10 ? '0' : '') + n; }

  function dataValida(g, m, a) {
    if (m < 1 || m > 12 || g < 1) return false;
    var giorni = [31, (a % 4 === 0 && (a % 100 !== 0 || a % 400 === 0)) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    return g <= giorni[m - 1];
  }

  function annoQuattroCifre(a) {
    a = parseInt(a, 10);
    if (a < 100) return a + (a > 40 ? 1900 : 2000);
    return a;
  }

  // "16-bis" / "16 bis" / "16bis" -> "16bis" (forma URN) e "16-bis" (forma leggibile)
  function normNumero(n) {
    if (!n) return null;
    return n.toLowerCase().replace(/[\s-]+/g, '');
  }

  // ---------------------------------------------------------------------------
  // 2. Riconoscimento dell'atto subito dopo "art. N, comma M, lett. x), della ..."
  // ---------------------------------------------------------------------------
  var RE_COMMA = new RegExp('\\s*,?\\s*(?:comma|co\\.|c\\.)\\s*(\\d+(?:[\\s-]?' + LATINI + ')?)', 'iy');
  var RE_LETT = new RegExp('\\s*,?\\s*lett(?:era|\\.)?\\s*([a-z](?:[\\s-]?' + LATINI + ')?)\\)?', 'iy');
  // connettori tra articolo e atto: "del", "della", "dell'", ", ", "ex", "di cui al"
  var RE_CONN = /\s*,?\s*(?:di cui al(?:l['’]|la|lo)?\s*|del(?:la|lo|l['’]|\s)|dell['’]|ex\s)?\s*/iy;
  var TIPI_ATTO = 'd\\.\\s?lgs\\.?|decreto legislativo|d\\.\\s?l\\.|decreto[- ]legge|d\\.\\s?p\\.\\s?r\\.|dpr|decreto del presidente della repubblica|r\\.\\s?d\\.|regio decreto|legge|l\\.';
  var RE_TIPO_TOKEN = new RegExp('(' + TIPI_ATTO + ')', 'iy');
  var RE_DATA_TESTO = new RegExp('\\s*,?\\s*(?:del\\s+)?(\\d{1,2})(?:°|º)?\\s+' + MESI_RE + '\\s+(\\d{4})\\s*,?', 'iy');
  var RE_DATA_NUM = /\s*,?\s*(?:del\s+)?(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})\s*,?/iy;
  var RE_NUM = /\s*(?:(?:n(?:r|um)?\.?|numero)\s*)?(\d+)(?![\d]|[.,]\d)/iy;
  var RE_ANNO = /(?:\s*\/\s*(\d{4}|\d{2})(?!\d)|\s*,?\s*del(?:l['’]\s*anno|\s+anno)?\s+(\d{4})(?!\d))/iy;
  var RE_SS = /\s*,?\s*(?:e\s+)?(?:ss\.|segg\.|seguenti)/iy;
  var RE_ALTRI = new RegExp('\\s*(?:,|e|ed)\\s*(\\d+(?:[\\s-]?' + LATINI + ')?)(?![\\d\\/]|\\.\\d)', 'iy');
  var RE_COMMA_PRIMA = new RegExp('comma\\s+(\\d+(?:[\\s-]?' + LATINI + ')?)\\s*,?\\s*(?:dell[\'’]\\s*|del\\s+)$', 'i');
  var RE_RELATIVO = /(?:medesim[ao]|stess[ao]|citat[ao]|predett[ao])\s+(?:legge|decreto(?:\s+legislativo)?|d\.\s?lgs\.?|testo unico)/iy;

  function tipoUrn(t) {
    t = t.toLowerCase().replace(/\s+/g, '');
    if (/^d\.?lgs|^decretolegislativo/.test(t)) return 'decreto.legislativo';
    if (/^d\.?l\.?$|^decreto-?legge/.test(t)) return 'decreto.legge';
    if (/^d\.?p\.?r|^decretodelpresidente/.test(t)) return 'decreto.presidente.repubblica';
    if (/^r\.?d|^regiodecreto/.test(t)) return 'regio.decreto';
    return 'legge';
  }

  function provaRegex(re, testo, pos) {
    re.lastIndex = pos;
    var m = re.exec(testo);
    return m && m.index === pos ? m : null;
  }

  function leggiAtto(testo, pos, ultimoAtto) {
    var m, i;
    for (i = 0; i < ALIAS.length; i++) {
      m = provaRegex(ALIAS[i].re, testo, pos);
      if (m) return { atto: costruisciAtto(ALIAS[i].key), fine: pos + m[0].length };
    }
    m = provaRegex(RE_RELATIVO, testo, pos);
    if (m) return { atto: ultimoAtto || null, fine: pos + m[0].length, relativo: true };
    return leggiAttoTipizzato(testo, pos);
  }

  // "D.P.R. 26 ottobre 1972, n. 633" | "DPR 633/72" | "D.Lgs. n. 472 del 1997" | "D.P.R. 26/10/1972 n. 633"
  // | "D.P.R. n. 633 del 26 ottobre 1972" | "legge n. 190/2014"
  function leggiAttoTipizzato(testo, pos) {
    var m = provaRegex(RE_TIPO_TOKEN, testo, pos);
    if (!m) return null;
    var tipo = tipoUrn(m[1]);
    var p = pos + m[0].length;
    var g = null, mm = null, anno = null, numero = null;

    function leggiData(at) {
      var d = provaRegex(RE_DATA_TESTO, testo, at);
      if (d) return { g: parseInt(d[1], 10), m: MESI[d[2].toLowerCase()], a: parseInt(d[3], 10), fine: at + d[0].length };
      d = provaRegex(RE_DATA_NUM, testo, at);
      if (d) return { g: parseInt(d[1], 10), m: parseInt(d[2], 10), a: parseInt(d[3], 10), fine: at + d[0].length };
      return null;
    }

    var d1 = leggiData(p);
    if (d1) { g = d1.g; mm = d1.m; anno = d1.a; p = d1.fine; }
    var n = provaRegex(RE_NUM, testo, p);
    if (!n) return null;                       // "legge di bilancio": nessun numero
    numero = parseInt(n[1], 10); p += n[0].length;
    if (!d1) {
      var a = provaRegex(RE_ANNO, testo, p);
      if (a) { anno = a[1] ? annoQuattroCifre(a[1]) : parseInt(a[2], 10); p += a[0].length; }
      else {
        var d2 = leggiData(p);
        if (d2) { g = d2.g; mm = d2.m; anno = d2.a; p = d2.fine; }
      }
    }
    if (!anno) return null;                    // "L. n. 3" senza anno: non affidabile
    var data = null;
    if (g && mm) {
      if (!dataValida(g, mm, anno)) return { atto: { errore: 'Data dell\'atto inesistente' }, fine: p };
      data = anno + '-' + pad(mm) + '-' + pad(g);
    }
    return { atto: costruisciAtto(tipo + ':' + numero + ':' + anno, data, tipo, numero, anno), fine: p };
  }

  function costruisciAtto(key, dataTesto, tipo, numero, anno) {
    var noto = ATTI[key];
    var parti = key.split(':');
    tipo = tipo || parti[0]; numero = numero || parseInt(parti[1], 10); anno = anno || parseInt(parti[2], 10);
    var data = dataTesto || (noto && noto.data) || null;
    // Normattiva usa "decreto.del.presidente.della.repubblica" negli URN dei D.P.R. (verificato sull'API)
    var tipoUrnNir = tipo === 'decreto.presidente.repubblica' ? 'decreto.del.presidente.della.repubblica' : tipo;
    var urn = 'urn:nir:stato:' + tipoUrnNir + ':' + (data || anno) + ';' + numero + (noto && noto.allegato ? ':' + noto.allegato : '');
    var avviso = null;
    if (dataTesto && noto && noto.data !== dataTesto) avviso = 'La data indicata non coincide con quella dell\'atto in archivio (' + noto.data.split('-').reverse().join('/') + ')';
    return {
      key: key, tipo: tipo, numero: numero, anno: anno, data: data,
      soloAnno: !data, urn: urn, nome: noto ? noto.nome : null, avviso: avviso
    };
  }

  // ---------------------------------------------------------------------------
  // 3. Estrazione
  // ---------------------------------------------------------------------------
  var RE_ART = new RegExp('\\b(art(?:icol[oi]|t)?)\\.?\\s*(\\d+(?:[\\s-]?' + LATINI + ')?)(?![\\d/])', 'gi');
  var RE_ATTO_SOLO = new RegExp('(?:^|[^a-zà-ù.])(' + TIPI_ATTO + ')', 'gi');
  var RE_CASS = new RegExp(
    '\\bCass(?:azione)?\\.?(?:\\s*(?:civ|pen|trib)\\.?)?\\s*,?' +
    '(?:\\s*(?:sez(?:ione)?\\.?|ss\\.?\\s*uu\\.?|sezioni unite)\\s*(?:un(?:ite)?\\.?|[IVX]+|\\d+|trib(?:utaria)?\\.?|lav(?:oro)?\\.?)?\\s*,?)?' +
    '(?:\\s*(?:sent(?:enza)?|ord(?:inanza)?)\\.?\\s*,?)?' +
    '(?:\\s*(\\d{1,2})\\s+' + MESI_RE + '\\s+(\\d{4})\\s*,?)?' +
    '\\s*n\\.\\s*(\\d+)(?:\\s*\\/\\s*(\\d{4}))?', 'gi');
  var RE_PRASSI = new RegExp(
    '\\b(circ(?:olare)?|ris(?:oluzione)?|risposta(?: a interpello)?)\\.?' +
    '(?:\\s*(?:dell[\'’]\\s*)?(?:Agenzia delle Entrate|AdE|Ag\\.\\s*Entrate))?' +
    '\\s*n\\.\\s*(\\d+)(?:\\s*\\/\\s*E)?' +
    '(?:\\s*(?:del|\\/)\\s*(?:(\\d{1,2})\\s+' + MESI_RE + '\\s+)?(\\d{4}))?', 'gi');

  /**
   * Estrae le citazioni da un elenco di paragrafi.
   * @param {Array<{testo:string, sezione?:string}>} paragrafi
   * @returns {Array} citazioni con posizione (indice paragrafo, offset, testo)
   */
  function estraiCitazioni(paragrafi) {
    var out = [];
    var ultimoAtto = null;
    paragrafi.forEach(function (p, ip) {
      var testo = p.testo || '';
      var trovate = [];

      // 3a. Norme
      RE_ART.lastIndex = 0;
      var m;
      while ((m = RE_ART.exec(testo)) !== null) {
        var inizio = m.index;
        var pos = m.index + m[0].length;
        var plurale = /^(artt|articoli)$/i.test(m[1]);
        var c = { genere: 'norma', articolo: normNumero(m[2]), articoloTesto: m[2] };
        // "comma 1 dell'art. 19"
        var prima = RE_COMMA_PRIMA.exec(testo.slice(Math.max(0, inizio - 40), inizio));
        if (prima) { c.comma = normNumero(prima[1]); c.commaTesto = prima[1]; inizio = inizio - prima[0].length; }
        // "artt. 2423, 2424 e 2425"
        if (plurale) {
          var ma; c.altriArticoli = [];
          while ((ma = provaRegex(RE_ALTRI, testo, pos))) { c.altriArticoli.push({ articolo: normNumero(ma[1]), articoloTesto: ma[1] }); pos += ma[0].length; }
          if (!c.altriArticoli.length) delete c.altriArticoli;
        }
        var mss = provaRegex(RE_SS, testo, pos);
        if (mss) { c.seguenti = true; pos += mss[0].length; }
        if (!c.comma) {
          var mc = provaRegex(RE_COMMA, testo, pos);
          if (mc) { c.comma = normNumero(mc[1]); c.commaTesto = mc[1]; pos += mc[0].length; }
        }
        var ml = provaRegex(RE_LETT, testo, pos);
        if (ml) { c.lettera = ml[1].toLowerCase().replace(/[\s-]+/g, ''); pos += ml[0].length; }
        if (!mss) { mss = provaRegex(RE_SS, testo, pos); if (mss) { c.seguenti = true; pos += mss[0].length; } }
        var mconn = provaRegex(RE_CONN, testo, pos);
        var posAtto = mconn ? pos + mconn[0].length : pos;
        var la = leggiAtto(testo, posAtto, ultimoAtto);
        if (la) {
          c.atto = la.atto; c.relativo = !!la.relativo; pos = la.fine;
          if (la.atto && !la.atto.errore) ultimoAtto = la.atto;
        } else {
          c.atto = null;
        }
        // togli punteggiatura finale dalla porzione citata
        var porzione = testo.slice(inizio, pos).replace(/[\s,;:]+$/, '');
        c.testo = porzione; c.inizio = inizio; c.fine = inizio + porzione.length;
        trovate.push(c);
        RE_ART.lastIndex = Math.max(pos, RE_ART.lastIndex);
      }

      // 3a-bis. Atti citati senza articolo ("il DPR 633/72", "la L. 190/2014")
      RE_ATTO_SOLO.lastIndex = 0;
      while ((m = RE_ATTO_SOLO.exec(testo)) !== null) {
        var st = m.index + m[0].length - m[1].length;
        var dentro = trovate.some(function (t) { return st >= t.inizio && st < t.fine; });
        if (dentro) continue;
        var lt = leggiAttoTipizzato(testo, st);
        if (!lt || !lt.atto) continue;
        var ta = testo.slice(st, lt.fine).replace(/[\s,;:]+$/, '');
        trovate.push({ genere: 'atto', atto: lt.atto, testo: ta, inizio: st, fine: st + ta.length });
        if (!lt.atto.errore) ultimoAtto = lt.atto;
        RE_ATTO_SOLO.lastIndex = Math.max(lt.fine, RE_ATTO_SOLO.lastIndex);
      }

      // 3b. Cassazione
      RE_CASS.lastIndex = 0;
      while ((m = RE_CASS.exec(testo)) !== null) {
        var cs = { genere: 'sentenza', testo: m[0].replace(/[\s,;:]+$/, ''), inizio: m.index };
        cs.fine = cs.inizio + cs.testo.length;
        cs.numero = parseInt(m[4], 10);
        if (m[1] && m[2] && m[3]) {
          cs.giorno = parseInt(m[1], 10); cs.mese = MESI[m[2].toLowerCase()]; cs.anno = parseInt(m[3], 10);
          cs.dataValida = dataValida(cs.giorno, cs.mese, cs.anno);
          if (m[5] && parseInt(m[5], 10) !== cs.anno) cs.avviso = 'Anno del numero diverso dall\'anno della data';
        } else if (m[5]) {
          cs.anno = parseInt(m[5], 10); cs.dataValida = null;
        }
        trovate.push(cs);
      }

      // 3c. Prassi
      RE_PRASSI.lastIndex = 0;
      while ((m = RE_PRASSI.exec(testo)) !== null) {
        var cp = { genere: 'prassi', testo: m[0].replace(/[\s,;:]+$/, ''), inizio: m.index };
        cp.fine = cp.inizio + cp.testo.length;
        cp.tipo = m[1].toLowerCase().indexOf('circ') === 0 ? 'circolare' : (m[1].toLowerCase().indexOf('ris') === 0 && m[1].toLowerCase().indexOf('risposta') !== 0 ? 'risoluzione' : 'risposta a interpello');
        cp.numero = parseInt(m[2], 10);
        if (m[3] && m[4]) cp.dataValida = dataValida(parseInt(m[3], 10), MESI[m[4].toLowerCase()], parseInt(m[5], 10));
        trovate.push(cp);
      }

      trovate.sort(function (a, b) { return a.inizio - b.inizio; });
      trovate.forEach(function (c) {
        c.paragrafo = ip; c.sezione = p.sezione || null;
        out.push(c);
      });
    });
    out.forEach(function (c, i) { c.id = i + 1; });
    return out;
  }

  // ---------------------------------------------------------------------------
  // 4. Esiti formali (prima di interrogare Normattiva)
  // ---------------------------------------------------------------------------
  var ESITI = {
    ok: { etichetta: 'Verificata', gravita: 4 },
    mod: { etichetta: 'Modificata', gravita: 2 },
    dub: { etichetta: 'Pertinenza dubbia', gravita: 1 },
    ko: { etichetta: 'Non trovata', gravita: 0 },
    nv: { etichetta: 'Non verificabile', gravita: 3 }
  };

  function esitoFormale(c) {
    if (c.genere === 'sentenza') {
      if (c.dataValida === false) return { esito: 'ko', bloccante: true, motivo: 'La data indicata non esiste: gli estremi non sono validi. Possibile citazione generata da IA.' };
      return { esito: 'nv', motivo: 'Giurisprudenza non ancora collegata in questa versione: verifica su SentenzeWeb.' + (c.avviso ? ' ' + c.avviso + '.' : '') };
    }
    if (c.genere === 'prassi') {
      if (c.dataValida === false) return { esito: 'ko', bloccante: true, motivo: 'La data del documento di prassi non esiste.' };
      return { esito: 'nv', motivo: 'Fonte di prassi non collegata: verifica manuale.' };
    }
    if (c.genere === 'atto' && c.atto && c.atto.errore) return { esito: 'ko', bloccante: true, motivo: c.atto.errore + '.' };
    if (!c.atto) return { esito: 'nv', motivo: 'Atto non riconosciuto: aggiungilo al dizionario o scrivi gli estremi completi.' };
    if (c.atto.errore) return { esito: 'ko', bloccante: true, motivo: c.atto.errore + '.' };
    return null; // serve la verifica su Normattiva
  }

  function urnArticolo(c) {
    if (!c.atto || !c.atto.urn) return null;
    return c.atto.urn + '~art' + c.articolo;
  }

  function linkNormattiva(urn, dataVig) {
    return 'https://www.normattiva.it/uri-res/N2Ls?' + urn + (dataVig ? '!vig=' + dataVig : '');
  }

  // ---------------------------------------------------------------------------
  // 5. Analisi del testo restituito da Normattiva
  // ---------------------------------------------------------------------------
  var ENTITA = { nbsp: ' ', amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", egrave: 'è', eacute: 'é', agrave: 'à',
    aacute: 'á', ograve: 'ò', oacute: 'ó', ugrave: 'ù', uacute: 'ú', igrave: 'ì', iacute: 'í', Egrave: 'È', Eacute: 'É',
    Agrave: 'À', Ograve: 'Ò', Ugrave: 'Ù', Igrave: 'Ì', laquo: '«', raquo: '»', ldquo: '“', rdquo: '”', lsquo: '‘',
    rsquo: '’', ndash: '–', mdash: '—', deg: '°', ordm: 'º', ordf: 'ª', euro: '€', hellip: '…', middot: '·', sect: '§' };

  function testoDaHtml(html) {
    return String(html || '')
      .replace(/[\r\n]+/g, ' ')                               // gli a capo del sorgente HTML non contano
      .replace(/<(script|style)[\s\S]*?<\/\1>/gi, ' ')
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<\/(p|div|li|h\d)>/gi, '\n')
      .replace(/<(p|div|li|h\d)(\s[^>]*)?>/gi, '\n')
      .replace(/<[^>]+>/g, '')
      .replace(/&#(\d+);/g, function (_, n) { return String.fromCharCode(parseInt(n, 10)); })
      .replace(/&#x([0-9a-f]+);/gi, function (_, n) { return String.fromCharCode(parseInt(n, 16)); })
      .replace(/&([a-z]+);/gi, function (m, n) { return ENTITA[n] !== undefined ? ENTITA[n] : m; })
      .replace(/[ \t\f\v ]+/g, ' ')
      .replace(/ *\n[\s]*/g, '\n')
      .trim();
  }

  var RE_COMMA_RIGA = new RegExp('^(\\d+(?:[\\s-]?' + LATINI + ')?)\\.\\s+', 'i');
  var RE_LETTERA_RIGA = new RegExp('^((?:[a-z]{1,2}|\\d+)(?:[\\s-]?' + LATINI + ')?(?:\\.\\d+)?)\\)\\s+', 'i');

  /**
   * Struttura il testo di un articolo per la lettura:
   * { titolo, rubrica, preambolo, commi: [{ numero, implicito, righe: [{ lettera|null, testo }] }] }
   * Nei codici e nei testi più vecchi (c.c., D.P.R. 633/1972) i commi non sono numerati: ogni capoverso è un comma.
   * Normattiva spezza le righe dove c'è testo modificato "((...))": quelle righe vengono ricucite.
   */
  function strutturaNorma(testo) {
    var righe = String(testo || '').split('\n').map(function (r) { return r.trim(); }).filter(Boolean);
    var out = { titolo: null, rubrica: null, preambolo: [], commi: [] };
    if (righe.length && /^Art(?:icolo)?\.?\s*\d+/i.test(righe[0]) && righe[0].length < 40) out.titolo = righe.shift().replace(/\.$/, '');
    if (righe.length && righe[0].length < 160 && !RE_COMMA_RIGA.test(righe[0]) &&
        (/^\(?\(?\s*\(.*\)\s*\)?\)?\.?$/.test(righe[0]) || (righe.length > 1 && !/[.;:]$/.test(righe[0]) && /^[A-ZÀ-Ù(]/.test(righe[0])))) {
      out.rubrica = righe.shift().replace(/^\(\(\s*|\s*\)\)$/g, '').replace(/\.$/, '').replace(/^\((.*)\)$/, '$1').trim();
    }
    // ricuci le righe spezzate dalle modifiche: continua se la riga precedente non chiude la frase
    // o se la riga inizia in minuscolo / con punteggiatura
    var cucite = [];
    righe.forEach(function (r) {
      var prec = cucite[cucite.length - 1];
      var continua = prec && !RE_COMMA_RIGA.test(r) && !RE_LETTERA_RIGA.test(r) &&
        (/^(\(\()?\s*[,;:.)a-zà-ù]/.test(r) || !/[.;:]\s*(\)\))?$/.test(prec));
      if (continua) cucite[cucite.length - 1] = prec + (/^(\(\()?\s*[,;:.)]/.test(r) ? '' : ' ') + r;
      else cucite.push(r);
    });
    var numerati = cucite.some(function (r) { return /^1\.\s/.test(r); });
    var cur = null, n = 0;
    cucite.forEach(function (r) {
      var mc = numerati ? RE_COMMA_RIGA.exec(r) : null, ml = RE_LETTERA_RIGA.exec(r);
      if (mc) { cur = { numero: mc[1].replace(/\s+/g, '-'), implicito: false, righe: [{ lettera: null, testo: r.slice(mc[0].length) }] }; out.commi.push(cur); }
      else if (ml && cur) cur.righe.push({ lettera: ml[1], testo: r.slice(ml[0].length) });
      else if (!numerati) { n++; cur = { numero: String(n), implicito: true, righe: [{ lettera: null, testo: r }] }; out.commi.push(cur); }
      else if (cur) cur.righe.push({ lettera: null, testo: r });
      else out.preambolo.push(r);
    });
    return out;
  }

  /** Pseudonimizza dati personali ricorrenti prima di inviare testo a un servizio esterno. */
  function pseudonimizza(t) {
    return String(t || '')
      .replace(/\b[A-Z]{6}\d{2}[A-EHLMPR-T]\d{2}[A-Z]\d{3}[A-Z]\b/gi, '[CODICE FISCALE]')
      .replace(/\bIT\d{2}[A-Z]\d{10}[0-9A-Z]{12}\b/gi, '[IBAN]')
      .replace(/\b(?:P\.?\s?IVA|partita IVA)\s*:?\s*\d{11}\b/gi, '[PARTITA IVA]')
      .replace(/\b\d{11}\b/g, '[NUMERO]')
      .replace(/[\w.+-]+@[\w-]+\.[\w.]+/g, '[EMAIL]')
      .replace(/(?:\+39\s?)?\b3\d{2}[\s.]?\d{6,7}\b/g, '[TELEFONO]');
  }

  /** "((ARTICOLO ABROGATO DAL D.LGS. 19 GENNAIO 2026, N. 10))" -> { cosa: 'articolo', da: 'D.LGS. 19 GENNAIO 2026, N. 10' } */
  function abrogazione(testo) {
    var m = /\(\(\s*(ARTICOLO|PROVVEDIMENTO|COMMA|CAPO|TITOLO)\s+ABROGAT[OA]\s+(?:DAL(?:LA|L['’])?|DA)\s*([^)]*?)\s*\)\)/i.exec(testo || '');
    if (!m) return null;
    var corpo = String(testo).replace(m[0], '').replace(/Art(?:icolo)?\.?\s*\d+[a-z-]*\.?/i, '').trim();
    if (corpo.length > 400) return null; // l'abrogazione riguarda solo una parte del testo
    return { cosa: m[1].toLowerCase(), da: m[2].replace(/\s+,/g, ',').trim() };
  }

  /** Numero dell'articolo letto nell'intestazione del testo ("Art. 16-bis") */
  function articoloRestituito(testo) {
    var m = new RegExp('Art(?:icolo)?\\.?\\s*(\\d+(?:[\\s-]?' + LATINI + ')?)', 'i').exec(testo || '');
    return m ? normNumero(m[1]) : null;
  }

  var ORDINE_LATINI = ['bis', 'ter', 'quater', 'quinquies', 'sexies', 'septies', 'octies', 'novies', 'decies'];

  /** Estrae il testo di un comma ("54", "54bis") dal testo dell'articolo. */
  function estraiComma(testo, comma) {
    if (!testo || !comma) return null;
    var mm = /^(\d+)([a-z]*)$/.exec(comma);
    if (!mm) return null;
    var n = parseInt(mm[1], 10), suff = mm[2];
    var etichetta = suff ? n + '[\\s-]?' + suff : String(n);
    var reInizio = new RegExp('(?:^|\\n|\\s)(' + etichetta + ')\\.\\s', 'i');
    var mi = reInizio.exec(testo);
    if (!mi) {
      // codici con commi non numerati: il comma N è l'N-esimo capoverso
      var st = strutturaNorma(testo);
      if (!suff && st.commi.length && st.commi[0].implicito && st.commi[n - 1]) {
        return st.commi[n - 1].righe.map(function (r) { return (r.lettera ? r.lettera + ') ' : '') + r.testo; }).join('\n');
      }
      return null;
    }
    var start = mi.index + mi[0].length - mi[0].replace(/^\s+/, '').length;
    var resto = testo.slice(mi.index + mi[0].length);
    // possibili commi successivi: n-(latino successivo) oppure n+1
    var candidati = [];
    var idx = suff ? ORDINE_LATINI.indexOf(suff) + 1 : 0;
    for (var k = idx; k < ORDINE_LATINI.length; k++) candidati.push(n + '[\\s-]?' + ORDINE_LATINI[k]);
    candidati.push(String(n + 1));
    var reFine = new RegExp('(?:\\n|\\s)(?:' + candidati.join('|') + ')\\.\\s', 'i');
    var mf = reFine.exec(resto);
    var fine = mf ? mi.index + mi[0].length + mf.index : testo.length;
    return testo.slice(start, fine).trim();
  }

  function contieneLettera(testo, lettera) {
    if (!testo || !lettera) return null;
    var l = lettera.replace(/^([a-z])(.*)$/, function (_, a, b) { return a + (b ? '[\\s-]?' + b : ''); });
    return new RegExp('(?:^|\\s|\\n)' + l + '\\)\\s', 'i').test(testo);
  }

  function normalizzaPerConfronto(t) {
    return String(t || '').toLowerCase().replace(/\(\(|\)\)/g, '').replace(/\s+/g, ' ').trim();
  }

  /** Confronto per parole (LCS) — restituisce segmenti {tipo:'=','-','+', testo} */
  function confrontaParole(a, b, limite) {
    limite = limite || 1500;
    var A = String(a || '').split(/\s+/).filter(Boolean), B = String(b || '').split(/\s+/).filter(Boolean);
    if (A.length > limite || B.length > limite) return null;
    var n = A.length, m = B.length, i, j;
    var dp = new Array(n + 1);
    for (i = 0; i <= n; i++) { dp[i] = new Uint16Array(m + 1); }
    for (i = n - 1; i >= 0; i--) for (j = m - 1; j >= 0; j--)
      dp[i][j] = A[i] === B[j] ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
    var out = [];
    function push(tipo, w) {
      var u = out[out.length - 1];
      if (u && u.tipo === tipo) u.testo += ' ' + w; else out.push({ tipo: tipo, testo: w });
    }
    i = 0; j = 0;
    while (i < n && j < m) {
      if (A[i] === B[j]) { push('=', A[i]); i++; j++; }
      else if (dp[i + 1][j] >= dp[i][j + 1]) { push('-', A[i]); i++; }
      else { push('+', B[j]); j++; }
    }
    while (i < n) push('-', A[i++]);
    while (j < m) push('+', B[j++]);
    return out;
  }

  /** "AAAAMMGG" -> "AAAA-MM-GG" (99999999 = vigenza illimitata -> null) */
  function dataApi(s) {
    s = String(s || '');
    if (!/^\d{8}$/.test(s) || s === '99999999') return null;
    return s.slice(0, 4) + '-' + s.slice(4, 6) + '-' + s.slice(6, 8);
  }

  /** Data di riferimento suggerita dal testo ("periodo d'imposta 2022" -> 2022-12-31) */
  function suggerisciDataRiferimento(testoDocumento, oggi) {
    var m = /periodo d['’]\s?imposta\s+(\d{4})/i.exec(testoDocumento || '') ||
            /esercizio\s+(\d{4})/i.exec(testoDocumento || '') ||
            /anno d['’]\s?imposta\s+(\d{4})/i.exec(testoDocumento || '');
    if (m) return { data: m[1] + '-12-31', fonte: '«' + m[0] + '»' };
    return { data: oggi, fonte: 'data odierna' };
  }

  /**
   * Valuta l'esito di una citazione normativa a partire dalle due risposte di Normattiva
   * (testo vigente alla data di riferimento e testo vigente oggi).
   * Le risposte sono già normalizzate: {trovato, multiplo, html, inizioVig, fineVig, errore}
   */
  function it(d) { return d && /^\d{4}-\d{2}-\d{2}$/.test(d) ? d.slice(8, 10) + '/' + d.slice(5, 7) + '/' + d.slice(0, 4) : d; }

  function valutaNorma(c, rif, oggi, dataRif) {
    var urn = urnArticolo(c);
    var base = { urn: urn, link: linkNormattiva(urn, dataRif) };
    function r(o) { for (var k in base) if (!(k in o)) o[k] = base[k]; return o; }

    if (rif.errore) return r({ esito: 'nv', motivo: 'Normattiva non raggiungibile dal plugin (' + rif.errore + '). Apri il link per verificare.' });
    if (rif.multiplo) return r({ esito: 'nv', motivo: 'Normattiva restituisce più atti per questi estremi: indica la data completa dell\'atto.' });
    if (!rif.trovato && rif.nonAncoraVigente) {
      return r({ esito: 'ko', bloccante: true, motivo: 'Alla data di riferimento (' + it(dataRif) + ') l\'articolo non esisteva ancora o non era vigente (Normattiva).' });
    }
    if (!rif.trovato) {
      if (c.atto.soloAnno) return r({ esito: 'nv', motivo: 'Atto non individuato con il solo anno: indica la data completa.' });
      return r({ esito: 'ko', bloccante: true, motivo: 'Normattiva non restituisce l\'atto o l\'articolo indicato.' });
    }

    var testoRif = testoDaHtml(rif.html);
    var artRest = articoloRestituito(testoRif);
    if (artRest && artRest !== c.articolo) {
      return r({ esito: 'ko', bloccante: true, motivo: 'Articolo non trovato: Normattiva restituisce l\'art. ' + artRest + ' al posto dell\'art. ' + c.articoloTesto + '.' });
    }
    if (rif.inizioVig && rif.inizioVig > dataRif) {
      return r({ esito: 'nv', motivo: 'Normattiva restituisce una versione in vigore dal ' + it(rif.inizioVig) + ', successiva alla data di riferimento (' + it(dataRif) + '): l\'articolo non esisteva ancora oppure la versione storica non è disponibile. Verifica sul link.' });
    }

    var abrRif = abrogazione(testoRif);
    if (abrRif) {
      return r({ esito: 'ko', bloccante: true, abrogato: true, motivo: 'Alla data di riferimento (' + it(dataRif) + ') l\'' + abrRif.cosa + ' risulta abrogato' + (abrRif.da ? ' (' + abrRif.da + ')' : '') + (rif.inizioVig ? ', dal ' + it(rif.inizioVig) : '') + '.' });
    }
    var testoOggiIntero = oggi && !oggi.errore && oggi.trovato ? testoDaHtml(oggi.html) : null;
    var abrOggi = abrogazione(testoOggiIntero);
    if (abrOggi) {
      return r({ esito: 'mod', abrogato: true, motivo: 'Vigente al ' + it(dataRif) + ', ma oggi l\'' + abrOggi.cosa + ' risulta abrogato' + (abrOggi.da ? ' (' + abrOggi.da + ')' : '') + (oggi.inizioVig ? ', dal ' + it(oggi.inizioVig) : '') + '. Verifica quale testo si applica.', estrattoRif: c.comma ? (estraiComma(testoRif, c.comma) || testoRif) : testoRif });
    }
    var parteRif = testoRif, parteOggi = testoOggiIntero;
    var note = [];
    if (c.comma) {
      var cr = estraiComma(testoRif, c.comma);
      if (!cr) return r({ esito: 'nv', motivo: 'Comma ' + c.commaTesto + ' non individuato automaticamente nel testo vigente al ' + it(dataRif) + ': verifica manuale.', estrattoRif: testoRif.slice(0, 600) });
      parteRif = cr;
      if (parteOggi) parteOggi = estraiComma(parteOggi, c.comma);
      if (c.lettera && contieneLettera(cr, c.lettera) === false) {
        return r({ esito: 'dub', motivo: 'Nel comma ' + c.commaTesto + ' non risulta la lettera ' + c.lettera + '): controlla il riferimento.', estrattoRif: cr });
      }
    }
    if (c.atto.avviso) note.push(c.atto.avviso);
    if (!artRest) note.push('Numero di articolo non leggibile nel testo restituito: controlla il link');

    if (parteOggi !== null && normalizzaPerConfronto(parteRif) !== normalizzaPerConfronto(parteOggi)) {
      return r({
        esito: 'mod',
        motivo: 'Vigente al ' + it(dataRif) + ', ma il testo è cambiato successivamente' + (rif.fineVig ? ' (versione valida fino al ' + it(rif.fineVig) + ')' : '') + '. Verifica quale versione si applica.' + (note.length ? ' ' + note.join('. ') + '.' : ''),
        estrattoRif: parteRif, estrattoOggi: parteOggi, confronto: confrontaParole(parteRif, parteOggi)
      });
    }
    return r({
      esito: 'ok',
      motivo: 'Esiste ed è vigente alla data di riferimento. Pertinenza non valutata in questa versione.' + (note.length ? ' ' + note.join('. ') + '.' : ''),
      estrattoRif: parteRif
    });
  }

  /** Atto citato senza articolo: basta che Normattiva lo trovi. */
  function valutaAtto(c, r) {
    var link = linkNormattiva(c.atto.urn, null);
    if (r.errore) return { esito: 'nv', urn: c.atto.urn, link: link, motivo: 'Normattiva non raggiungibile dal plugin (' + r.errore + ').' };
    if (r.multiplo) return { esito: 'nv', urn: c.atto.urn, link: link, motivo: 'Più atti corrispondono a questi estremi: indica la data completa.' };
    if (!r.trovato) return c.atto.soloAnno
      ? { esito: 'nv', urn: c.atto.urn, link: link, motivo: 'Atto non individuato con il solo anno: indica la data completa.' }
      : { esito: 'ko', bloccante: true, urn: c.atto.urn, link: link, motivo: 'Normattiva non restituisce l\'atto indicato.' };
    return { esito: 'ok', urn: c.atto.urn, link: link, motivo: 'L\'atto esiste in Normattiva' + (c.atto.nome ? ' (' + c.atto.nome + ')' : '') + '. Citazione senza articolo: nessun controllo di vigenza.' + (c.atto.avviso ? ' ' + c.atto.avviso + '.' : '') };
  }

  return {
    ATTI: ATTI, valutaAtto: valutaAtto, leggiAttoTipizzato: leggiAttoTipizzato, ESITI: ESITI,
    estraiCitazioni: estraiCitazioni, esitoFormale: esitoFormale, valutaNorma: valutaNorma,
    urnArticolo: urnArticolo, linkNormattiva: linkNormattiva,
    testoDaHtml: testoDaHtml, strutturaNorma: strutturaNorma, pseudonimizza: pseudonimizza, articoloRestituito: articoloRestituito, abrogazione: abrogazione, estraiComma: estraiComma,
    contieneLettera: contieneLettera, confrontaParole: confrontaParole, dataApi: dataApi,
    suggerisciDataRiferimento: suggerisciDataRiferimento, dataValida: dataValida
  };
});
