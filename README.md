# Controllo citazioni — componente aggiuntivo per Word (v0.4)

Pannello laterale per Word che legge il documento, trova le citazioni e le controlla:

| Livello | Cosa fa nella v0.1 |
|---|---|
| 1 · Esistenza | Interroga l'API Open Data di **Normattiva**. Verifica che l'articolo restituito sia quello citato e che il comma e la lettera esistano. Segnala le sentenze con data impossibile (es. 30 febbraio). |
| 2 · Vigenza alla data | Confronta il testo del comma **vigente alla data di riferimento** con quello **vigente oggi** e mostra le differenze. |
| 3 · Pertinenza (IA) | Non ancora attiva: richiede un server per custodire la chiave API (prevista nella v0.2). |

Sul documento il plugin aggiunge **riquadri colorati** attorno alle citazioni (content control) e **commenti** con l'esito. Può inserire una **nota di verifica** e un **registro** in fondo al documento. Il testo del documento non viene mai riscritto. Il pulsante "Rimuovi marcature" toglie tutto.

**Novità v0.4 — prassi e giurisprudenza tributaria (DEF, MEF)**
- Riconosce anche: Cassazione in tutte le forme comuni («Cass., sez. V, ord. 4 novembre 2026, n. 28437», «Cass. n. 28437/2026», «Cassazione, ordinanza n. … del …»), giudici tributari di merito («CGT I grado Milano, sent. n. 100/2024», «CTR Lombardia n. …», «CTP …»), circolari, risoluzioni e risposte a interpello («circolare n. 32/E del 5 dicembre 2023», «circ. AdE n. 32/E/2023», «risposta a interpello n. 100/2024»).
- **Senza Worker:** pulsante **«Cerca su DEF»**, che apre la ricerca già compilata sulla Banca dati *Documentazione Economica e Finanziaria* del MEF (def.finanze.it).
- **Con il Worker** (sezione 6): verifica **automatica** su DEF.
  - **Verificata**, con data e oggetto ufficiali e link al documento.
  - **Data diversa**, se numero e anno coincidono ma la data no.
  - **Non verificabile**, se il documento non è in DEF: DEF raccoglie una *selezione* della giurisprudenza tributaria, quindi l'assenza non prova che la sentenza non esista.
  - Per le sentenze: **massima** del Ce.R.D.E.F. nel dettaglio. Con l'IA attiva, la pertinenza della frase viene valutata sulla massima.
- Normattiva viene sempre interrogata direttamente; il Worker serve solo per DEF e IA.

**Novità v0.3**
- Voce del tasto destro più visibile: **«§ Consulta norma (Normattiva)»**, con icona (lente con §).
- Pannello **Consulta norma** riprogettato:
  - scheda con atto, articolo, rubrica e stato (in vigore / abrogato / modifiche future);
  - tre sezioni: **Testo**, **Versioni**, **Riepilogo**;
  - testo diviso in commi numerati, con il comma citato evidenziato e una barra per saltare da un comma all'altro;
  - testo inserito o modificato evidenziato (in Normattiva è tra doppie parentesi);
  - dimensione del carattere regolabile (A− / A+);
  - versioni su una linea del tempo, con le voci «Leggi» e «Cosa è cambiato».
- Codici e testi più vecchi (c.c., D.P.R. 633/1972), dove i commi non sono numerati: ogni capoverso conta come un comma. «art. 2423, comma 2, c.c.» trova il secondo capoverso.
- **Livello 3 · Pertinenza con IA** (facoltativo, sezione 9): l'IA confronta la frase del parere con il testo della norma vigente alla data e segnala le citazioni che non sostengono l'affermazione.

**Novità v0.2**
- **Consulta norma**: seleziona un riferimento nel testo (es. «art. 19 del DPR 633/72»), tasto destro → **Consulta norma**. Il pannello si apre con:
  - il testo vigente alla data scelta (solo il comma citato, oppure l'articolo intero);
  - il **riepilogo multivigenza**: le versioni dell'articolo ricostruite dalle date di vigenza di Normattiva (fino a 10), la versione applicabile alla data, le differenze tra due versioni consecutive;
  - i pulsanti **Inserisci come commento** (aggancia il riepilogo alla selezione) e **Copia**.
  Con il pannello aperto, se cambi selezione la consultazione si aggiorna (casella «Aggiorna quando cambio selezione»).
- **Riconoscimento ampliato**: «artt. 2423 e 2424 c.c.», «art. 2423 e ss. c.c.», «cc» senza punti, «art. 19, D.P.R. 633/1972», date numeriche («26/10/1972»), «n. 472 del 1997», «comma 1 dell'art. 19», e gli atti citati senza articolo («il DPR 633/72», verificati solo come esistenza).
- **Test connessione** più completo: verifica anche se Normattiva restituisce davvero le versioni storiche.

Il menu del tasto destro compare solo sul **testo semplice**: non compare su tabelle, elenchi puntati o caselle di testo. Lì usa la scheda «Consulta norma» del pannello.

Funziona senza server: è una pagina statica pubblicata su **GitHub Pages**, come il plugin Excel.

---

## Contenuto della cartella

```
manifest.xml        ← il file da caricare in Word (va personalizzato, passo 3)
taskpane.html/.css/.js  pannello
core.js             logica di estrazione e verifica (nessuna dipendenza)
commands.html       file richiesto da Word per il pulsante nella barra
assets/             icone
proxy/worker.js     proxy di riserva, solo se serve (passo 6)
test/core.test.js   test della logica (node test/core.test.js)
.nojekyll           evita che GitHub Pages rielabori i file
```

## 1. Crea il repository su GitHub
1. Su github.com: **New repository** → nome `controllo-citazioni` → **Public** → Create.
2. **Add file → Upload files** e trascina **tutto il contenuto** di questa cartella, compresa la cartella `assets`. Conferma con **Commit changes**.
   *(Se usi Git dal Terminale: `git init`, `git add .`, `git commit -m "v0.1"`, `git push`.)*

## 2. Attiva GitHub Pages
**Settings → Pages → Build and deployment → Source: Deploy from a branch → Branch: `main` / `(root)` → Save.**
Dopo 1–2 minuti la pagina è online all'indirizzo `https://TUO-UTENTE.github.io/controllo-citazioni/`.
Prova ad aprire nel browser `…/controllo-citazioni/taskpane.html`: deve comparire il pannello (fuori da Word mostrerà solo l'interfaccia).

## 3. Personalizza il manifest
Nel `manifest.xml` sostituisci **tutte** le occorrenze di `TUO-UTENTE` con il tuo nome utente GitHub. Da Terminale, nella cartella del progetto:

```bash
sed -i '' 's#TUO-UTENTE#il-tuo-utente#g' manifest.xml
```

Se hai dato al repository un nome diverso da `controllo-citazioni`, sostituisci anche quello. Facoltativo: cambia `<ProviderName>` con il nome dello studio. Poi ricarica su GitHub il manifest aggiornato.

Controllo facoltativo del manifest con lo strumento ufficiale Microsoft (serve Node.js): `npx office-addin-manifest validate manifest.xml`

## 4. Carica il plugin in Word per Mac
1. Chiudi Word.
2. Finder → **Cmd+Shift+G** → incolla:
   `~/Library/Containers/com.microsoft.Word/Data/Documents/wef`
   (se la cartella `wef` non esiste, creala).
3. Copia lì dentro il tuo `manifest.xml` personalizzato.
4. Apri Word e un documento → **Home → Componenti aggiuntivi** → **Controllo citazioni**. Nella scheda Home compare anche il pulsante del gruppo "Citazioni".

## 5. Primo utilizzo
1. Icona ⚙︎ in alto a destra nel pannello → inserisci il tuo **nome** → **Salva**.
2. Premi **Test connessione**:
   - `✓ Statuto del contribuente, art. 10: restituito art. 10` → tutto a posto;
   - `⚠ Codice civile…: restituito art. X` → il numero di allegato del codice civile va corretto in `core.js` (voce `regio.decreto:262:1942`, campo `allegato`);
   - `✗ … rete o CORS` → vai al passo 6.
3. Controlla la **data di riferimento**: il plugin la ricava dal testo (es. «periodo d'imposta 2022» → 31/12/2022), ma puoi cambiarla.
4. **Controlla documento** e leggi gli esiti nel pannello. "Dettaglio" mostra il confronto tra le versioni della norma.

## 6. Attivare il Worker (DEF e, se vuoi, IA)
Il browser dentro Word non può leggere i siti di DEF (blocco di sicurezza CORS). Il Worker è un piccolo servizio gratuito su Cloudflare che interroga DEF per conto del plugin. **Un solo Worker** serve sia per DEF sia per l'IA.
1. Crea un account gratuito su cloudflare.com → **Workers & Pages → Create → Worker** → dagli un nome (es. `controllo-citazioni`) → **Deploy**.
2. **Edit code** → cancella il contenuto, incolla tutto `proxy/worker.js` → **Deploy**. In cima al file `ORIGINE` deve essere `https://stefanoscoz.github.io`.
3. Copia l'indirizzo del Worker (`https://controllo-citazioni.<tuo-nome>.workers.dev`).
4. Nel plugin: ⚙︎ → **Indirizzo del Worker** → incolla → **Salva** → **Test connessione**. Deve comparire:
   - `✓ Worker raggiungibile`
   - `✓ DEF: circolare 32/E/2023 trovata`
5. Solo per l'IA: sezione 9 (chiave API come Secret nello stesso Worker).

Il piano gratuito di Cloudflare consente 100.000 richieste al giorno: un controllo ne usa poche decine.

## 7. Aggiornare il plugin
1. Ricarica su GitHub i file modificati e aspetta 1–2 minuti.
2. Se è cambiato il **manifest**, chiudi Word e sostituisci il file nella cartella `wef`.
3. Riapri Word. I file del pannello sono richiamati con un numero di versione (`?v=…` in `taskpane.html`), così Word scarica quelli nuovi senza svuotare la cache.

**Non svuotare la cache di Office** (`OsfWebHost`): rimuove anche gli altri componenti aggiuntivi installati dallo store, come Claude per Word ed Excel, che vanno poi reinstallati.

**Debug** (vedere gli errori): da Terminale
`defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`,
poi tasto destro nel pannello → **Ispeziona elemento**. Funziona solo con Office scaricato da Office.com, non dal Mac App Store.

## 9. Attivare la pertinenza con IA (facoltativo)
Serve una chiave API Claude (console.anthropic.com), che resta nel Worker e non va mai nella pagina pubblica.
1. Usa il Worker della sezione 6: non ne serve un altro.
2. Nel Worker: **Settings → Variables and Secrets → Add** → tipo **Secret**, nome `ANTHROPIC_API_KEY`, valore la tua chiave → **Deploy**. Facoltativo: variabile `MODELLO` per cambiare modello.
3. Nel plugin: ⚙︎ → spunta **Verifica di pertinenza con IA** → **Salva**. Il Test connessione deve indicare «chiave API presente».

Cosa viene inviato: la frase attorno alla citazione, pseudonimizzata (codici fiscali, IBAN, partite IVA, email e telefoni sostituiti), più il testo della norma restituito da Normattiva. I **nomi propri non vengono rimossi**: evita di attivare la funzione su documenti con dati di clienti che non vuoi inviare. L'esito dell'IA è sempre indicato come «da validare» (art. 13 L. 132/2025). Il costo delle chiamate è a tuo carico secondo il listino API.

## 8. Ampliare il dizionario degli atti
In `core.js`, oggetto `ATTI`: aggiungi una riga `'tipo:numero:anno': { data: 'AAAA-MM-GG', nome: '…' }`. Questo serve per le citazioni brevi senza data (es. "D.Lgs. 472/1997"). Tipi: `legge`, `decreto.legislativo`, `decreto.legge`, `decreto.presidente.repubblica`, `regio.decreto`. Le denominazioni abituali (TUIR, c.c., Statuto…) sono nell'elenco `ALIAS`.

Dopo ogni modifica: `node test/core.test.js` (serve Node.js).

---

## Limiti noti della v0.1 — da conoscere prima dell'uso
- **Non collaudato dal vivo su Normattiva.** La logica è testata con risposte simulate. Il formato reale delle risposte e l'accettazione della sintassi `!vig=` nell'API vanno confermati con il Test connessione. Se `!vig=` non viene accettato, il plugin ripiega sul testo attuale e lo dichiara nell'esito.
- **Sentenze:** solo controllo formale delle date. L'esistenza su SentenzeWeb non è ancora verificata (esito "Non verificabile").
- **Prassi** (circolari, risoluzioni, interpelli): sempre "Non verificabile".
- **Abrogazione implicita:** Normattiva non la registra, quindi il plugin non può segnalarla.
- Riconosce le forme di citazione più comuni (art./comma/lett. + atto, alias, "della medesima legge"). Citazioni scritte in modo insolito possono sfuggire: il plugin le mostra come "Atto non riconosciuto".
- Gli esiti sono un **supporto**: la verifica e la responsabilità restano del professionista (art. 13 L. 132/2025; art. 21 Codice deontologico).

**Fonte dei testi:** Normattiva (IPZS). Testi non ufficiali; fa fede la Gazzetta Ufficiale.
