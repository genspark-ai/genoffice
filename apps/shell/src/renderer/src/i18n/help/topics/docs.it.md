# Docs: elaborazione testi

Docs è il processore di testi simile a Word: legge e scrive veri file .docx con un'impaginazione fedele a WYSIWYG.

## La barra multifunzione

Le schede: **Home / Inserisci / Disegno / Layout / Progettazione / Riferimenti / Revisione / Visualizza**, più le schede contestuali che compaiono per l'oggetto selezionato (Progettazione tabella e layout, Formato immagine, Intestazione e piè di pagina).

- **Home**: appunti; carattere (incluse dimensioni CJK e segni di enfasi) con **Cancella tutta la formattazione** e un interruttore per **Mostra/nascondi segni di formattazione**; paragrafo (allineamento/rientro/interlinea/elenchi) più **Definisci nuovo punto elenco / Definisci nuovo formato numero / Elenco a più livelli**, che salvano nel documento i tuoi propri stili di elenco; stili (Titolo 1-6/Normale/Citazione, modificabili) con un **Riquadro Stili** per l'elenco completo.
- **Inserisci**: interruzioni di pagina e di sezione; tabelle (una griglia righe × colonne, oppure **Inserisci tabella…** per una dimensione esatta); immagini, forme, caselle di testo; **Frontespizio** e **Pagina vuota** da una galleria predefinita; **Grafico**; **Capolettera**; **WordArt**; campi (data, ora, numero di pagina, pagine totali, nome file); collegamenti ipertestuali, **Segnalibro** e riferimenti incrociati; commenti; intestazione e piè di pagina e numeri di pagina; simboli ed equazioni.
  - **Grafico** inserisce un vero oggetto grafico con i propri dati — a barre, a linee o a torta — e non un'immagine. _Modifica dati_ di Word apre i numeri che ci sono dietro.
- **Disegno**: inchiostro sulla pagina, nel gruppo **Strumenti di disegno** — **Seleziona** torna alla modifica del testo, poi **Penna**, **Evidenziatore** e **Gomma** (un clic o un passaggio di mano cancella l'intero tratto). Accanto, **Stile penna** / **Stile evidenziatore** è un solo controllo con campioni di colore e una fila di spessori; la sua etichetta segue lo strumento attivo. L'inchiostro viene salvato nel documento come annotazione che fluttua sopra il testo, quindi sopravvive al salvataggio e alla riapertura, e **Cancella tutto** nel gruppo successivo lo rimuove del tutto.
- **Layout**: margini, orientamento e formato carta, colonne, rientri e spaziature dei paragrafi.
- **Progettazione**: temi, set di colori, filigrana, bordi di pagina.
- **Riferimenti**: indice (aggiornabile), note a piè di pagina e di fine documento, didascalie, riferimenti incrociati.

  ![La scheda Riferimenti](img/docs-references.png)

- **Revisione**: **Editor** controlla l'intero documento per ortografia, grammatica e punteggiatura; **Traduci**; controllo ortografico; commenti (**Commenti con IA** lavora quelli aperti); revisioni con le viste Tutte le modifiche/Revisioni semplici, accetta/rifiuta e **Riepilogo revisioni IA**; conteggio delle parole; **Confronta** con un altro file; **Proteggi documento**.
- **Visualizza**: cinque modi per guardare il file — **Layout di stampa**, **Layout Web**, **Struttura**, **Modalità di lettura** e **Anteprima pagine**; riduci/ingrandisci/100 %/larghezza pagina/una pagina; **Pannello IA**; **Modalità scura**; righello, griglia e riquadro di spostamento; **Nuova scheda**, **Dividi** e **Cambia scheda**; la finestra di dialogo ricercabile delle **scorciatoie da tastiera**.
  - **Modalità scura** scurisce la pagina e la tela intorno, mai il nastro: è la divisione di Word fra una superficie di modifica scura e una finestra scura. La scelta viene ricordata e in ogni caso prevail sull'aspetto dell'app.
  - **Dividi** apre un secondo riquadro sotto che scorre in modo indipendente e rispecchia il primo; chiudilo con la × sul suo bordo.

## Il riquadro di navigazione

**Visualizza ▸ Riquadro di spostamento** apre un pannello laterale che contiene la
struttura dei titoli del documento, una casella di ricerca su tutto il documento e
una miniatura per ogni pagina. Lo stato del pannello viene ricordato tra un avvio e
l'altro, così un documento che consulti tramite la struttura resta navigabile.

**La struttura** è l'albero dei titoli. Fai clic destro su un titolo della struttura
per piegarlo o riorganizzarlo, non solo per spostarti:

- **Comprimi / Espandi** su un titolo comprimono l'intero sottoalbero di quel titolo:
  il capitolo scompare, il suo testo resta nel documento.
- **Comprimi tutto / Espandi tutto** comprimono o espandono
  tutto in un colpo solo. Su un rapporto lungo è la differenza tra una struttura
  leggibile e un muro di testo.
- **Mostra livelli titolo** filtra l'albero alle profondità che ti interessano, così
  _Mostra titolo 1_ ti lascia un indice che puoi davvero scorrere.
- **Alza di livello / Abbassa di livello** cambiano il livello del titolo e, con esso,
  il livello che ogni titolo sottostante eredita: è così che un capitolo diventa una
  sezione.
- **Nuovo titolo prima / dopo** ne inserisce uno nella posizione del cursore, senza
  uscire dal pannello.
- **Elimina** rimuove il titolo _e tutto ciò che c'è sotto_, ed è quello a cui fare
  attenzione: elimina un sottoalbero, non una riga.
- **Seleziona titolo e contenuto** seleziona dal titolo fino alla fine del suo
  sottoalbero, pronto per modificare l'intera sezione.

## Menu contestuale

Fai clic destro ovunque nel corpo del testo — il menu si adatta all'elemento su cui hai cliccato. I gruppi principali:

- **Appunti**: taglia / copia / incolla / **incolla come testo semplice**.
- **Carattere, paragrafo**: cambia famiglia e dimensione, grassetto/corsivo/sottolineato, allineamento/rientro/interlinea senza passare dalla barra multifunzione.
- **Sinonimi**: elenca i sinonimi della parola selezionata; fai clic su uno per sostituirla.
- **Traduci** (IA): traduce la selezione nella lingua di destinazione (inglese, cinese semplificato, giapponese, coreano, francese, tedesco, spagnolo, ...) tramite il pannello IA.
- **Nuovo commento**: allega un commento alla selezione.
- **Ortografia** (su una parola scritta male): sostituzioni suggerite, ignora tutto, aggiungi al dizionario, imposta la lingua di correzione.
- **Collegamento ipertestuale**: apri / modifica / copia collegamento / rimuovi collegamento.
- **Immagine**: visualizza immagine, salva immagine come, **circonda testo** (nella linea / quadrato a sinistra e destra / sopra e sotto / dietro il testo / davanti al testo), ordine di disposizione.
- **Campi** (indice, numeri di pagina): aggiorna campo / mostra/nascondi codici di campo / modifica campo.
- **Numerazione elenco** (all'interno di un elenco): riavvia numerazione / continua numerazione / cambia livello dell'elenco / imposta il valore iniziale della numerazione.
- **Tabella** (cursore dentro una tabella): inserisci righe/colonne, unisci / dividi celle, dividi tabella, adatta automaticamente, allineamento delle celle, distribuisci righe/colonne, proprietà della tabella, menu di eliminazione, seleziona.

## Modifica

- Trova e sostituisci (ctrl+F / ctrl+H): distinzione maiuscole/minuscole, corrispondenza parola intera, espressioni regolari.
- Pennello di formattazione; annulla/ripeti illimitato; opzioni di incolla.
- Tabelle: unire/dividere celle, operazioni su righe e colonne, bordi e sfondi, ordinamento, formule.
- Immagini: circondatura del testo, ritaglio, compressione; tela di disegno.

## Tipografia CJK

- Compressione della punteggiatura e regole di interruzione di linea (kinsoku) allineate a Word; conversione larghezza piena/metà.
- I caratteri candidati coprono i nomi di famiglia CJK più comuni di Windows e macOS.

## IA

- Pulsante IA nella barra multifunzione e pannello laterale: riscrittura, espansione, traduzione, sintesi, inserimento di tabelle preimpostate, più istruzioni libere.
- Ogni intervento dell'IA crea prima uno snapshot; puoi tornare indietro dall'elenco delle versioni e il ripristino stesso è annullabile.

## Salvataggio ed esportazione

- Salva .docx riscrivendo solo i paragrafi modificati — il contenuto non toccato resta identico byte per byte.
- Esporta in PDF (con l'impaginazione corrente) e in immagini pagina per pagina.

## Stampa

ctrl+P tramite la finestra di dialogo di sistema, con pagine WYSIWYG.
