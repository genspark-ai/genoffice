# PDF: lettura, annotazione e oscuramento

L'editor PDF ha cinque schede nella barra multifunzione: **Home / Annota / Modifica / Pagine / Visualizza**. Legge e scrive: il testo è modificabile, il contenuto può essere oscurato e firmato e i moduli si possono compilare.

## Lettura e navigazione

- Barra laterale sinistra: ** miniature** (fai clic per saltare, con l'intervallo visibile evidenziato) oppure **struttura** (segnalibri, quando presenti).
- Zoom: il controllo del rapporto in basso a destra; ctrl+rotellina cambia lo zoom a scatti.
- Rotazione: per singola pagina o per tutte dal menu Pagine; le rotazioni vengono riscritte al salvataggio.
- Ricerca: ctrl+F a testo completo con tutte le occorrenze evidenziate.
- PDF cifrati: una richiesta della password (in una piccola finestra separata) li apre; la password viene usata solo per questa sessione.

## Selezione del testo e marcatura (Annota)

Prova su un qualsiasi paragrafo:

1. **Trascina il mouse su una frase** — al rilascio sopra il testo compare una barra di annotazione:

![La barra di annotazione dopo aver selezionato il testo](img/pdf-highlight.png)

2. Scegli **evidenziazione** (il campione giallo apre una tavolozza di colori), **sottolineatura** o **barratura**; **Chiedi all'IA** invia la selezione insieme alla tua domanda al pannello IA.
3. Per annullare un'annotazione, riseleziona con il trascinamento lo stesso passaggio e fai clic sul pulsante già attivo sulla barra (un interruttore in stile Word), oppure selezionalo e premi Canc.

Per sapere cosa aspettarti:

- Trascinando sul testo compare una barra a comparsa: **evidenziazione / sottolineatura / barratura / copia / Chiedi all'IA**.
- I colori vengono dalla tavolozza; **applicare la stessa marcatura a un intervallo già marcato la rimuove** (interruttore in stile Word).
- Le marcature già salvate nel file possono essere selezionate ed eliminate (⋯ menu o Canc).
- **Nota**: mentre è attivo uno strumento di disegno, il livello di testo non è selezionabile — lo strumento si disattiva da solo dopo ogni inserimento, così torni in modalità selezione per l'azione successiva.

## Strumenti di disegno (Annota)

Sei strumenti: **inchiostro, rettangolo, ellisse, freccia, nota**, più il **riquadro di oscuramento** nella scheda Modifica.

- Ogni strumento è un interruttore: fai clic per attivarlo; **si disattiva da solo una volta posizionata una forma** (per continuare, fai clic di nuovo sullo strumento); fare clic sullo strumento già attivo lo disattiva comunque.
- L'inchiostro segue lo spessore del tratto; rettangolo, ellisse e freccia si trascinano; i colori vengono dalla tavolozza di disegno.
- Le forme posizionate si possono selezionare, eliminare, trascinare e (rettangolo ed ellisse) ridimensionare.
- **L'oscuramento, il flusso completo** (per nascondere una riga di testo):

  1. Scheda Annota ▸ fai clic su **Oscura area** (lo strumento si attiva).
  2. **Trascina un riquadro sul contenuto** — viene coperto da un segno tratteggiato e nella barra degli strumenti compaiono i pulsanti **cancella contrassegni / applica oscuramenti**:

  ![La pagina dopo aver segnato un'oscuramento](img/pdf-redact.png)

  3. Fai clic su **Applica oscuramenti** e conferma — il risultato è una copia di lavoro in cui il testo e le immagini coperti vengono rimossi fisicamente (non semplicemente coperti) e l'operazione non è annullabile; il documento originale resta intatto.

  Hai sbagliato? Il pulsante cancella contrassegni elimina i segni attuali, così puoi ridisegnare.

## Post-it e fili di commenti

- Lo **strumento nota** posiziona un puntino e apre una scheda a margine per il testo (il nome dell'autore è configurabile); confermando, viene salvata come annotazione di testo PDF standard.
- Fai clic su un puntino per aprire il filo: **rispondi** (fili piatti in stile WPS/Acrobat), **modifica** il tuo commento, **elimina** un commento o l'intero filo.
- Le modifiche in corso restano valide finché il salvataggio non riscrive il nuovo testo nella stessa annotazione nel file, mantenendo intatte le catene di risposte.

## Modifica del contenuto PDF (Modifica)

- **Modifica testo**: fai clic sul testo per modificarlo blocco per blocco (motore pdfium; caratteri abbinati per quanto possibile).
- **Inserisci testo**: inserisci testo ricercabile con il carattere, la dimensione e il colore che preferisci.
- **Inserisci immagine / timbro**.
- Moduli: i campi AcroForm si compilano direttamente; i valori vengono scritti al salvataggio.

## Firme

- **Firma a inchiostro**: disegnala; può essere collegata a un campo firma di un modulo.
- **Firma come immagine**: posiziona un'immagine come firma.
- Le firme salvate si possono riutilizzare.

## Operazioni sulle pagine (Pagine)

- **Ruota / elimina / riordina**: trascina le miniature per riordinarle; l'eliminazione chiede conferma.
- **Estrai pagine**: esporta le pagine selezionate in un nuovo PDF.
- **Dividi**: per intervalli, in più file.
- **Unisci**: aggiunge altri PDF. Le dimensioni sono sommate **prima** che venga letto qualsiasi contenuto e un totale oltre **1 GiB viene rifiutato** con un errore leggibile (per mantenere limitata la memoria).
- Le modifiche a livello di pagina vengono riscritte al salvataggio successivo; Salva con nome… lascia intatto l'originale.

## Esportazione e stampa

- **Esporta come Word…**: conversione locale in .docx.
- **Stampa**: ordine e rotazioni correnti tramite la finestra di sistema; sono supportati gli intervalli di pagine.

## Salvataggio

- Il normale salvataggio, manuale o automatico, riscrive annotazioni e modifiche nel file (in modo atomico).
- **L'oscuramento passa dal proprio flusso di applicazione** e produce una copia, lasciando intatto l'originale, così il contenuto sensibile non resta in esso.
