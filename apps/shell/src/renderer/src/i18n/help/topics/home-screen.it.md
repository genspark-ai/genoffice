# La schermata Home: dove si trovano i tuoi file

Home è la pagina iniziale di GenOffice: una barra di navigazione a sinistra, elenchi di file e schede di creazione rapida a destra.

![La schermata Home](img/home-screen.png)

## Navigazione laterale

- **Recenti**: i file che hai aperto di recente. Ogni riga è datata — oggi, ieri o la data.
- **Preferiti**: i file che hai messo tra i preferiti. Passa il puntatore su una riga di file e fai clic sulla stella per aggiungerlo o rimuoverlo.
- **Guida dell’utente**: apre questa guida.
- **Genspark Projects**: dopo aver effettuato l'accesso al tuo account Genspark, mostra i progetti che hai creato sul web con Genspark AI; fai clic su uno per continuare a modificarlo nel browser. Sono disponibili ricerca, ordinamento per data, aggiornamento e caricamento di altri elementi.
- **Cartelle**: le directory che aggiungi alla barra laterale con **Aggiungi cartella…**, oppure che ci trascini dal gestore di file. Ognuna diventa una radice che puoi aprire, riempire di sottocartelle, rinominare e rimuovere; una radice che va offline viene mostrata come non disponibile e può essere tolta dall'elenco. **Nuova cartella** ne crea un'altra.

Qui non c’è alcuna voce Cestino. I file eliminati finiscono nel cestino di sistema, e recuperarli è compito del sistema operativo.

## L'elenco dei file

Ogni riga mostra un'icona, il nome del file, la data di modifica e altro. Il **⋯ menu** della riga offre:

- **Apri** e **Mostra nella cartella** per individuare il file nel gestore di file.
- **Copia percorso**.
- **Sposta nella cartella…**: apre un selettore di cartelle e sposta davvero il file; se la destinazione contiene già un file con lo stesso nome, puoi saltare, sovrascrivere o rinominare.
- **Rinomina**: rinomina direttamente lì, con l'estensione conservata automaticamente.
- **Aggiungi ai preferiti / Rimuovi dai preferiti** — i preferiti sopravvivono al riavvio e seguono il file quando lo rinomini.
- **Duplica**: crea una copia nella stessa cartella.
- **Elimina**: sposta il file nel cestino di sistema — non è un'eliminazione definitiva.
- **Rimuovi dall'elenco**, nella vista **Recenti** di primo livello, per far sparire una voce senza toccare il file.

### Più file insieme

Spunta la casella di una riga, o fai ⌘/ctrl-clic, per comporre una selezione; la casella dell'intestazione seleziona tutto ciò che è attualmente elencato, e una barra sopra l'elenco indica (**{n} selezionati**) quanti sono selezionati e offre **Sposta nella cartella…** ed **Elimina file** per tutto l'insieme. Puoi anche trascinare una selezione multipla su una cartella nella barra laterale.

## Ricerca

La casella di ricerca in alto filtra contemporaneamente due cose:

- **Nomi dei file**: filtro rapido per nome.
- **Contenuto dei file**: GenOffice indicizza i tuoi file in background (il testo dentro docx/xlsx/pptx/pdf/md/html, con il ricorso all'OCR per i PDF scansionati), così cercando nel corpo del testo trovi file. L'ambito e le opzioni si impostano nelle impostazioni di ricerca.

## Schede di avvio rapido

Le schede sopra gli elenchi creano un nuovo documento in un solo passaggio. Facendo clic su una scheda viene creato un file di quel tipo e aperto il relativo editor — puoi metterti a scrivere subito oppure lasciare che sia l'IA a preparare una bozza per te (ogni editor ha un **pulsante AI** nella barra multifunzione e **Chiedi all'IA** nel menu contestuale della selezione).

Il nuovo file viene creato nella cartella attualmente selezionata nella barra laterale; se non è selezionata alcuna, va nella cartella predefinita.

Cosa fa ogni scheda:

- **AI Docs** (.docx): un documento di testo vuoto nell'editor Docs. Il file viene scritto su disco solo al **primo salvataggio**; i nuovi documenti si aprono con il pannello IA espanso (puoi disattivarlo in Impostazioni → «Apri il pannello IA nei nuovi documenti»).
- **AI Sheets** (.xlsx): un foglio di calcolo vuoto nell'editor Sheets. Finché non salvi, su disco non esiste alcun file — il nome è prenotato per il primo salvataggio; dopo la prima generazione con l'IA il file può anche essere rinominato automaticamente in base al suo contenuto.
- **AI Slides** (.pptx): una presentazione vuota nell'editor Slides.
- **AI Markdown** (.md): un documento Markdown vuoto nell'editor Markdown.
- **AI HTML** (.html): una pagina web vuota nell'editor HTML.
- **AI PDF** (.pdf): diverso dagli altri — crea **immediatamente** un vero PDF vuoto di una sola pagina nella cartella di destinazione e lo apre come file normale (l'editor PDF lavora su file reali). Utile per annotare, oscurare o aggiungere testo; al primo salvataggio il file può essere rinominato automaticamente in base al suo contenuto.
- **Apri file locale**: un selettore di file di sistema per Word (.docx/.doc), Excel (.xlsx/.xlsm/.xls/.csv/.tsv), PowerPoint (.pptx/.ppt), PDF, Markdown (.md/.markdown) e pagine web (.html/.htm). È possibile selezionare più file; ogni file ottiene la propria scheda.

> Suggerimento: File ▸ Nuovo nella barra dei menu crea gli stessi tipi di documento (⌘N/Ctrl+N crea per impostazione predefinita un documento di testo); trascinando un file nella finestra lo si apre.

## Progetti cloud (Genspark Projects)

- Il primo utilizzo richiede l'accesso al tuo account Genspark (flusso con codice dispositivo: GenOffice mostra un codice e tu completi l'accesso nel browser).
- L'elenco dei progetti si sincronizza con il web; Apri nel browser porta a quella pagina per continuare.
- Non effettuare l'accesso non compromette nessuna delle funzioni locali.
