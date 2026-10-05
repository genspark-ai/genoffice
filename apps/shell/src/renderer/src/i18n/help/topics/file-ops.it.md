# Operazioni sui file: rinomina, eliminazione, esportazione

Questo capitolo raccoglie le operazioni sui file comuni a tutti gli editor; le opzioni di esportazione specifiche di ciascun editor sono nel suo capitolo.

Il **⋯ menu** della riga (compare passando il puntatore su una riga di file) raccoglie queste azioni:

![Il ⋯ menu di una riga di file](img/file-ops.png)

## Rinomina

Due punti d'ingresso, un unico insieme di controlli:

- **⋯ ▸ Rinomina** su una riga di Home.
- **Doppio clic su una scheda di file** per rinominare direttamente lì (vedi [Schede e gestione delle finestre](help://tabs-and-windows)).

Regole: l'estensione viene conservata automaticamente; i caratteri non ammessi, i punti finali e i nomi riservati (CON/NUL e simili) vengono rifiutati con un messaggio; anche un conflitto di nome nella stessa cartella viene bloccato. Viene rinominato il vero file su disco, e gli elenchi dei recenti e dei preferiti si aggiornano di conseguenza.

## Eliminazione

- **⋯ ▸ Elimina** su Home: sposta il file nel **cestino di sistema**, ripristinabile dal sistema operativo.
- Dopo l'eliminazione compare per qualche secondo un avviso con la possibilità di annullare — annullare rimette il file al suo posto.

## Duplicazione

**⋯ ▸ Duplica** crea una copia di <name> nella stessa cartella e la apre in una nuova scheda; in caso di nome duplicato viene aggiunto automaticamente un contatore.

## Salva e Salva con nome

- **⌘S / ctrl+S** salva il file corrente; un file senza titolo chiede prima percorso e nome.
- **Salva con nome…** scrive un file nuovo e lascia intatto l'originale; le modifiche successive riguardano il file nuovo.
- Ogni salvataggio è atomico (file temporaneo + rinomina); un'uscita a metà scrittura non può corrompere il file.
- Il salvataggio automatico interviene solo dopo il primo salvataggio manuale (vedi [Avvio rapido](help://getting-started)).

## Esportazione in PDF

- **Docs**: File ▸ Esporta come PDF… (oppure il pulsante nella barra multifunzione), con l'impaginazione corrente.
- **Slides**: l'esportazione rasterizza pagina per pagina, con l'avanzamento per le presentazioni grandi.
- **Sheets**: l'esportazione segue l'impaginazione di stampa.
- Le esportazioni vengono renderizzate in una finestra nascosta e salvate dove scegli tu.

## Esportazione in Word / immagini

- **Esporta come Word…** nel PDF: trasforma il PDF in un file .docx (conversione locale; le impaginazioni complesse vengono riprodotte per quanto possibile).
- **Docs** può esportare le pagine come immagini (PNG per pagina).

## Stampa

File ▸ Stampa… in ogni editor (⌘P/ctrl+P) apre la finestra di stampa di sistema; i PDF si stampano con l'ordine delle pagine e le rotazioni correnti.

## Dove finiscono i file senza titolo

Il percorso che scegli al primo salvataggio è la loro casa; prima di allora il documento esiste solo in memoria. Il salvataggio automatico subentra solo dopo quel primo salvataggio.
