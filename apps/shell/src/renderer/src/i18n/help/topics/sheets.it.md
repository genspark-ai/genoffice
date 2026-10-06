# Sheets: fogli di calcolo

Sheets è l'editor simile a Excel; il calcolo avviene in un processo separato, un motore scritto in Rust (un eventuale crash lì non porta giù l'app). Apre e salva veri file .xlsx; i file .csv e .tsv si aprono come tabelle.

## L'interfaccia

- **Barra multifunzione**: otto schede, descritte una alla volta qui sotto.
- **Barra delle formule**: mostra e modifica la formula della cella attiva; sono supportate le funzioni comuni.
- **Schede dei fogli** (in basso): aggiungi / rinomina / elimina / sposta i fogli.
- **Modifica di una cella**: doppio clic oppure digita e basta; Invio conferma e scende, Tab va a destra, Esc annulla (le abitudini di Excel).
- **Scorciatoie**: in linea con la famiglia Excel (ctrl+C/V/X, ctrl+Z/Y, ctrl+F, ...).

## Le schede della barra multifunzione

- **Home**: carattere, riempimento, bordi, formati numerici (valuta/percentuale/migliaia, aumento e riduzione dei decimali), allineamento, unione, inserimento e dimensione di righe e colonne, formattazione condizionale, formatta come tabella, stili delle celle, appunti e pennello di formattazione, ordinamento e filtro.
- **Inserisci**: forme, icone, simboli, equazioni, schermate e altro.
- **Layout di pagina**: colori e caratteri del tema, interruttori per linee della griglia e intestazioni in stampa, anteprima delle interruzioni di pagina.
- **Formule**: Somma automatica e inserimento delle funzioni, definizione dei nomi (anche dalla selezione), traccia dei precendenti e dei dipendenti, Finestra di controllo, ricalcolo del foglio o della cartella di lavoro.
- **Dati**: ordinamento e filtro (compreso il filtro avanzato e la cancellazione del filtro), da testo a colonne, unione di cartelle di lavoro, aggiorna tutto.
- **Revisione**: sfoglia i commenti (mostra, precedente/successivo), traduci.
- **Visualizza**: interruttori per linee della griglia e intestazioni, zoom, visualizzazione Normale / anteprima delle interruzioni di pagina.
- **Struttura grafico**: compare con un grafico selezionato — tipo di grafico, stili e colori, modifica dell'intervallo di dati.

La scheda Dati, pulsante per pulsante (da sinistra a destra nell'immagine):

![La scheda Dati](img/sheets-data.png)

- **Tabella pivot**: costruisce una tabella pivot dall'intervallo corrente; trascina i campi per aggregare.
- **Aggiorna**: ricalcola i dati della tabella pivot corrente.
- **Da testo/CSV**: importa un file .csv/.txt come nuovo foglio, suddiviso in base al delimitatore.
- **Unisci cartelle di lavoro**: porta dentro questo file i fogli di altri file .xlsx.
- **Aggiorna tutto**: ricalcola tutte le tabelle pivot e tutte le origini dati esterne.
- **Ordina** (menu a tendina): crescente / decrescente / ordinamento personalizzato (regole su più colonne).
- **Filtro**: aggiunge menu a tendina ▼ alla riga di intestazione; spunta i valori da mantenere.
- I piccoli pulsanti impilati accanto: **Cancella** (riporta tutte le righe), **Riapplica** (riesegue il filtro corrente), **Avanzato** (filtra usando un'area di criteri).
- **Da testo a colonne** (menu a tendina): divide una colonna in più colonne in base al delimitatore o a una larghezza fissa.
- **Compilazione rapida**: dai un esempio e il resto della colonna si riempie seguendo l'esempio (ctrl+E).
- **Rimuovi duplicati**: elimina le righe duplicate in base alle colonne selezionate.
- **Convalida dati** (menu a tendina): regole di immissione per la selezione (elenchi a tendina, intervalli di numeri...).
- **Consolida**: aggrega più intervalli in un unico punto per categoria.
- **Analisi di simulazione** (menu a tendina): ricerca del valore — risolve una cella di input in modo che una cella con formula raggiunga il valore obiettivo.
- **Raggruppa / Annulla raggruppamento** (menu a tendina): gruppi di righe o colonne con compressione ed espansione.
- **Subtotale**: inserisce righe di subtotale per ogni categoria.

La scheda Formule, pulsante per pulsante:

![La scheda Formule](img/sheets-formulas.png)

- **Inserisci funzione** (fx): cerca le funzioni con una procedura guidata per gli argomenti.
- **Sommatoria automatica** (menu a tendina): SOMMA con un clic, più media/conteggio/massimo/minimo.
- **Usate di recente / Finanziarie / Logiche / Testo / Data e ora / Ricerca e riferimento / Matematiche e trigonometriche / Altre**: sfoglia e inserisci funzioni per categoria.
- **Gestore nomi**: visualizza, crea ed elimina intervalli con nome.
- **Definisci nome** (menu a tendina): assegna un nome alla selezione; **Usa in formula** inserisce un nome esistente; **Crea da selezione** assegna i nomi agli intervalli in base alla riga o colonna di intestazione.
- **Traccia precendenti / Traccia dipendenti**: frecce blu che mostrano da dove arrivano i dati di una formula e dove essa alimenta il foglio; **Rimuovi frecce** le cancella.
- **Mostra formule**: le celle mostrano la formula stessa invece del risultato.
- **Controllo errori**: individua e spiega gli errori di formula.
- **Finestra di controllo**: blocca le celle che ti interessano e osserva il loro valore in tempo reale.
- **Opzioni di calcolo** (menu a tendina): ricalcolo automatico o manuale; in modalità manuale **Calcola ora / Calcola foglio** lo avvia a comando.

## Numeri e formattazione

- Formati numerici: generale, numero, valuta, percentuale, data/ora, frazione, notazione scientifica e altro.
- Allineamento, a capo, celle unite, bordi e riempimenti.
- Altezze delle righe e larghezze delle colonne trascinando; doppio clic sul bordo per adattare automaticamente.

## Dati

**Ordinamento e filtro** (per esempio, decrescente su una colonna):

1. Fai clic su **una qualsiasi cella di quella colonna** (non serve selezionare tutta la colonna).
2. Scheda Home ▸ **Ordina e filtra** ▸ **Decrescente**; le righe si riordinano insieme per intero (l'area viene ordinata come un blocco unico).
3. Per regole personalizzate (più colonne, per colore): stesso percorso, **Ordinamento personalizzato**.
4. Filtro: seleziona la riga di intestazione e fai clic su **Ordina e filtra ▸ Filtro** — ogni intestazione riceve un menu a tendina ▼ in cui spuntare i valori da mantenere; cancellando il filtro torna tutto.

- Ordinamento e filtro.
- Blocca finestre.
- .csv / .tsv: si aprono direttamente come tabella (il tsv delimitato da tabulazioni viene letto come uno solo); il salvataggio scrive indietro nel formato originale.

## Menu contestuali

- **Nella griglia**: il menu proprio dell'editor (Univer) — taglia/copia/incolla, inserisci ed elimina righe/colonne, nascondi, unisci celle, blocca finestre e le altre voci di uso quotidiano.
- **Sulla barra di stato in basso**: scegli quali statistiche mostrare nella barra di stato (media / conteggio / somma, ...); la scelta viene ricordata.
- **Su una scheda di foglio in basso**: aggiungi / rinomina / elimina / colora / nascondi fogli (menu schede di Univer).
- Il menu contestuale della barra delle schede in alto è descritto in [Schede e gestione delle finestre](help://tabs-and-windows).

## IA

- Il pannello IA laterale: seleziona un intervallo e dai istruzioni in linguaggio naturale (riformattare, generare dati, scrivere formule).
- Le modifiche dell'IA si possono annullare dal pannello.

## Salvataggio ed esportazione

- Salva .xlsx (formule e formati conservati); Salva con nome…; l'esportazione in PDF segue l'impaginazione di stampa.
- Il salvataggio automatico segue la regola globale (attivo dopo il primo salvataggio manuale).

## Stabilità

- Il motore di calcolo in Rust è isolato in un processo separato rispetto all'interfaccia: se dati estremi lo fanno terminare, vedi un messaggio e un tentativo di ripristino della sessione — non un crash dell'app.
