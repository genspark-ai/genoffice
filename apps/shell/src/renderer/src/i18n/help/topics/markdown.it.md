# L'editor Markdown

L'editor Markdown apre i file .md / .markdown con un'esperienza fatta di sorgente più anteprima renderizzata.

- **Apri**: da Home oppure File ▸ Apri…; funziona anche dalla riga di comando.
- **Modifica**: editing in testo semplice; le estensioni GFM (tabelle, elenchi di attività, barratura, collegamenti automatici) vengono renderizzate nell'anteprima.
- **Anteprima**: in tempo reale; le risorse relative come le immagini si risolvono accanto al documento.
- **Salvataggio**: fedele byte per byte — il BOM, i CRLF e la presenza di una riga finale vuota vengono preservati; un salvataggio senza modifiche non riscrive il file.
- **Trova e sostituisci**: ctrl+F cerca nel sorgente; Sostituisci tutto scrive il risultato.
- **IA**: i pulsanti predefiniti permettono all'assistente di riscrivere, estendere o tradurre il documento.

## La barra degli strumenti

Una riga di pulsanti sopra l'editor (passa il puntatore per le descrizioni):

![La barra degli strumenti Markdown](img/md-toolbar.png)

- **File e cronologia**: Salva, Salva con nome…, Annulla, Ripeti, Trova; a destra l'interruttore **Salvataggio automatico** scrive le modifiche su disco a intervalli regolari.
- **Pulsante AI**: apre il pannello IA; accanto trovi le impostazioni predefinite riscrivi / estendi / traduci.
- **Stile paragrafo** (menu a tendina): passa dal testo del corpo ai vari livelli di titolo.
- **Formattazione in linea**: **grassetto**, _corsivo_, ~~barratura~~, `codice in linea`, collegamento.
- **Elenchi**: puntato, numerato, elenco di attività.
- **Inserisci**: tabella, immagine, linea orizzontale.
- **Proprietà**: inserisce o salta al blocco YAML front matter in cima al file.
- **Struttura**: salta in base alla gerarchia dei titoli.
- **Controllo ortografico**: attiva o disattiva il controllo ortografico per questo documento.

Tre esempi rapidi:

- **Titolo**: metti il cursore sulla riga ▸ menu a tendina dello stile paragrafo ▸ «Titolo 1».
- **Tabella**: fai clic su **Inserisci tabella** ▸ trascina per scegliere il numero di righe e colonne ▸ digita nelle celle; l'anteprima la renderizza subito.
- **Elenco di attività**: seleziona qualche riga ▸ fai clic su **Elenco di attività** ▸ ogni riga diventa `- [ ]`, resa come caselle di spunta nell'anteprima.
