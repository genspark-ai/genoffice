# L'editor HTML

L'editor HTML apre i file .html / .htm con due modalità: **anteprima** (la pagina renderizzata) e **sorgente**.

- **Anteprima**: rendering reale; i fogli di stile e le immagini relativi vengono caricati accanto al file.
- **Ispettore dell'anteprima**: fai clic per selezionare un elemento, doppio clic per modificarne il testo sul posto, elimina dalla barra degli strumenti e Chiedi all'IA sulla selezione.
- **Modalità sorgente**: modifica l'HTML; ctrl+F per cercare, Sostituisci tutto salva il markup riscritto.
- **Salvataggio**: fedele byte per byte (BOM/CRLF/riga finale preservati); i salvataggi senza modifiche non riscrivono nulla.
- **Zoom**: ctrl+rotellina / pizzicata ridimensiona l'anteprima; ctrl+Z dentro l'anteprima annulla l'ultima modifica.

## La barra degli strumenti

Fai clic su un qualsiasi elemento nell'anteprima e sopra di esso compare una barra degli strumenti:

![La barra mobile sopra un elemento selezionato](img/html-toolbar.png)

- **File e cronologia**: Salva, Salva con nome…, Annulla, Ripeti, Trova; l'interruttore **Salvataggio automatico** scrive le modifiche a intervalli regolari.
- commutatore **Anteprima / Sorgente**; **Presenta** mostra la pagina a schermo intero.
- **Formattazione**: grassetto, corsivo, aumento e riduzione della dimensione del carattere; il **pannello degli stili** dell'elemento selezionato (colori e altro).
- **Inserisci**: titolo, paragrafo, tabella, immagine (tramite collegamento), altro.
- **Azioni sull'immagine** (con un'immagine selezionata): ritaglio, **rimuovi sfondo**, sostituisci, blocca proporzioni.
- **Azioni sull'elemento** (con un elemento selezionato nell'ispettore dell'anteprima): elimina, duplica, sposta su/sposta giù.
- **Pulsante AI**: apre il pannello IA; puoi fare domande direttamente sull'elemento selezionato.
