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

## Esportazione

Menu File, tutto locale e tutto chiede dove mettere il risultato:

- **Esporta come Word…** e **Esporta come PDF…** scrivono un vero .docx o .pdf.
- **Esporta come HTML a file singolo…** scrive un unico .html con le immagini incorporate. Non sovrascriverà il file che avete aperto e vi dice quante immagini non è riuscito a incorporare.

## Inserisci scheletro

Per una pagina vuota, **Inserisci ▸ Inserisci scheletro** scrive un documento minimo in modalità standard:

```html
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <title></title>
  </head>
  <body></body>
</html>
```

Ogni parte c'è per una ragione, ed è per questo che è un comando e non qualcosa da digitare:

- il **doctype**, altrimenti l'anteprima funziona in modalità quirks, dove il dimensionamento dei box e il layout delle tabelle seguono regole diverse da quelle che ti aspetti;
- il **`lang`**, altrimenti uno screen reader non ha una lingua in cui leggere la pagina e il browser sceglie un font e un correttore ortografico per la lingua sbagliata;
- il **charset**, altrimenti una pagina di testo non latino può decodificarsi come mojibake.

Un meta viewport è deliberatamente assente: questo documento viene renderizzato in un riquadro desktop senza alcun viewport mobile che ne risenta.

Il `lang` segue la lingua dell'interfaccia dell'app, così lo scheletro che inserisci è quello per cui il tuo tooling è già configurato. Poi puoi modificarlo liberamente.

La voce compare solo in modalità modifica e solo mentre il documento è vuoto — quando c'è contenuto, non c'è nulla _in cui_ inserire uno scheletro.
