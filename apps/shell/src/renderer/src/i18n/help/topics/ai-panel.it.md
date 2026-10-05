# Il pannello dell'assistente IA

Ogni editor può richiamare il pannello IA: seleziona qualcosa, dai un'istruzione, guarda il risultato che arriva in streaming.

## Come aprirlo e usarlo

![Il pannello IA in Docs](img/ai-panel.png)

- Punti d'ingresso: il **pulsante AI** nella barra multifunzione di ogni editor, **Chiedi all'IA** nei menu contestuali oppure Chiedi all'IA sulla barra di annotazione.
- Descrivi l'attività in linguaggio naturale (riscrivi questo / trasforma questa colonna in percentuali / riorganizza l'impaginazione di questa pagina...) e premi Invio.
- Le risposte vengono renderizzate **in streaming**; quando l'IA ha bisogno di strumenti (leggere il documento, modificarlo, eseguire uno script) li usa e prosegue fino alla fine.
- **Interrompi**: interrompe il turno in corso in qualsiasi momento.

## Cosa sa fare

- **Docs**: riscrivere/espandere/tradurre/riassumere, inserire tabelle e immagini, modificare la formattazione; ogni turno crea prima uno snapshot.
- **Sheets**: formule, riempimento dei dati, trasformazioni in blocco, formattazione.
- **Slides**: generazione dell'intera presentazione, regolazione del layout, riscrittura dei testi.
- **PDF**: domande e risposte e riassunti sul testo o sulle pagine selezionate.
- **Markdown / HTML**: riscrivere, estendere, tradurre.

## Ripristino e sicurezza

- Il pannello di Docs mantiene un **elenco delle versioni**: uno snapshot per turno, puoi tornare a qualunque versione e il ripristino stesso è annullabile con Ctrl+Z. Gli snapshot sopravvivono alla riapertura del documento.
- Le modifiche dell'IA passano dalla stessa pipeline di modifica di quelle manuali (annullabili, subordinate al salvataggio) — niente aggira la tua conferma di salvataggio.

## Privacy

- Le istruzioni e il contenuto del documento pertinente vengono inviati al **servizio modello che hai configurato** (Genspark ospitato o un endpoint personalizzato, capitolo successivo); senza configurazione non viene inviato nulla.
- I file locali non vengono caricati da nessun'altra parte; le chiavi BYOK vivono solo nelle intestazioni delle richieste — mai su disco o nei registri.
