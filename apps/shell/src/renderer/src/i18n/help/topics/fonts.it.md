# Caratteri: font di sistema e famiglie scaricabili

## Gli elenchi dei caratteri

I controlli dei caratteri si trovano nella scheda Home di ogni editor (qui sotto: il gruppo dei caratteri in Slides):

![Il gruppo dei caratteri nella barra multifunzione di Slides](img/fonts.png)

Il selettore dei caratteri di ogni editor unisce: i caratteri installati localmente + i candidati comuni della piattaforma (le famiglie più diffuse di Windows e macOS, western e CJK, compresi i nomi localizzati con cui i caratteri CJK si presentano sui propri sistemi). Quando il sistema riesce a enumerare i caratteri locali l'elenco viene raggruppato in base a ciò che esiste davvero; in caso contrario ricade sull'elenco completo dei candidati.

## Caratteri scaricabili (Slides)

- Il selettore dei caratteri di Slides apre un **catalogo di caratteri**: una selezione curata di font OFL che copre molti sistemi di scrittura; ogni famiglia fornisce i pesi normale e grassetto.
- Scegliendone uno, viene scaricato e installato dal CDN (con checksum fissato) nell'archivio caratteri dell'app — disponibile in seguito a tutti i documenti, senza alcuna installazione a livello di sistema.
- Le installazioni risiedono in una directory privata dell'app e spariscono con l'app.

## Tipografia CJK

- I caratteri CJK in Docs ottengono la compressione della punteggiatura e le regole di interruzione di linea (kinsoku) allineate a Word (vedi il capitolo su Docs).
- Le famiglie CJK del catalogo coprono cinese semplificato e tradizionale, giapponese e coreano, sia serif sia sans-serif.

## Gestione dei caratteri locali

- La voce per installare i file di caratteri locali carica i file .ttf/.otf dal disco nell'archivio caratteri dell'app.
- L'archivio è suddiviso per famiglia; più pesi della stessa famiglia convivono.
