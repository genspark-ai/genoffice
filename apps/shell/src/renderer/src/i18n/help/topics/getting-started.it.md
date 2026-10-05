# Avvio rapido: l'interfaccia e le basi

GenOffice è una suite per ufficio che funziona interamente sulla tua macchina: una sola finestra, una fila di schede, con sei editor — Docs (elaborazione testi), Sheets (fogli di calcolo), Slides (presentazioni), PDF, Markdown e HTML. I file sono veri file .docx / .xlsx / .pptx / .pdf, completamente intercambiabili con Word, Excel e PowerPoint. Nessuna connessione di rete richiesta.

## Panoramica dell'interfaccia

![La schermata Home](img/home-screen.png)

La finestra è composta da tre parti:

- **Barra delle schede (in alto)**: ogni file aperto è una scheda. La scheda Home, all'estrema sinistra, è sempre presente e non può essere chiusa; le altre schede sono i tuoi documenti. Fai doppio clic su una scheda per rinominare il file direttamente lì.
- **Area di contenuto**: l'editor (o Home) appartenente alla scheda attiva.
- **Barra dei menu**: nella barra dei menu di sistema su macOS, in cima alla finestra su Windows/Linux. I menu File/Modifica/Visualizza si adattano all'editor attivo.

## Creare un documento

Uno qualsiasi di questi:

- Fai clic su una scheda di creazione rapida nella sezione [Avvio rapido](help://getting-started) di **Home** (AI Docs, AI Sheets, AI Slides, ...).
- Menu **File ▸ Nuovo**: Docs (⌘N/ctrl+N), Sheets, Slides, Markdown, HTML o PDF.
- Trascina un file sulla finestra, oppure fai doppio clic su di esso nel gestore di file (se GenOffice è l'app predefinita).

Un nuovo documento si apre senza titolo; il file su disco viene creato solo al primo salvataggio.

## Apertura dei file

- Il menu **File ▸ Apri…** (⌘O/ctrl+O) apre il selettore di sistema: .docx / .doc / .xlsx / .xls / .csv / .tsv / .pptx / .ppt / .pdf / .md / .html.
- Fai clic su qualsiasi voce dell'elenco **Recenti** di Home.
- Anche `genoffice <file>` eseguito da un terminale apre i file.

## Il modello di salvataggio

- **Salvataggio manuale**: ⌘S/ctrl+S, oppure File ▸ Salva / Salva con nome…. Al primo salvataggio ti verranno richiesti percorso e nome.
- **Il salvataggio automatico** si attiva solo dopo che hai salvato il file manualmente almeno una volta — un PDF che hai soltanto letto non viene mai riscritto di nascosto. Il salvataggio automatico interviene poco dopo che il contenuto è cambiato.
- Se chiudi una scheda con modifiche non salvate, ti viene prima chiesto di scegliere tra Salva / Scarta / Annulla.
- Ogni scrittura è atomica (file temporaneo + rinomina), così un'interruzione di corrente non può lasciare metà file.

## Scorciatoie comuni

| Azione               | macOS | Windows / Linux |
| -------------------- | ----- | --------------- |
| Nuovo documento      | ⌘N    | ctrl+N          |
| Apri                 | ⌘O    | ctrl+O          |
| Salva                | ⌘S    | ctrl+S          |
| Chiudi scheda        | ⌘W    | ctrl+W          |
| Apri questo manuale | F1    | F1              |

Le scorciatoie interne a ogni editor (pennello di formattazione, trova e sostituisci, operazioni sulle tabelle, ...) si trovano nei rispettivi capitoli; Docs dispone inoltre di una finestra di dialogo ricercabile con tutte le scorciatoie da tastiera (vedi il suo capitolo).

## Dove andare poi

- Dove si trovano i file: [La schermata Home](help://home-screen).
- Gestire molti file aperti: [Schede e gestione delle finestre](help://tabs-and-windows).
- Far fare il lavoro all'IA: [Il pannello dell'assistente IA](help://ai-panel).
- Lingua, tema, app predefinite: [Impostazioni, lingua, tema e integrazioni MCP](help://settings-integrations).
