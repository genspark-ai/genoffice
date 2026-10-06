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

| Azione                        | macOS | Windows / Linux |
| ----------------------------- | ----- | --------------- |
| Nuovo documento               | ⌘N    | ctrl+N          |
| Apri                          | ⌘O    | ctrl+O          |
| Salva                         | ⌘S    | ctrl+S          |
| Chiudi scheda                 | ⌘W    | ctrl+W          |
| Apri questo manuale           | F1    | F1              |
| Riduci la barra multifunzione | ⌥⌘R   | Ctrl+F1         |

**Ridurre la barra multifunzione** funziona in ogni editor. La riga di schede resta al suo posto e la banda dei comandi sottostante viene nascosta; la scheda selezionata fa anche da controllo di riduzione, quindi mentre la barra multifunzione è ridotta nessuna scheda è selezionata e premere una scheda qualsiasi riporta la banda. Un doppio clic su una scheda fa lo stesso. Il modo in cui l'ha lasciata viene ricordato per ogni editor.

Le scorciatoie interne a ogni editor (pennello di formattazione, trova e sostituisci, operazioni sulle tabelle, ...) si trovano nei rispettivi capitoli; Docs dispone inoltre di una finestra di dialogo ricercabile con tutte le scorciatoie da tastiera (**⌘/**) (vedi il suo capitolo).

## Le scorciatoie Option+Comando

Option+Comando è il livello che Word riserva ai salti strutturati, e GenOffice lo riempie allo stesso modo. Docs ne prende la maggior parte, Sheets ne prende due di sue per la parità con Excel; una scorciatoia vale ovunque.

**Docs**

| Scorciatoia (macOS) | A cosa serve                   | Windows / Linux    |
| ------------------- | ------------------------------ | ------------------ |
| ⌥⌘1 / ⌥⌘2 / ⌥⌘3     | Titolo 1 / 2 / 3               | Ctrl+Alt+1 / 2 / 3 |
| ⌥⌘0                 | Normale                        | Ctrl+Alt+0         |
| ⌥⌘M                 | Paragrafo                      | Ctrl+Alt+M         |
| ⌥⌘A                 | Nuovo commento                 | Ctrl+Alt+A         |
| ⌥⌘F                 | Inserisci nota a piè di pagina | Ctrl+Alt+F         |
| ⌥⌘E                 | Inserisci nota di chiusura     | Ctrl+Alt+D         |
| ⌥⌘G                 | Vai a                          | Ctrl+G             |

Due di esse cambiano su Windows, per la stessa ragione per cui Word le separa. **macOS si è preso ⌥⌘D** — mostra e nasconde il Dock —, quindi la nota di chiusura è ⌥⌘E sul Mac e Ctrl+Alt+D ovunque altrove. E **Vai a** perde il tasto Alt: Ctrl+G, mentre la scorciatoia del Mac lo porta con sé.

**Sheets**, mentre la griglia ha il focus

| Scorciatoia (macOS) | A cosa serve  | Windows / Linux |
| ------------------- | ------------- | --------------- |
| ⌥⌘0                 | Bordi esterni | Ctrl+Shift+7    |
| ⌥⌘−                 | Nessun bordo  | Ctrl+Shift+−    |

Windows non è una riscrittura della coppia del Mac. Excel per Mac dà a Sheets **entrambe** — ⌘⇧7 e ⌥⌘0 sono due tasti per gli stessi bordi esterni —, quindi su Windows il comando conserva lo slot Ctrl+Shift che aveva già e il livello Option è semplicemente assente.

Attenzione: **⌥⌘0 significa Normale in Docs e Bordi esterni in Sheets**. Non compaiono mai nello stesso editor, quindi in uso non collide nulla, ma ⌥⌘0 è già impegnato e non è disponibile come scorciatoia globale.

**Ogni editor**: **⌥⌘R / Ctrl+F1** riduce la barra multifunzione, come descritto sopra.

Così ⌥⌘D resta libero perché GenOffice lo usi su macOS, se un futuro comando lo vorrà.

## Dove andare poi

- Dove si trovano i file: [La schermata Home](help://home-screen).
- Gestire molti file aperti: [Schede e gestione delle finestre](help://tabs-and-windows).
- Far fare il lavoro all'IA: [Il pannello dell'assistente IA](help://ai-panel).
- Lingua, tema, app predefinite: [Impostazioni, lingua, tema e integrazioni MCP](help://settings-integrations).
