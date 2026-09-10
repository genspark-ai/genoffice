<p align="center">
  <a href="https://genoffice.ai/">
    <picture>
      <source srcset="../assets/readme/hero-dark.webp" media="(prefers-color-scheme: dark)">
      <img src="../assets/readme/hero.webp" alt="GenOffice — la suite Office AI open source: Docs, Sheets, Slides, PDF, Markdown e HTML con un pannello AI integrato" width="100%">
    </picture>
  </a>
</p>

<h1 align="center">GenOffice</h1>

<p align="center"><b>La prima suite office AI open source e completa al mondo.</b><br>
File Word, Excel, PowerPoint e PDF, modificati da te e dalla tua AI, salvati negli stessi formati reali.</p>

<p align="center">
  <a href="../../LICENSE"><img src="https://img.shields.io/github/license/genspark-ai/genoffice" alt="License: Apache-2.0"></a>
  <a href="https://github.com/genspark-ai/genoffice/releases/latest"><img src="https://img.shields.io/github/v/release/genspark-ai/genoffice" alt="Latest release"></a>
  <a href="https://github.com/genspark-ai/genoffice/releases"><img src="https://img.shields.io/github/downloads/genspark-ai/genoffice/total" alt="Downloads"></a>
  <a href="https://github.com/genspark-ai/genoffice/stargazers"><img src="https://img.shields.io/github/stars/genspark-ai/genoffice?style=flat" alt="GitHub stars"></a>
</p>

<p align="center"><a href="../../README.md">English</a> · <a href="README.zh-CN.md">简体中文</a> · <a href="README.zh-TW.md">繁體中文</a> · <a href="README.ja.md">日本語</a> · <a href="README.ko.md">한국어</a> · <a href="README.de.md">Deutsch</a> · <a href="README.fr.md">Français</a> · <a href="README.es.md">Español</a> · <a href="README.pt-BR.md">Português (Brasil)</a> · <a href="README.ru.md">Русский</a> · <b>Italiano</b> · <a href="README.nl.md">Nederlands</a> · <a href="README.pl.md">Polski</a> · <a href="README.cs.md">Čeština</a> · <a href="README.id.md">Bahasa Indonesia</a> · <a href="README.ms.md">Bahasa Melayu</a> · <a href="README.th.md">ไทย</a> · <a href="README.hi.md">हिन्दी</a> · <a href="README.ar.md">العربية</a> · <a href="README.he.md">עברית</a></p>

<p align="center">
  <a href="#download"><b>Download</b></a> ·
  <a href="https://genoffice.ai/"><b>Sito web</b></a> ·
  <a href="https://genoffice.ai/join"><b>Community</b></a> ·
  <a href="../../PRIVACY.md"><b>Privacy</b></a>
</p>

GenOffice è un'alternativa gratuita e open source a Microsoft Office per
macOS, Windows e Linux. Apre e salva file nativi `.docx`, `.xlsx` e `.pptx`,
permette di modificare PDF, Markdown e HTML, e affianca un agente AI a ogni
documento — non una chat incollata a fianco, ma un editor che legge il file,
applica la modifica e ti mostra esattamente cosa ha toccato.

- **Formati reali, preservati byte per byte.** Viene riscritto solo ciò che
  modifichi. Tutto il resto del file resta identico byte per byte, così i
  documenti continuano a funzionare in Word, Excel e PowerPoint.
- **AI che puoi controllare.** Le modifiche arrivano come revisioni tracciate
  e diff con ripristino a un clic. I fogli di calcolo ottengono formule live,
  non numeri incollati. Le presentazioni e le pagine vengono generate
  direttamente sulla canvas e restano completamente modificabili.
- **Locale per progettazione.** I file si aprono, si modificano, si salvano e
  si convertono sul tuo computer. Le conversioni PDF → Word / Excel /
  PowerPoint, Markdown → Word e HTML → Word avvengono tutte in locale. Solo
  le chiamate AI lasciano la macchina, verso il provider che scegli.
- **Le tue chiavi, o nessuna.** Accedi con Genspark e non dovrai configurare
  nulla, oppure usa la tua chiave per Claude, OpenAI, Gemini, DeepSeek, Kimi,
  GLM, Qwen, Doubao, MiniMax, Grok, Mistral, OpenRouter o qualsiasi endpoint
  compatibile con OpenAI, inclusi i server locali.

**Scaricalo:** [macOS](https://github.com/genspark-ai/genoffice/releases/latest) (Apple Silicon e Intel) ·
[Windows](https://github.com/genspark-ai/genoffice/releases/latest) (x64 e Arm) ·
[Linux](https://github.com/genspark-ai/genoffice/releases/latest) (deb, rpm, AppImage) —
dettagli e requisiti nella sezione [Download](#download).

## Demo

Sei app, un solo pannello AI. Ogni screenshot mostra l'app reale su macOS,
con l'AI guidata dal prompt che puoi leggere nel pannello.

### 1 · Docs — apri e modifica `.docx` con un'AI che puoi controllare

<table>
<tr>
<td width="50%"><img src="../assets/readme/docs-report.webp" alt="GenOffice Docs visualizza una pagina di relazione annuale a due colonne con un'immagine di copertina a piena larghezza, una tabella KPI con sfondo colorato, intestazione e piè di pagina, a zoom 80% con il pannello AI ridotto a icona"></td>
<td width="50%"><img src="../assets/readme/docs-ai.webp" alt="GenOffice Docs: una panoramica aziendale con un'immagine banner; l'AI ha reso più conciso il paragrafo Overview e inserito una nuova sezione a elenco puntato, e il pannello offre un ripristino a un clic"></td>
</tr>
<tr>
<td><b>Apre il file esattamente come lo impagina Word</b> — sezioni a due colonne, immagini a piena pagina, tabelle con sfondo colorato, intestazioni e piè di pagina, impaginazione sulle stesse metriche di riga di Word. Stili, commenti, revisioni tracciate, equazioni e input a penna vengono preservati senza alterazioni.</td>
<td><b>Chiedi la modifica</b> — l'AI legge i blocchi di cui ha bisogno, riscrive l'Overview e inserisce una nuova sezione a elenco puntato. Ogni intervento dell'AI è uno snapshot che puoi ripristinare; con <b>Revisioni</b> attivate, le modifiche arrivano come revisioni in stile Word.</td>
</tr>
</table>

### 2 · Sheets — file `.xlsx` con formule e grafici live, non numeri incollati

<table>
<tr>
<td width="50%"><img src="../assets/readme/sheets-ai.webp" alt="GenOffice Sheets: l'AI ha aggiunto un foglio Summary con i ricavi per area e categoria usando formule SUMIF, più un grafico a colonne, e segnala 43 modifiche applicate con un pulsante Annulla"></td>
<td width="50%"><img src="../assets/readme/sheets-qa.webp" alt="GenOffice Sheets: alla domanda su quale area ha guidato i ricavi del Q2, l'AI risponde Europa con la suddivisione per categoria e cita come link le celle utilizzate, accanto al foglio Orders"></td>
</tr>
<tr>
<td><b>Costruiscilo</b> — da una singola frase l'agente aggiunge un foglio Summary con vere <code>SUMIF</code> per area e categoria, inserisce un grafico a colonne e applica le 43 modifiche come un unico batch annullabile.</td>
<td><b>Interrogalo</b> — le domande sulla cartella di lavoro ottengono risposta con il ragionamento e le celle esatte usate come citazioni cliccabili. Dietro le quinte: un motore <code>.xlsx</code> Rust proprietario, tabelle pivot, filtri dati, formattazione condizionale e tracciamento delle formule.</td>
</tr>
</table>

### 3 · Slides — da un prompt a una presentazione `.pptx`

<img src="../assets/readme/slides-generate.webp" alt="Time-lapse di GenOffice Slides che genera il deck per investitori Aurora Home: l'IA pianifica la narrazione nel pannello, le diapositive appaiono sulla tela una dopo l'altra e il deck finito si conclude con la richiesta finale" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/slides-cover.webp" alt="GenOffice Slides: la slide di copertina di una presentazione per investitori Aurora Home generata dall'AI sulla canvas, con il prompt originale di una riga e il riepilogo dell'AI su cosa ha creato, nel pannello"></td>
<td width="50%"><img src="../assets/readme/slides-ai.webp" alt="GenOffice Slides: la slide di chiusura progettata della stessa presentazione di 11 slide, con la striscia di miniature a sinistra e il pannello AI che riassume la storyline"></td>
</tr>
<tr>
<td><b>Una riga in ingresso</b> — "Crea una presentazione per investitori di 10 slide per Aurora Home…". GenOffice pianifica la storyline, ricerca i numeri e disegna ogni slide sulla canvas come una vera <code>.pptx</code>.</td>
<td><b>Una presentazione finita in uscita</b> — undici slide progettate con tipografia e immagini coerenti e una call to action di chiusura; continua a modificare con master, layout, guide intelligenti e ritaglio non distruttivo, oppure chiedi al pannello di cambiare stile, riscrivere e riordinare.</td>
</tr>
</table>

### 4 · PDF — modifica il testo dei PDF sul posto, converti PDF in Word in locale

<table>
<tr>
<td width="50%"><img src="../assets/readme/pdf-edit.webp" alt="GenOffice PDF: la modalità Modifica testo delimita ogni blocco di testo della pagina per la modifica sul posto, mentre il pannello AI risponde a una domanda sul report con citazioni di pagina"></td>
<td width="50%"><img src="../assets/readme/pdf-convert.webp" alt="GenOffice Docs mostra un documento Word convertito in locale dal PDF Helios quarterly review, aperto in una seconda scheda accanto al PDF originale"></td>
</tr>
<tr>
<td><b>Modifica direttamente nella pagina</b> — la modalità Modifica testo delimita ogni blocco di testo per riscriverlo sul posto; il content stream viene riscritto tramite PDFium con i font originali, non con un'annotazione che copre il testo. Chiedi all'AI informazioni su un report lungo e ottieni risposte con citazioni di pagina.</td>
<td><b>Converti in locale</b> — <b>PDF Converter → PDF to Word</b> produce un <code>.docx</code> modificabile che si apre in Docs accanto all'originale, con titoli, righe di statistiche e paragrafi intatti. Le conversioni verso Excel e PowerPoint funzionano allo stesso modo; le pagine scansionate passano attraverso l'OCR di sistema.</td>
</tr>
</table>

### 5 · HTML — un generatore AI di pagine e interfacce, a partire da un design brief

Descrivi a cosa serve la pagina e a chi è destinata. L'AI propone prima un
**design brief** — hook, palette, tipografia e direzioni di stile — poi
costruisce un unico file `.html` autosufficiente basato su quei token.

<img src="../assets/readme/html-restyle-motion.webp" alt="Time-lapse di GenOffice HTML che ristilizza la landing page Lumen: una richiesta di Restyle nel pannello trasforma la pagina scura Midnight Studio nella calda versione Solar Daybreak, mantenendo invariate tutte le sezioni e tutti i testi" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/html-ai.webp" alt="GenOffice HTML: una landing page generata per una lampada da scrivania solare nella direzione scura Midnight Studio, mostrata nella live preview con il pannello AI che riassume la pagina appena creata"></td>
<td width="50%"><img src="../assets/readme/html-restyle.webp" alt="La stessa landing page Lumen ristilizzata dall'AI nella calda direzione Solar Daybreak: sfondo carta, titoli serif e un accento arancione, con tutte le sezioni e tutti i testi mantenuti"></td>
</tr>
<tr>
<td><b>Generata da un solo prompt</b> — un hero d'impatto, card delle funzionalità, pricing e un modulo waitlist per Lumen, realizzati nella direzione Midnight Studio. Clicca su qualsiasi elemento per cambiarne lo stile, fai doppio clic per modificare il testo, oppure passa alla vista sorgente di CodeMirror.</td>
<td><b>Stesso design, nuova direzione</b> — una singola richiesta di <b>Restyle</b> sostituisce i token del brief e la pagina si adatta: carta calda, serif editoriale, accento arancio sole, senza riscrivere nulla. Presenta a schermo intero, oppure esporta come PDF o come documento Word nativo e modificabile.</td>
</tr>
</table>
<table>
<tr>
<td width="50%"><img src="../assets/readme/html-dashboard.webp" alt="GenOffice HTML: un'interfaccia dashboard personale generata per un designer freelance in stile lino caldo, con una barra laterale sinistra, un saluto in serif e quattro card di metriche"></td>
<td width="50%"><img src="../assets/readme/html-report.webp" alt="GenOffice HTML: un report dati generato sul mercato dei veicoli elettrici in stile broadsheet, con una testata serif, una cifra principale di 17,3 milioni e una riga di statistiche"></td>
</tr>
<tr>
<td><b>Mockup UI</b> — lo starter "personal dashboard" trasforma una persona in un layout funzionante: barra laterale, saluto, sparkline delle ore fatturabili, card di fatture e utilizzo, tutto HTML reale che puoi consegnare a un developer.</td>
<td><b>Data stories</b> — lo starter "data report" costruisce un broadsheet editoriale: testata serif, un numero in evidenza, una riga di statistiche separata da un filetto, grafici SVG inline e una nota metodologica.</td>
</tr>
</table>

### 6 · Markdown — un block editor su file `.md` semplici, con Ask AI

<table>
<tr>
<td width="50%"><img src="../assets/readme/markdown-ai.webp" alt="GenOffice Markdown: un paragrafo selezionato mostra un popover Ask AI con un'istruzione digitata e chip di suggerimento come Migliora, Rendi più conciso, Espandi e Correggi la grammatica, più i pulsanti Invia ora e Aggiungi alla coda"></td>
<td width="50%"><img src="../assets/readme/markdown-render.webp" alt="GenOffice Markdown visualizza un documento di note di lancio con una tabella, un diagramma di flusso Mermaid e un elenco di attività, con i prompt di partenza del pannello AI a sinistra"></td>
</tr>
<tr>
<td><b>Ask AI su una selezione</b> — seleziona qualsiasi passaggio e appare un chip <b>Ask AI</b>: digita un'istruzione o scegli un suggerimento, invia subito, oppure accoda diverse modifiche ancorate ed eseguile in un'unica passata. Lo stesso punto di accesso esiste in ogni app.</td>
<td><b>Visualizzato, salvato come Markdown semplice</b> — titoli, elenchi, tabelle, immagini, blocchi di codice e diagrammi Mermaid in un block editor Tiptap, riscritti come <code>.md</code> semplice, con un'esportazione <b>Markdown → Word</b> completamente locale.</td>
</tr>
</table>

## Perché GenOffice

- **Open source**, Apache-2.0, sviluppato in modo aperto su GitHub.
- **Tuo da eseguire.** App native per macOS, Windows e Linux; i file restano
  sul tuo disco e ogni modifica, salvataggio e conversione avviene sul tuo
  computer.
- **File Office reali.** `.docx`, `.xlsx` e `.pptx` nativi, con preservazione
  byte per byte: le parti del file che non hai toccato vengono copiate
  esattamente come erano.
- **Un'AI che modifica direttamente il documento.** Revisioni tracciate in
  Docs, formule e grafici live in Sheets, slide disegnate sulla canvas, ogni
  intervento dell'AI genera uno snapshot da cui puoi tornare indietro.
- **Il tuo modello, la tua chiave.** Accedi con Genspark, oppure usa una
  chiave per Claude, OpenAI, Gemini, DeepSeek e altri, con server locali e
  qualsiasi endpoint compatibile con OpenAI incluso.
- **PDF fatto bene.** Modifica il testo direttamente nella pagina e converti
  PDF in Word, Excel o PowerPoint in locale, con OCR di sistema per le
  scansioni.
- **Anche Markdown e HTML**, con lo stesso pannello AI e l'esportazione locale
  in Word.
- **Gratuito**, per singoli utenti e team.

## Backend AI

**Accedi con Genspark** e non c'è nulla da configurare: le chiamate ai
modelli passano attraverso il proxy Genspark (famiglie Claude, GPT e Gemini)
e gli agenti hanno accesso a ricerca web e per immagini, generazione di
immagini e analisi di immagini/audio/video.

**Oppure usa la tua chiave.** In Settings → AI trovi Claude, OpenAI, Gemini,
DeepSeek, Kimi, GLM, Qwen, Doubao, MiniMax, Grok, Mistral, OpenRouter e
OpenCode Zen/Go, più uno slot personalizzato per qualsiasi endpoint
compatibile con OpenAI (base URL + chiave), inclusi i server con modelli
locali. Ricerca e media hanno i propri provider per singola funzionalità
sotto **AI Media & Search**: Serper o Tavily per la ricerca web, e OpenAI,
Gemini, Doubao/Seedream, GLM, Grok, Qwen, MiniMax o qualsiasi endpoint di
immagini compatibile con OpenAI per la generazione di immagini e l'analisi
di immagini/video.

L'intera suite include i temi chiaro, scuro e di sistema. I temi cambiano
solo ciò che vedi a schermo: esportazioni, stampe e file salvati mantengono
sempre i colori originali del documento.

<a id="download"></a>

## Download

| Piattaforma                                | Requisiti                                             | Download                                                                                   |
| ------------------------------------------ | ----------------------------------------------------- | ------------------------------------------------------------------------------------------ |
| **macOS** — Apple Silicon (arm64)          | macOS 11+                                             | [Ultima versione `.dmg` (arm64)](https://github.com/genspark-ai/genoffice/releases/latest) |
| **macOS** — Intel (x64)                    | macOS 11+                                             | [Ultima versione `.dmg` (x64)](https://github.com/genspark-ai/genoffice/releases/latest)   |
| **Windows** (x64, la maggior parte dei PC) | Windows 10+, Intel/AMD                                | [Ultimo installer `-x64.exe`](https://github.com/genspark-ai/genoffice/releases/latest)    |
| **Windows** su Arm (ARM64)                 | Windows 11 su Arm (Snapdragon X e simili)             | [Ultimo installer `-arm64.exe`](https://github.com/genspark-ai/genoffice/releases/latest)  |
| **Linux** — Debian / Ubuntu                | x86_64, glibc 2.34+ (Ubuntu 22.04 o successivo)       | [Ultimo `.deb`](https://github.com/genspark-ai/genoffice/releases/latest)                  |
| **Linux** — Fedora / RHEL / openSUSE       | x86_64, glibc 2.34+ (Fedora 35+, RHEL 9+, Leap 15.6+) | [Ultimo `.rpm`](https://github.com/genspark-ai/genoffice/releases/latest)                  |
| **Linux** — altre distribuzioni            | x86_64, glibc 2.34+, FUSE 2                           | [Ultimo `.AppImage`](https://github.com/genspark-ai/genoffice/releases/latest)             |

Tutte le build provengono da `main`; gli installer per macOS e Windows sono
firmati. Le versioni precedenti sono disponibili nella pagina
[Releases](https://github.com/genspark-ai/genoffice/releases).

<details>
<summary><b>Installazione su Linux</b></summary>

Il deb si installa con apt — scarica le dipendenze e aggiunge GenOffice al
menu delle applicazioni:

```bash
sudo apt install ./genoffice_<version>_amd64.deb
```

Su Fedora / famiglia RHEL / openSUSE, installa invece l'rpm:

```bash
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / RHEL family
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

L'AppImage viene eseguito così com'è: installa il runtime FUSE 2
(`sudo apt install libfuse2`; su Ubuntu 24.04 il pacchetto è `libfuse2t64`),
rendi il file eseguibile, poi eseguilo:

```bash
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

</details>

## Come funziona

Sette app Electron — Docs, Sheets, Slides, PDF, Markdown, HTML e la shell a
schede — condividono un unico livello motore fatto di package TypeScript
puri più un sidecar Rust per `.xlsx`. Il file originale è sempre la fonte di
verità: le modifiche vengono applicate come patch mirate, e tutto ciò che
l'editor non ha toccato sopravvive intatto al ciclo di apertura e
salvataggio.

```
open docx ─► archive original by hash (never touched)
          ─► parse word/document.xml into a block tree, each block anchored to its original XML
          ─► Tiptap editor (manual + AI editing, dirty tracking)
save      ─► dirty blocks → OOXML fragments (referencing existing styles only)
          ─► splice into the original document.xml; untouched blocks keep their bytes
          ─► repack the zip; every other entry is copied byte-for-byte
```

Il tour package per package (motori docx/pptx, `pdf2docx`, `html2docx`, il
core dell'agente e i provider) si trova in
[CONTRIBUTING.md](../../CONTRIBUTING.md#engine-packages).

## Sviluppo

```bash
npm install
npm run fixtures     # generate test .docx fixtures
npm test             # engine + app unit tests (docs/sheets/slides need no display)
npm run typecheck    # tsc --noEmit across every workspace
npm run dev          # all six editors + shell against Vite dev servers
npm run dev:docs     # a single app (same pattern works per workspace)
npm run dist:mac     # package macOS dmg (regenerates third-party notices)
npm run dist:win     # package Windows nsis installer
npm run dist:linux   # package Linux AppImage + deb + rpm
```

L'app sheets richiede inoltre una toolchain Rust per il suo sidecar xlsx
(`cargo` nel PATH); `npm run build -w @genoffice/sheets` la compila
automaticamente. Consulta [CONTRIBUTING.md](../../CONTRIBUTING.md) per i
controlli che ogni modifica deve superare e per come vengono accettate le
pull request.

## Community

GenOffice è in sviluppo attivo e il tuo feedback lo plasma.

- **Segnala un bug o richiedi una funzionalità** su
  [GitHub Issues](https://github.com/genspark-ai/genoffice/issues).
- **Unisciti alla chat di gruppo di GenOffice** su
  [GenTeam](https://genoffice.ai/join) per parlare con il team e con altri
  utenti.
- **Metti una stella al repo** se GenOffice ti è utile — è il modo migliore
  per supportare il progetto.

## FAQ

<details>
<summary><b>GenOffice è gratuito?</b></summary>

Sì. GenOffice è gratuito e open source sotto licenza Apache-2.0 — nessuna
prova, nessun piano a pagamento per le app in sé.

</details>

<details>
<summary><b>GenOffice può aprire file Microsoft Word, Excel e PowerPoint?</b></summary>

Sì. GenOffice apre e salva file nativi `.docx`, `.xlsx` e `.pptx`. Il
salvataggio preserva i byte: le parti del file che non hai toccato vengono
riscritte byte per byte, così i documenti continuano a funzionare in
Microsoft Office.

</details>

<details>
<summary><b>GenOffice funziona offline?</b></summary>

La modifica dei documenti è completamente locale — i file non lasciano mai
il tuo computer per essere aperti, modificati, salvati o convertiti. Le
funzionalità AI (agenti, ricerca, strumenti per le immagini) richiedono una
connessione di rete, con l'accesso Genspark oppure con una tua chiave API
per il modello.

</details>

<details>
<summary><b>GenOffice può modificare file PDF?</b></summary>

Sì — una vera modifica di testo e immagini nel PDF, che riscrive il content
stream della pagina preservando i font originali, non con annotazioni che
coprono il testo.

</details>

<details>
<summary><b>GenOffice può convertire PDF in Word, Excel o PowerPoint?</b></summary>

Sì — completamente in locale: estrazione a livello di carattere con PDFium
più un'analisi del layout basata sulla geometria, nessun servizio cloud,
nessun upload. Sono coperte anche le pagine scansionate: su macOS e Windows
l'OCR di sistema le legge, così vengono convertite in testo modificabile
invece che in un'immagine di pagina.

</details>

<details>
<summary><b>Posso usare un mio modello AI o una mia chiave API?</b></summary>

Sì. Oltre all'accesso Genspark senza chiavi, GenOffice supporta l'uso della
tua chiave per Claude, OpenAI, Gemini, DeepSeek, Kimi, GLM, Qwen, Doubao,
MiniMax, Grok, Mistral, OpenRouter e OpenCode Zen/Go, più qualsiasi endpoint
compatibile con OpenAI — inclusi i server con modelli locali. Ricerca,
generazione di immagini e analisi di immagini/video richiedono le proprie
chiavi in Settings → AI Media & Search.

</details>

<details>
<summary><b>GenOffice può convertire HTML in Word?</b></summary>

Sì — Export as Word nell'app HTML produce un `.docx` nativo e modificabile,
interamente in locale. La pagina viene renderizzata nel Chromium integrato e
ridotta a vere strutture Word: titoli, paragrafi, elenchi, tabelle, card,
righe KPI, campi modulo e sfondi di pagina; solo gli elementi visivi senza
un equivalente in Word (grafici, icone, box decorati) vengono incorporati
come immagini.

</details>

<details>
<summary><b>GenOffice raccoglie dati?</b></summary>

Le build ufficiali inviano per impostazione predefinita analytics d'uso
limitate, e puoi disattivare l'invio in qualsiasi momento in Settings →
General. Le analytics non inviano mai il contenuto dei documenti, i nomi dei
file, i percorsi dei file, l'identità dell'account o gli indirizzi email.
Consulta [GenOffice Privacy](../../PRIVACY.md) per l'elenco completo di
eventi e dati raccolti.

</details>

## Sicurezza

Consulta [SECURITY.md](../../SECURITY.md) per la postura di sicurezza dei
processi (sandboxing del renderer, validazione IPC, filtro dei link esterni)
e i modelli di minaccia per i contenuti generati dall'AI.

## Ringraziamenti

GenOffice non sarebbe possibile senza questi progetti open source:

- [Electron](https://www.electronjs.org/) — il runtime desktop per ogni app.
- [Univer](https://github.com/dream-num/univer) (Apache-2.0) — il core
  dell'interfaccia per fogli di calcolo che Sheets estende.
- [PDFium](https://pdfium.googlesource.com/pdfium/) (BSD-3-Clause, incluso
  tramite [@embedpdf/pdfium](https://github.com/embedpdf/embed-pdf-viewer))
  — il motore del content stream dietro la vera modifica di testo e immagini
  nei PDF.
- [pdf.js](https://github.com/mozilla/pdf.js) (Apache-2.0) e
  [pdf-lib](https://github.com/Hopding/pdf-lib) (MIT) — rendering PDF e
  assemblaggio dei documenti.
- [Tiptap](https://tiptap.dev/) / [ProseMirror](https://prosemirror.net/) —
  i block editor di Docs e Markdown.
- [CodeMirror](https://codemirror.net/) (MIT) — l'editor sorgente nell'app
  HTML.
- [Konva](https://konvajs.org/) — rendering su canvas per Slides e per i
  grafici di Sheets.
- [HarfBuzz](https://github.com/harfbuzz/harfbuzz) (wasm) — metriche di
  text-shaping per le scritture complesse.
- [calamine](https://github.com/tafia/calamine) e
  [IronCalc](https://github.com/ironcalc/IronCalc) — i livelli di lettura e
  calcolo del sidecar Rust per xlsx.
- [libeot](https://github.com/umanwizard/libeot) (MPL-2.0) — il decoder
  MicroType Express per i font PowerPoint incorporati, portato in
  TypeScript.
- [React](https://react.dev/) (MIT) — il livello dell'interfaccia utente di
  ogni app.
- [Mermaid](https://mermaid.js.org/) (MIT) e [KaTeX](https://katex.org/)
  (MIT) — diagrammi e formule matematiche in Markdown e Docs.
- [opentype.js](https://opentype.js.org/) (MIT) — l'analisi dei font per le
  metriche e la ricerca dei glifi.
- [JSZip](https://stuk.github.io/jszip/) (MIT) e
  [fast-xml-parser](https://github.com/NaturalIntelligence/fast-xml-parser)
  (MIT) — i livelli del contenitore OOXML e XML.
- [Fluent UI System Icons](https://github.com/microsoft/fluentui-system-icons)
  (MIT) — il set di icone usato nelle barre multifunzione.
- [electron-updater](https://www.electron.build/) (MIT) — gli aggiornamenti
  all'interno dell'app.
- I font Liberation, Carlito, Caladea e Noto CJK (OFL/Apache-2.0) — font per
  documenti incluse nel bundle.

`npm run notices` rigenera il riepilogo delle licenze di terze parti incluse
(`tools/gen-third-party-notices.mjs`); tutte le dipendenze a runtime sono
MIT/Apache-2.0/BSD-3-Clause/OFL.

## Licenza

GenOffice è distribuito con licenza [Apache License 2.0](../../LICENSE), con
un'unica eccezione: la directory `ee/` è riservata a futuri moduli
enterprise ed è coperta dalla
[GenOffice Enterprise License](../../ee/LICENSE).

I nomi e i loghi GenOffice e Genspark sono marchi di Mainfunc, Inc. La
licenza Apache-2.0 non concede il permesso di utilizzarli (vedi la sezione
6); i fork dovrebbero usare un proprio branding.
