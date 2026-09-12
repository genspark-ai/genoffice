<p align="center">
  <a href="https://genoffice.ai/">
    <picture>
      <source srcset="../assets/readme/hero-dark.webp" media="(prefers-color-scheme: dark)">
      <img src="../assets/readme/hero.webp" alt="GenOffice – die Open-Source-KI-Office-Suite: Docs, Sheets, Slides, PDF, Markdown und HTML mit integriertem KI-Panel" width="100%">
    </picture>
  </a>
</p>

<h1 align="center">GenOffice</h1>

<p align="center"><b>Die weltweit erste vollwertige Open-Source-KI-Office-Suite.</b><br>
Word-, Excel-, PowerPoint- und PDF-Dateien, bearbeitet von dir und deiner KI, gespeichert in den echten Formaten.</p>

<p align="center">
  <a href="../../LICENSE"><img src="https://img.shields.io/github/license/genspark-ai/genoffice" alt="Lizenz: Apache-2.0"></a>
  <a href="https://github.com/genspark-ai/genoffice/releases/latest"><img src="https://img.shields.io/github/v/release/genspark-ai/genoffice" alt="Neueste Version"></a>
  <a href="https://github.com/genspark-ai/genoffice/releases"><img src="https://img.shields.io/github/downloads/genspark-ai/genoffice/total" alt="Downloads"></a>
  <a href="https://github.com/genspark-ai/genoffice/stargazers"><img src="https://img.shields.io/github/stars/genspark-ai/genoffice?style=flat" alt="GitHub-Sterne"></a>
</p>

<p align="center"><a href="../../README.md">English</a> · <a href="README.es.md">Español</a> · <a href="README.pt-BR.md">Português (Brasil)</a> · <b>Deutsch</b> · <a href="README.fr.md">Français</a> · <a href="README.zh-CN.md">简体中文</a> · <a href="README.zh-TW.md">繁體中文</a> · <a href="README.ko.md">한국어</a> · <a href="README.ja.md">日本語</a> · <a href="README.ar.md">العربية</a> · <a href="README.ru.md">Русский</a> · <a href="README.it.md">Italiano</a> · <a href="README.nl.md">Nederlands</a> · <a href="README.pl.md">Polski</a> · <a href="README.cs.md">Čeština</a> · <a href="README.id.md">Bahasa Indonesia</a> · <a href="README.ms.md">Bahasa Melayu</a> · <a href="README.th.md">ไทย</a> · <a href="README.hi.md">हिन्दी</a> · <a href="README.he.md">עברית</a></p>

<p align="center">
  <a href="#download"><b>Download</b></a> ·
  <a href="https://genoffice.ai/"><b>Website</b></a> ·
  <a href="https://genoffice.ai/join"><b>Community</b></a> ·
  <a href="../../PRIVACY.md"><b>Datenschutz</b></a>
</p>

GenOffice ist eine kostenlose Open-Source-Alternative zu Microsoft Office für
macOS, Windows und Linux. Es öffnet und speichert native `.docx`-, `.xlsx`-
und `.pptx`-Dateien, bearbeitet PDF, Markdown und HTML und stellt jedem
Dokument einen KI-Agenten zur Seite – keine angeflanschte Chat-Box, sondern
ein Editor, der die Datei liest, die Änderung vornimmt und genau zeigt, was
er verändert hat.

- **Echte Formate, byteerhaltend.** Nur das, was du bearbeitest, wird neu
  geschrieben. Alles andere in der Datei bleibt Byte für Byte erhalten,
  sodass Dokumente in Word, Excel und PowerPoint weiter funktionieren.
- **KI, die du überprüfen kannst.** Änderungen kommen als nachverfolgte
  Änderungen und Diffs mit Rückgängig-Funktion per Klick an. Tabellen
  erhalten echte Formeln statt eingefügter Zahlen. Präsentationen und Seiten
  werden direkt auf der Arbeitsfläche erzeugt und bleiben vollständig
  bearbeitbar.
- **Lokal von Grund auf.** Dateien werden auf deinem Rechner geöffnet,
  bearbeitet, gespeichert und konvertiert. PDF → Word / Excel / PowerPoint,
  Markdown → Word und HTML → Word laufen vollständig auf dem Gerät. Nur die
  KI-Aufrufe verlassen den Rechner – zu dem Anbieter, den du wählst.
- **Deine Schlüssel oder keine.** Melde dich mit Genspark an und du brauchst
  keine Schlüssel, oder bring deinen eigenen Schlüssel für Claude, OpenAI,
  Gemini, DeepSeek, Kimi, GLM, Qwen, Doubao, MiniMax, Grok, Mistral,
  OpenRouter, API Route, Requesty oder jeden OpenAI-kompatiblen Endpunkt mit – auch lokale Server
  werden unterstützt.

**Los geht's:** [macOS](https://github.com/genspark-ai/genoffice/releases/latest) (Apple Silicon und Intel) ·
[Windows](https://github.com/genspark-ai/genoffice/releases/latest) (x64 und Arm) ·
[Linux](https://github.com/genspark-ai/genoffice/releases/latest) (deb, rpm, AppImage) —
Details und Voraussetzungen unter [Download](#download).

## Demo

Sechs Apps, ein KI-Panel. Jeder Screenshot zeigt die echte App auf macOS,
wobei die KI durch den Prompt gesteuert wird, den du im Panel nachlesen
kannst.

### 1 · Docs – `.docx` öffnen und bearbeiten mit einer KI, die du überprüfen kannst

<table>
<tr>
<td width="50%"><img src="../assets/readme/docs-report.webp" alt="GenOffice Docs stellt eine zweispaltige Geschäftsberichtsseite mit vollbreitem Titelbild, schattierter KPI-Tabelle, Kopf- und Fußzeile bei 80 % Zoom mit eingeklapptem KI-Panel dar"></td>
<td width="50%"><img src="../assets/readme/docs-ai.webp" alt="GenOffice Docs: eine Unternehmensübersicht mit einem Banner-Bild; die KI hat die Übersicht gestrafft und einen neuen Abschnitt mit Aufzählungspunkten eingefügt, und das Panel bietet ein Rückgängig per Klick"></td>
</tr>
<tr>
<td><b>Öffnet die Datei genau so, wie Word sie layoutet</b> — zweispaltige Abschnitte, randlose Bilder, schattierte Tabellen, Kopf- und Fußzeilen, Seitenumbruch nach Words Zeilenmetrik. Formatvorlagen, Kommentare, nachverfolgte Änderungen, Formeln und Tinteneingaben bleiben beim Roundtrip unverändert.</td>
<td><b>Frag nach der Änderung</b> — die KI liest die Blöcke, die sie braucht, schreibt die Übersicht neu und fügt einen neuen Abschnitt mit Aufzählungspunkten ein. Jeder KI-Durchgang ist ein Snapshot, den du zurücksetzen kannst; ist <b>Änderungen nachverfolgen</b> aktiviert, kommen die Bearbeitungen als Word-typische Überarbeitungen an.</td>
</tr>
</table>

### 2 · Sheets – `.xlsx` mit echten Formeln und Diagrammen statt eingefügter Zahlen

<table>
<tr>
<td width="50%"><img src="../assets/readme/sheets-ai.webp" alt="GenOffice Sheets: die KI hat ein Summary-Tabellenblatt mit Umsatz nach Region und Kategorie über SUMIF-Formeln sowie ein Säulendiagramm hinzugefügt und meldet 43 angewendete Änderungen mit einer Rückgängig-Schaltfläche"></td>
<td width="50%"><img src="../assets/readme/sheets-qa.webp" alt="GenOffice Sheets: auf die Frage, welche Region den Q2-Umsatz anführte, antwortet die KI mit Europa samt Aufschlüsselung nach Kategorie und verlinkt die verwendeten Zellen als Belege, neben dem Orders-Tabellenblatt"></td>
</tr>
<tr>
<td><b>Bau es</b> — aus einem einzigen Satz fügt der Agent ein Summary-Tabellenblatt mit echten <code>SUMIF</code>s nach Region und Kategorie hinzu, fügt ein Säulendiagramm ein und wendet die 43 Änderungen als einen einzigen, rückgängig machbaren Vorgang an.</td>
<td><b>Frag es</b> — Fragen zur Arbeitsmappe werden mit der Begründung und den genau verwendeten Zellen als klickbare Belege beantwortet. Im Hintergrund: eine hausgemachte Rust-<code>.xlsx</code>-Engine, Pivot-Tabellen, Datenschnitte, bedingte Formatierung und Formel-Tracing.</td>
</tr>
</table>

### 3 · Slides – vom Prompt zur `.pptx`-Präsentation

<img src="../assets/readme/slides-generate.webp" alt="Zeitraffer von GenOffice Slides bei der Erstellung des Investoren-Decks für Aurora Home: Die KI plant die Storyline im Panel, die Folien erscheinen nacheinander auf der Leinwand, und das fertige Deck endet mit dem abschließenden Ask" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/slides-cover.webp" alt="GenOffice Slides: die Titelfolie einer von der KI erstellten Investoren-Präsentation für Aurora Home auf der Arbeitsfläche, mit dem ursprünglichen Ein-Zeilen-Prompt und der Zusammenfassung der KI im Panel, was sie erstellt hat"></td>
<td width="50%"><img src="../assets/readme/slides-ai.webp" alt="GenOffice Slides: die gestaltete Abschlussfolie derselben 11-Folien-Präsentation, mit der Miniaturansicht-Leiste links und dem KI-Panel, das den Handlungsbogen zusammenfasst"></td>
</tr>
<tr>
<td><b>Eine Zeile rein</b> — „Erstelle eine 10-Folien-Investoren-Präsentation für Aurora Home…“. GenOffice plant den Handlungsbogen, recherchiert die Zahlen und entwirft jede Folie direkt auf der Arbeitsfläche als echte <code>.pptx</code>.</td>
<td><b>Eine fertige Präsentation raus</b> — elf gestaltete Folien mit einheitlicher Typografie, Bildsprache und einem abschließenden Call-to-Action; weiter bearbeiten mit Master-Folien, Layouts, intelligenten Hilfslinien und zerstörungsfreiem Zuschneiden, oder das Panel bitten, neu zu stylen, umzuschreiben und neu zu ordnen.</td>
</tr>
</table>

### 4 · PDF – PDF-Text direkt bearbeiten, PDF lokal in Word konvertieren

<table>
<tr>
<td width="50%"><img src="../assets/readme/pdf-edit.webp" alt="GenOffice PDF: Der Textbearbeitungsmodus umrahmt jeden Textblock auf der Seite zur direkten Bearbeitung, während das KI-Panel eine Frage zum Bericht mit Seitenangaben beantwortet"></td>
<td width="50%"><img src="../assets/readme/pdf-convert.webp" alt="GenOffice Docs zeigt ein Word-Dokument, das lokal aus dem PDF des Helios-Quartalsberichts konvertiert wurde, geöffnet in einem zweiten Tab neben dem Original-PDF"></td>
</tr>
<tr>
<td><b>Direkt auf der Seite bearbeiten</b> — der Textbearbeitungsmodus umrahmt jeden Textblock zum direkten Neutippen; der Inhaltsstream wird über PDFium mit den Originalschriften neu geschrieben, nicht mit einer überdeckenden Anmerkung. Frag die KI zu einem langen Bericht und erhalte Antworten mit Seitenangaben.</td>
<td><b>Lokal konvertieren</b> — <b>PDF-Konverter → PDF zu Word</b> erzeugt eine bearbeitbare <code>.docx</code>, die neben der Quelle in Docs geöffnet wird, mit intakten Überschriften, Kennzahlenzeilen und Absätzen. Die Ziele Excel und PowerPoint funktionieren genauso; gescannte Seiten laufen durch die System-OCR.</td>
</tr>
</table>

### 5 · HTML – ein KI-Seiten- und UI-Builder, zuerst das Design-Briefing

Sag, wofür die Seite gedacht ist und für wen. Die KI schlägt zunächst ein
**Design-Briefing** vor – Aufhänger, Farbpalette, Typografie und
Stilrichtungen – und baut dann eine einzige, in sich geschlossene `.html`-Datei
nach diesen Vorgaben.

<img src="../assets/readme/html-restyle-motion.webp" alt="Zeitraffer von GenOffice HTML beim Neugestalten der Lumen-Landingpage: Eine einzige Restyle-Anfrage im Panel verwandelt die dunkle Midnight Studio-Seite in die warme Solar Daybreak-Version, während jeder Abschnitt und der gesamte Text unverändert bleiben" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/html-ai.webp" alt="GenOffice HTML: eine generierte Landingpage für eine Solar-Schreibtischlampe in der dunklen Stilrichtung „Midnight Studio“, gezeigt in der Live-Vorschau mit dem KI-Panel, das die gerade erstellte Seite zusammenfasst"></td>
<td width="50%"><img src="../assets/readme/html-restyle.webp" alt="Dieselbe Lumen-Landingpage, von der KI in die warme Stilrichtung „Solar Daybreak“ umgestylt: Papierhintergrund, Serifen-Überschriften und ein Orange-Akzent, wobei jeder Abschnitt und der gesamte Text erhalten bleiben"></td>
</tr>
<tr>
<td><b>Aus einem Prompt generiert</b> — ein kräftiger Hero-Bereich, Feature-Karten, Preise und ein Warteliste-Formular für Lumen, gebaut in der Stilrichtung „Midnight Studio“. Klick auf ein beliebiges Element, um es umzustylen, doppelklicke, um Text zu bearbeiten, oder wechsle in die CodeMirror-Quellansicht.</td>
<td><b>Gleiches Design, neue Richtung</b> — eine einzige <b>Restyle</b>-Anfrage tauscht die Vorgaben des Briefings aus, und die Seite folgt: warmes Papier, editorielle Serifenschrift, sonnenoranger Akzent, ohne dass etwas neu geschrieben wird. Im Vollbild präsentieren oder als PDF oder natives, bearbeitbares Word-Dokument exportieren.</td>
</tr>
</table>
<table>
<tr>
<td width="50%"><img src="../assets/readme/html-dashboard.webp" alt="GenOffice HTML: ein generiertes persönliches Dashboard-UI für eine freiberufliche Designerin im warmen Leinen-Stil, mit einer linken Leiste, einer Begrüßung in Serifenschrift und vier Kennzahlen-Karten"></td>
<td width="50%"><img src="../assets/readme/html-report.webp" alt="GenOffice HTML: ein generierter Datenreport zum E-Auto-Markt im Broadsheet-Stil, mit einem Serifen-Kopf, einer Schlagzeile von 17,3 Millionen und einer Kennzahlenzeile"></td>
</tr>
<tr>
<td><b>UI-Mockups</b> — die Vorlage „Personal Dashboard“ macht aus einer Persona ein funktionierendes Layout: linke Leiste, Begrüßung, Sparkline für abrechenbare Stunden, Rechnungs- und Auslastungskarten – alles echtes HTML, das du direkt an eine Entwicklerin oder einen Entwickler weitergeben kannst.</td>
<td><b>Datengeschichten</b> — die Vorlage „Data Report“ baut ein editorielles Broadsheet: Serifen-Kopf, eine Schlagzeilenzahl, eine durch Linien abgetrennte Kennzahlenzeile, eingebettete SVG-Diagramme und eine Methodik-Notiz.</td>
</tr>
</table>

### 6 · Markdown – ein Block-Editor für einfaches `.md`, mit Ask AI

<table>
<tr>
<td width="50%"><img src="../assets/readme/markdown-ai.webp" alt="GenOffice Markdown: In einem ausgewählten Absatz erscheint ein Ask-AI-Popover mit einer eingegebenen Anweisung und Vorschlags-Chips wie „Polish“, „Make more concise“, „Expand“ und „Fix grammar“, außerdem den Schaltflächen „Send now“ und „Add to queue“"></td>
<td width="50%"><img src="../assets/readme/markdown-render.webp" alt="GenOffice Markdown stellt ein Launch-Notes-Dokument mit einer Tabelle, einem Mermaid-Flussdiagramm und einer Aufgabenliste dar, mit den Start-Prompts des KI-Panels links"></td>
</tr>
<tr>
<td><b>Ask AI zu einer Auswahl</b> — markiere eine beliebige Textstelle, und ein <b>Ask AI</b>-Chip erscheint: eine Anweisung eingeben oder einen Vorschlag auswählen, sofort senden, oder mehrere verankerte Änderungen in eine Warteschlange legen und in einem Durchgang ausführen. Denselben Einstiegspunkt gibt es in jeder App.</td>
<td><b>Gerendert, gespeichert als reines Markdown</b> — Überschriften, Listen, Tabellen, Bilder, Codeblöcke und Mermaid-Diagramme in einem Tiptap-Block-Editor, zurückgeschrieben als reines <code>.md</code>, mit einem vollständig lokalen Markdown-→-Word-Export.</td>
</tr>
</table>

## Warum GenOffice

- **Open Source**, Apache-2.0, offen entwickelt auf GitHub.
- **Läuft auf deinem Rechner.** Native Apps für macOS, Windows und Linux;
  Dateien bleiben auf deiner Festplatte, und jede Bearbeitung, jedes Speichern
  und jede Konvertierung findet auf deinem Rechner statt.
- **Echte Office-Dateien.** Native `.docx`, `.xlsx` und `.pptx`, byteerhaltend:
  Teile der Datei, die du nicht angefasst hast, werden unverändert kopiert.
- **Eine KI, die das Dokument selbst bearbeitet.** Nachverfolgte Änderungen in
  Docs, echte Formeln und Diagramme in Sheets, Folien werden direkt auf die
  Arbeitsfläche gezeichnet, und jeder KI-Schritt hinterlässt einen Snapshot,
  den du zurückrollen kannst.
- **Dein Modell, dein Schlüssel.** Melde dich mit Genspark an oder bring einen
  eigenen Schlüssel für Claude, OpenAI, Gemini, DeepSeek und mehr mit – lokale
  Server und jeder OpenAI-kompatible Endpunkt sind ebenfalls möglich.
- **PDF richtig gemacht.** Text direkt auf der Seite bearbeiten und PDF lokal
  in Word, Excel oder PowerPoint konvertieren, mit System-OCR für Scans.
- **Auch Markdown und HTML**, mit demselben KI-Panel und lokalem Export nach
  Word.
- **Kostenlos**, für Einzelpersonen wie für Teams.

## KI-Backends

**Mit Genspark anmelden**, und es gibt nichts zu konfigurieren:
Modellaufrufe laufen über den Genspark-Proxy (die Modellfamilien Claude, GPT
und Gemini), und die Agenten erhalten Web- und Bildsuche, Bildgenerierung
sowie Bild-/Audio-/Videoanalyse.

**Oder bring deinen eigenen Schlüssel mit.** Unter Einstellungen → KI stehen
Claude, OpenAI, Gemini, DeepSeek, Kimi, GLM, Qwen, Doubao, MiniMax, Grok,
Mistral, OpenRouter, API Route, Requesty und OpenCode Zen/Go zur Wahl, plus ein freier Slot für
jeden OpenAI-kompatiblen Endpunkt (Basis-URL + Schlüssel), einschließlich
lokaler Modell-Server. Suche und Medien haben eigene Anbieter je Fähigkeit
unter KI-Medien & Suche: Serper oder Tavily für die Websuche und OpenAI,
Gemini, Doubao/Seedream, GLM, Grok, Qwen, MiniMax oder jeder
OpenAI-kompatible Bild-Endpunkt für Bildgenerierung und Bild-/Videoanalyse.

Die gesamte Suite bringt ein helles, ein dunkles und einen
Systemdesign-Modus mit. Designs ändern nur, was auf dem Bildschirm zu sehen
ist: Exporte, Ausdrucke und gespeicherte Dateien behalten immer die eigenen
Farben des Dokuments.

<a id="download"></a>

## Download

| Plattform                            | Voraussetzungen                                       | Download                                                                                    |
| ------------------------------------ | ----------------------------------------------------- | ------------------------------------------------------------------------------------------- |
| **macOS** – Apple Silicon (arm64)    | macOS 11+                                             | [Neueste `.dmg` (arm64)](https://github.com/genspark-ai/genoffice/releases/latest)          |
| **macOS** – Intel (x64)              | macOS 11+                                             | [Neueste `.dmg` (x64)](https://github.com/genspark-ai/genoffice/releases/latest)            |
| **Windows** (x64, die meisten PCs)   | Windows 10+, Intel/AMD                                | [Neuester `-x64.exe`-Installer](https://github.com/genspark-ai/genoffice/releases/latest)   |
| **Windows** auf Arm (ARM64)          | Windows 11 auf Arm (Snapdragon X und Ähnliche)        | [Neuester `-arm64.exe`-Installer](https://github.com/genspark-ai/genoffice/releases/latest) |
| **Linux** – Debian / Ubuntu          | x86_64, glibc 2.34+ (Ubuntu 22.04 oder neuer)         | [Neueste `.deb`](https://github.com/genspark-ai/genoffice/releases/latest)                  |
| **Linux** – Fedora / RHEL / openSUSE | x86_64, glibc 2.34+ (Fedora 35+, RHEL 9+, Leap 15.6+) | [Neueste `.rpm`](https://github.com/genspark-ai/genoffice/releases/latest)                  |
| **Linux** – andere Distributionen    | x86_64, glibc 2.34+, FUSE 2                           | [Neueste `.AppImage`](https://github.com/genspark-ai/genoffice/releases/latest)             |

Alle Builds stammen aus `main`; die macOS- und Windows-Installer sind
signiert. Ältere Versionen findest du auf der Seite
[Releases](https://github.com/genspark-ai/genoffice/releases).

<details>
<summary><b>Installation unter Linux</b></summary>

Das deb-Paket wird mit apt installiert – es zieht die Abhängigkeiten nach
und fügt GenOffice dem Anwendungsmenü hinzu:

```bash
sudo apt install ./genoffice_<version>_amd64.deb
```

Installiere auf Fedora / RHEL-Familie / openSUSE stattdessen das rpm-Paket:

```bash
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / RHEL family
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

Das AppImage läuft direkt: installiere die FUSE-2-Laufzeit
(`sudo apt install libfuse2`; auf Ubuntu 24.04 heißt das Paket
`libfuse2t64`), mache die Datei ausführbar und starte sie dann:

```bash
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

</details>

## Wie es funktioniert

Sieben Electron-Apps — Docs, Sheets, Slides, PDF, Markdown, HTML und die
Tab-Shell — teilen sich eine Engine-Schicht aus reinen TypeScript-Paketen
plus einem Rust-Sidecar für `.xlsx`. Die Originaldatei ist immer die Quelle
der Wahrheit: Änderungen werden als eng begrenzte Patches angewendet, und
alles, was der Editor nicht berührt hat, überlebt den Roundtrip unverändert.

```
open docx ─► archive original by hash (never touched)
          ─► parse word/document.xml into a block tree, each block anchored to its original XML
          ─► Tiptap editor (manual + AI editing, dirty tracking)
save      ─► dirty blocks → OOXML fragments (referencing existing styles only)
          ─► splice into the original document.xml; untouched blocks keep their bytes
          ─► repack the zip; every other entry is copied byte-for-byte
```

Die Paket-für-Paket-Tour (docx-/pptx-Engines, `pdf2docx`, `html2docx`, der
Agent-Kern und die Provider) findest du in
[CONTRIBUTING.md](../../CONTRIBUTING.md#engine-packages).

## Entwicklung

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

Die Sheets-App benötigt zusätzlich eine Rust-Toolchain für ihren
xlsx-Sidecar (`cargo` im PATH); `npm run build -w @genoffice/sheets`
kompiliert ihn automatisch. Die Prüfungen, die jede Änderung bestehen muss,
und wie Pull Requests landen, stehen in
[CONTRIBUTING.md](../../CONTRIBUTING.md).

## Community

GenOffice wird aktiv weiterentwickelt, und dein Feedback prägt es mit.

- **Melde einen Bug oder wünsche dir ein Feature** in
  [GitHub Issues](https://github.com/genspark-ai/genoffice/issues).
- **Tritt dem GenOffice-Gruppenchat** auf
  [GenTeam](https://genoffice.ai/join) bei, um mit dem Team und anderen
  Nutzern zu sprechen.
- **Gib dem Repo einen Stern**, wenn dir GenOffice nützlich ist – das ist
  die beste Art, das Projekt zu unterstützen.

## FAQ

<details>
<summary><b>Ist GenOffice kostenlos?</b></summary>

Ja. GenOffice ist kostenlos und Open Source unter der Apache-2.0-Lizenz –
keine Testversion, keine kostenpflichtige Stufe für die Apps selbst.

</details>

<details>
<summary><b>Kann GenOffice Microsoft Word-, Excel- und PowerPoint-Dateien öffnen?</b></summary>

Ja. GenOffice öffnet und speichert native `.docx`-, `.xlsx`- und
`.pptx`-Dateien. Das Speichern ist byteerhaltend: Teile der Datei, die du
nicht angefasst hast, werden Byte für Byte zurückgeschrieben, sodass
Dokumente in Microsoft Office weiter funktionieren.

</details>

<details>
<summary><b>Funktioniert GenOffice offline?</b></summary>

Die Dokumentbearbeitung ist vollständig lokal – Dateien verlassen deinen
Rechner nie, um geöffnet, bearbeitet, gespeichert oder konvertiert zu
werden. Die KI-Funktionen (Agenten, Suche, Bildwerkzeuge) benötigen eine
Netzwerkverbindung, entweder über eine Genspark-Anmeldung oder deinen
eigenen Modell-API-Schlüssel.

</details>

<details>
<summary><b>Kann GenOffice PDF-Dateien bearbeiten?</b></summary>

Ja – echte PDF-Text- und Bildbearbeitung, die den Seiteninhalts-Stream unter
Erhalt der Originalschriften neu schreibt, statt überdeckender Anmerkungen.

</details>

<details>
<summary><b>Kann GenOffice PDF in Word, Excel oder PowerPoint konvertieren?</b></summary>

Ja – vollständig lokal: zeichenweise Extraktion über PDFium plus
geometriebasierte Layoutanalyse, kein Cloud-Dienst, kein Upload. Auch
gescannte Seiten werden abgedeckt: Unter macOS und Windows liest die
System-OCR sie aus, sodass sie zu bearbeitbarem Text statt zu einem
Seitenbild konvertiert werden.

</details>

<details>
<summary><b>Kann ich mein eigenes KI-Modell oder meinen eigenen API-Schlüssel verwenden?</b></summary>

Ja. Neben der schlüssellosen Genspark-Anmeldung unterstützt GenOffice das
Mitbringen eigener Schlüssel für Claude, OpenAI, Gemini, DeepSeek, Kimi,
GLM, Qwen, Doubao, MiniMax, Grok, Mistral, OpenRouter, API Route, Requesty und OpenCode Zen/Go,
plus jeden OpenAI-kompatiblen Endpunkt – einschließlich lokaler
Modell-Server. Suche, Bildgenerierung und Bild-/Videoanalyse benötigen
eigene Schlüssel unter Einstellungen → KI-Medien & Suche.

</details>

<details>
<summary><b>Kann GenOffice HTML in Word konvertieren?</b></summary>

Ja – **Als Word exportieren** in der HTML-App erzeugt vollständig lokal
eine native, bearbeitbare `.docx`. Die Seite wird im eingebauten Chromium
gerendert und auf echte Word-Strukturen reduziert: Überschriften, Absätze,
Listen, Tabellen, Karten, KPI-Zeilen, Formularfelder und Seitenhintergründe;
nur visuelle Elemente ohne Word-Entsprechung (Diagramme, Icons, dekorierte
Boxen) werden als Bilder eingebettet.

</details>

<details>
<summary><b>Erfasst GenOffice irgendwelche Daten?</b></summary>

Offizielle gepackte Builds senden standardmäßig begrenzte Nutzungsanalysen,
und du kannst die Übermittlung jederzeit unter Einstellungen → Allgemein
deaktivieren. Die Analysen senden niemals Dokumentinhalte, Dateinamen,
Dateipfade, Kontoidentität oder E-Mail-Adressen. Die vollständigen Angaben
zu Ereignissen und Daten stehen in
[GenOffice-Datenschutz](../../PRIVACY.md).

</details>

## Sicherheit

Die Sicherheitsarchitektur der Prozesse (Renderer-Sandboxing,
IPC-Validierung, Filterung externer Links) und die Bedrohungsmodelle für
KI-generierte Inhalte stehen in [SECURITY.md](../../SECURITY.md).

## Danksagungen

GenOffice wäre ohne diese Open-Source-Projekte nicht möglich:

- [Electron](https://www.electronjs.org/) – die Desktop-Laufzeitumgebung für jede App.
- [Univer](https://github.com/dream-num/univer) (Apache-2.0) – der
  Tabellenkalkulations-UI-Kern, den Sheets erweitert.
- [PDFium](https://pdfium.googlesource.com/pdfium/) (BSD-3-Clause, gebündelt
  über [@embedpdf/pdfium](https://github.com/embedpdf/embed-pdf-viewer)) –
  die Content-Stream-Engine hinter der echten PDF-Text- und
  Bildbearbeitung.
- [pdf.js](https://github.com/mozilla/pdf.js) (Apache-2.0) und
  [pdf-lib](https://github.com/Hopding/pdf-lib) (MIT) – PDF-Rendering und
  Dokumentenzusammenstellung.
- [Tiptap](https://tiptap.dev/) / [ProseMirror](https://prosemirror.net/) –
  die Block-Editoren in Docs und Markdown.
- [CodeMirror](https://codemirror.net/) (MIT) – der Quelltext-Editor in HTML.
- [Konva](https://konvajs.org/) – Canvas-Rendering für Slides und
  Sheets-Diagramme.
- [HarfBuzz](https://github.com/harfbuzz/harfbuzz) (wasm) –
  Text-Shaping-Metriken für komplexe Schriftsysteme.
- [calamine](https://github.com/tafia/calamine) und
  [IronCalc](https://github.com/ironcalc/IronCalc) – die Lese- und
  Berechnungsschichten des Rust-xlsx-Sidecars.
- [libeot](https://github.com/umanwizard/libeot) (MPL-2.0) – der
  MicroType-Express-Decoder für eingebettete PowerPoint-Schriften, nach
  TypeScript portiert.
- [React](https://react.dev/) (MIT) – die UI-Schicht jeder App.
- [Mermaid](https://mermaid.js.org/) (MIT) und [KaTeX](https://katex.org/)
  (MIT) – Diagramme und Mathematik in Markdown und Docs.
- [opentype.js](https://opentype.js.org/) (MIT) – Font-Parsing für Metriken
  und Glyphen-Lookups.
- [JSZip](https://stuk.github.io/jszip/) (MIT) und
  [fast-xml-parser](https://github.com/NaturalIntelligence/fast-xml-parser)
  (MIT) – die OOXML-Container- und XML-Schichten.
- [Fluent UI System Icons](https://github.com/microsoft/fluentui-system-icons)
  (MIT) – das Icon-Set in den Menübändern.
- [electron-updater](https://www.electron.build/) (MIT) – In-App-Updates.
- Die Schriften Liberation, Carlito, Caladea und Noto CJK (OFL/Apache-2.0) –
  mitgelieferte Dokumentschriften.

`npm run notices` erzeugt die gebündelte Zusammenfassung der
Drittanbieter-Lizenzen neu (`tools/gen-third-party-notices.mjs`); alle
Laufzeitabhängigkeiten stehen unter MIT/Apache-2.0/BSD-3-Clause/OFL.

## Lizenz

GenOffice steht unter der [Apache License 2.0](../../LICENSE), mit einer
Ausnahme: Das Verzeichnis `ee/` ist für zukünftige Enterprise-Module
reserviert und unterliegt der
[GenOffice Enterprise License](../../ee/LICENSE).

Die Namen und Logos GenOffice und Genspark sind Marken von Mainfunc, Inc.
Die Apache-2.0-Lizenz gewährt keine Berechtigung zu ihrer Nutzung (siehe
Abschnitt 6); Forks sollten ihr eigenes Branding verwenden.
