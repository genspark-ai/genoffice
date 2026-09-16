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
  <a href="#command-line-and-agent-skill"><b>CLI</b></a> ·
  <a href="#mcp-server"><b>MCP</b></a> ·
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
  OpenRouter, Requesty oder jeden OpenAI-kompatiblen Endpunkt mit – auch lokale Server
  werden unterstützt.
- **Skriptbar und bereit für Agenten.** Die App bringt eine
  `genoffice`-Befehlszeile und einen Skill für Claude Code, Codex, Cursor,
  Gemini CLI, GitHub Copilot, OpenCode und Windsurf mit, sodass ein
  Coding-Agent echte Office-Dateien auf deinem Rechner erstellen,
  konvertieren, lesen und bearbeiten kann, ohne ein Fenster zu öffnen.

**Los geht's:** [macOS](https://github.com/genspark-ai/genoffice/releases/latest) (Apple Silicon und Intel) ·
[Windows](https://github.com/genspark-ai/genoffice/releases/latest) (x64 und Arm) ·
[Linux](https://github.com/genspark-ai/genoffice/releases/latest) (deb, rpm, AppImage) —
Details und Voraussetzungen unter [Download](#download).

## Demo

Sechs Apps, ein KI-Panel und eine Befehlszeile für deinen Coding-Agenten.
Jeder Screenshot zeigt die echte App auf macOS, wobei die KI durch den
Prompt gesteuert wird, den du im Panel nachlesen kannst.

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

### 7 · CLI – dein Coding-Agent steuert GenOffice, auf deinem Rechner

GenOffice bringt eine `genoffice`-Befehlszeile und einen Agent-Skill mit.
Installiere den Skill, und Claude Code, Codex, Cursor, Gemini CLI, GitHub
Copilot, OpenCode oder Windsurf können echte Office-Dateien über dieselben
Engines wie die Apps erstellen, konvertieren, lesen und bearbeiten, ohne ein
Fenster zu öffnen.

<img src="../assets/readme/cli-deck-in-app.webp" alt="GenOffice Slides zeigt eine Sonnensystem-Präsentation mit acht Folien, die ein Coding-Agent über die genoffice-Befehlszeile erstellt hat: die Titelfolie auf der Arbeitsfläche, acht Miniaturansichten links und das geöffnete KI-Panel" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/cli-slides-grid.webp" alt="Die acht gerenderten Folien der Sonnensystem-Präsentation nebeneinander: Titel, Zeitleiste der Erforschung, vier Kennzahlen, Balkendiagramm der Planetendurchmesser, Gesteinsplaneten gegen Riesen, die 99,8 % der Sonne als Hero-Zahl, das Raster der vier Riesen und die Kernaussagen"></td>
<td width="50%"><img src="../assets/readme/cli-integrations.webp" alt="GenOffice-Einstellungen, Seite Integrationen: der genoffice-Skill ist in Claude Code installiert, mit Installieren-Schaltflächen neben Codex und Cursor"></td>
</tr>
<tr>
<td><b>Ein Prompt an deinen Agenten</b> — „Erstelle eine Präsentation mit acht Folien über das Sonnensystem.“ Der Agent liest den Skill, schreibt ein Stylesheet, eine Gliederung und eine Seitenspezifikation pro Folie, generiert die beiden Fotos mit <code>genoffice image</code> und lässt <code>genoffice slides check</code> alles zurückweisen, was überläuft oder sich überlappt, bevor <code>genoffice create</code> die <code>.pptx</code> zusammensetzt und <code>slides render</code> pro Folie ein PNG zum Anschauen zurückgibt.</td>
<td><b>Einmal installieren, unter Einstellungen → Integrationen</b> — GenOffice listet die Coding-Agenten auf, die es auf diesem Rechner findet, und schreibt den Skill in jeden, den du auswählst. Oder lade den Skill als Zip herunter oder führe <code>npx skills add genspark-ai/genoffice</code> aus. Befehle und der vollständige Workflow stehen unter <a href="#command-line-and-agent-skill">Befehlszeile und Agent-Skill</a>.</td>
</tr>
</table>

### 8 · MCP – dieselben Tools über das Model Context Protocol

Jeder `genoffice`-Befehl ist auch ein MCP-Tool. Claude Code, Claude Desktop,
Cursor und jeder andere MCP-Client können `genoffice mcp` selbst starten,
ganz ohne Skill-Installation und ohne geöffnetes Fenster, und erhalten 29
Tools plus die Op-Referenzen als Ressourcen. Ein zweiter HTTP-Server in der
App lässt einen Agenten ein Word-Dokument in einem sichtbaren Editor-Tab
aufbauen, während du zusiehst.

<img src="../assets/readme/mcp-deck-motion.webp" alt="Zeitraffer von Claude Code beim Bau einer achtseitigen Investoren-Präsentation über erneuerbare Energien über den genoffice-MCP-Server: sucht nach Abbildungen und Fotos, prüft jedes Kandidatenbild mit media, deck_start schreibt das Stylesheet und die Gliederung, deck_page fügt Seite für Seite eine geprüfte Seite hinzu, deck_build stellt die .pptx zusammen und slides_render liefert ein Bild jeder Folie zurück; die fertige Präsentation öffnet sich danach in GenOffice Slides" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/mcp-deck-in-app.webp" alt="GenOffice Slides zeigt die achtseitige Präsentation Renewable Energy 2026, die Claude Code über den genoffice-MCP-Server erstellt hat: die Titelfolie mit einem Foto eines Windparks auf der Arbeitsfläche und acht Miniaturansichten links"></td>
<td width="50%"><img src="../assets/readme/mcp-integrations.webp" alt="GenOffice-Einstellungen, Seite Integrationen, MCP-Teil: der einzeilige claude mcp add-Befehl für Claude Code, der JSON-Block für Cursor, Claude Desktop und andere MCP-Clients, und die Option für den lokalen HTTP-Server darunter"></td>
</tr>
<tr>
<td><b>Ein Prompt, 38 Tool-Aufrufe, kein Terminal</b> — „Erstelle eine achtseitige Investoren-Präsentation über erneuerbare Energien im Jahr 2026, mit einem echten Foto auf dem Titelblatt und überall sonst, wo ein Foto hilft.“ Der Agent holt die Abbildungen und die Fotos mit <code>search</code>, fragt <code>media</code>, ob jedes Kandidatenbild eine echte Fotografie ist, ruft <code>deck_start</code> mit einem Stylesheet und einer Gliederung auf, dann <code>deck_page</code> einmal pro Folie; jede Seite wird gegen die Gliederung und die Palette geprüft, bevor sie behalten wird, <code>deck_build</code> stellt die <code>.pptx</code> zusammen, <code>slides_audit</code> sucht nach Überlauf, <code>slides_render</code> liefert pro Folie ein PNG als Bildinhalt zurück, den das Modell ansehen kann, und <code>deck_replace</code> überarbeitet die drei Seiten, die ihm nicht gefallen haben.</td>
<td><b>Einmal verbinden, unter Einstellungen → Integrationen</b> — kopiere die <code>claude mcp add</code>-Zeile für Claude Code oder den JSON-Block in Cursor, Claude Desktop oder einen anderen MCP-Client. Option B schaltet den lokalen HTTP-Server für den sichtbaren Word-Editor ein. Beides wird unter <a href="#mcp-server">MCP-Server</a> beschrieben.</td>
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
- **Skriptbar.** Eine `genoffice`-Befehlszeile, ein Agent-Skill und ein
  MCP-Server stellen jede Engine in den Dienst von Claude Code, Claude
  Desktop, Codex, Cursor und anderen Agenten – weiterhin auf dem Gerät.
- **Kostenlos**, für Einzelpersonen wie für Teams.

## KI-Backends

**Mit Genspark anmelden**, und es gibt nichts zu konfigurieren:
Modellaufrufe laufen über den Genspark-Proxy (die Modellfamilien Claude, GPT
und Gemini), und die Agenten erhalten Web- und Bildsuche, Bildgenerierung
sowie Bild-/Audio-/Videoanalyse.

**Oder bring deinen eigenen Schlüssel mit.** Unter Einstellungen → KI stehen
Claude, OpenAI, Gemini, DeepSeek, Kimi, GLM, Qwen, Doubao, MiniMax, Grok,
Mistral, OpenRouter, Requesty und OpenCode Zen/Go zur Wahl, plus ein freier Slot für
jeden OpenAI-kompatiblen Endpunkt (Basis-URL + Schlüssel), einschließlich
lokaler Modell-Server. Suche und Medien haben eigene Anbieter je Fähigkeit
unter KI-Medien & Suche: Serper oder Tavily für die Websuche und OpenAI,
Gemini, Doubao/Seedream, GLM, Grok, Qwen, MiniMax oder jeder
OpenAI-kompatible Bild-Endpunkt für Bildgenerierung und Bild-/Videoanalyse.

Die gesamte Suite bringt ein helles, ein dunkles und einen
Systemdesign-Modus mit. Designs ändern nur, was auf dem Bildschirm zu sehen
ist: Exporte, Ausdrucke und gespeicherte Dateien behalten immer die eigenen
Farben des Dokuments.

<a id="command-line-and-agent-skill"></a>

## Befehlszeile und Agent-Skill

Alles, was die Apps mit einer Datei tun können, kann die
`genoffice`-Befehlszeile aus einem Terminal heraus: Word, Excel, PowerPoint,
PDF, Markdown und HTML inspizieren, konvertieren, erstellen, lesen und
bearbeiten – auf denselben Engines, ohne Fenster. Sie wird mit GenOffice
installiert, braucht keine eigene Laufzeitumgebung und schickt nie ein
Dokument irgendwohin. Zusammen mit dem mitgelieferten **Agent-Skill** macht
sie aus einem Coding-Agenten einen Dokumentenarbeiter, der echte
Office-Dateien statt Markdown-Näherungen liefert.

**Funktioniert mit:** Claude Code, Codex, Cursor, Gemini CLI, GitHub Copilot,
OpenCode und Windsurf von Haus aus, mit jedem anderen Agenten, der Skills
liest, sowie über den [MCP-Server](#mcp-server) mit Claude Desktop und jedem
MCP-Client.

### Skill installieren

| Wie                                           | Was passiert                                                                                                                                                                                                             |
| --------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| **Einstellungen → Integrationen** in der App  | Listet die auf diesem Rechner gefundenen Agenten auf; ein Klick schreibt den Skill in jeden, den du auswählst. Eine **Aktualisieren**-Schaltfläche erscheint, wenn eine GenOffice-Version einen neueren Skill mitbringt. |
| **Als Zip herunterladen** auf derselben Seite | Das Layout, das claude.ai, die Claude-Desktop-Apps und andere Assistenten als hochgeladenen Skill akzeptieren.                                                                                                           |
| `npx skills add genspark-ai/genoffice`        | Installiert aus diesem Repository in jeden Skills-kompatiblen Agenten.                                                                                                                                                   |

Starte dann einen neuen Chat und bitte um ein Dokument. Der Skill bringt dem
Agenten bei, wann er zu `genoffice` greifen soll, wie er eine Datei vor dem
Bearbeiten liest und wie er seine eigene Arbeit prüft.

### Schnellstart im Terminal

```bash
genoffice --version
genoffice info report.docx --json                  # headings and blocks; or sheets, slides, pages
genoffice convert report.md --to pdf               # md/html/docx/xlsx/pptx → pdf, pdf → docx/xlsx/pptx, …
genoffice create --type docx --from notes.md --out notes.docx
genoffice create --type xlsx --from table.json --out sales.xlsx   # "=SUM(B2:B9)" cells stay live formulas
genoffice docs read report.docx --range 0-9 --json # then `docs apply --ops edits.json` edits in place
genoffice render report.docx --out shots/          # one PNG per page, to look at what you made
genoffice open sales.xlsx                          # hand the result to the editor
```

Jeder Befehl gibt eine einzeilige Zusammenfassung aus, oder mit `--json` ein
einzelnes JSON-Objekt. Bearbeitungen sind atomar: Eine zurückgewiesene
Operation lässt die Datei unangetastet und kommt mit einer erklärenden
Fehlermeldung zurück. `genoffice help` listet die aktuellen Befehle auf; die
vollständige Referenz steht in
[packages/cli/README.md](../../packages/cli/README.md).

### Was der Agent tatsächlich ausführt

Die Sonnensystem-Präsentation in der Demo oben brauchte einen einzigen Prompt
in Claude Code. Dahinter folgte der Agent dem gestuften Workflow des Skills,
und die CLI prüfte jede Stufe, bevor die nächste begann:

```bash
genoffice capabilities --json                        # which cloud tools GenOffice has configured
genoffice guide slides design                        # the deck workflow and layout library
genoffice image "the eight planets in a row …" --aspect 16:9 --out deck/assets/cover.jpg
genoffice slides check deck/outline.json --json      # 8 pages, no findings
genoffice slides check deck/pages/01.json --json     # builds one slide, audits overflow and overlap
…                                                    # one page file per slide, fixed until each check is clean
genoffice create --type pptx --spec deck/pages --outline deck/outline.json --out deck/solar-system.pptx --json
genoffice slides render deck/solar-system.pptx --out deck/shots --json
genoffice slides audit deck/solar-system.pptx --json    # 8 slides, no layout issues
genoffice slides replace deck/solar-system.pptx --slide 4 --spec deck/pages/05.json --json
genoffice open deck/solar-system.pptx
```

Innerhalb von `genoffice` findet kein Modellaufruf statt: Der Agent denkt,
die CLI baut und prüft, und das Ergebnis öffnet sich in GenOffice oder
PowerPoint als gewöhnliche `.pptx`.

<a id="mcp-server"></a>

### MCP-Server

Dieselben Befehle stehen auch als
[Model-Context-Protocol](https://modelcontextprotocol.io)-Tools zur
Verfügung, für Assistenten, die kein Terminal ausführen können oder denen du
lieber keins gibst. Es gibt zwei Wege hinein, beide mit kopierfertigen
Snippets unter **Einstellungen → Integrationen → MCP**:

| Weg                                 | Was es ist                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 |
| ----------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **A · `genoffice mcp`** (empfohlen) | Ein Stdio-Server, den der Assistent selbst startet; GenOffice muss dafür nicht geöffnet sein. Ein Tool pro Befehl (`info`, `convert`, `create_docx`, `create_xlsx`, `create_pptx`, `create_pdf`, `docs_read` / `docs_apply` / `docs_check`, `sheet_*`, `slides_*`, `render`, `guide`, `search`, `image`, `media`, `open`) plus der gestufte Deck-Ablauf `deck_start` → `deck_page` → `deck_build` → `deck_replace`. Ops, Specs und Markdown werden inline übergeben, sodass auch ein Client ohne Dateisystem funktioniert. |
| **B · Lokaler HTTP-Server**         | Läuft innerhalb der GenOffice-App unter `http://127.0.0.1:3093/mcp` (Streamable HTTP, mit Legacy-SSE). Seine Tools steuern einen sichtbaren Word-Editor-Tab: `create_session`, `insert_content`, `replace_blocks`, `apply_ops`, `read_document`, `save_session`, und du siehst zu, wie das Dokument entsteht. Standardmäßig aus; im selben Einstellungsbereich einschalten.                                                                                                                                                |

```bash
# Claude Code
claude mcp add --transport stdio genoffice -- genoffice mcp
```

```jsonc
// Cursor, Claude Desktop oder ein anderer MCP-Client
{ "mcpServers": { "genoffice": { "command": "genoffice", "args": ["mcp"] } } }
```

`genoffice` ist hier die CLI, die in der App enthalten ist (unter macOS
`/Applications/GenOffice.app/Contents/Resources/cli/genoffice`; der
Einstellungsbereich zeigt den genauen Pfad für deine Installation an). Der
Server bringt seine eigenen Workflow-Anweisungen mit und stellt die
Op-Referenzen als `genoffice://guide/*`-Ressourcen bereit, sodass kein
Skill nötig ist; der Skill und der MCP-Server können nebeneinander
bestehen, und der Assistent wählt einen davon. Cloud-Funktionen (`search`,
`image`, `media`) laufen weiterhin über den in GenOffice konfigurierten
Anbieter; alles andere läuft lokal, und `GENOFFICE_ALLOWED_ROOTS`
beschränkt jedes Tool auf die Ordner, die du angibst.

Die Präsentation über erneuerbare Energien aus der Demo oben zeigt, wie ein
einzelner Prompt in Claude Code mit nur dem `genoffice`-MCP-Server auf
Protokollebene aussieht:

```text
capabilities · guide(slides, spec) · guide(slides, design)
search(query) ×4                         → IEA, BNEF and IRENA figures for the slides
search(query, images) ×7 · media(url, ask) ×7
                                         → candidate photos, each one checked to be a real photograph
deck_start(dir, style, outline)          → outline checked: 8 pages to write
deck_page(dir, 0, page) … deck_page(dir, 7, page)
                                         → each page checked against the outline and the palette; one page sent again
deck_build(dir, out)                     → renewables-2026.pptx, no image failures
slides_audit(file) · slides_render(file, out)
                                         → no layout findings; 8 PNGs come back as image content
deck_replace(dir, n, page) ×3 · slides_render(file, out)
                                         → three pages fixed after looking at the renders
```

Achtunddreißig Aufrufe, etwa dreizehn Minuten, und der Assistent hat nie
ein Terminal angefasst: Die Abbildungen, die Fotos, die Guides, die
Prüfungen und die Renderings liefen alle als MCP-Tool-Ergebnisse. Nur
`search` und `media` haben die Maschine verlassen, zum in GenOffice
konfigurierten Provider.

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
GLM, Qwen, Doubao, MiniMax, Grok, Mistral, OpenRouter, Requesty und OpenCode Zen/Go,
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
<summary><b>Kann ich GenOffice aus Claude Code, Codex, Cursor oder einem Skript steuern?</b></summary>

Ja. GenOffice installiert eine `genoffice`-Befehlszeile, die dieselben
Engines ohne Fenster ausführt: Dokumente aus einem Terminal oder einem Skript
inspizieren, konvertieren, erstellen, lesen und bearbeiten, mit
`--json`-Ausgabe für Programme. Der mitgelieferte Agent-Skill bringt Claude
Code, Codex, Cursor, Gemini CLI, GitHub Copilot, OpenCode und Windsurf bei,
sie zu nutzen; installiere ihn unter **Einstellungen → Integrationen**. Siehe
[Befehlszeile und Agent-Skill](#command-line-and-agent-skill).

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
