<p align="center">
  <a href="https://genoffice.ai/">
    <picture>
      <source srcset="../assets/readme/hero-dark.webp" media="(prefers-color-scheme: dark)">
      <img src="../assets/readme/hero.webp" alt="GenOffice — het open-source AI-kantoorpakket: Docs, Sheets, Slides, PDF, Markdown en HTML met een ingebouwd AI-paneel" width="100%">
    </picture>
  </a>
</p>

<h1 align="center">GenOffice</h1>

<p align="center"><b>De eerste volwaardige open-source AI-officesuite ter wereld.</b><br>
Word-, Excel-, PowerPoint- en PDF-bestanden, bewerkt door jou en je AI, opgeslagen in de echte formaten.</p>

<p align="center">
  <a href="../../LICENSE"><img src="https://img.shields.io/github/license/genspark-ai/genoffice" alt="Licentie: Apache-2.0"></a>
  <a href="https://github.com/genspark-ai/genoffice/releases/latest"><img src="https://img.shields.io/github/v/release/genspark-ai/genoffice" alt="Nieuwste release"></a>
  <a href="https://github.com/genspark-ai/genoffice/releases"><img src="https://img.shields.io/github/downloads/genspark-ai/genoffice/total" alt="Downloads"></a>
  <a href="https://github.com/genspark-ai/genoffice/stargazers"><img src="https://img.shields.io/github/stars/genspark-ai/genoffice?style=flat" alt="GitHub-sterren"></a>
</p>

<p align="center"><a href="../../README.md">English</a> · <a href="README.es.md">Español</a> · <a href="README.pt-BR.md">Português (Brasil)</a> · <a href="README.de.md">Deutsch</a> · <a href="README.fr.md">Français</a> · <a href="README.zh-CN.md">简体中文</a> · <a href="README.zh-TW.md">繁體中文</a> · <a href="README.ko.md">한국어</a> · <a href="README.ja.md">日本語</a> · <a href="README.ar.md">العربية</a> · <a href="README.ru.md">Русский</a> · <a href="README.it.md">Italiano</a> · <b>Nederlands</b> · <a href="README.pl.md">Polski</a> · <a href="README.cs.md">Čeština</a> · <a href="README.id.md">Bahasa Indonesia</a> · <a href="README.ms.md">Bahasa Melayu</a> · <a href="README.th.md">ไทย</a> · <a href="README.hi.md">हिन्दी</a> · <a href="README.he.md">עברית</a></p>

<p align="center">
  <a href="#download"><b>Download</b></a> ·
  <a href="https://genoffice.ai/"><b>Website</b></a> ·
  <a href="https://genoffice.ai/join"><b>Community</b></a> ·
  <a href="../../PRIVACY.md"><b>Privacy</b></a>
</p>

GenOffice is een gratis, open-source alternatief voor Microsoft Office voor
macOS, Windows en Linux. Het opent en bewaart native `.docx`-, `.xlsx`- en
`.pptx`-bestanden, bewerkt PDF, Markdown en HTML, en zet bij elk document een
AI-agent — geen los chatvenstertje erbij geplakt, maar een editor die het
bestand leest, de wijziging doorvoert en je precies laat zien wat er is
aangepast.

- **Echte formaten, byte-behoudend.** Alleen wat je bewerkt, wordt herschreven.
  Al het overige in het bestand blijft byte-voor-byte behouden, zodat
  documenten gewoon blijven werken in Word, Excel en PowerPoint.
- **AI die je kunt controleren.** Wijzigingen komen binnen als tracked changes
  en diffs met terugdraaien in één klik. Spreadsheets krijgen live formules,
  geen geplakte getallen. Decks en pagina's worden op het canvas gegenereerd
  en blijven volledig bewerkbaar.
- **Lokaal by design.** Bestanden worden op je eigen machine geopend, bewerkt,
  opgeslagen en geconverteerd. PDF → Word / Excel / PowerPoint, Markdown →
  Word en HTML → Word draaien allemaal on-device. Alleen de AI-aanroepen
  verlaten de machine, naar de provider die jij kiest.
- **Jouw eigen keys, of geen enkele.** Log in met Genspark en je hoeft niets te
  regelen, of gebruik je eigen key voor Claude, OpenAI, Gemini, DeepSeek,
  Kimi, GLM, Qwen, Doubao, MiniMax, Grok, Mistral, OpenRouter, API Route, Requesty, of elk
  OpenAI-compatibel endpoint, inclusief lokale servers.

**Haal het hier:** [macOS](https://github.com/genspark-ai/genoffice/releases/latest) (Apple Silicon en Intel) ·
[Windows](https://github.com/genspark-ai/genoffice/releases/latest) (x64 en Arm) ·
[Linux](https://github.com/genspark-ai/genoffice/releases/latest) (deb, rpm, AppImage) —
details en vereisten in [Download](#download).

## Demo

Zes apps, één AI-paneel. Elke screenshot toont de echte app op macOS, met de
AI aangestuurd vanuit de prompt die je in het paneel kunt lezen.

### 1 · Docs — open en bewerk `.docx` met een AI die je kunt controleren

<table>
<tr>
<td width="50%"><img src="../assets/readme/docs-report.webp" alt="GenOffice Docs toont een pagina van een jaarverslag in twee kolommen met een cover-afbeelding over de volledige breedte, een KPI-tabel met arcering, koptekst en voettekst, op 80% zoom met het AI-paneel ingeklapt"></td>
<td width="50%"><img src="../assets/readme/docs-ai.webp" alt="GenOffice Docs: een bedrijfsoverzicht met een bannerafbeelding; de AI heeft de Overview compacter gemaakt en een nieuwe sectie met opsommingstekens toegevoegd, en het paneel biedt een terugdraaifunctie met één klik"></td>
</tr>
<tr>
<td><b>Opent het bestand precies zoals Word het opmaakt</b> — secties in twee kolommen, beeldvullende afbeeldingen, tabellen met arcering, koppen en voetteksten, paginering op basis van Words regelmetriek. Stijlen, opmerkingen, tracked changes, vergelijkingen en inkt komen ongewijzigd terug.</td>
<td><b>Vraag om de wijziging</b> — de AI leest de blokken die het nodig heeft, herschrijft de Overview en voegt een nieuwe sectie met opsommingstekens toe. Elke AI-beurt is een snapshot die je kunt terugdraaien; met <b>Track changes</b> ingeschakeld komen wijzigingen binnen als revisies in Word-stijl.</td>
</tr>
</table>

### 2 · Sheets — `.xlsx` met live formules en grafieken, geen geplakte getallen

<table>
<tr>
<td width="50%"><img src="../assets/readme/sheets-ai.webp" alt="GenOffice Sheets: de AI heeft een Summary-blad toegevoegd met omzet per regio en categorie via SUMIF-formules, plus een kolomdiagram, en meldt 43 toegepaste wijzigingen met een knop om ongedaan te maken"></td>
<td width="50%"><img src="../assets/readme/sheets-qa.webp" alt="GenOffice Sheets: gevraagd welke regio de omzet in Q2 aanvoerde, antwoordt de AI Europa met de uitsplitsing per categorie en verwijst als links naar de cellen die het heeft gebruikt, naast het Orders-blad"></td>
</tr>
<tr>
<td><b>Bouw het</b> — met één zin voegt de agent een Summary-blad toe met echte <code>SUMIF</code>-formules per regio en categorie, voegt een kolomdiagram in, en past de 43 wijzigingen toe als één batch die je in één keer ongedaan kunt maken.</td>
<td><b>Stel een vraag</b> — vragen over de werkmap komen terug met de redenering en de exacte cellen als klikbare verwijzingen. Onder de motorkap: een eigen Rust-<code>.xlsx</code>-engine, draaitabellen, slicers, voorwaardelijke opmaak en formule-traceren.</td>
</tr>
</table>

### 3 · Slides — van een prompt naar een `.pptx`-deck

<img src="../assets/readme/slides-generate.webp" alt="Timelapse van GenOffice Slides dat de Aurora Home investeerdersdeck genereert: de AI plant het verhaal in het paneel, dia's verschijnen één voor één op het canvas, en de voltooide deck eindigt met de afsluitende vraag" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/slides-cover.webp" alt="GenOffice Slides: de openingsslide van een door AI gegenereerd investeerdersdeck voor Aurora Home op het canvas, met de oorspronkelijke prompt van één regel en de samenvatting van de AI over wat het heeft gebouwd, in het paneel"></td>
<td width="50%"><img src="../assets/readme/slides-ai.webp" alt="GenOffice Slides: de vormgegeven slotslide van hetzelfde deck van 11 slides, met de miniaturenstrook links en het AI-paneel dat de verhaallijn samenvat"></td>
</tr>
<tr>
<td><b>Eén regel als input</b> — "Create a 10-slide investor pitch deck for Aurora Home…". GenOffice bepaalt de verhaallijn, zoekt de cijfers uit en zet elke slide als een echte <code>.pptx</code> op het canvas.</td>
<td><b>Een afgewerkt deck als resultaat</b> — elf vormgegeven slides met consistente typografie, beeldmateriaal en een afsluitende call-to-action; blijf bewerken met masters, layouts, slimme hulplijnen en niet-destructief bijsnijden, of laat het paneel de stijl, tekst en volgorde aanpassen.</td>
</tr>
</table>

### 4 · PDF — bewerk PDF-tekst direct op de pagina, converteer PDF naar Word on-device

<table>
<tr>
<td width="50%"><img src="../assets/readme/pdf-edit.webp" alt="GenOffice PDF: Edit text mode omkadert elk tekstblok op de pagina voor bewerking op de plek zelf, terwijl het AI-paneel een vraag over het rapport beantwoordt met paginaverwijzingen"></td>
<td width="50%"><img src="../assets/readme/pdf-convert.webp" alt="GenOffice Docs toont een Word-document dat lokaal is geconverteerd vanuit de kwartaalrapportage-PDF van Helios, geopend in een tweede tab naast de originele PDF"></td>
</tr>
<tr>
<td><b>Bewerk binnen de pagina</b> — Edit text mode omkadert elk tekstblok voor herschrijven op de plek zelf; de content stream wordt via PDFium herschreven met de originele fonts, geen overplaknotitie. Stel de AI een vraag over een lang rapport en krijg antwoorden met paginaverwijzingen.</td>
<td><b>Converteer on-device</b> — <b>PDF Converter → PDF to Word</b> levert een bewerkbare <code>.docx</code> op die naast de bron in Docs opent, met koppen, cijferrijen en alinea's intact. Excel- en PowerPoint-doelen werken op dezelfde manier; gescande pagina's gaan door de systeem-OCR.</td>
</tr>
</table>

### 5 · HTML — een AI-pagina- en UI-builder, eerst een design brief

Vertel waarvoor de pagina dient en voor wie hij bedoeld is. De AI stelt eerst
een **design brief** voor — hook, kleurenpalet, typografie en stijlrichtingen
— en bouwt daarna op basis van die tokens één zelfstandig `.html`-bestand.

<img src="../assets/readme/html-restyle-motion.webp" alt="Timelapse van GenOffice HTML dat de Lumen landingspagina van stijl verandert: één Restyle-verzoek in het paneel verandert de donkere Midnight Studio-pagina in de warme Solar Daybreak-versie, terwijl elke sectie en alle inhoud op hun plek blijven" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/html-ai.webp" alt="GenOffice HTML: een gegenereerde landingspagina voor een zonne-energie bureaulamp in de donkere stijlrichting Midnight Studio, getoond in de live preview met het AI-paneel dat de net gebouwde pagina samenvat"></td>
<td width="50%"><img src="../assets/readme/html-restyle.webp" alt="Dezelfde Lumen-landingspagina, door de AI omgezet naar de warme stijlrichting Solar Daybreak: papieren achtergrond, serif-koppen en een oranje accent, met behoud van elke sectie en alle tekst"></td>
</tr>
<tr>
<td><b>Gegenereerd vanuit één prompt</b> — een gedurfde hero, featurekaarten, prijzen en een wachtlijstformulier voor Lumen, gebouwd in de stijlrichting Midnight Studio. Klik op elk element om de stijl aan te passen, dubbelklik om tekst te bewerken, of schakel over naar de CodeMirror-broncodeweergave.</td>
<td><b>Zelfde ontwerp, nieuwe richting</b> — met één <b>Restyle</b>-verzoek wisselen de tokens van de brief en past de pagina zich aan: warm papier, redactionele serif, zonoranje accent, zonder dat er iets wordt herschreven. Presenteer in volledig scherm, of exporteer als PDF of een native bewerkbaar Word-document.</td>
</tr>
</table>
<table>
<tr>
<td width="50%"><img src="../assets/readme/html-dashboard.webp" alt="GenOffice HTML: een gegenereerde persoonlijke dashboard-UI voor een freelance designer in een warme linnen stijl, met een navigatiebalk links, een begroeting in serif en vier metriekkaarten"></td>
<td width="50%"><img src="../assets/readme/html-report.webp" alt="GenOffice HTML: een gegenereerd datarapport over de EV-markt in broadsheet-stijl, met een kopregel in serif, een hoofdcijfer van 17,3 miljoen en een rij met kerncijfers"></td>
</tr>
<tr>
<td><b>UI-mockups</b> — de starter "personal dashboard" zet een persona om in een werkende layout: navigatiebalk links, begroeting, sparkline voor factureerbare uren, facturen- en bezettingskaarten, allemaal echte HTML die je zo aan een developer kunt overdragen.</td>
<td><b>Datastories</b> — de starter "data report" bouwt een redactioneel broadsheet: kopregel in serif, één hoofdcijfer, een met lijnen afgescheiden rij kerncijfers, inline SVG-grafieken en een methodologienoot.</td>
</tr>
</table>

### 6 · Markdown — een blokeditor over platte `.md`, met Ask AI

<table>
<tr>
<td width="50%"><img src="../assets/readme/markdown-ai.webp" alt="GenOffice Markdown: bij een geselecteerde alinea verschijnt een Ask AI-popover met een getypte instructie en suggestiechips zoals Polish, Make more concise, Expand en Fix grammar, plus de knoppen Send now en Add to queue"></td>
<td width="50%"><img src="../assets/readme/markdown-render.webp" alt="GenOffice Markdown toont een document met lanceringsnotities met een tabel, een Mermaid-stroomdiagram en een taaklijst, met de startprompts van het AI-paneel aan de linkerkant"></td>
</tr>
<tr>
<td><b>Ask AI over een selectie</b> — selecteer een willekeurige passage en er verschijnt een <b>Ask AI</b>-chip: typ een instructie of kies een suggestie, verstuur die direct, of verzamel meerdere aan tekst verankerde wijzigingen en voer ze in één keer uit. Dezelfde toegang bestaat in elke app.</td>
<td><b>Weergegeven, opgeslagen als platte Markdown</b> — koppen, lijsten, tabellen, afbeeldingen, codeblokken en Mermaid-diagrammen in een Tiptap-blokeditor, terugweggeschreven als platte <code>.md</code>, met een volledig lokale <b>Markdown → Word</b>-export.</td>
</tr>
</table>

## Waarom GenOffice

- **Open source**, Apache-2.0, in het openbaar ontwikkeld op GitHub.
- **Van jou om te draaien.** Native apps voor macOS, Windows en Linux;
  bestanden blijven op je eigen schijf en elke bewerking, opslag en conversie
  gebeurt op je eigen machine.
- **Echte Office-bestanden.** Native `.docx`, `.xlsx` en `.pptx`,
  byte-behoudend: de delen van een bestand die je niet hebt aangeraakt, worden
  precies zo gekopieerd als ze waren.
- **Een AI die het document zelf bewerkt.** Tracked changes in Docs, live
  formules en grafieken in Sheets, slides die direct op het canvas worden
  getekend, elke AI-beurt een snapshot waarnaar je kunt teruggaan.
- **Je eigen model, je eigen key.** Log in met Genspark, of gebruik een key
  voor Claude, OpenAI, Gemini, DeepSeek en meer, inclusief lokale servers en
  elk OpenAI-compatibel endpoint.
- **PDF goed gedaan.** Bewerk tekst direct op de pagina en converteer PDF
  naar Word, Excel of PowerPoint on-device, met systeem-OCR voor scans.
- **Ook Markdown en HTML**, met hetzelfde AI-paneel en lokale export naar
  Word.
- **Gratis**, voor individuen en teams.

## AI-backends

**Log in met Genspark** en er is niets te configureren: modelaanroepen lopen
via de Genspark-proxy (de Claude-, GPT- en Gemini-families) en de agents
krijgen toegang tot web- en beeldzoeken, beeldgeneratie en analyse van beeld,
audio en video.

**Of gebruik je eigen key.** Onder Settings → AI vind je Claude, OpenAI,
Gemini, DeepSeek, Kimi, GLM, Qwen, Doubao, MiniMax, Grok, Mistral, OpenRouter, API Route, Requesty
en OpenCode Zen/Go, plus een aangepast slot voor elk OpenAI-compatibel
endpoint (base URL + key), inclusief lokale modelservers. Zoeken en media
hebben elk hun eigen providers per capability onder **AI Media & Search**:
Serper of Tavily voor webzoeken, en OpenAI, Gemini, Doubao/Seedream, GLM,
Grok, Qwen, MiniMax of elk OpenAI-compatibel images-endpoint voor
beeldgeneratie en analyse van beeld en video.

De hele suite wordt geleverd met een licht, donker en systeemthema. Thema's
veranderen alleen wat er op je scherm te zien is: exports, afdrukken en
opgeslagen bestanden behouden altijd de eigen kleuren van het document.

<a id="download"></a>

## Download

| Platform                             | Vereisten                                               | Download                                                                                    |
| ------------------------------------ | ------------------------------------------------------- | ------------------------------------------------------------------------------------------- |
| **macOS** — Apple Silicon (arm64)    | macOS 11+                                               | [Nieuwste `.dmg` (arm64)](https://github.com/genspark-ai/genoffice/releases/latest)         |
| **macOS** — Intel (x64)              | macOS 11+                                               | [Nieuwste `.dmg` (x64)](https://github.com/genspark-ai/genoffice/releases/latest)           |
| **Windows** (x64, de meeste pc's)    | Windows 10+, Intel/AMD                                  | [Nieuwste `-x64.exe`-installer](https://github.com/genspark-ai/genoffice/releases/latest)   |
| **Windows** op Arm (ARM64)           | Windows 11 op Arm (Snapdragon X en vergelijkbare chips) | [Nieuwste `-arm64.exe`-installer](https://github.com/genspark-ai/genoffice/releases/latest) |
| **Linux** — Debian / Ubuntu          | x86_64, glibc 2.34+ (Ubuntu 22.04 of nieuwer)           | [Nieuwste `.deb`](https://github.com/genspark-ai/genoffice/releases/latest)                 |
| **Linux** — Fedora / RHEL / openSUSE | x86_64, glibc 2.34+ (Fedora 35+, RHEL 9+, Leap 15.6+)   | [Nieuwste `.rpm`](https://github.com/genspark-ai/genoffice/releases/latest)                 |
| **Linux** — overige distributies     | x86_64, glibc 2.34+, FUSE 2                             | [Nieuwste `.AppImage`](https://github.com/genspark-ai/genoffice/releases/latest)            |

Alle builds komen van `main`; de macOS- en Windows-installers zijn gesigneerd.
Oudere versies staan op de pagina [Releases](https://github.com/genspark-ai/genoffice/releases).

<details>
<summary><b>Installeren op Linux</b></summary>

De deb installeer je met apt — dat haalt de afhankelijkheden op en voegt
GenOffice toe aan het toepassingenmenu:

```bash
sudo apt install ./genoffice_<version>_amd64.deb
```

Op Fedora, de RHEL-familie of openSUSE installeer je in plaats daarvan de rpm:

```bash
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / RHEL family
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

De AppImage draait direct: installeer de FUSE 2-runtime
(`sudo apt install libfuse2`; op Ubuntu 24.04 heet het pakket `libfuse2t64`),
maak het bestand uitvoerbaar en start het:

```bash
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

</details>

## Hoe het werkt

Zeven Electron-apps — Docs, Sheets, Slides, PDF, Markdown, HTML en de shell
met tabbladen — delen één engine-laag van pure TypeScript-packages plus een
Rust-sidecar voor `.xlsx`. Het originele bestand is altijd de source of
truth: wijzigingen worden toegepast als gerichte patches, en alles wat de
editor niet heeft aangeraakt, komt ongewijzigd terug uit de rondgang.

```
open docx ─► archive original by hash (never touched)
          ─► parse word/document.xml into a block tree, each block anchored to its original XML
          ─► Tiptap editor (manual + AI editing, dirty tracking)
save      ─► dirty blocks → OOXML fragments (referencing existing styles only)
          ─► splice into the original document.xml; untouched blocks keep their bytes
          ─► repack the zip; every other entry is copied byte-for-byte
```

De pakket-voor-pakket-rondleiding (docx/pptx-engines, `pdf2docx`,
`html2docx`, de agent-core en providers) vind je in
[CONTRIBUTING.md](../../CONTRIBUTING.md#engine-packages).

## Ontwikkeling

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

De sheets-app heeft daarnaast een Rust-toolchain nodig voor de xlsx-sidecar
(`cargo` op je PATH); `npm run build -w @genoffice/sheets` compileert die
automatisch. Zie [CONTRIBUTING.md](../../CONTRIBUTING.md) voor de checks die
elke wijziging moet doorstaan en hoe pull requests worden gemerged.

## Community

GenOffice wordt actief doorontwikkeld en jouw feedback bepaalt mee de
richting.

- **Meld een bug of vraag een functie aan** via
  [GitHub Issues](https://github.com/genspark-ai/genoffice/issues).
- **Doe mee aan de GenOffice-groepschat** op
  [GenTeam](https://genoffice.ai/join) om met het team en andere gebruikers te
  praten.
- **Geef de repo een ster** als je iets aan GenOffice hebt — dat is de beste
  manier om het project te steunen.

## Veelgestelde vragen

<details>
<summary><b>Is GenOffice gratis?</b></summary>

Ja. GenOffice is gratis en open source onder de Apache-2.0-licentie — geen
proefperiode, geen betaalde laag voor de apps zelf.

</details>

<details>
<summary><b>Kan GenOffice bestanden van Microsoft Word, Excel en PowerPoint openen?</b></summary>

Ja. GenOffice opent en bewaart native `.docx`-, `.xlsx`- en `.pptx`-bestanden.
Opslaan is byte-behoudend: delen van het bestand die je niet hebt aangeraakt,
worden byte-voor-byte terugweggeschreven, zodat documenten gewoon blijven
werken in Microsoft Office.

</details>

<details>
<summary><b>Werkt GenOffice offline?</b></summary>

Documenten bewerken is volledig lokaal — bestanden verlaten je machine nooit
om te worden geopend, bewerkt, opgeslagen of geconverteerd. De AI-functies
(agents, search, beeldtools) hebben een netwerkverbinding nodig, met ofwel
een Genspark-login of je eigen model-API-key.

</details>

<details>
<summary><b>Kan GenOffice PDF-bestanden bewerken?</b></summary>

Ja — echte bewerking van PDF-tekst en -afbeeldingen, waarbij de content
stream van de pagina wordt herschreven met behoud van de originele fonts,
niet met overplaknotities.

</details>

<details>
<summary><b>Kan GenOffice PDF converteren naar Word, Excel of PowerPoint?</b></summary>

Ja — volledig on-device: tekenniveau-extractie met PDFium plus
geometrie-gebaseerde layout-analyse, geen cloudservice, geen upload. Gescande
pagina's worden ook ondersteund: op macOS en Windows leest de systeem-OCR ze,
zodat ze worden omgezet naar bewerkbare tekst in plaats van een
pagina-afbeelding.

</details>

<details>
<summary><b>Kan ik mijn eigen AI-model of API-key gebruiken?</b></summary>

Ja. Naast de keyless Genspark-login ondersteunt GenOffice je eigen key voor
Claude, OpenAI, Gemini, DeepSeek, Kimi, GLM, Qwen, Doubao, MiniMax, Grok,
Mistral, OpenRouter, API Route, Requesty en OpenCode Zen/Go, plus elk OpenAI-compatibel endpoint —
inclusief lokale modelservers. Search, beeldgeneratie en analyse van beeld en
video gebruiken hun eigen keys onder Settings → AI Media & Search.

</details>

<details>
<summary><b>Kan GenOffice HTML converteren naar Word?</b></summary>

Ja — Export as Word in de HTML-app levert volledig on-device een native,
bewerkbare `.docx` op. De pagina wordt gerenderd in de ingebouwde Chromium en
teruggebracht tot echte Word-structuren: koppen, alinea's, lijsten, tabellen,
kaarten, KPI-rijen, formuliervelden en paginaachtergronden; alleen visuals
zonder Word-equivalent (grafieken, iconen, gestileerde vakken) worden als
afbeelding ingevoegd.

</details>

<details>
<summary><b>Verzamelt GenOffice gegevens?</b></summary>

De officiële, kant-en-klare builds versturen standaard beperkte
gebruiksanalyse, en je kunt rapportage op elk moment uitschakelen onder
Settings → General. Analytics verstuurt nooit documentinhoud, bestandsnamen,
bestandspaden, accountidentiteit of e-mailadressen. Zie
[GenOffice Privacy](../../PRIVACY.md) voor de volledige lijst van events en
gegevens die worden gedeeld.

</details>

## Beveiliging

Zie [SECURITY.md](../../SECURITY.md) voor de beveiligingsaanpak van de
processen (renderer-sandboxing, IPC-validatie, external-link gating) en de
dreigingsmodellen voor door AI gegenereerde content.

## Met dank aan

GenOffice zou niet mogelijk zijn zonder deze open-source projecten:

- [Electron](https://www.electronjs.org/) — de desktop-runtime voor elke app.
- [Univer](https://github.com/dream-num/univer) (Apache-2.0) — de
  spreadsheet-UI-kern waarop Sheets is gebouwd.
- [PDFium](https://pdfium.googlesource.com/pdfium/) (BSD-3-Clause, gebundeld
  via [@embedpdf/pdfium](https://github.com/embedpdf/embed-pdf-viewer)) — de
  content-stream-engine achter echte bewerking van PDF-tekst en
  -afbeeldingen.
- [pdf.js](https://github.com/mozilla/pdf.js) (Apache-2.0) en
  [pdf-lib](https://github.com/Hopding/pdf-lib) (MIT) — PDF-rendering en het
  samenstellen van documenten.
- [Tiptap](https://tiptap.dev/) / [ProseMirror](https://prosemirror.net/) —
  de blokeditors in Docs en Markdown.
- [CodeMirror](https://codemirror.net/) (MIT) — de broncode-editor in HTML.
- [Konva](https://konvajs.org/) — canvas-rendering voor Slides en
  Sheets-grafieken.
- [HarfBuzz](https://github.com/harfbuzz/harfbuzz) (wasm) —
  text-shaping-metrieken voor complexe schriftsystemen.
- [calamine](https://github.com/tafia/calamine) en
  [IronCalc](https://github.com/ironcalc/IronCalc) — de read- en calc-laag
  van de Rust-xlsx-sidecar.
- [libeot](https://github.com/umanwizard/libeot) (MPL-2.0) — de MicroType
  Express-decoder voor ingesloten PowerPoint-lettertypen, overgezet naar
  TypeScript.
- [React](https://react.dev/) (MIT) — de UI-laag van elke app.
- [Mermaid](https://mermaid.js.org/) (MIT) en [KaTeX](https://katex.org/)
  (MIT) — diagrammen en wiskunde in Markdown en Docs.
- [opentype.js](https://opentype.js.org/) (MIT) — lettertype-parsing voor
  metrieken en glyph-opzoeking.
- [JSZip](https://stuk.github.io/jszip/) (MIT) en
  [fast-xml-parser](https://github.com/NaturalIntelligence/fast-xml-parser)
  (MIT) — de OOXML-container- en XML-laag.
- [Fluent UI System Icons](https://github.com/microsoft/fluentui-system-icons)
  (MIT) — de icoonset in de linten.
- [electron-updater](https://www.electron.build/) (MIT) — updates binnen
  de app.
- Liberation, Carlito, Caladea en Noto CJK-fonts (OFL/Apache-2.0) —
  meegeleverde documentfonts.

`npm run notices` genereert het gebundelde overzicht van licenties van derden
opnieuw (`tools/gen-third-party-notices.mjs`); alle runtime-dependencies zijn
MIT/Apache-2.0/BSD-3-Clause/OFL.

## Licentie

GenOffice wordt uitgebracht onder de [Apache License 2.0](../../LICENSE), met
één uitzondering: de map `ee/` is gereserveerd voor toekomstige
enterprise-modules en valt onder de
[GenOffice Enterprise License](../../ee/LICENSE).

De namen en logo's van GenOffice en Genspark zijn handelsmerken van Mainfunc,
Inc. De Apache-2.0-licentie geeft geen toestemming om deze te gebruiken (zie
sectie 6); forks moeten hun eigen branding gebruiken.
