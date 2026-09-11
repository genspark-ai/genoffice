<p align="center">
  <a href="https://genoffice.ai/">
    <picture>
      <source srcset="docs/assets/readme/hero-dark.webp" media="(prefers-color-scheme: dark)">
      <img src="docs/assets/readme/hero.webp" alt="GenOffice — the open-source AI Office suite: Docs, Sheets, Slides, PDF, Markdown and HTML with a built-in AI panel" width="100%">
    </picture>
  </a>
</p>

<h1 align="center">GenOffice</h1>

<p align="center"><b>The world's first full-featured open-source AI Office suite.</b><br>
Word, Excel, PowerPoint and PDF files, edited by you and your AI, saved back in the real formats.</p>

<p align="center">
  <a href="LICENSE"><img src="https://img.shields.io/github/license/genspark-ai/genoffice" alt="License: Apache-2.0"></a>
  <a href="https://github.com/genspark-ai/genoffice/releases/latest"><img src="https://img.shields.io/github/v/release/genspark-ai/genoffice" alt="Latest release"></a>
  <a href="https://github.com/genspark-ai/genoffice/releases"><img src="https://img.shields.io/github/downloads/genspark-ai/genoffice/total" alt="Downloads"></a>
  <a href="https://github.com/genspark-ai/genoffice/stargazers"><img src="https://img.shields.io/github/stars/genspark-ai/genoffice?style=flat" alt="GitHub stars"></a>
</p>

<p align="center"><b>English</b> · <a href="docs/i18n/README.es.md">Español</a> · <a href="docs/i18n/README.pt-BR.md">Português (Brasil)</a> · <a href="docs/i18n/README.de.md">Deutsch</a> · <a href="docs/i18n/README.fr.md">Français</a> · <a href="docs/i18n/README.zh-CN.md">简体中文</a> · <a href="docs/i18n/README.zh-TW.md">繁體中文</a> · <a href="docs/i18n/README.ko.md">한국어</a> · <a href="docs/i18n/README.ja.md">日本語</a> · <a href="docs/i18n/README.ar.md">العربية</a> · <a href="docs/i18n/README.ru.md">Русский</a> · <a href="docs/i18n/README.it.md">Italiano</a> · <a href="docs/i18n/README.nl.md">Nederlands</a> · <a href="docs/i18n/README.pl.md">Polski</a> · <a href="docs/i18n/README.cs.md">Čeština</a> · <a href="docs/i18n/README.id.md">Bahasa Indonesia</a> · <a href="docs/i18n/README.ms.md">Bahasa Melayu</a> · <a href="docs/i18n/README.th.md">ไทย</a> · <a href="docs/i18n/README.hi.md">हिन्दी</a> · <a href="docs/i18n/README.he.md">עברית</a></p><!-- lang-switcher · public-hygiene: allow -->

<p align="center">
  <a href="#download"><b>Download</b></a> ·
  <a href="https://genoffice.ai/"><b>Website</b></a> ·
  <a href="https://genoffice.ai/join"><b>Community</b></a> ·
  <a href="PRIVACY.md"><b>Privacy</b></a>
</p>

GenOffice is a free, open-source alternative to Microsoft Office for macOS,
Windows and Linux. It opens and saves native `.docx`, `.xlsx` and `.pptx`
files, edits PDF, Markdown and HTML, and puts an AI agent next to every
document — not a chat box bolted on the side, but an editor that reads the
file, makes the change, and shows you exactly what it touched.

- **Real formats, byte-preserving.** Only what you edit is rewritten. Everything
  else in the file survives byte-for-byte, so documents keep working in Word,
  Excel and PowerPoint.
- **AI you can review.** Edits land as tracked changes and diffs with one-click
  rollback. Spreadsheets get live formulas, not pasted numbers. Decks and pages
  are generated onto the canvas and stay fully editable.
- **Local by design.** Files open, edit, save and convert on your machine.
  PDF → Word / Excel / PowerPoint, Markdown → Word and HTML → Word all run
  on-device. Only the AI calls leave the machine, to the provider you choose.
- **Your keys or none.** Sign in with Genspark and skip keys, or bring your own
  key for Claude, OpenAI, Gemini, DeepSeek, Kimi, GLM, Qwen, Doubao, MiniMax,
  Grok, Mistral, OpenRouter, Requesty, or any OpenAI-compatible endpoint, local
  servers included.

**Get it:** [macOS](https://github.com/genspark-ai/genoffice/releases/latest) (Apple Silicon and Intel) ·
[Windows](https://github.com/genspark-ai/genoffice/releases/latest) (x64 and Arm) ·
[Linux](https://github.com/genspark-ai/genoffice/releases/latest) (deb, rpm, AppImage) —
details and requirements in [Download](#download).

## Demo

Six apps, one AI panel. Every screenshot is the real app on macOS, with the
AI driven from the prompt you can read in the panel.

### 1 · Docs — open and edit `.docx` with an AI you can review

<table>
<tr>
<td width="50%"><img src="docs/assets/readme/docs-report.webp" alt="GenOffice Docs rendering a two-column annual-report page with a full-width cover image, shaded KPI table, header and footer, at 80% zoom with the AI panel collapsed"></td>
<td width="50%"><img src="docs/assets/readme/docs-ai.webp" alt="GenOffice Docs: a company overview with a banner image; the AI tightened the Overview and inserted a new bulleted section, and the panel offers a one-click roll back"></td>
</tr>
<tr>
<td><b>Opens the file as Word lays it out</b> — two-column sections, full-bleed images, shaded tables, headers and footers, pagination on Word's line metrics. Styles, comments, tracked changes, equations and ink round-trip untouched.</td>
<td><b>Ask for the edit</b> — the AI reads the blocks it needs, rewrites the Overview and inserts a new bulleted section. Every AI turn is a snapshot you can roll back; with <b>Track changes</b> on, edits arrive as Word-style revisions.</td>
</tr>
</table>

### 2 · Sheets — `.xlsx` with live formulas and charts, not pasted numbers

<table>
<tr>
<td width="50%"><img src="docs/assets/readme/sheets-ai.webp" alt="GenOffice Sheets: the AI added a Summary sheet with revenue by region and category using SUMIF formulas, plus a column chart, and reports 43 applied changes with an Undo button"></td>
<td width="50%"><img src="docs/assets/readme/sheets-qa.webp" alt="GenOffice Sheets: asked which region led Q2 revenue, the AI answers Europe with the category breakdown and cites the cells it used as links, next to the Orders sheet"></td>
</tr>
<tr>
<td><b>Build it</b> — from one sentence the agent adds a Summary sheet with real <code>SUMIF</code>s by region and category, inserts a column chart, and applies the 43 changes as a single undoable batch.</td>
<td><b>Ask it</b> — questions about the workbook come back with the reasoning and the exact cells used as clickable citations. Under the hood: an in-house Rust <code>.xlsx</code> engine, pivot tables, slicers, conditional formatting and formula tracing.</td>
</tr>
</table>

### 3 · Slides — from a prompt to a `.pptx` deck

<img src="docs/assets/readme/slides-generate.webp" alt="Time-lapse of GenOffice Slides generating the Aurora Home investor deck: the AI plans the storyline in the panel, slides appear on the canvas one after another, and the finished deck ends on the closing ask" width="100%">

<table>
<tr>
<td width="50%"><img src="docs/assets/readme/slides-cover.webp" alt="GenOffice Slides: the cover slide of an AI-generated Aurora Home investor deck on the canvas, with the original one-line prompt and the AI's summary of what it built in the panel"></td>
<td width="50%"><img src="docs/assets/readme/slides-ai.webp" alt="GenOffice Slides: the designed closing slide of the same 11-slide deck, with the thumbnail strip on the left and the AI panel summarizing the storyline"></td>
</tr>
<tr>
<td><b>One line in</b> — "Create a 10-slide investor pitch deck for Aurora Home…". GenOffice plans the storyline, researches the numbers, and drafts every slide onto the canvas as a real <code>.pptx</code>.</td>
<td><b>A finished deck out</b> — eleven designed slides with consistent typography, imagery and a closing call to action; keep editing with masters, layouts, smart guides and non-destructive cropping, or ask the panel to restyle, rewrite and reorder.</td>
</tr>
</table>

### 4 · PDF — edit PDF text in place, convert PDF to Word on-device

<table>
<tr>
<td width="50%"><img src="docs/assets/readme/pdf-edit.webp" alt="GenOffice PDF: Edit text mode outlines every text block on the page for in-place editing while the AI panel answers a question about the report with page citations"></td>
<td width="50%"><img src="docs/assets/readme/pdf-convert.webp" alt="GenOffice Docs showing a Word document converted locally from the Helios quarterly review PDF, opened in a second tab beside the original PDF"></td>
</tr>
<tr>
<td><b>Edit inside the page</b> — Edit text mode outlines every text block for in-place retyping; the content stream is rewritten through PDFium with the original fonts, not a cover-up annotation. Ask the AI about a long report and get answers with page citations.</td>
<td><b>Convert on-device</b> — <b>PDF Converter → PDF to Word</b> produces an editable <code>.docx</code> that opens in Docs next to the source, headings, stat rows and paragraphs intact. Excel and PowerPoint targets work the same way; scanned pages go through the system OCR.</td>
</tr>
</table>

### 5 · HTML — an AI page and UI builder, design brief first

Say what the page is for and who it is for. The AI proposes a **design brief**
first — hook, palette, typography and style directions — then builds a single
self-contained `.html` file against those tokens.

<img src="docs/assets/readme/html-restyle-motion.webp" alt="Time-lapse of GenOffice HTML restyling the Lumen landing page: one Restyle request in the panel turns the dark Midnight Studio page into the warm Solar Daybreak version while every section and all copy stay in place" width="100%">

<table>
<tr>
<td width="50%"><img src="docs/assets/readme/html-ai.webp" alt="GenOffice HTML: a generated landing page for a solar desk lamp in the dark Midnight Studio direction, shown in the live preview with the AI panel summarizing the page it just built"></td>
<td width="50%"><img src="docs/assets/readme/html-restyle.webp" alt="The same Lumen landing page restyled by the AI into the warm Solar Daybreak direction: paper background, serif headlines and an orange accent, with every section and all copy kept"></td>
</tr>
<tr>
<td><b>Generated from one prompt</b> — a bold hero, feature cards, pricing and a waitlist form for Lumen, built in the Midnight Studio direction. Click any element to restyle it, double-click to edit text, or switch to the CodeMirror source view.</td>
<td><b>Same design, new direction</b> — one <b>Restyle</b> request swaps the brief's tokens and the page follows: warm paper, editorial serif, sun-orange accent, nothing rewritten. Present fullscreen, or export as PDF or a native editable Word document.</td>
</tr>
</table>
<table>
<tr>
<td width="50%"><img src="docs/assets/readme/html-dashboard.webp" alt="GenOffice HTML: a generated personal dashboard UI for a freelance designer in a warm linen style, with a left rail, serif greeting and four metric cards"></td>
<td width="50%"><img src="docs/assets/readme/html-report.webp" alt="GenOffice HTML: a generated EV-market data report in a broadsheet style, with a serif masthead, a 17.3 million headline figure and a stat row"></td>
</tr>
<tr>
<td><b>UI mockups</b> — the "personal dashboard" starter turns a persona into a working layout: left rail, greeting, billable-hours sparkline, invoice and utilization cards, all real HTML you can hand to a developer.</td>
<td><b>Data stories</b> — the "data report" starter builds an editorial broadsheet: serif masthead, one headline number, a rule-separated stat row, inline SVG charts and a methodology note.</td>
</tr>
</table>

### 6 · Markdown — a block editor over plain `.md`, with Ask AI

<table>
<tr>
<td width="50%"><img src="docs/assets/readme/markdown-ai.webp" alt="GenOffice Markdown: a selected paragraph shows an Ask AI popover with a typed instruction and suggestion chips such as Polish, Make more concise, Expand and Fix grammar, plus Send now and Add to queue buttons"></td>
<td width="50%"><img src="docs/assets/readme/markdown-render.webp" alt="GenOffice Markdown rendering a launch-notes document with a table, a Mermaid flowchart and a task list, with the AI panel's starter prompts on the left"></td>
</tr>
<tr>
<td><b>Ask AI about a selection</b> — select any passage and an <b>Ask AI</b> chip appears: type an instruction or pick a suggestion, send it now, or queue several anchored edits and run them in one pass. The same entry exists in every app.</td>
<td><b>Rendered, saved as plain Markdown</b> — headings, lists, tables, images, code blocks and Mermaid diagrams in a Tiptap block editor, written back as plain <code>.md</code>, with a fully local <b>Markdown → Word</b> export.</td>
</tr>
</table>

## Why GenOffice

- **Open source**, Apache-2.0, built in the open on GitHub.
- **Yours to run.** Native apps for macOS, Windows and Linux; files stay on your
  disk and every edit, save and conversion happens on your machine.
- **Real Office files.** Native `.docx`, `.xlsx` and `.pptx`, byte-preserving:
  the parts of a file you did not touch are copied exactly as they were.
- **An AI that edits the document itself.** Tracked changes in Docs, live
  formulas and charts in Sheets, slides drawn onto the canvas, every AI turn a
  snapshot you can roll back.
- **Your model, your key.** Sign in with Genspark, or bring a key for Claude,
  OpenAI, Gemini, DeepSeek and more, local servers and any OpenAI-compatible
  endpoint included.
- **PDF done properly.** Edit text inside the page, and convert PDF to Word,
  Excel or PowerPoint on-device, with system OCR for scans.
- **Markdown and HTML too**, with the same AI panel and local export to Word.
- **Free**, for individuals and teams alike.

## AI backends

**Sign in with Genspark** and there is nothing to configure: model calls route
through the Genspark proxy (Claude, GPT and Gemini families) and the agents get
web and image search, image generation, and image/audio/video analysis.

**Or bring your own key.** Settings → AI lists Claude, OpenAI, Gemini,
DeepSeek, Kimi, GLM, Qwen, Doubao, MiniMax, Grok, Mistral, OpenRouter, Requesty
and OpenCode Zen/Go, plus a custom slot for any OpenAI-compatible endpoint (base
URL + key), including local model servers. Search and media have their own
per-capability providers under **AI Media & Search**: Serper or Tavily for web
search, and OpenAI, Gemini, Doubao/Seedream, GLM, Grok, Qwen, MiniMax or any
OpenAI-compatible images endpoint for image generation and image/video
analysis.

The whole suite ships light, dark and system themes. Themes only change what
is on screen: exports, prints and saved files always keep the document's own
colors.

## Download

| Platform                             | Requirements                                          | Download                                                                                  |
| ------------------------------------ | ----------------------------------------------------- | ----------------------------------------------------------------------------------------- |
| **macOS** — Apple Silicon (arm64)    | macOS 11+                                             | [Latest `.dmg` (arm64)](https://github.com/genspark-ai/genoffice/releases/latest)         |
| **macOS** — Intel (x64)              | macOS 11+                                             | [Latest `.dmg` (x64)](https://github.com/genspark-ai/genoffice/releases/latest)           |
| **Windows** (x64, most PCs)          | Windows 10+, Intel/AMD                                | [Latest `-x64.exe` installer](https://github.com/genspark-ai/genoffice/releases/latest)   |
| **Windows** on Arm (ARM64)           | Windows 11 on Arm (Snapdragon X and similar)          | [Latest `-arm64.exe` installer](https://github.com/genspark-ai/genoffice/releases/latest) |
| **Linux** — Debian / Ubuntu          | x86_64, glibc 2.34+ (Ubuntu 22.04 or newer)           | [Latest `.deb`](https://github.com/genspark-ai/genoffice/releases/latest)                 |
| **Linux** — Fedora / RHEL / openSUSE | x86_64, glibc 2.34+ (Fedora 35+, RHEL 9+, Leap 15.6+) | [Latest `.rpm`](https://github.com/genspark-ai/genoffice/releases/latest)                 |
| **Linux** — other distributions      | x86_64, glibc 2.34+, FUSE 2                           | [Latest `.AppImage`](https://github.com/genspark-ai/genoffice/releases/latest)            |

All builds come from `main`; the macOS and Windows installers are signed.
Older versions are on the [Releases](https://github.com/genspark-ai/genoffice/releases) page.

<details>
<summary><b>Installing on Linux</b></summary>

The deb installs with apt — it pulls in the dependencies and adds GenOffice
to the applications menu:

```bash
sudo apt install ./genoffice_<version>_amd64.deb
```

On Fedora / RHEL-family / openSUSE, install the rpm instead:

```bash
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / RHEL family
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

The AppImage runs in place: install the FUSE 2 runtime
(`sudo apt install libfuse2`; on Ubuntu 24.04 the package is `libfuse2t64`),
make the file executable, then run it:

```bash
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

</details>

## How it works

Seven Electron apps — Docs, Sheets, Slides, PDF, Markdown, HTML and the
tabbed shell — share one engine layer of pure TypeScript packages plus a Rust
sidecar for `.xlsx`. The original file is always the source of truth: edits
are applied as narrow patches, and everything the editor did not touch
survives the round trip untouched.

```
open docx ─► archive original by hash (never touched)
          ─► parse word/document.xml into a block tree, each block anchored to its original XML
          ─► Tiptap editor (manual + AI editing, dirty tracking)
save      ─► dirty blocks → OOXML fragments (referencing existing styles only)
          ─► splice into the original document.xml; untouched blocks keep their bytes
          ─► repack the zip; every other entry is copied byte-for-byte
```

The package-by-package tour (docx/pptx engines, `pdf2docx`, `html2docx`, the
agent core and providers) lives in [CONTRIBUTING.md](CONTRIBUTING.md#engine-packages).

## Development

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

The sheets app additionally needs a Rust toolchain for its xlsx sidecar
(`cargo` on PATH); `npm run build -w @genoffice/sheets` compiles it
automatically. See [CONTRIBUTING.md](CONTRIBUTING.md) for the checks every
change must pass and how pull requests land.

## Community

GenOffice is in active development and your feedback shapes it.

- **Report a bug or request a feature** in
  [GitHub Issues](https://github.com/genspark-ai/genoffice/issues).
- **Join the GenOffice group chat** on
  [GenTeam](https://genoffice.ai/join) to talk to the team and other users.
- **Star the repo** if GenOffice is useful to you — it is the best way to
  support the project.

## FAQ

<details>
<summary><b>Is GenOffice free?</b></summary>

Yes. GenOffice is free and open-source under the Apache-2.0 license — no
trial, no paid tier for the apps themselves.

</details>

<details>
<summary><b>Can GenOffice open Microsoft Word, Excel and PowerPoint files?</b></summary>

Yes. GenOffice opens and saves native `.docx`, `.xlsx` and `.pptx` files.
Saving is byte-preserving: parts of the file you didn't touch are written
back byte-for-byte, so documents keep working in Microsoft Office.

</details>

<details>
<summary><b>Does GenOffice work offline?</b></summary>

Document editing is fully local — files never leave your machine to be
opened, edited, saved or converted. The AI features (agents, search, image
tools) need a network connection, with either a Genspark sign-in or your own
model API key.

</details>

<details>
<summary><b>Can GenOffice edit PDF files?</b></summary>

Yes — real PDF text and image editing that rewrites the page content stream
with the original fonts preserved, not cover-up annotations.

</details>

<details>
<summary><b>Can GenOffice convert PDF to Word, Excel or PowerPoint?</b></summary>

Yes — entirely on-device: PDFium character-level extraction plus
geometry-based layout analysis, no cloud service, no upload. Scanned pages
are covered too: on macOS and Windows the system OCR reads them, so they
convert to editable text rather than a page image.

</details>

<details>
<summary><b>Can I use my own AI model or API key?</b></summary>

Yes. Besides the keyless Genspark sign-in, GenOffice supports bring your own
key for Claude, OpenAI, Gemini, DeepSeek, Kimi, GLM, Qwen, Doubao, MiniMax,
Grok, Mistral, OpenRouter, Requesty and OpenCode Zen/Go, plus any OpenAI-compatible
endpoint — including local model servers. Search, image generation and
image/video analysis take their own keys under Settings → AI Media & Search.

</details>

<details>
<summary><b>Can GenOffice convert HTML to Word?</b></summary>

Yes — Export as Word in the HTML app produces a native, editable `.docx`
entirely on-device. The page is rendered in the built-in Chromium and reduced
to real Word structures: headings, paragraphs, lists, tables, cards, KPI rows,
form fields and page backgrounds; only visuals with no Word counterpart
(charts, icons, decorated boxes) are embedded as pictures.

</details>

<details>
<summary><b>Does GenOffice collect any data?</b></summary>

Official packaged builds send limited usage analytics by default, and you can
disable reporting at any time under Settings → General. Analytics never sends
document content, file names, file paths, account identity or email
addresses. See [GenOffice Privacy](PRIVACY.md) for the complete event and
data disclosures.

</details>

## Security

See [SECURITY.md](SECURITY.md) for the process security posture (renderer
sandboxing, IPC validation, external-link gating) and the threat models for
AI-generated content.

## Acknowledgements

GenOffice would not be possible without these open-source projects:

- [Electron](https://www.electronjs.org/) — the desktop runtime for every app.
- [Univer](https://github.com/dream-num/univer) (Apache-2.0) — the spreadsheet
  UI core that Sheets extends.
- [PDFium](https://pdfium.googlesource.com/pdfium/) (BSD-3-Clause, bundled via
  [@embedpdf/pdfium](https://github.com/embedpdf/embed-pdf-viewer)) — the
  content-stream engine behind true PDF text and image editing.
- [pdf.js](https://github.com/mozilla/pdf.js) (Apache-2.0) and
  [pdf-lib](https://github.com/Hopding/pdf-lib) (MIT) — PDF rendering and
  document assembly.
- [Tiptap](https://tiptap.dev/) / [ProseMirror](https://prosemirror.net/) —
  the block editors in Docs and Markdown.
- [CodeMirror](https://codemirror.net/) (MIT) — the source editor in HTML.
- [Konva](https://konvajs.org/) — canvas rendering for Slides and Sheets
  charts.
- [HarfBuzz](https://github.com/harfbuzz/harfbuzz) (wasm) — text-shaping
  metrics for complex scripts.
- [calamine](https://github.com/tafia/calamine) and
  [IronCalc](https://github.com/ironcalc/IronCalc) — the read and calc layers
  of the Rust xlsx sidecar.
- [libeot](https://github.com/umanwizard/libeot) (MPL-2.0) — the MicroType
  Express decoder for embedded PowerPoint fonts, ported to TypeScript.
- [React](https://react.dev/) (MIT) — the UI layer of every app.
- [Mermaid](https://mermaid.js.org/) (MIT) and [KaTeX](https://katex.org/)
  (MIT) — diagrams and math in Markdown and Docs.
- [opentype.js](https://opentype.js.org/) (MIT) — font parsing for metrics
  and glyph lookup.
- [JSZip](https://stuk.github.io/jszip/) (MIT) and
  [fast-xml-parser](https://github.com/NaturalIntelligence/fast-xml-parser)
  (MIT) — the OOXML container and XML layers.
- [Fluent UI System Icons](https://github.com/microsoft/fluentui-system-icons)
  (MIT) — the icon set across the ribbons.
- [electron-updater](https://www.electron.build/) (MIT) — in-app updates.
- Liberation, Carlito, Caladea, and Noto CJK fonts (OFL/Apache-2.0) — bundled
  document fonts.

`npm run notices` regenerates the bundled third-party license summary
(`tools/gen-third-party-notices.mjs`); all runtime dependencies are
MIT/Apache-2.0/BSD-3-Clause/OFL.

## License

GenOffice is licensed under the [Apache License 2.0](LICENSE), with one
exception: the `ee/` directory is reserved for future enterprise modules and
is covered by the [GenOffice Enterprise License](ee/LICENSE).

The GenOffice and Genspark names and logos are trademarks of Mainfunc, Inc.
The Apache-2.0 license does not grant permission to use them (see section 6);
forks should use their own branding.
