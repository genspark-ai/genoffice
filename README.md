# [GenOffice](https://genoffice.ai/)

**The world's first full-featured open-source AI Office suite.**

[![License: Apache-2.0](https://img.shields.io/github/license/genspark-ai/genoffice)](LICENSE)
[![Latest release](https://img.shields.io/github/v/release/genspark-ai/genoffice)](https://github.com/genspark-ai/genoffice/releases/latest)
[![Downloads](https://img.shields.io/github/downloads/genspark-ai/genoffice/total)](https://github.com/genspark-ai/genoffice/releases)
[![GitHub stars](https://img.shields.io/github/stars/genspark-ai/genoffice?style=flat)](https://github.com/genspark-ai/genoffice/stargazers)
![Platforms: macOS | Windows | Linux](https://img.shields.io/badge/platform-macOS%20%7C%20Windows%20%7C%20Linux-lightgrey)

[Website](https://genoffice.ai/) · [Download](https://github.com/genspark-ai/genoffice/releases/latest) · [Demo](https://www.youtube.com/watch?v=B2pLdMX95v4)

GenOffice is a free, open-source alternative to Microsoft Office for macOS,
Windows, and Linux, built around AI editing as a first-class workflow rather
than a bolted-on chat box. It opens and saves the real Microsoft Office
formats — Word (`.docx`), Excel (`.xlsx`), PowerPoint (`.pptx`) — and edits
PDF and Markdown too: a word processor, spreadsheet, presentation editor,
PDF editor, and Markdown editor as six Electron apps sharing one engine
layer.

[![Meet GenOffice — the world's first full-featured open-source AI Office (video)](https://img.youtube.com/vi/B2pLdMX95v4/maxresdefault.jpg)](https://www.youtube.com/watch?v=B2pLdMX95v4)

[Watch the demo video on YouTube](https://www.youtube.com/watch?v=B2pLdMX95v4)

## Features

- **Real PDF editing** — retype text and edit images in the page itself, original fonts preserved.
- **Microsoft Word–compatible, byte-preserving `.docx` editing** — only what you touched changes; Word never notices.
- **Word-faithful pagination** — page breaks land where Word puts them.
- **Excel-compatible spreadsheets** — in-house engine with a Rust `.xlsx` sidecar, own charts, pivot tables, slicers.
- **PowerPoint-compatible presentations** — in-house `.pptx` engine with masters, layouts, smart guides, non-destructive crop.
- **Markdown to Word, fully local** — the same OOXML engine, no Pandoc, no cloud.
- **AI that edits documents** — block-level edits with snapshots and diffs, document-aware agents.
- **Agent tools built in** — web/image search, image generation, media analysis.
- **Light / dark / system themes.**
- **macOS, Windows, Linux.**
- **Free & open-source (Apache-2.0).**

## Download

| Platform                             | Requirements                                          | Download                                                                                                                         |
| ------------------------------------ | ----------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------- |
| **macOS** — Apple Silicon (arm64)    | macOS 11+                                             | [GenOffice-0.6.101-arm64.dmg](https://github.com/genspark-ai/genoffice/releases/download/v0.6.101/GenOffice-0.6.101-arm64.dmg)   |
| **macOS** — Intel (x64)              | macOS 11+                                             | [GenOffice-0.6.101.dmg](https://github.com/genspark-ai/genoffice/releases/download/v0.6.101/GenOffice-0.6.101.dmg)               |
| **Windows** (x64)                    | Windows 10+                                           | [GenOfficeSetup-v0.6.101.exe](https://github.com/genspark-ai/genoffice/releases/download/v0.6.101/GenOfficeSetup-v0.6.101.exe)   |
| **Linux** — Debian / Ubuntu          | x86_64, glibc 2.34+ (Ubuntu 22.04 or newer)           | [genoffice_0.6.101_amd64.deb](https://github.com/genspark-ai/genoffice/releases/download/v0.6.101/genoffice_0.6.101_amd64.deb)   |
| **Linux** — Fedora / RHEL / openSUSE | x86_64, glibc 2.34+ (Fedora 35+, RHEL 9+, Leap 15.6+) | [genoffice-0.6.101.x86_64.rpm](https://github.com/genspark-ai/genoffice/releases/download/v0.6.101/genoffice-0.6.101.x86_64.rpm) |
| **Linux** — other distributions      | x86_64, glibc 2.34+, FUSE 2                           | [GenOffice-0.6.101.AppImage](https://github.com/genspark-ai/genoffice/releases/download/v0.6.101/GenOffice-0.6.101.AppImage)     |

All builds come from `main`; the macOS and Windows installers are signed.
Older versions are on the [Releases](https://github.com/genspark-ai/genoffice/releases) page.

### Installing on Linux

The deb installs with apt — it pulls in the dependencies and adds GenOffice
to the applications menu:

```bash
sudo apt install ./genoffice_0.6.101_amd64.deb
```

On Fedora / RHEL-family / openSUSE, install the rpm instead:

```bash
sudo dnf install ./genoffice-0.6.101.x86_64.rpm     # Fedora / RHEL family
sudo zypper install ./genoffice-0.6.101.x86_64.rpm  # openSUSE
```

The AppImage instead runs in place: install the FUSE 2 runtime
(`sudo apt install libfuse2`; on Ubuntu 24.04 the package is `libfuse2t64`),
make the file executable, then run it:

```bash
chmod +x GenOffice-0.6.101.AppImage
./GenOffice-0.6.101.AppImage
```

## Apps

| App             | Product                | What it is                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      |
| --------------- | ---------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `apps/docs`     | **GenOffice Docs**     | `.docx` word processor. Byte-preserving round trip: only dirty paragraphs are regenerated (paragraph patch), everything else in the original file is kept byte-for-byte, so opening and saving never breaks layout in Word. Paginated view whose line metrics reproduce the original document's layout, tracked changes, comments, styles, equations, ink.                                                                                                                                                                                                      |
| `apps/sheets`   | **GenOffice Sheets**   | `.xlsx` spreadsheet. UI built on the open-source [Univer](https://github.com/dream-num/univer) core (Apache-2.0) with a large layer of in-house extensions; `.xlsx` import/export runs through an in-house Rust sidecar (calamine + IronCalc), charts are rendered in-house (Konva), plus pivot tables, slicers, conditional formatting, and formula tracing.                                                                                                                                                                                                   |
| `apps/slides`   | **GenOffice Slides**   | `.pptx` presentations. In-house `.pptx` parse/render/edit engine with masters, charts, cropping, ink, and text shaping (HarfBuzz metrics).                                                                                                                                                                                                                                                                                                                                                                                                                      |
| `apps/pdf`      | **GenOffice PDF**      | `.pdf` viewer/editor on [pdf.js](https://github.com/mozilla/pdf.js) (Apache-2.0) + [pdf-lib](https://github.com/Hopding/pdf-lib) (MIT): annotations, forms, outlines, stamps, signatures, page operations, and printing support. True text editing — paragraph selection with in-block reflow, alignment restoration, original-font preservation — and content-stream image insert/edit, all rewriting page content streams through [PDFium](https://pdfium.googlesource.com/pdfium/) wasm (BSD-3-Clause) with subset-embedded fonts — no cover-up annotations. |
| `apps/markdown` | **GenOffice Markdown** | `.md` / `.markdown` editor: Tiptap block editor over plain Markdown files — headings, lists, tables, images, code blocks — saved back as plain Markdown, hosted in shell tabs.                                                                                                                                                                                                                                                                                                                                                                                  |
| `apps/shell`    | **GenOffice**          | The suite shell: home screen, tabbed hosting of the five editors, light/dark/system theme, auto-update.                                                                                                                                                                                                                                                                                                                                                                                                                                                         |

Every app embeds the same AI panel: block-granular AI editing with version
snapshots and diffs in docs, a tool-calling agent over workbook/slide/PDF
state in the others.

The whole suite ships light / dark / system UI themes built on shared design
tokens (`packages/ui`), with a CI guard that keeps chrome colors on the token
system. Document surfaces stay light in dark mode — Word-style dark chrome
around white paper — so files render and export identically in both themes.

**AI 模型供应商。** 应用内 AI 面板与智能体所使用的大模型后端现已支持用户自选 —— 包括 Genspark（默认，需登录）、Claude、Gemini、DeepSeek、OpenAI、自定义 OpenAI 兼容接口，或本地 OpenCode 服务 —— 均在 **设置 → 通用 → AI** 中配置。详见 [AI 模型供应商](#ai-模型供应商)。

## AI 模型供应商

GenOffice 不再写死单一 AI 后端：**设置 → 通用 → AI**（shell 设置）可让你选择由哪家供应商为各个应用中的 AI 面板与智能体提供能力、挑选具体模型，并在需要密钥的供应商处粘贴自己的 API Key。相关设置按安装实例存储，绝不会写入文档内容。

<img width="1006" height="764" alt="001" src="https://github.com/user-attachments/assets/50f2c105-0c8c-425c-a06f-4f8ee4e2e591" />
<img width="998" height="756" alt="002" src="https://github.com/user-attachments/assets/eb534f87-aa0c-4c84-ab22-0021f2d409a0" />

| 供应商 | API Key | Base URL | 说明 |
| --- | --- | --- | --- |
| Genspark | 无需 | — | 默认。登录 Genspark 账号即可；请求经由 Genspark 代理（Claude、GPT、Gemini 系列），并通过 `packages/ai-search` 解锁 gsk 工具接口（网页/图片搜索、图片生成、媒体分析、语音转写）。 |
| Claude | 需要（`sk-ant-…`） | — | Anthropic API。 |
| Gemini | 需要（`AIza…`） | — | Google Gemini API。 |
| DeepSeek | 需要（`sk-…`） | — | DeepSeek API。 |
| OpenAI | 需要（`sk-…`） | — | OpenAI API。 |
| 自定义 | 需要 | 需要 | 任意 OpenAI 兼容接口 —— 指向你自己的网关。 |
| OpenCode | 无需 | 需要 | 本地 `opencode serve` 进程。若未运行，GenOffice 会自动拉起（默认 `http://127.0.0.1:3456`）。智能体循环运行在 opencode 内部；每个会话都以 deny-all 权限创建，因此 opencode 自身的工具（bash、文件编辑等）永远无法触碰你的机器。 |

**模型选择。** 每个供应商都内置一份精选模型列表 —— Genspark 提供 `claude-opus-4-7`、`claude-opus-4-8`、`claude-sonnet-4-6`、`claude-haiku-4-5`、`gpt-5.2`、`gemini-3.1-pro-preview`、`gemini-3-flash-preview` 等。**自定义** 与 **OpenCode** 供应商则接受自由输入的模型 id：OpenCode 可写成 `provider/model-id`（如 `anthropic/claude-opus-4-7`），或仅写 `model-id` 以使用服务端默认模型。

供应商抽象与流式调用位于 `packages/ai-provider`；设置界面在 `apps/shell/src/renderer/src/SettingsModal.tsx`，设置通过 `apps/shell/src/main/app-settings.ts` 经 `getAiSettings` / `setAiSettings` IPC 持久化。

## Engine packages

All pure TypeScript, no Electron dependency, unit-tested (except the UI kit):

- `packages/docx-engine` — docx parsing → block tree (with `docxIndex`
  anchors and passthrough), OOXML fragment generation, byte-level paragraph
  patching.
- `packages/pptx-engine` / `packages/pptx-render` — pptx model and rendering.
- `packages/file-parse` — text extraction for AI attachments (office formats,
  text formats).
- `packages/agent-core` — the AI agent loop and skill composition shared by
  every app.
- `packages/ai-provider` — provider abstraction and streaming for the model
  backends.
- `packages/ai-search` — Genspark auth + web/image search tools.
- `packages/i18n`, `packages/ui`, `packages/project-store`,
  `packages/electron-utils` — shared i18n core, React UI kit, recent-files
  store, and Electron main-process helpers.

## Development

```bash
npm install
npm run fixtures     # generate test .docx fixtures
npm test             # engine + app unit tests (docs/sheets/slides need no display)
npm run typecheck    # tsc --noEmit across every workspace
npm run dev          # all five editors + shell against Vite dev servers
npm run dev:docs     # a single app (same pattern works per workspace)
npm run dist:mac     # package macOS dmg (regenerates third-party notices)
npm run dist:win     # package Windows nsis installer
npm run dist:linux   # package Linux AppImage + deb + rpm
```

The sheets app additionally needs a Rust toolchain for its xlsx sidecar
(`cargo` on PATH); `npm run build -w @genoffice/sheets` compiles it
automatically.

Local UI/e2e driver scripts (Playwright + Electron, for local acceptance, not
committed by default) live in [`scripts/drivers/`](scripts/drivers/README.md).

## Architecture notes (docx round trip)

```
open docx ─► archive original by hash (never touched)
          ─► docx-engine parses word/document.xml top-level elements (w:p / w:tbl / …)
          ─► Block tree, each block anchored by docxIndex + original XML slice
          ─► Tiptap streaming editor (manual + AI editing, dirty tracking)
save      ─► dirty blocks → OOXML fragments (referencing existing styles only)
          ─► splice into original document.xml (untouched blocks keep original bytes)
          ─► repack zip; all other entries copied byte-for-byte
```

The same philosophy holds in sheets and slides: the original file is the
source of truth, edits are applied as narrow patches, and everything the
editor didn't touch survives the round trip untouched.

## FAQ

**Is GenOffice free?**
Yes. GenOffice is free and open-source under the Apache-2.0 license — no
trial, no paid tier for the apps themselves.

**Can GenOffice open Microsoft Word, Excel, and PowerPoint files?**
Yes. GenOffice opens and saves native `.docx`, `.xlsx`, and `.pptx` files.
Saving is byte-preserving: parts of the file you didn't touch are written
back byte-for-byte, so documents keep working in Microsoft Office.

**Does GenOffice work offline?**
Document editing is fully local — files never leave your machine to be
opened, edited, or saved. The AI features (agents, search, image tools) sign
in to a Genspark account and need a network connection.

**Can GenOffice edit PDF files?**
Yes — real PDF text and image editing that rewrites the page content stream
with the original fonts preserved, not cover-up annotations.

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
- [Konva](https://konvajs.org/) — canvas rendering for Slides and Sheets
  charts.
- [HarfBuzz](https://github.com/harfbuzz/harfbuzz) (wasm) — text-shaping
  metrics for complex scripts.
- [calamine](https://github.com/tafia/calamine) and
  [IronCalc](https://github.com/ironcalc/IronCalc) — the read and calc layers
  of the Rust xlsx sidecar.
- Liberation, Carlito, Caladea, and Noto CJK fonts (OFL/Apache-2.0) — bundled
  document fonts.

## Third-party notices

`npm run notices` regenerates the bundled third-party license summary
(`tools/gen-third-party-notices.mjs`); all runtime dependencies are
MIT/Apache-2.0/BSD-3-Clause/OFL, and the bundled fonts (Liberation, Carlito,
Caladea, Noto CJK subsets) are OFL/Apache.

## License

GenOffice is licensed under the [Apache License 2.0](LICENSE), with one
exception: the `ee/` directory is reserved for future enterprise modules and
is covered by the [GenOffice Enterprise License](ee/LICENSE).

The GenOffice and Genspark names and logos are trademarks of Mainfunc, Inc.
The Apache-2.0 license does not grant permission to use them (see section 6);
forks should use their own branding.
