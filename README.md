# GenOffice — OpenRouter & NVIDIA NIM Fork

An AI-native desktop office suite for Windows/macOS with a word processor, spreadsheet, presentations and PDF tools. This fork extends the upstream GenOffice project with **bring-your-own-key AI access through OpenRouter and direct NVIDIA NIM**, removing the requirement to use Genspark sign-in for the primary AI workflow.

> **Fork status:** active experimental fork. OpenRouter and NVIDIA NIM integration is being tested across Docs, Sheets and Slides. Some visual/multimodal features may require a vision-capable model.

## Fork AI features

- **OpenRouter BYOK** — enter your own OpenRouter API key inside the application.
- **OpenRouter Free Router** — use `openrouter/free` when available.
- **NVIDIA NIM Direct** — use an NVIDIA `nvapi-...` key without routing through OpenRouter.
- **Multiple NVIDIA models** — experiment with supported Nemotron, Kimi, DeepSeek, Qwen and MiniMax endpoints as available from NVIDIA.
- **Shared provider configuration** — Docs, Sheets and Slides are being routed through the same configurable AI provider layer.
- **Streaming + tool calling** — retains GenOffice's agent-based document, workbook and presentation editing workflow.
- **Real provider errors** — provider/capacity errors should be surfaced directly instead of being converted into a Genspark login prompt.
- **Local key configuration** — API credentials are configured in the desktop application rather than embedded into the Windows build.

### Current provider choices

| Provider | Key | Typical use |
| --- | --- | --- |
| **OpenRouter** | `sk-or-v1-...` | Flexible model access and OpenRouter Free Router |
| **NVIDIA NIM (Direct)** | `nvapi-...` | Direct NVIDIA-hosted models including Nemotron and other available NIM endpoints |

Model availability, free access and rate limits are controlled by the respective provider and can change over time. A model being listed in the UI does not guarantee that every API key has access to it.

### Multimodal note

Some Slides operations, such as visual analysis/beautification, can send rendered slide imagery to the AI model. Those operations require a **multimodal/vision-capable model and endpoint**. Text-only models can still work for normal chat, document generation and compatible tool-calling tasks but may fail on visual slide operations.

## Download / Windows builds

This fork includes a GitHub Actions workflow named **Build Windows EXE**. Windows installers are produced as workflow artifacts after the customization scripts, full TypeScript typecheck and packaging steps complete successfully.

Do not use the upstream download links below if you specifically want the OpenRouter/NVIDIA modifications; build/download the Windows artifact from this fork's Actions workflow.

## Apps

| App | Product | What it is |
| --- | --- | --- |
| `apps/docs` | **GenOffice Docs** | `.docx` word processor with paginated editing, tracked changes, comments, styles, equations and AI editing. |
| `apps/sheets` | **GenOffice Sheets** | `.xlsx` spreadsheet editor built around Univer with GenOffice extensions, import/export, charts, pivots and AI workbook tools. |
| `apps/slides` | **GenOffice Slides** | `.pptx` presentation editor with an in-house parse/render/edit engine and AI slide tools. |
| `apps/pdf` | **GenOffice PDF** | `.pdf` viewer/editor using pdf.js + pdf-lib. |
| `apps/shell` | **GenOffice** | Electron suite shell hosting the editors and shared configuration. |

Every editor uses GenOffice's agent architecture: the AI can do more than return chat text; supported models can call tools that operate on document, workbook and presentation state.

## AI architecture in this fork

```text
Docs / Sheets / Slides
          │
          ▼
   Shared AI / Agent layer
          │
     ┌────┴────┐
     ▼         ▼
 OpenRouter  NVIDIA NIM
     │         │
     ▼         ▼
 selected    selected
  model       model
```

The goal is to keep provider-specific authentication and routing out of individual editor behavior so that a provider failure does not incorrectly trigger legacy Genspark authentication UI.

## Engine packages

The project keeps the upstream GenOffice engine structure:

- `packages/docx-engine` — DOCX parsing, block tree and OOXML patching.
- `packages/pptx-engine` / `packages/pptx-render` — PPTX model and rendering.
- `packages/file-parse` — text extraction for AI attachments.
- `packages/agent-core` — shared AI agent loop and skill composition.
- `packages/ai-provider` — provider abstraction and streaming model backends.
- `packages/ai-search` — upstream search/auth-related functionality; some legacy Genspark-related code may still remain while the fork migration continues.
- `packages/i18n`, `packages/ui`, `packages/project-store`, `packages/electron-utils` — shared infrastructure.

## Development

```bash
npm install
npm run fixtures
npm test
npm run typecheck
npm run dev
npm run dev:docs
npm run dist:mac
npm run dist:win
```

The Sheets app additionally needs a Rust toolchain for its XLSX sidecar.

## Windows fork build

The Windows workflow applies the fork customization scripts before compilation, including the OpenRouter route changes, suite-level Genspark fallback cleanup, NVIDIA provider integration and expanded NVIDIA model choices. It then runs:

```text
npm ci
→ apply fork customization scripts
→ npm run typecheck
→ npm run dist:win
→ upload installer artifact
```

## Architecture notes

GenOffice uses **Electron + React/TypeScript** for the desktop application, with shared engine packages and native/Rust components where appropriate. The original office file remains the source of truth and editors aim to apply narrow changes so untouched file content survives round trips.

## Security

API keys are user-supplied credentials. Do not commit API keys to this repository or bake personal keys into public builds. Treat documents sent to an AI provider according to that provider's privacy and data-handling terms.

See [SECURITY.md](SECURITY.md) for the upstream process-security posture and AI-generated-content threat model.

## Upstream project

This repository is a fork/customization of **GenOffice by Mainfunc/Genspark**. The original project contains the core office editors and agent architecture. This fork's main focus is configurable direct AI-provider access.

## License

GenOffice is licensed under the [Apache License 2.0](LICENSE), with the upstream exception for the `ee/` directory described in [`ee/LICENSE`](ee/LICENSE).

The **GenOffice** and **Genspark** names and logos are trademarks of Mainfunc, Inc. The Apache-2.0 license does not grant trademark rights. This fork should use its own branding if distributed as an independent product.
