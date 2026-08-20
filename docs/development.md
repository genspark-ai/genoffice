# Development

KĀRYA is an npm workspaces monorepo. This page covers the layout, the
provider architecture, how to add a provider, how tests work, and how to
build.

## Repository structure

```
apps/
  docs/        KĀRYA Docs — .docx word processor
  sheets/      KĀRYA Sheets — .xlsx spreadsheet (+ Rust xlsx sidecar)
  slides/      KĀRYA Slides — .pptx presentation editor
  pdf/         KĀRYA PDF — .pdf viewer/editor
  markdown/    KĀRYA Markdown — .md editor
  shell/       KĀRYA — suite shell (home, tabs, packaging)
packages/
  agent-core/  provider-agnostic AI agent loop + skills
  ai-provider/ provider abstraction, chat/streaming, Ollama discovery
  ai-search/   Genspark auth + web/image search tools
  docx-engine/ pptx-engine/ pptx-render/ file-parse/ ...
  i18n/ ui/ project-store/ electron-utils/
tools/         one-off scripts (icons, i18n insertion, notices)
ee/            reserved enterprise modules (covered by a separate license)
```

Each app is an [electron-vite](https://electron-vite.org) project with three
process layers: `src/main` (Electron main), `src/preload`, and `src/renderer`
(React). The shell packages the five editor modules (`out/` trees) as
`extraResources` so all editors run in shell tabs.

## Provider architecture

```
UI (AI settings dialog, AI panel)
        ↓
AI settings (provider, model, baseUrl, apiKey)
        ↓
AI provider abstraction (packages/ai-provider)
        ↓
Provider adapter (OpenAI-compatible HTTP, Ollama /api/tags, Genspark proxy)
        ↓
Network / local server
```

- `packages/ai-provider/src/providers.ts` — the registry: `AI_PROVIDERS`
  (id, label, models, defaultModel, key policy, base URL policy),
  `isProviderConfigured`, `resolveAiSettings` (legacy migration).
- `packages/ai-provider/src/chat.ts` / `stream.ts` — chat completions and
  streaming over the provider's base URL.
- `packages/ai-provider/src/ollama.ts` — `listOllamaModels` (GET `/api/tags`,
  5s timeout) and `ollamaListStatus` (maps a discovery result to a
  user-facing status: connected / not-running / invalid / unknown).
- `packages/agent-core` — the agent loop and skills. It consumes the
  provider interface and never branches on `provider === 'ollama'`; provider
  behavior lives in `ai-provider`.
- `packages/ui/src/AiSettingsDialog.tsx` — the shared provider settings UI
  used by all five editors.

### Adding a provider

1. Add an entry to `AI_PROVIDERS` in `packages/ai-provider/src/providers.ts`
   with the models your provider supports and its key/base-URL policy.
2. If the provider uses a non-OpenAI protocol, add an adapter alongside
   `chat.ts`/`stream.ts`; if it is OpenAI-compatible, nothing else is
   needed.
3. The shared settings dialog and `isProviderConfigured` pick it up
   automatically — no per-app changes.
4. Add unit tests in `packages/ai-provider/tests` (see below).

## How Ollama works

- Discovery: `listOllamaModels(baseUrl)` hits `/api/tags` with a 5-second
  timeout and returns model names + parameter sizes.
- Caching: the settings dialog caches discovery per endpoint for 10 seconds
  so opening the dialog is cheap; errors are never cached, and the Refresh
  button forces a fresh probe.
- Status: `ollamaListStatus` classifies the result; the dialog maps statuses
  to localized strings. Raw error text never reaches the UI.
- No polling, no auto-download, no auto-start, no auto-removal of models.

## Tests

- Every workspace uses Vitest (`npm test`), plus `cargo test` for the Rust
  xlsx sidecar. Tests never require a real Ollama or network access — the
  `ai-provider` suite mocks HTTP.
- Coverage areas: provider registration and configuration, Ollama discovery
  and status classification (running / not running / invalid endpoint /
  empty list / multiple models / removed model), chat, streaming, tool
  calls, cancellation, error handling, and provider isolation.
- UI tests live in `packages/ui/tests` (jsdom) and per-app `tests/` dirs.

## Building

```bash
npm install
npm run fixtures       # generate test .docx fixtures
npm test               # engine + app unit tests
npm run typecheck      # tsc --noEmit across every workspace
npm run dev            # all five editors + shell against Vite dev servers
npm run dev:docs       # a single app (same pattern per workspace)
npm run dist:win       # notices + build:all + Windows NSIS installer
npm run dist:mac       # macOS dmg/zip (requires the universal xlsx sidecar)
npm run dist:linux     # AppImage + deb + rpm
```

### The Rust xlsx sidecar

Sheets' `.xlsx` import/export runs through `apps/sheets/native/xlsx-engine`
(Rust: calamine + IronCalc). The Windows shell installer expects the sidecar
at `target/x86_64-pc-windows-gnu/release/xlsx-sidecar.exe`:

```bash
rustup target add x86_64-pc-windows-gnu   # one-time
npm run native:build -w @genoffice/sheets # host (msvc) build for dev
cargo build --release --target x86_64-pc-windows-gnu \
  --manifest-path apps/sheets/native/xlsx-engine/Cargo.toml \
  --config apps/sheets/native/xlsx-engine/.cargo/config.toml
```

The `.cargo/config.toml` statically links the CRT for both targets, so the
shipped sidecar has no VC++ / mingw runtime DLL dependency. Building the GNU
target on Windows requires a mingw-w64 GCC toolchain (`choco install mingw`).

### Icons and branding

The KĀRYA icon set (per-app `build/icon.{png,ico,icns}`, the shell's Linux
`build/icons/` set, and the in-app lockup) is generated by
`tools/gen-karya-icons.mjs`:

```bash
node tools/gen-karya-icons.mjs
```

The in-app logo lockup lives at
`apps/shell/src/renderer/src/assets/karya-logo.svg` and is intentionally
monochrome because the shell inverts it via CSS in dark mode.
