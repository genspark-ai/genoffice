# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

Guidance for AI agents and human contributors working in this repo.

## Commands

Node ≥ 22.12, npm ≥ 10. A Rust toolchain (`cargo`) is needed for the sheets
xlsx sidecar (`apps/sheets/native/xlsx-engine`).

```bash
npm install
npm run fixtures        # regenerate fixtures/generated (after docx-engine changes; CI diffs it)
npm run dev             # all six editor renderers (Vite :5173–5178) + shell
npm run dev:docs        # one app standalone
npm run build:all       # build every app + cli + shell (needed before e2e)
npm run lint            # ESLint, 0 errors required
npm run typecheck       # tsc --noEmit per workspace
npm test                # every workspace's vitest (sheets also runs cargo tests + sidecar build)
npm run format          # Prettier on uncommitted changed/new files only (never reformat the tree)
npm run format:check -- --base origin/main
npm run test:e2e        # Playwright against the built Electron app, serial
```

Single workspace / single test (vitest):

```bash
npm test -w @genoffice/docx-engine
npx vitest run -c packages/docx-engine/vitest.config.ts packages/docx-engine/tests/foo.test.ts -t "name"
cd apps/docs && npx vitest run tests/some.test.ts
```

Each workspace has its own `vitest.config.ts` (fixtures resolve relative to the
workspace), so run tests from inside it or pass its config.

CI also runs the repo-specific gates in `tools/` (`npm run check:*`). The
diff-based ones compare against `FORMAT_BASE_REF` / `--base <ref>`. Notable
rules they enforce on new/changed lines:

- `check:inline-rules`, `check:table-layout-rules`, `check:font-rules` — Docs
  text-layout, table column-width, and font face/line-factor decisions live
  only in `apps/docs/src/renderer/text-layout/`, `table-layout/` and
  `text-layout/font-rules/`, not in emitters, callers or the parser.
- `check:page-cut-readers` — pagination preview / PDF export consume the
  canvas PageCut data (`pagination-pass.ts`); don't re-slice or re-measure.
- `check:font-metrics-provenance`, `check:generated-fonts` — line-metric
  numbers must cite the Word measurement; `*.generated.*` font files come from
  the measurement manifest, never hand edits.
- `check:word-truths` — `docs/word-truths-index.json` must match fixtures.
- `check:skill-version` — a body change to `skills/*/SKILL.md` needs a
  `metadata.version` bump.
- `check:english-comments` — code comments and docs are English-only;
  non-English text belongs only in i18n tables and fixtures.

## Architecture

npm-workspaces monorepo; Electron + React + TypeScript.

- `apps/shell` is the only Electron entry point that ships. Its main process
  imports each editor's main module directly as TS source
  (`apps/{docs,sheets,slides,pdf,markdown,html}/src/main/*-main.ts`) and
  bundles them, which is why editor main-process changes need a shell rebuild.
  The shell also owns home/file-index search (node:sqlite), the control server
  used by the CLI, the MCP server, headless export, and updates.
- Each editor app has `src/main` (Electron main), `src/preload`,
  `src/renderer` (React UI, served by its own Vite port in dev), `src/shared`
  (IPC contracts, command tables), and `tests/`.
- `packages/*` are pure TS engines with no Electron dependency: format engines
  (`docx-engine`, `pptx-engine`/`pptx-render`/`pptx-ops`, `xlsx-gateway`),
  converters (`pdf2docx`, `html2docx`, `hwp2docx`, `doc2docx`, `ppt2pptx`,
  `odf2ooxml`, `compound-file`), AI (`agent-core` loop + skills,
  `ai-provider`, `ai-search`), and shared `ui`, `i18n`, `project-store`,
  `electron-utils`, `zip-gate`. `packages/cli` is the `genoffice` CLI.
- Sheets' xlsx import/export runs in a Rust sidecar process.
- Legacy/ODF/HWP formats open as converted OOXML copies; saving always writes
  OOXML.

**Round-trip fidelity is the core invariant.** The original file is archived
and treated as source of truth. Docx: `docx-engine` parses top-level body
elements into a block tree anchored by `docxIndex` + original XML slice; on
save only dirty blocks are regenerated and spliced back, and every other zip
entry is copied byte-for-byte. Sheets and slides follow the same narrow-patch
approach. Changes to open/save paths need a round-trip test proving untouched
content survives byte-for-byte.

## Repo conventions

- GitHub `main` is a mirror advanced by squashed `Sync snapshot` commits from a
  private tree; external PRs are imported by maintainers rather than merged.
- Playwright/Electron acceptance drivers and Office-comparison scripts are
  local-only (gitignored); never wire them into CI.
- Prefer a new module over growing an already-large file.

## Theming rules (mandatory)

The suite supports light / dark / system UI themes. The switching mechanism is a
`data-theme` attribute on `<html>` plus CSS custom properties defined once in
`packages/ui/src/tokens.css` (light defaults in `:root`, overrides in
`[data-theme='dark']`, and a `prefers-color-scheme` media-query fallback for
system mode).

1. **UI chrome colors must use semantic tokens.** Never write raw `#hex` /
   `rgb()` in renderer CSS rules or chrome-related inline styles — reference
   `var(--surface)`, `var(--text)`, `var(--hover)`, etc. from
   `packages/ui/src/tokens.css`. Raw values are allowed only on custom-property
   definition lines (`--x: #...;` — token, accent, or app-scoped variable
   definitions). CI enforces this for new/changed renderer CSS lines
   (`tools/check-theme-colors.mjs`).
2. **Every new token gets both values.** Adding a token means adding it to all
   three blocks in `tokens.css` (light, dark, system-dark fallback).
3. **Accent colors stay per-app.** Each app defines `--accent` /
   `--accent-dark` / `--accent-soft` (and its dark-adjusted values) in its own
   `styles.css`. Shared rules reference `var(--accent)` and inherit the app's
   brand color.
4. **Document content is never re-authored by the theme.** Page surfaces, cell
   fills, slide content, PDF page bitmaps, export/print stylesheets, chart
   palettes, highlight color maps, stamps, and WordArt presets are document
   data: they stay hardcoded, must not reference chrome tokens, and every
   save/export/print path must produce identical output in both themes. A
   Word/Excel-style _dark page_ (Sheets via Univer's `darkMode`, Docs via
   `apps/docs/src/renderer/editor/dark-page.ts`) is a display-time remap only:
   the authored color stays the real declaration, the remapped twin lives in
   a screen-only `--dk-*` / `.page-dark` layer, and print/export never see it.
5. **Canvas-drawn UI affordances go through a constants table.** Konva/canvas
   editing chrome (selection frames, guides, handles) reads from the app's
   canvas color table (e.g. `canvas-colors.ts`) keyed by the current theme —
   no inline hex in draw calls.

## Build gotchas

- App main-process code (`apps/*/src/main`) is compiled into the **shell**
  build. After changing it, rebuild the shell or the change silently does not
  run.
- In dev mode, preload changes require a rebuild — a stale preload leaves the
  renderer blank.
- Workspace packages listed in an app's `dependencies` must also be added to
  the `externalizeDepsPlugin` `exclude` list, or the packaged app crashes on
  launch.
- `useI18n()`'s `t` is not referentially stable; never put it in a hook
  dependency array. Store the key and translate at render time.

## Docs commands (mandatory)

The Docs menu bar, ribbon and right-click menu run one set of commands, so an
action behaves the same wherever it is started.

1. **One id per action.** Register it in `apps/docs/src/shared/commands.ts`
   and give it a handler in `apps/docs/src/renderer/commands/menu-commands.ts`.
   What the handler needs from the app goes on `CommandHost`.
2. **Entry points run the command.** A ribbon or right-click control calls
   `runCommand(id, payload)` (or `<RbCommand cmd="…">`); it does not reach
   into the editor itself. CI enforces this for new and changed click
   handlers in `Ribbon.tsx`, `ribbon-*.tsx` and `ContextMenu.tsx`
   (`tools/check-docs-commands.mjs`).
3. **Not everything is a command.** Opening a control's own dropdown or dialog
   is UI state. A handler that is neither takes a
   `// command-exempt: <why>` comment on one of the three lines above it.

## UI strings (i18n)

- Large dictionaries are sharded per locale: `i18n/strings-<domain>.ts` is a
  thin aggregator over `i18n/<domain>/<lang>.ts` (one file per language, `zh`
  defines the key set). Add a new key to `zh.ts` and to every sibling shard;
  the `satisfies Record<keyof typeof zh, string>` on each shard turns a
  missing or extra key into a type error. Never grow the aggregator back into
  a single 19-locale object.
