# Contributing to ForaOffice

<!-- Modified for ForaOffice from GenOffice contributor guidance. -->

Thanks for your interest in contributing. This document covers the local
setup, the checks a change must pass, and the conventions used in this
repository.

## How changes land here

Development and issue tracking happen in [lanbingtech/foraoffice](https://github.com/lanbingtech/foraoffice).
This fork does not use GenOffice's private-tree snapshot/import workflow. Discuss
bugs and proposals in GitHub Issues. Maintainers implement accepted work on
focused branches, review the resulting changes and validation, then integrate it
into the target branch. Do not infer that an upstream PR or release contains a
fork-specific change.

Preserve contributor authorship in commits and integrations. An issue is complete
only when its acceptance criteria have evidence and the reviewed source, tests and
documentation are committed; link the relevant commit or integration in the issue.
Use pull requests when the maintainer's chosen integration requires them, not as a
second request tracker. Existing issue labels describe readiness, not permission
to perform an unapproved external action.

The current sensitive-file owner is [@turbinefish](https://github.com/turbinefish),
as listed in CODEOWNERS. A CODEOWNERS entry requests review; enforcement also
depends on repository rules and must not be assumed enabled.

## Repository layout

- `apps/*` — the seven Electron apps (docs, sheets, slides, pdf, markdown, html, shell).
  Each app is an npm workspace with its own `src/main` (Electron main
  process), `src/renderer` (React UI), and `tests/`.
- `packages/*` — pure TypeScript engine and shared packages (no Electron
  dependency, unit-tested): docx/pptx engines, AI agent core, providers,
  i18n, UI kit.
- `apps/sheets/native/xlsx-engine` — Rust xlsx engine (runs as a sidecar process) for xlsx import/export.

## Product neutrality

ForaOffice is a general-purpose Office editor intended for integration by any
host and operation by any AI agent through standard protocols such as MCP.
Currently implemented adapters and tested clients are a supported subset, not a
restriction on the product's intended consumers. Distinguish this goal from
combinations actually verified; portable integration is currently implemented
for Docs only.

Keep document editing, fidelity, capability discovery and operation results
independent of host identity, model provider and industry workflows. Hosts own
business rules, domain syntax, task orchestration and publication policy; expose
their checks through generic extension contracts. Technical object protection
and host authorization are distinct. Fix integration defects at the responsible
shared boundary after checking the authoritative standard; consumer-specific
branches and private agent handshakes do not belong in the editor core.

## Docs SDK contribution boundary

Only Docs currently has the portable host/runtime work. Read the
[SDK guide](apps/docs/sdk/README.md), [integration overview](docs/docs-sdk-integration.md)
and [delivery policy](docs/docs-sdk-releases.md) before changing its public surface.
The editor must not acquire a dependency on a particular model, conversation API,
host application's storage schema or private IPC protocol. Inject native and
application services at the existing host boundary; use standard MCP at the public
agent transport boundary.

When changing a public editor handle, host service, tool, scope/lifecycle behavior,
file/ribbon extension or instruction hook, update its authoritative SDK reference,
version/compatibility notes and affected examples. Reuse public-interface tests.
Preserve untouched document data and the distinction between live unsaved content
and file-level processing. UI chrome uses semantic theme tokens; theme changes
must not rewrite authored document colors or alter export/print output. New UI
strings belong in every locale shard, not in an enlarged dictionary aggregator.

Public contributor guidance must be available in tracked documents. Local agent
configuration is intentionally ignored and must not be published wholesale or
made a prerequisite for following this guide. Specs, task graphs, progress and
acceptance evidence live in GitHub Issues; do not add repository-local process
ledgers or temporary investigation files. Leave unrelated user changes intact and
account for task-owned tracked/untracked files before committing.

External test environments are read-only unless the user explicitly approves a
specific write. This includes login/session creation, configuration, deployments,
container changes and cleanup. A local development request is not such approval.

## Engine packages

All pure TypeScript, no Electron dependency, unit-tested (except the UI kit):

- `packages/docx-engine` — docx parsing → block tree (with `docxIndex`
  anchors and passthrough), OOXML fragment generation, byte-level paragraph
  patching.
- `packages/pptx-engine` / `packages/pptx-render` — pptx model and rendering.
- `packages/pdf2docx` — local PDF → DOCX conversion: PDFium character-level
  extraction, pure-geometry layout analysis, rebuild through `docx-engine`;
  the same analysis drives the PDF app's PowerPoint and Excel exports.
- `packages/html2docx` — local HTML → DOCX conversion: the page is rendered in
  the app's own Chromium, reduced in-browser to a document intent tree, and
  written as native OOXML with the `docx` library; only visuals with no Word
  counterpart are screenshotted. Drives the HTML app's Export as Word.
- `packages/hwp2docx` — local HWP/HWPX → DOCX conversion: HWPX (OWPML)
  parsed directly, binary `.hwp` read through rhwp (MIT, WebAssembly) into
  HWPX first; character and paragraph shapes, tables, pictures, sections,
  headers and footers rebuilt through `docx-engine`, anything else counted
  in `dropped`. Docs opens `.hwp`/`.hwpx` through it as converted `.docx`
  copies, like `.odt`.
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

### Architecture notes (docx round trip)

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

## Getting started

Prerequisites: Node 24+ for Docs SDK build/consumption, npm 10+, and a Rust toolchain (`cargo` on PATH,
needed only for the sheets xlsx sidecar).

```bash
npm install
npm run fixtures     # generate test .docx fixtures (one-time, and after docx-engine changes)
npm run dev          # all editors + shell against Vite dev servers
npm run dev:docs     # or run a single app
```

## Checks every change must pass

Run the checks relevant to your change locally; CI runs the configured repository gates. Docs SDK metadata/build changes also require the package and host checks described in the delivery guide:

```bash
npm run format:check # Prettier check for uncommitted changed/new files
npm run lint         # ESLint across the repo (0 errors required; warnings allowed)
npm run typecheck    # tsc --noEmit across every workspace
npm test             # engine + app unit tests (also runs the Rust sidecar tests)
npm run licenses     # production dependency licenses within the permissive allowlist
```

Formatting is intentionally incremental: existing files are not reformatted
unless they are part of your change. Run these exact commands before committing:

```bash
npm run format                              # format uncommitted changed/new files
npm run format:check                        # verify uncommitted changed/new files
npm run format:check -- --base origin/main  # verify committed files on your branch
```

CI supplies the PR or push base automatically and checks only files changed from
that base. This keeps the formatter gate useful without creating a repository-wide
formatting diff.

## Building installers

Run these from the repository root — they regenerate the third-party
notices and build all seven apps before packaging:

```bash
npm run dist:mac   # dmg + zip
npm run dist:win   # nsis installer
```

Without Apple or Windows signing credentials in the environment these produce
unsigned artifacts: code signing and notarization are skipped with a warning
rather than failing. That is the expected result for a contributor build.

On macOS, packaging from a repository that lives on an exFAT/FAT32 volume (an
external USB drive, for example) fails because the OS writes hidden `._`
AppleDouble sidecar files next to the build output and electron-builder trips
over them. Point the output directory at an APFS path instead of moving the
repository:

```bash
BUILD_DIR=/tmp/genoffice-release npm run dist:mac
```

`dist:win` takes the xlsx sidecar from the MinGW cross-compilation path when it
is there (that is what the release pipeline builds). On a Windows host the plain
`npm run build:all` compiles it with the host toolchain into
`target/release/xlsx-sidecar.exe` instead, and packaging accepts that build as
long as the host architecture matches the packaged one (`GENOFFICE_WIN_ARM64=1`
still needs a cross build for `aarch64-pc-windows-msvc`).

## Environment variables

None are required — the apps run with all of these unset. They exist for
testing and local overrides:

| Variable                                                                                                               | Effect                                                                                      |
| ---------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------- |
| `BUILD_DIR`                                                                                                            | Override the electron-builder output directory (default `apps/shell/release`)               |
| `GENOFFICE_USER_DATA`                                                                                                  | Override the Electron userData directory (test isolation)                                   |
| `GENOFFICE_LANG`                                                                                                       | Force the UI language instead of following the OS locale                                    |
| `GENOFFICE_FAKE_UPDATE`                                                                                                | Exercise the updater UI without a real release feed                                         |
| `GSK_API_KEY`, `GSK_CLI_PATH`                                                                                          | Genspark credentials / CLI location for the built-in AI provider                            |
| `AI_SEARCH_DISABLE_GSK`, `SERPER_API_KEY`, `SERPLY_API_KEY`, `SEARCHAPI_API_KEY`, `TAVILY_API_KEY`, `PARALLEL_API_KEY` | Disable the gsk search backend / supply a Serper, Serply, SearchApi, Tavily or Parallel key |
| `XLSX_SIDECAR_PATH`, `XLSX_OPEN_PATH`, `XLSX_DEBUG_PORT`                                                               | Point at a locally built xlsx sidecar and its debug port                                    |
| `*_DEV_PORT`, `*_RENDERER_URL`                                                                                         | Per-app Vite dev server ports and renderer URLs (set by `npm run dev`)                      |

In the inherited desktop consumer, AI features degrade rather than break without credentials: requests surface an
inline sign-in prompt, and web search falls back to keyless backends (Parallel's
free Search MCP, then DuckDuckGo).

## Coding conventions

- **English only** in code, comments, commit messages, and docs. User-facing
  strings go through the i18n resources (`src/renderer/i18n/`, plus the inline
  main-process dictionaries in `src/main/`), which are the only places
  non-English text belongs (plus test fixture text).
- TypeScript everywhere; avoid adding new `any` surfaces where a precise type
  is cheap.
- Tests live in `apps/*/tests` and `packages/*/tests` (vitest). New engine
  behavior needs a unit test; renderer-only UI tweaks generally don't.
- Keep ad-hoc Office comparison drivers and sensitive repro files local and ignored. Maintained SDK/browser/Electron integration tests belong in the existing test suites and CI; temporary evidence belongs in ignored build output or outside the repository.
- Keep files from growing without bound: if you are adding a substantial new
  concern to an already-large file, prefer a new module.

## Commit and PR guidelines

- Small, focused commits with imperative English subject lines
  (e.g. `fix docx table border round-trip`, `add slides chart legend parsing`).
- A PR should explain _why_ the change is needed, and mention which of the
  checks above you ran.
- File format fidelity is the core product promise: for changes touching
  open/save paths (docx/xlsx/pptx), include a round-trip test proving
  untouched content survives byte-for-byte.

## Reporting bugs and requesting features

Use the issue templates. For suspected security issues, do **not** open a
public issue — follow [SECURITY.md](SECURITY.md).

## Code of conduct

All community spaces follow the
[Contributor Covenant](CODE_OF_CONDUCT.md); participation implies acceptance.

## License and CLA

There is no CLA. Contributions are accepted under the [Apache License 2.0](LICENSE)
that covers the open-source project: inbound = outbound, per Apache-2.0 section 5.
Contributors retain their copyright; do not replace upstream ownership notices or
invent a legal entity from a GitHub account name.

Retain applicable upstream, dependency and font notices in source and distributed
artifacts. Modified upstream files must prominently state they were changed, as
required by Apache-2.0 section 4; a project changelog alone is not assumed to satisfy
that file-level requirement. Preserve the license text itself. See [NOTICE](NOTICE)
for origin and distribution attribution. Product branding uses ForaOffice;
GenOffice/Genspark names remain appropriate for origin and compatibility references.

The unused upstream enterprise placeholder has been removed from the fork's active
policy and payloads. This does not relicense upstream enterprise material or any
third-party component. Any separately licensed material introduced later must have
its boundary explicitly reviewed and disclosed.
