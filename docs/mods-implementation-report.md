# Mod support in GenOffice — what was built, and how

This is the closing report for the feature-assistant work: every kind that was
implemented, the mechanism behind each one, what is an interface versus what is host
code, where the script interface actually works today, and what is left.

The design lives in docs/mods-design.md; the nineteen changes are documented
one by one in docs/mods-pr1-mod-interface.md … docs/mods-pr19-packs.md.

---

## 1. The shape of the thing

Two assistants, deliberately separate:

|                                           | who                                             | may do                                                                                                      |
| ----------------------------------------- | ----------------------------------------------- | ----------------------------------------------------------------------------------------------------------- |
| **feature assistant** (the mod assistant) | the panel on Home (Cmd/Ctrl+K for its commands) | add, update and delete features in a registry; it cannot touch a document and has no write tool             |
| **native assistant** (each app's own)     | docs, sheets, slides, pdf, markdown, html       | use the features: list them, dry-run a script feature read-only, read a knowledge source; **no apply tool** |

A **feature** is an extension. It is one JSON record on disk with a `kind`:
`slot`, `view` (inside a slot), `overlay`, `appearance`, `script`, `hook`,
`knowledge`, `command`. A "mod" is a feature, or a pack of them.

The dividing line, unchanged through all nineteen changes:

> **Nothing a feature does may change document rendering, saving, exporting or
> printing.** Features draw chrome, react to events, run read-only code, offer
> material to read, and register actions. Writes happen only where the user asked for
> them, through the engine's own operation path — never through a feature's own
> authority.

## 2. Where the code lives

| Layer           | Path                                                                                                                                  | What it is                                                  |
| --------------- | ------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------- |
| contracts       | packages/agent-core/src/feature-*.ts, slot-overlay.ts, feature-script.ts, feature-hooks.ts, feature-knowledge.ts, feature-commands.ts | validation and vocabulary, free of electron, unit-tested    |
| shared renderer | packages/ui/src/mods/*                                                                                                                | one slot/overlay renderer for the shell and all six editors |
| host            | apps/shell/src/main/feature-*.ts, script-host.ts, features-ipc.ts                                                                     | the only place features are read, run, mounted or written   |
| panel           | apps/shell/src/renderer/src/ModsAssistant.tsx, CommandPalette.tsx                                                                     | the feature assistant's UI                                  |
| assistants      | apps/{docs,sheets,slides,pdf,markdown,html}/src/renderer/ai                                                                           | the native assistant's tools (`createModsSkill`)            |

Every host capability is registered once in the **shell main process** (which is the
process that serves all the app IPC channels), and reached from a renderer only
through that app's own preload bridge.

## 3. The eight kinds

| kind                | contract              | host behaviour                                                                                                                                                                                                                                                                                                                                                                                                                                                       | where the user sees it                                  |
| ------------------- | --------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------- |
| **slot** (K3)       | feature-slots.ts      | host draws declared elements (text, button, input, select, list, file-list, view, overlay, divider, markdown, group) with its own components                                                                                                                                                                                                                                                                                                                         | home sidebar/card/panel, six editor rails               |
| **view** (K5)       | feature-slots.ts      | a real `WebContentsView` mounted over the rectangle a slot reports; https anywhere, http only on loopback; sandboxed, no preload, window.open goes to the browser                                                                                                                                                                                                                                                                                                    | a browser panel a feature declared                      |
| **overlay** (K4)    | slot-overlay.ts       | an inert SVG layer over the content area: shapes in fractions, tones from the host palette, no pointer events                                                                                                                                                                                                                                                                                                                                                        | annotations/guides                                      |
| **appearance** (K8) | feature-appearance.ts | re-points chrome variables by name (background, accent family) and sets the markdown reader's paper theme (`markdown`: one of the app's own picker themes — notion-light/dark, drake-dark, blubook, nord, newsprint — or a custom { paper, ink, accent?, font? } the validator checks for WCAG contrast; the user's ribbon pick always wins, a feature decides only while the ribbon is on default; screen media only; text, saving, print and export never touched) | panel, command palette, rails, the md reader's paper    |
| **script** (K1)     | feature-script.ts     | read-only dry run, and a confirmed apply, in a `node:vm` sandbox                                                                                                                                                                                                                                                                                                                                                                                                     | panel's Scripts section; the native assistants' dry run |
| **hook** (K2)       | feature-hooks.ts      | runs on `document.opened`, `document.closed`, `document.will-save`; the last one can **veto a save**                                                                                                                                                                                                                                                                                                                                                                 | panel's Hooks section                                   |
| **knowledge** (K7)  | feature-knowledge.ts  | lists a folder/file source and reads one entry by name, through the document reader and its bounds                                                                                                                                                                                                                                                                                                                                                                   | the assistant reads it on demand                        |
| **command** (K6)    | feature-commands.ts   | an action in the Cmd/Ctrl+K palette, with the same action vocabulary a slot button has                                                                                                                                                                                                                                                                                                                                                                               | the palette                                             |

## 4. Mechanism notes worth keeping

- **The registry is three layers of validation.** The tool validates, the store
  validates again on write and again on read, and the IPC boundary validates every
  payload from a renderer. A malformed record cannot reach the renderer.
- **Actions are one vocabulary.** `open`, `ask-assistant`, `panel`,
  `run-feature` — used by slot buttons, lists, inputs, commands. A host that cannot
  serve an action draws the control **disabled** rather than letting it fail.
- **Slot data is host data.** A `file-list` names a source (`recent-files`,
  `starred-files`); the host reads it and substitutes only `$path`, `$name`,
  `$folder` into the action. A slot never holds a filesystem handle.
- **The save veto is central.** The guard wraps `ipcMain.handle` **once, early**,
  inspects only save-ish channel names, and calls straight through when no hook
  exists. A broken hook is ignored; a veto becomes the invoke's rejection, so the
  writer is never reached. No app changed.
- **Knowledge entries are names, not paths.** Only a name the listing returned
  resolves, and only inside the declared folder; `../secret.md` and absolute escapes
  are refused.
- **Packs are untrusted input.** Every record is re-validated on the way in, the
  version is checked, the review names what a pack declares (including "can stop a
  save" and "runs code") before anything is written, and an import never replaces an
  existing feature unless asked. Imports keep each feature's **own id** — a second
  import skips rather than duplicating.

## 5. Scripts: interface, host, and what is _not_ there

This is the part worth being precise about.

**There are no built-in script features.** Nothing is seeded: the registry directory
(`userData/features`) starts empty, and the only code that creates a record is the
store when the panel or an import asks it to. A script is a feature _the user's own
feature assistant wrote_, and there is no CLI `script` command either.

**What is interface, and what is host code.** The _interface_ is the contract
(feature-script.ts: source limits, module syntax refused, the `$.docs` API summary)
plus the channels. The _host_ is new native code in the shell main
(script-host.ts): a `node:vm` sandbox where the API is built **inside** the context
from plain JSON, code generation off, no `Promise`, no timers, the whole run
synchronous under a vm timeout. That host is not a set of features; it is the thing
the interface is an interface _to_. It was **revived from parked work** — the same
sandbox once existed as an unmerged CLI command, preserved in
docs/mods-kind-script.md — and moved into the shell main, where a document can be
opened headlessly and a registry entry can be run on demand.

**Two capabilities, two gates.** `runScriptReadOnly` (dry run) never writes, and
`applyScriptToDocument` writes only when the caller passes `confirm: true`, the
engine accepts the whole batch atomically, and the file is not open in an editor.

### Where the script interface works today

The host dispatches on the document's extension (script-shared.ts): `.docx` →
`$.docs`, `.xlsx` → `$.sheets`, `.md`/`.html` → `$.text`, `.pptx` → `$.deck`,
`.pdf` → `$.pdf`. A feature may declare `script.format`; a run against a file of
another format is refused with the mismatch named. The reads each format exposes
are bounded and refuse over the cap with a reason.

| app                | reads the registry | dry run (`features:run`) | apply (`features:apply`)            | assistant tools (`createModsSkill`)         |
| ------------------ | ------------------ | ------------------------ | ----------------------------------- | ------------------------------------------- |
| shell (Home panel) | yes                | yes                      | **yes** (the user's two-step Apply) | — (it composes the _feature_ skill instead) |
| docs               | yes                | yes                      | no                                  | yes                                         |
| sheets             | yes                | yes                      | no                                  | yes                                         |
| slides             | yes                | yes                      | no                                  | yes                                         |
| pdf                | yes                | yes                      | no                                  | yes                                         |
| markdown           | yes                | yes                      | no                                  | yes                                         |
| html               | yes                | yes                      | no                                  | yes                                         |

So: a feature's script can be **dry-run** from any of the six editors' assistants,
and **applied** only by the user, in the panel, against a document they pick. What
changed with per-format support is _what the dry run and the apply can act on_:

- **sheets** (`$.sheets`): `sheets()`, `usedRange(sheet)`, `range("Sheet1!A2:F200")`
  as a 2-D value grid, `formulas(sheet?)`, `formats(range)`, `find(query)`, and
  `propose(ops)` with the workbook DSL (the in-app `propose_operations`
  vocabulary). The snapshot is read straight from the package with the gateway
  (no native sidecar), capped at 50 000 non-empty cells; a run over the cap is
  refused with the count. The write goes through `runWorkbookDsl` +
  `writeWorkbook` — atomic, formulas evaluated like any other save; the dry run
  reports a diff summary (cells affected, then the first op lines).
- **markdown / html** (`$.text`): `get()` and `lines()`; the script _returns_ the
  new full text and the host makes one whole-document edit, reported as a line
  diff summary. No engine, no ops.
- **deck** (`$.deck`): `slides()` / `elements(slide)` with real EMU geometry and
  text, and `propose(ops)` with the slides app's own op vocabulary
  (`setText`, `addElement`, …) — one dialect, not a second one. The run is
  `runTxn` in atomic isolation; an apply saves and reports the layout audit
  (overlap / out-of-bounds / overflow) with the result.
- **pdf** (`$.pdf`): read-side only, on purpose — `info()` and
  `pages(from, to, maxChars?)` (snapshot holds up to 200 pages × 4 000 chars).
  The script returns a report and the run hands it back as `result`. There is no
  apply path yet: pdf writes would need the shared headless pdf ops/save module
  and a bounded batch input (§8.2.6), which still do not exist.
- **docs** (`$.docs`): unchanged, and still the reference for the two gates.

### What is still true, stated plainly

- from an editor's assistant, `dry_run_mod_feature` now produces a meaningful
  plan for every format it reads — including spreadsheets, which was the gap
  that motivated §8;
- writes still land only through the user's Apply in the panel (or the docs app's
  own engine), never through a model;
- the sandbox is synchronous in the shell main, so the cell caps and the vm
  timeout are the budget; a utility process is the named next step (§8.2.2);
- `$.files` (bounded batch input across several files) is not exposed anywhere.

## 6. Verification

Final state of the repository after the last change:

| check                                 | result                                                                                                                                          |
| ------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------- |
| `npm run build:all`                   | **BUILD_EXIT:0**, 0 errors                                                                                                                      |
| packages/agent-core tests             | **229 passed**                                                                                                                                  |
| apps/shell tests                      | **861 passed** (88 files)                                                                                                                       |
| typecheck, each package on its own    | agent-core 0, ui 0, cli 0, docs 0, sheets 0, slides 0, pdf 0, markdown 0, html 0, shell **0 in its own sources**                                |
| theme / english-only / prettier gates | pass                                                                                                                                            |
| UI evidence                           | the preview harness renders the real components and stylesheets with headless Chromium and (for interactive parts) drives them with real clicks |

The shell's whole-project typecheck also reports 18 errors in files this work never
touched (`../docs/src/renderer/ai/tools.ts`, `../docs/src/renderer/editor/extensions.ts`,
`../docs/src/renderer/i18n/locale.tsx`, `packages/cli/src/dom.ts` missing
`@types/jsdom`) — pre-existing, and absent when those packages are typechecked
themselves.

**What could not be verified on this machine.** The built Electron app stops
answering about two seconds after launch here (an unclicked probe confirmed it, and
it predates this work), and `page.screenshot` returns a blank frame. So: mounting a
real `WebContentsView`, a real save under the hook guard, the real file dialogs, and
an end-to-end run through the packaged app were **not** exercised. Those paths are
covered by unit tests over their decision logic and by the IPC payload shapes, and
the gaps are named in the relevant PR documents rather than papered over.

## 7. What is left, in the order I would do it

Per-app script support is specified in §8; the order below follows that analysis.
Items 1–5 are done as of the per-format script host (format dispatch, bounded
per-format reads, the sheets write path, `$.text`, the deck convergence and the
pdf read side); what remains of each is named.

1. **Sheets scripts first** — the real pain: repetitive or complex computation that the
   assistant can only do today by looping `read_range` and `propose_operations`.
   Needs format dispatch, a range-shaped read API, a bounded write path, and execution
   off the main process (§8.2.2). _Done except the utility process: the caps and the
   vm timeout are the budget until then._
2. **Markdown next** — the cheapest place to prove the win: text in, text out, no
   engine, and the app already has a whole-document write. Do it early enough to
   measure "N tool calls vs 1 script call" for every later format (§8.2.3). _The
   interface is done (`$.text`); the measurement is still to produce._
3. **Html** — same class as markdown, with structure (§8.2.4). _Done through the
   same `$.text` channel; a structural `$.html.nodes()` read was deliberately not
   added._
4. **Slides — converge, do not add.** The app already has a script tool
   (`execute_slide_script`) with read-write combined, one atomic transaction, one undo,
   rollback and a layout audit. A mod script must target that same environment rather
   than introduce a second dialect (§8.2.5). _Done: the deck scripts propose the
   slides app's op vocabulary, run atomically and report the layout audit._
5. **Pdf — different axis, not "no value".** Its win is the read side (condensing a
   long document into a short report) and cross-file batch work, which needs a bounded
   batch input the sandbox does not have today (§8.2.6). _The read side is done
   (`$.pdf` + the report as the run's result); cross-file batch and the write side
   remain._
6. **Fold the knowledge kind into the knowledge base**, then extend the base
   (more source kinds, per-source controls, an index with freshness, better retrieval,
   citations that survive) — §9. Nothing is implemented there yet.
7. **Let an editor's assistant apply, behind the user.** The channel exists; an apply
   tool for a model should come only with an explicit confirm in the conversation, so
   the plan is read before the write.
8. **More hook events, each with a real emitter.** `document.saved` (after a save),
   `selection.changed`, `document.exported`. An event nobody raises is not offered —
   the same rule that keeps `ribbon.button` out of the place list.
9. **The remaining places: ribbon, menus, status bar.** They need each app's ribbon
   metadata, and a place is only added to the contract once a host draws it.
10. **Pack provenance.** Signatures or a provenance field; today a pack says what it
    declares, not who wrote it.
11. **Per-feature permissions in the panel.** One screen that lists, per feature: which
    paths it reads, which events it observes, whether it can stop a save.
12. **Sandbox hardening.** The current sandbox is a capability boundary for code the
    user or their assistant wrote, not a hardened boundary for hostile code; that needs
    a separate process and a stricter channel.
13. **Marketplace/distribution.** A different problem from trust, and out of scope so far.
14. **Smaller polish.** Shared-package labels are English with host overrides; the
    panel's own empty-state copy could name the kinds it supports; the pack review could
    offer "select which features to install".

## 8. Script support per app: value, and what to expose

**Status: implemented, one format short of this analysis.** The host dispatches on
the document's extension and serves each format's vocabulary (see §5): sheets,
markdown, html, deck and pdf read-side are live; the pdf _write_ side is not (it
needs the shared headless pdf ops/save module and a bounded batch input), and the
utility process of §8.2.2 is still the named next step. The decisions below are
kept as written — they are the reasons behind what shipped.

### 8.1 The criterion is not "which app", it is "what is the assistant forced to do"

All six assistants already have a **granular operation channel**: `apply_ops` in
slides, pdf, markdown and html; `propose_operations` in sheets; the script kind's own
propose/apply in docs. So every app suffers the same thing — repetitive work becomes N
tool calls, N model round trips, N chances to drift. What differs is only:

| question                                    | why it decides the shape                                                                                                                                                                    |
| ------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| structure or text?                          | a text format needs **no engine**: the script returns the new text and the app applies one edit. A binary format needs a headless engine to apply operations.                               |
| how much is already covered by batch tools? | pdf already ships whole-file operations (`extract_pages`, `merge_pages`, `split_pdf`, `set_watermark`, `set_header_footer`, `crop_pages`…); a script that re-does them is only indirection. |
| is the expensive part reading or writing?   | for pdf it is **reading** (a 200-page page-by-page read into context is ruinous); for sheets and slides it is usually writing.                                                              |

A script's value is therefore never "a new capability"; it is **collapsing a loop out of
the token loop into one deterministic call**, with a reviewable result.

### 8.2 Per app

#### 8.2.1 docs — the reference implementation (shipped)

- Good for: bulk document edits proposed as operations and validated by the engine —
  heading levels, find/replace, TOC, fields, table operations, styles.
- Today: `$.docs.blocks()` / `$.docs.text()` / `$.docs.propose(ops)`; a read-only dry
  run (`features:run`, exposed in all six editors and the shell) and a write
  (`features:apply`) that requires an explicit confirm, an atomic engine batch, a file
  not open in an editor, and a one-step replace.
- What the other formats should **copy**: the dry-run-before-apply rule, the two gates,
  three-layer validation, and the panel's Scripts section (pick a document → dry run →
  Apply).
- What they should **not** copy: the `$.docs` _dialect itself_. Slides already
  embodies a better contract (§8.2.5; §8.3.1).

#### 8.2.2 sheets — highest value

- Genuinely good for: repetitive or complex work over ranges — cleaning, normalising,
  deduplicating, cross-sheet joins, one-off analysis, and **generating** formulas across
  a range instead of computing the same thing N times.
- Must **not** be the tool for: row-local, stable arithmetic. A formula is auditable,
  recomputes when inputs change and undoes cleanly; a script that writes computed
  values goes stale. The script API must therefore be able to write **formulas**, not
  only values.
- Foundation to reuse (already in the tree): `packages/cli/src/formats/xlsx.ts`,
  `xlsx-dsl.ts`, `xlsx-gateway-ops.ts`, `xlsx-batch-order.ts`, `xlsx-catalog.ts`,
  `xlsx-check.ts`; `@genoffice/xlsx-gateway` (`./*` is importable); the app's own
  operations through `propose_operations`.
- What to expose to the script:
  `$.sheets.sheets()`, `$.sheets.usedRange(sheet)`, `$.sheets.range('Sheet1!A2:F2000')`
  (values), `$.sheets.formats(range)`, `$.sheets.find(query)` — every read **bounded**
  (cells and rows), refusing over the cap with a reason.
- What the write side must be: the gateway's own operations (values, formulas, formats),
  applied **atomically**; the dry run returns a **diff summary** (cells/rows affected
  plus the first N changes), not a plan list — a wrong write in a workbook is much
  costlier than in a text document.
- Operational requirement: **run off the main process.** The shell main serves every
  app, and the sandbox is synchronous, so a 100k-cell loop would freeze every window.
  Hard milliseconds budget plus a cell cap first; a utility process (message protocol)
  next.

#### 8.2.3 markdown — cheapest, and the place to prove the mechanism

- Genuinely good for: deterministic document-wide transforms — TOC generation, heading
  renumbering, table alignment, link checking, frontmatter batch edits.
- Must **not** be for: prose rewriting (the model does that well already).
- Foundation: none needed. The app already has a whole-document write
  (`write_document`) and an ops channel (`apply_ops`).
- What to expose: the document as text (`$.text`), and either the new text back (one
  edit, easy diff) or proposed ops (more diffable). Both are acceptable; pick one and
  keep it consistent with html.
- Why first among the cheap ones: it is the fastest way to produce the evidence that
  justifies the rest — one real repetitive task measured as N tool calls versus 1 script
  call.

#### 8.2.4 html — same class as markdown, with structure

- Genuinely good for: batch element/structure edits, injecting or refreshing a block
  across many places, extracting or inserting tabular data, link/attribute sweeps.
- Foundation: `read_source`, `write_document`, `apply_ops`, `get_outline` already
  exist; the source is text, so again no engine is required.
- What to expose: the source as text plus the same write channel chosen for markdown.
  If a structural read is wanted (`$.html.nodes(selector)`), it must be bounded and
  read-only; the write stays textual/op-based.

#### 8.2.5 slides — converge, do not add a second dialect

- The app **already has a script tool**: `execute_slide_script` in
  `apps/slides/src/renderer/ai/slides-skill.ts`, described there as the preferred way
  to edit a page's existing elements. Its contract is the one to converge on:
  read-write combined (the script receives the real geometry and text of every element,
  so no read call first), one **atomic transaction** with rollback, **one undo step**,
  a constrained synchronous DSL with no ambient globals, and a **layout audit**
  (overlap / out-of-bounds / text overflow) returned with the result.
  `apps/slides/src/renderer/ai/slide-qc.ts` already runs a quality loop whose tool
  allow-list is `read_slide` + `execute_slide_script`.
- Genuinely good for: a brand pass over N slides, notes, consistent re-layout, decks
  built from data.
- Foundation for a deck-wide (not page-wide) script:
  `packages/cli/src/formats/pptx.ts` — `openDeck`, `blankDeck`, `saveDeck`,
  `parseOps`, `applyOps`, `describeDeck`, `describeLayouts`.
- What to do: let a mod script target **that** environment and op vocabulary, rather
  than introducing the docx-shaped sandbox. Two script dialects inside one app is the
  failure mode to avoid.

#### 8.2.6 pdf — a different axis: read-side, and batch

- Genuinely good for:
  1. **Read-side condensation** — scan every page and return a short result (matches,
     extracted rows, anomalies). This is the affordable way to work on a 200-page file,
     and it is the case where a script changes what is _possible_, not just what is fast.
  2. **Conditional page rules** — inspect a page, then decide: markup, redact, crop,
     drop. Today this is `search_text` + `read_pages` + per-page `markup_text` /
     `crop_pages`, i.e. N round trips.
  3. **Cross-file batch** — the same watermark, header, page range or split applied to
     the N files the user picked.
- Must **not** be for: one-off operations a single batch tool already covers (merging
  three files, one watermark); anything needing visual judgement; a terminal pdf nobody
  consumes.
- Foundation: `apps/pdf/src/main/save-pdf.ts` (`savePdfToPath`) and the ops surface in
  `apps/pdf/src/renderer/ai/pdf-skill.ts` (+ `tools.ts`, `stamps.ts`,
  `ocr-layer.tsx`, `search.ts`); `pdf-lib`, `@embedpdf/pdfium`, `pdfjs-dist`. Headless
  **reading** already works: `packages/cli/src/formats/pdf.ts` exports `pdfInfo`,
  `pdfText`, `convertPdf`.
- The missing pieces are packaging, not capability: a **shared headless pdf ops/save
  module** (the docx and pptx formats have one; pdf does not), and a **bounded batch
  input** — cross-file work is inexpressible today because the sandbox deliberately has
  no filesystem access.
- Shape: closer to a **command plus a work queue** (kind K6) than to a per-document
  script.

### 8.3 Interface exposure: the handoff list

#### 8.3.1 One contract, per-format vocabularies

Define the script contract **once** and give each format a vocabulary behind it: a
constrained synchronous DSL, `$` as the only root, `$.input`, `$.log(...)`,
`$.budget`, one namespace per format (`$.docs`, `$.sheets`, `$.deck`, `$.text`,
`$.pdf`, `$.files`), `$.event.veto(reason)` only where the host is asking permission
(`document.will-save`). Semantics to keep identical across formats:

- one call = one **atomic batch**; any refused operation rolls the whole thing back;
- **one undo step** for the app;
- the run returns a **summary/audit** (what would change, plus any warning) — for
  binary formats a diff summary, for decks a layout audit;
- a dry run and a real apply run the **same executor**, so the plan means what the write
  will do.

#### 8.3.2 What the host must expose to the script

Bounded, format-specific reads (cells/rows/pages/blocks/chars caps, refusing with a
reason), the format's op vocabulary for proposals, `log`, `budget`, and — for the
batch cases — a **bounded `$.files` list** containing only the files the user selected
for this run. Nothing wider.

#### 8.3.3 What the host must expose to the app and its assistant

- `featuresList` — already exposed in six editor preloads and the shell.
- `features:run` (read-only dry run) — already exposed in the six editor preloads and
  the shell.
- `features:apply` — **deliberately exposed only in the shell today** (the panel's
  two-step Apply, pressed by the user). If an editor is given one, it must arrive with
  an explicit in-conversation confirm and still satisfy the existing gates: explicit
  confirm, atomic engine batch, and the file not open in an editor.
- **No model gets an apply tool** without that confirm. The mod assistant's tool surface
  must keep having no write path at all.

#### 8.3.4 What must never be exposed

The filesystem, the network, ambient globals, `require`/`process`, `Promise`/timers,
raw engine objects or document handles, any unbounded read, and the rendering,
printing and export paths. A script's authority is exactly the API the host built
inside the sandbox; when a format needs more reach (cross-file batch, heavier
computation), widen it through **bounded host-provided inputs or helpers**, never by
opening a general capability.

#### 8.3.5 Where execution runs, and how writes land

- Execution today is the **shell main process** with a **synchronous** sandbox, which
  serves every app — a slow script freezes all windows. Move it to a utility
  process/worker behind a message protocol before any format that can do real work
  (sheets first).
- Writes must land through the **app's own operation channel** (`apply_ops` /
  `propose_operations` / the shared format module), so undo, dirty state and
  collaborative state stay coherent. Never write a file underneath an editor that has it
  open — the docx apply already refuses that, and every format should.
- The renderer holds live editable state; do not run a script _inside_ the renderer
  unless the format chooses the slides pattern deliberately (restricted DSL in-renderer,
  atomic, one undo, audit) and documents why.

#### 8.3.6 Evidence to produce per format (acceptance)

For each format, before claiming it works:

1. one real repetitive task, measured **N tool calls vs 1 script call** — round trips,
   wall time, and tokens;
2. a dry run whose output a human can review (diff summary / plan) and which matches the
   apply;
3. a rollback proof: an operation the engine refuses leaves document state and file
   bytes unchanged;
4. a failed script (throw, timeout, budget) leaves the file untouched and reports why;
5. the applied write verified by **reading the file back** with the same reader, not by
   asserting a return value.

## 9. Knowledge: fold the mod knowledge kind into the knowledge base

**Status: implemented — all three migration phases (§9.7).** A knowledge feature
registers a source of the knowledge base: the base's corpus is the starred files
plus every source the features declare, one retrieval path, one Ask UI. The
per-source controls (enable/disable, reindex, the missing-path flag), the source
line in the model's context and the coverage line on every search are live; the
two `list_knowledge` / `read_knowledge` tools are gone from `createModsSkill`
(the host reader functions stay — the base reads through them), and the
assistants' tool surface for knowledge is the base's own search and page reads.
The record shape, the pack format and the knowledge validation are unchanged, so
packs carrying a knowledge feature still import and still validate. What remains
of §9.5: retrieval ranking quality (5), the scope restriction (8), and the index
freshness UI beyond the reindex button (3).

### 9.1 The decision

One concept, not two. The **native knowledge base** owns retrieval, page reads,
citations and the Ask UI; a mod **`knowledge` feature becomes a source registration** —
"this folder (or file) is part of my knowledge base" — instead of a parallel, smaller
knowledge base of its own.

### 9.2 Why (the reasoning, kept short)

1. **The mod system has no other sanctioned read path.** Slots, appearance and overlays
   carry no content; hooks carry events; the script sandbox deliberately has no
   filesystem. Without a source declaration an extension can only be fed content by the
   model, or make the user attach the same files again and again.
2. **Today's kind is the weaker duplicate.** It offers a list of at most 50 names and
   read-one-by-name; it has **no search, no page model, no citations and no UI** (the
   panel never mentions `knowledge`). The native base has all four.
3. **Absolute paths do not travel.** A pack carrying `/Users/me/Work/standards` means
   nothing on another machine, so "shareable knowledge" was never real.
4. **Two things named knowledge confuse everyone** — including the review that produced
   this section.

Folding keeps the capability (an extension can point the assistant at local material)
and deletes the duplication (one corpus, one retrieval path, one UI).

### 9.3 What exists on each side today

|            | native knowledge base                                                                           | mod knowledge kind                                                                                   |
| ---------- | ----------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------- |
| code       | `packages/agent-core/src/knowledge-base-skill.ts`, `apps/shell/src/renderer/src/KbAskPanel.tsx` | `packages/agent-core/src/feature-knowledge.ts`, host `apps/shell/src/main/feature-knowledge-host.ts` |
| corpus     | the reader's **starred files**                                                                  | a declared **folder** or **file**                                                                    |
| shapes     | `KbFileInfo`, `KbSearchHit`, `KbReadPage` — search hits, page reads                             | `KnowledgeEntry` names, one at a time (`KNOWLEDGE_MAX_ENTRIES = 50`)                                 |
| citations  | yes — `openSourcePath` plus the file-nav scheme (`FILE_NAV_SCHEME`)                             | none                                                                                                 |
| UI         | the Ask panel                                                                                   | none                                                                                                 |
| tools      | the knowledge-base skill                                                                        | `list_knowledge`, `read_knowledge` in `createModsSkill`                                              |
| written by | the app's own settings and flows                                                                | the feature assistant (`features-skill.ts`), or a pack                                               |

### 9.4 Target design

- **The corpus is a set of sources.** `starred` stays the built-in source; `folder`
  and `file` sources arrive from a knowledge feature. Retrieval never cares where a
  source came from.
- **Because retrieval and citations are native, a mod-provided folder gets search, page
  reads and clickable citations for free.** That is the point: the mod contributes a
  _source_, not a second engine.
- **One writer per source, one reader for all.** The feature declares
  `{ source, path, label }`; the host validates it (unchanged), and the knowledge base
  reads it through the reader the mod host already has (`listKnowledgeEntries`,
  `pathForEntry`, `readKnowledgeEntry`). The traversal guard, the entry cap and the
  read bounds stay where they are.
- **Do not rename the kind.** `knowledge` is part of the pack format, and renaming it
  would force a pack version bump for no user-visible gain. Keep the record shape; call
  it a **source** in the UI and in documentation.
- **Retire the duplicate tools, not the capability.** Recommendation: the knowledge
  base's search and read tools become the single path for the model, and
  `list_knowledge` / `read_knowledge` leave `createModsSkill` in a later step
  (§9.7 phase 3). Keeping both means two answers to "what does the assistant know",
  which is the problem being fixed.
- **Existing knowledge features keep working.** An existing feature becomes a source; a
  pack that carries one keeps importing, because validation is unchanged.

### 9.5 Extending the knowledge base

The point of folding is not only to delete code; it is to make the knowledge base worth
folding into. Ordered by value per unit of work:

1. **More source kinds.** Starred files (exists), a declared folder, a single file.
   Later, if there is a reason: the documents currently open in the app, and the files
   attached to the conversation.
2. **Per-source controls.** Enable/disable, include/exclude globs, a per-source file cap
   and byte cap, and an explicit "reindex". A disabled source is not searched and says
   so in the answer's coverage line. _Shipped: enable/disable with the coverage line,
   the globs — a feature's `knowledge.include` / `knowledge.exclude` name patterns
   narrow what a source contributes, enforced in the host listing (so an excluded name
   also stops resolving) — and the per-source index mode: `knowledge.index` is
   `full-text` (default; content in the search index) or `names` (files listed,
   readable and citable, content never indexed — the coverage line reports such a
   source as not searched). Decision recorded: retrieval **invocation** stays
   knowledge-base-only (the search tools exist only in the base's skill); the mod
   assistant shapes how a source is indexed. When the base ever gains semantic
   retrieval, the extension point is this same field (an `index: "embedding"` value
   in `KNOWLEDGE_INDEX_MODES` plus the engine behind it) — the vocabulary offers no
   mode no host serves today, which is why `embedding` is not in the enum yet.
   Reindex was deliberately left to the knowledge-base side; per-source caps are
   still open._
3. **An index with freshness.** Incremental by path, mtime and size; a visible
   "n documents, last indexed …" per source; a stale marker when files changed. All
   bounded — the index stays a cache, never a second copy of the documents.
4. **Retrieval quality.** Rank with path and title weighting and recency; group hits by
   document so the model reads fewer, better chunks; de-duplicate near-identical hits;
   always return enough context for a citation.
5. **Citations that survive.** Keep opening the source file (the file-nav scheme). Add
   the quote and its page or offset, and show a **relative** path in shared output
   rather than the absolute path of the author's machine.
6. **The Ask panel grows a Sources section** — add a folder or file, enable/disable,
   reindex, see counts and size — and the mods panel shows a knowledge feature as a
   _source_ with a jump to the knowledge base, not as a kind of its own.
7. **One assistant tool surface**: search the base, read a page of a hit, and report
   coverage honestly ("searched 3 sources, 41 documents, showing 6 hits").
8. **Scope, when it is needed**: let a request restrict the search to one source, or to
   the documents related to the current file, and say which scope was used.
9. **Missing paths must be visible, not silent.** After a pack import, a source whose
   path does not exist on this machine is kept, flagged "needs attention", and never
   searched — the honest fix for the absolute-path problem.
10. **Unchanged safety**: read-only, name and page addressed, bounded counts and bytes,
    traversal-guarded, no network from a mod, and the knowledge base never writes a
    document. It informs the assistant; it never edits a file.

### 9.6 What must not happen

- No second retrieval engine inside the mod system: if the knowledge base cannot answer
  something, the fix belongs in the knowledge base.
- No embeddings or model calls from the mod host; if semantic retrieval is ever added it
  is a knowledge-base decision with its own cost story, not a side effect of a mod.
- No document writes, ever — reading material is not editing it.

### 9.7 Migration, in three steps

| phase | what changes                                                                                                                                                            | what stays compatible                                                 |
| ----- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------- |
| 1     | the knowledge base learns the source registry; a knowledge feature registers a source; the Ask panel lists sources                                                      | the record shape, the pack format, the host reader, the two old tools |
| 2     | the assistant's knowledge tools go through the base (search, page reads, citations); the two mod tools are marked deprecated in prompt text only                        | nothing breaks; both tools still answer                               |
| 3     | `list_knowledge` and `read_knowledge` leave `createModsSkill`; the host reader functions stay, because the base uses them; a knowledge feature is presented as a source | packs with a knowledge feature still import and still validate        |

All three shipped in one change: the base's corpus is assembled from the sources
(shell main, `knowledge-base.ts` `assembleCorpus`), the sources ride the file
index as extra paths so retrieval and freshness are the index's own, `kb:sources`
/ `kb:source-set-enabled` / `kb:reindex` serve the Ask panel's Sources section,
the six editors' assistants mount the knowledge-base skill over the same `kb:*`
channels, and the two mod knowledge tools were removed in the same stroke since
nothing else consumed them.

### 9.8 Acceptance (what the implementer must show)

1. A folder source is searchable, and its hits carry citations that open the right file.
2. A disabled source is not searched, and the coverage line says so.
3. A source whose path is missing warns in the UI and does not break any panel.
4. Caps hold and refuse with a readable reason (entries per folder, files per source,
   bytes per read).
5. The same question asked with the source on and off differs in a way that shows the
   source's material, and the citation resolves to that file.
6. The knowledge base still cannot write: no save, export or print path is touched, and
   the existing starred-files behaviour is unchanged.

### 9.9 Out of scope

A RAG service, a hosted index, embeddings, a document-management UI, and any write path
from the knowledge base.

## 10. Index of the changes

```
docs/mods-design.md                     the design (kinds, red lines, PR order)
docs/mods-pr1-mod-interface.md          registry, IPC, bridge, the panel
docs/mods-pr2-declarative-slots.md      slots: places, elements, actions, file-list
docs/mods-pr3-appearance.md             appearance: names, not colours
docs/mods-pr4-script-kind.md            script: read-only dry run
docs/mods-pr5-native-view.md            the native view slot (a browser)
docs/mods-pr6-overlay-slot.md           the canvas overlay slot
docs/mods-pr7-editor-slots.md           the renderer moves to @genoffice/ui; Docs
docs/mods-pr8-five-editors.md           slides, pdf, markdown, html
docs/mods-pr9-sheets.md                 sheets
docs/mods-pr10-script-apply.md          apply, with its two gates
docs/mods-pr11-script-run.md            the panel's Scripts section
docs/mods-pr12-mods-tools.md            the native assistant's read-only tools
docs/mods-pr13-six-assistants.md        the tools in all six editors
docs/mods-pr14-hook-kind.md             hooks: observe
docs/mods-pr15-hook-activity.md         hooks: activity in the panel
docs/mods-pr16-hook-veto.md             hooks: veto a save
docs/mods-pr17-knowledge-kind.md        knowledge sources
docs/mods-pr18-command-kind.md          commands and the palette
docs/mods-pr19-packs.md                 packs: export, review, install
docs/mods-kind-script.md                the parked CLI script work this revived
```
