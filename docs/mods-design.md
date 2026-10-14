# Mods — design (draft for review)

Status: **proposal for review**. Nothing described here is implemented. Interface
names and shapes are sketches; all of them are open to change.

This document covers four things:

1. the two assistants — a _mod assistant_ that works on **capability**, and a
   _native assistant_ that edits **documents** — and the rule that separates them;
2. the **features a mod can add** — a script function is one of them, not the
   system's foundation — and who writes and calls each;
3. the **three slots** through which a mod can appear in the UI;
4. the **interfaces to expose**, written out for review.

---

## 1. Scope

### 1.1 Goal

Open a customization surface so that a user — or an AI acting for the user — can
extend how GenOffice looks and behaves: add tools and commands, add a script that
performs a repetitive job, add a panel, add a visualization over a document page.

### 1.2 The red line

There is exactly **one authoritative document path**: how a document is laid out,
rendered, saved, exported and printed. Mods never join it, never replace it and
never patch it.

Adding _new_ rendering is allowed, as long as it is additive:

| Allowed                                                                              | Not allowed                                                               |
| ------------------------------------------------------------------------------------ | ------------------------------------------------------------------------- |
| A display-only overlay above the page (highlight, annotation, ruler, comment bubble) | Changing pagination, line breaking, table layout, slide element drawing   |
| A separate secondary view of the document model (thumbnail wall, mind map, preview)  | Changing what save / export / print produce                               |
| An extra exporter producing an additional artifact                                   | Injecting logic into the save / export / print path                       |
| Writing standard content into the document (picture, shape, table, chart)            | Storing content only the mod can render, leaving a hole for everyone else |

Content written into a document must be renderable by the engine on its own, or
degrade to a static placeholder. "The document is empty without this mod" is not
acceptable.

### 1.3 Acceptance test

> **With every mod disabled, a document must open, display, save, export and
> print exactly as it does today.**

This is the single test the whole design is judged by. It is mechanical: run the
existing engine tests with no mods loaded, and diff the output of a document that
had mods applied to it. Any design that cannot pass it is out of scope.

---

## 2. Current state

Facts this design builds on.

- **The capability seam already exists.** AgentSkill in
  packages/agent-core/src/skill.ts is a small contract: a system-prompt section,
  JSON-Schema tools, per-turn context, a tool executor, and a post-run
  verifyResponse correction. composeSkills() merges several skills and rejects
  duplicate tool names. The loop assembles
  system = runtimePreamble() + skill.systemPrompt + systemSuffix(), and
  tools = skill.tools (packages/agent-core/src/loop.ts).
- **The loop only notifies; it does not ask.** AgentLoopEvents (onText,
  onToolStart, onToolExecuted, onTurnEnd, onDone, onError) carries no return
  value and cannot rewrite or veto anything. There is a snapshot hook for
  one-click rollback and a post-hoc verifyResponse, but no pre-execution
  interception.
- **The AI surface is duplicated.** Six apps each own a panel, a transport and a
  skill set (apps/<app>/src/renderer/ai/); the shell owns Home, the AI settings
  and the shared ai:stream handler. Adding one capability means touching all of
  them — the knowledge-base work touched 32 files for one feature.
- **Home already has the file APIs a switcher needs**: recents, starred,
  searchFiles, listFolder, folderRoots, openPath in
  apps/shell/src/shared/home-api.ts.
- **Two external control surfaces exist**: the bundled CLI and genoffice mcp
  (stdio), plus the in-app MCP server on loopback
  (apps/shell/src/main/mcp/mcp-server.ts). They prove out-of-process control;
  they are not a mod host.
- **Renderers are sandboxed**: sandbox: true, contextIsolation: true,
  nodeIntegration: false (apps/shell/src/main/index.ts). Everything a renderer
  can do goes through the hand-written preload allowlist. Third-party code
  therefore cannot be loaded into a renderer, which is why the script host lives
  outside them (section 8).
- **A script sandbox already exists, in one place.** The slides layout script
  (apps/slides/src/renderer/ai/layout-script.ts) runs a restricted JS-shaped
  program: it receives a snapshot of the page's real geometry, computes, and
  calls a fixed vocabulary that **only collects operations**; the caller then
  dispatches the collected operations through the existing edit IPC. It is
  limited to slides, one-shot and not user-manageable — but it is the execution
  model this design generalizes.
- **Ops already are a declarative vocabulary.** docs|sheet|slides apply --ops
  accepts a list of operations, guide <domain> --json emits the per-op schema the
  validator uses, and apply already supports --dry-run. The protocol in section 4
  fronts exactly this vocabulary.

**Conclusion:** the missing pieces are (a) a pre-execution event seam, (b) one
capability registry instead of seven copies, (c) the interface that lets the mod
assistant add features, and (d) a way to load third-party code outside the
renderer.

---

## 3. The two assistants

### 3.1 Division of labour

|                          | **Native assistant** (document assistant)         | **Mod assistant** (feature assistant)               |
| ------------------------ | ------------------------------------------------- | --------------------------------------------------- |
| Job                      | Edit documents, and call the features a mod added | Add features                                        |
| Where it lives           | The AI panel of each of the six apps              | Home, plus one global shortcut                      |
| Document access          | Read **and write**                                | **Read only**                                       |
| Adds a feature           | —                                                 | **Yes**, of whatever kind fits (section 4)          |
| Writes a script function | —                                                 | **Yes**, when a script is the right kind of feature |
| Calls a script function  | **Yes**, inside its own run                       | Never                                               |
| Registers UI slots       | —                                                 | Yes                                                 |
| Modifies the document    | Yes, through the existing review + undo path      | Never                                               |

### 3.2 The rule

**The mod assistant has no document writer in its toolbox.** The restriction is
not a convention that is written down and checked: the capability profile handed
to it simply does not contain a write method. This is what makes "read-only"
reliable.

### 3.3 Hand-off

A feature the mod assistant registers becomes usable **inside the native
assistant's run**. For a script function, the native assistant calls it like any
other tool — whether the mod assistant just wrote it or a mod shipped it. Because
the call happens inside the native assistant's own run, review, snapshot-based
undo and attribution are inherited for free.

### 3.4 Indirect permission — the one real risk

The mod assistant cannot write documents, but it can add a script function, and
the native assistant calls it. A feature the mod assistant adds can therefore
reach the document indirectly.

Mitigation, in order of importance:

1. The mod API is the **only** capability surface. A script gets no raw file,
   process or network handle: everything goes through the API in section 6.4.
2. What a function can reach follows from the API surface it is handed, not from
   anything the writer declares; the host checks every call at run time, and the
   audit can print the reach without running anything.
3. The first time the native assistant runs a function that would write, the user
   sees what it would do and confirms. Confirmation can be remembered per
   function, and revoked.
4. Anything that arrives at run time is untrusted until confirmed: show what it
   was handed and what it would do, then run.

### 3.5 How a user tells them apart

The capability ranges are the distinction, and they must be visible:

- the mod assistant never offers "change this paragraph for me" — it has no
  document writer;
- the native assistant never offers "add a feature to GenOffice" — it calls the
  features that already exist;
- every feature a mod contributes is labelled with the mod that provided it.

---

## 4. What a mod can add

A mod adds **features**. There is no single mechanism: the mod layer exposes a set
of extension kinds, and a mod uses whichever kind fits the feature it is adding.
The mod assistant is the one that adds them.

A **script function is one of those kinds**. It is not the foundation of the mod
system and not a peer of it — it is simply what a mod reaches for when the feature
being added is "do this repeatable job". A mod might add only a button and never a
script; another might add a script and nothing else.

| #   | Kind                | Adds                                                                           | Called by            | When             |
| --- | ------------------- | ------------------------------------------------------------------------------ | -------------------- | ---------------- |
| K1  | Script function     | a JS/TS function the AI writes against the exposed API, to do a repeatable job | the native assistant | first kind built |
| K2  | Event hook          | observe, rewrite or veto a tool call or a prompt                               | the loop             | later            |
| K3  | Declarative slot    | a card, button or panel described as data and drawn by the host                | the host renders it  | later            |
| K4  | Canvas overlay slot | a display-only layer above the document page                                   | the host renders it  | later            |
| K5  | Native view slot    | an embedded web view (browser, terminal, visualization)                        | the host mounts it   | later            |
| K6  | Command             | a named action in the palette, ribbon or a context menu                        | the user             | later            |
| K7  | Knowledge source    | a corpus the assistant may search                                              | the native assistant | later            |

K3, K4 and K5 are the **three slots** of section 5, counted separately because
each is a different mechanism and needs its own work.

Adding one kind must not block the others: the registry holds whatever kinds
exist, and the API surface is shaped by the kinds already built rather than by a
hypothetical full set.

### 4.1 Script functions: an exposed interface, not a fixed format

The point of this kind is one sentence: **the mod layer exposes an interface to
the AI, and the AI writes functions against it.** How the AI then writes them is
the AI's business.

This document therefore deliberately does **not** fix the shape of a script: not
a step list, not a required parameter block, not a manifest the writer has to
author. Three things are fixed, and only three:

- **what a function can reach** — the API surface it is handed (section 6.4), and
  nothing else;
- **who calls it** — the native assistant, because calling a function is document
  work;
- **what the host guarantees when it runs** — the sandbox, a time budget, and a
  single undo unit.

Everything else stays open until a first implementation exists. Nailing down a
format before anyone has written a real function is the fastest way to lock in
the wrong one.

- **Who writes one.** The mod assistant, when a function is the right kind of
  feature for what the user asked for. A mod author shipping functions with a mod
  is the same thing one step earlier.
- **Who calls one.** The native assistant. The mod assistant never calls one.

Beyond the exposed API nothing constrains the author: no required style, no
required declaration. What a function can reach follows from the API it was
handed, not from anything it says about itself.

### 4.2 Execution model

1. The host builds a **snapshot** of what the script may read (open document,
   selection, requested range) and passes it in.
2. The script runs in the host sandbox with only the mod API bound — no raw
   filesystem, process or network handle.
3. Every API call is checked against the API surface the function was handed.
4. Mutating calls are collected and dispatched through the existing engine path,
   atomically, inside the native assistant's run.
5. On rejection, over-budget or timeout, nothing is applied.

Consequences: what the function was handed, and what it would do, can be seen
before it runs; a run is a single undo unit; and a script that overruns its budget
is stopped rather than allowed to wedge the editor.

### 4.3 Lifecycle

create → store → dry-run → run → undo

The mod assistant creates and dry-runs; the native assistant calls. Storage lives
under .genoffice/scripts/<id>/, or inside a mod's own directory when a function
ships with the mod.

### 4.5 Capabilities this is meant to reach

The kinds above are the mechanism. These are the concrete capabilities they are
for, named so the work can be checked against something a person would ask for:

| Capability                                       | Built from          | Notes                                                       |
| ------------------------------------------------ | ------------------- | ----------------------------------------------------------- |
| Add a browser panel                              | K5 native view      | host-mounted webview in its own partition; no editor access |
| Add a new script function                        | K1                  | the AI writes it, the native assistant calls it             |
| Add a file / navigation side panel               | K3 declarative slot | reuses the existing file APIs; the slot supplies the view   |
| Change chrome: ribbon, menus, status bar, tabs   | K3 declarative slot | themed host components, so dark mode and i18n follow        |
| Change the app background / accent (chrome only) | K8 appearance       | tokens on the shell chrome; never the document surface      |
| Draw over a document page (annotations, rulers)  | K4 canvas overlay   | screen-only layer; save, export and print never see it      |
| Add commands, workflows, knowledge sources       | K6, K7              | the pieces a mod composes into something bigger             |

Every row is subject to section 1.2: the document rendering path is untouched,
and the acceptance test in section 1.3 still decides whether it shipped.

### 4.4 Relationship to the other kinds

Nothing about a script is special to the mod layer. Registering one puts a tool in
the native assistant's toolbox the same way registering a command or a slot puts
something in the interface: same API surface, same capability check, same audit.
If a later kind needs something the script kind did not, that is a change to the
shared surface, not a second system.

---

## 5. The three slots

Slots are named places where a mod may contribute UI. They are additive: a slot
renders alongside what GenOffice draws, or replaces a marked element within it —
it never changes what the document canvas paints.

### 5.1 Declarative slot

The mod returns a **data description** of a small tree built from a fixed element
set; GenOffice renders it with its own components.

- Elements (initial set): Box, Text, Button, Input, Select, Link, Markdown, Code,
  Image, List, Progress.
- Why: theme tokens, dark mode, text scaling and i18n stay under the host's
  control, and no third-party code runs in the renderer.
- Slots it can fill: AI panel (header, composer-adjacent, message actions, tool
  card), Home (card, panel, entry), ribbon (button, group), context menu (item),
  status bar, tab bar, editor side panels.

### 5.2 Native view slot

The host mounts a **native view** — an Electron webview / BrowserView — into the
slot.

- Use cases: an embedded browser, a terminal, a third-party visualization
  library.
- The view runs in its **own session partition** (separate storage), with node
  integration and preload disabled. It cannot reach the editor or the document.
- It talks to the rest of GenOffice only through the capability API — for example
  "insert this page's text into the document" is a tool call, not DOM access.
- The declarative element set cannot express a web page; this slot exists for
  exactly that case.

### 5.3 Canvas overlay slot

A **display-only layer above the document page**.

- Use cases: AI highlights, annotations, comment bubbles, rulers, guides, custom
  markers.
- The host provides four things: page-to-screen **coordinate mapping** (scroll,
  zoom, page turn), a **redraw signal**, **hit-testing** callbacks, and the
  **theme color table**.
- It is screen-only: the save, export and print paths do not load this layer at
  all. The dark-page remap in apps/docs/src/renderer/editor/dark-page.ts is the
  precedent for that separation.
- It has a per-frame **time budget**; a mod that exceeds it is skipped for that
  frame, so an expensive overlay cannot stall the editor.

### 5.4 Where each slot can appear

| Slot           | Home | AI panel (6 apps) |             Editor chrome             |  Document canvas   |
| -------------- | :--: | :---------------: | :-----------------------------------: | :----------------: |
| Declarative    | yes  |        yes        | yes (ribbon, menus, status bar, tabs) |         no         |
| Native view    | yes  |        yes        |                  no                   |         no         |
| Canvas overlay |  no  |        no         |                  no                   | yes (display only) |

---

## 6. Interfaces to expose (draft)

### 6.1 Mod package

    .genoffice/mods/<id>/
      mod.json          # manifest
      register.js       # entry: export function register(on, options) {}
      scripts/          # optional, scripts shipped with the mod
      assets/

### 6.2 Manifest (mod.json)

    {
      "id": "acme.heading-tools",
      "name": "Heading Tools",
      "version": "1.0.0",
      "apiVersion": "0.1",
      "entry": "register.js",
      "hooks": ["tool.call", "prompt.section", "ui.render"],
      "capabilities": ["docs.read", "ui.slot", "script.register"],
      "slots": ["app.aiPanel.header", "ribbon.group"]
    }

hooks and capabilities are the two lists the validator prints without running
the mod.

### 6.3 Events

    on(event: string, matcher?: Record<string, unknown>, handler: Hook): Registration
    interface Registration { catch(handler: (error: HookError) => void): void }

    type Hook = (
      $: ModsApi,
      e: Readonly<Record<string, unknown>>,
      next: Next,
    ) => Promise<unknown> | unknown

    interface Next {
      (e?: Record<string, unknown>): Promise<unknown>   // continue the chain
      readonly signal: AbortSignal
      readonly budget: { ms: number; remainingMs: number }
      readonly origin: { mod: string; tier: "prepend" | "user" | "append" | "builtin" }
    }

A handler either passes the event on (next(e)), rewrites it
(next({ ...e, text })), or answers it by returning an object and never calling
next.

Draft event set, and where each would be implemented:

| Event                           | Fires                              | Handler may return              | Implementation site                 |
| ------------------------------- | ---------------------------------- | ------------------------------- | ----------------------------------- |
| tool.call                       | Before a tool executes             | next(e) / { deny } / { result } | AgentLoop, before skill.executeTool |
| tool.describe                   | When a tool description is sent    | { description }                 | tool assembly / composeSkills       |
| prompt.section                  | When the system prompt is composed | { text } / { text: null }       | loop.ts system assembly             |
| prompt.submit                   | A user message is submitted        | next({ ...e, text }) / { drop } | panel composer                      |
| turn.start                      | A turn begins                      | next(e)                         | AgentLoop                           |
| turn.step                       | Before one request to the model    | next({ ...e, model })           | AgentLoop.startTurn                 |
| turn.complete                   | A turn ended                       | { text } (a note line)          | AgentLoop                           |
| script.run                      | Before a script executes           | next(e) / { deny }              | script host                         |
| command.run                     | A mod command runs                 | { text } / next(e)              | command registry                    |
| ui.render                       | A slot is drawn                    | a declarative tree              | each slot host                      |
| ui.press / ui.input / ui.select | A mod-drawn control is used        | next(e)                         | slot host                           |
| session.start                   | The app session loads              | next(e)                         | shell                               |

### 6.4 The capability API

The API is handed to every hook as its first argument, and it is what a script is
written against. It is the **only** way a mod or a script affects anything outside
its own code — which is why the static audit is possible.

    interface ModsApi {
      readonly mod: { id: string; root: string; version: string }
      readonly docs: DocsApi        // mod profile: read-only
      readonly sheets: SheetsApi    // mod profile: read-only
      readonly slides: SlidesApi    // mod profile: read-only
      readonly pdf: PdfApi          // mod profile: read-only
      readonly files: FilesApi      // read, list, open; no delete/move
      readonly script: ScriptApi    // register, list; no run in the mod profile
      readonly ui: UiApi            // slot registration, open pane, toast, ask, pick
      readonly command: CommandApi  // register, run, list
      readonly store: StoreApi      // per-mod key/value
      readonly state: StateApi      // reactive state
      readonly clock: ClockApi      // now, sleep, after, every
      readonly model?: ModelApi     // gated capability
      readonly http?: HttpApi       // gated capability
      readonly process?: ProcessApi // gated capability
    }

**Two profiles, one interface.** The native assistant receives the write-capable
profile; mod callers receive the read-only profile on the document namespaces.
The write methods exist in neither the type nor the runtime object handed to a
mod. http, process and any document writer are absent unless the user granted
them.

Read methods on the document namespaces are deliberately the same ones the tools
already use today (read structure, read range, read text), so nothing new has to
be invented for the common case.

### 6.5 Script interface (shape deliberately open)

The essence is an API handed to the AI. The module shape below is one possible
form and is expected to change once a first implementation exists; nothing in it
is a contract the author has to satisfy.

    // The AI writes a function against the API it is handed:
    //   export default async function ($, input) { ... }

    interface ScriptApi {
      register(input: { name: string; source: string }): Promise<{ id: string }>
      list(): Promise<ScriptInfo[]>
      dryRun(id: string, input?: Record<string, unknown>): Promise<{ plan: unknown }>
      // run/apply is deliberately absent from the mod profile:
      // calling a script is the native assistant's job (section 4.1).
    }

    // Illustrative only. The id and name come from the registration call, and what
    // the script may reach comes from the API it is handed, not from a declaration.
    interface ScriptInfo {
      id: string
      name: string
      createdAt: string
    }

Registering a script makes it a callable tool in the native assistant's toolbox.
The mod assistant may write a new script; the native assistant calls it, whether
it was just written or shipped by a mod.

### 6.6 Audit

    genoffice mod validate <dir>   # prints hooks: and calls: without running anything
    genoffice mod list             # installed mods, grants, and what each can reach

Because every external effect flows through the capability API, both listings can
be generated from the API surface a feature is handed and the static shape of the
code, not from a hand-written declaration.

---

## 7. Permission model

Two independent axes, because read-only is not automatically safe:

|               | Cannot reach the network                | Can reach the network                                    |
| ------------- | --------------------------------------- | -------------------------------------------------------- |
| **Read-only** | Default. Granted on install.            | Requires consent: readable documents can be exfiltrated. |
| **Can write** | Requires consent, plus review and undo. | Denied by default.                                       |

Tiers:

- **T0 — default**: read documents, fill slots, register commands and scripts,
  register a capability surface.
- **T1 — granted**: a script the native assistant runs may write, but only into
  the current document or selection, always reviewable, always a single undo
  unit.
- **T2 — denied unless explicitly granted**: delete or move files, spawn
  processes, network access.

Plus: a per-mod grant revocable at any time, a global safe mode that disables
every mod, and an in-place affordance — right-clicking a mod's contribution shows
which mod it came from and offers to disable it without a trip to Home.

---

## 8. Where things run

| Piece                                                            | Location                             | Why                                                                 |
| ---------------------------------------------------------------- | ------------------------------------ | ------------------------------------------------------------------- |
| Install, trust, grants, enable/disable, audit, the mod assistant | **Home**                             | One control centre; available with no document open                 |
| Capability definitions, the protocol and the mod registry        | **Shared middle layer**              | Every surface reads one source; Home cannot supply document context |
| Mod code and the script host                                     | **Main or a utility process**        | Renderers are sandboxed and cannot load third-party code            |
| Slots and overlays                                               | **Each window**                      | They render where the user is looking                               |
| Creating a script function                                       | **Mod assistant**                    | A script is one kind of feature the mod layer adds                  |
| Calling a script function                                        | **Native assistant, inside its run** | One writer; existing review and undo paths                          |

One shortcut (mod.palette.open, user-remappable) opens a single palette listing
every entry point contributed by installed mods. Mods do not each claim a key.

---

## 9. Delivery plan (PR sequence)

### 9.1 How to slice

Do not ship a "mod platform" PR. It would be too large to review, impossible to
revert in one piece, and it would carry five unproven ideas at once.

Slice so that every PR:

- is independently mergeable and independently revertible;
- has something to run or something to look at when it lands;
- leaves the document path untouched, so section 1.3 still passes;
- does not depend on a later PR.

Order by **risk**, not by layer. The one genuinely unproven assumption in this
whole design is whether a JS/TS script behind a restricted protocol can produce
safe, auditable, undoable document edits. Everything after that — assistants,
slots, distribution — is conventional product plumbing around a core that is
already known to work.

Start with **one domain (docs) and one host (the CLI)**, the way the
knowledge-base work made docs the first panel consumer. The other five apps, the
overlay and the native view each come later as their own PR.

### 9.2 The sequence

| #                                                                                 | PR                                                                                                                         | Carries                                                                                                                                                         | Done when                                                                                                                                                                               |
| --------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 1                                                                                 | The interface to the mod assistant: a feature registry and the tools it is handed, with no kind executing                  | the registry (one JSON record per feature), the IPC boundary, the preload bridge, the AgentSkill, a read-only document reader, and the Home panel that calls it | from Home, asking for a feature records it; it can be listed, read, updated and deleted and survives a restart; the assistant never claims a feature runs, and never changes a document |
| 2                                                                                 | K1 — a script function: read-only, dry-run, CLI host                                                                       | the API surface for one domain, the sandbox host, budgets, script run --dry-run                                                                                 | a script reads a document, computes and prints a plan; a call outside the API it was handed is refused; an over-budget script is stopped                                                |
| 3                                                                                 | K1 — script functions that write                                                                                           | mutating calls collected and dispatched through the existing ops path, first-write confirmation, one undo unit                                                  | a script output matches the equivalent apply --ops exactly; the whole run is a single undo                                                                                              |
| 4                                                                                 | K1 inside the app: the native assistant calls scripts                                                                      | utility-process host, registration as a tool, the consent step, docs first                                                                                      | in docs, ask for a repeatable job; the assistant calls a script and the change is reviewable and undoable                                                                               |
| 5                                                                                 | K2 — event hooks                                                                                                           | tool.call, tool.describe, prompt.section, turn.step, turn.complete, with next, budgets and abort                                                                | hooks can observe, rewrite and veto; with no subscriber the loop behaves exactly as it does today                                                                                       |
| 6                                                                                 | K3 — the declarative slot (first of the three slots; K4 the canvas overlay and K5 the native view follow as their own PRs) | the element set, the slot names, and host renderers for the AI panel header, Home card, ribbon button and context menu item                                     | a mod adds a button and a card; theme, dark mode, scaling and i18n follow the host                                                                                                      |
| 7                                                                                 | Packaging and trust                                                                                                        | mod.json, the loader in the host process, grants, mod validate / mod list, safe mode                                                                            | a mod installs from a directory, can be disabled, and its hooks and calls are listed without running it                                                                                 |
| After that, each of these is its own PR: the canvas overlay slot, the native view |
| slot, fan-out of the assistants to the other five apps, and                       |
| marketplace/distribution.                                                         |

### 9.3 Why this order

- **Risk first.** PR 1 and PR 2 settle the only real unknown. Everything later is
  plumbing around a core that has already been shown to work.
- **Value early.** After PR 2 a user can automate a document with a script; after
  PR 3 the assistant can write one. Neither requires mods.
- **The first two PRs do not touch the UI.** No i18n (21 locales), no theme
  tokens, no Electron. They are the easiest PRs in this repository to review.
- **Hooks deliberately come late.** The script kind does not need the event
  seam; only a mod that intercepts the assistant does. Shipping the seam later
  keeps PR 1 to PR 4 small.
- **Distribution last.** A package format, a trust UI and an audit are only
  meaningful once there is something worth distributing.

### 9.4 Out of scope for every PR in the sequence

- no change to document rendering, save, export or print — section 1.3 is the
  gate;
- no third-party code loaded into a renderer (section 8);
- no compatibility promise: the protocol ships marked experimental, with
  apiVersion reserved for the day it stabilizes;
- no marketplace, no signing, no remote install.

### 9.5 If only one PR can be done

PR 1. It is the smallest thing that proves the model end to end — the first
feature kind a mod can add, with its API, its budget and its host —
and because it writes nothing, there is no security argument to have while the
shape is still moving.

---

## 10. Open questions

1. **Naming.** "Mod" is also the name of the Claude Code feature this is modelled
   on. Do we keep it, or use "extension" / "add-on"? And is "mod assistant" the
   right name for the capability-side assistant?
2. **Shortcut.** Which key opens the palette, and do we need per-app overrides?
3. **Context for the mod assistant.** It lives in Home, but it must read
   documents. Does it read the currently open document (needs cross-window
   context) or only files the user picks or has open? The former is more useful,
   the latter is simpler.
4. **Scope of a mod.** Per user, per workspace, or per document?
5. **Distribution.** Do we ship a marketplace now, or only local install plus the
   existing Settings → Integrations surface?
6. **API stability.** At 0.1.0 the codebase moves fast. Do mods get an
   apiVersion with a compatibility promise from day one, or is the surface marked
   experimental until it settles?
7. **Obligations on mod UI.** CI enforces semantic theme tokens, 21 locales and
   English-only comments. Third-party mods cannot be held to that, so the host
   must supply themed components and i18n injection — confirm this is a hard
   requirement rather than a documented expectation.
8. **Scripts shipped by a mod vs written by the mod assistant.** Same format,
   different trust: one is reviewed at install, the other is new code written on
   the spot. Confirm that a function written at run time always requires a visible
   confirmation, with its code and what it would do, before its first write.
9. **Script sandbox profile.** JS/TS is the assumed language for the exposed API.
   Confirm which globals the sandbox exposes, whether a function may import
   packages, and what the time and memory budgets are.
10. **Native view slot and the browser.** Is an embedded browser a mod feature or
    a built-in GenOffice tab that mods may place and configure? The latter is
    cheaper and more consistent.
