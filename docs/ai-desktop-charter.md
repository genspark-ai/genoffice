# AI desktop charter — the agent workbench

Positioning decision (2026-10-29): **GenOffice is an agent workbench for local
office work, not an OS desktop replacement.** The in-app desktop is the agent's
work surface; tray, global hotkey, deep links, CLI and MCP are its nervous
system.

The one-sentence test for every feature below: does it help the user open the
machine in the morning and finish the day's office work inside this one window,
by delegating to the agent?

## Principles

1. **Every desktop capability is also an agent tool.** If the agent cannot see
   or change a piece of desktop state, that piece is not part of the product.
   The CLI → MCP → schema single-source derivation
   (`packages/cli/src/mcp/server.ts:102`) makes this cheap: add a command
   domain, get script, agent and remote channels for free.
2. **The engine is the moat; the desktop is the showroom.** Byte-preserving
   OOXML editing, the local FTS5+CJK index and reviewable AI edits are what
   Atlas/Comet/Copilot/Puter do not do. Desktop chrome only displays that
   advantage.
3. **Do not fight the OS.** No system-level window management, no global
   context-menu injection, no drawing into the real desktop layer. Electron's
   sandbox + WebContentsView architecture cannot get there; system-side reach
   goes through the CLI/MCP so Raycast, Shortcuts and coding agents can drive
   GenOffice instead.
4. **The fork stays a thin overlay.** Desktop features live in shell's home
   surface; zero engine-layer changes; upstream is tracked continuously (see
   Survival preconditions).
5. **One gate, one command.** Every roadmap item ships with a copy-pasteable
   acceptance check.

## Verified current state

All rows confirmed by reading the code on 2026-10-29.

| fact                                                                                                                                    | evidence                                                                                                                                         |
| --------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------ |
| Desktop is a view mode of Home, mutually exclusive with recent/starred/ask/cloud/folder                                                 | `apps/shell/src/renderer/src/Home.tsx:1511`                                                                                                      |
| Icon set is closed: 10 hardcoded app tiles + folder roots + last 8 recent files; no arbitrary files, no widget container                | `apps/shell/src/renderer/src/DesktopView.tsx:387`                                                                                                |
| No `globalShortcut`, no `setLoginItemSettings`, no `setAsDefaultProtocolClient` anywhere in `apps/` or `packages/`                      | repo-wide grep, zero hits                                                                                                                        |
| No MCP tool or control-socket command reads or writes desktop state (wallpaper, icons, todos, open windows)                             | `apps/shell/src/main/mcp/mcp-server.ts:21`, `apps/shell/src/main/control-server.ts:15`                                                           |
| Ctrl+K palette sources = static actions + recents + skills + plugins; it never queries the file index and has no AI fallback            | `apps/shell/src/renderer/src/CommandPalette.tsx:101`                                                                                             |
| Ask AI corpus is locked to the starred set, though the index already holds everything                                                   | `apps/shell/src/main/knowledge-base.ts:32`                                                                                                       |
| Local index: FTS5 contentless + versioned CJK bigram tokenizer, 22 extensions, 200k file cap, single-threaded extractor worker          | `apps/shell/src/main/file-index/store.ts:94`, `tokenize.ts`, `folder-tree.ts:30`, `scan.ts:17`                                                   |
| agent-core has no confirm/approval/permission path; a tool call executes immediately; `run()` drops new instructions while busy         | `packages/agent-core/src/loop.ts:494`, `:613`                                                                                                    |
| pi replaced the protocol-stream layer only; AgentLoop still owns tool execution and the ReAct loop in the renderer                      | `packages/agent-pi-bridge/src/index.ts:207` (drop-in `streamForProvider` replacement), `docs/pi-migration-charter.md:15` (zero renderer changes) |
| Home.tsx is 3,906 lines; shell main/index.ts is 7,164 lines                                                                             | `wc -l`                                                                                                                                          |
| Wallpapers + templates ship only in the Windows package, but packaged mode reads them from `resources/` on every platform               | `apps/shell/electron-builder.cjs:645` vs `apps/shell/src/main/desktop-ipc.ts:223`, `templates-ipc.ts:28`                                         |
| Fork has no `upstream` remote and no release workflow (`.github/workflows/` = ci.yml only); update feed is `GENOFFICE_UPDATE_URL`-gated | `git remote -v`; `apps/shell/electron-builder.cjs:807`                                                                                           |

## Red team — four challenges, four counters

| challenge                                                                                                                          | counter                                                                                                                                                                                                                                  |
| ---------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| "Why not macOS + a ChatGPT client?" The workbench has no system presence at all: a 200 MB window must be opened first.             | Add the smallest system footprint: tray + global hotkey + login item + `genoffice://` deep link. Days of work, and the workbench story becomes real.                                                                                     |
| "Electron can never be a desktop layer." True by architecture.                                                                     | Concede the boundary. System-side reach is outsourced to CLI/MCP; the in-app desktop is repositioned as the agent's work surface, not a Finder replacement.                                                                              |
| "A solo fork vs a 1.17M-line upstream is a maintenance cliff; your differentiator is a thin UI layer upstream can copy in a week." | Track upstream with CI trial-merges, keep fork changes confined to shell's home surface, and stake the moat on what upstream cannot copy: the local index + byte-preserving OOXML + reviewable-edit combination, shipped as agent tools. |
| "Platform vendors absorb you (Windows ODR, File Explorer MCP connector, macOS Spotlight actions)."                                 | Occupy the space they do not: agents that natively understand OOXML and rewrite local documents byte-preservingly; local-first, open-source agent governance (the open ODR).                                                             |

## Competitive whitespace (2025–2026 landscape)

- ChatGPT Atlas: browser as assistant; agent acts on logged-in web pages;
  memories live in the cloud.
  <https://openai.com/index/introducing-chatgpt-atlas>
- Perplexity Comet: the browsing session collapsed into one conversation.
  <https://www.perplexity.ai/hub/blog/introducing-comet>
- Windows agentic direction: On-device Agent Registry (ODR) registers and
  isolates MCP servers; File Explorer gets an MCP connector; Agent Launchers.
  Closed, Windows-only. <https://learn.microsoft.com/en-us/windows/ai/mcp/overview>
- Raycast: launcher as agent console; closed, macOS-only. <https://www.raycast.com/ai>
- Puter: icon desktop in the browser; files in the cloud, AI not core.
  <https://github.com/HeyPuter/puter>

Unclaimed ground: structured office documents as an agent work surface (browser
agents only understand DOM; Microsoft's connector does generic file I/O; Office
agents are locked to the 365 cloud), local-first AI workbench, open-source
agent registration/permission/audit. GenOffice already holds all three
ingredients; the work below is wiring them together.

## Survival preconditions (before any feature work)

**P-1. Packaging regression.** Move `build/templates` and `build/wallpapers`
from the `win` block to top-level `extraResources` in
`apps/shell/electron-builder.cjs:635`; today mac/linux installers ship the
desktop features empty.
Gate: `npm run dist:mac && find apps/shell/dist -name wallpapers -type d`
prints a path.

**P-2. Upstream tracking.** `git remote add upstream
https://github.com/genspark-ai/genoffice.git`, plus a daily CI job that
trial-merges `upstream/main` and fails on conflict.
Gate: `git fetch upstream` succeeds; CI job is green or files a conflict
issue automatically.

## Workstreams

### A. Wire the desktop to the agent (P0 — the vision divider)

- New CLI command domain `desktop.*`: `get-state`, `set-wallpaper`, `arrange`,
  `todo add|list`, `windows list`. Registered once in the CLI registry, they
  automatically surface as MCP tools and agent-callable verbs. Reuse the tool
  registration mechanism at `apps/shell/src/main/mcp/mcp-server.ts:21`.
- Ctrl+K becomes the intent bus: `buildRows`
  (`apps/shell/src/renderer/src/CommandPalette.tsx:245`) becomes a provider
  registry — file-index hits, open tabs, settings entries, skills, MCP tools
  all register as sources; on no match, the query falls through to Ask AI as
  natural language.
- Ask AI corpus generalizes from "starred" to "configurable scope" (default:
  all indexed folder roots; starred becomes one scope). Swap the starred `Set`
  for a scope predicate in `knowledge-base.ts` and relax the `kb:read` check
  accordingly.

Gates:

- `genoffice desktop get-state --json` prints wallpaper, icon layout and todos.
- Type a string with no command match in Ctrl+K → an AI answer streams.
- Ask AI answers a question about an indexed but unstarred file and cites it.

### B. System presence (P0)

Tray icon with Ask AI / palette quick actions; a global hotkey
(`globalShortcut`) that summons the palette without focusing the window;
`app.setLoginItemSettings` behind a settings toggle; register the `genoffice://`
protocol so browsers, CLI and automations can invoke workbench actions.

Gates:

- With the window closed to tray, pressing the hotkey opens the palette.
- `open "genoffice://ask?q=hello"` launches/focuses the app into Ask AI.

### C. Desktop: launcher → work surface (P1)

- Icon sources become providers (apps / recents / folders / widgets /
  plugins); arbitrary files and folders can be dragged onto the desktop
  (alias or a dedicated Desktop directory); `~/Desktop` mapping.
- Widgets via the existing tear-off mechanism: a "pin to desktop" mode for
  detached tabs (borderless, low level, reattach without state loss —
  `apps/shell/src/main/detached-windows.ts:37`) replaces the hand-written
  TodoWidget.
- Ask AI can reference the items currently visible on the desktop.

Gates:

- Drag a `.docx` from Finder onto the desktop → icon appears, double-click
  opens it in a tab.
- Pin a sheets tab to the desktop, quit, relaunch → widget restores with its
  document state.

### D. Agent safety base and long tasks (P1 — prerequisite for cross-app action)

- Confirmation hook: `AgentToolDef` gains a risk annotation and
  `AgentLoopOptions` a host-injected `confirmTool(call): Promise<boolean>`,
  enforced before execution in `packages/agent-core/src/loop.ts`. Tool
  execution stays in agent-core (pi only streams single turns), so this is a
  single-point change.
- `run()` returns a task handle (pause/resume/cancel/subscribe); busy loop
  queues instead of dropping (`loop.ts:613`); transcripts and run state
  persist in project-store for restart recovery and a background-task list UI.
- Shell-level desktop loop: `composeSkills` aggregates per-app skills proxied
  over IPC (tool names prefixed per app); read-only skills first, write
  skills behind the confirmation hook.
- pi parity and legacy retirement: close the `onActivity` gap in
  `packages/agent-pi-bridge/src/index.ts:226` (fine-grained UI activity during
  long tool-argument streams), then work the unmet retirement criteria in
  `docs/pi-migration-charter.md:122` — real-key tests for anthropic/gemini,
  Genspark credential injection on the pi branch, a decision on the frozen
  codex provider — and delete the legacy protocol streams.

Gates:

- A high-risk tool call (delete, overwrite) opens a confirmation UI; denying
  it aborts the call; unit tests cover both paths in agent-core.
- Ask the shell agent to summarize two open documents from different apps →
  one answer citing both.
- `GENOFFICE_AI_BACKEND=legacy` no longer appears in the tree.

### E. Engineering foundations (P1–P2)

- Split `Home.tsx` (3,906 lines, ~60 `useState`) into a reducer state machine
  - per-surface files; split shell `main/index.ts` (7,164 lines) into
    lifecycle/window/menu/home-ipc modules.
- Unify the four preference stores (desktop-layout, desktop-prefs, todos,
  localStorage `home.view`) into one versioned settings store with schema
  migration.
- Wallpaper bytes move from in-memory data URLs to a `wallpaper://` protocol.
- Fork release pipeline: public release workflow (build → sign →
  `latest*.yml` → static hosting), reusing the forward-only guards in
  `scripts/update-feed-utils.cjs`; without it fork users never receive
  updates.
- AppDescriptor registry: converge `createView/isDirty/requestClose/
agentControl` into a registered descriptor so app-routing, TabManager and
  the MCP bridge all look up one table (removes the static imports at
  `apps/shell/src/main/index.ts:177`).

Gates:

- `npm run typecheck` and `npm test` stay green after each split.
- A fork GitHub Release produces installable artifacts and an update feed.

## Roadmap

| horizon   | ships                                                 | acceptance  |
| --------- | ----------------------------------------------------- | ----------- |
| this week | P-1 packaging fix, P-2 upstream tracking              | gates above |
| 2 weeks   | workstream B (system presence)                        | gates above |
| 1 month   | workstream A (desktop ⇆ agent wiring)                 | gates above |
| 2 months  | workstream C + D confirmation hook                    | gates above |
| 3 months  | workstream E + release pipeline; legacy stack retired | gates above |

## Non-goals

- Replacing Finder/Explorer or the OS desktop; system-level window
  management; global context-menu injection.
- Executing third-party code inside plugins (the plugin format stays
  prompt-only by design).
- A vector/embedding retrieval layer — deferred until workstream A lands and
  lexical recall is measured to be the bottleneck.
- Cloud sync of the desktop layout or index. Local-first is the positioning.
