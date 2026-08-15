# Plan 025: Roll the shared Codex controls into PDF and Markdown

> **Implementation instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update this plan's status row in
> `plans/README.md` after implementation and review.
>
> **Drift check (run first)**:
> `git diff --stat cf1f3c1..HEAD -- apps/pdf apps/markdown packages/ui packages/agent-core plans`
> Plan 023 is expected drift. If PDF/Markdown excerpts no longer match, stop
> and refresh this plan before implementation.

## Status

- **Priority**: P1
- **Effort**: M
- **Risk**: MED
- **Depends on**: Plan 023
- **Category**: direction
- **Planned at**: commit `cf1f3c1`, 2026-08-16

## Why this matters

PDF and Markdown already send tool-capable agent turns through the suite AI
handler, but their panels fetch settings only when sending and expose no
settings mutation, auth, capability, or settings-event APIs. They form the
second rollout shape: settings should become mounted panel state rather than
top-level App state. This plan adds the approved Codex controls while leaving
PDF and Markdown editing skills, history, and document mutations untouched.

## Current state

- `apps/pdf/src/renderer/ai/AiPanel.tsx:85` holds only a nullable settings ref;
  lines 219-222 fetch settings immediately before each run.
- `apps/markdown/src/renderer/ai/AiPanel.tsx:113` has the same ref-only shape;
  lines 342-345 fetch settings immediately before each run.
- Both AgentLoops are constructed once with ref getters
  (`apps/pdf/src/renderer/ai/AiPanel.tsx:101-136` and
  `apps/markdown/src/renderer/ai/AiPanel.tsx:160-171`), so updating the ref
  changes the next send without recreating or canceling the loop.
- `apps/pdf/src/shared/ipc.ts:587-595` and
  `apps/markdown/src/shared/ipc.ts:46-53` define pass-through AI channels from
  the shell. Neither includes set-settings, settings-changed, or Codex methods.
- `apps/pdf/src/preload/index.ts:71-77` and
  `apps/markdown/src/preload/index.ts:57-63` expose get/stream/cancel only.
- `apps/pdf/src/renderer/ai/transport.ts` and
  `apps/markdown/src/renderer/ai/transport.ts` do not resolve safe Codex error
  codes to localized text.
- `apps/pdf/src/renderer/i18n/strings.ts` and
  `apps/markdown/src/renderer/i18n/strings.ts` are full 19-locale dictionaries.
  Docs `apps/docs/src/renderer/i18n/strings-ai.ts` is the approved source for
  Codex labels/errors; copy those translations exactly.
- PDF and Markdown renderers are hosted by the shell in production, where
  `apps/shell/src/main/index.ts:2816` already registers the Codex-aware handler.
  Their standalone main processes are intentionally deferred to Plan 026.

## Commands you will need

| Purpose            | Command                                    | Expected on success |
| ------------------ | ------------------------------------------ | ------------------- |
| PDF tests          | `npm run test -w @genoffice/pdf`           | all pass            |
| PDF typecheck      | `npm run typecheck -w @genoffice/pdf`      | exit 0              |
| Markdown tests     | `npm run test -w @genoffice/markdown`      | all pass            |
| Markdown typecheck | `npm run typecheck -w @genoffice/markdown` | exit 0              |
| Theme gate         | `npm run check:theme-colors`               | exit 0              |

## Scope

**In scope**:

- `apps/pdf/src/shared/ipc.ts`
- `apps/pdf/src/preload/index.ts`
- `apps/pdf/src/renderer/main.tsx`
- `apps/pdf/src/renderer/ai/AiPanel.tsx`
- `apps/pdf/src/renderer/ai/transport.ts`
- `apps/pdf/src/renderer/i18n/strings.ts`
- `apps/pdf/src/renderer/styles.css`
- One focused PDF provider-control test under `apps/pdf/tests/`
- `apps/markdown/src/shared/ipc.ts`
- `apps/markdown/src/preload/index.ts`
- `apps/markdown/src/renderer/main.tsx`
- `apps/markdown/src/renderer/ai/AiPanel.tsx`
- `apps/markdown/src/renderer/ai/transport.ts`
- `apps/markdown/src/renderer/i18n/strings.ts`
- `apps/markdown/src/renderer/styles.css`
- One focused Markdown provider-control test under `apps/markdown/tests/`
- `plans/README.md`

**Out of scope**:

- PDF/Markdown main-process AI handler registration; Plan 026 owns it.
- PDF or Markdown editor tools, prompts, snapshots, history persistence,
  exports, image generation/search, or file handling.
- Top-level App state changes unless a live code constraint makes panel-owned
  settings impossible; if so, STOP rather than widening scope.
- Per-app provider settings or an app-specific copy of shared controller code.
- Provider transport, OAuth, credentials, or shared-control redesign.

## Git workflow

- Work on the current branch; do not push or open a PR.
- Keep PDF and Markdown changes separable if commits are requested later.
- Suggested subjects: `add Codex controls to PDF` and
  `add Codex controls to Markdown`.

## Steps

### Step 1: Extend PDF and Markdown preload contracts

Add typed channel constants and preload methods for `ai:set-settings`,
`ai:settings-changed`, the five Codex account/capability operations, and the
existing get/stream methods. Import safe account/capability/error types from
`@genoffice/ai-provider` into the shared interfaces.

Each preload must expose:

- `setAiSettings(settings)`;
- `onAiSettingsChanged(handler)` returning unsubscribe;
- `aiCodexStatus()`;
- `aiCodexLogin()`;
- `aiCodexCancelLogin()`;
- `aiCodexLogout()`; and
- `aiCodexCapabilities()`.

Do not expose credentials, authorization URLs, or raw provider failures.

**Verify**:

```sh
npm run typecheck -w @genoffice/pdf
npm run typecheck -w @genoffice/markdown
```

Expected: both exit 0.

### Step 2: Make PDF settings mounted, evented panel state

Replace the ref-only PDF settings flow with `AiSettings | null` state plus a
ref assigned every render. Load settings when the panel mounts, subscribe to
external settings changes, and unsubscribe on unmount. `updateAiSettings`
updates state/ref immediately and persists through preload; an external event
updates state/ref only and must not write back.

Do not recreate `AgentLoop`. Its existing ref getter must observe the current
settings at the next stream start. While initial settings are loading, keep the
composer visible but disable Send; do not allow a non-null assertion to run
with no settings.

Adapt `window.pdfApi` to the shared control API and render the approved three
surfaces in the existing panel header/body/composer locations. Merge existing
panel notices into the shared banner input only where the current PDF panel has
one; do not invent new notices. Use the controller's `sendDisabled` in addition
to loading/busy state.

Import shared control CSS once, use Plan 023's error resolver in transport, and
copy all approved Codex label/error translations into the PDF dictionaries.

**Verify**:

```sh
npm run typecheck -w @genoffice/pdf
npm run test -w @genoffice/pdf
```

Expected: exit 0; existing PDF AI tool tests remain green.

### Step 3: Make Markdown settings mounted, evented panel state

Apply the same mounted-state/ref pattern in Markdown. Preserve its chat
resolution/persistence, snapshots, rollback, search skill, and file-path refs.
Initial loading and signed-out Codex must disable Send without disabling draft
editing or Stop.

Adapt `window.markdownApi`, render the shared controls at the same semantic
positions, import shared CSS, use the shared error resolver, and add the
approved translations to every Markdown locale.

**Verify**:

```sh
npm run typecheck -w @genoffice/markdown
npm run test -w @genoffice/markdown
```

Expected: exit 0; existing Markdown AI tool tests remain green.

### Step 4: Add thin panel integration tests

Add one jsdom test per app, following Docs' mount/mock pattern without copying
its complete controller matrix. Each test must prove:

- provider order and selection callback;
- pending/signed-out Codex disables Send;
- capabilities make the model control visible;
- external settings update the mounted control immediately with no write-back;
- the next mocked stream receives the new settings snapshot; and
- existing app-specific tool dependencies are not invoked merely by opening
  provider controls.

**Verify**:

```sh
npm run test -w @genoffice/pdf
npm run test -w @genoffice/markdown
```

Expected: all tests pass.

### Step 5: Run the two-app gate

```sh
npm run typecheck -w @genoffice/ui
npm run typecheck -w @genoffice/pdf
npm run typecheck -w @genoffice/markdown
npm run test -w @genoffice/pdf
npm run test -w @genoffice/markdown
npm run check:theme-colors
git diff --check
```

Expected: every command exits 0.

## Test plan

- One PDF sidebar integration test for preload adapter, auth gating,
  capabilities, external settings, and next-send snapshot.
- One Markdown equivalent.
- Existing app tool tests remain characterization coverage for app-owned
  behavior.
- Shared controller edge cases remain covered once in Plan 023.

## Done criteria

- [ ] PDF and Markdown display the approved shared Codex controls in shell mode.
- [ ] Both panels load settings before send and react immediately to suite
      settings events without write-back loops.
- [ ] The next send uses the latest global settings; active sends are not
      canceled or mutated.
- [ ] Pending/loading/signed-out states cannot dereference null settings or send
      unauthorized Codex requests.
- [ ] All 19 locales include the approved Codex labels/errors.
- [ ] Focused tests, typechecks, theme gate, and `git diff --check` pass.
- [ ] No files outside Scope are modified.
- [ ] `plans/README.md` marks Plan 025 DONE after review.

## STOP conditions

Stop and report back if:

- Either panel cannot hold settings state without moving unrelated editor state
  into its top-level App.
- The AgentLoop captures settings instead of reading its existing ref at stream
  start.
- The shared controls need PDF/Markdown-specific forks.
- Any credential/token data would cross preload.
- Adding the controls requires changing PDF/Markdown skills, mutations, or
  persistence contracts.
- A focused verification fails twice after a reasonable correction.

## Maintenance notes

- Keep the state and ref assigned together; the state renders controls and the
  ref supplies the next AgentLoop request.
- Provider settings are suite state, not chat history. Do not persist them in a
  document project or localStorage.
- Standalone failures are expected until Plan 026 and are not a reason to add
  app-local provider logic here.
