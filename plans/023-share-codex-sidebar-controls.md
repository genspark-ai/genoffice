# Plan 023: Share the settled Codex sidebar controls and synchronize AI settings

> **Implementation instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update this plan's status row in
> `plans/README.md` after implementation and review.
>
> **Drift check (run first)**:
> `git diff --stat cf1f3c1..HEAD -- packages/ui packages/agent-core apps/docs plans`
> If any in-scope file changed since this plan was written, compare the
> "Current state" excerpts against the live code before proceeding; on a
> mismatch, treat it as a STOP condition.

## Status

- **Priority**: P1
- **Effort**: M
- **Risk**: MED
- **Depends on**: Plans 001-021 (DONE)
- **Category**: tech-debt
- **Planned at**: commit `cf1f3c1`, 2026-08-16

## Why this matters

Docs contains the approved Codex sidebar UX, but its provider/auth/capability
state and three UI surfaces currently live inside a 2,000-line app-specific
panel. Copying that logic into four more panels would create five independent
implementations of auth races, capability normalization, keyboard behavior,
and accessibility. This plan extracts only the proven Codex sidebar chrome,
keeps each app's agent and tools app-owned, and makes the existing suite-wide
settings file update every already-open tab before later plans add consumers.

## Current state

- `apps/docs/src/renderer/ai/AiPanel.tsx:346-553` owns Codex account state,
  selection/request race guards, capability loading, and normalization.
- `apps/docs/src/renderer/ai/AiPanel.tsx:1038-1142` mutates provider, model,
  reasoning, and service-tier settings.
- `apps/docs/src/renderer/ai/AiPanel.tsx:1371-1448` renders the provider select
  and signed-out/error banner; lines 1688-1884 render the composer model menu
  and use `sendDisabled` for the auth preflight.
- `apps/docs/src/renderer/ai/transport.ts:5-30` maps safe IPC error codes to
  localized text. The other app transports do not have this mapping.
- `apps/docs/src/main/docs-main.ts:2484-2515` sanitizes the one
  `ai-settings.json`; lines 2601-2603 save it but do not notify other renderers.
- `apps/shell/src/main/index.ts:2816` registers this Docs-owned AI handler once
  for every app tab in the packaged suite.
- `packages/ui/src/AiComposer.tsx` is the established shared-renderer pattern:
  behavior is shared, labels and app-specific slots are supplied by callers,
  and CSS uses stable shared class names.
- `apps/docs/tests/ai-panel-collapse.test.ts:97-508` is the characterization
  suite for the approved provider order, auth preflight, stale async guards,
  model/reasoning/speed selection, reset, and collapse behavior. Preserve these
  assertions while moving their implementation.
- `CLAUDE.md` requires semantic theme tokens for renderer chrome and all new
  tokens in light, dark, and system-dark blocks. Codex controls are chrome, not
  document content.

The approved product contract is:

1. Provider, Codex model, reasoning effort, and service tier are suite-global.
2. A change from any sidebar is broadcast to every already-open renderer.
3. The new settings affect the next send; an in-flight request keeps the
   settings snapshot with which it started.
4. Credentials never cross the main/preload boundary; only redacted account
   status and validated capabilities do.

## Commands you will need

| Purpose               | Command                                 | Expected on success |
| --------------------- | --------------------------------------- | ------------------- |
| UI typecheck          | `npm run typecheck -w @genoffice/ui`    | exit 0              |
| Agent transport tests | `npm run test -w @genoffice/agent-core` | all pass            |
| Docs tests            | `npm run test -w @genoffice/docs`       | all pass            |
| Docs typecheck        | `npm run typecheck -w @genoffice/docs`  | exit 0              |
| Theme gate            | `npm run check:theme-colors`            | exit 0              |

## Scope

**In scope**:

- `packages/ui/package.json`
- `packages/ui/src/index.ts`
- `packages/ui/src/AiProviderControls.tsx` (create)
- `packages/ui/src/ai-provider-controls.css` (create)
- `packages/agent-core/src/electron-transport.ts`
- `packages/agent-core/src/index.ts`
- `packages/agent-core/tests/electron-transport.test.ts`
- `apps/docs/src/main/docs-main.ts`
- `apps/docs/src/preload/index.ts`
- `apps/docs/src/shared/ipc.ts`
- `apps/docs/src/renderer/App.tsx`
- `apps/docs/src/renderer/main.tsx`
- `apps/docs/src/renderer/ai/AiPanel.tsx`
- `apps/docs/src/renderer/ai/transport.ts`
- `apps/docs/src/renderer/styles.css`
- `apps/docs/tests/ai-panel-collapse.test.ts`
- One new focused Docs/shared-controls test under `apps/docs/tests/` if keeping
  the existing file focused is clearer
- `plans/README.md`

**Out of scope**:

- Sheets, Slides, PDF, or Markdown wiring; Plans 024 and 025 own it.
- Standalone main-process parity; Plan 026 owns it.
- Provider transport, OAuth, credential storage, or tool execution changes.
- A generic chat panel or generic Electron IPC framework.
- Per-app provider settings or interrupting/restarting an in-flight run when
  settings change.
- Visual redesign of the approved Docs controls.

## Git workflow

- Work on the current branch; do not push or open a PR.
- Keep the extraction and settings-broadcast changes reviewable as separate
  logical commits if committing is requested later.
- Suggested subjects: `share Codex sidebar controls` and
  `synchronize AI settings across tabs`.

## Steps

### Step 1: Extract safe IPC error resolution into the existing transport seam

In `packages/agent-core/src/electron-transport.ts`, add a small exported helper
and message-map type that convert `IpcErrorCode` to caller-supplied localized
strings. It must cover every current code (`timeout`, `credits`, all Codex
codes, and `provider-failure`) and return the unknown fallback for missing or
unrecognized input. Keep translation text outside `agent-core`.

Update Docs transport to call this helper instead of owning its switch. Export
the helper/type through `packages/agent-core/src/index.ts`. Extend the existing
transport test with table-driven coverage for all codes and the unknown case.

**Verify**:

```sh
npm run test -w @genoffice/agent-core
npm run typecheck -w @genoffice/agent-core
npm run typecheck -w @genoffice/docs
```

Expected: all commands exit 0 and the existing transport semantics are
unchanged.

### Step 2: Extract the approved Codex controller and three UI surfaces

Create `packages/ui/src/AiProviderControls.tsx`. Move, rather than rewrite, the
settled Docs behavior into:

- one hook/controller for provider switching, status/login/cancel lifecycle,
  capability loading, stale-result guards, model/reasoning/tier normalization,
  reset, popover positioning/keyboard/outside-click behavior, and
  `sendDisabled`;
- a provider-select component for the panel header;
- an auth/error banner component for the panel body; and
- a model/reasoning/speed component for the composer footer.

The controller must accept structural props rather than a Docs global:

- current `AiSettings` and `onSettingsChange(next)`;
- a redacted API adapter with status/login/cancel/capabilities methods;
- localized labels/error resolver supplied by the app; and
- an optional non-Codex notice so Docs can retain its attachment notice in the
  same banner area.

It must expose state/callbacks used by the three view components plus
`sendDisabled`. It must never read credentials or import an app preload global.
Use types from `@genoffice/ai-provider`; add only that existing workspace
dependency to `packages/ui`. Do not add an external dependency.

Move the matching `.ai-provider-*`, `.ai-codex-*`, and necessary popover rules
from Docs into `packages/ui/src/ai-provider-controls.css`, using existing
semantic tokens and `--accent`. Export the component module and stylesheet
subpath from `packages/ui/package.json`/`src/index.ts`.

**Verify**:

```sh
npm run typecheck -w @genoffice/ui
npm run check:theme-colors
```

Expected: exit 0; no raw renderer chrome colors are introduced.

### Step 3: Recompose Docs from the shared controls without changing UX

Import the shared stylesheet once from `apps/docs/src/renderer/main.tsx` and
replace the moved controller/UI code in `AiPanel.tsx` with the shared hook and
three components. Adapt `window.desktop` through the structural API prop. Keep
the components at their exact approved placements: provider selector in the
header, auth/error banner above chat, and model menu in the composer footer.

Delete only CSS and helpers made dead by the extraction. Keep Docs-specific
attachment notice, track-changes control, agent loop, tools, prompts, and chat
history local. `transport.ts` must supply all localized messages to the shared
error resolver.

Update the existing characterization tests to mock the structural API as they
do today. The tests must still prove:

- only Genspark and ChatGPT Codex appear, in that order;
- Codex send is disabled while status is pending or signed out;
- login, stale provider switches, and stale capability results are safe;
- model/reasoning/speed changes and reset produce normalized settings; and
- collapse does not discard the controller state.

**Verify**:

```sh
npm run test -w @genoffice/docs -- ai-panel-collapse.test.ts
npm run typecheck -w @genoffice/docs
```

Expected: all tests pass and typecheck exits 0.

### Step 4: Broadcast sanitized suite-global settings

Add a typed `onAiSettingsChanged(handler)` subscription to the Docs preload
and shared IPC interface. In `registerAiIpc`, sanitize and persist settings as
today, then send the sanitized value on `ai:settings-changed` to every live
renderer `webContents`. The payload contains provider configuration only, never
Codex credentials or account data. Do not broadcast before persistence
succeeds.

In Docs `App.tsx`, subscribe once, update local `settings`, and unsubscribe on
unmount. Keep the existing optimistic local update in `updateAiSettings`; the
echoed broadcast is an idempotent reconciliation. The AgentLoop transport
already snapshots settings at stream start, so do not cancel or recreate an
active loop.

Add a focused test proving an externally delivered settings value updates the
rendered controls but does not invoke `setAiSettings` again (no event loop).

**Verify**:

```sh
npm run test -w @genoffice/docs
npm run typecheck -w @genoffice/docs
```

Expected: all pass; an external event updates Docs once without another IPC
write.

### Step 5: Run the focused shared-foundation gate

```sh
npm run typecheck -w @genoffice/ui
npm run test -w @genoffice/agent-core
npm run test -w @genoffice/docs
npm run typecheck -w @genoffice/docs
npm run check:theme-colors
git diff --check
```

Expected: every command exits 0.

## Test plan

- Extend `packages/agent-core/tests/electron-transport.test.ts` for every safe
  error-code mapping and unknown fallback.
- Preserve and adapt `apps/docs/tests/ai-panel-collapse.test.ts` as the shared
  control characterization suite.
- Add one renderer test for external settings delivery, idempotent local state,
  and no write-back loop.
- Do not add screenshot tests or a new test framework.

## Done criteria

- [ ] Docs uses shared controller/select/banner/model components with no UX or
      behavioral regression.
- [ ] Shared controls never import an app preload global or receive credentials.
- [ ] Every safe IPC error code resolves through one shared helper.
- [ ] Persisted settings broadcast to every open renderer only after sanitizing
      and saving.
- [ ] External settings update Docs without a write-back loop or active-run
      cancellation.
- [ ] Focused tests, typechecks, theme gate, and `git diff --check` pass.
- [ ] No files outside Scope are modified.
- [ ] `plans/README.md` marks Plan 023 DONE after review.

## STOP conditions

Stop and report back if:

- The approved Docs behavior cannot be preserved without changing its visible
  layout or interaction contract.
- The shared controls would need credentials or Electron APIs in the renderer.
- Settings are not actually shared through the Docs handler in shell mode.
- Broadcasting requires exposing shell tab-manager internals to renderer code.
- Any new dependency other than the existing workspace `ai-provider` package
  appears necessary.
- A focused verification fails twice after a reasonable correction.

## Maintenance notes

- Add future provider-specific sidebar chrome to this shared seam only after a
  second app needs it; do not turn it into a generic chat framework.
- Review async request counters carefully: a response from a provider selected
  earlier must never overwrite the current provider's state.
- `ai:settings-changed` is non-secret suite state. Account status remains
  request/response because it can change independently of provider settings.
