# Plan 024: Roll the shared Codex controls into Sheets and Slides

> **Implementation instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update this plan's status row in
> `plans/README.md` after implementation and review.
>
> **Drift check (run first)**:
> `git diff --stat cf1f3c1..HEAD -- apps/sheets apps/slides packages/ui packages/agent-core plans`
> Plan 023 is expected drift. If Sheets/Slides excerpts no longer match, stop
> and refresh this plan before implementation.

## Status

- **Priority**: P1
- **Effort**: M
- **Risk**: MED
- **Depends on**: Plan 023
- **Category**: direction
- **Planned at**: commit `cf1f3c1`, 2026-08-16

## Why this matters

Sheets and Slides already use the shared provider settings and streaming
transport, but neither sidebar lets the user select Codex, sign in, or choose
validated Codex capabilities. Both apps keep settings in their top-level App,
making them one rollout shape. This plan wires the shared controls without
changing workbook/slide agents, generation overrides, attachments, or editing
tools, and subscribes both mounted tabs to suite-global settings changes.

## Current state

- `apps/sheets/src/renderer/App.tsx:582-584` owns `aiSettings` and the ref read
  by the AgentLoop; lines 1126-1128 load it once.
- `apps/sheets/src/renderer/ExcelShell.tsx:434-455` renders `AiChatPanel` but
  passes no settings or update callback.
- `apps/sheets/src/renderer/ai/AiChatPanel.tsx` is presentation-focused and has
  no provider/auth/model controls.
- `apps/sheets/src/preload/index.ts:258-265` exposes get/set settings; the
  `DesktopApi` at `apps/sheets/src/shared/desktop-api.ts:1963-1977` exposes
  streaming and Genspark account methods only.
- `apps/sheets/src/shared/desktop-api.ts:1731-1737` defines a strict provider
  config schema containing only `apiKey`, `model`, and `baseUrl`. It currently
  rejects the Codex `reasoningEffort` and `serviceTier` fields.
- `apps/slides/src/renderer/App.tsx:330` owns settings and lines 1002-1004 load
  them once; `AiPanel` receives settings at lines 2665-2677 but no update
  callback.
- `apps/slides/src/renderer/ai/AiPanel.tsx:254-288` declares its props and owns
  the slide-specific AgentLoop/tool/generation behavior.
- `apps/slides/src/preload/index.ts:288-326` exposes get/set/stream and Genspark
  methods but no Codex or settings-change methods; the matching interface is
  `apps/slides/src/shared/ipc.ts:1329-1337`.
- `apps/sheets/src/renderer/ai/transport.ts` and
  `apps/slides/src/renderer/ai/transport.ts` use `createIpcTransport` but do not
  resolve Codex machine error codes to localized messages.
- App strings follow `defineStrings`: every key added to the Chinese dictionary
  must exist in all 18 other locale dictionaries. Docs
  `apps/docs/src/renderer/i18n/strings-ai.ts` is the approved translation source
  for Codex labels and errors; reuse those translations exactly.
- Slides' generation-only Anthropic model override at
  `apps/slides/src/renderer/ai/AiPanel.tsx:688-699` is intentional. Codex must
  continue using the user-selected Codex model and must not receive the
  Anthropic-only override.

## Commands you will need

| Purpose          | Command                                  | Expected on success              |
| ---------------- | ---------------------------------------- | -------------------------------- |
| Sheets tests     | `npm run test -w @genoffice/sheets`      | all pass, including native gates |
| Sheets typecheck | `npm run typecheck -w @genoffice/sheets` | exit 0                           |
| Slides tests     | `npm run test -w @genoffice/slides`      | all pass                         |
| Slides typecheck | `npm run typecheck -w @genoffice/slides` | exit 0                           |
| Theme gate       | `npm run check:theme-colors`             | exit 0                           |

## Scope

**In scope**:

- `apps/sheets/src/shared/ipc-channels.ts`
- `apps/sheets/src/shared/desktop-api.ts`
- `apps/sheets/src/preload/index.ts`
- `apps/sheets/src/renderer/main.tsx`
- `apps/sheets/src/renderer/App.tsx`
- `apps/sheets/src/renderer/ExcelShell.tsx`
- `apps/sheets/src/renderer/ai/AiChatPanel.tsx`
- `apps/sheets/src/renderer/ai/transport.ts`
- `apps/sheets/src/renderer/i18n/strings-ai.ts`
- `apps/sheets/src/renderer/styles.css`
- One focused Sheets provider-control test under `apps/sheets/tests/`
- `apps/slides/src/shared/ipc.ts`
- `apps/slides/src/preload/index.ts`
- `apps/slides/src/renderer/main.tsx`
- `apps/slides/src/renderer/App.tsx`
- `apps/slides/src/renderer/ai/AiPanel.tsx`
- `apps/slides/src/renderer/ai/transport.ts`
- `apps/slides/src/renderer/i18n/strings-ai.ts`
- `apps/slides/src/renderer/styles.css`
- `apps/slides/tests/ai-panel-collapse.test.ts`
- `plans/README.md`

**Out of scope**:

- Sheets or Slides main-process provider execution; Plan 026 owns standalone
  parity, while packaged shell mode already uses the shared handler.
- Workbook/slide prompts, skills, tools, deterministic planner, history, or
  attachments.
- Slides generation pipeline/model override changes.
- PDF and Markdown; Plan 025 owns them.
- App-specific forks of the shared controls or per-app provider settings.
- OAuth, credentials, provider transport, or shared-control redesign.

## Git workflow

- Work on the current branch; do not push or open a PR.
- Keep Sheets and Slides changes as separable logical commits if committing is
  requested later.
- Suggested subjects: `add Codex controls to Sheets` and
  `add Codex controls to Slides`.

## Steps

### Step 1: Extend both preload contracts with safe shared AI methods

Add channel constants/methods for:

- `ai:set-settings` where absent;
- `ai:settings-changed` subscription;
- `ai:codex-status`;
- `ai:codex-login`;
- `ai:codex-cancel-login`;
- `ai:codex-logout`; and
- `ai:codex-capabilities`.

Use exported account/capability/error types from `@genoffice/ai-provider`.
Preloads must validate response shape consistently with each app's existing
style, but must never expose token fields. `onAiSettingsChanged` returns an
unsubscribe function.

For Sheets, extend `aiProviderConfigSchema` with optional validated
`reasoningEffort` and `serviceTier` fields matching Docs' allowlist/pattern.
Keep the schema strict and add schema tests within the new focused test or an
existing desktop API schema test. This is required because every stream embeds
the settings object.

**Verify**:

```sh
npm run typecheck -w @genoffice/sheets
npm run typecheck -w @genoffice/slides
```

Expected: both exit 0; preload-exposed objects satisfy their shared API types.

### Step 2: Make Sheets consume and publish suite-global settings

In Sheets `App.tsx`, add one `updateAiSettings(next)` callback that updates
state/ref immediately and awaits or fire-and-forgets `setAiSettings` using the
existing app style. Subscribe once to `onAiSettingsChanged`, update local state
only, and unsubscribe on unmount. Do not write the received value back.

Thread `aiSettings` and `updateAiSettings` through `ExcelShell.tsx` into
`AiChatPanel`. In the panel, adapt `window.desktopApi` to Plan 023's structural
Codex API and render the shared provider selector/banner/model control at the
same semantic locations as Docs. Pass the controller's `sendDisabled` into the
existing `AiComposer`; keep `aiBusy`, attachment notice, new-chat, undo, and
deterministic fallback behavior unchanged.

Import the shared control stylesheet from `renderer/main.tsx`; add only
Sheets-specific layout overrides to local CSS when the existing panel geometry
requires them.

Update transport error resolution with Plan 023's shared helper. Add the Docs
Codex labels/error strings to every Sheets locale by copying the approved
translations exactly.

**Verify**:

```sh
npm run typecheck -w @genoffice/sheets
npm run test -w @genoffice/sheets
```

Expected: exit 0; existing workbook tests remain green.

### Step 3: Make Slides consume and publish suite-global settings

Add `updateAiSettings(next)` and an external settings subscription in Slides
`App.tsx` with the same no-write-back rule. Pass the update callback to
`AiPanel`. Adapt `window.slidesApi` to the shared controller and place the
selector/banner/model control consistently with Docs while preserving the
Slides header actions and composer layout.

Use `sendDisabled` on every user-send entry point that uses the normal current
settings. Do not disable Stop. Do not change or apply Codex controls to the
Anthropic-only deck-generation override; when provider is Codex, generation
must keep the selected Codex configuration.

Import the shared stylesheet, map error codes through the Plan 023 helper, and
copy approved Codex translations into every Slides locale.

**Verify**:

```sh
npm run typecheck -w @genoffice/slides
npm run test -w @genoffice/slides
```

Expected: exit 0; all existing slide agent/generation/collapse tests pass.

### Step 4: Add thin per-app integration tests

For Sheets, add one jsdom test that mounts `AiChatPanel` with mocked APIs and
asserts:

- the provider order is Genspark then ChatGPT Codex;
- selecting Codex calls `onSettingsChange`;
- signed-out/pending Codex disables Send but not draft editing;
- validated capabilities render the model menu; and
- an external settings callback updates the parent state without another
  `setAiSettings` call.

For Slides, extend `ai-panel-collapse.test.ts` with the same integration seam,
plus an assertion that the panel remains mounted across collapse. Do not repeat
the full shared controller matrix already covered by Docs; test only app wiring.

**Verify**:

```sh
npm run test -w @genoffice/sheets
npm run test -w @genoffice/slides
```

Expected: both pass with the new app-level checks.

### Step 5: Run the two-app gate

```sh
npm run typecheck -w @genoffice/ui
npm run typecheck -w @genoffice/sheets
npm run typecheck -w @genoffice/slides
npm run test -w @genoffice/sheets
npm run test -w @genoffice/slides
npm run check:theme-colors
git diff --check
```

Expected: every command exits 0.

## Test plan

- Sheets schema accepts valid Codex settings and rejects invalid reasoning/tier
  input without weakening strict validation.
- One Sheets sidebar integration test covers API wiring and global updates.
- Extend Slides' existing collapse test for provider/auth/model wiring.
- Rely on Plan 023's shared characterization tests for controller edge cases;
  do not clone them into each app.

## Done criteria

- [ ] Sheets and Slides display the approved shared Codex controls.
- [ ] Both apps can persist a setting and react immediately to another tab's
      `ai:settings-changed` event without a write-back loop.
- [ ] Codex settings survive Sheets' strict request validation.
- [ ] Pending/signed-out Codex disables all normal Send paths, not Stop or draft
      editing.
- [ ] Slides' Anthropic-only generation override remains unchanged.
- [ ] All focused tests, typechecks, theme gate, and `git diff --check` pass.
- [ ] No files outside Scope are modified.
- [ ] `plans/README.md` marks Plan 024 DONE after review.

## STOP conditions

Stop and report back if:

- Either app bypasses `aiSettingsRef` and captures stale settings inside its
  AgentLoop.
- Sheets schema cannot accept Codex fields without weakening request boundary
  validation.
- Slides requires changing its generation contract or tool definitions.
- The shared controls need an app-specific fork rather than a small structural
  adapter/layout override.
- A renderer receives any credential/token field.
- A focused verification fails twice after a reasonable correction.

## Maintenance notes

- Review every send entry point in Slides; it has normal chat, generation, and
  corrective runs, and only provider-authored sends should be gated.
- Keep the suite settings subscription one-way: main broadcasts persisted
  values; renderers update local state but never echo event payloads.
- App-specific panel content stays local. Only provider chrome and safe error
  handling are shared.
