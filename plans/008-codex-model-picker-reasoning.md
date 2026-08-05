# Plan 008: Add an available Codex model and reasoning picker

> **Implementation instructions**: Follow this plan step by step. Run every
> verification command. Update `plans/README.md` only after implementation and
> review pass. Do not commit or push.
>
> **Drift check (run first)**: `git diff --stat 4da673d -- <in-scope paths>` and
> `git diff --stat -- <in-scope paths>`. The current branch contains the Docs
> Codex mod as uncommitted work; compare the excerpts below with live files.
> Stop if the provider/settings or IPC shapes differ materially.

## Status

- **Priority**: P1
- **Effort**: M
- **Risk**: MED — private backend capability discovery and request compatibility
- **Depends on**: `plans/006-docs-codex-mode.md`
- **Category**: direction
- **Planned at**: commit `4da673d`, 2026-08-04 (with current uncommitted Codex mod)

## Why this matters

Docs currently exposes provider selection but hides the Codex model and sends a
single hard-coded model (`gpt-5.5`). The private Codex Responses endpoint can
reject an otherwise valid authenticated request when model or reasoning fields
are unsupported, leaving users unable to select a compatible account model.
This plan adds a main-process capability boundary, a Codex-only model picker,
persisted selection, and reasoning effort passed through the Responses request
without exposing credentials or trusting renderer-supplied model lists.

## Current state

- `packages/ai-provider/src/providers.ts:66-75` — `openai-codex` metadata has
  static `models: ['gpt-5.5']` and `defaultModel: 'gpt-5.5'`.
- `packages/ai-provider/src/types.ts:40-55` — `CodexAdapterRequest` has model
  but no reasoning setting; `AiProviderConfig` persists only `apiKey`, model,
  and optional base URL.
- `packages/ai-provider/src/codex.ts:66-78` — request body sends model,
  instructions, input, tools, `store: false`, and `stream: true`; reasoning is
  absent. Match existing stateless request construction and never execute tools
  in this module.
- `apps/docs/src/renderer/ai/AiPanel.tsx:735-755` — header has one provider
  `<select>` and no model/reasoning controls. `changeProvider` updates
  `AiSettings` and `App.tsx:572-576` persists settings through typed IPC.
- `apps/docs/src/shared/ipc.ts:160-205` — settings IPC is typed; all secrets
  must remain main-process-only. Existing Codex auth IPC returns only
  `CodexAccountStatus`.
- `apps/docs/src/main/docs-main.ts:2510-2575` — main process reads Codex auth,
  then calls provider streaming. Any capability request must use this boundary,
  not renderer credentials.
- Existing conventions: provider metadata is the source for labels/defaults;
  settings updates are immutable in `AiPanel`/`App`; tests use Vitest and
  `apps/docs/tests/ai-panel-collapse.test.ts` for rendered controls and mocked
  `window.desktop`; provider request tests live in
  `packages/ai-provider/tests/codex.test.ts`.

### Capability contract

Add a typed main-process operation that returns only model IDs and supported
reasoning efforts for the currently authenticated Codex account. It may call
the backend capability endpoint with main-process credentials, but must not
persist or return access/refresh tokens, account headers, raw response bodies,
or arbitrary URLs. If the endpoint is unavailable or its response shape cannot
be verified with mocked tests, stop and report instead of inventing a fallback.

Reasoning values must be an explicit allowlist (for example, `none`, `low`,
`medium`, `high`) derived from capability data or a documented provider
contract. Send selected effort using the Codex Responses schema
`reasoning: { effort: <value> }`; omit `reasoning` when effort is `none`.

## Commands you will need

| Purpose            | Command                                       | Expected on success     |
| ------------------ | --------------------------------------------- | ----------------------- |
| Provider tests     | `npm test -w @genoffice/ai-provider`          | all provider tests pass |
| Docs tests         | `npm test -w @genoffice/docs`                 | all Docs tests pass     |
| Provider typecheck | `npm run typecheck -w @genoffice/ai-provider` | exit 0                  |
| Docs typecheck     | `npm run typecheck -w @genoffice/docs`        | exit 0                  |
| Formatting         | `npm run format:check`                        | exit 0                  |

## Scope

**In scope** (only these files; create tests where noted):

- `packages/ai-provider/src/types.ts`
- `packages/ai-provider/src/providers.ts`
- `packages/ai-provider/src/codex.ts`
- `packages/ai-provider/src/index.ts`
- `packages/ai-provider/tests/codex.test.ts`
- `packages/ai-provider/tests/providers.test.ts`
- `apps/docs/src/main/docs-main.ts`
- `apps/docs/src/shared/ipc.ts`
- `apps/docs/src/preload/index.ts`
- `apps/docs/src/renderer/ai/AiPanel.tsx`
- `apps/docs/src/renderer/i18n/strings-ai.ts`
- `apps/docs/tests/ai-panel-collapse.test.ts` (or the existing Docs AI UI test file if it has moved)
- `plans/README.md`

**Out of scope**:

- Other providers' model pickers or settings UX.
- Reading `~/.codex/auth.json` or adding a CLI/runtime dependency.
- Renderer access to tokens, account IDs, capability URLs, or raw backend
  diagnostics.
- Changes to `AgentLoop`, Docs tools, tool schemas, OAuth storage, or the
  Genspark path.
- Live-account tests as a substitute for mocked capability/request tests.

## Steps

### Step 1: Define model/reasoning types and safe defaults

Extend provider-side types with a Codex reasoning-effort union, capability
shape, and request field. Keep persisted settings JSON-compatible and migrate
missing values deterministically. Preserve legacy settings and ensure a model
outside the verified capability list cannot be sent by default.

**Verify**: `npm run typecheck -w @genoffice/ai-provider` and provider metadata
tests pass; request types contain no credential fields.

### Step 2: Implement capability retrieval and request translation

Add injectable capability fetching in the provider/main boundary. Validate
response schema, deduplicate model IDs, preserve a safe fallback only when
explicitly provided by the verified contract, and normalize unsupported model
or reasoning selections before a request. Update `buildCodexRequest` to emit
the selected model and optional `reasoning` object. Keep `store: false`, Docs
tool schemas, and existing message/image translation unchanged.

**Verify**: extend `packages/ai-provider/tests/codex.test.ts` for capability
schema validation, model selection, each reasoning value, omitted `none`, and
redaction; `npm test -w @genoffice/ai-provider` passes.

### Step 3: Expose typed main/preload capability IPC

Add a Docs IPC method returning only the validated capability shape. Main reads
Codex auth via `CodexAuthService`, performs capability fetch, and maps 401/403,
network, malformed, and unsupported responses to safe user-facing errors.
Preload exposes the typed method. Never accept a renderer URL or token.

**Verify**: Docs typecheck/tests pass; IPC tests assert no token/account header
or raw response body appears in any renderer-visible result.

### Step 4: Add Codex-only model and reasoning controls

In `AiPanel`, render model and reasoning selects only when provider is
`openai-codex`. Load capabilities on Codex selection, show a localized loading
and unavailable state, retain the last valid selection when refresh fails, and
persist model/reasoning through existing `onSettingsChange`/settings IPC.
Do not alter controls for Genspark or API-key providers. Make labels/options
accessible and use existing compact header styling.

**Verify**: Docs UI tests cover Codex-only visibility, capability options,
persisted selection, disabled/unavailable state, and no controls for Genspark;
`npm test -w @genoffice/docs` passes.

### Step 5: Complete integration and review

Wire selected reasoning into every Codex stream request, confirm a model change
does not reset auth or conversation history, and update localized strings for
all supported locales with safe English fallback. Review diff against the
in-scope list and update Plan 008 status to DONE only after all gates pass.

**Verify**: run provider tests, Docs tests, both typechecks, and
`npm run format:check`; all exit 0. `git diff --check` is clean and no
out-of-scope files changed.

## Test plan

- Provider tests: capability schema, model allowlist, reasoning serialization,
  unsupported-selection fallback, malformed/secret-bearing error redaction.
- Docs tests: Codex-only controls, loading/error state, model/reasoning change
  persistence, IPC payload redaction, and stream request propagation.
- Structural pattern: use existing `packages/ai-provider/tests/codex.test.ts`,
  `providers.test.ts`, and `apps/docs/tests/ai-panel-collapse.test.ts`.

## Done criteria

- [ ] Codex model picker displays only validated available model IDs.
- [ ] Reasoning picker displays only validated efforts; `none` omits field.
- [ ] Selected model/effort persist and reach Codex request body.
- [ ] No token, account ID, raw backend body, or arbitrary URL reaches renderer.
- [ ] Non-Codex provider UX and Genspark auth behavior unchanged.
- [ ] Provider/Docs tests, typechecks, and format check pass.
- [ ] No files outside Scope changed; README status row updated.

## STOP conditions

- Backend capability endpoint is undocumented/unavailable and no safe validated
  model list can be established from existing provider contract.
- Reasoning schema differs from `reasoning: { effort }` and cannot be mocked
  and verified without guessing.
- Capability data would require renderer-held credentials or raw diagnostics.
- Implementer must modify `AgentLoop`, tool schemas, OAuth storage, or another
  provider to make picker work.
- Any verification gate fails twice or current-state excerpts drift materially.

## Maintenance notes

- Backend model availability is account- and time-dependent; keep capability
  validation centralized and review fallback behavior when Codex changes.
- Reviewers should check that picker options come from validated main-process
  data and that persisted settings cannot bypass the allowlist.
- Live endpoint compatibility remains Plan 007's operational gate; this plan
  must not weaken Docs tool boundaries to accommodate it.
