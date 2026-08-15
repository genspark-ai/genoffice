# Plan 026: Give every standalone app the same Codex-capable AI runtime

> **Implementation instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, review the full diff and update this
> plan's status row in `plans/README.md`.
>
> **Drift check (run first)**:
> `git diff --stat cf1f3c1..HEAD -- packages/ai-provider apps/docs apps/sheets apps/slides apps/pdf apps/markdown apps/shell plans`
> Plans 023-025 are expected drift. Recheck all current-state excerpts before
> touching main-process code; mismatches are a STOP condition.

## Status

- **Priority**: P1
- **Effort**: L
- **Risk**: MED
- **Depends on**: Plans 023, 024, and 025
- **Category**: tech-debt
- **Planned at**: commit `cf1f3c1`, 2026-08-16

## Why this matters

The packaged suite already routes all app tabs through one Codex-capable Docs
AI handler, but standalone development does not match production: Sheets and
Slides force Genspark, while PDF and Markdown register no generic AI handlers.
Copying the full stream/auth/capability branch into four apps would immediately
diverge. This plan extracts the now-proven, Electron-independent request
orchestration into the existing provider package, keeps channel registration
and sender validation app-owned, and verifies all five sidebar integrations in
both the unified shell and standalone development paths.

## Current state

- `apps/shell/src/main/index.ts:2816` imports and registers
  `apps/docs/src/main/docs-main.ts::registerAiIpc` once for the packaged suite.
- Docs `registerAiIpc` at `apps/docs/src/main/docs-main.ts:2537-2734` owns:
  settings sanitization/persistence, Genspark credential injection, Codex auth
  status/login/logout/capabilities, capability validation, stream lifecycle,
  safe error codes, activity pings, cancellation, and stream completion.
- The reusable pieces already live in `@genoffice/ai-provider`:
  `CodexAuthService`, `fetchCodexCapabilities`, `streamForProvider`, provider
  defaults/resolution, safe error types, and the Node credential/callback
  adapter. The package intentionally imports no Electron API.
- Sheets standalone `registerSheetsAiIpc` begins at
  `apps/sheets/src/main/sheets-main.ts:2114`; lines 2123-2128 force Genspark.
  It validates renderer input with Zod and calls `sessionFor(event)` at the
  trust boundary. Preserve both.
- Slides standalone `registerAiIpc` begins at
  `apps/slides/src/main/ai-ipc.ts:59`; lines 63-68 force Genspark. Its
  `registerSlidesOnlyAiIpc` contains app-specific image/media channels and must
  remain separate.
- PDF standalone calls only `registerPdfIpc` at
  `apps/pdf/src/main/pdf-main.ts:1128-1139`; Markdown standalone likewise calls
  only `registerMarkdownIpc` at
  `apps/markdown/src/main/markdown-main.ts:762-773`.
- PDF already depends on `@genoffice/ai-search` for image generation. Markdown
  does not, although standalone Genspark and web search require it.
- `packages/ai-provider/src/codex-auth-node.ts` is the established Node-only
  subpath. Do not add Electron imports to `packages/ai-provider`.
- `CLAUDE.md` warns that app main-process code is compiled into the shell build;
  final verification must include `npm run build:all`.

The approved boundaries are:

1. Provider/auth/request orchestration may be shared only as pure TypeScript;
   Electron channel registration remains in each app.
2. Sender/session validation stays at each app's existing IPC trust boundary.
3. Credentials remain in the Node auth service and never cross IPC.
4. Suite settings are one file in shell mode; standalone apps use their own
   normal Electron `userData/ai-settings.json`.
5. App-specific search, image generation, media analysis, and editor-tool
   channels are preserved.

## Commands you will need

| Purpose            | Command                                       | Expected on success          |
| ------------------ | --------------------------------------------- | ---------------------------- |
| Provider tests     | `npm run test -w @genoffice/ai-provider`      | all pass                     |
| Provider typecheck | `npm run typecheck -w @genoffice/ai-provider` | exit 0                       |
| App typecheck      | `npm run typecheck`                           | exit 0 across all workspaces |
| App tests          | `npm test`                                    | all workspace tests pass     |
| Build              | `npm run build:all`                           | all six apps build           |
| Static             | `npm run lint && npm run check:theme-colors`  | exit 0; lint has no errors   |

## Scope

**In scope**:

- `packages/ai-provider/src/main-runtime.ts` (create)
- `packages/ai-provider/src/index.ts`
- `packages/ai-provider/tests/main-runtime.test.ts` (create)
- `apps/docs/src/main/docs-main.ts`
- `apps/docs/src/main/codex-auth-main.ts` only if its factory must accept the
  shared runtime without changing auth behavior
- `apps/sheets/src/main/sheets-main.ts`
- `apps/sheets/src/shared/desktop-api.ts` only for request/runtime schema drift
  discovered while applying Plan 024
- `apps/slides/src/main/ai-ipc.ts`
- `apps/pdf/src/main/pdf-main.ts`
- `apps/markdown/package.json`
- `apps/markdown/src/main/markdown-main.ts`
- Minimal main-process tests under existing app test directories if a thin IPC
  adapter contains non-trivial app-owned branching
- `plans/README.md`

**Out of scope**:

- Renderer layout or shared control changes; Plans 023-025 own them.
- A new workspace package or Electron import in `@genoffice/ai-provider`.
- Changes to OAuth endpoints, credential paths/permissions, callback lifecycle,
  provider wire format, model catalog semantics, or AgentLoop.
- Removing app sender/session validation.
- Consolidating app-specific search/image/media/tool IPC.
- Changing suite-global settings into per-app settings.
- Push, PR creation, or plan-file deletion.

## Git workflow

- Work on the current branch; do not push or open a PR.
- Keep pure runtime extraction, app adapters, and final validation as separate
  logical commits if committing is requested later.
- Suggested subjects: `share AI main runtime`,
  `enable Codex in standalone editors`, and
  `validate suite-wide Codex rollout`.

## Steps

### Step 1: Characterize the current Docs main runtime

Before extracting, add tests in `packages/ai-provider/tests/main-runtime.test.ts`
against the target pure interface using injected auth, Genspark-key, capability,
stream, and clock/fetch collaborators. The characterization matrix must cover:

- settings sanitization permits only Genspark/Codex and normalizes Codex
  reasoning/tier fields;
- Genspark injects its runtime key without persisting it;
- Codex calls `getContext`, validates selected model/reasoning/tier, and passes
  context to `streamForProvider`;
- auth/capability/stream failures return safe codes, never raw Codex bodies;
- pings, deltas, tool calls, stop reason, done, and cancellation preserve their
  current chunk contract; and
- concurrent request IDs own independent abort controllers.

Start with failing tests against the intended public API; do not move Electron
code into the package to make them pass.

**Verify**: `npm run test -w @genoffice/ai-provider -- main-runtime.test.ts`
→ tests execute and fail only because the target module/API does not yet exist.

### Step 2: Extract a pure `AiMainRuntime`

Create `packages/ai-provider/src/main-runtime.ts` and export it from the package
root. It may depend on existing provider/auth modules and standard Web APIs,
but not Electron or filesystem APIs.

The runtime should own only proven shared behavior:

- `sanitizeAiSettings(value)` returning the allowed Genspark/Codex settings;
- redacted Codex status/login/cancel/logout/capabilities methods delegated to
  an injected `CodexAuthService`;
- `stream(request, send, localizedFallbacks)` with settings validation,
  Genspark key injection, capability checks, pings, safe chunks, and cleanup;
- `cancel(requestId)`.

Dependencies must be injected structurally: auth service, Genspark key getter,
stream/capability functions where tests need substitution, and localized
fallback strings for non-Codex errors. The runtime must not read/write settings,
open browsers, register channels, inspect Electron events, or execute tools.

Do not create interfaces for hypothetical transports. This one runtime has five
known consumers and mirrors the already-running Docs behavior.

**Verify**:

```sh
npm run test -w @genoffice/ai-provider -- main-runtime.test.ts
npm run test -w @genoffice/ai-provider
npm run typecheck -w @genoffice/ai-provider
rg -n "from 'electron'|from \"electron\"" packages/ai-provider
```

Expected: tests/typecheck pass; the final search returns no matches.

### Step 3: Make Docs the first thin Electron adapter

Refactor only the generic AI portion of `registerAiIpc` to instantiate/use
`AiMainRuntime`. Keep Docs-owned file persistence, browser opening, Genspark
login/search handlers, channel names, shell-wide settings broadcast from Plan
023, and `webContents` sending in Docs.

For every renderer→main request, preserve current validation/sanitization and
safe result shape. `ai:set-settings` must persist the runtime-sanitized value
before broadcasting it. The shell must continue importing the same exported
`registerAiIpc` symbol; do not change shell lifecycle code.

Run the existing Docs Codex/provider tests before continuing. A behavioral
change here is a failed extraction, not an opportunity to revise the contract.

**Verify**:

```sh
npm run test -w @genoffice/docs
npm run typecheck -w @genoffice/docs
npm run typecheck -w @genoffice/shell
```

Expected: all pass.

### Step 4: Replace Sheets and Slides standalone provider branches

In Sheets, keep `registerSheetsAiIpc`, `sessionFor(event)`, Zod parsing, settings
file path, Genspark login/search, and per-session stream ownership at the IPC
boundary. Delegate sanitized provider/auth/capability/stream behavior to one
runtime instance. Remove the forced-Genspark normalization. Ensure every auth
and stream handler validates the sender using the existing session convention.

In Slides, keep `registerAiIpc` as the generic standalone adapter and leave
`registerSlidesOnlyAiIpc` untouched. Delegate to the runtime, remove forced
Genspark, and retain Slides' localized main-process fallbacks and rescue fetch.

Both apps construct `CodexAuthService` with the existing Node credential store,
callback adapter, `safeExternalUrl`, and `shell.openExternal`, matching
`apps/docs/src/main/codex-auth-main.ts`; do not copy storage/callback
implementations.

**Verify**:

```sh
npm run typecheck -w @genoffice/sheets
npm run test -w @genoffice/sheets
npm run typecheck -w @genoffice/slides
npm run test -w @genoffice/slides
```

Expected: all pass; searches for forced provider assignment in these handlers
return no matches:

```sh
rg -n "settings\.provider = 'genspark'" apps/sheets/src/main apps/slides/src/main
```

### Step 5: Register generic AI handlers in PDF and Markdown standalone mode

Add a small, idempotent generic AI registration function in each app main file
using the shared runtime. Reuse each app's existing Node/fs helpers where
available to read and atomically write `userData/ai-settings.json`; do not add a
settings library. Register the function only from standalone startup. In shell
mode, Docs still owns generic `ai:*` handlers, so `createPdfView` and
`createMarkdownView` must not double-register them.

PDF already has `@genoffice/ai-search`; reuse it for Genspark status/login and
the generic search functions required by its tools. Preserve PDF-owned
`ai:generate-image` behavior.

Add existing workspace dependency `@genoffice/ai-search` to Markdown and use
it for standalone Genspark status/login and web search. Do not add an external
dependency. Both apps construct Codex auth exactly as in Step 4 and register
settings-changed broadcasts (one renderer in standalone mode).

**Verify**:

```sh
npm run typecheck -w @genoffice/pdf
npm run test -w @genoffice/pdf
npm run typecheck -w @genoffice/markdown
npm run test -w @genoffice/markdown
```

Expected: all pass; standalone startup paths call generic registration once,
while hosted view creation does not.

### Step 6: Run automated repository gates

```sh
npm run format:check
npm run check:theme-colors
npm run lint
npm run typecheck
npm test
npm run build:all
npm run licenses
npm audit --omit=dev --audit-level=high
git diff --check
```

Expected: every command exits 0; lint has no errors; audit has no high/critical
production advisory.

### Step 7: Exercise suite-global behavior in the unified shell

Run the unified app with a non-privileged test account. Across Docs, Sheets,
Slides, PDF, and Markdown:

1. Select Codex in one sidebar; all already-open sidebars update without reload.
2. Change model, reasoning, and speed in a second app; all sidebars reflect it.
3. Confirm the change affects the next send only; an active run is not canceled
   or retargeted.
4. Signed-out status disables Send in every app and offers the same login flow.
5. Sign in once; every app obtains redacted status/capabilities without another
   credential file or token in renderer devtools.
6. Run one tool-capable edit in each app and verify its app-owned tools still
   operate on the correct document type.
7. Switch back to Genspark in any app; all sidebars update and next sends use
   Genspark.
8. Restart the suite; global selection and Codex model/reasoning/speed persist.

Record pass/fail and platform without account IDs, tokens, callback URLs, or
raw provider payloads.

### Step 8: Exercise standalone parity

Launch each app through its workspace `dev` command (one at a time so shared
channel names cannot collide). For Sheets, Slides, PDF, and Markdown verify:

- Genspark and Codex appear;
- Codex status/login/capabilities work;
- one normal tool-capable turn completes;
- cancel settles the UI;
- provider/model settings persist across restart; and
- no "No handler registered" error appears.

If a standalone command needs a representative file, use a repository fixture
and do not overwrite it; save to a temporary copy.

## Test plan

- New provider-package runtime tests cover shared orchestration and cancellation
  without Electron.
- Existing provider/auth tests remain authoritative for OAuth/storage/wire
  semantics.
- Existing app suites cover thin adapters and app-owned tools.
- Manual unified-shell matrix proves live suite-global settings and one login.
- Manual standalone matrix proves registration parity and persistence.

## Done criteria

- [ ] One Electron-independent `AiMainRuntime` owns shared settings/auth/stream
      orchestration; `@genoffice/ai-provider` still has no Electron import.
- [ ] Docs remains the sole generic AI handler in unified shell mode.
- [ ] Sheets and Slides standalone no longer force Genspark.
- [ ] PDF and Markdown standalone register all generic AI handlers exactly once.
- [ ] Every renderer→main trust boundary preserves its existing sender/session
      validation.
- [ ] Unified settings changes propagate live across all five app sidebars and
      apply to the next send only.
- [ ] No credential/token/account identifier crosses preload or appears in logs.
- [ ] Full format, theme, lint, typecheck, test, build, license, audit, and diff
      gates pass.
- [ ] Unified-shell and standalone manual matrices pass and are recorded safely.
- [ ] No files outside Scope are modified.
- [ ] `plans/README.md` marks Plan 026 DONE after review.

## STOP conditions

Stop and report back if:

- Sharing the runtime requires an Electron import in `ai-provider` or a new
  workspace package.
- Any app would need to import another app's main module in standalone mode.
- Preserving Sheets sender validation or per-session cancellation is
  incompatible with the shared runtime API.
- PDF/Markdown hosted views would double-register generic `ai:*` handlers.
- A credential, token, authorization URL, or raw Codex response would cross IPC
  or enter logs/tests.
- Extraction changes provider request, OAuth, storage, or model semantics.
- A full gate or live scenario fails twice after a reasonable correction.

## Maintenance notes

- `AiMainRuntime` is orchestration, not an Electron framework: future apps
  still own channels, storage location, localization, sender validation, and
  app-specific tools.
- Review cancellation ownership carefully. A request ID must be scoped to the
  runtime instance/session that started it and removed in `finally`.
- When adding a new provider-specific setting, update runtime sanitization,
  preload schemas, shared controls, and the global-settings event together.
