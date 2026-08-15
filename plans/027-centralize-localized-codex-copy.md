# Plan 027: Centralize localized Codex UI copy

- Status: DONE
- Priority: P1
- Effort: L
- Risk: MED
- Depends on: Plans 023–026
- Planned at: commit 2c56c4f, 2026-08-16

## Outcome

Make Codex UI copy a single, typed, localized source for Docs, Sheets, Slides,
PDF, and Markdown. The shared source must own the Codex brand/chrome labels,
reasoning-effort display values, and IPC error copy. App renderers may compose
the shared catalog into their existing dictionaries, but they must not define
their own Codex label maps, error maps, or codex-strings.ts files.

The public behavior remains the same except for fixing drift: every app uses
the active locale for all eight supported reasoning IDs, including xhigh as the
display label “Extra High” in English, and every app resolves the same IPC error
code to the same localized message.

This is a targeted implementation plan, not a full repository audit. The
current worktree already contains the accepted Plans 023–026 implementation
and follow-up UI edits; those changes are intentionally outside this plan and
must not be reverted or reformatted.

## Why this boundary

packages/i18n already provides the dependency-free typed primitives:

- Lang and LANGS in packages/i18n/src/index.ts:1-42.
- defineStrings in packages/i18n/src/index.ts:150-160, which enforces the
  same key set for all 19 supported locales.
- createI18n in packages/i18n/src/index.ts:185-194, which provides the
  locale lookup and interpolation behavior.

The Codex effort IDs and error contract belong to @genoffice/ai-provider, not
to a renderer or generic UI package. packages/ai-provider/src/types.ts:23-38
already defines CodexReasoningEffort and the Codex error codes, and
@genoffice/ui already depends on @genoffice/ai-provider. Therefore the
smallest safe dependency direction is:

```text
@genoffice/i18n  <-  @genoffice/ai-provider  <-  @genoffice/ui  <-  app renderers
```

@genoffice/ai-provider will consume the generic i18n core and export a pure
TypeScript Codex catalog/factory. This avoids making the generic i18n package
know about Codex and avoids making the provider depend on a UI package. No
React, Electron, new workspace package, or new runtime dependency is needed.

## Current state and evidence

The defect is architectural, not a single bad translation:

1. packages/ui/src/AiProviderControls.tsx:26-37 requires callers to provide
   reasoningLabel(value) and resolveError(code), so every app can implement the
   contract differently. The control consumes the callback in AiCodexModelControl
   around :589-714.
2. Docs has three separate label objects in
   apps/docs/src/renderer/ai/AiPanel.tsx: the main model control has a full
   title-case map, the auth banner passes raw effort IDs, and the composer
   repeats another full map.
3. Sheets and Slides each define a complete local effort map in their
   providerLabels helper. PDF and Markdown return the raw ID for every value
   other than none.
4. Four apps have copied renderer catalogs:

   - apps/sheets/src/renderer/i18n/codex-strings.ts
   - apps/slides/src/renderer/i18n/codex-strings.ts
   - apps/pdf/src/renderer/i18n/codex-strings.ts
   - apps/markdown/src/renderer/i18n/codex-strings.ts

   Docs instead embeds Codex keys directly in
   apps/docs/src/renderer/i18n/strings-ai.ts. The copies include different
   key names (aiSwitchModel versus aiSwitchModelTitle) and do not provide
   localized reasoning values.

5. All five renderer transports duplicate the same codexErrorText mapping from
   resolveIpcErrorCode in their respective transport.ts files. Slides also
   resolves a stream error directly in AiPanel.tsx.
6. packages/agent-core/src/electron-transport.ts:14-36 is already the canonical
   IPC error-code normalization boundary. The new Codex catalog must consume
   that normalized code; it must not duplicate transport parsing.

The implementation must begin by rechecking these symbols and paths. If any
have moved or the shared UI contract is materially different, stop and update
this plan before editing source.

## Target design

### 1. Provider-owned typed catalog

Create packages/ai-provider/src/codex-i18n.ts with no React or Electron imports.
It should export:

- codexStrings, created with defineStrings, containing one complete set of
  keys for all Lang values.
- CodexUiLabels, the structural label contract consumed by the shared UI.
- getCodexUiLabels(lang: Lang): CodexUiLabels, backed by createI18n.
- resolveCodexError(code: IpcErrorCode | string | undefined, lang: Lang): string
  or an equivalent resolver exposed through the returned label object. There
  must be one implementation of the error-code-to-copy mapping.

The catalog key set must cover the existing shared Codex chrome and the new
reasoning values:

```text
aiCodexBrand
aiCodexLoginBtn
aiCodexSignInRequired
aiSwitchModelTitle
aiCodexModelLabel
aiCodexReasoningLabel
aiCodexSpeedLabel
aiResetDefault
aiCodexModelsUnavailable

aiCodexReasoningNone
aiCodexReasoningMinimal
aiCodexReasoningLow
aiCodexReasoningMedium
aiCodexReasoningHigh
aiCodexReasoningXHigh
aiCodexReasoningMax
aiCodexReasoningUltra

aiCodexTimeout
aiCodexCreditsExhausted
aiCodexAuthRequired
aiCodexAuthExpired
aiCodexAuthTemporary
aiCodexCapabilitiesUnavailable
aiCodexRateLimited
aiCodexRequestRejected
aiCodexInvalidStream
aiCodexInvalidToolCall
aiCodexProviderFailure
aiCodexUnknownError
```

Use the existing localized Codex wording as the baseline, reconcile duplicate
variants once, and add a natural translation for every reasoning key in all
19 locales. English must be exactly None, Minimal, Low, Medium, High,
Extra High, Max, and Ultra. Do not display the raw IDs as a fallback for a
supported effort. If an existing locale has no defensible translation, stop
for a language decision instead of silently shipping mixed case or raw IDs.

Keep the effort mapping exhaustive and type-driven:

```ts
const reasoningKeys: Record<CodexReasoningEffort, CodexStringKey> = { ... }
```

Adding a new union member in types.ts must then fail the catalog typecheck
until its localized key and display mapping are added.

CodexUiLabels should preserve the existing UI fields so this remains a small
API migration, but its callbacks must accept the normalized error input needed
by the shared resolver. AiProviderControlLabels in @genoffice/ui should become
a type alias/re-export of this provider-owned contract, preserving the existing
UI export for callers while preventing each app from inventing a different
shape.

### 2. App composition and renderer use

Each app should obtain one label bundle from the active locale and pass that
same object to its ProviderSelect, AuthBanner, and ModelControl instances:

```ts
const labels = getCodexUiLabels(lang)
```

Replace every inline providerLabels/reasoningLabel map in:

- apps/docs/src/renderer/ai/AiPanel.tsx
- apps/sheets/src/renderer/ai/AiChatPanel.tsx
- apps/slides/src/renderer/ai/AiPanel.tsx
- apps/pdf/src/renderer/ai/AiPanel.tsx
- apps/markdown/src/renderer/ai/AiPanel.tsx

Use the active lang from each app's existing useI18n() result. PDF and
Markdown currently do not expose the module-level getLang helper used by the
other apps; add that minimal read-only accessor to their locale modules only
if the transport closure needs it. Do not put the unstable t function in a
React dependency array; follow the existing CLAUDE.md i18n guidance.

For app dictionaries, import codexStrings from @genoffice/ai-provider and
compose its locale entry into the existing defineStrings dictionary. Remove
the four renderer-local catalog files and remove the inline Codex definitions
from Docs. Existing app-specific AI prompts, tools, generic error copy, and
non-Codex strings remain app-owned. Normalize Slides' legacy aiSwitchModel
call sites to the shared aiSwitchModelTitle key rather than creating a second
shared alias.

### 3. Shared error resolution

Replace the five copies of codexErrorText in:

- apps/docs/src/renderer/ai/transport.ts
- apps/sheets/src/renderer/ai/transport.ts
- apps/slides/src/renderer/ai/transport.ts
- apps/pdf/src/renderer/ai/transport.ts
- apps/markdown/src/renderer/ai/transport.ts

with the provider resolver. Preserve each transport's public return shape,
IPC validation, timeout behavior, and caller-facing fallback behavior. The only
moved responsibility is mapping a normalized code to localized Codex copy.
Migrate Slides' direct stream-error call to the same resolver. After the
migration, app transports must not contain a local resolveIpcErrorCode map.

The shared resolver should cover every current IpcErrorCode: timeout, credits,
auth-required, auth-expired, auth-temporary, capabilities-unavailable,
rate-limit, request-rejected, invalid-stream, invalid-tool-call,
provider-failure, plus the unknown fallback. Preserve the existing distinction
between auth-required, auth-expired, and transient auth failures.

## In-scope files

Create or modify only the following implementation files, plus the plan index
and lockfile required by dependency metadata:

```text
packages/ai-provider/src/codex-i18n.ts                 (create)
packages/ai-provider/src/index.ts
packages/ai-provider/package.json
packages/ai-provider/tests/codex-i18n.test.ts          (create)
packages/ui/src/AiProviderControls.tsx
packages/ui/src/index.ts                               (only if exports need adjustment)

apps/docs/src/renderer/ai/AiPanel.tsx
apps/docs/src/renderer/ai/transport.ts
apps/docs/src/renderer/i18n/strings-ai.ts
apps/sheets/src/renderer/ai/AiChatPanel.tsx
apps/sheets/src/renderer/ai/transport.ts
apps/sheets/src/renderer/i18n/strings-ai.ts
apps/slides/src/renderer/ai/AiPanel.tsx
apps/slides/src/renderer/ai/transport.ts
apps/slides/src/renderer/i18n/strings-ai.ts
apps/pdf/src/renderer/ai/AiPanel.tsx
apps/pdf/src/renderer/ai/transport.ts
apps/pdf/src/renderer/i18n/locale.tsx              (only if getLang is needed)
apps/pdf/src/renderer/i18n/strings.ts
apps/markdown/src/renderer/ai/AiPanel.tsx
apps/markdown/src/renderer/ai/transport.ts
apps/markdown/src/renderer/i18n/locale.tsx          (only if getLang is needed)
apps/markdown/src/renderer/i18n/strings.ts

apps/sheets/src/renderer/i18n/codex-strings.ts       (delete)
apps/slides/src/renderer/i18n/codex-strings.ts       (delete)
apps/pdf/src/renderer/i18n/codex-strings.ts          (delete)
apps/markdown/src/renderer/i18n/codex-strings.ts     (delete)

package-lock.json
plans/README.md
```

Adding @genoffice/i18n to packages/ai-provider/package.json is in scope; the
five apps already depend on @genoffice/ai-provider, @genoffice/i18n, and
@genoffice/ui, so no app package dependency should be added unless
verification proves one is missing. Do not add a new package.

## Out of scope

- Any Electron main/preload IPC channel, auth storage, runtime orchestration,
  agent loop, or provider transport behavior from Plans 023–026.
- The existing chatbox title, model-picker alignment, send button layout, or
  CSS changes already present in the worktree.
- App-owned prompts, tools, document mutations, message rendering, and
  non-Codex provider copy.
- A generic cross-app localization framework or a new i18n abstraction beyond
  the existing defineStrings/createI18n primitives.
- Changing supported locales or changing the CodexReasoningEffort IDs.
- Reverting, squashing, or reformatting the pre-existing dirty worktree.

## Implementation sequence

### Step 1 — Add the shared catalog at the provider boundary

1. Re-read packages/ai-provider/src/types.ts, packages/i18n/src/index.ts, and
   the four existing codex-strings.ts files. Record any key or wording
   differences before deleting a copy.
2. Add @genoffice/i18n to packages/ai-provider/package.json and update
   package-lock.json with the repository's normal package-manager command.
3. Implement codex-i18n.ts with defineStrings, createI18n, the complete
   19-locale catalog, the exhaustive reasoning map, and the single error map.
4. Export the catalog, CodexUiLabels, factory, and resolver from
   packages/ai-provider/src/index.ts.
5. Add packages/ai-provider/tests/codex-i18n.test.ts before app migration.
   It must assert:

   - every LANGS entry returns every catalog key with a non-empty value;
   - every CodexReasoningEffort maps to a non-empty localized label for every
     locale;
   - English reasoning values exactly match the approved title-case labels;
   - all known IPC codes and an unknown string resolve to non-empty copy;
   - auth-required, auth-expired, and auth-temporary remain distinct.

Verification:

```sh
npm run typecheck -w @genoffice/ai-provider
npm run test -w @genoffice/ai-provider -- codex-i18n.test.ts
```

Expected result: both commands exit 0; the test enumerates all 19 locales,
all eight effort IDs, and all current normalized error codes.

### Step 2 — Make shared UI consume the provider contract

1. Change AiProviderControlLabels in packages/ui/src/AiProviderControls.tsx
   to reuse the exported CodexUiLabels contract. Keep the existing UI type
   export for compatibility unless the compiler proves it is unused.
2. Do not move locale lookup into React or Electron. The UI receives already
   localized labels, exactly as it does today; only the source of those labels
   changes.
3. Verify AiCodexModelControl, AiProviderAuthBanner, and ProviderSelect still
   render the same control states and accessibility text.

Verification:

```sh
npm run typecheck -w @genoffice/ui
npm run test -w @genoffice/ui
```

Expected result: no consumer needs a local reasoningLabel implementation, and
all existing UI tests pass.

### Step 3 — Migrate app catalogs and label consumers

1. In each app's renderer i18n dictionary, compose the provider-owned locale
   entries so existing typed t('aiCodex...') call sites remain valid while the
   values have one definition.
2. Delete the four duplicate codex-strings.ts files. Remove every inline Codex
   key definition from Docs' strings-ai.ts. Remove the Slides-only
   aiSwitchModel spelling and migrate its consumers to aiSwitchModelTitle.
3. Replace every app-local provider label helper and every inline Docs label
   object with one getCodexUiLabels(lang) result shared by all controls in that
   panel. This must fix the Docs auth-banner raw-ID path as part of the
   migration.
4. Keep the active locale reactive using the existing app useI18n state. Do not
   cache a label bundle across locale changes.

Verification:

```sh
npm run typecheck -w @genoffice/docs
npm run typecheck -w @genoffice/sheets
npm run typecheck -w @genoffice/slides
npm run typecheck -w @genoffice/pdf
npm run typecheck -w @genoffice/markdown
npm test --workspaces --if-present
```

Expected result: all five apps compile and their existing tests pass; changing
the locale changes reasoning, control, and Codex chrome labels together.

### Step 4 — Migrate transport and direct error consumers

1. Replace each app's local codexErrorText implementation with the shared
   resolver while keeping the current transport function signatures.
2. Ensure the resolver reads the current app locale at call time. Add the
   minimal getLang export to PDF/Markdown locale modules only if their
   transport cannot otherwise receive the active lang without stale state.
3. Migrate Slides' direct codexErrorText stream-error call to the same shared
   resolver. Remove any now-unused local translation imports and helpers.
4. Confirm error messages remain localized for all five apps, including
   timeout/credits and unknown errors, not only the errors present in the
   copied Codex catalogs.

Verification:

```sh
rg -n "codexErrorText|resolveIpcErrorCode\\(" \
  apps/docs/src/renderer/ai apps/sheets/src/renderer/ai \
  apps/slides/src/renderer/ai apps/pdf/src/renderer/ai \
  apps/markdown/src/renderer/ai
rg -n "reasoningLabel:|reasoningLabel\\(" \
  apps/docs/src/renderer/ai apps/sheets/src/renderer/ai \
  apps/slides/src/renderer/ai apps/pdf/src/renderer/ai \
  apps/markdown/src/renderer/ai
rg --files apps | rg 'renderer/i18n/codex-strings\\.ts'
```

Expected result: the first two searches return no app-owned mapping/helper
definitions (shared imports and UI consumption are allowed); the last search
returns no files. If a search finds a legitimate unrelated use, narrow the
check and document it rather than retaining a duplicate Codex map.

### Step 5 — Run the suite gates and review the diff

Run from the repository root:

```sh
npm run typecheck
npm test
npm run lint
npm run format:check
npm run check:theme-colors
npm run build:all
npm run licenses
npm audit --omit=dev --audit-level=high
git diff --check
```

Expected result: all commands exit 0. Lint warnings are acceptable only when
they are pre-existing and clearly identified; new errors are not.

Then review:

```sh
git diff --stat
git diff -- packages/ai-provider packages/ui apps/docs apps/sheets apps/slides apps/pdf apps/markdown package-lock.json
git status --short
```

Confirm the diff contains only the in-scope implementation plus the
pre-existing dirty files listed before execution. Do not use git reset,
git checkout, or broad cleanup to make the worktree appear clean.

## Test matrix

The implementation is complete only when these invariants are covered:

| Invariant                                             | Test/check                                                     |
| ----------------------------------------------------- | -------------------------------------------------------------- |
| Every supported locale has the same Codex key set     | codex-i18n.test.ts over LANGS and defineStrings typecheck      |
| Every reasoning ID has a localized display value      | codex-i18n.test.ts over every CodexReasoningEffort and LANGS   |
| English uses the approved title-case labels           | exact assertions for all eight values                          |
| Error mapping is centralized and complete             | shared resolver test over all IpcErrorCode values plus unknown |
| Shared controls accept only the common label contract | @genoffice/ui typecheck/tests                                  |
| Five apps use the same locale-aware labels            | five app typechecks and existing AI/control tests              |
| No copied catalog or transport map remains            | the rg checks in Step 4                                        |
| Existing app behavior remains intact                  | workspace tests, lint, build, and manual diff review           |

## Done criteria

- @genoffice/ai-provider owns the only Codex catalog and the only reasoning
  effort display mapping.
- All 19 locales have non-empty Codex chrome, reasoning, and error values.
- xhigh renders as localized “Extra High” rather than xhigh, Xhigh, or a
  per-app variant.
- Docs, Sheets, Slides, PDF, and Markdown obtain labels from the shared
  provider factory; no app defines reasoningLabel or a Codex error map.
- The four copied codex-strings.ts files are deleted and Docs no longer embeds
  a second Codex catalog.
- All current IPC error codes retain their distinct semantics and localized
  fallback behavior.
- Package metadata and package-lock.json are synchronized.
- Every command in Steps 1–5 passes, and git diff --check is clean.
- plans/README.md contains Plan 027 with status DONE only after the
  implementation and review are complete. It remains TODO while this plan is
  only being written.

## STOP conditions

Stop before source edits and report the discrepancy if:

- any listed type, file, or shared UI contract has materially changed;
- adding the provider catalog would create a dependency cycle;
- an app dictionary cannot compose the provider catalog without weakening its
  typed key guarantees;
- an existing locale lacks a defensible localized value for a new reasoning or
  error key and no language decision is available;
- an app-specific use of a shared Codex key has a meaning that cannot be
  reconciled without changing product behavior;
- the current error normalization contract no longer includes the listed IPC
  codes;
- tests or typechecks fail for an unrelated pre-existing reason and the
  failure cannot be isolated after one focused retry;
- the implementation would require touching auth, main/preload IPC, runtime,
  layout, or other out-of-scope files beyond the conditional locale accessors;
- a formatter or package-manager operation proposes changes outside the
  in-scope files and the pre-existing dirty set.

## Maintenance contract

Future Codex effort IDs must be added in three places together: the
CodexReasoningEffort union, the shared localized catalog, and the exhaustive
reasoning-map test. Future Codex error codes must be mapped in the provider
resolver once; app transports must call it and never add a switch or local
translation table. App-specific AI copy remains local unless it becomes part
of this same shared Codex contract.

## Execution record

- STATUS: COMPLETE WITH BASELINE TEST DEVIATION
- STEPS:
  - Step 1: done — provider typecheck passed; Codex localization suite passed
    with 11 files and 161 tests.
  - Step 2: done — shared UI typecheck passed. The UI package has no test
    script, so there was no additional UI test command to run.
  - Step 3: done — all five app typechecks passed; serial AI/control tests
    passed for Sheets (3), PDF plus i18n (60), and Markdown (2); Slides'
    full suite passed (43 files, 413 tests).
  - Step 4: done — all five transports use resolveCodexError; no app-owned
    reasoning formatter, error table, or copied catalog remains.
  - Step 5: done except for the documented baseline test failure — root
    typecheck, format check, lint, theme check, licenses, audit, and build all
    passed.
- STOPPED BECAUSE: none for Plan 027 implementation. The root npm test command
  stops at the pre-existing Docs assertion in
  apps/docs/tests/ai-panel-collapse.test.ts:185, which still expects
  GenSpark AI while the already-requested source behavior is Genspark. That
  test and title behavior were dirty before Plan 027 and are explicitly out
  of scope; no unrelated test or title edit was made.
- FILES CHANGED: provider Codex catalog, exports, dependency metadata, and
  tests; shared UI label contract; five app AI panels and transports; five app
  i18n compositions/accessors; four duplicate catalog deletions; lockfile and
  plan index.
- NOTES: The concurrent full app run also exposed Sheets Rust-sidecar and PDF
  font-test timeouts under the existing environment. Serial Plan-027-focused
  tests passed, and the independent Docs failure was reproduced directly.
