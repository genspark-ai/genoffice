# Plan 009: Refine Docs provider and composer controls

> **Implementation instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise.
>
> **Drift check (run first)**: `git diff --stat 4da673d..HEAD -- apps/docs/src/renderer/ai/AiPanel.tsx apps/docs/src/renderer/styles.css apps/docs/src/renderer/i18n/strings-ai.ts apps/docs/tests/ai-panel-collapse.test.ts`
> The working tree already contains the Codex provider implementation. This
> plan is scoped to the UI pass on top of that state; compare the current
> excerpts below before editing.

## Status

- **Priority**: P1
- **Effort**: M
- **Risk**: MED
- **Depends on**: 001–008 are complete for the provider/model capability path
- **Category**: tech-debt
- **Planned at**: commit `4da673d`, 2026-08-05

## Why this matters

The Docs AI panel exposes provider, model, reasoning, and track-changes state in
one crowded header/footer treatment. The selected provider also replaces the
product title, so switching to ChatGPT Codex changes the panel identity. This
pass makes provider choice a stable header control, keeps model/reasoning close
to Send where users act on them, and gives Track changes an explanatory choice
surface instead of a state-changing pill.

## Current state

- `apps/docs/src/renderer/ai/AiPanel.tsx` owns provider settings, Codex
  capability loading, track-changes persistence, and the composer.
- `apps/docs/src/renderer/ai/AiPanel.tsx:852-919` currently renders the
  selected provider label as the title, then provider/model/reasoning selects
  together in the header action group.
- `apps/docs/src/renderer/ai/AiPanel.tsx:782-788` toggles track changes directly
  and accepts pending AI revisions when switching off.
- `apps/docs/src/renderer/ai/AiPanel.tsx:1104-1141` passes attachment and track
  controls to `AiComposer` through `footerStart`. `AiComposer` already accepts
  arbitrary React nodes there, so `packages/ui/src/AiComposer.tsx` does not need
  to change.
- `apps/docs/src/renderer/styles.css:3945-3990` styles the panel header and
  actions; there is no dedicated `.ai-provider-select` style, so the browser
  default select chrome is currently visible.
- `apps/docs/src/renderer/styles.css:4922-4995` styles the composer footer and
  Track changes as a bordered pill.
- `apps/docs/tests/ai-panel-collapse.test.ts:80-183` mounts the real panel in
  jsdom and already verifies Codex capability options/settings updates. The
  existing test expects three `.ai-provider-select` elements for Codex and must
  be rewritten around the new provider header select plus composer popover.
- `apps/docs/src/renderer/i18n/strings-ai.ts` requires every language dictionary
  to match the Chinese key set. Reuse existing localized model/reasoning and
  Track changes title strings for popup labels/descriptions; update the
  existing `aiPanelTitle` brand value to `GenSpark AI` in every dictionary.

## Commands you will need

| Purpose       | Command                                                        | Expected on success                                                                     |
| ------------- | -------------------------------------------------------------- | --------------------------------------------------------------------------------------- |
| Drift check   | `git diff --stat 4da673d..HEAD -- <in-scope paths>`            | No committed drift relevant to this plan; compare any output with the current excerpts. |
| Red test      | `npm run test -w @genoffice/docs -- ai-panel-collapse.test.ts` | Fails on the newly added UI behavior before production changes.                         |
| Focused tests | `npm run test -w @genoffice/docs -- ai-panel-collapse.test.ts` | All tests in the file pass.                                                             |
| Typecheck     | `npm run typecheck -w @genoffice/docs`                         | Exit 0, no TypeScript errors.                                                           |
| Lint          | `npm run lint -- --quiet`                                      | Exit 0, no ESLint errors.                                                               |
| Format check  | `npm run format:check`                                         | Exit 0, no formatting differences.                                                      |

## Scope

**In scope** (the only files this plan may modify):

- `apps/docs/src/renderer/ai/AiPanel.tsx`
- `apps/docs/src/renderer/styles.css`
- `apps/docs/src/renderer/i18n/strings-ai.ts`
- `apps/docs/tests/ai-panel-collapse.test.ts`

**Out of scope**:

- `packages/ui/src/AiComposer.tsx`; its existing `footerStart` extension point is
  sufficient.
- Provider contracts, OAuth, IPC, capability fetching, AgentLoop behavior, and
  document revision semantics.
- Sheets, Slides, PDF, ribbon Track Changes, and any global theme rewrite.
- Git commit, push, merge, or PR creation.

## Design constraints

- Keep `AI_PROVIDERS` as the provider selector source; all existing providers
  remain available.
- Header visible title is the constant product label `GenSpark AI`, regardless
  of selected provider.
- Provider select is last in the header action group, aligned to the far right,
  with transparent/text-only styling and no visible border or outline.
- Codex model/reasoning control appears only when `settings.provider` is
  `openai-codex` and capability data has a selected model. One compact trigger
  sits immediately before Send. Its popover lists available model choices and
  the current model's available reasoning efforts, using existing
  `changeCodexModel` and `changeCodexReasoning` callbacks.
- Track changes remains persisted under `TRACK_CHANGES_KEY`. Its trigger opens a
  popover with two clearly selectable choices using localized
  `aiTrackOnTitle`/`aiTrackOffTitle` explanations. Selecting off preserves the
  current behavior of accepting pending AI revisions. Clicking outside or
  pressing Escape closes either popover without changing state.
- Use real buttons and ARIA state (`aria-expanded`, `aria-haspopup`, menu/menuitem
  or equivalent) so keyboard users can discover the controls. Do not introduce
  native `<select>` controls for model/reasoning inside the composer.

## Steps

### Step 1: Add failing behavior tests

In `apps/docs/tests/ai-panel-collapse.test.ts`, add tests before production
changes for these behaviors:

1. Codex panel title remains `GenSpark AI`; header has one provider select only;
   model and reasoning choices are not rendered as header selects.
2. Codex composer trigger opens one popover containing the capability-provided
   models and current model's reasoning efforts. Clicking a model and an effort
   calls `onSettingsChange` with the expected Codex settings.
3. Track changes trigger opens a popover with two explanatory options. Choosing
   on marks the trigger active; choosing off clears it and persists `0`.

Model the mount/cleanup and async capability setup on the existing tests in the
same file. Assert user-visible behavior and settings payloads, not React
implementation details. Use stable classes/roles/data attributes only where
needed for the test contract.

**Verify**: `npm run test -w @genoffice/docs -- ai-panel-collapse.test.ts` → the
new tests fail because the current header still has model/reasoning selects and
the current Track changes button has no popover. Fix test setup errors until the
failure is specifically about the missing behavior; do not edit production code
before observing this failure.

### Step 2: Implement stateful popover controls in `AiPanel`

In `apps/docs/src/renderer/ai/AiPanel.tsx`:

- Add one state value for the active composer popover (`model` or `track`), plus
  refs/effect cleanup for outside pointer events and Escape. Close the popover
  when switching away from Codex.
- Keep existing Codex capability loading and settings callbacks unchanged.
  Derive the active model from `codexCapabilities` and current settings; do not
  invent model IDs or reasoning levels when capabilities are unavailable.
- Replace `{selectedProvider.label}` in the panel title with constant `GenSpark AI`.
- Remove Codex model/reasoning selects and loading/error text from the header.
  Keep the provider selector in the header, after new-chat/collapse actions, and
  preserve all `AI_PROVIDERS` options and `changeProvider` behavior.
- Replace direct Track changes toggling with a trigger/popover. Both choices
  call one explicit setter; the off path invokes the existing `acceptChanges`
  behavior. Do not alter revision application logic.
- Add a compact Codex model/reasoning trigger to `footerStart`, immediately
  before Send. Render model and effort option buttons in its popover. Keep the
  attachment control and Track changes control before it. Add a flex spacer for
  non-Codex providers so Send stays right-aligned when no model control exists.

**Verify**: `npm run test -w @genoffice/docs -- ai-panel-collapse.test.ts` → all
new and existing panel tests pass.

### Step 3: Match the reference interaction with scoped CSS

In `apps/docs/src/renderer/styles.css`:

- Style `.ai-provider-select` as a borderless, transparent, text-only control
  with a compact chevron/appearance treatment, focus-visible accessibility,
  and no outline in its resting state. Keep it last/right in the header action
  group and prevent long provider names from breaking the title.
- Add popover card styles for model/reasoning and Track changes: absolute
  positioning above the composer, panel surface, border/shadow, rounded corners,
  readable row spacing, hover/selected states, and wrapping descriptions.
- Replace the current bordered `.ai-track-btn` pill styling with a lightweight
  text control; active state uses the existing brand color without a chip border.
- Style the compact model trigger as a quiet footer control immediately before
  Send. Ensure the composer remains usable at the existing 280px minimum panel
  width and long localized labels wrap inside popovers rather than shifting Send.

**Verify**: `npm run format:check` → exit 0. `npm run typecheck -w
@genoffice/docs` → exit 0.

### Step 4: Normalize the visible product label

In `apps/docs/src/renderer/i18n/strings-ai.ts`, change every existing
`aiPanelTitle` value from `Genspark` to `GenSpark AI` so accessibility labels and
future uses agree with the fixed header product label. Do not add untranslated
keys or alter unrelated AI copy.

**Verify**: `rg -n "aiPanelTitle:" apps/docs/src/renderer/i18n/strings-ai.ts` →
all 19 values are `GenSpark AI`; `npm run typecheck -w @genoffice/docs` exits 0.

### Step 5: Review and full verification

Read the complete diff. Confirm every hunk maps to a step above, especially:

- provider selector remains the only header select;
- Codex capability choices still update `AiSettings` and model changes preserve
  a valid reasoning effort;
- popovers close on outside click/Escape and do not submit the composer;
- switching Track changes off still accepts pending AI revisions;
- no provider transport or document revision code changed.

Run focused and repository checks.

**Verify**:

- `npm run test -w @genoffice/docs -- ai-panel-collapse.test.ts` → all pass.
- `npm run typecheck -w @genoffice/docs` → exit 0.
- `npm run lint -- --quiet` → exit 0.
- `npm run format:check` → exit 0.
- `git diff --check` → no whitespace errors.
- `git status --short` → only the four in-scope files are modified by this plan
  (pre-existing user changes in those files are allowed; no out-of-scope file
  may be added).

## Test plan

- Extend `apps/docs/tests/ai-panel-collapse.test.ts`, following its existing
  `mount`, `createEditor`, `panelProps`, async capability, and cleanup patterns.
- Cover provider/title placement contract, Codex model/effort selection, Track
  changes two-choice behavior, persistence, and popover dismissal.
- Run the red test before editing production code, then rerun focused tests after
  each implementation step.

## Done criteria

- [ ] `GenSpark AI` stays visible in panel header for every provider.
- [ ] Provider selector is the only header select, appears at far-right end, and
      has no visible border/outline in its resting style.
- [ ] Codex model/reasoning control is inside composer immediately before Send;
      header no longer renders those controls.
- [ ] Model and reasoning selections update existing settings callbacks using
      capability data.
- [ ] Track changes opens an explanatory two-choice popover; off path accepts
      pending AI revisions and persists correctly.
- [ ] Focused Docs tests, Docs typecheck, lint, format check, and `git diff --check`
      pass.
- [ ] No out-of-scope file is modified.

## Execution record

STATUS: COMPLETE

- **Steps**: Step 1 RED observed; Steps 1–4 implemented; Step 5 reviewed.
- **Focused verification**: Docs panel test file passed (7/7).
- **Full Docs verification**: `npm run test -w @genoffice/docs` passed (55 files,
  457 tests); Docs typecheck, production build, format check, and
  `git diff --check` passed.
- **Scoped lint**: ESLint passed for the touched TypeScript files.
- **Deviation**: Root `npm run lint -- --quiet` remains red on pre-existing,
  out-of-scope Codex edits at `apps/docs/src/main/docs-main.ts:2337`
  (`prefer-const`) and `apps/docs/src/shared/ipc.ts:27`
  (`CodexReasoningEffort` unused). UI changes did not cause either error; no
  out-of-scope fixes were made.
- **Files changed by this plan**: `apps/docs/src/renderer/ai/AiPanel.tsx`,
  `apps/docs/src/renderer/styles.css`, `apps/docs/src/renderer/i18n/strings-ai.ts`,
  and `apps/docs/tests/ai-panel-collapse.test.ts`. Existing provider changes in
  other files were preserved.

## STOP conditions

Stop and report if:

- `AiComposer` cannot place a footer control before Send without changing
  `packages/ui/src/AiComposer.tsx`.
- Codex capabilities are unavailable in the test/runtime path and the model
  control would require inventing fallback options.
- Current code excerpts differ in a way that changes provider/settings ownership
  or revision semantics.
- A required visual change needs global theme or non-Docs app edits.
- Any verification command fails twice after a focused fix attempt.

## Maintenance notes

- Future providers with model/reasoning metadata should be added to the same
  composer control contract rather than reintroducing header selects.
- Keep the popover's option labels tied to capability data; never turn the model
  list into a free-form renderer-authored string.
- Review Track changes off-path carefully: it intentionally accepts existing AI
  revisions, while ribbon Track Changes remains separate and out of scope.
