# Plan 028: Prune brittle Codex presentation assertions

> **Implementation instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving on. Do
> not commit, push, or open a PR. When complete, update this plan's status row
> in `plans/README.md` after implementation and review.
>
> **Drift check (run first)**: `git diff --stat 3723808..HEAD -- apps/docs/tests/ai-panel-collapse.test.ts apps/sheets/tests/ai-provider-controls.test.ts packages/ai-provider/tests/codex-i18n.test.ts packages/ai-provider/tests/providers.test.ts plans/README.md`
> If any in-scope test has changed, compare the current state below with the
> live file. Stop on a material mismatch rather than broadening this cleanup.

## Status

- **Priority**: P1
- **Effort**: S
- **Risk**: LOW
- **Depends on**: none
- **Category**: tests
- **Planned at**: commit `3723808`, 2026-08-16

## Why this matters

The new Codex rollout includes several renderer tests that assert product
copy, capitalization, sibling order, and incidental menu counts. They do not
protect an interaction contract and already cause the full test suite to fail:
the Docs test expects `GenSpark AI` while the product now renders `Genspark`.

Delete assertions whose only purpose is to freeze presentation or localized
copy. Preserve tests that protect authentication gating, settings mutation,
draft preservation, retry, and the account-backed provider defaults. This is
a test-only reduction; do not add a replacement visual-testing framework or
new component abstractions.

## Current state

- `apps/docs/tests/ai-panel-collapse.test.ts` exercises the Docs `AiPanel`
  through its public DOM controls. It also currently contains Codex tests with
  presentational assertions.
- `apps/sheets/tests/ai-provider-controls.test.ts` verifies the shared control
  integration and provider selection callback.
- `packages/ai-provider/tests/codex-i18n.test.ts` correctly validates that
  every supported locale has non-empty Codex copy; its exact English-label
  snapshot is redundant.
- `packages/ai-provider/tests/providers.test.ts` validates default persisted
  provider settings. The Codex test should keep account-backed/default-model
  behavior without treating display copy as a provider-runtime contract.

The relevant current assertions are:

```ts
// apps/docs/tests/ai-panel-collapse.test.ts:98-124
it('orders provider, new conversation, and collapse controls', async () => {
  // ...
  expect(provider?.nextElementSibling).toBe(headerButtons[0])
  expect(providerOptions).toEqual(['Genspark', 'ChatGPT Codex'])
})

// apps/docs/tests/ai-panel-collapse.test.ts:185-193
expect(container.querySelector('.ai-panel-title')?.textContent).toContain('GenSpark AI')
expect(container.querySelectorAll('.ai-panel-header .ai-provider-select')).toHaveLength(1)
expect(container.querySelectorAll('[data-codex-menu-item]')).toHaveLength(3)

// apps/docs/tests/ai-panel-collapse.test.ts:273-279
expect(header.nextElementSibling).toBe(banner)
expect(banner.textContent).toContain('登录 ChatGPT 后才能使用 ChatGPT Codex。')
expect(banner.querySelectorAll('button')).toHaveLength(1)
```

The already useful behavioral checks in that file include:

```ts
// apps/docs/tests/ai-panel-collapse.test.ts:195-229
expect(onSettingsChange).toHaveBeenLastCalledWith(/* selected model/effort/tier */)

// apps/docs/tests/ai-panel-collapse.test.ts:314-332
expect(aiCodexLogin).toHaveBeenCalledTimes(1)
expect(container.querySelector<HTMLTextAreaElement>('.ai-input-box textarea')!.value).toBe(
  'Keep this draft',
)
expect(container.querySelector<HTMLButtonElement>('.ai-send-btn')!.disabled).toBe(false)
```

The repository uses Vitest and renderer-level DOM tests. Match the existing
`act`, `mount`, and `cleanup` style; do not export test-only helpers from
production code. Root commands are `npm test`, `npm run typecheck`,
`npm run lint`, and `npm run format:check`.

## Commands you will need

| Purpose        | Command                                                                                                                                                                                                                                | Expected on success                                                                                      |
| -------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------- |
| Drift check    | `git diff --stat 3723808..HEAD -- apps/docs/tests/ai-panel-collapse.test.ts apps/sheets/tests/ai-provider-controls.test.ts packages/ai-provider/tests/codex-i18n.test.ts packages/ai-provider/tests/providers.test.ts plans/README.md` | No unexpected committed drift in scope                                                                   |
| Docs test      | `npm test -w @genoffice/docs -- tests/ai-panel-collapse.test.ts`                                                                                                                                                                       | All Docs panel tests pass, including the formerly failing Codex test                                     |
| Sheets test    | `npm test -w @genoffice/sheets -- tests/ai-provider-controls.test.ts`                                                                                                                                                                  | All selected Sheets tests pass                                                                           |
| Provider tests | `npm test -w @genoffice/ai-provider -- tests/codex-i18n.test.ts tests/providers.test.ts`                                                                                                                                               | Both selected provider test files pass                                                                   |
| Full tests     | `npm test`                                                                                                                                                                                                                             | Every workspace test passes                                                                              |
| Typecheck      | `npm run typecheck`                                                                                                                                                                                                                    | Exit 0, no TypeScript errors                                                                             |
| Lint           | `npm run lint`                                                                                                                                                                                                                         | Exit 0, no lint errors                                                                                   |
| Formatting     | `npm run format:check`                                                                                                                                                                                                                 | Exit 0 for committed project files; separately report any pre-existing untracked-tool formatting failure |
| Diff hygiene   | `git diff --check`                                                                                                                                                                                                                     | No whitespace errors                                                                                     |

## Scope

**In scope** (the only files to modify):

- `apps/docs/tests/ai-panel-collapse.test.ts`
- `apps/sheets/tests/ai-provider-controls.test.ts`
- `packages/ai-provider/tests/codex-i18n.test.ts`
- `packages/ai-provider/tests/providers.test.ts`
- `plans/README.md` (status row only)

**Out of scope** (do not touch):

- All production source, CSS, i18n catalogs, and display labels.
- Functional Codex tests in `packages/ai-provider/tests/auth.test.ts`,
  `codex-auth-node.test.ts`, `codex.test.ts`, and `main-runtime.test.ts`.
- Existing non-Codex panel tests, E2E setup, snapshots, or new visual-test
  tooling.
- The uncommitted anti-slop plugin configuration and `tools/oxlint/`.

## Git workflow

- Work on the operator-provided branch and preserve unrelated dirty changes.
- Do not commit, push, merge, or open a PR.

## Steps

### Step 1: Remove Docs-only presentation checks while retaining interaction contracts

In `apps/docs/tests/ai-panel-collapse.test.ts`:

1. Delete the entire `orders provider, new conversation, and collapse controls`
   test (current lines 98-129). It asserts only control order and display
   strings.
2. Delete the entire `keeps model and reasoning controls hidden for non-Codex
providers` test (current lines 131-139). Its sole assertion counts provider
   selects and does not test model/reasoning visibility.
3. Rename the test beginning at current line 141 to describe its retained
   behavior: selecting a Codex model, reasoning effort, and service tier
   updates settings. Remove title, header-placement, and option/menu-count
   assertions. Keep the trigger presence required to perform the interaction
   and keep the three `onSettingsChange` expectations.
4. In the preflight test, keep the signed-out banner's presence, disabled Send,
   and no-stream/no-capabilities checks. Delete sibling placement, translated
   banner copy, and button-count assertions.
5. In the retry test, keep draft preservation, retry availability, and the
   second login call. Delete the exact localized failure text assertion.
6. In the non-auth stream-error test, keep the error element's presence and
   absence of the login action. Delete the exact localized error text.

Do not alter mocks, user flows, or production selectors merely to simplify
these tests.

**Verify**: `npm test -w @genoffice/docs -- tests/ai-panel-collapse.test.ts`
→ all tests pass and no expectation mentions `GenSpark AI`, Chinese user copy,
`nextElementSibling`, or menu/option counts for the Codex controls.

### Step 2: Remove redundant provider display-copy expectations

1. In `apps/sheets/tests/ai-provider-controls.test.ts`, rename the first test
   to describe provider selection persistence. Remove the expected ordered
   option text array; retain dispatching `openai-codex`, the settings callback,
   and the no-write assertion.
2. In `packages/ai-provider/tests/codex-i18n.test.ts`, delete the entire
   `uses the approved English reasoning labels` test. The preceding every-locale
   non-empty check remains the localization contract.
3. In `packages/ai-provider/tests/providers.test.ts`, remove only
   `label: 'ChatGPT Codex'` from the Codex provider object assertion. Retain
   `defaultModel`, `requiresApiKey`, and the complete persisted default
   settings assertion.

**Verify**: `npm test -w @genoffice/sheets -- tests/ai-provider-controls.test.ts`
and `npm test -w @genoffice/ai-provider -- tests/codex-i18n.test.ts tests/providers.test.ts`
→ all selected tests pass.

### Step 3: Run gates and review the reduced test surface

Run every command in the table. Inspect the diff and confirm it removes only
test expectations or obsolete test bodies: no production source, generated
files, test harness changes, or anti-slop files may be changed. Update Plan
028's status in `plans/README.md` to `DONE` only after the full suite passes.

**Verify**: `npm test`, `npm run typecheck`, `npm run lint`, and
`git diff --check` all succeed; `git diff --name-only` is limited to the four
test files and the status-row edit (besides unrelated pre-existing changes).

## Test plan

- The retained Docs test must prove that each selected Codex model, reasoning
  effort, and service tier is passed to `onSettingsChange`.
- The retained signed-out/preflight test must prove send remains disabled and
  neither stream nor capabilities starts before authentication resolves.
- The retained login test must prove a draft survives login and duplicate login
  clicks do not issue a second request.
- The retained Sheets/provider tests must prove `openai-codex` is selectable,
  account-backed, and defaults to the expected runtime settings.
- Do not add tests for literal localized phrases, DOM sibling order, control
  order, headings, or number of menu entries.

## Done criteria

- [ ] The root `npm test` passes; the old Docs `GenSpark AI` failure is gone.
- [ ] Functional Codex authentication, selection, draft, and error-routing
      assertions remain in the four in-scope test files.
- [ ] No in-scope test asserts exact Codex product copy, localized message
      text, control sibling order, or incidental menu/option counts.
- [ ] `npm run typecheck`, `npm run lint`, and `git diff --check` succeed.
- [ ] No files outside scope are modified, apart from unrelated pre-existing
      working-tree changes.
- [ ] Plan 028 is marked `DONE` in `plans/README.md` only after review.

## STOP conditions

Stop and report instead of improvising if:

- Removing an assertion also removes the only test of a named functional
  contract in the test plan.
- A selector is needed for a user interaction but can only be removed by
  changing production source.
- The full suite fails outside the known Docs title assertion after the focused
  tests pass.
- Any requested cleanup requires changing locale content or Codex runtime
  defaults rather than tests.

## Maintenance notes

UI copy belongs in the shared localization catalog and should be reviewed as
product content, not frozen in renderer interaction tests. When a visual
layout is a product requirement, cover it with a deliberately owned E2E or
visual contract; do not reintroduce unit assertions about DOM sibling order.
