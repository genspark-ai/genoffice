# Plan 017: Prepare a focused, validated Docs Codex PR

> **Implementation instructions**: This is the final integration gate. Begin
> only after Plans 011-016 are DONE. Run every automated and live check, preserve
> redacted results in the PR description, then remove internal planning artifacts
> from the product diff. Do not push or open a PR unless the operator explicitly
> asks.
>
> **Drift check (run first)**:
> `git diff --stat 1878b30..HEAD -- apps/docs packages/ai-provider packages/ui plans .github/pull_request_template.md .github/workflows/ci.yml`
> Review all drift because this plan validates the complete branch.

## Status

- **Priority**: P1
- **Effort**: M
- **Risk**: MED
- **Depends on**: Plans 011, 012, 013, 014, 015, and 016
- **Category**: dx, tests, docs
- **Planned at**: commit `1878b30`, 2026-08-15

## Why this matters

The branch is a coherent vertical Docs integration, but it currently includes
internal plan documents, one changed file that fails formatting, and no recorded
live-account acceptance result. GenOffice's PR template requires a focused diff,
all CI gates, validation disclosure, and screenshots for visible changes. Finish
those gates here; keep the later visual redesign and office-wide rollout in
separate PRs.

## Current state

- There is no repository `CONTRIBUTING.md`. The applicable contribution rules
  are `.github/pull_request_template.md` and `.github/workflows/ci.yml`.
- The PR template requires `npm run format:check`, `npm run lint`,
  `npm run typecheck`, `npm test`, focused scope, i18n resources for user-facing
  strings, and before/after evidence for visible changes.
- CI additionally runs license checks, theme-token checks, deterministic fixture
  generation, Sheets compatibility, builds all apps, and Electron E2E.
- At commit `1878b30`,
  `FORMAT_BASE_REF=origin/main npm run format:check` reports
  `apps/docs/src/main/docs-main.ts` as unformatted.
- `plans/001-010` and this hardening series are internal implementation handoff
  artifacts. They do not exist on `origin/main` and should not ship in the
  product PR.
- The approved architecture is GenOffice-owned OAuth/transport feeding the
  existing GenOffice `AgentLoop` and Docs tools. It is modeled on established
  Codex/Pi behavior but does not depend on Pi or copy/read Pi credential files.
- This PR may retain the functional Docs provider/model/reasoning controls needed
  to use the integration. Do not add the later Impeccable redesign or roll the
  feature into other office apps here.

## Commands you will need

| Purpose      | Command                                                  | Expected on success                              |
| ------------ | -------------------------------------------------------- | ------------------------------------------------ |
| Format write | `FORMAT_BASE_REF=origin/main npm run format`             | exit 0; changed files formatted                  |
| Format check | `FORMAT_BASE_REF=origin/main npm run format:check`       | exit 0                                           |
| Theme tokens | `FORMAT_BASE_REF=origin/main npm run check:theme-colors` | exit 0                                           |
| Licenses     | `npm run licenses`                                       | exit 0                                           |
| Lint         | `npm run lint`                                           | exit 0                                           |
| Typecheck    | `npm run typecheck`                                      | exit 0                                           |
| Unit tests   | `npm test`                                               | exit 0                                           |
| Build        | `npm run build:all`                                      | exit 0                                           |
| E2E          | `npm run test:e2e`                                       | exit 0, or a documented environment-only blocker |

## Scope

**In scope**:

- All product files already changed by the Docs Codex integration and Plans 011-016
- Formatting changes limited to files already changed on this branch
- `plans/` status reconciliation and final removal from the product PR
- A PR description based on `.github/pull_request_template.md` (outside the repo
  unless the operator requests a draft file)
- Redacted screenshots/recording and live-account validation notes

**Out of scope**:

- New provider/model selector visual design
- Sheets, Slides, PDF, Markdown, or shared cross-app rollout
- Pi migration, Pi dependency, Codex CLI runtime dependency, or auth-contract changes
- Unrelated refactors, formatting, fixture updates, or flaky-test fixes
- Rewriting git history, pushing, or opening a PR without explicit operator approval

## Git workflow

- Stay on the operator-provided branch.
- Do not rewrite history. If the operator wants separate commits, create new
  forward commits only.
- Do not push or open a PR. Hand the validated branch and PR text back to the operator.

## Steps

### Step 1: Confirm all prerequisite plans and scope

Read Plans 011-016 and verify every done criterion against the live code and
test output. Update their README statuses only when evidence exists. Inspect the
branch file list and ensure every product file supports the Docs Codex vertical
integration.

This is one focused product PR: provider transport/auth + required Docs IPC/UI.
Do not artificially split mutually dependent layers into non-runnable PRs. The
next PRs are (1) Impeccable-led UI redesign and (2) office-wide rollout.

**Verify**:

- `git diff --name-only origin/main...HEAD` → every non-`plans/` file is part of
  the Docs Codex integration or a required shared provider/composer change.
- `rg -n "\| 01[1-6].*DONE" plans/README.md` → six matches.

If an unrelated product file is present, report it to the operator; do not delete
or revert user work without approval.

### Step 2: Format only the branch's changed files

Run the repository formatter with `FORMAT_BASE_REF=origin/main`. Inspect the
result and confirm it touched only files already changed by the integration.

**Verify**:

- `FORMAT_BASE_REF=origin/main npm run format` → exit 0.
- `FORMAT_BASE_REF=origin/main npm run format:check` → exit 0.
- `git diff --check` → no whitespace errors.
- `git diff --name-only` → no unrelated files appeared.

### Step 3: Run the full local CI-equivalent gates

Run, in this order:

1. `npm run licenses`
2. `FORMAT_BASE_REF=origin/main npm run check:theme-colors`
3. `npm run lint`
4. `npm run typecheck`
5. `npm test`
6. `npm run fixtures`
7. `git diff --exit-code -- fixtures/generated`
8. `npm run fixtures -w @genoffice/sheets`
9. `npm run compat -w @genoffice/sheets`
10. `npm run build:all`
11. `npm run test:e2e`

Record exact command, pass/fail, date, OS, and any environment-only limitation.
Do not fix unrelated failures inside this plan. If fixtures change, STOP; do not
commit generated drift unrelated to Codex.

**Verify**: every command exits 0. E2E may be marked blocked only when the local
machine lacks a CI prerequisite and all other gates pass; preserve the exact
non-secret failure for the operator and rely on GitHub CI before merge.

### Step 4: Complete redacted live-account acceptance

With a maintainer-approved ChatGPT account, run the built Docs app and record
only pass/fail plus non-sensitive diagnostics. Verify:

1. select Codex, sign in, and load the account-specific model list;
2. restart the app and confirm protected credential reuse without another login;
3. generate a text-only answer;
4. complete one Docs tool mutation and confirm only caller-supplied Docs/file
   tools were offered/executed;
5. run a multi-turn continuation;
6. cancel an active generation and return to an idle composer;
7. simulate or safely induce a transient refresh failure and confirm credentials
   remain; confirm a mocked/controlled invalid refresh returns to signed-out state;
8. sign out and confirm the credential file is removed;
9. on Linux, confirm `basic_text` produces the safe unavailable-storage message;
10. confirm renderer state, settings JSON, IPC chunks, logs, screenshots, and
    error text contain no access token, refresh token, account ID, authorization
    header, raw provider body, or credential path content.

Never record the account email, token, prompt, generated document content, or
raw OAuth/provider payload in plans, logs, screenshots, or PR text.

**Verify**: a redacted checklist contains ten pass results, or a precise blocker
is reported. Any failed security/auth item is a STOP condition.

### Step 5: Prepare UI evidence and PR text

Capture redacted before/after screenshots or a short recording showing the Docs
provider selector, sign-in state, and model/reasoning selector. Use a blank or
synthetic document and hide account-identifying data.

Draft the PR body from `.github/pull_request_template.md`:

- summarize the GenOffice-owned architecture and why it is needed;
- state that maintainer approval exists for the auth contract;
- state that Pi/Codex CLI is not a runtime dependency and no external credential
  file is copied/read;
- list every validation command and live result;
- disclose any check not run and why;
- attach the redacted UI evidence;
- explicitly defer visual redesign and cross-app rollout.

**Verify**: every PR-template checkbox has evidence or an explicit explanation.

### Step 6: Remove internal planning artifacts from the product diff

After all implementation agents have finished using them and the PR text contains
the validation record, mark Plan 017 DONE in `plans/README.md`, then remove the
branch-added `plans/` directory from the product PR. Confirm `origin/main` has no
tracked `plans/` files before removal. This step intentionally removes this plan
itself; the PR body becomes the durable validation record.

Do not remove any plan directory that appears on the base branch or contains
unrelated user files. If either is true, STOP and ask the operator how to archive
the handoff documents.

**Verify**:

- `git ls-tree -r origin/main --name-only | rg '^plans/'` → no output.
- `git diff --name-only origin/main...HEAD | rg '^plans/'` → no output after removal.
- `git status --short` → only intended product changes and removal of branch-added
  plan files remain.

## Test plan

- Automated: exact local equivalents of every PR/CI gate above.
- Live: ten-item redacted acceptance checklist across auth, capability loading,
  streaming, tool execution, cancellation, refresh classification, logout, Linux
  storage, and secret boundaries.
- Visual: redacted screenshots/recording for the existing functional UI only.

## Done criteria

- [ ] Plans 011-016 are verified DONE before final integration testing.
- [ ] Format, theme, license, lint, typecheck, unit, fixture, compatibility, and
      build gates exit 0.
- [ ] E2E exits 0 or has a documented environment-only blocker pending CI.
- [ ] All ten live-account checks pass with no sensitive data recorded.
- [ ] The final diff contains no unrelated changes or `plans/` artifacts.
- [ ] PR text follows the repository template and identifies deferred work.
- [ ] No Pi/Codex CLI runtime dependency or renderer credential path exists.
- [ ] The operator receives the validated branch state and PR text; nothing was pushed.

## STOP conditions

- Any prerequisite plan is not DONE or its done criteria cannot be reproduced.
- A token, account identifier, raw provider payload, or credential material is
  visible outside the main-process auth boundary.
- Live auth/tool behavior fails, requires weakening the Docs tool boundary, or
  requires an undocumented workaround.
- Formatting/fixtures introduce unrelated diffs.
- The base branch contains a `plans/` directory or the current one has unrelated user files.
- Any CI-equivalent gate fails twice after a reasonable in-scope correction.

## Maintenance notes

Keep this PR as the stable Docs vertical slice. The next work should use
Impeccable to redesign the provider/model UX against GenOffice conventions, then
reuse the settled provider/auth boundary across office apps. A Pi migration is a
separate measured spike, justified only if maintaining protocol/auth compatibility
cost becomes materially higher than adopting Pi's provider layer.
