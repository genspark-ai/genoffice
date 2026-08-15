# Plan 022: Validate and package the focused Docs Codex PR

> **Implementation instructions**: This is the final gate after Plans 019-021.
> Run every command and live scenario. Do not mark DONE or remove handoff plans
> until all gates pass. Stop and report on any STOP condition.
>
> **Drift check (run first)**:
> `git diff --stat 1878b30..HEAD -- apps/docs packages/ai-provider packages/agent-core plans`
> Review all drift: this plan validates the final branch, not one symbol. The
> audited working tree was uncommitted at the planned-at SHA.

## Status

- **Priority**: P1
- **Effort**: M
- **Risk**: LOW
- **Depends on**: Plans 019, 020, and 021
- **Category**: tests
- **Planned at**: commit `1878b30`, 2026-08-15

## Why this matters

Automated checks are green on the audited tree, but a real account flow and
final PR hygiene still have to prove the approved contract. This gate keeps the
product PR to one Docs vertical slice, excludes internal plans, and confirms
that shared seams are ready for later apps without rolling them out now.

## Current state

- Audit on 2026-08-15: full tests, typecheck, format, lint (zero errors), working
  tree theme check, production dependency audit, and `git diff --check` passed.
  `npm run build:all` also passed/was run; repeat it after Plans 019-021.
- `FORMAT_BASE_REF=origin/main npm run check:theme-colors` initially saw colors
  in committed branch state already corrected in the working tree. The final
  committed diff must pass against `origin/main`.
- `origin/main` does not contain `plans/`; branch-only plans must not ship.
- Approved storage: `~/.genoffice/codex-auth.json`, parent `0700`, file `0600`,
  fresh OAuth, and no Pi/Codex CLI credential import.
- `CONTRIBUTING.md` lists the exact gates and focused-PR expectations.

## Commands you will need

| Purpose  | Command                                                              | Expected on success        |
| -------- | -------------------------------------------------------------------- | -------------------------- |
| Format   | `npm run format:check && npm run format:check -- --base origin/main` | exit 0                     |
| Theme    | `FORMAT_BASE_REF=origin/main npm run check:theme-colors`             | exit 0                     |
| Static   | `npm run lint && npm run typecheck`                                  | exit 0; lint has 0 errors  |
| Tests    | `npm test`                                                           | exit 0                     |
| Licenses | `npm run licenses`                                                   | exit 0                     |
| Build    | `npm run build:all`                                                  | exit 0; all six apps build |

## Scope

**In scope**:

- Test-only fixes in files changed by Plans 019-021
- `plans/README.md` status before cleanup
- Deleting branch-only `plans/` after all gates and after the operator preserves
  the handoff elsewhere
- Preparing PR description text; do not publish

**Out of scope**:

- Provider/model selector redesign
- Other office-app rollout
- Pi migration or credential import
- New OAuth/storage architecture
- Push/PR publication without explicit instruction

## Git workflow

- Work on the current branch; do not push/open a PR.
- Suggested final commit subject: `add ChatGPT Codex login to Docs`.

## Steps

### Step 1: Run complete repository gates

```sh
npm run format:check
npm run format:check -- --base origin/main
FORMAT_BASE_REF=origin/main npm run check:theme-colors
npm run lint
npm run typecheck
npm test
npm run licenses
npm run build:all
npm audit --omit=dev --audit-level=high
git diff --check
```

Expected: every command exits 0; lint may show existing warnings but no errors;
audit reports no high/critical production vulnerabilities.

### Step 2: Exercise Docs with a non-privileged test account

Verify manually:

1. Without `codex-auth.json`, Codex shows localized signed-out state and send is disabled.
2. Login opens OpenAI auth only after port 1455 listens; cancel/retry recover.
3. Successful login writes only the approved JSON path. Inspect modes, never contents.
4. Restart Docs: account, capabilities, model, and reasoning restore.
5. Run read-only, edit/tool, cancel, and follow-up prompts successfully.
6. Disconnect network: signed-in identity remains with localized transient guidance; file remains.
7. Logout removes the file, clears account UI, and disables sending.
8. During login, open a second window: overlap does not strand either callback.

Record pass/fail and platform in PR notes without secrets, IDs, authorization
URLs, or raw payloads.

Operator report: scenarios 1–8 PASS. Platform was not recorded in this handoff.

### Step 3: Review final architecture and diff

```sh
rg -n "codex-auth.json|beginCodexCallback|codexCredentialStore" apps packages
rg -n "from 'electron'|from \"electron\"" packages/ai-provider
rg -n "refreshToken|accessToken" apps/docs/src/renderer apps/docs/src/preload
git diff --stat origin/main...HEAD
git diff --name-only origin/main...HEAD
```

Expected: one shared Node adapter owns store/callback; no Electron import in the
provider package; no credential crosses renderer/preload; only the Docs slice
and shared seams changed, with no speculative other-app wiring.

### Step 4: Remove plans and prepare PR summary

After Steps 1-3 pass and the operator preserves this handoff, remove branch-only
`plans/` files from the product diff. If `origin/main` begins tracking them,
STOP instead.

Prepare but do not publish a PR description covering:

- Why: maintainer-approved ChatGPT Codex login/provider support in Docs.
- Architecture: shared provider/auth core and Node adapter; Docs Electron/UI.
- Scope: Docs only; redesign and office rollout follow separately.
- Security: PKCE/state, loopback callback, safe codes, no renderer credentials,
  atomic `0600` JSON.
- Validation: exact automated commands and live scenarios/platform.

**Verify**:

```sh
git diff --name-only origin/main...HEAD | rg '^plans/'
git status --short
```

Expected: first command has no matches; status shows only intended product work
or is clean after an operator-created commit.

## Test plan

- Full suite and all six builds.
- Fresh/persisted login, transient refresh, cancellation, overlap, stream/tool,
  follow-up, and logout with a real account.
- Final security/diff searches and plan-artifact exclusion.

## Done criteria

- [ ] Plans 019-021 are DONE and reviewed.
- [ ] Every automated gate passes on the final diff.
- [x] Every live scenario passes and is recorded safely.
- [ ] JSON modes/path match the approved convention.
- [ ] No plans or other-app rollout remain in product diff.
- [ ] Focused PR description is ready but unpublished.

## STOP conditions

- Live OAuth/provider validation is unavailable or fails twice.
- `origin/main` materially changes gates or begins tracking `plans/`.
- A credential/token/account ID appears in IPC, logs, output, or PR text.
- Fixing a gate requires redesign, other-app rollout, Pi, or contract changes.

## Maintenance notes

Next: use Impeccable for provider/model UI as a separate change, then integrate
the shared adapter/provider into each app's existing IPC/tool path without
copying Docs auth code.
