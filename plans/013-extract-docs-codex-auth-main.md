# Plan 013: Extract Docs Codex auth from `docs-main.ts`

> **Implementation instructions**: This is a behavior-preserving extraction.
> Follow the steps exactly, run each verification gate, and stop if behavior must
> change. Update Plan 013 in `plans/README.md` after review.
>
> **Drift check (run first)**:
> `git diff --stat 1878b30..HEAD -- apps/docs/src/main/docs-main.ts apps/docs/src/main/codex-auth-main.ts packages/ai-provider/src/index.ts`
> If `docs-main.ts` auth symbols no longer match the excerpts below, stop and report.

## Status

- **Priority**: P1
- **Effort**: S
- **Risk**: LOW
- **Depends on**: Plan 011
- **Category**: tech-debt
- **Planned at**: commit `1878b30`, 2026-08-15

## Why this matters

`docs-main.ts` currently owns document lifecycle, AI IPC, encrypted credential
storage, a localhost OAuth server, browser launch, and auth error mapping. The
Codex-specific block adds roughly 150 lines to an already large main-process
module and deep-imports another workspace's source. Extracting only the
Electron-specific auth wiring gives the security-sensitive code a testable home
without refactoring the broader AI IPC or inventing a shared cross-app layer.

## Current state

- `apps/docs/src/main/docs-main.ts:13-14` imports Node HTTP and Electron
  `safeStorage`; the Codex block is their only auth-specific consumer.
- `apps/docs/src/main/docs-main.ts:2487-2638` defines credential validation and
  storage, callback-server lifecycle, the `CodexAuthService` singleton, browser
  launch, and safe auth errors inline.
- `apps/docs/src/main/docs-main.ts:60` deep-imports
  `../../../../packages/ai-provider/src/auth`, bypassing the workspace export.
- `packages/ai-provider/src/index.ts:24-31` exports auth types but not
  `CodexAuthService`.
- `apps/slides/src/main/ai-ipc.ts` is the repository exemplar for extracting a
  focused main-process AI concern while keeping `registerAiIpc()` at the app boundary.
- The unified shell imports Docs' `registerAiIpc()`. Do not change channel
  ownership or register handlers twice.

## Commands you will need

| Purpose            | Command                                                                                                                | Expected on success |
| ------------------ | ---------------------------------------------------------------------------------------------------------------------- | ------------------- |
| Provider typecheck | `npm run typecheck -w @genoffice/ai-provider`                                                                          | exit 0              |
| Docs typecheck     | `npm run typecheck -w @genoffice/docs`                                                                                 | exit 0              |
| Focused tests      | `npm test -w @genoffice/ai-provider -- tests/auth.test.ts && npm test -w @genoffice/docs -- tests/codex-error.test.ts` | both exit 0         |
| Format             | `FORMAT_BASE_REF=origin/main npm run format:check`                                                                     | exit 0              |

## Scope

**In scope**:

- `apps/docs/src/main/docs-main.ts`
- `apps/docs/src/main/codex-auth-main.ts` (create)
- `packages/ai-provider/src/index.ts`
- `plans/README.md` for status only

**Out of scope**:

- AI IPC channel names or renderer/preload types
- Moving all of `registerAiIpc()` out of `docs-main.ts`
- Sharing this Electron module with Sheets, Slides, PDF, Markdown, or Shell
- Any OAuth, storage, callback, or user-visible behavior change
- Pi migration or a new package/dependency

## Git workflow

- Work on the operator-provided branch. This extraction should be one reviewable
  commit if the operator requests commits.
- Do not push or open a PR.

## Steps

### Step 1: Export the existing service through the package boundary

In `packages/ai-provider/src/index.ts`, export `CodexAuthService` from `./auth`
alongside the existing auth types. Do not expose private token parsing helpers or
constants.

Update `docs-main.ts` to import the class from `@genoffice/ai-provider`; remove
the relative deep import.

**Verify**: `npm run typecheck -w @genoffice/ai-provider && npm run typecheck -w @genoffice/docs` → exit 0.

### Step 2: Move only Electron-specific auth wiring

Create `apps/docs/src/main/codex-auth-main.ts` and move these existing symbols
without changing their behavior:

- credential path/constants and `isCodexCredentials()`;
- `codexCredentialStore()`;
- `beginCodexCallback()`;
- the lazy `CodexAuthService` singleton and `getCodexAuth()`;
- `safeCodexAuthError()`.

The new module may import `app`, `safeStorage`, and `shell` from Electron;
`createServer` and filesystem functions from Node; `safeExternalUrl` from
`@genoffice/electron-utils`; and auth types/service from
`@genoffice/ai-provider`. Use `join(app.getPath('userData'),
'codex-credentials.bin')` for the same current path.

Export only the functions consumed by `docs-main.ts` plus
`codexCredentialStore` and `beginCodexCallback` for focused tests in Plans 014
and 015. Keep the singleton module-local.

Remove imports from `docs-main.ts` that became unused. Leave
`registerAiIpc()` and its calls to `getCodexAuth()`/`safeCodexAuthError()` in
place through imports from the new module.

**Verify**:

- `rg -n "CODEX_CREDENTIALS_PATH|beginCodexCallback|function codexCredentialStore|function safeCodexAuthError" apps/docs/src/main/docs-main.ts` → no matches.
- `rg -n "packages/ai-provider/src/auth" apps/docs/src` → no matches.
- `npm run typecheck -w @genoffice/docs` → exit 0.

### Step 3: Prove the extraction is behavior-preserving

Run existing provider auth and Docs error tests; do not add speculative tests in
this extraction because Plans 014 and 015 add behavior-specific coverage to the
new module.

**Verify**:

- `npm test -w @genoffice/ai-provider -- tests/auth.test.ts` → all pass.
- `npm test -w @genoffice/docs -- tests/codex-error.test.ts` → all pass.
- `FORMAT_BASE_REF=origin/main npm run format:check` → exit 0.
- `git diff --check` → no whitespace errors.

## Test plan

No new behavior test is required. Existing auth-service and safe-error tests are
the characterization gate. Plans 014 and 015 will create
`apps/docs/tests/codex-auth-main.test.ts` around the extracted storage and
callback functions.

## Done criteria

- [ ] `docs-main.ts` contains no credential-store or callback-server implementation.
- [ ] `docs-main.ts` imports `CodexAuthService` only through `@genoffice/ai-provider`.
- [ ] AI IPC registration remains in `docs-main.ts` and is still registered once.
- [ ] No OAuth/storage/callback behavior changed.
- [ ] Provider and Docs typechecks/tests pass.
- [ ] Only in-scope files changed.
- [ ] Plan 013 is marked DONE in `plans/README.md`.

## STOP conditions

- Extraction creates an import cycle between `docs-main.ts` and the new module.
- The unified shell requires a different user-data path or handler owner than
  `app.getPath('userData')` and current behavior cannot be preserved.
- A behavior change appears necessary to complete the move; leave it for Plans
  014-016 and report the reason.
- A gate fails twice after a reasonable correction.

## Maintenance notes

This module is intentionally Docs-owned. Cross-app rollout should first prove
the same needs in another app, then extract a shared Electron boundary; do not
pre-build that abstraction here. Review that tokens remain main-process-only.
