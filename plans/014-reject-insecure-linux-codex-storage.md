# Plan 014: Reject insecure Linux Codex credential storage

> **Implementation instructions**: Follow this security plan exactly and run
> each gate. Do not add a plaintext fallback. Update Plan 014 in
> `plans/README.md` only after implementation and review.
>
> **Drift check (run first)**:
> `git diff --stat 1878b30..HEAD -- apps/docs/src/main/codex-auth-main.ts apps/docs/tests/codex-auth-main.test.ts`
> This plan expects Plan 013's extracted module. If it does not exist, STOP.

## Status

- **Priority**: P1
- **Effort**: S
- **Risk**: LOW
- **Depends on**: Plan 013
- **Category**: security
- **Planned at**: commit `1878b30`, 2026-08-15

## Why this matters

Electron can report encryption as available on Linux while using its
`basic_text` backend, which protects data with a hardcoded plaintext password.
The current check would write long-lived OAuth access and refresh tokens in that
mode. GenOffice must refuse sign-in/storage when the selected Linux backend is
`basic_text`; there is no acceptable plaintext fallback for account credentials.

## Current state

- Before Plan 013, the implementation is at
  `apps/docs/src/main/docs-main.ts:2505-2536`; after Plan 013 it must be in
  `apps/docs/src/main/codex-auth-main.ts` with the same behavior.
- Both `get()` and `set()` currently check only
  `safeStorage.isEncryptionAvailable()`.
- Electron's official `safeStorage` documentation states that Linux data is
  unprotected when `getSelectedStorageBackend()` returns `basic_text`:
  <https://www.electronjs.org/docs/latest/api/safe-storage>.
- Existing failure text is `Secure credential storage is unavailable`; preserve
  this safe, renderer-facing contract.

## Commands you will need

| Purpose           | Command                                                        | Expected on success |
| ----------------- | -------------------------------------------------------------- | ------------------- |
| Focused Docs test | `npm test -w @genoffice/docs -- tests/codex-auth-main.test.ts` | exit 0              |
| Docs typecheck    | `npm run typecheck -w @genoffice/docs`                         | exit 0              |
| Format            | `FORMAT_BASE_REF=origin/main npm run format:check`             | exit 0              |

## Scope

**In scope**:

- `apps/docs/src/main/codex-auth-main.ts`
- `apps/docs/tests/codex-auth-main.test.ts` (create if Plan 015 has not created it)
- `plans/README.md` for status only

**Out of scope**:

- A custom encryption scheme, password prompt, keyring dependency, or plaintext fallback
- Migrating to Electron's asynchronous safe-storage API
- Deleting an existing credential file merely because the backend is currently unavailable
- OAuth refresh/login semantics or UI redesign

## Git workflow

- Work on the operator-provided branch; do not push or open a PR.
- Keep the storage predicate and its tests together as one logical change.

## Steps

### Step 1: Add Linux backend characterization tests

In `apps/docs/tests/codex-auth-main.test.ts`, mock Electron `app` and
`safeStorage` before dynamically importing `codex-auth-main.ts`. Use a temporary
directory for `app.getPath('userData')` and clean it after the test. Add compact
tests proving:

1. Linux + `isEncryptionAvailable() === true` + backend `basic_text` causes
   `codexCredentialStore().set()` to reject with the existing secure-storage
   error and writes no credential file.
2. Linux + a protected backend such as `gnome_libsecret` permits the existing
   encrypted write path.
3. Non-Linux behavior depends only on `isEncryptionAvailable()` and never calls
   the Linux-only backend method.

Stub `process.platform` only for the duration of each test and restore it.
Never place realistic tokens or secrets in fixtures.

**Verify**: `npm test -w @genoffice/docs -- tests/codex-auth-main.test.ts` → the
`basic_text` case fails before the production guard is added.

### Step 2: Fail closed for `basic_text`

In `apps/docs/src/main/codex-auth-main.ts`, centralize the current availability
check in the smallest local helper used by both `get()` and `set()`. It must
throw `Secure credential storage is unavailable` when:

- encryption is unavailable on any platform; or
- `process.platform === 'linux'` and
  `safeStorage.getSelectedStorageBackend() === 'basic_text'`.

Do not treat `unknown` as secure if it occurs after app readiness. If tests or
current startup ordering prove the call can occur before readiness, STOP and
report rather than guessing a policy. Do not delete an existing file on this
availability error; a protected backend may become available later.

**Verify**: `npm test -w @genoffice/docs -- tests/codex-auth-main.test.ts` → all
storage tests pass.

### Step 3: Run Docs gates

**Verify**:

- `npm run typecheck -w @genoffice/docs` → exit 0.
- `FORMAT_BASE_REF=origin/main npm run format:check` → exit 0.
- `git diff --check` → no whitespace errors.

## Test plan

Use Vitest module mocks and a temporary user-data directory. Test the public
behavior of `codexCredentialStore()`; do not export a separate platform policy
function merely to test it.

## Done criteria

- [ ] Linux `basic_text` can neither read nor write Codex credentials.
- [ ] Protected Linux backends retain the existing encrypted behavior.
- [ ] Other platforms retain the existing availability check.
- [ ] Availability failure does not delete the encrypted credential file.
- [ ] Focused tests, Docs typecheck, and formatting pass.
- [ ] Only in-scope files changed.
- [ ] Plan 014 is marked DONE in `plans/README.md`.

## STOP conditions

- Electron 43.3.0 does not expose `getSelectedStorageBackend()` as documented.
- The function is called before Electron app readiness and returns `unknown` in
  a real supported startup path.
- A secure implementation would require storing a password or key beside the ciphertext.
- A gate fails twice after a reasonable correction.

## Maintenance notes

Reviewers should verify that `basic_text` fails closed and that no catch block
turns this into a corrupt-file deletion. Electron now recommends asynchronous
safe storage, but that migration is deliberately deferred until its availability
and key-rotation behavior can be designed separately.
