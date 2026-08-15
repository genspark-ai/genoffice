# Plan 029: Reject insecure Codex credential files on read

> **Implementation instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving on. Do
> not commit, push, or open a PR. When complete, update this plan's status row
> in `plans/README.md` after implementation and review.
>
> **Drift check (run first)**: `git diff --stat 3723808..HEAD -- packages/ai-provider/src/codex-auth-node.ts packages/ai-provider/tests/codex-auth-node.test.ts plans/README.md`
> If the credential-file helpers or tests have changed since this plan was
> written, compare the live code to the excerpts below. Stop if the adapter
> contract or supported platforms have materially changed.

## Status

- **Priority**: P1
- **Effort**: S
- **Risk**: MED — authentication must fail safely rather than read a file that
  violates the credential-storage contract
- **Depends on**: `plans/018-migrate-codex-auth-to-json.md`
- **Category**: security
- **Planned at**: commit `3723808`, 2026-08-16

## Why this matters

Plan 018 intentionally stores OAuth refresh credentials in
`~/.genoffice/codex-auth.json` and writes them with mode `0600`. The write path
repairs permissions, but the read path currently accepts an existing path
without checking file type or permission bits. A refresh credential that was
made group/world-readable, replaced with a symlink, or otherwise does not meet
the storage contract must never be loaded.

Harden the shared Node adapter at the single read boundary. On POSIX, reject
non-regular, symlinked, or group/world-accessible credential files with the
existing generic credential error; preserve the file for manual recovery. Do
not create a new storage abstraction, alter OAuth, or introduce a dependency.

## Current state

- `packages/ai-provider/src/codex-auth-node.ts` is pure Node code shared by
  the suite's Electron main runtimes. It owns the JSON credential path and
  `codexCredentialStore()`; renderers do not read credential files.
- `writeCredentialFile()` already creates its temporary file with `0600` and
  applies `chmod(path, 0o600)` after rename.
- `readCredentialFile()` immediately passes the path to `readFile()`, then
  parses the data. It has no file metadata check.
- `packages/ai-provider/tests/codex-auth-node.test.ts` creates an isolated
  `GENOFFICE_AUTH_DIR` before each test and already verifies the write mode.

```ts
// packages/ai-provider/src/codex-auth-node.ts:77-95
await mkdir(directory, { recursive: true, mode: 0o700 })
const handle = await open(temporaryPath, 'wx', 0o600)
// ... write and fsync ...
await rename(temporaryPath, path)
await chmod(path, 0o600)

// packages/ai-provider/src/codex-auth-node.ts:98-115
async function readCredentialFile(path: string): Promise<CodexCredentials | undefined> {
  let raw: string
  try {
    raw = await readFile(path, 'utf8')
  } catch (error) {
    if (fileNotFound(error)) return undefined
    throw credentialFileError()
  }
  // JSON validation; malformed data is removed
}

// packages/ai-provider/tests/codex-auth-node.test.ts:54-69
await store.set(credentials)
expect(statSync(path).mode & 0o777).toBe(0o600)
await expect(store.get()).resolves.toEqual(credentials)
```

Use Node's built-in `node:fs` constants and `node:fs/promises` APIs only. The
project convention is to return the generic `credentialFileError()` to avoid
exposing filesystem details. Keep malformed JSON behavior unchanged: malformed
regular credential files are removed and treated as signed out.

## Commands you will need

| Purpose            | Command                                                                                                                                           | Expected on success                                                                                      |
| ------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------- |
| Drift check        | `git diff --stat 3723808..HEAD -- packages/ai-provider/src/codex-auth-node.ts packages/ai-provider/tests/codex-auth-node.test.ts plans/README.md` | No unexpected committed drift in scope                                                                   |
| Focused test       | `npm test -w @genoffice/ai-provider -- tests/codex-auth-node.test.ts`                                                                             | All credential-store and callback tests pass                                                             |
| Provider typecheck | `npm run typecheck -w @genoffice/ai-provider`                                                                                                     | Exit 0                                                                                                   |
| Full tests         | `npm test`                                                                                                                                        | Every workspace test passes                                                                              |
| Full typecheck     | `npm run typecheck`                                                                                                                               | Exit 0, no TypeScript errors                                                                             |
| Lint               | `npm run lint`                                                                                                                                    | Exit 0, no lint errors                                                                                   |
| Formatting         | `npm run format:check`                                                                                                                            | Exit 0 for committed project files; separately report any pre-existing untracked-tool formatting failure |
| Diff hygiene       | `git diff --check`                                                                                                                                | No whitespace errors                                                                                     |

## Scope

**In scope** (the only files to modify):

- `packages/ai-provider/src/codex-auth-node.ts`
- `packages/ai-provider/tests/codex-auth-node.test.ts`
- `plans/README.md` (status row only)

**Out of scope** (do not touch):

- OAuth URLs, PKCE, callback lifecycle, token refresh, or user-facing error
  strings.
- Renderer, preload, Electron IPC, settings, and any app-specific runtime.
- Credential JSON schema, write format, normal missing-file behavior, or
  malformed-regular-file cleanup.
- Directory-permission migration, secure-storage/keychain work, custom
  encryption, a shared filesystem framework, dependencies, or the uncommitted
  anti-slop plugin configuration.

## Git workflow

- Work on the operator-provided branch and preserve unrelated dirty changes.
- Do not commit, push, merge, or open a PR.

## Steps

### Step 1: Add focused POSIX regressions before changing the reader

In `packages/ai-provider/tests/codex-auth-node.test.ts`, extend the existing
`codexCredentialStore` suite using its temporary auth directory and dummy
fixture credentials:

1. Write valid credentials through the store, change the resulting file to
   mode `0644` on POSIX, and assert `get()` rejects with the generic
   `ChatGPT sign-in credentials unavailable` error. Assert the file still
   exists; an insecure file must not be deleted automatically.
2. On POSIX only, replace the credential path with a symlink to a valid
   fixture file and assert `get()` rejects with the same generic error and does
   not remove either path. Skip this case on Windows because POSIX permission
   and no-follow semantics are not portable there.
3. Keep the existing malformed JSON test unchanged: a regular malformed file
   remains a signed-out cleanup case, not an insecure-file case.

Use the Node standard-library test setup already in this file. Do not export
or directly unit-test a new internal helper.

**Verify**: `npm test -w @genoffice/ai-provider -- tests/codex-auth-node.test.ts`
→ the two new tests initially fail only because the reader lacks metadata
validation.

### Step 2: Make the one credential read boundary fail closed on POSIX

In `packages/ai-provider/src/codex-auth-node.ts`, change only the code path
that opens a credential file:

1. On POSIX, open the file without following symlinks using Node's native file
   flags, inspect the opened handle with `stat()`, and reject unless it is a
   regular file with no group or other permission bits (`mode & 0o077 === 0`).
   Read the JSON through that verified handle, then close it in `finally`.
2. Preserve the current missing-file result (`undefined`) for `ENOENT`.
   Map unsafe type/mode/symlink/open failures to `credentialFileError()`; do
   not expose OS error details.
3. On Windows, retain the existing `readFile(path, 'utf8')` behavior because
   POSIX file modes and `O_NOFOLLOW` are not portable there. Keep this platform
   branch local to the credential reader; do not add a general platform layer.
4. Do not call `removeCredentialFile()` for rejected unsafe paths. Only a file
   that was safely opened and then fails JSON/schema validation should follow
   the existing malformed-file cleanup path.

Prefer `node:fs` / `node:fs/promises` imports already used by the module. Do
not introduce a time-of-check/time-of-use `lstat()` followed by a separate
`readFile()` on POSIX; validate the same file handle from which bytes are read.

**Verify**: `npm test -w @genoffice/ai-provider -- tests/codex-auth-node.test.ts`
→ all existing tests plus the insecure-mode and symlink regressions pass.

### Step 3: Run gates and review the security boundary

Run every command in the table. Inspect the final diff and confirm all
credential failures remain generic, no token values appear in output/source,
and only the single reader changes. Update Plan 029's status to `DONE` only
after the focused and full suites pass.

**Verify**: `npm test`, `npm run typecheck`, `npm run lint`, and
`git diff --check` all succeed; `git diff --name-only` is limited to the
source file, its test, and the plan status row (besides unrelated pre-existing
changes).

## Test plan

- Existing: valid write/read with `0600`, malformed regular JSON cleanup,
  logout deletion, and callback lifecycle.
- New POSIX regressions: a mode-`0644` credential file and a credential-path
  symlink fail with the generic error and remain on disk.
- Compatibility: missing credentials still return signed out; valid `0600`
  credentials still load exactly as before.

## Done criteria

- [ ] On POSIX, `get()` never reads a symlink, non-regular file, or file with
      group/other permission bits set.
- [ ] Insecure files are not deleted or replaced during a read attempt.
- [ ] Valid `0600` credentials, missing files, logout, and malformed regular
      JSON retain their present behavior.
- [ ] `npm test -w @genoffice/ai-provider -- tests/codex-auth-node.test.ts`,
      `npm test`, `npm run typecheck`, `npm run lint`, and `git diff --check`
      succeed.
- [ ] No files outside scope are modified, apart from unrelated pre-existing
      working-tree changes.
- [ ] Plan 029 is marked `DONE` in `plans/README.md` only after review.

## STOP conditions

Stop and report instead of improvising if:

- The target Node runtime does not expose a safe no-follow open flag on a
  supported POSIX platform.
- The current adapter must change its public `CodexCredentialStore` interface
  or credential JSON format to implement the guard.
- The symlink regression cannot be created in the test environment for a
  reason other than Windows platform behavior.
- The full suite fails for a reason unrelated to the pre-existing brittle Docs
  test or the changed credential-store tests.

## Maintenance notes

The write path's `0600` mode is a storage invariant only if reads enforce it.
Any future credential migration or alternate reader must preserve the same
fail-closed POSIX behavior and must not weaken it with a separate metadata
check followed by a path-based read.
