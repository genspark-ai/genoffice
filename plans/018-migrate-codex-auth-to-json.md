# Plan 018: Move Codex auth to GenOffice-owned JSON storage

> **Implementation instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving on. The
> operator selected this plan for immediate execution; do not commit, push, or
> open a PR. When complete, update the status row in `plans/README.md`.
>
> **Revision**: The operator chose a clean break after the first implementation.
> Do not retain or import the old safeStorage credential; existing users sign in
> again.
>
> **Drift check (run first)**:
> `git diff --stat 1878b30..HEAD -- apps/docs/src/main/codex-auth-main.ts apps/docs/tests/codex-auth-main.test.ts`
> The working tree already contains the extracted auth module from Plans
> 013–016; compare the current excerpts below against the live files before
> editing. Stop if the auth boundary has moved or the credential contract has
> changed.

## Status

- **Priority**: P1
- **Effort**: M
- **Risk**: MED — the new file contains bearer credentials in plaintext and
  users must complete a fresh login
- **Depends on**: `plans/013-extract-docs-codex-auth-main.md`,
  `plans/016-refresh-aware-codex-account-status.md`
- **Category**: migration / security
- **Planned at**: commit `1878b30`, 2026-08-15

## Why this matters

GenOffice currently stores Codex OAuth credentials through Electron
`safeStorage`, which causes macOS Keychain/safe-storage prompts and makes the
login unavailable when that platform service is unavailable. Pi and OpenCode
use an app-owned JSON credential file with restrictive permissions; GenOffice
already uses the same general pattern for its Genspark identity in
`~/.genoffice/auth.json`.

Move new Codex credentials to `~/.genoffice/codex-auth.json`, using a separate
file so Genspark credentials and Codex credentials cannot overwrite each other.
Existing users intentionally sign in again. The old encrypted file is not read,
and normal JSON reads, writes, refreshes, and logout must never call
`safeStorage`.

This follows the storage approach, not the external file locations: do not
read or modify Pi's `~/.pi/agent/auth.json`, OpenCode's
`~/.local/share/opencode/auth.json`, or Codex CLI's `~/.codex/auth.json`.

## Current state

- `apps/docs/src/main/codex-auth-main.ts:15-16` derives the new JSON path from
  `GENOFFICE_AUTH_DIR` or `~/.genoffice`; the old `userData` credential path is
  intentionally not part of the runtime contract.
- `apps/docs/src/main/codex-auth-main.ts:125-136` implements every
  `CodexCredentialStore` method against the JSON file only.
- `apps/docs/src/main/docs-main.ts:2554-2581` calls the store only through
  `getCodexAuth()` for status, login, cancel, and logout. Keep this IPC and
  renderer boundary unchanged.
- `packages/ai-provider/src/auth.ts:11-20` is the provider contract. The
  store receives `{ accessToken, refreshToken, accountId, expiresAt, email? }`;
  do not change that provider-facing shape for a storage-only migration. Its
  storage comment must describe main-process storage without requiring
  encryption.
- `packages/ai-search/src/genoffice-auth.ts:111-161` is the closest in-repo
  convention: an app-owned JSON file under `~/.genoffice`, a test-directory
  override through `GENOFFICE_AUTH_DIR`, and mode `0600` on writes. Its
  `auth.json` is the Genspark file and must remain untouched.
- Upstream reference behavior: Pi documents OAuth credentials in
  `~/.pi/agent/auth.json` and OpenCode writes OAuth entries with `type`,
  `access`, `refresh`, and `expires` to its `auth.json`. These references are
  inspiration for the file-backed format, not files GenOffice should import.

### Target storage contract

Use `GENOFFICE_AUTH_DIR` when set (test isolation), otherwise
`join(homedir(), '.genoffice')`. The Codex file is
`<auth-dir>/codex-auth.json` and contains one validated OAuth object:

```json
{
  "type": "oauth",
  "access": "<access token>",
  "refresh": "<refresh token>",
  "expires": 0,
  "accountId": "<account id>",
  "email": "<optional email>"
}
```

`expires` is the same absolute epoch-millisecond value as the provider's
`expiresAt`. The adapter maps `access`/`refresh`/`expires` to the existing
`accessToken`/`refreshToken`/`expiresAt` fields at the main-process boundary.
The plan file and tests must never contain real credentials.

## Commands you will need

| Purpose        | Command                                                                                                          | Expected on success                           |
| -------------- | ---------------------------------------------------------------------------------------------------------------- | --------------------------------------------- |
| Drift check    | `git diff --stat 1878b30..HEAD -- apps/docs/src/main/codex-auth-main.ts apps/docs/tests/codex-auth-main.test.ts` | no committed drift in the in-scope auth files |
| Focused tests  | `npm test -w @genoffice/docs -- tests/codex-auth-main.test.ts`                                                   | all auth-main tests pass                      |
| Docs typecheck | `npm run typecheck -w @genoffice/docs`                                                                           | exit 0                                        |
| Formatting     | `npm run format:check`                                                                                           | exit 0                                        |
| Full tests     | `npm test`                                                                                                       | all workspace tests pass                      |
| Full typecheck | `npm run typecheck`                                                                                              | exit 0                                        |
| Build          | `npm run build:all`                                                                                              | all six app builds pass                       |
| Diff hygiene   | `git diff --check`                                                                                               | no whitespace errors                          |

## Scope

**In scope** (the only files to modify):

- `apps/docs/src/main/codex-auth-main.ts`
- `apps/docs/tests/codex-auth-main.test.ts`
- `packages/ai-provider/src/auth.ts` for the stale storage-interface comment
- `plans/README.md` for the status row and dependency note

**Out of scope** (do not touch):

- `packages/ai-provider/src/auth.ts` provider-facing credential types and
  behavior (only its storage-interface comment is in scope)
- `apps/docs/src/main/docs-main.ts`, preload channels, renderer auth state, or
  the OAuth URLs/callback protocol
- `packages/ai-search/src/genoffice-auth.ts` and its Genspark `auth.json`
- Pi, OpenCode, or Codex CLI credential files; no cross-tool credential import
- A shared multi-provider auth registry, keyring replacement, custom encryption,
  new dependency, password prompt, or UI redesign
- Plans 011–017 except the `plans/README.md` status/dependency maintenance

## Git workflow

- Work on the operator-provided branch and preserve unrelated dirty changes.
- Do not commit, push, merge, or open a PR.

## Steps

### Step 1: Replace storage tests with the migration contract

In `apps/docs/tests/codex-auth-main.test.ts`, keep the existing Electron
module mock and callback-server tests, but change the credential-store cases to
use a temporary `GENOFFICE_AUTH_DIR`. Add assertions for:

1. `set()` writes valid JSON to `codex-auth.json`, maps the provider fields to
   `type/access/refresh/expires/accountId/email`, creates the file with mode
   `0600`, and `get()` maps it back to the exact existing credentials object.
2. There is no legacy fallback: a missing JSON file returns signed out and
   requires a fresh OAuth login.
3. `delete()` removes the JSON file, and malformed JSON is treated as signed
   out without exposing file contents.

Use dummy fixture strings such as `access` and `refresh`, never realistic
tokens. Restore `GENOFFICE_AUTH_DIR` in `afterEach`.

**Verify**: `npm test -w @genoffice/docs -- tests/codex-auth-main.test.ts` →
the clean-break storage assertions are present; they may fail until Step 2
implements the target behavior.

### Step 2: Implement the file-backed Codex credential adapter

In `apps/docs/src/main/codex-auth-main.ts`:

- Add the home-directory/test-override path helpers and a private stored-file
  type. Keep the file separate from the existing Genspark `auth.json`.
- Add strict parsing/mapping for the target JSON object. Require non-empty
  strings for `access`, `refresh`, and `accountId`, a finite numeric `expires`,
  and optional string `email`; reject arrays, unknown roots, and missing fields.
- Write JSON with the standard library only. Create the parent directory with
  `mkdir(..., { recursive: true, mode: 0o700 })`; create a unique temporary
  file with mode `0600`, write pretty JSON plus a trailing newline, rename it
  into place, and clean up the temporary file on ordinary errors. Ensure the
  final file is mode `0600` so an existing file cannot retain broader bits.
- Make normal `get()`, `set()`, and `delete()` operate only on
  `codex-auth.json`. A missing file means signed out. A malformed JSON file is
  removed and treated as signed out; non-ENOENT filesystem failures propagate
  without deleting data.

**Verify**: `npm test -w @genoffice/docs -- tests/codex-auth-main.test.ts` →
all JSON format, permission, validation, and no-safeStorage tests pass.

### Step 3: Remove the legacy storage boundary

In `apps/docs/src/main/codex-auth-main.ts`, remove the old compatibility path:

- Remove the Electron `safeStorage` import, the legacy path, the availability
  guard, and all decrypt/legacy migration helpers.
- Keep `get()` limited to `codex-auth.json`; a missing file is signed out and
  must not inspect or delete any old credential.
- Remove the obsolete safe-storage error mapping. The user-visible recovery is
  the normal ChatGPT sign-in flow.

**Verify**: `npm test -w @genoffice/docs -- tests/codex-auth-main.test.ts` →
all JSON, clean-break, and callback tests pass; `rg -n
"safeStorage|codex-credentials" apps/docs/src apps/docs/tests` returns no
matches.

### Step 4: Run the repository gates and review the diff

Run the focused tests, Docs typecheck, formatting, full tests, full typecheck,
build, and `git diff --check` from the command table. Read the complete diff
and confirm that only the in-scope paths changed, that no credential
value appears in source/tests/plans, and that the renderer/IPC code is
unchanged.

**Verify**: every command in the table exits 0; `git status --short` lists only
the expected in-scope changes plus unrelated pre-existing changes; no
`safeStorage` or `codex-credentials.bin` reference remains in the Docs runtime
source or tests.

## Test plan

Follow the existing Vitest style in `apps/docs/tests/codex-auth-main.test.ts`:
mock Electron before dynamically importing the main-process module, isolate
filesystem state with `mkdtempSync`, and assert public store behavior rather
than exporting helper functions. The security-sensitive regression cases are
normal JSON operation, `0600` permissions, missing-file signed-out behavior,
malformed-file cleanup, and fresh-login recovery. Keep the callback lifecycle
and safe-error tests already covering Plans 015–016.

## Done criteria

- [x] New Codex credentials are stored as `~/.genoffice/codex-auth.json` (or
      the `GENOFFICE_AUTH_DIR` test override), not in Electron safeStorage.
- [x] The JSON file is validated, written atomically through a `0600` temp file,
      and ends with mode `0600`.
- [x] No Codex runtime source or test imports or calls `safeStorage`.
- [x] Existing `codex-credentials.bin` data is not read; users sign in again.
- [x] No Pi/OpenCode/Codex CLI credential path is read or modified.
- [x] `npm test -w @genoffice/docs -- tests/codex-auth-main.test.ts` exits 0.
- [x] `npm run typecheck -w @genoffice/docs` exits 0.
- [x] `npm run format:check`, `npm test`, `npm run typecheck`,
      `npm run build:all`, and `git diff --check` exit 0.
- [x] Only in-scope files are modified by this plan, and Plan 018 is marked
      DONE in `plans/README.md`.

## STOP conditions

Stop and report instead of improvising if:

- The live auth module no longer matches the current-state excerpts or the
  provider credential contract has changed.
- The app must share one file with Genspark or an external Pi/OpenCode/Codex
  installation to preserve an existing user login; that is a separate schema
  and ownership decision.
- The platform cannot create/replace the file with mode `0600` without adding a
  dependency or broadening credential exposure.
- The clean-break behavior would require reading or importing an old credential
  to preserve a login.
- Any focused or repository verification command fails twice after a reasonable
  focused correction, or the change requires an out-of-scope file.

## Maintenance notes

Reviewers should inspect the boundary mapping and the `0600` write ordering
first. The plaintext JSON is intentionally user-owned and compatible with the
Pi/OpenCode style, so file permissions and redacted errors are the security
contract; do not casually move the tokens into renderer settings or logs. A
future explicit “import credentials from Pi/OpenCode” feature would need its
own schema/version and user-visible consent flow; it is not implied by this
migration.
