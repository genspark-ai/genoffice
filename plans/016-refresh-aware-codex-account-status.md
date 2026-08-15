# Plan 016: Make Codex account status refresh-aware

> **Implementation instructions**: Execute after Plans 011 and 013. Run each
> verification gate and update Plan 016 in `plans/README.md` only after review.
>
> **Drift check (run first)**:
> `git diff --stat 1878b30..HEAD -- packages/ai-provider/src/auth.ts packages/ai-provider/tests/auth.test.ts apps/docs/src/main/codex-auth-main.ts apps/docs/tests/codex-auth-main.test.ts apps/docs/src/renderer/ai/AiPanel.tsx`
> Compare live status behavior with the excerpts below. Do not patch renderer
> symptoms if the service can provide the correct state.

## Status

- **Priority**: P1
- **Effort**: S
- **Risk**: LOW
- **Depends on**: Plans 011 and 013
- **Category**: bug
- **Planned at**: commit `1878b30`, 2026-08-15

## Why this matters

`status()` currently reports `loggedIn: true` whenever a credential blob exists,
even if its access token is expired and refresh will immediately delete it.
Docs can therefore enable send and load capabilities from stale account state.
Make the provider service validate expired credentials and report confirmed
expiry without turning transient provider failures into a fake logout.

## Current state

- `packages/ai-provider/src/auth.ts:104-109` checks only `store.get()` and never
  calls `getContext()`.
- `packages/ai-provider/src/auth.ts:159-169` already owns expiry, refresh, and
  concurrent-refresh deduplication. Reuse it; do not duplicate expiry math.
- After Plan 011, confirmed HTTP 400/401 refresh rejection deletes credentials,
  while transient failures preserve them and throw a bounded temporary error.
- `apps/docs/src/main/docs-main.ts:2708-2713` returns `auth.status()` over IPC and
  otherwise maps a thrown error to `loggedIn: false`.
- `apps/docs/src/renderer/ai/AiPanel.tsx:408` disables Codex send unless
  `codexAccount.loggedIn === true`; lines 427-466 request status whenever Codex
  is selected. Correcting the service result fixes the root state path.

## Commands you will need

| Purpose         | Command                                                                               | Expected on success |
| --------------- | ------------------------------------------------------------------------------------- | ------------------- |
| Auth tests      | `npm test -w @genoffice/ai-provider -- tests/auth.test.ts`                            | exit 0              |
| Docs auth tests | `npm test -w @genoffice/docs -- tests/codex-auth-main.test.ts`                        | exit 0              |
| Typecheck       | `npm run typecheck -w @genoffice/ai-provider && npm run typecheck -w @genoffice/docs` | exit 0              |
| Format          | `FORMAT_BASE_REF=origin/main npm run format:check`                                    | exit 0              |

## Scope

**In scope**:

- `packages/ai-provider/src/auth.ts`
- `packages/ai-provider/tests/auth.test.ts`
- `apps/docs/src/main/codex-auth-main.ts` only if its safe error mapper needs the
  new temporary message
- `apps/docs/tests/codex-auth-main.test.ts` only for that mapper
- `plans/README.md` for status only

**Out of scope**:

- `AiPanel.tsx`, CSS, preload, or IPC schema changes unless provider tests prove
  the existing `CodexAccountStatus` shape cannot express the result
- Background polling, timers, token refresh scheduling, or capability caching
- Logging users out on transient refresh failures
- Pi migration or cross-app rollout

## Git workflow

- Work on the operator-provided branch; do not push or open a PR.
- Keep the status contract and tests in one logical change.

## Steps

### Step 1: Specify the four account-status states in tests

Extend `packages/ai-provider/tests/auth.test.ts` with status cases:

1. no credentials → `{ loggedIn: false }`, no network call;
2. fresh credentials → `{ loggedIn: true, email? }`, no network call;
3. expired credentials + successful refresh → logged in with the preserved email
   and updated stored credentials;
4. expired credentials + confirmed invalid refresh → logged out with a safe
   expiry error after deletion;
5. expired credentials + transient refresh failure → still logged in with the
   preserved email and a bounded temporary error; credentials remain stored.

The status object must never include access token, refresh token, account ID, or
raw provider errors.

**Verify**: `npm test -w @genoffice/ai-provider -- tests/auth.test.ts` → expired
status tests fail against the current existence-only implementation.

### Step 2: Reuse `getContext()` for expired status

Update `CodexAuthService.status()`:

- return logged out immediately when no credential exists;
- return logged in immediately for a credential outside the existing refresh skew;
- for an expired/near-expiry credential, call the existing `getContext()` path;
- after a failed refresh, inspect the store again: absence means confirmed
  logout/expiry; presence means a transient error and remains logged in;
- return only bounded user-facing messages in `CodexAccountStatus.error`.

Do not add a second refresh lock, expiry constant, or background refresh. Preserve
email from the stored credential; `getContext()` intentionally returns only the
request credential subset.

If the temporary error string crosses Docs' safe mapper during a send/capability
failure, add one explicit mapping in `safeCodexAuthError()` within the extracted
`codex-auth-main.ts`. Do not broaden raw error passthrough.

**Verify**: `npm test -w @genoffice/ai-provider -- tests/auth.test.ts` → all five
status states pass.

### Step 3: Confirm existing Docs UI consumes the corrected state

Do not change `AiPanel.tsx` unless this check fails. The existing Codex-selection
effect calls `aiCodexStatus()`, and `codexSendDisabled` derives from
`loggedIn !== true`. Confirm by reading the final diff and, if an existing
AiPanel test harness can cover it without new scaffolding, add one compact case
showing a returned logged-out status disables send.

**Verify**:

- `rg -n "const codexSendDisabled = .*loggedIn !== true" apps/docs/src/renderer/ai/AiPanel.tsx` → one match.
- `npm test -w @genoffice/docs -- tests/codex-auth-main.test.ts` → pass if the
  safe mapper changed; otherwise this command may report no new tests but exits 0.

### Step 4: Run package gates

**Verify**:

- `npm run typecheck -w @genoffice/ai-provider && npm run typecheck -w @genoffice/docs` → exit 0.
- `FORMAT_BASE_REF=origin/main npm run format:check` → exit 0.
- `git diff --check` → no whitespace errors.

## Test plan

Use the existing `memoryStore()` and injected clock/fetch. Keep renderer tests
out unless the existing harness makes the state-to-send-disabled assertion
small; the provider service is the root cause and primary test boundary.

## Done criteria

- [ ] Status distinguishes absent, fresh, refreshed, expired, and temporarily unavailable states.
- [ ] Confirmed invalid refresh reports logged out.
- [ ] Transient refresh failure reports logged in and retains credentials.
- [ ] Status never exposes credential or provider payload fields.
- [ ] Existing Docs send gating consumes the result without duplicated refresh logic.
- [ ] Focused tests, typechecks, and formatting pass.
- [ ] No out-of-scope files changed.
- [ ] Plan 016 is marked DONE in `plans/README.md`.

## STOP conditions

- Correct UI state requires exposing tokens or raw errors to the renderer.
- The existing `CodexAccountStatus` shape cannot distinguish retained credentials
  from confirmed logout without an IPC-breaking change.
- Plan 011's transient-vs-confirmed classification is absent or materially different.
- A gate fails twice after a reasonable correction.

## Maintenance notes

Status is on-demand validation, not a health monitor. If future UX needs
continuous account health, design that separately; do not add polling to this
service. Reviewers should ensure temporary outages never become logout.
