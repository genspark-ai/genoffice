# Plan 015: Harden the Codex localhost callback lifecycle

> **Implementation instructions**: Follow the steps and gates in order. Preserve
> the approved OAuth parameters and localhost-only binding. Update Plan 015 in
> `plans/README.md` after review.
>
> **Drift check (run first)**:
> `git diff --stat 1878b30..HEAD -- packages/ai-provider/src/auth.ts packages/ai-provider/tests/auth.test.ts apps/docs/src/main/codex-auth-main.ts apps/docs/tests/codex-auth-main.test.ts`
> This plan expects the extraction from Plan 013. Stop if the symbols have moved
> elsewhere or their contract has changed.

## Status

- **Priority**: P1
- **Effort**: M
- **Risk**: MED
- **Depends on**: Plan 013
- **Category**: bug, security
- **Planned at**: commit `1878b30`, 2026-08-15

## Why this matters

The browser currently opens immediately after `server.listen()` is called, not
after the callback socket is actually listening. A port collision can therefore
open an unusable login page before the app reports failure. In addition, any
request with a missing/wrong state rejects and closes the active login, allowing
unrelated localhost traffic to terminate the flow. Bind first, ignore invalid
callbacks after a safe 400 response, and settle the login exactly once.

## Current state

- `packages/ai-provider/src/auth.ts:34-39` gives `CodexCallbackHandle` only
  `wait`, `cancel`, and optional `complete`.
- `packages/ai-provider/src/auth.ts:114-131` calls `beginCallback()` and then
  opens the browser without awaiting server readiness.
- Before extraction, `apps/docs/src/main/docs-main.ts:2590-2595` responds 400 to
  a bad callback and immediately rejects the entire login.
- Before extraction, `apps/docs/src/main/docs-main.ts:2600-2602` installs an
  error handler and calls `listen()`, but exposes no readiness promise.
- The existing five-minute timeout, `127.0.0.1` binding, exact callback path,
  deferred browser success response, and PKCE state check must remain.

## Commands you will need

| Purpose             | Command                                                                               | Expected on success |
| ------------------- | ------------------------------------------------------------------------------------- | ------------------- |
| Provider auth tests | `npm test -w @genoffice/ai-provider -- tests/auth.test.ts`                            | exit 0              |
| Docs callback tests | `npm test -w @genoffice/docs -- tests/codex-auth-main.test.ts`                        | exit 0              |
| Typecheck           | `npm run typecheck -w @genoffice/ai-provider && npm run typecheck -w @genoffice/docs` | exit 0              |
| Format              | `FORMAT_BASE_REF=origin/main npm run format:check`                                    | exit 0              |

## Scope

**In scope**:

- `packages/ai-provider/src/auth.ts`
- `packages/ai-provider/tests/auth.test.ts`
- `apps/docs/src/main/codex-auth-main.ts`
- `apps/docs/tests/codex-auth-main.test.ts`
- `plans/README.md` for status only

**Out of scope**:

- Changing port 1455, redirect URI, OAuth scopes, issuer, or browser UX copy
- Binding beyond `127.0.0.1`
- Device-code flow, custom URI schemes, Pi, or Codex CLI integration
- Supporting multiple simultaneous login flows

## Git workflow

- Work on the operator-provided branch. Do not push or open a PR.
- Keep provider contract and Electron callback implementation in one logical change.

## Steps

### Step 1: Characterize readiness ordering in the provider service

Add a required `ready: Promise<void>` to `CodexCallbackHandle`. Update the auth
test helper and existing callback fixtures. Add a deferred-readiness test proving
`CodexAuthService.login()` does not call `openBrowser()` until `ready` resolves,
then continues through the existing callback/exchange path.

Add a rejection case proving readiness failure does not open the browser and
still calls `complete(false)`/cleanup exactly once where applicable.

**Verify**: `npm test -w @genoffice/ai-provider -- tests/auth.test.ts` → the new
ordering test fails before `login()` awaits readiness.

### Step 2: Expose actual server readiness and single settlement

In `beginCodexCallback()` inside `apps/docs/src/main/codex-auth-main.ts`:

- return a `ready` promise that resolves only from the server's `listening` event;
- reject readiness and `wait` safely on a bind/listen error;
- ensure timeout, cancel, bind error, successful callback, and `complete()` close
  the server and settle promises at most once;
- preserve `timeout.unref()` and the five-minute limit;
- keep the server bound to `127.0.0.1`.

Avoid a new lifecycle class. A few local closures and one settlement guard match
the current implementation.

**Verify**: `npm run typecheck -w @genoffice/docs` → exit 0.

### Step 3: Reject bad requests without terminating the valid login

For a wrong path, retain 404 and continue listening. For missing code/state or
state mismatch, send a generic 400 response and continue listening; do not call
the login rejection function. The provider service still compares the returned
state before exchange as defense in depth.

In `apps/docs/tests/codex-auth-main.test.ts`, start the callback on a test-only
port accepted as an optional function parameter (production default remains
1455), await `ready`, then prove:

1. an invalid-state request receives 400 and `handle.wait` remains pending;
2. a later correct-state/code request resolves `handle.wait`;
3. `complete(true)` sends the pending 200 response and closes the server;
4. a port collision rejects `ready` and does not hang;
5. cancel and timeout close the server and reject once.

Use fake timers only for timeout and restore them. Use placeholder codes/states.

**Verify**: `npm test -w @genoffice/docs -- tests/codex-auth-main.test.ts` → all
callback lifecycle tests pass without leaked handles.

### Step 4: Await readiness before browser launch

In `CodexAuthService.login()`, await `callback.ready` immediately before
`openBrowser()`. Preserve active-login cancellation, state comparison,
credential-save-before-browser-success, and final active-handle cleanup.

**Verify**:

- `npm test -w @genoffice/ai-provider -- tests/auth.test.ts` → all pass.
- `npm test -w @genoffice/docs -- tests/codex-auth-main.test.ts` → all pass and
  Vitest exits without open-handle warnings.
- `npm run typecheck -w @genoffice/ai-provider && npm run typecheck -w @genoffice/docs` → exit 0.
- `FORMAT_BASE_REF=origin/main npm run format:check` → exit 0.

## Test plan

- Provider tests own ordering between callback readiness and browser launch.
- Docs main-process tests own actual HTTP server behavior and cleanup.
- Reuse existing injected callback fixtures in `auth.test.ts`; do not perform a
  real OAuth exchange or open a browser.

## Done criteria

- [ ] Browser launch happens only after localhost callback readiness.
- [ ] Port bind failure never opens the browser and settles promptly.
- [ ] Invalid localhost requests cannot terminate a valid pending login.
- [ ] Valid callback, complete, cancel, timeout, and error paths close once.
- [ ] Server remains on `127.0.0.1` and production port remains 1455.
- [ ] Focused tests, both typechecks, and formatting pass.
- [ ] No out-of-scope files changed.
- [ ] Plan 015 is marked DONE in `plans/README.md`.

## STOP conditions

- The approved issuer redirect cannot tolerate waiting for socket readiness.
- Tests reveal the callback must bind to a non-loopback interface.
- Correct cleanup requires a new process-global callback coordinator.
- A verification gate fails twice after a reasonable correction.

## Maintenance notes

The callback handle is the sole lifecycle contract between provider auth and
Electron. Future callback transports must preserve `ready`, `wait`, `cancel`,
and deferred `complete`; do not leak HTTP objects into `@genoffice/ai-provider`.
