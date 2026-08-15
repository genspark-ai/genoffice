# Plan 019: Make Codex login single-flight and callback completion idempotent

> **Implementation instructions**: Follow this plan step by step. Run every
> verification command before moving on. If a STOP condition occurs, stop and
> report; do not improvise. Update this plan's row in `plans/README.md` only
> after implementation and review.
>
> **Drift check (run first)**:
> `git diff --stat 1878b30..HEAD -- packages/ai-provider/src/auth.ts packages/ai-provider/tests/auth.test.ts apps/docs/src/main/codex-auth-main.ts apps/docs/tests/codex-auth-main.test.ts`
> The current implementation is uncommitted at the planned-at SHA. Compare the
> symbols below against the working tree too. Behavioral mismatch is a STOP.

## Status

- **Priority**: P1
- **Effort**: S
- **Risk**: LOW
- **Depends on**: none
- **Category**: bug
- **Planned at**: commit `1878b30`, 2026-08-15

## Why this matters

One Docs window prevents a second click, but `CodexAuthService` is a
main-process singleton callable by multiple windows and future apps. A second
login currently replaces `activeLogin`, while a second valid callback replaces
the pending HTTP response. That can strand the first browser request and make
cancellation target the wrong attempt. One shared single-flight guard and one
callback guard fix both root causes.

## Current state

- `packages/ai-provider/src/auth.ts:115-189`: `login()` always creates a new
  callback and assigns `this.activeLogin = callback`.
- `apps/docs/src/main/codex-auth-main.ts:139-240`: after the first valid callback
  sets `settled`, the server stays open for token exchange, but the handler still
  assigns `pendingResponse = response` before `settle()` notices it is settled.
- Test patterns: `packages/ai-provider/tests/auth.test.ts` and
  `apps/docs/tests/codex-auth-main.test.ts`.
- Preserve the approved OAuth contract, port `1455`, timeout, and
  `~/.genoffice/codex-auth.json` behavior.

## Commands you will need

| Purpose        | Command                                                                               | Expected on success |
| -------------- | ------------------------------------------------------------------------------------- | ------------------- |
| Provider tests | `npm test -w @genoffice/ai-provider`                                                  | exit 0              |
| Docs test      | `npm test -w @genoffice/docs -- --run tests/codex-auth-main.test.ts`                  | exit 0              |
| Typecheck      | `npm run typecheck -w @genoffice/ai-provider && npm run typecheck -w @genoffice/docs` | exit 0              |
| Format         | `npm run format:check`                                                                | exit 0              |

## Scope

**In scope**:

- `packages/ai-provider/src/auth.ts`
- `packages/ai-provider/tests/auth.test.ts`
- `apps/docs/src/main/codex-auth-main.ts`
- `apps/docs/tests/codex-auth-main.test.ts`
- `plans/README.md` (status only)

**Out of scope**:

- OAuth parameters, token exchange, refresh policy, credential schema, or port
- Renderer state guards
- A queue, multi-account manager, or new concurrency abstraction

## Git workflow

- Work on the current branch; do not create, push, or open a PR.
- If asked to commit: `fix overlapping Codex login attempts`.

## Steps

### Step 1: Make `CodexAuthService.login()` single-flight

Keep one in-flight login promise on the service. A concurrent `login()` call
must return that attempt and must not call `beginCallback`, generate another
authorization URL, or replace the callback. Clear promise and callback only
when the attempt settles. Keep `cancelLogin()` and `logout()` targeting it.
Use a direct implementation; do not add a queue or public abstraction.

Add a service test starting two logins before resolving the callback. Assert
`beginCallback` and `openBrowser` run once, both calls resolve to the same
account result, and cancellation targets that callback. Also prove a later
login can start after cancellation/settlement.

**Verify**: `npm test -w @genoffice/ai-provider` -> all tests pass.

### Step 2: Reject duplicate valid callback requests

In `beginCodexCallback`, check `settled` before assigning `pendingResponse`.
Finish a later valid request immediately with a bounded non-success response;
do not expose code/state or replace the first response. The first response must
still receive `complete(true|false)`.

Add a regression test sending two valid callbacks before `complete(true)`.
The second finishes without becoming pending, and the first receives the
successful completion page when completed.

**Verify**: `npm test -w @genoffice/docs -- --run tests/codex-auth-main.test.ts` -> all pass.

### Step 3: Run focused gates

```sh
npm run typecheck -w @genoffice/ai-provider
npm run typecheck -w @genoffice/docs
npm run format:check
git diff --check
```

Expected: every command exits 0.

## Test plan

- Concurrent calls share one callback/browser flow.
- Cancellation settles shared callers and permits a later login.
- A duplicate valid HTTP callback cannot replace the first response.
- Existing invalid-state, collision, timeout, cancellation, and storage tests
  remain green.

## Done criteria

- [ ] At most one login exists per `CodexAuthService` instance.
- [ ] Duplicate callbacks cannot replace `pendingResponse`.
- [ ] Regression tests fail on old behavior and pass after the fix.
- [ ] Focused tests, typecheck, formatting, and diff check pass.

## STOP conditions

- Fixing this requires changing OAuth parameters or callback port.
- Concurrent callers are newly required to open separate browser flows.
- A verification fails twice after a reasonable in-scope correction.

## Maintenance notes

The process-wide attempt matches the single fixed localhost port. Add
per-account concurrency only if simultaneous accounts and distinct callback
endpoints become a real product requirement.
