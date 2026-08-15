# Plan 011: Preserve Codex credentials across transient refresh failures

> **Implementation instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the next
> step. If a STOP condition occurs, report it instead of improvising. When done,
> update this plan's status row in `plans/README.md`.
>
> **Drift check (run first)**:
> `git diff --stat 1878b30..HEAD -- packages/ai-provider/src/auth.ts packages/ai-provider/tests/auth.test.ts`
> If either file changed, compare the excerpts below with the live code before
> proceeding. A behavioral mismatch is a STOP condition.

## Status

- **Priority**: P1
- **Effort**: S
- **Risk**: MED
- **Depends on**: none
- **Category**: bug, security
- **Planned at**: commit `1878b30`, 2026-08-15

## Why this matters

An expired access token currently causes the refresh token and account metadata
to be deleted for every refresh error, including a dropped connection, rate
limit, or provider outage. That converts a recoverable service failure into a
forced sign-in and destroys the only locally stored recovery credential. Delete
credentials only when the token endpoint has positively rejected the refresh.

The maintainer has approved this GenOffice-owned OAuth contract. Do not replace
it with Pi, read Pi/Codex CLI credential files, or add a Pi dependency.

## Current state

- `packages/ai-provider/src/auth.ts:180-191` catches every failure from
  `exchange()`, deletes the store, and throws `ChatGPT sign-in expired`:

  ```ts
  try {
    const refreshed = await this.exchange(...)
    await this.deps.store.set(refreshed)
    return this.context(refreshed)
  } catch {
    await this.deps.store.delete()
    throw new Error('ChatGPT sign-in expired')
  }
  ```

- `packages/ai-provider/src/auth.ts:198-203` discards the response status inside
  a generic `Error`, so `refresh()` cannot distinguish invalid credentials from
  transient failures.
- `packages/ai-provider/tests/auth.test.ts:171-188` already establishes the
  expected destructive case: an HTTP 401 refresh response deletes credentials.
- Tests use injected `fetch`, `clock`, and an in-memory store. Extend this
  pattern; do not introduce a second auth service or retry framework.

## Commands you will need

| Purpose            | Command                                                    | Expected on success              |
| ------------------ | ---------------------------------------------------------- | -------------------------------- |
| Provider tests     | `npm test -w @genoffice/ai-provider -- tests/auth.test.ts` | exit 0; all auth tests pass      |
| Provider typecheck | `npm run typecheck -w @genoffice/ai-provider`              | exit 0; no TypeScript errors     |
| Format check       | `FORMAT_BASE_REF=origin/main npm run format:check`         | exit 0; no changed-file warnings |

## Scope

**In scope** (the only source/test files to modify):

- `packages/ai-provider/src/auth.ts`
- `packages/ai-provider/tests/auth.test.ts`
- `plans/README.md` for the status update

**Out of scope**:

- OAuth URLs, client ID, scopes, PKCE, or token payload shape
- Automatic retries or backoff
- Docs IPC/UI behavior; Plan 016 handles account-status consistency
- Pi, `@mariozechner/pi-ai`, Codex CLI files, or renderer credential storage

## Git workflow

- Work on the operator-provided branch; do not create another branch unless asked.
- Keep this as one focused logical change. Existing history uses concise
  conventional-style subjects such as `fix(docx): ...`.
- Do not push, rewrite history, or open a PR unless the operator asks.

## Steps

### Step 1: Characterize destructive and non-destructive refresh failures

In `packages/ai-provider/tests/auth.test.ts`, retain the existing HTTP 401 test
and add table-driven or compact tests proving:

1. HTTP 400 and HTTP 401 are confirmed refresh rejection responses: `getContext()`
   rejects with the safe expired message and calls `store.delete()` once.
2. A network rejection, HTTP 408/429/5xx response, and malformed HTTP 200 token
   response do **not** call `store.delete()`.
3. A later `getContext()` can retry after a transient failure; the existing
   `refreshPromise.finally()` must not leave a rejected promise cached.

Do not assert or log response bodies or credential values beyond the existing
test fixtures.

**Verify**: `npm test -w @genoffice/ai-provider -- tests/auth.test.ts` → the new
tests fail only because current code deletes credentials indiscriminately.

### Step 2: Preserve enough HTTP classification for refresh

In `packages/ai-provider/src/auth.ts`, add the smallest internal error shape that
retains an OAuth token response's numeric HTTP status. Keep it private to this
file unless another compiled caller genuinely needs it. Preserve the existing
safe login error text (`ChatGPT sign-in failed (HTTP NNN)`).

Change `refresh()` so only a token endpoint HTTP 400 or 401 deletes the stored
credentials and becomes `ChatGPT sign-in expired`. For network failures,
timeouts, HTTP 408/429/5xx, and invalid success payloads, retain the credentials
and throw a bounded error such as `ChatGPT sign-in temporarily unavailable`.
Never include the raw response body, URL, access token, or refresh token.

Do not add retries: retaining the credential and clearing `refreshPromise` is
enough to let the next user action retry.

**Verify**: `npm test -w @genoffice/ai-provider -- tests/auth.test.ts` → all tests
pass, including the new preservation and retry cases.

### Step 3: Run the package gates

**Verify**:

- `npm run typecheck -w @genoffice/ai-provider` → exit 0.
- `FORMAT_BASE_REF=origin/main npm run format:check` → exit 0 for both in-scope files.
- `git diff --check` → no whitespace errors.

## Test plan

- Extend `packages/ai-provider/tests/auth.test.ts` using the `memoryStore()` and
  injected `fetch` pattern at lines 23-53.
- Cover confirmed invalidation, transient preservation, malformed success, and
  retry after failure.
- Do not add live OAuth tests or network calls.

## Done criteria

- [ ] HTTP 400/401 refresh rejection deletes credentials and reports expiry.
- [ ] Network, 408, 429, 5xx, and malformed-success failures retain credentials.
- [ ] A transient failure does not poison the concurrent-refresh promise.
- [ ] Provider auth tests and typecheck pass.
- [ ] No raw token response content crosses the error boundary.
- [ ] No source/test file outside the in-scope list changed.
- [ ] `plans/README.md` marks Plan 011 DONE after review.

## STOP conditions

- The issuer uses a confirmed-invalid refresh response outside HTTP 400/401 and
  the distinction cannot be made without parsing a sensitive response body.
- Correct classification requires changing the approved OAuth request contract.
- A verification command fails twice after a reasonable correction.
- The fix appears to require renderer or app-specific auth state changes.

## Maintenance notes

Reviewers should focus on the deletion predicate: it must fail closed for
confirmed invalid credentials without treating provider availability as account
revocation. If the issuer later exposes a stable typed OAuth error contract, add
only the minimum validated code needed to refine this predicate.
