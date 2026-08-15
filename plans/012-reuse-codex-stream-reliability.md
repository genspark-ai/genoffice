# Plan 012: Route Codex through GenOffice streaming reliability helpers

> **Implementation instructions**: Follow this plan in order and run each gate.
> Stop on the conditions below rather than inventing a parallel streaming stack.
> When done, update Plan 012 in `plans/README.md`.
>
> **Drift check (run first)**:
> `git diff --stat 1878b30..HEAD -- packages/ai-provider/src/codex.ts packages/ai-provider/src/types.ts packages/ai-provider/src/stream.ts packages/ai-provider/tests/codex.test.ts`
> Compare all changed symbols with the excerpts below before editing.

## Status

- **Priority**: P1
- **Effort**: M
- **Risk**: MED
- **Depends on**: none
- **Category**: bug, tech-debt
- **Planned at**: commit `1878b30`, 2026-08-15

## Why this matters

Codex streaming bypasses the shared rescue fetch, connect/idle watchdog, and
wire-activity callback used by every other GenOffice provider. It also treats a
clean EOF as success even when the provider never sent `response.completed`, so
a truncated response can become a false `done` turn. Reusing the established
helpers fixes VPN/proxy resilience and completion semantics without another
transport abstraction.

## Current state

- `packages/ai-provider/src/codex.ts:258-271` calls global `fetch` and iterates
  `sseLines(response.body)` without a watchdog or byte callback.
- `packages/ai-provider/src/codex.ts:329-334` throws on explicit failure but
  returns successfully at EOF without observing `response.completed`.
- `packages/ai-provider/src/fetch.ts:19-31` already provides `aiFetch()`, which
  retries network-layer failures once through the Electron-injected rescue fetch.
- `packages/ai-provider/src/watchdog.ts:35-68` already provides
  `createStreamWatchdog()`, translating connect/idle aborts into `AiTimeoutError`.
- `packages/ai-provider/src/stream.ts:31-39` defines `StreamCallbacks.onActivity`,
  and the Anthropic/Gemini/OpenAI paths call it when bytes arrive.
- `packages/ai-provider/src/stream.ts:908-920` forwards only `signal`,
  `onDelta`, and `onToolCall` to Codex.
- `packages/ai-provider/tests/codex.test.ts:192-233` already includes a valid
  `response.completed` event but does not prove the event is required.

## Commands you will need

| Purpose        | Command                                                     | Expected on success             |
| -------------- | ----------------------------------------------------------- | ------------------------------- |
| Codex tests    | `npm test -w @genoffice/ai-provider -- tests/codex.test.ts` | exit 0; Codex tests pass        |
| Provider tests | `npm test -w @genoffice/ai-provider`                        | exit 0; all provider tests pass |
| Typecheck      | `npm run typecheck -w @genoffice/ai-provider`               | exit 0                          |
| Format         | `FORMAT_BASE_REF=origin/main npm run format:check`          | exit 0                          |

## Scope

**In scope**:

- `packages/ai-provider/src/codex.ts`
- `packages/ai-provider/src/types.ts`
- `packages/ai-provider/src/stream.ts`
- `packages/ai-provider/tests/codex.test.ts`
- `plans/README.md` for status only

**Out of scope**:

- Changes to request translation, tool execution, model catalog validation, or
  Codex OAuth
- New timeout constants, retry libraries, or provider-specific watchdogs
- Renderer busy-state behavior
- Treating `[DONE]` as the required Responses API completion event

## Git workflow

- Use the operator-provided branch and one focused commit if committing is requested.
- Do not push or open a PR.

## Steps

### Step 1: Add regression tests for the shared reliability contract

Extend `packages/ai-provider/tests/codex.test.ts` to prove:

1. a stream containing `response.completed` resolves;
2. EOF or `[DONE]` without `response.completed` rejects with a bounded protocol
   error and never looks like success;
3. received SSE bytes invoke `onActivity`;
4. an initial network rejection can use the existing rescue-fetch path;
5. caller cancellation still surfaces as `AbortError`, while watchdog expiry
   surfaces as `AiTimeoutError`.

Use fake timers only for the watchdog case and restore global fetch, timers, and
`setRescueFetch(null)` after each affected test so provider tests remain isolated.

**Verify**: `npm test -w @genoffice/ai-provider -- tests/codex.test.ts` → the new
completion/activity/reliability tests fail against the current implementation.

### Step 2: Reuse `aiFetch`, `createStreamWatchdog`, and byte activity

In `streamCodexResponse()`:

- create one `StreamWatchdog` linked to `request.signal`;
- execute the entire fetch-and-read operation inside `watchdog.guard()`;
- pass `watchdog.signal` to `aiFetch()`;
- call `watchdog.touch()` after response headers and on every received body chunk;
- pass a byte callback to `sseLines()` that also calls `request.onActivity?.()`.

Add optional `onActivity` to `CodexAdapterRequest` and forward
`StreamCallbacks.onActivity` in the `openai-codex` branch of
`streamForProvider()`. Follow the existing provider implementations in
`packages/ai-provider/src/stream.ts`; do not create a Codex-only copy.

Use the same `aiFetch` default for capability requests while preserving the
existing injectable fetch parameter used by capability tests. Do not retry HTTP
responses; `aiFetch` retries only thrown network failures.

**Verify**: `npm test -w @genoffice/ai-provider -- tests/codex.test.ts` → rescue,
activity, cancellation, and timeout tests pass.

### Step 3: Require the terminal Responses event

Track whether a parsed event has `type === 'response.completed'`. After the SSE
iterator ends, throw a bounded invalid-stream error unless that terminal event
was observed. `[DONE]` is transport noise in this implementation and must not
substitute for the protocol terminal event. Preserve explicit `error` and
`response.failed` handling, malformed-event bounds, and completed tool-call
emission.

**Verify**: `npm test -w @genoffice/ai-provider -- tests/codex.test.ts` → valid
completion resolves; truncated EOF and `[DONE]`-only streams reject.

### Step 4: Run package gates

**Verify**:

- `npm test -w @genoffice/ai-provider` → all tests pass.
- `npm run typecheck -w @genoffice/ai-provider` → exit 0.
- `FORMAT_BASE_REF=origin/main npm run format:check` → exit 0.
- `git diff --check` → no whitespace errors.

## Test plan

Model tests after the existing `streamCodexResponse` cases in
`packages/ai-provider/tests/codex.test.ts:192-417`. Add the minimum cases listed
above; do not duplicate shared watchdog unit tests already covered elsewhere.

## Done criteria

- [ ] No Codex streaming request calls global `fetch` directly.
- [ ] Codex uses the shared rescue fetch and watchdog.
- [ ] Codex wire bytes reach `StreamCallbacks.onActivity`.
- [ ] Only `response.completed` makes a normal SSE EOF successful.
- [ ] Cancellation and timeout remain distinguishable.
- [ ] Provider tests, typecheck, and formatting pass.
- [ ] No out-of-scope files changed.
- [ ] Plan 012 is marked DONE in `plans/README.md`.

## STOP conditions

- The shared helpers cannot support Codex without changing another provider's behavior.
- The live protocol is proven not to emit `response.completed` for successful
  streamed Responses; report the captured event _types only_, never payloads.
- Testing requires exposing credentials, URLs, or response bodies.
- A gate fails twice after a reasonable correction.

## Maintenance notes

Any future streaming provider should enter through `aiFetch`, the shared
watchdog, and `StreamCallbacks`. Reviewers should reject provider-local timeout
or rescue mechanisms unless the common contract is demonstrably insufficient.
