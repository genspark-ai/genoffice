# Plan 003: Add Codex streaming and cancellation

> Planning gate: stop after the provider emits the existing stream contract and
> all focused tests pass.

## Status

- Priority: P1
- Effort: M
- Depends on: `plans/002-codex-request-translation.md`
- Risk: MED — event ordering, partial tool calls, and abort behavior
- Planned at: commit `4da673d`, 2026-08-04

## Objective

Translate Codex Responses SSE events into GenOffice’s existing chunks/tool-call
callbacks. The provider must never execute tools.

## Scope

- `packages/ai-provider/src/codex.ts`
- `packages/ai-provider/src/stream.ts` routing
- `packages/ai-provider/tests/codex.test.ts`
- `packages/ai-provider/src/index.ts` if exports require adjustment

## Steps

1. Map text deltas, completed function-call arguments, completion, and provider
   errors to existing `AiStreamChunk` semantics.
2. Accumulate fragmented and parallel tool arguments, including malformed JSON
   and bounded diagnostic errors.
3. Wire abort signals so cancellation closes the fetch stream and does not leave
   request state or refresh promises alive.
4. Keep ordinary OpenAI, Genspark, and other provider routes unchanged.

## Verify gate

Run `npm test -w @genoffice/ai-provider`,
`npm test -w @genoffice/agent-core`, and
`npm run format:check`. Tests must cover text-only output, tools, 401/429/5xx,
malformed events, and cancellation.

## STOP conditions

- Events cannot map to the current stream/tool-call contract.
- Cancellation causes late tool execution or stale callbacks.
- Existing provider or agent-core tests regress.
