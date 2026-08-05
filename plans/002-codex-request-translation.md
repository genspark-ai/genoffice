# Plan 002: Translate GenOffice requests to Codex Responses

> Planning gate: stop after mocked request-body tests pass. Do not add OAuth or
> Docs integration yet.

## Status

- Priority: P1
- Effort: M
- Depends on: `plans/001-codex-provider-contract.md`
- Risk: MED — protocol and tool-history compatibility
- Planned at: commit `4da673d`, 2026-08-04

## Objective

Implement the provider-only request adapter modeled on Pi/OpenCode’s Codex
Responses transport. GenOffice remains responsible for prompts, history, and
tool execution.

## Scope

- `packages/ai-provider/src/codex.ts` (create)
- `packages/ai-provider/src/stream.ts` only for routing if necessary
- `packages/ai-provider/tests/codex.test.ts` (create)

## Steps

1. Convert GenOffice system instructions and user/assistant/tool messages to the
   Codex Responses input representation.
2. Convert existing `AgentToolDef` schemas to function tools without adding
   Codex-owned tools or changing schemas.
3. Preserve supported image data URLs and disable remote conversation storage
   unless required by the tested protocol.
4. Assert mocked request URLs, headers, instructions, messages, tools, model,
   and session fields. Use only fake opaque credentials.

## Verify gate

Run `npm run typecheck -w @genoffice/ai-provider` and
`npm test -w @genoffice/ai-provider`. Confirm the adapter has no Electron,
renderer, filesystem, or tool-execution imports.

## STOP conditions

- Existing GenOffice tool-call history cannot be represented.
- The provider requires a second harness or built-in Codex tools.
- License or bundling constraints prevent isolating the provider layer.
