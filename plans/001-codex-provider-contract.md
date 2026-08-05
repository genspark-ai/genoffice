# Plan 001: Define the Codex provider contract

> Planning gate: stop after this plan and inspect the public types and metadata
> before implementing network behavior.

## Status

- Priority: P1
- Effort: S
- Depends on: none
- Risk: LOW
- Planned at: commit `4da673d`, 2026-08-04

## Objective

Introduce the distinct account-backed Codex provider identity and the internal
request/auth types needed by later plans. This gate must not make network calls,
add OAuth, touch Electron, or change `AgentLoop`.

## Scope

- `packages/ai-provider/src/types.ts`
- `packages/ai-provider/src/providers.ts`
- `packages/ai-provider/src/index.ts`
- provider metadata/settings tests

## Steps

1. Add a provider ID such as `openai-codex` and metadata that does not require
   an API key or pretend that ChatGPT OAuth is an OpenAI API key.
2. Define an injected, in-memory auth context containing only the values the
   provider layer needs. Keep it out of `AiSettings` and renderer-facing types.
3. Define the narrow adapter boundary for instructions, messages, tools, model,
   abort signal, and stream callbacks.
4. Add metadata and settings-resolution tests, including unknown-provider
   rejection and legacy settings behavior.

## Verify gate

Run `npm run typecheck -w @genoffice/ai-provider` and
`npm test -w @genoffice/ai-provider`. Confirm there are no Electron, renderer,
Pi agent, or Codex app-server imports.

## STOP conditions

- A credential field must be added to `AiSettings`.
- The provider contract requires changing `AgentLoop`.
- The selected dependency pulls in a coding-agent harness rather than provider
  protocol code.
