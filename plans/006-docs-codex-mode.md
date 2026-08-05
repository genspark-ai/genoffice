# Plan 006: Enable Codex mode in GenOffice Docs

> Planning gate: stop after the deterministic fake-provider Docs run passes.
> Live-account validation belongs to Plan 007.

## Status

- Priority: P1
- Effort: M
- Depends on: `plans/003-codex-streaming-cancellation.md`,
  `plans/005-docs-auth-ipc.md`
- Risk: MED — provider selection and tool-loop integration
- Planned at: commit `4da673d`, 2026-08-04

## Objective

Make Codex selectable in Docs while preserving GenOffice’s complete harness:
`AgentLoop`, Docs prompts/history, `createDocsSkill`, `createFilesSkill`, and
local tool execution.

## Scope

- `apps/docs/src/main/docs-main.ts`
- `apps/docs/src/shared/ipc.ts`
- `apps/docs/src/preload/index.ts`
- `apps/docs/src/renderer/App.tsx`
- `apps/docs/src/renderer/ai/AiPanel.tsx`
- `apps/docs/src/renderer/ai/transport.ts` only if types require it
- `apps/docs/src/renderer/i18n/strings-ai.ts`
- relevant provider and Docs tests

## Steps

1. Remove the Docs-only forced Genspark selection and add safe Codex metadata
   while preserving legacy settings and API-key providers.
2. Make stream/chat handlers account-aware: Codex does not require an API key;
   Genspark and ordinary providers retain their current behavior.
3. Generalize login/error UX for the selected provider, including localized
   ChatGPT sign-in/status/logout actions without displaying raw auth errors.
4. Add a fake-provider test proving one Docs tool call executes locally, its
   result returns to the provider, and the final response reaches the panel.
5. Confirm the request includes only Docs/file tools and no Codex shell,
   filesystem, app-server, or built-in agent tools.

## Verify gate

Run `npm test -w @genoffice/docs`,
`npm run typecheck -w @genoffice/docs`, and the provider tests. Perform a fake
Codex Docs run covering tool use, cancellation, retry, document mutation, and
history persistence.

## STOP conditions

- `AgentLoop` or message/history semantics must change.
- The model receives tools outside the Docs/file skill set.
- Genspark or ordinary API-key paths regress.
