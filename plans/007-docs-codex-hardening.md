# Plan 007: Harden, validate, and package the Docs mod

> Final gate: this is the only plan that requires live subscription testing.
> Record pass/fail results without recording account, token, prompt, or document
> content.

## Status

- Priority: P1
- Effort: M
- Depends on: `plans/006-docs-codex-mode.md`
- Risk: MED — live endpoint and packaging behavior
- Planned at: commit `4da673d`, 2026-08-04
- Implementation status: BLOCKED — live ChatGPT subscription/account unavailable in this environment; live checklist not executed.

## Objective

Validate the personal mod against a real account, harden operational failures,
confirm the Docs tool boundary, and ensure the packaged app does not depend on
the installed Codex CLI at runtime.

## Scope

- Codex provider/auth tests and bounded error handling
- Docs files touched by Plan 006
- `apps/docs/package.json` and `package-lock.json` only if packaging requires it
- concise setup/security notes in `README.md`, `CONTRIBUTING.md`, or `SECURITY.md`

## Steps

1. Normalize auth expiry, refresh failure, rate-limit, unsupported-model,
   malformed-SSE, timeout, and network errors into safe user-facing messages.
2. With a real account, test text, one Docs tool turn, multi-turn continuation,
   cancellation, retry, and an image attachment if supported.
3. Inspect renderer globals, settings, IPC chunks, logs, and errors for secret
   leakage; verify only Docs/file tools are sent.
4. Build and package Docs; verify startup without a session fails gracefully and
   no global Codex CLI is required.
5. Document the personal-mod assumptions and known dependence on the Codex
   backend transport.

## Verify gate

Run `npm run format:check`, `npm run lint`, `npm run typecheck`, `npm test`,
`npm run licenses`, `npm run build -w @genoffice/docs`, and
`npm run dist:dir -w @genoffice/docs`. Complete the redacted live-account
checklist and inspect the final diff.

## STOP conditions

- Live behavior requires weakening the Docs tool boundary.
- Packaging introduces renderer-visible credentials or a CLI runtime dependency.
- The endpoint requires undocumented workarounds that cannot be tested safely.
