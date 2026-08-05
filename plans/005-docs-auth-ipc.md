# Plan 005: Expose the authentication boundary through Docs IPC

> Planning gate: stop after typed IPC and redaction tests pass. Do not select
> Codex as the active Docs provider yet.

## Status

- Priority: P1
- Effort: S
- Depends on: `plans/004-codex-oauth-credentials.md`
- Risk: HIGH — Electron boundary and secret exposure
- Planned at: commit `4da673d`, 2026-08-04

## Objective

Expose only account status, login initiation, cancellation where supported, and
logout to Docs. The renderer must never receive access tokens, refresh tokens,
authorization codes, account headers, or raw provider diagnostics.

## Scope

- `apps/docs/src/main/docs-main.ts`
- `apps/docs/src/preload/index.ts`
- `apps/docs/src/shared/ipc.ts`
- existing Docs IPC/preload tests

## Steps

1. Add typed account IPC methods beside the existing Genspark methods.
2. Open the login URL through the existing safe external-link path.
3. Validate all renderer inputs in main and return redacted status/errors only.
4. Add bridge-shape, redaction, logout, and login-failure tests.

## Verify gate

Run `npm run typecheck -w @genoffice/docs`,
`npm test -w @genoffice/docs`, and inspect renderer-visible IPC payloads for
secret absence.

## STOP conditions

- Any secret crosses the preload bridge.
- OAuth state or callback handling is moved into renderer code.
- Existing Genspark authentication IPC regresses.
