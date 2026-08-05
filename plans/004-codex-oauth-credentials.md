# Plan 004: Implement GenOffice-owned Codex OAuth

> Planning gate: stop after credential-service tests pass. No Docs UI or IPC
> wiring is included in this gate.

## Status

- Priority: P1
- Effort: M
- Depends on: `plans/001-codex-provider-contract.md`
- Risk: HIGH — credential storage and refresh
- Planned at: commit `4da673d`, 2026-08-04

## Objective

Implement provider-owned OAuth login, refresh, account identification, secure
storage, and logout in the main-process/provider boundary. Do not read or copy
`~/.codex/auth.json`; the CLI is not a runtime dependency of the mod.

## Scope

- `packages/ai-provider/src/auth.ts` or equivalent
- `packages/ai-provider/src/index.ts`
- `packages/ai-provider/tests/auth.test.ts` (create)

## Steps

1. Implement PKCE/device login, callback/state validation, token exchange,
   expiry detection, refresh, logout, and redacted account status.
2. Store credentials in a protected main-process credential store; never in
   `AiSettings`, logs, snapshots, renderer state, or IPC payloads.
3. Deduplicate concurrent refreshes and delete invalid credentials after a
   confirmed refresh failure.
4. Use injectable HTTP, clock, browser, and storage dependencies for tests.

## Verify gate

Run provider typecheck/tests. Test state mismatch, timeout, expiry, refresh,
logout, non-200 responses, redaction, and concurrent refresh. Search the repo
for token-field references and confirm none are renderer-facing.

## STOP conditions

- Tokens cannot be protected or deleted.
- Login requires renderer-held credentials.
- OAuth behavior cannot be tested with mocked issuer/transport boundaries.
