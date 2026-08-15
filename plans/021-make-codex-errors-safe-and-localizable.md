# Plan 021: Make Codex failures safe and localizable at the shared boundary

> **Implementation instructions**: Follow this plan step by step. Run every
> verification command and stop on any STOP condition. Prefer a small stable
> code union over app-specific message parsing. Update `plans/README.md` after
> implementation and review.
>
> **Drift check (run first)**:
> `git diff --stat 1878b30..HEAD -- packages/ai-provider packages/agent-core/src/electron-transport.ts packages/agent-core/tests/electron-transport.test.ts apps/docs/src/main apps/docs/src/renderer/ai apps/docs/src/renderer/i18n/strings-ai.ts apps/docs/tests`
> The audited code is uncommitted at this SHA. Compare the live code with
> Current state; behavioral drift is a STOP condition.

## Status

- **Priority**: P1
- **Effort**: M
- **Risk**: MED
- **Depends on**: `plans/020-extract-codex-auth-node-adapter.md`
- **Category**: security
- **Planned at**: commit `1878b30`, 2026-08-15

## Why this matters

The stream embeds provider payloads and tool arguments in raw `Error.message`
values, then relies on a Docs-only regex sanitizer before IPC. Another app's
existing handler could forward them raw. The path also returns hardcoded
English despite GenOffice's i18n rule. Safe codes at the shared boundary make
every integration secure by default and leave user copy in the renderer.

## Current state

- `packages/ai-provider/src/codex.ts:248-257` includes tool arguments in an
  error; lines 298-302 include malformed SSE; lines 348-351 throw raw provider
  failure text.
- `apps/docs/src/main/codex-error.ts` converts errors to English with regexes.
  `docs-main.ts` sends those strings, plus English model/reasoning errors.
- `packages/ai-provider/src/auth.ts:8-9,23-28` puts English status text in
  `CodexAccountStatus.error`.
- `packages/agent-core/src/electron-transport.ts` already maps stable `timeout`
  and `credits` codes to renderer-localized text. Extend this seam; do not add a
  second transport.
- `apps/docs/src/renderer/i18n/strings-ai.ts` contains every locale's AI keys.
- `CONTRIBUTING.md` requires user-facing copy in renderer i18n or main dictionaries.

## Commands you will need

| Purpose         | Command                              | Expected on success |
| --------------- | ------------------------------------ | ------------------- |
| Provider tests  | `npm test -w @genoffice/ai-provider` | exit 0              |
| Transport tests | `npm test -w @genoffice/agent-core`  | exit 0              |
| Docs tests      | `npm test -w @genoffice/docs`        | exit 0              |
| Typecheck       | `npm run typecheck`                  | exit 0              |
| Build           | `npm run build -w @genoffice/docs`   | exit 0              |

## Scope

**In scope**:

- `packages/ai-provider/src/auth.ts`
- `packages/ai-provider/src/codex.ts`
- `packages/ai-provider/src/types.ts`
- `packages/ai-provider/src/index.ts`
- `packages/ai-provider/tests/auth.test.ts`
- `packages/ai-provider/tests/codex.test.ts`
- `packages/agent-core/src/electron-transport.ts`
- `packages/agent-core/tests/electron-transport.test.ts`
- `apps/docs/src/main/codex-auth-main.ts`
- `apps/docs/src/main/codex-error.ts` (delete if unused)
- `apps/docs/src/main/docs-main.ts`
- `apps/docs/src/renderer/ai/transport.ts`
- `apps/docs/src/renderer/ai/AiPanel.tsx`
- `apps/docs/src/renderer/i18n/strings-ai.ts`
- affected tests under `apps/docs/tests/`
- `plans/README.md` (status only)

**Out of scope**:

- Non-Codex provider behavior
- Raw response bodies, prompts, tool arguments, tokens, or URLs in errors/logs
- New telemetry/error dependencies
- UI redesign or other app wiring
- One translation key per HTTP status; group by actionable user outcome

## Git workflow

- Work on the current branch; do not push/open a PR.
- If asked to commit: `localize safe Codex failures`.

## Steps

### Step 1: Define a minimal safe Codex error taxonomy

Define/export one small stable error-code union for actionable cases:
authentication required/expired, temporary auth/network failure, timeout,
unavailable model/capabilities, rate limit, rejected request, invalid stream,
invalid tool call, and generic provider failure. Reuse `timeout` where it
already exists. Do not encode raw messages/bodies in renderer-facing data.

Have provider paths throw/return a typed error carrying a code and optional HTTP
status/safe diagnostic code only. Remove SSE payloads, tool JSON, and provider
failure text from thrown messages. Keep enough typed structure for deterministic
classification without regexing provider content.

Change `CodexAccountStatus` to carry a code instead of English `error`. Update
tests for confirmed-expired versus transient refresh failure.

**Verify**: `npm test -w @genoffice/ai-provider` -> all pass, including
assertions that raw secret/payload snippets are absent.

### Step 2: Route codes through the existing IPC transport

Extend `AiStreamChunk` and matching `IpcStreamChunk` with the codes. Generalize
`createIpcTransport` with one optional error-code resolver callback; retain
existing timeout/credits behavior so sibling apps need no changes.

Add tests proving a known Codex code resolves locally and unknown/missing codes
fall back to `unknownErrorText()` without showing a raw main-process message.

**Verify**: `npm test -w @genoffice/agent-core` -> all pass.

### Step 3: Localize Docs at the renderer boundary

Replace Docs regex sanitization and hardcoded account/model/reasoning messages
with codes. Map codes in `apps/docs/src/renderer/ai/transport.ts` and the
account/capability UI using `strings-ai.ts` keys.

Add the smallest useful keys to every locale. Reuse existing localized generic
or timeout copy where a separate action is unnecessary. Remove the unused
setter-only `codexCapabilitiesLoading` state at `AiPanel.tsx:312`.

Delete `codex-error.ts` and its tests if no caller remains; otherwise reduce it
to typed-code classification with no user copy or message regexes.

**Verify**:

```sh
npm test -w @genoffice/docs
npm run build -w @genoffice/docs
rg -n "argumentsJson\.slice|payload\.slice|event\.error\?\.message|safeCodexStreamError" packages/ai-provider apps/docs/src/main
```

Expected: tests/build exit 0; search returns no unsafe/obsolete matches.

### Step 4: Run shared gates

```sh
npm run typecheck
npm run lint
npm run format:check
git diff --check
```

Expected: all exit 0; lint may keep pre-existing warnings but has zero errors.

## Test plan

- Provider errors exclude tool JSON, SSE payloads, response bodies, tokens,
  URLs, and raw provider text.
- HTTP/auth/timeout/protocol/tool categories yield deterministic codes.
- Transient refresh keeps credentials; confirmed expiry removes them.
- Renderer resolves known codes via i18n and localizes generic fallback.
- Existing non-Codex timeout/credits behavior remains unchanged.

## Done criteria

- [ ] Shared Codex paths expose codes, not English/raw payloads.
- [ ] Docs renders all Codex errors through i18n.
- [ ] Another app can consume errors safely without a Docs helper.
- [ ] Setter-only loading state is removed.
- [ ] Provider, transport, Docs tests, checks, and build pass.

## STOP conditions

- A code would need raw provider/request content to be useful.
- This requires changing every non-Codex provider's behavior.
- Locale dictionary shape prevents following the existing convention.
- A verification fails twice after a reasonable in-scope correction.

## Maintenance notes

Keep the taxonomy user-action-oriented. Add a code only when the UI can offer
meaningfully different guidance. Diagnostics stay bounded and typed, never in
renderer messages.
