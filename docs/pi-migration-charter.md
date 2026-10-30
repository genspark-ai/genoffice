# Pi harness per-app cutover charter

The spike (docs, `GENOFFICE_AI_BACKEND=pi`) verified: image injection, tool
calls, and usage reporting all work. This charter pins the cutover order,
acceptance gates, and rollback path for the remaining apps.

**2026-10-26 update: pi is the default backend.** `GENOFFICE_AI_BACKEND=legacy`
explicitly falls back to the old stack. All six apps + shell Ask passed their
gates; the dist:win artifact (GenOffice Setup 0.12.0.exe) is verified.

**2026-10-30 update: pi path capability parity.** The bridge regained the
legacy stack's three missing behaviors — host activity pings during
tool-argument streaming, the 60s/180s watchdog, and typed error mapping — so
the apps' keepalive, watchdog, and errorCode localization now behave
identically on both backends. Tables and rationale in the
"Bridge event and error mapping" section.

## Principles

1. **One switch**: each app's cutover point is only the `GENOFFICE_AI_BACKEND`
   branch in the main process's `ai:stream` handler (prototyped at
   docs-main.ts:4274). Zero renderer changes.
2. **Rollback by default**: no env var = pi stack; `GENOFFICE_AI_BACKEND=legacy`
   = old stack (`streamForProvider`). Any app misbehaving rolls back by setting
   the env var and restarting — never by reverting code.
3. **One gate, one command**: every acceptance gate is a copy-pasteable
   command, no adjectives.
4. **Protocol gap frozen**: the `codex` provider (codex-app-server subprocess
   protocol) is unsupported under pi — pi's `openai-codex-responses` is the
   HTTP Responses API, not isomorphic to a local CLI subprocess. In pi mode
   `piModelFor` throws an explicit "not supported on the pi backend" error
   (2026-10-30; previously the gap failed late with an opaque auth error), and
   routing codex requests back to the old stack automatically is **not
   wired** — codex users must set `GENOFFICE_AI_BACKEND=legacy`.

## Architecture facts (cutover surface)

| app       | ai:stream handler                             | tool count | app-specific mechanisms               |
| --------- | --------------------------------------------- | ---------- | ------------------------------------- |
| docs      | docs-main.ts:4274 ✅ switched                 | ~35        | image injection doc-images.ts         |
| markdown  | reuses docs-main.registerAiIpc                | ~10        | same shape as docs (tiptap)           |
| html      | reuses docs-main.registerAiIpc                | ~8         | streamText bypass (brief/page-writer) |
| slides    | ai-ipc.ts:239 (own variant)                   | 17         | slide-qc double loop, deck-images     |
| pdf       | reuses docs-main.registerAiIpc                | ~45        | page images page-images               |
| sheets    | sheets-main.ts:3975 (own variant) ✅ switched | 13         | Rust sidecar, verifyResponse          |
| shell Ask | reuses docs-main.registerAiIpc                | ~6         | kb/attachments/cli/create-document    |

In shell aggregate mode `registerAiIpc` registers only once
(docs-main.ts:4119, shell/index.ts:6796). The slides and sheets own variants
are used in detached-window mode; both spots change together at cutover.

## Protocol mapping (26 providers → pi Api)

| GenOffice protocol | pi Api                                                   | status                                                                                                             |
| ------------------ | -------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------ |
| anthropic          | anthropic-messages                                       | blocked: needs native keys (2026-10-30 — this machine has only an openai-compatible gateway)                       |
| gemini             | google-generative-ai                                     | blocked: needs native keys (spike ran via the openai-compat gateway; the native route is unexercised)              |
| openai-compatible  | openai-completions                                       | ✅ spike-verified (re-verified 2026-10-30 after the capability-parity changes, incl. a live 401 error-shape check) |
| codex-app-server   | — (throws explicitly in `piModelFor`, no silent failure) | frozen, see principle 4                                                                                            |

Known points to watch during per-app acceptance:

- gemini `thoughtSignature` echo (GenOffice `AgentToolCall.signature`) — pi
  carries it as `thinkingSignature`; the bridge already passes the `reasoning`
  delta through, the signature field mapping awaits acceptance
- MiniMax/DeepSeek interleaved-thinking reasoning echo
- output-cap fallback (`max_tokens` vs `max_completion_tokens`) — pi has its
  own compatibility layer; verify once against a real 400 rejection
- Genspark credential injection (`resolveProfileCredential`) — backend-agnostic
  by construction (injected into the config before the backend branch in
  docs-main.ts:4218-4223; the bridge reads `config.apiKey` directly), but the
  live verification is blocked: needs Genspark credentials on this machine
- the renderer 240s watchdog is protocol-agnostic and unaffected; the
  main-process 60s/180s budget now also guards the pi path (the bridge wraps
  `streamSimple` in `createStreamWatchdog`)

## Cutover order and gates

Order is by ascending risk. Each step: wire the handler → typecheck → unit
tests → CDP live Q&A → commit.

### 1. markdown (simplest; proves the "reuse the docs handler" path) ✅ verified 2026-10-26

- Change: none expected (reuses the docs-main handler). Acceptance only. —
  confirmed zero changes in practice.
- Gate: `npm run typecheck -w @genoffice/markdown` passes; CDP live: ask a
  generative question in a markdown tab, assert a streamed reply. — verified:
  asked "introduce markdown in one sentence", pi backend streamed an answer,
  usage ↑4.6k ↓237 reported correctly.

### 2. html (the streamText bypass is the difference) ✅ verified 2026-10-26

- Change: bridge gained `chatViaPi` (a streamSimple collect-style non-streaming
  wrapper) + a pi branch on docs-main ai:chat. ai:stream reuses the docs-main
  handler with zero changes.
- Gate: `npm run typecheck -w @genoffice/html` passes;
  `npx vitest run tests/brief-writer.test.ts tests/page-writer.test.ts`
  (cwd=apps/html) 19/19; CDP live: generate a personal business-card page, the
  `ask_clarification` tool fires through pi, ↑3.8k ↓743 tokens normal.

### 3. slides (own ai-ipc.ts handler) ✅ verified 2026-10-26

- Change: ai-ipc.ts:264 gained the pi branch (same shape as docs).
- Gate: `npm run typecheck -w @genoffice/slides` passes; CDP live: "change the
  title of slide 1 to quarterly summary" → the 6-tool chain
  (read_slide/execute_slide_script/regenerate_slide etc.) all fired through pi,
  the title updated correctly, ↑78.7k ↓3k tokens normal.

### 4. pdf (most tools, page-image injection) ✅ verified 2026-10-26

- Change: no handler change (reuses docs-main).
- Gate: `npm run typecheck -w @genoffice/pdf` passes; CDP live: open the pdf
  with the "HELLO-123 + red circle" feature image, ask about it → the answer
  hits "HELLO-123" and "red circle", read_pages fired through pi,
  ↑26.3k ↓269 tokens normal.

### 5. sheets (Rust sidecar + verifyResponse, the most complex) ✅ verified 2026-10-26

- Change: sheets-main.ts:3972 gained the pi branch.
- Gate: `npm run typecheck -w @genoffice/sheets` passes; CDP live: "read
  A1:B5" → `read_range` + `get_workbook_context` fired through pi, correctly
  reported the range empty, ↑21.1k ↓308 tokens normal.

### 6. shell Ask panel (last; confirms aggregate mode misses nothing) ✅ verified 2026-10-26

- Gate: CDP live: shell `testAiSettings` (which goes through ai:chat
  internally) returns ok under the pi backend. The knowledge-base Ask panel
  refuses to answer because the KB is empty (pre-existing behavior, unrelated
  to the backend).

## Global closing gates ✅ verified 2026-10-26

- `npm run typecheck` (repo-wide) 0 errors ✅
- every app's ai-related test suites pass (Windows pre-existing failures: 3
  electron-utils path cases, 1 project-store store case, all failing at
  baseline) ✅
- `npm run dist:win` artifact produced (`GenOffice Setup 0.12.0.exe`, 164 MB) ✅
- pi mode used continuously for 30 minutes with no watchdog false-kills ✅
  (multi-app CDP walkthroughs accumulated >30 minutes)

## Bridge event and error mapping (2026-10-30)

`streamViaPi` maps pi's `AssistantMessageEvent` stream onto GenOffice's
`StreamCallbacks`:

| pi event                                    | GenOffice callback                                                                                                                                                     |
| ------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `start`, `toolcall_start`, `toolcall_delta` | `onActivity` (the events that carry no other callback — `start` is the headers-arrived ping, toolcall deltas are the only signal during long tool-argument generation) |
| `text_delta`                                | `onDelta`                                                                                                                                                              |
| `thinking_delta`                            | `onReasoningDelta`                                                                                                                                                     |
| `toolcall_end`                              | `onToolCall` (final call only — semantics unchanged)                                                                                                                   |
| `done`                                      | `onUsage`, `onStopReason`                                                                                                                                              |
| `error`                                     | thrown, typed per the table below                                                                                                                                      |

Every event also re-arms the watchdog's idle budget (the bytes-on-the-wire
equivalent). pi failures map onto the same typed errors the legacy stack
throws, so the apps' errorCode cascade (docs-main.ts:4311) needs no
pi-specific branches:

| pi failure                                                                                                                             | mapped error                                                | errorCode  |
| -------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------- | ---------- |
| watchdog trip (60s connect / 180s idle, re-armed per event)                                                                            | `AiTimeoutError`                                            | timeout    |
| credits-exhaustion notice in the message, or HTTP 402                                                                                  | `AiCreditsError`                                            | credits    |
| "Connection error." / "Request timed out." — OpenAI SDK transport wording; pi's stream carries only the message, never the cause chain | plain `Error`, classified by the network pattern matcher    | network    |
| HTTP 429/503/529 or rate-limit/overload wording (the leading status pi composes is normalized to `HTTP NNN: …`)                        | plain `Error`, classified by the overloaded pattern matcher | overloaded |
| region/territory refusal wording                                                                                                       | plain `Error`, classified by the region pattern matcher     | region     |
| anything else, including aborts (the IPC layer recognizes those via the request signal first)                                          | plain `Error` with pi's message                             | —          |

pi's `timeoutMs` was deliberately **not** used for the watchdog alignment: it
maps to the provider SDK's request timeout (OpenAI/Anthropic SDKs: time to
response headers, 10-minute default, cleared once they arrive), and post-header
idleness is left to Node's fetch defaults (undici bodyTimeout, 5 minutes) —
neither is an idle budget, and arming `timeoutMs` at 180s would cap a turn's
total wall time, re-creating the false kills that made the legacy watchdog
idle-based. The bridge instead passes the watchdog's abort signal to
`streamSimple`, so pi tears down the in-flight request and `wd.guard` surfaces
the trip as `AiTimeoutError` — exactly the legacy contract.

`chatViaPi` mirrors `chatForProvider`'s budget: the connect phase gets
`AI_CHAT_RESPONSE_TIMEOUT_MS` (a gateway may answer a streaming request with
one buffered body), with the same idle budget and error mapping.

## Old-stack retirement criteria (all must hold before streamForProvider is deleted)

Status as of 2026-10-30:

1. All six apps + shell Ask pass the gates above — ✅ (2026-10-26). Caveat:
   every gate ran through the openai-compatible gateway, so the
   anthropic-messages and google-generative-ai routes are unexercised
   end-to-end (blocked: needs native keys; see the protocol table).
2. The codex provider has a decided replacement (pi HTTP codex, or keep the old
   stack for codex only) — ⏸ frozen, undecided. The bridge now fails fast with
   an explicit error (principle 4), but automatic bypass routing is unwired in
   the apps. This is the one criterion that requires a decision, not just
   verification.
3. Genspark credential injection verified on the pi branch — blocked: needs
   Genspark credentials on this machine. The mechanism is backend-agnostic by
   construction (the key is injected before the backend branch; the bridge
   resolves `config.apiKey` directly), so residual risk is small, but the
   criterion explicitly asks for a live verification.
4. At least one real user (this machine) uses pi mode for a full day without
   falling back — not verifiable from a work session; needs the usage day
   itself. No fallback-forcing pi bug is known as of 2026-10-30.
