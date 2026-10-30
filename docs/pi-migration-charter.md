# Pi harness per-app cutover charter

The spike (docs, `GENOFFICE_AI_BACKEND=pi`) verified: image injection, tool
calls, and usage reporting all work. This charter pins the cutover order,
acceptance gates, and rollback path for the remaining apps.

**2026-10-26 update: pi is the default backend.** `GENOFFICE_AI_BACKEND=legacy`
explicitly falls back to the old stack. All six apps + shell Ask passed their
gates; the dist:win artifact (GenOffice Setup 0.12.0.exe) is verified.

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
   HTTP Responses API, not isomorphic to a local CLI subprocess. In pi mode the
   codex endpoint bypasses through the old stack (the bridge throws an explicit
   error, no silent failure).

## Architecture facts (cutover surface)

| app | ai:stream handler | tool count | app-specific mechanisms |
|---|---|---|---|
| docs | docs-main.ts:4274 ✅ switched | ~35 | image injection doc-images.ts |
| markdown | reuses docs-main.registerAiIpc | ~10 | same shape as docs (tiptap) |
| html | reuses docs-main.registerAiIpc | ~8 | streamText bypass (brief/page-writer) |
| slides | ai-ipc.ts:239 (own variant) | 17 | slide-qc double loop, deck-images |
| pdf | reuses docs-main.registerAiIpc | ~45 | page images page-images |
| sheets | sheets-main.ts:3975 (own variant) ✅ switched | 13 | Rust sidecar, verifyResponse |
| shell Ask | reuses docs-main.registerAiIpc | ~6 | kb/attachments/cli/create-document |

In shell aggregate mode `registerAiIpc` registers only once
(docs-main.ts:4119, shell/index.ts:6796). The slides and sheets own variants
are used in detached-window mode; both spots change together at cutover.

## Protocol mapping (26 providers → pi Api)

| GenOffice protocol | pi Api | status |
|---|---|---|
| anthropic | anthropic-messages | not spike-tested (no key) |
| gemini | google-generative-ai | not spike-tested (ran via openai-compat gateway) |
| openai-compatible | openai-completions | ✅ spike-verified |
| codex-app-server | — (unsupported, throws and bypasses) | frozen, see principle 4 |

Known points to watch during per-app acceptance:
- gemini `thoughtSignature` echo (GenOffice `AgentToolCall.signature`) — pi
  carries it as `thinkingSignature`; the bridge already passes the `reasoning`
  delta through, the signature field mapping awaits acceptance
- MiniMax/DeepSeek interleaved-thinking reasoning echo
- output-cap fallback (`max_tokens` vs `max_completion_tokens`) — pi has its
  own compatibility layer; verify once against a real 400 rejection
- Genspark credential injection (`resolveProfileCredential`) must also work on
  the pi branch
- the double watchdogs (main 180s / renderer 240s) are protocol-agnostic and
  unaffected

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

## Old-stack retirement criteria (all must hold before streamForProvider is deleted)

1. All six apps + shell Ask pass the gates above
2. The codex provider has a decided replacement (pi HTTP codex, or keep the old
   stack for codex only)
3. Genspark credential injection verified on the pi branch
4. At least one real user (this machine) uses pi mode for a full day without
   falling back
