# GenOffice Mail P5 — AI Phase Spec (skills, scoping, IPC, ownership)

Status: SPEC ONLY, no behavior change. Lanes build verbatim from sections
2-11; any deviation is a SPEC-DELTA event owned by Lina.

Provenance: base `c2feb03f` (`fix/mail-p4-resync-purge`: QA GREEN stack
T `8f290511` + Y `a5a446c0` + O `4c466a4b` + BUG-1 fix `c2feb03f`),
branch `feat/mail-p5-ai-spec`, local-only (no upstream PR per standing
instruction; fakes/synthetic mail only, no real creds).
Normative inputs: plan `genoffice-mail-plan.md` section 3 (AI integration)

- roadmap P4 acceptance; PII rules from `genoffice-mail-security-spec.md`
  section 4 (normative — violation fails sec review).

Base-tree note: on first touch the `t_41aadd5d` checkout showed a transient
`M apps/mail/src/main/store/store.ts` plus mass-deletion flapping; the
workspace `.git` metadata and then the worktree were being destroyed by an
unknown concurrent actor. No uncommitted BUG-1 duplicate existed — the
BUG-1 fix (`purgeFolder` alias, `store.ts:706-711`) is committed in
`c2feb03f`. The base for this spec was re-materialized from rescued loose
objects + main-repo history, checked out at `c2feb03f` with a CLEAN tree
(`git status` empty, `git fsck` clean). Lanes build from THIS branch, never
from `t_41aadd5d`.

Discipline (all lanes): English-only comments (`check-english-comments`
gate); license allowlist — NO new prod deps, pure TS on
`@genoffice/agent-core` + `@genoffice/ai-provider` (both UNCHANGED, no
edits under `packages/`); local-dev only; renderer CSS uses semantic
tokens only (`tools/check-theme-colors.mjs`).

---

## 1. Skill set (composed, packages UNCHANGED)

Four mail skills, all implementing `AgentSkill`
(`packages/agent-core/src/skill.ts:15-44`: `id`, `systemPrompt`, `tools`,
optional `buildContext()`, `executeTool(call, signal?)`, optional
`verifyResponse()`), merged with `composeSkills`
(`packages/agent-core/src/skill.ts:46`) and driven by `AgentLoop`
(`packages/agent-core/src/loop.ts:201`, options incl. `transport`
at `loop.ts:54-55`, `DEFAULT_MAX_TURNS` at `loop.ts:82`).
Streaming rides `streamForProvider`
(`packages/ai-provider/src/stream.ts:18-24`: provider-routed streaming
tool-calling turn; BYOK key or signed-in Genspark proxy, never any other
recipient). One-shot fallback is `chatForProvider`
(`packages/ai-provider/src/chat.ts:10-16`). Transport type `AgentTransport`
lives in `packages/agent-core/src/types.ts` (re-exported via
`packages/agent-core/src/index.ts:10`); Electron IPC construction is
`createIpcTransport` (`packages/agent-core/src/electron-transport.ts:73`).

Markdown precedent (mirror, do not fork): `apps/markdown/src/renderer/ai/`
— `AiPanel.tsx` wiring (`AgentLoop` at `:359`, `transport` at `:360`,
`composeSkills('markdown+search', ...)` at `:361`,
`createMarkdownSkill` at `:362`, `createSearchSkill` at `:373`,
`createElectronTransport` at `:270-271`); `transport.ts:6`
`createElectronTransport(getSettings)` over `window.markdownApi`
(`onAiStream`/`aiStream`/`aiStreamCancel`); `tools.ts:114` `AGENT_TOOLS`

- `tools.ts:437` `executeTool`; `markdown-skill.ts:69`
  `createMarkdownSkill`, `search-skill.ts:9` `createSearchSkill`.

Mail starts from `apps/mail/src/renderer/ai/` (currently `AiStub.tsx` +
`ai.css` stub only). New files (K lane, section 10):

- `triage-skill.ts` — `createTriageSkill()`. Context: headers
  (from/to/subject/date) + folder + snippet + labels/flags ONLY, via
  `triage-headers` retrieval (section 4). Tools: `rank_folder`
  (digest buckets: priority / needs-reply / newsletter / noise),
  `apply_label` (user-confirmed ONLY — never auto-move, plan section 3).
  `buildContext()` attaches the folder header list; bodies are NOT
  attached, NOT fetchable from this skill.
- `summarize-skill.ts` — `createSummarizeSkill()`. Context: selected
  thread/folder bodies, fetched ONLY on explicit user action via
  `fetch-bodies` (section 4). Tools: `summarize_thread`,
  `catch_up_folder`. Every summary bullet carries a message-id citation.
- `compose-assist-skill.ts` — `createComposeAssistSkill()`. Context:
  draft text + cited thread context the user opened. Tools:
  `draft_from_bullets`, `rewrite_tone`, `reply_from_thread`. Output is
  TEXT returned to the renderer for insertion (section 5) — the skill
  never sends, never touches outbox IPC, never harvests addresses beyond
  recipients the user typed.
- `search-skill.ts` (mail flavor) — `createMailSearchSkill()`. Tools over
  `triage-headers` + `fetch-bodies`: `search_mail` (FTS, headers-first),
  `quote_message` (per-claim message-id citation payload). No
  cross-account joins: every tool call echoes the request scope and the
  skill refuses results whose account is outside it.
- `mail-tools.ts` — `MAIL_AGENT_TOOLS` + `executeMailTool` (mirror of
  markdown `tools.ts:114`/`tools.ts:437`), calling ONLY
  `window.mailApi.aiRetrieve` (section 6). Tool names globally unique
  across the four skills (`composeSkills` requires it).

`verifyResponse` (skill.ts:33-43) is REQUIRED on ask-mail/summarize:
reject any final text containing a factual claim without a
`[message-id]` citation, forcing one corrective turn.

## 2. PII matrix (normative; violation fails sec review)

Verbatim from security spec section 4:

| Context                      | MAY see                                                       | MUST NEVER see                                         |
| ---------------------------- | ------------------------------------------------------------- | ------------------------------------------------------ |
| Triage (default)             | Headers (from/to/subject/date), folder, snippet, labels/flags | Bodies, attachments, other accounts' mail              |
| Summarize / catch-up         | Selected thread/folder bodies, on explicit user action        | Mail outside the selected account scope                |
| Compose-assist               | Draft text + cited thread context the user opened             | Address-book harvesting beyond recipients user typed   |
| Ask-mail                     | Retrieved messages WITH per-claim message-id citations        | Cross-account joins; un-cited claims presented as fact |
| Logs / telemetry / crash     | Allowlisted fields only (accountId hash, op outcome, latency) | Bodies, snippets, subjects, addresses, tokens, blobs   |
| Skill traces / eval fixtures | Synthetic mail only                                           | Real user mail in checked-in fixtures                  |

Enforcement points: `MailEnvelope` is headers+snippet by construction —
`toEnvelope` (`apps/mail/src/main/store/store.ts`, `toEnvelope` fn)
maps rows to ref/uid/threadId/from/to/cc/subject/date/flags/labels/
snippet/hasAttach and NEVER body text, so triage tool output cannot leak
bodies even on a renderer bug. Bodies travel ONLY through
`getMessageBody` (`store.ts:576`) behind `fetch-bodies`. Eval fixtures
use `mail-fake-data.ts` (main) and `devFakes.ts` (renderer) synthetic
mail; skill traces never touch real accounts.

## 3. Account-scoping envelope (every AI call)

The picker-selected account set travels with EVERY AI call and is
re-checked where results render:

```ts
interface AiScope {
  accountIds: string[]
} // picker-selected, non-empty, max 8
```

- Renderer: `AiPanel` holds the scope from the account picker; every
  `aiRetrieve` call carries it; rendered results re-check
  `result.scope` against the live picker (stale-scope results are
  discarded, same pattern as `AiStub`'s `data-account`
  in `AiStub.tsx`).
- Main (`mail:ai-retrieve` handler, W lane): rejects unknown accountIds;
  opens ONLY the per-account SQLite files for scope members
  (`<accountId>.sqlite`, per-account files per `store.ts` header) —
  cross-account joins are impossible by construction; the in-store guard
  `fts(query, accountIds)` (`store.ts:730-733`, refuses callers outside
  its own account) stays as second layer.
- Bodies fetched on demand for the cited set only — never
  whole-mailbox-to-context (security spec rule 3).
- AI calls go ONLY to the user's configured provider (BYOK key or
  Genspark proxy via existing `AI_CHANNELS.stream`); no training use,
  no secondary recipients (security spec rule 2). Local FTS stays LOCAL
  (security spec rule 5).

## 4. Retrieval API (Farah/Nadia joint — exact shape)

One new invoke channel (section 6); three ops. Triage runs on headers +
FTS; bodies lazily.

```ts
// renderer -> main (W lane implements; K lane consumes via mail-tools.ts)
type AiRetrieveOp = 'triage-headers' | 'fetch-bodies' | 'cite-resolve'
interface AiRetrieveRequest {
  scope: AiScope
  op: AiRetrieveOp
  query?: string // triage-headers: FTS query, headers+snippet ranked
  folderId?: string
  limit?: number // default 50 (FTS_DEFAULT_LIMIT, store.ts:53)
  refs?: MailMessageRef[] // fetch-bodies / cite-resolve: explicit user-cited set
}
interface AiRetrieveResponse {
  scope: AiScope // echo — renderer discards on mismatch
  envelopes: MailEnvelope[] // headers+snippet ONLY (toEnvelope shape); bodies NEVER here
  bodies?: { ref: MailMessageRef; text: string }[] // fetch-bodies only, cited set only
  citations: { uid: string; folderId: string; msgIdHeader?: string }[]
}
```

Main-side mapping (W lane, imports store/sync APIs, never edits them):

- `triage-headers` → `SqliteMailStore.searchLocal(query, { limit })`
  (`store.ts:735-738`: `buildFtsMatch` quoting at `store.ts:256` — user
  text NEVER reaches MATCH raw; bm25-ranked FTS plus escaped-LIKE
  substring fallback for CJK/mid-token queries at `store.ts:749+`).
  Envelopes only.
- `fetch-bodies` → `getMessageBody(accountId, folderId, uid)`
  (`store.ts:576`) per ref, ONLY for refs in scope;Snippet/bodyText
  inline cap `INLINE_BODY_CAP_BYTES` (`store.ts`) respected.
- `cite-resolve` → header-only resolve of refs to citable ids (no bodies).
- Dev/fakes parity: `FakeMailApi` path (current `mail-read-handlers.ts`
  serves fakes at `:134-150`) stays for renderer dev + stubbed-provider
  e2e; real path above is used in production. W lane does NOT change
  existing `fetch-body`/`search-local` handlers.
- Tokenizer lock: FTS5 `unicode61` PLAIN (`migrations.ts:73`; the
  `remove_diacritics=2` wording in older spec text does not match the
  Node-bundled SQLite build — see SPEC-DELTA D2). `node:sqlite` verdict
  stands: `sqlite.ts:15` is the ONLY `node:sqlite` importer in the repo
  (`sqlite.ts:2-4`).

## 5. Compose-assist insertion contract (AI proposes, human sends)

Output lands in the compose box as EDITABLE text; the human sends via the
existing send path. Insertion point: `ComposeHost.tsx` — `bodyHtml` state
at `:118`, uncontrolled `contentEditable` div at `:338-341`
(`onInput` → `setBodyHtml(innerHTML)`; deliberately NOT
`dangerouslySetInnerHTML`-per-keystroke at `:120-122`).

- UI lane adds ONE optional prop to `ComposeHost`
  (e.g. `assistText: string | null` + `onAssistConsumed()`): an effect
  appends the assist HTML via `setBodyHtml(current => current + html)`
  then clears — typing caret behavior from `34d0a468` is preserved.
- K lane returns plain assist text from `compose-assist-skill` tools;
  renderer (UI lane) inserts it. K lane NEVER touches `ComposeHost`,
  compose IPC, or outbox. No auto-send, no draft overwrite without the
  user's keystrokes staying intact.

## 6. MAIL_CHANNELS IPC additions (typed + zod end-to-end)

Sheets precedent (`apps/sheets/src/shared/ipc-channels.ts` +
`desktop-api.ts`): renderer bridge stays zod-free with hand-rolled shape
guards (mail already follows this — `mail-bridge.ts:10-15`), main
handlers re-validate with zod schemas from `shared/ipc-schemas.ts`
(existing: `mailRefSchema` `:17`, `mailSearchLocalSchema` `:162`,
`mailSaveDraftSchema` `:200`, `mailSyncNowSchema` `:216`,
`mailOutboxListSchema` `:237`; channels table `MAIL_CHANNELS`
in `shared/ipc.ts:21-48`; `MailChannel` type at `:55`).

Exactly ONE new channel (W lane):

- `mail:ai-retrieve` — `MAIL_CHANNELS.aiRetrieve = 'mail:ai-retrieve'`
  (`shared/ipc.ts`, appended next to `:44-48` sync block).
- `mailAiRetrieveSchema` (`shared/ipc-schemas.ts`, appended):
  `{ scope: z.object({ accountIds: z.array(accountIdSchema).min(1).max(8) }).strict(), op: z.enum(['triage-headers','fetch-bodies','cite-resolve']), query: z.string().max(500).optional(), folderId: z.string().max(256).optional(), limit: z.number().int().min(1).max(200).optional(), refs: z.array(mailMessageRefSchema).max(50).optional() }.strict()`
  with refinements: `fetch-bodies`/`cite-resolve` REQUIRE `refs`;
  `triage-headers` REQUIRES `query`.
- Streaming reuses the GENERIC channels (shell-owned, UNCHANGED):
  `AI_CHANNELS.stream = 'ai:stream'`, `streamChunk`, `streamCancel`
  (`shared/ipc.ts:178-187`), already bridged for mail at
  `mail-bridge.ts:174-181` (`aiGskStatus`/`aiStream`/`aiStreamCancel`/
  `onAiStream`) and stubbed in `devFakes.ts:416-432`. No new
  `mail:ai-stream*` channels — one retrieve channel keeps the IPC
  surface minimal (threat T5).

## 7. Redaction hook — DECISION: extend `sync/redact.ts`

DECISION (locked): EXTEND `apps/mail/src/main/sync/redact.ts`. New file
rejected — this module is already the single choke point on all sync
log/error paths with `sanitizeError` (`redact.ts:16-31`: one-line cap,
EMAIL/TOKEN/LONG_B64 → `[address]`/`[secret]`/`[blob]`) and the
`containsPii` test helper (`redact.ts:33-44`), covered by
`sync-redact.test.ts`. (Plan-named `agent-core/redact-core.ts` does NOT
exist — verified, see module header.)

W lane adds (same file, same owner Y→W handoff documented in section 10):

- `redactForPrompt(text: string): string` — EMAIL/TOKEN/B64 rules PLUS
  `Received:` / `Authentication-Results:` header blocks, IPv4/IPv6
  literals, `-----BEGIN *-` key material → `[redacted]`; applied to
  EVERYTHING crossing into logs AND prompts (retrieval context assembly
  in `mail:ai-retrieve`, skill-trace strings).
- `stripTransportHeaders(raw: string): string` — drops the full
  Received-chain + auth internals before any header block enters a
  prompt (threat T9: metadata leakage).
- Unit tests with mail fixtures in `sync-redact.test.ts` style
  (hostile-mail-shaped inputs: multi-hop Received, tokens, key blocks).

## 8. Renderer plan (UI lane)

- New `AiPanel.tsx` in `apps/mail/src/renderer/ai/`, mirroring markdown
  `AiPanel.tsx` wiring (`AgentLoop` `:359` + `transport` `:360` +
  `composeSkills([...])` `:361`): `createMailTransport` (same shape as
  markdown `transport.ts:6`, over `window.mailApi.onAiStream/aiStream/
aiStreamCancel`), `composeSkills('mail-ai', intro, [triage,
summarize, composeAssist, mailSearch])`, `AgentLoop<MailSnapshot>`.
- DELETE `AiStub.tsx` (stub rail served its purpose; its 4 disabled
  buttons + `data-account` scoping note are superseded by the panel).
  Mount swap in `App.tsx`: replace `<AiStub accountId={...} />`
  (`App.tsx:261`, sidebar under `FolderTree`) with
  `<AiPanel scope={selection accountIds} onInsertAssist={...}
onOpenCitation={openMessage} />`.
- Citation-click opens the message in the reader: citations call
  `openMessage(envelope)` (`App.tsx:137-150`: `setSelectedUid` + mark
  seen), the same callback `MessageList` uses via `onSelect`
  (`App.tsx:284`, `MessageList.tsx:27-98`); reader renders via
  `MessageReader` (`MessageReader.tsx:38`).
- `ai.css`: keep existing `.mail-ai` / `.mail-ai-actions` classes
  (token-only: `var(--border)`, `var(--surface)`, `var(--text-muted)`);
  ADD panel/streaming/error/citation classes — semantic tokens ONLY
  (theme gate).
- i18n: `mail.ai.*` in `apps/mail/src/renderer/i18n/strings.ts`
  (SINGLE-FILE, all locales inline — follow the existing file, NOT the
  shard pattern; SPEC-DELTA D4). Existing 6 keys (`strings.ts:89-94`
  en, `:185-190` zh + 17 more locales). ADD 7 keys, zh FIRST then all
  18 siblings: `mail.ai.title`, `mail.ai.placeholder`,
  `mail.ai.send`, `mail.ai.stop`, `mail.ai.error`,
  `mail.ai.emptyResult`, `mail.ai.cite`.
- `devFakes.ts`: ADDITIVE AI stubs only — extend the existing
  `getAiSettings`/`aiStream` stubs (`devFakes.ts:416-432`) with an
  `aiRetrieve` fake serving synthetic triage/bodies/citations. No other
  fake changes.

## 9. Mount-port note (UI lane step one — EXTERNAL INPUT)

Shell mount `703c0f16` (`t_1dd0c7f9`: OutboxBanner/Panel/
OfflineIndicator mount) is a SIBLING of `c2feb03f` off `4c466a4b` and is
NOT in this stack — verified absent: `703c0f16` resolves in NO
`genoffice` checkout on this machine (hunt across all
`kanban/workspaces/*/genoffice/.git` negative).

UI lane step one: locate the `t_1dd0c7f9` branch/workspace holding
`703c0f16`, port its 5 renderer-only files:
`OutboxBanner.tsx`, `OutboxPanel.tsx`, `OfflineIndicator.tsx`,
`OutboxApi.ts`, `outbox.css` (all under
`apps/mail/src/renderer/components/`, cf. current same-named files from
the `4c466a4b` stack) + the mount snippet (App shell sidebar/banner area
adjacent to `App.tsx:261`). If the source workspace is unrecoverable,
FALLBACK: rebuild the mount from this stack's existing components
(same 5 files exist at `c2feb03f`) and record the rebuild as SPEC-DELTA.
Either way UI lane records the final file list + mount location in its
completion metadata. No main/ipc/preload changes in the port
(renderer-only).

## 10. File-ownership lock (collision-free — no shared-file edits)

- **W/provider lane** owns: `apps/mail/src/main/ai/*` (new; retrieval
  context assembly + `mail:ai-retrieve` handler — IMPORTS store/sync
  APIs, never edits them) + the ONE appended AI block in
  `src/main/*ipc*.ts` handler registration + `src/shared/ipc.ts` ONE
  appended `aiRetrieve` entry + `src/shared/ipc-schemas.ts` ADDITIVE
  `mailAiRetrieveSchema` + `src/preload/mail-bridge.ts` ADDITIVE
  `aiRetrieve` method + its tests (`ai-retrieve.test.ts`,
  redaction-prompt tests). Touches NOTHING else.
- **K/skills lane** owns: `apps/mail/src/renderer/ai/*-skill.ts`
  (4 new files) + `mail-tools.ts` (new) + its tests
  (`ai-skills.test.ts`, two-account leak fixtures). Consumes W's
  `window.mailApi.aiRetrieve`; NEVER touches main/ipc/preload/devFakes.
- **UI lane** owns: `renderer/ai/AiPanel.tsx` (new) + `AiStub.tsx`
  (delete) + `ai.css` (extend) + `App.tsx` mount-line edits ONLY
  (`:9` import, `:261` swap) + `ComposeHost.tsx` ONE additive optional
  prop + `devFakes.ts` ADDITIVE AI stubs + `i18n/strings.ts` ADDITIVE
  7 `mail.ai.*` keys × 19 locales + its tests
  (`ai-panel.test.tsx`, citations-open-messages). No other shared-file
  edits. Mount-port (section 9) first.
- **Docs lane** owns: `apps/mail/docs/*` follow-ups + README/help-lang
  updates. No code.
- Shared files (`ipc.ts`, `ipc-schemas.ts`, `mail-bridge.ts`,
  `strings.ts`, `devFakes.ts`, `App.tsx`, `ComposeHost.tsx`) are
  APPEND/ADDITIVE-ONLY per above; lanes never edit each other's blocks.
  `packages/*` UNTOUCHED by all lanes.

## 11. Test plan (all lanes; P4 regression stays green)

- Stubbed-provider AI e2e (K+UI): `AgentLoop` over fake transport +
  `FakeMailApi.aiRetrieve`; triage→summarize→compose→ask flows assert
  scope echo + citations present.
- Two-account leak fixtures (K+W, sec-gating): accounts A/B with
  disjoint synthetic mail; assert NO B-envelope in A-scoped triage,
  NO B-body in A-scoped summarize, `fts()` guard unit test, stale-scope
  render discard test. Any leak = lane FAIL.
- Redaction unit tests (W): mail fixtures incl. multi-hop Received,
  bearer tokens, key blocks, IPs; assert `containsPii(redacted) ===
false` and header blocks gone from prompt strings.
- Citations-open-messages (UI): cited claim click → `openMessage` →
  reader shows cited uid (`App.tsx:137` path).
- P4 regression: full mail suite (existing 204+ tests incl.
  `store-sqlite`, `sync-*`, `mail-*`) stays green; `typecheck`,
  `eslint`, theme-colors, english-comments gates green.

## 12. SPEC-DELTA vs plan section 3 (build lanes carry verbatim)

- D1 `node:sqlite` verdict STANDS (`sqlite.ts:15`; sole importer rule
  at `sqlite.ts:2-4`). No better-sqlite3, no second driver.
- D2 FTS5 tokenizer is `unicode61` PLAIN (`migrations.ts:73`); older
  `remove_diacritics=2` wording does not match the Node-bundled build
  (noted at `migrations.ts:6-7`). CJK/mid-token recall comes from the
  escaped-LIKE fallback (`store.ts:749+`); triage queries MUST go
  through `buildFtsMatch` (`store.ts:256`).
- D3 `agent-core/redact-core.ts` does NOT exist — redaction extends
  `apps/mail/src/main/sync/redact.ts` (section 7), not a new file, not
  agent-core.
- D4 Mail i18n is SINGLE-FILE `strings.ts` (all 19 locales inline),
  not sharded — UI lane follows the existing file, zh first.
- D5 Current `fetch-body`/`search-local` handlers serve FAKES
  (`mail-read-handlers.ts:134-150`); P5 retrieval (section 4) targets
  the real `SqliteMailStore` with fakes parity, and does NOT modify
  those handlers.
- D6 Naming: plan roadmap calls AI "P4-AI" and hardening "P5"; this
  execution runs AI as P5-SPEC with P4 = sync/offline (done). Phase
  acceptance for THIS card: stubbed-provider AI e2e + redaction tests
  - no-leak tests green (plan P4-AI acceptance, same bar, new number).
- D7 Streaming reuses generic `ai:stream*` channels (section 6); there
  are NO `mail:ai-stream*` channels. The sole new channel is
  `mail:ai-retrieve`.
- D8 Base recovery (provenance header): spec branch re-materialized at
  `c2feb03f` clean; lanes must NOT use the `t_41aadd5d` workspace path
  (destroyed during this spec) — build lanes clone from THIS branch.
