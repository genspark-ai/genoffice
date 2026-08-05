# ChatGPT Codex Preflight Sign-In Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use
> `superpowers:subagent-driven-development` (recommended) or
> `superpowers:executing-plans` to implement this plan task-by-task. Steps use
> checkbox (`- [ ]`) syntax for tracking.
>
> **Implementation instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving on. If a
> STOP condition occurs, stop and report instead of improvising. After the
> implementation and review pass, update this plan's status in
> `plans/README.md`.
>
> **Drift check (run first)**:
> `git diff --stat 4da673d..HEAD -- packages/ui/src/AiComposer.tsx apps/docs/src/renderer/ai/AiPanel.tsx apps/docs/src/renderer/i18n/strings-ai.ts apps/docs/src/renderer/styles.css apps/docs/tests/ai-panel-collapse.test.ts`
> The working tree already contains the uncommitted Codex provider and panel UI.
> Compare the Current state excerpts below against live code before editing.

**Goal:** Selecting `ChatGPT Codex` while signed out immediately shows a sign-in
banner and prevents every send path until browser OAuth succeeds.

**Architecture:** `AiPanel` owns renderer-safe Codex auth UI state and reads it
through the existing `aiCodexStatus()` IPC boundary. It gates capability loading
and all message starts on confirmed auth, while `AiComposer` accepts a generic
external send-disable signal shared safely with other apps. Existing main-process
OAuth, encrypted credential storage, provider transport, and Genspark behavior
remain unchanged.

**Tech Stack:** React 19, TypeScript, Electron IPC, Vitest/jsdom, localized
`defineStrings` dictionaries, app-scoped CSS.

## Status

- **Priority**: P1
- **Effort**: M
- **Risk**: MED — async auth state affects composer submission and capability loading
- **Depends on**: `plans/009-ai-panel-provider-composer-popovers.md`
- **Category**: bug
- **Planned at**: commit `4da673d`, 2026-08-05

## Global constraints

- Keep OAuth and credentials in the Electron main process. Tokens must never
  enter renderer state, settings, logs, history, or IPC results.
- Use only existing `aiCodexStatus()` and `aiCodexLogin()` renderer APIs. The
  latter already opens the system browser through main-process `shell.openExternal`.
- Keep Genspark and API-key provider behavior unchanged.
- Do not create a failed user/assistant chat turn merely to discover sign-out.
- Disable mouse and Enter-key submission until Codex auth is positively known
  to be logged in; also guard programmatic preset, Continue, and Retry paths.
- Keep the user's typed draft and attachments while signing in or switching
  providers.
- Use accurate localized copy in every existing language dictionary. Do not
  introduce English fallback text into non-English dictionaries.

---

## Why this matters

Current Codex UX discovers missing OAuth credentials only after a message is
submitted. That produces an undelivered user message, an assistant error, and
only then an inline `Sign in to ChatGPT` action. Provider selection already has
a safe status IPC, so sign-in can happen before any message or document action
starts.

## Current state

- `apps/docs/src/renderer/ai/AiPanel.tsx:72-94` stores auth recovery on chat
  entries. `loginActionForProvider()` always assigns a Codex login action after
  any Codex stream error, even when the failure is unrelated to auth.

  ```ts
  interface ChatEntry {
    // ...
    loginProvider?: 'genspark' | 'openai-codex'
  }

  export function loginActionForProvider(provider, loggedIn?) {
    if (provider === 'openai-codex') return 'openai-codex'
    if (provider === 'genspark' && loggedIn === false) return 'genspark'
    return undefined
  }
  ```

- `apps/docs/src/renderer/ai/AiPanel.tsx:238-279` calls
  `aiCodexCapabilities()` immediately when Codex is selected. Signed-out state
  is therefore inferred indirectly from a capability error instead of checked
  through `aiCodexStatus()`.
- `apps/docs/src/renderer/ai/AiPanel.tsx:527-580` adds Codex recovery UI only
  after `AgentLoop` reports an error. `apps/docs/src/renderer/ai/AiPanel.tsx:990-1003`
  renders that action inside the failed assistant message.
- `apps/docs/src/renderer/ai/AiPanel.tsx:647-664` starts runs without an auth
  precondition. This path is also used by preset auto-run, Continue, and Retry.
- `packages/ui/src/AiComposer.tsx:13-62` exposes `busy` but no external send
  gate. Its single `canSend` value controls both Enter and button submission:

  ```ts
  const canSend = value.trim().length > 0 && !busy
  ```

- `apps/docs/src/renderer/ai/AiPanel.tsx:884-915` places the header immediately
  before `.ai-chat`. The new banner belongs between those elements.
- `apps/docs/tests/ai-panel-collapse.test.ts:203-297` explicitly tests the old
  failure-first Codex recovery action. The same file already mounts the real
  panel, mocks `window.desktop`, types into the controlled textarea, and checks
  Codex controls.
- `apps/docs/src/main/docs-main.ts:2498-2512` already returns redacted
  `CodexAccountStatus` values for status/login. No main-process change is needed.
- Localization convention: `apps/docs/src/renderer/i18n/strings-ai.ts` requires
  every dictionary to match the Chinese key set. Styling convention: Docs AI
  surfaces use app-scoped `.ai-*` classes and existing color/radius variables in
  `apps/docs/src/renderer/styles.css`.

## Commands

| Purpose                 | Command                                                                                                                                                               | Expected on success                                            |
| ----------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------- |
| Focused red/green tests | `npm run test -w @genoffice/docs -- ai-panel-collapse.test.ts`                                                                                                        | New tests fail before implementation, then all file tests pass |
| Docs typecheck          | `npm run typecheck -w @genoffice/docs`                                                                                                                                | Exit 0, no TypeScript errors                                   |
| Shared UI typecheck     | `npm run typecheck -w @genoffice/ui`                                                                                                                                  | Exit 0, no TypeScript errors                                   |
| Scoped lint             | `npx eslint packages/ui/src/AiComposer.tsx apps/docs/src/renderer/ai/AiPanel.tsx apps/docs/src/renderer/i18n/strings-ai.ts apps/docs/tests/ai-panel-collapse.test.ts` | Exit 0, no ESLint errors                                       |
| Format check            | `npm run format:check`                                                                                                                                                | Exit 0, no formatting differences                              |
| Whitespace              | `git diff --check`                                                                                                                                                    | No output                                                      |

## Scope

**In scope** — only these production/test files may change:

- `packages/ui/src/AiComposer.tsx`
- `apps/docs/src/renderer/ai/AiPanel.tsx`
- `apps/docs/src/renderer/i18n/strings-ai.ts`
- `apps/docs/src/renderer/styles.css`
- `apps/docs/tests/ai-panel-collapse.test.ts`
- `plans/README.md` — status update only after implementation/review

**Out of scope**:

- `apps/docs/src/main/docs-main.ts`, preload, and shared IPC contracts; required
  status/login APIs already exist.
- `packages/ai-provider/**`; OAuth, token refresh, encrypted storage, and Codex
  transport do not change.
- Sheets, Slides, PDF, and their composer styling or auth behavior.
- Genspark's existing post-error sign-in action.
- Logout/account-management UI, token display, automatic browser launch on
  provider selection, or global auth state.
- Commit, push, merge, or PR creation unless the operator separately requests it.

## Git workflow

- Branch, if requested: `codex/codex-preflight-sign-in`.
- Match repo history with a short imperative English commit subject, for example
  `show Codex sign-in before sending`.
- Do not push or open a PR unless explicitly instructed.

---

### Task 1: Add failing preflight-auth UI tests

**Files:**

- Modify/test: `apps/docs/tests/ai-panel-collapse.test.ts:104-297`

**Interfaces:**

- Consumes: existing `window.desktop.aiCodexStatus(): Promise<CodexAccountStatus>`
  and `aiCodexLogin(): Promise<CodexAccountStatus>` mocks.
- Produces: regression contract for banner placement, OAuth invocation, send
  gating, successful recovery, failed recovery, and removal of Codex's inline
  post-error login action.

- [ ] **Step 1: Update authenticated Codex test fixtures**

  Every test that selects `openai-codex` for model/reasoning or streaming must
  mock `aiCodexStatus` as `{ loggedIn: true }`. Preserve existing capability and
  stream mocks. This makes those tests explicit about the precondition instead
  of accidentally depending on unknown auth state.

- [ ] **Step 2: Replace old helper/recovery tests with signed-out preflight tests**

  Remove the `loginActionForProvider` import and the test named `keeps Codex
sign-in recovery visible when stored status is stale`. Replace the old
  `renders the safe Codex diagnostic and recovery action` contract with tests
  using this state progression:

  ```ts
  const aiCodexStatus = vi.fn().mockResolvedValue({ loggedIn: false })
  const aiCodexLogin = vi.fn().mockResolvedValue({ loggedIn: true })
  const aiCodexCapabilities = vi.fn().mockResolvedValue({
    models: [{ id: 'gpt-5.5', reasoningEfforts: ['none', 'low', 'high'] }],
  })
  ```

  Assert all cases below with visible roles/classes rather than component state:

  1. Before `aiCodexStatus` resolves, typed input cannot enable Send and Enter
     cannot call `aiStream`.
  2. `{ loggedIn: false }` renders `.ai-codex-auth-banner` after
     `.ai-panel-header` and before `.ai-chat`.
  3. Banner contains localized required-sign-in copy and exactly one
     `Sign in to ChatGPT` button.
  4. Signed-out state does not call `aiCodexCapabilities` or `aiStream`.
  5. Clicking banner action calls `aiCodexLogin` once; while its promise is
     pending, button is disabled and exposes `aria-busy="true"`.
  6. Login success removes banner, calls capabilities, preserves typed draft,
     and enables Send.
  7. Login result `{ loggedIn: false, error: 'safe error' }` leaves banner and
     draft in place, displays `safe error`, and permits another login attempt.
  8. Switching away from Codex hides banner and restores ordinary provider send
     behavior. A login promise resolving after that switch must not restore a
     stale renderer status; reselecting Codex performs a fresh status check.
  9. A non-auth Codex stream error while status remains logged in renders the
     ordinary message error but no ChatGPT login button.

- [ ] **Step 3: Run focused tests and confirm RED**

  Run:

  ```bash
  npm run test -w @genoffice/docs -- ai-panel-collapse.test.ts
  ```

  Expected: new assertions fail because no preflight banner exists,
  capabilities load before status, and `AiComposer` cannot be externally gated.
  Fix only test setup errors until failures describe those missing behaviors.

---

### Task 2: Add one external send gate to shared `AiComposer`

**Files:**

- Modify: `packages/ui/src/AiComposer.tsx:13-62`
- Test through: `apps/docs/tests/ai-panel-collapse.test.ts`

**Interfaces:**

- Consumes: optional `sendDisabled?: boolean` from any app-level composer owner.
- Produces: `canSend === false` for both button clicks and Enter whenever
  `sendDisabled` is true. Existing callers omit it and retain current behavior.

- [ ] **Step 1: Add optional prop and fold it into existing gate**

  Use this exact public shape and calculation:

  ```ts
  readonly sendDisabled?: boolean | undefined

  const canSend = value.trim().length > 0 && !busy && !sendDisabled
  ```

  Destructure `sendDisabled = false`. Do not disable the textarea: users must be
  able to compose and retain a draft while signed out. Do not create a second
  keyboard condition; both Enter and the button must keep using `canSend`.

- [ ] **Step 2: Typecheck shared UI**

  Run:

  ```bash
  npm run typecheck -w @genoffice/ui
  ```

  Expected: exit 0. Existing Docs, Sheets, and PDF callers remain source
  compatible because the prop is optional.

---

### Task 3: Implement Codex auth state machine and run guard

**Files:**

- Modify: `apps/docs/src/renderer/ai/AiPanel.tsx:175-301`
- Modify: `apps/docs/src/renderer/ai/AiPanel.tsx:527-580`
- Modify: `apps/docs/src/renderer/ai/AiPanel.tsx:621-743`

**Interfaces:**

- Consumes: `CodexAccountStatus`, `window.desktop.aiCodexStatus()`, and
  `window.desktop.aiCodexLogin()`.
- Produces: panel-local `codexAccount`, `codexLoginPending`,
  `codexSendDisabled`, and a ref-backed guard used by every `runWith` caller.

- [ ] **Step 1: Model status without credential data**

  Import `CodexAccountStatus` from `../../shared/ipc`. Add panel state and a ref:

  ```ts
  const [codexAccount, setCodexAccount] = useState<CodexAccountStatus | null>(null)
  const [codexLoginPending, setCodexLoginPending] = useState(false)
  const codexAccountRef = useRef<CodexAccountStatus | null>(null)
  codexAccountRef.current = codexAccount

  const codexSendDisabled = settings.provider === 'openai-codex' && codexAccount?.loggedIn !== true
  ```

  `null` means status is unknown/checking. It must block sending but must not be
  treated as a confirmed signed-out warning.

- [ ] **Step 2: Check status whenever Codex becomes selected**

  Add an effect keyed by `settings.provider`. On Codex selection, set
  `codexAccount` to `null`, call `aiCodexStatus()`, then store its redacted
  result. Use an `active` flag so a late result cannot update after selection
  changes/unmounts. On an unexpected rejected IPC promise, store
  `{ loggedIn: false, error: tModule('aiUnknownError') }`. When switching away,
  clear both Codex account and login-pending state; never launch OAuth
  automatically. Clearing account prevents one render of stale `loggedIn: true`
  state when Codex is selected again.

- [ ] **Step 3: Gate capability loading on confirmed login**

  Change the existing capability effect precondition from provider-only to:

  ```ts
  if (settings.provider !== 'openai-codex' || codexAccount?.loggedIn !== true) {
    setCodexCapabilities(null)
    return
  }
  ```

  Include `codexAccount?.loggedIn` in its dependency list. Keep existing model
  validation and `onSettingsChange` behavior unchanged. This prevents
  `aiCodexCapabilities()` from becoming a second, failure-based auth probe.

- [ ] **Step 4: Guard every run path before mutating chat/history**

  At the start of `runWith`, before clearing input or persisting anything, add:

  ```ts
  if (
    settingsRef.current.provider === 'openai-codex' &&
    codexAccountRef.current?.loggedIn !== true
  ) {
    return
  }
  ```

  This covers Send, Enter, preset `autoRun`, Continue, Retry, and direct future
  callers. Signed-out attempts must not change input, attachments, chat,
  snapshots, history, busy state, or document content.

- [ ] **Step 5: Make banner OAuth action update auth state**

  Keep `startLogin` provider-aware for Genspark, but replace its Codex branch
  with this state transition:

  1. Ignore repeat calls while `codexLoginPending` is true.
  2. Set pending true, await `aiCodexLogin()`, and store its returned safe
     `CodexAccountStatus` directly.
  3. Preserve draft, attachments, chat, and settings.
  4. On rejected IPC, store a localized safe failure without exposing raw data.
  5. Before storing the resolved result, confirm
     `settingsRef.current.provider === 'openai-codex'`. If the user switched
     providers while the browser was open, ignore the renderer result; the
     credential store still completes and a later Codex selection rechecks it.
  6. Clear pending in `finally`.

  Do not use `attachNotice` for Codex sign-in results. Success is represented by
  banner removal; failure remains actionable inside the banner.

- [ ] **Step 6: Remove Codex failure-first recovery while preserving Genspark**

  Replace `ChatEntry.loginProvider` with the prior Genspark-only
  `loginRequired?: boolean`. Delete exported `loginActionForProvider()`.
  In `onError`:

  - For Codex, call `aiCodexStatus()` and update `codexAccount`. If refresh/token
    handling deleted invalid credentials, this changes the panel back to the
    signed-out banner and blocks later sends.
  - For Genspark, keep the current `aiGskStatus()` check and inline Genspark
    login action.
  - For every other provider, add no auth action.

  The failed message may retain its normal diagnostic and `undelivered` marker,
  but it must never contain a ChatGPT sign-in button.

- [ ] **Step 7: Pass gate to `AiComposer`**

  Add:

  ```tsx
  sendDisabled = { codexSendDisabled }
  ```

  Do not disable the textarea, attachments, provider switcher, model control,
  Track changes control, collapse control, or Stop button for an already-running
  request.

- [ ] **Step 8: Run focused tests**

  Run:

  ```bash
  npm run test -w @genoffice/docs -- ai-panel-collapse.test.ts
  ```

  Expected: auth state and send-gating tests pass. Banner tests may remain red
  until Task 4 adds markup/copy/styles.

---

### Task 4: Render and style signed-out banner

**Files:**

- Modify: `apps/docs/src/renderer/ai/AiPanel.tsx:884-915`
- Modify: `apps/docs/src/renderer/i18n/strings-ai.ts`
- Modify: `apps/docs/src/renderer/styles.css:3945-4056`

**Interfaces:**

- Consumes: confirmed `codexAccount.loggedIn === false`,
  `codexLoginPending`, `startLogin()`, existing `aiCodexLoginBtn`, and optional
  safe `codexAccount.error`.
- Produces: accessible `.ai-codex-auth-banner` placed between header and chat.

- [ ] **Step 1: Add localized required-sign-in copy**

  Add `aiCodexSignInRequired` to every dictionary in
  `strings-ai.ts`, preserving dictionary parity. English source copy:

  ```text
  Sign in to ChatGPT before using ChatGPT Codex.
  ```

  Translate that meaning accurately for all 19 existing locales. Reuse
  `aiCodexLoginBtn` for the action. Do not add a second success toast or duplicate
  the main-process error strings.

- [ ] **Step 2: Insert semantic banner between header and chat**

  Render only for confirmed signed-out state:

  ```tsx
  {settings.provider === 'openai-codex' && codexAccount?.loggedIn === false && (
    <div className="ai-codex-auth-banner" role="status">
      <div className="ai-codex-auth-copy">
        <span>{t('aiCodexSignInRequired')}</span>
        {codexAccount.error && <span className="ai-codex-auth-error">{codexAccount.error}</span>}
      </div>
      <button
        type="button"
        className="ai-codex-auth-login"
        disabled={codexLoginPending}
        aria-busy={codexLoginPending}
        onClick={() => void startLogin()}
      >
        {t('aiCodexLoginBtn')}
      </button>
    </div>
  )}
  <div ref={logRef} className="ai-chat" ...>
  ```

  Banner must be a sibling between `.ai-panel-header` and `.ai-chat`, not a chat
  message. This keeps conversation/history unchanged and banner fixed while chat
  scrolls.

- [ ] **Step 3: Style warning and action using existing tokens**

  Add scoped `.ai-codex-auth-*` styles near header/chat styles:

  - horizontal flex layout with wrapping at the panel's 280px minimum width;
  - warning-tinted background/border with readable light/dark-theme token colors;
  - compact text and error subtext;
  - clear primary action, hover, focus-visible, and disabled/pending states;
  - no absolute positioning and no overlap with `.ai-chat` or composer.

  Reuse existing CSS variables. Do not alter global banners or unrelated panel
  spacing.

- [ ] **Step 4: Run focused tests and typechecks**

  Run:

  ```bash
  npm run test -w @genoffice/docs -- ai-panel-collapse.test.ts
  npm run typecheck -w @genoffice/docs
  npm run typecheck -w @genoffice/ui
  ```

  Expected: every command exits 0; all focused tests pass.

---

### Task 5: Review removal, localization, and regression boundaries

**Files:**

- Review all in-scope files.
- Update: `plans/README.md` after all gates pass.

**Interfaces:**

- Consumes: completed Tasks 1-4.
- Produces: verified implementation with no Codex failure-first login action and
  no regression to other providers.

- [ ] **Step 1: Search for obsolete Codex recovery paths**

  Run:

  ```bash
  rg -n "loginActionForProvider|loginProvider" apps/docs/src/renderer/ai/AiPanel.tsx apps/docs/tests/ai-panel-collapse.test.ts
  ```

  Expected: no matches. `loginRequired` and `.ai-login-btn` may remain only for
  Genspark.

- [ ] **Step 2: Verify localization parity**

  Run:

  ```bash
  rg -n "aiCodexSignInRequired:" apps/docs/src/renderer/i18n/strings-ai.ts
  ```

  Expected: 19 matches, one per existing locale, with no placeholder text.

- [ ] **Step 3: Run final verification**

  Run:

  ```bash
  npm run test -w @genoffice/docs -- ai-panel-collapse.test.ts
  npm run typecheck -w @genoffice/docs
  npm run typecheck -w @genoffice/ui
  npx eslint packages/ui/src/AiComposer.tsx apps/docs/src/renderer/ai/AiPanel.tsx apps/docs/src/renderer/i18n/strings-ai.ts apps/docs/tests/ai-panel-collapse.test.ts
  npm run format:check
  git diff --check
  ```

  Expected: all commands exit 0; `git diff --check` prints nothing.

- [ ] **Step 4: Inspect scope and update plan status**

  Run `git status --short`. Confirm this implementation added changes only to
  the five in-scope production/test files, aside from pre-existing user changes.
  Read the full diff and verify each new hunk maps to this plan. Then update Plan
  010 from `TODO` to `DONE` in `plans/README.md`, recording any verification
  exception instead of silently weakening a gate.

## Test plan

- Extend `apps/docs/tests/ai-panel-collapse.test.ts`; use its existing
  `mount`, `panelProps`, `typeInto`, async `act`, and desktop mock patterns.
- Cover unknown/checking, signed out, OAuth pending, OAuth success, OAuth safe
  failure/retry, provider switch, mouse send, Enter send, capability gating,
  draft preservation, and non-auth stream errors.
- Keep tests deterministic: deferred promises model pending login; no real
  browser, network, timer, or credential store is used.
- Existing authenticated model/reasoning tests remain and gain explicit
  `aiCodexStatus` mocks.

## Done criteria

- [ ] Selecting ChatGPT Codex calls `aiCodexStatus()` without sending a message.
- [ ] Unknown/checking and signed-out Codex states disable both Send and Enter.
- [ ] Confirmed signed-out state shows warning banner between header and chat.
- [ ] Banner action calls `aiCodexLogin()` exactly once per click; Electron main
      process remains responsible for opening browser OAuth.
- [ ] OAuth pending state prevents duplicate clicks; success removes banner and
      enables capabilities/send; safe failure stays visible and retryable.
- [ ] Typed draft, attachments, chat, settings, and document state survive auth.
- [ ] Signed-out preset, Continue, Retry, and direct `runWith` paths cannot create
      chat/history/document side effects.
- [ ] Codex capability fetching happens only after confirmed login.
- [ ] Codex post-failure inline login action and helper are gone; Genspark's
      existing sign-in recovery still works.
- [ ] All 19 locale dictionaries contain translated required-sign-in copy.
- [ ] Focused tests, Docs/UI typechecks, scoped lint, format check, and whitespace
      check pass.
- [ ] No out-of-scope implementation file is modified.

## STOP conditions

Stop and report rather than improvising if:

- Live code no longer exposes `aiCodexStatus()` and `aiCodexLogin()` with
  `CodexAccountStatus` results through preload.
- Correct behavior would require OAuth tokens or raw provider errors in renderer
  state.
- `AiComposer` cannot gate Enter and button through one optional prop without
  changing existing Sheets/PDF behavior.
- Provider selection is moved outside `AiPanel`, making panel-local status
  ownership stale or duplicated.
- Correct implementation requires changing Codex transport, OAuth protocol,
  credential storage, or Genspark auth behavior.
- Any verification command fails twice after one focused correction.

## Maintenance notes

- Future account-backed providers should use an explicit status contract, not
  infer auth from failed model or message requests.
- Reviewers should scrutinize stale async results during rapid provider changes,
  the ref-backed `runWith` guard, and preservation of Genspark recovery.
- A future account/logout menu may lift auth state above `AiPanel`; that is
  intentionally deferred until another surface consumes the same state.
