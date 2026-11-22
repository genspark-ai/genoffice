# Local-only fork: overview and plan

Status: Phase 1 (egress guard) done; Phases 2 to 4 not started (2026-11-22).

This fork of [genspark-ai/genoffice](https://github.com/genspark-ai/genoffice)
must never send documents, settings, credentials or usage data off the
machine. The code itself may stay public and CI may keep running on GitHub;
only the running app is restricted.

## Decisions

| Topic           | Decision                                                                                                                               |
| --------------- | -------------------------------------------------------------------------------------------------------------------------------------- |
| AI features     | Hide the AI panels and AI entry points behind one flag, so they can return later with a local model. Do not delete the code.           |
| Allowed hosts   | Loopback only: `localhost`, `127.0.0.1`, `::1`. No LAN addresses.                                                                      |
| Upstream        | Keep merging upstream. Unused network code is blocked at runtime and disabled at a few seams instead of deleted, so merges stay small. |
| GitHub workflow | Keep `.github/workflows/ci.yml`. It sees only the code, never user documents.                                                          |

## Build and run

Prerequisites: Node 22.12 or newer (`.nvmrc` says 22), npm 10 or newer, and a
Rust toolchain (`cargo` on `PATH`) for the Sheets xlsx engine. Run every
command from the repository root.

One-time setup (also after pulling upstream):

```bash
npm ci
```

```bash
npm run fixtures
```

Develop with live reload (all six editors on Vite ports 5173 to 5178, plus the
shell):

```bash
npm run dev
```

Build everything and run the built app unpacked:

```bash
npm run build:all
```

```bash
npx electron apps/shell
```

`npm run shell` does both steps in one. After changing code in
`apps/*/src/main`, rebuild the shell (`npm run build -w @genoffice/shell`), or
the change does not run.

An unpacked run keeps its profile in `GenOffice-Dev` under the OS app-data
folder (`~/Library/Application Support/GenOffice-Dev` on macOS), apart from any
installed GenOffice. To use a throwaway profile instead:

```bash
GENOFFICE_USER_DATA="$HOME/genoffice-scratch" npx electron apps/shell
```

Keep that path short: macOS refuses socket paths over 104 characters, and the
control server's socket lives in the profile folder.

Package an installer, with no update-feed or analytics variables set:

```bash
npm run dist:mac
```

The other targets are `npm run dist:win` and `npm run dist:linux`; the output
goes to `apps/shell/release` (override with `BUILD_DIR`). Without signing
credentials the installers are unsigned, which is expected for a local build.

Check the guard after a session; an empty or missing file means nothing tried
to leave the machine:

```bash
cat "$HOME/Library/Application Support/GenOffice-Dev/egress-blocked.log"
```

The packaged app writes to `GenOffice` instead of `GenOffice-Dev`.

Run the checks:

```bash
npm run lint && npm run typecheck && npm test
```

```bash
cd apps/shell && npx vitest run tests/egress-guard.test.ts
```

## Goal

- **Stays working:** opening, editing, saving and converting docx, xlsx, pptx,
  PDF, Markdown and HTML; local file search; the `genoffice` CLI.
- **Disabled:** auto-update, analytics, Genspark sign-in and search, cloud
  projects, web search, search reranking, image hosting, remote fonts and
  remote resources in HTML previews, every cloud AI provider.
- **Hidden:** the AI panels in every editor (code kept, flag off).
- **Rule:** the packaged app opens no non-loopback connection, from the main
  process, a renderer, an iframe or a child process it starts.

## Where network traffic can start

All outbound traffic starts in Electron main processes or in child processes
they spawn. The editor UIs are already locked down by their Content Security
Policy.

- **Renderers** (`apps/*/src/renderer`): every editor's `index.html` CSP limits
  `connect-src` to `'self'`, `ws://localhost:*` and in-app media schemes.
- **Shell main** (`apps/shell/src/main`): the one shipped Electron entry. It
  bundles each editor's main module and owns the updater, analytics, cloud
  projects, file-index reranking, the control server and the MCP server.
- **Editor main modules** (`apps/*/src/main`): run inside the shell process and
  call `net.fetch` for AI rescue fetches and slide fonts.
- **Engine packages** (`packages/ai-provider`, `packages/ai-search`): hold every
  AI and web-search endpoint, and spawn the `gsk` CLI and the Codex app
  server, which can reach the network on their own.
- **Format engines** (`docx-engine`, `pptx-*`, `pdf2docx`, converters): local
  code with no network use.

One guard in the shell main process therefore covers almost everything; child
processes are the gap it cannot see, so they are disabled at their call sites.

## Outbound connections found

A build from this fork already sends no analytics and checks no update feed,
because both need keys injected at package time (`apps/shell/electron-builder.cjs`).
Nine other paths can still reach the network.

| Path                            | Contacts                                             | Where                                                    | On in a fork build?                   | Action                   |
| ------------------------------- | ---------------------------------------------------- | -------------------------------------------------------- | ------------------------------------- | ------------------------ |
| GitHub star count               | `api.github.com/repos/genspark-ai/genoffice`         | `apps/shell/src/main/index.ts` (`fetchGithubStars`)      | Yes, when Settings > About opens      | Guard blocks; seam patch |
| Genspark sign-in, search, AI    | `www.genspark.ai` (or `GSK_BASE_URL`), `gsk` CLI     | `packages/ai-search/src/genoffice-auth.ts`, `gsk.ts`     | When signed in                        | Disable at seam          |
| Cloud projects list             | `www.genspark.ai` project API                        | `apps/shell/src/main/cloud-projects.ts`                  | When signed in                        | Disable at seam          |
| Cloud AI providers (about 40)   | OpenAI, Anthropic, Gemini, DeepSeek, OpenRouter, ... | `packages/ai-provider/src/providers.ts`, `registry.ts`   | When a key is entered                 | Loopback endpoints only  |
| Web search fallbacks            | DuckDuckGo, Parallel, Serper, Tavily, Exa, ...       | `packages/ai-search/src/index.ts`, `parallel-mcp.ts`     | Yes; DuckDuckGo and Parallel keyless  | Disable at seam          |
| Search reranking (TypeSafe Jev) | `api.typesafe.ai`, OpenRouter, Cloudflare            | `apps/shell/src/main/file-index/decider.ts`              | Opt-in                                | Disable at seam          |
| Markdown image upload           | `smms.app`, GitHub contents API                      | `apps/markdown/src/main/image-host.ts`                   | Opt-in                                | Disable at seam          |
| Slide font download             | Any URL the font store is given                      | `apps/slides/src/main/font-store.ts`                     | When a deck needs a font (to confirm) | Guard blocks             |
| HTML preview of user pages      | Whatever the page's scripts, images and links load   | `apps/html/src/renderer/preview/PreviewFrame.tsx`        | Yes, with remote resources            | Guard blocks (session)   |
| Auto-update                     | Feed from `GENOFFICE_UPDATE_URL`                     | `apps/shell/src/main/updater.ts`, `electron-builder.cjs` | No feed baked in                      | Guard blocks; leave code |
| Analytics (GA4)                 | `www.google-analytics.com`                           | `apps/shell/src/main/analytics.ts`                       | No keys                               | Guard blocks; leave code |
| Links you click                 | `genoffice.ai/join`, GitHub releases, help           | `index.ts`, `packages/electron-utils/src/github-menu.ts` | On click, in your browser             | Confirm before opening   |

The MCP server binds `127.0.0.1` and the control server uses a local socket;
neither is outbound. The CLI can serve to a network on request, which the
guard must refuse for non-loopback binds.

## Keeping upstream merges cheap

Deleting upstream code would make every merge conflict. Instead the fork
relies on three layers, so most upstream changes merge untouched:

1. **Deny by default at runtime.** The egress guard blocks every non-loopback
   connection, including ones upstream adds later. A new upstream feature that
   phones home simply fails closed; nothing needs to be noticed first.
2. **Few, marked seams.** Fork edits to upstream files are one-line gates that
   read a single flag module. Each carries a `// local-only:` comment so
   `git grep "local-only:"` lists the whole patch set, and conflicts are easy to
   spot. Everything else lives in new fork-only files that upstream never
   touches.
3. **Tripwire check.** `tools/check-no-egress.mjs` lists hard-coded
   `http(s)://` hosts in `apps/*/src` and `packages/*/src` that are not in the
   allowlist. Run on an upstream merge, it names every new endpoint so you can
   decide whether it needs a seam.

Fork-only files (upstream never has these):

- `apps/shell/src/main/local-only/egress-guard.ts`
- `apps/shell/src/main/local-only/allowlist.ts`
- `packages/ui/src/local-only.ts` (flags shared by every renderer)
- `tools/check-no-egress.mjs`
- `docs/local-only-fork.md` (this file)

Merge procedure:

```bash
git remote add upstream https://github.com/genspark-ai/genoffice.git
git config rerere.enabled true
git fetch upstream
git switch -c merge/upstream-$(date +%F) main
git merge upstream/main
npm run check:no-egress -- --base main
npm run typecheck && npm test
```

Then confirm every `// local-only:` seam still exists (`git grep -c "local-only:"`
matches the count recorded below) before merging the branch into `main`.

## Plan

Run `npm run lint`, `npm run typecheck` and `npm test` after each phase.

### Phase 1: egress guard (done)

Blocks everything before any feature is touched. Code lives in
`apps/shell/src/main/local-only/`; `install.ts` is the first import of
`apps/shell/src/main/index.ts` (the one `// local-only:` seam so far), so the
guard runs before any other module, editor main modules included.

| Layer           | What it does                                                                                                                                              |
| --------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Node sockets    | `net.Socket.prototype.connect` fails non-loopback hosts with `EGENOFFICE_OFFLINE` before DNS. Covers http, https, tls, ws and Node `fetch`.               |
| Node `fetch`    | `globalThis.fetch` rejects non-loopback URLs early with `OfflineBlockedError`.                                                                            |
| Node `listen`   | `net.Server.prototype.listen` binds an unspecified host to `127.0.0.1` and throws on a non-loopback host. Pipes and socket paths pass.                    |
| Chromium proxy  | Every session gets a black-hole proxy (`http://127.0.0.1:9`, loopback bypassed). Covers renderers, iframes, `net.fetch` and Chromium background requests. |
| Chromium filter | `webRequest.onBeforeRequest` on every session cancels non-loopback `http`, `https`, `ws`, `wss` and `ftp`.                                                |
| External links  | `shell.openExternal` asks before opening a public web link in the browser.                                                                                |
| Child processes | `HTTP_PROXY`, `HTTPS_PROXY`, `ALL_PROXY` point at the black hole; `NO_PROXY` lists loopback.                                                              |
| Log             | One JSON line per new (kind, host) in `userData/egress-blocked.log`: time, kind, host, caller frame. No URLs or content.                                  |

Verified: 14 unit tests in `apps/shell/tests/egress-guard.test.ts`; the full
shell suite passes; a real Electron 43 run blocked Node `fetch`, `net.fetch`
and a `BrowserWindow` load of a public page, and the proxy alone fails closed
(`ERR_PROXY_CONNECTION_FAILED`, no direct fallback).

Known gaps, left for later phases:

- UDP (`dgram`) and direct `dns` calls are not patched. Nothing in the tree
  is known to use them for egress.
- Child processes that ignore proxy variables (the `gsk` CLI, the Codex app
  server) can still connect; Phase 2 stops them from being spawned.
- The `genoffice` CLI runs as its own process and is not guarded; Phase 2
  covers its network paths.
- An editor started on its own (`npm run dev:docs`) does not load the shell
  guard; only the shell is shipped.
- The confirm dialog text is English only, to keep the patch out of the
  i18n shards.
- On Linux, Chromium cannot download spellcheck dictionaries any more.

### Phase 2: disable at the seams (2 days)

Gate, do not delete. Each edit is a one-line check of a flag from
`packages/ui/src/local-only.ts` (renderers) or the main-process allowlist,
marked `// local-only:`.

1. **Hide AI panels.** Add `AI_PANELS_ENABLED = false` to
   `packages/ui/src/local-only.ts`. Gate the panel mount and the
   `AiPanelSideButton` in each app: `apps/docs/src/renderer/ai/AiPanel.tsx`,
   `apps/sheets/src/renderer/ai/AiChatPanel.tsx`, and the `AiPanel.tsx` in
   `slides`, `pdf`, `markdown` and `html`. Hide AI entries in the Docs command
   table and ribbon through the existing `unavailable-entries.ts` mechanism,
   and hide Ask AI on the shell home screen.
2. **AI providers, loopback only.** In `packages/ai-provider/src/registry.ts`
   (or `ai-settings-guard.ts`), accept only custom OpenAI-compatible endpoints
   whose host is loopback (Ollama, LM Studio, llama.cpp). Cloud providers stay
   in the code but never appear in settings.
3. **Genspark and cloud projects.** Make `hasGskAuth()` return false and skip
   spawning the `gsk` CLI and Codex app server; return an empty snapshot from
   `cloud-projects.ts`.
4. **Web search.** Return no backends from `packages/ai-search/src/index.ts`,
   so agent skills register no web or image search tools.
5. **Small features.** Skip `fetchGithubStars`, the TypeSafe reranker in
   `file-index/decider.ts`, and the Markdown image hosts in `image-host.ts`;
   hide the GitHub and community menu items in
   `packages/electron-utils/src/github-menu.ts`.
6. **Updater and analytics.** Leave the code; it is already inert without the
   build-time keys, and the guard blocks it if keys ever appear. Add an assert
   in `electron-builder.cjs` that fails packaging when `GENOFFICE_UPDATE_URL`
   or GA4 variables are set.

Record the seam count here once done: `git grep -c "local-only:"` = _TBD_.

### Phase 3: harden the build (1 day)

1. Add `tools/check-no-egress.mjs` and an npm script `check:no-egress`; add it
   to CI next to the other `check:*` gates. Allowlist: OOXML and W3C schema
   URIs, loopback, documentation links in comments.
2. Change the app id, product name and userData folder in
   `electron-builder.cjs` and `apps/shell/src/main`, so the fork never shares
   settings or stored keys with an official GenOffice install.
3. On first launch, delete stored cloud credentials (Genspark, provider keys)
   found in the settings store.

### Phase 4: verify and package (half a day)

Run the checklist below, then `npm run dist:mac` (or `dist:win`, `dist:linux`)
with no `GENOFFICE_UPDATE_URL` or GA4 variables set.

## Verification

The build passes only when a full session shows zero non-loopback
connections, measured outside the app.

- [ ] Egress-guard unit tests: `fetch`, `net.fetch`, `https.request` and a
      renderer request to a public host fail; `127.0.0.1` succeeds.
- [ ] `npm run check:no-egress` passes.
- [ ] Packaged app under a host firewall in deny mode (LuLu or Little Snitch on
      macOS, OpenSnitch on Linux): open, edit and save one docx, xlsx, pptx,
      PDF, Markdown and HTML file, run a file search, open Settings. No
      connection attempt appears.
- [ ] Same session with the network unplugged: nothing hangs or errors.
- [ ] `egress-blocked.log` is empty after that session.
- [ ] `npm test`, `npm run typecheck`, `npm run lint` and the existing
      `check:*` gates stay green.

The firewall run is the real proof: it also catches child processes and
native code that the in-app guard cannot see.
