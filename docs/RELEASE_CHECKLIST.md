# Release checklist

Use this before publishing a KĀRYA build. It mirrors what Phase 8/9 verified
for the first fork release and applies to every subsequent one.

## Code health

- [ ] `npm run typecheck` — 0 errors across all workspaces
- [ ] `npm test` — full monorepo suite green (17/17 packages)
- [ ] `npm run lint` (or `eslint .`) — 0 errors
- [ ] `npm run fixtures` ran before the test suite where required

## Build

- [ ] `npm run build:all` — all five editor modules + shell render cleanly
- [ ] xlsx sidecar exists for every packaged platform:
      `apps/sheets/native/xlsx-engine/target/release/xlsx-sidecar(.exe)`
      and, for the Windows shell installer,
      `target/x86_64-pc-windows-gnu/release/xlsx-sidecar.exe`
- [ ] `npm run notices` regenerates `apps/shell/build/THIRD-PARTY-NOTICES.txt`
- [ ] `node tools/gen-karya-icons.mjs` produced the current icon set
      (only needed when branding changes)

## Windows installer

- [ ] `npm run dist:win` completes and produces
      `apps/shell/release/KĀRYA Setup <version>.exe` (NSIS, x64)
- [ ] Install on a clean VM (no Node/rust tooling) — app launches, editors
      open from the shell tabs, `.xlsx` opens (sidecar present under
      `Resources/native/`)
- [ ] Desktop shortcut name is **KĀRYA**; window title reads KĀRYA
- [ ] Install → launch → close → relaunch → uninstall leaves documents and
      `%APPDATA%\KĀRYA` intact
- [ ] First run after an existing GenOffice install migrates user data
- [ ] Paths with spaces and non-ASCII user names work

## AI providers

- [ ] Ollama running: status chip shows **Connected**, models listed
- [ ] Ollama not running: status shows **Ollama is not running** (no crash,
      no stack trace in the UI)
- [ ] Invalid endpoint: status shows **Invalid endpoint**
- [ ] Selected model removed: notice shown, no silent switch
- [ ] Genspark sign-in + chat works (cloud provider)
- [ ] One direct vendor key (e.g. OpenAI or Claude) works end-to-end
- [ ] No cloud fallback: with Ollama selected and Ollama down, requests fail
      with the local-AI error, never route to a cloud provider

## Security

- [ ] Production logs contain no API keys, no `Authorization` headers, and
      no unnecessary document content
- [ ] API keys are stored only in the local settings file under user data
- [ ] No prompt/document-content telemetry added

## License / attribution

- [ ] `LICENSE` and `NOTICE` untouched
- [ ] `ee/` untouched
- [ ] README still states the fork relationship and upstream attribution
- [ ] About dialog shows the GenOffice-derived attribution

## Docs

- [ ] README download/install links match the actual artifacts
- [ ] `docs/installation.md`, `docs/ollama.md`, `docs/providers.md` accurate
      for this version

## Git

- [ ] Working tree reviewed: `git status`, `git diff --stat`
- [ ] Commit message summarizes the change; release tagged with the app
      version
- [ ] `LICENSE`/`NOTICE` byte-identical to upstream before tagging
