# Plan 020: Extract the reusable Codex Node auth adapter from Docs

> **Implementation instructions**: Follow this plan step by step and run every
> verification command. Stop on a STOP condition; do not invent a broader IPC
> framework. Update `plans/README.md` after implementation and review.
>
> **Drift check (run first)**:
> `git diff --stat 1878b30..HEAD -- packages/ai-provider apps/docs/src/main/codex-auth-main.ts apps/docs/tests/codex-auth-main.test.ts`
> The audited implementation is uncommitted at this SHA. Compare the live
> symbols with Current state. Behavioral drift is a STOP condition.

## Status

- **Priority**: P1
- **Effort**: M
- **Risk**: LOW
- **Depends on**: `plans/019-make-codex-login-single-flight.md`
- **Category**: tech-debt
- **Planned at**: commit `1878b30`, 2026-08-15

## Why this matters

The provider service is reusable, but its required JSON store and localhost
callback live in a private Docs module. Sheets or Slides would have to import
another app's source or duplicate roughly 200 lines. Moving only the Node
adapter into the provider package leaves each app with a tiny Electron wrapper,
so later rollout is integration rather than rewrite.

## Current state

- `apps/docs/src/main/codex-auth-main.ts:1-241` mixes reusable Node concerns
  (atomic JSON storage and callback server) with the Docs Electron singleton at
  lines 243-255.
- Electron is used only by `getCodexAuth()` for `shell.openExternal`; URL
  validation uses `@genoffice/electron-utils`.
- `packages/ai-provider/src/auth.ts` defines `CodexCredentialStore` and
  `CodexCallbackHandle`; its `package.json` exports only `.`.
- `apps/docs/tests/codex-auth-main.test.ts` covers mode `0600`, read/write,
  malformed removal, readiness, collisions, invalid requests, cancellation,
  timeout, and successful completion.
- `CONTRIBUTING.md` permits pure TypeScript packages but no Electron dependency.

## Commands you will need

| Purpose        | Command                                                                               | Expected on success |
| -------------- | ------------------------------------------------------------------------------------- | ------------------- |
| Provider tests | `npm test -w @genoffice/ai-provider`                                                  | exit 0              |
| Docs test      | `npm test -w @genoffice/docs -- --run tests/codex-auth-main.test.ts`                  | exit 0              |
| Typecheck      | `npm run typecheck -w @genoffice/ai-provider && npm run typecheck -w @genoffice/docs` | exit 0              |
| Build          | `npm run build -w @genoffice/docs`                                                    | exit 0              |

## Scope

**In scope**:

- `packages/ai-provider/package.json`
- `packages/ai-provider/src/codex-auth-node.ts` (new)
- `packages/ai-provider/tests/codex-auth-node.test.ts` (new)
- `apps/docs/src/main/codex-auth-main.ts`
- `apps/docs/tests/codex-auth-main.test.ts`
- `plans/README.md` (status only)

**Out of scope**:

- Root `@genoffice/ai-provider` browser-safe exports
- Electron or IPC inside `packages/ai-provider`
- Other office-app wiring
- A generic auth framework, new package, or dependency
- Path/schema changes: keep `~/.genoffice/codex-auth.json`,
  `GENOFFICE_AUTH_DIR`, directory `0700`, file `0600`, and atomic rename

## Git workflow

- Work on the current branch; do not push or open a PR.
- If asked to commit: `share the Codex main-process auth adapter`.

## Steps

### Step 1: Add an explicit Node-only provider subpath

Create `packages/ai-provider/src/codex-auth-node.ts` and move the reusable file
store and callback server into it. Export only the factories used by apps
(`codexCredentialStore` and `beginCodexCallback`); keep parsing/file internals
private.

Add `"./codex-auth-node": "./src/codex-auth-node.ts"` to package exports. Do
not re-export it from `src/index.ts`: the explicit subpath keeps Node built-ins
out of renderer imports of the root.

Move reusable tests to
`packages/ai-provider/tests/codex-auth-node.test.ts`, preserving every existing
case and Plan 019's duplicate-callback regression.

**Verify**: `npm test -w @genoffice/ai-provider` -> all tests pass.

### Step 2: Reduce Docs to Electron composition

Change `apps/docs/src/main/codex-auth-main.ts` to import the factories from
`@genoffice/ai-provider/codex-auth-node`. Keep only `CodexAuthService`
composition, `safeExternalUrl`, and `shell.openExternal`. Inject `fetch`,
`Date.now`, store, browser opener, and callback without duplicating internals.

Keep/reduce the Docs test to its app-specific composition boundary; do not
repeat the moved adapter matrix in two workspaces.

**Verify**:

```sh
npm test -w @genoffice/docs -- --run tests/codex-auth-main.test.ts
npm run build -w @genoffice/docs
```

Expected: both exit 0.

### Step 3: Verify the package boundary

```sh
rg "from 'electron'|from \"electron\"" packages/ai-provider
rg "node:" packages/ai-provider/src/index.ts
npm run typecheck -w @genoffice/ai-provider
npm run typecheck -w @genoffice/docs
git diff --check
```

Expected: both searches return no matches; remaining commands exit 0.

## Test plan

- Storage schema, permissions, malformed-file, and callback lifecycle cases
  move to the provider workspace and stay green.
- Docs tests cover only its Electron composition.
- Production build proves the Node subpath does not contaminate renderer code.

## Done criteria

- [ ] No reusable store/server implementation remains in Docs.
- [ ] Node adapter is exposed only through an explicit subpath.
- [ ] `@genoffice/ai-provider` has no Electron import/dependency.
- [ ] Approved JSON path/schema/permissions are unchanged.
- [ ] Provider/Docs tests, typechecks, build, and diff checks pass.

## STOP conditions

- Tooling cannot expose a Node subpath without bundling it into root imports.
- Extraction requires Electron in `packages/ai-provider`.
- Plan 019 is incomplete or callback behavior differs.
- A verification fails twice after a reasonable in-scope correction.

## Maintenance notes

Future apps compose this adapter with their own Electron browser opener; they
must not import `apps/docs` or copy this implementation. Defer a generic AI IPC
framework until multiple apps prove the same missing abstraction.
