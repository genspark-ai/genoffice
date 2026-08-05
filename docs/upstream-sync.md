# Keeping a provider-enabled fork current

GenOffice source updates and user AI configuration are separate concerns. The
provider-enabled fork should rebase its source branch from the upstream
repository, while each installed copy keeps its own `ai-settings.json` and
`ai-credentials.json` in Electron `userData`. The credentials file contains
safeStorage ciphertext and must never be committed or copied into a pull
request.

## One-time repository setup

Keep the fork as `origin` and add the official repository as `upstream`:

```sh
git remote rename origin fork
git remote add origin https://github.com/YOUR_ACCOUNT/genoffice.git
git remote add upstream https://github.com/genspark-ai/genoffice.git
git fetch --all --prune
```

Use the actual fork URL. Do not put API keys, signed URLs, or a private package
registry token in a remote URL.

## Safe local update

First make sure the working tree is clean. Commit or stash source changes and
close running builds before syncing. The helper creates a local backup ref and
refuses tracked or symbolic-link settings paths.

Run a fetch-only plan. A dry run does not create a branch or modify the worktree:

```sh
npm run sync:upstream -- --remote upstream --branch main
```

Apply the rebase and protect the installed user's two files:

```sh
npm run sync:upstream -- --apply \
  --settings-path "$GENOFFICE_USER_DATA/ai-settings.json" \
  --credentials-path "$GENOFFICE_USER_DATA/ai-credentials.json"
```

`GENOFFICE_USER_DATA` is the test/dev user-data directory when set. For a
packaged macOS app, resolve the actual Electron user-data directory from the
app before running the command. The helper copies the files as bytes into a
mode-0700 temporary directory, rebases the source branch, and restores the
files byte-for-byte even when the rebase stops with conflicts. It never prints
their contents.

The helper refuses to start with a dirty worktree and never runs `git reset
--hard`, `git clean`, or a forced push. It leaves a conflict active for manual
resolution:

```sh
git status
# resolve source conflicts only
git add <resolved-source-files>
git rebase --continue
```

The backup ref printed by the command can be used to inspect or recover the
pre-rebase source state:

```sh
git show backup/upstream-sync-<timestamp>
git diff backup/upstream-sync-<timestamp>..HEAD
```

Delete a backup ref only after the rebase and tests are verified:

```sh
git branch -d backup/upstream-sync-<timestamp>
```

## Pull-request and release flow

Keep provider work in focused commits so upstream conflicts are small:

1. Fetch upstream and create a backup ref.
2. Rebase the fork branch onto `upstream/main`.
3. Resolve only source conflicts; never resolve by replacing user-data files.
4. Run formatting, lint, typecheck, unit tests, and the provider mock E2E suite.
5. Push the fork branch and open a pull request against the fork's `main`.
6. Build and publish fork installers from the fork's own signed release feed.

The official GenOffice update feed must not be used for a modified fork. An
official installer can overwrite a fork binary and remove the fork's provider
runtime, even though Electron userData remains. Keep the fork's app ID and
update endpoint stable across releases; if the app ID changes, provide a
one-time userData migration before enabling updates.

## Automated upstream PRs

The optional scheduled workflow should only fetch upstream and open a review
branch. It must not merge, publish an installer, or touch user-data paths.
Review the generated PR with the same tests as a manual rebase. A conflict
should fail the workflow and require a maintainer to resolve it; no workflow
should force-push `main`.

## Credential and settings rules

- `ai-settings.json` contains provider IDs, models, endpoints, active task
  selections, and credential references. It must not contain API-key values.
- `ai-credentials.json` contains encrypted values produced by Electron
  `safeStorage`. It is machine-bound and must not be copied to another
  computer or checked into Git.
- Legacy `apiKey` values are migrated once in the main process. The sanitized
  versioned settings file is written only after the key has been accepted by
  the OS credential store.
- If safeStorage is unavailable, saving a new credential fails without writing
  a plaintext fallback. Existing legacy data remains available for a later
  retry after the OS credential store is available.
- Renderer snapshots contain only `credentialConfigured: boolean`; the key is
  resolved for provider requests in the main process.

Run the migration and workflow safety tests directly when changing this area:

```sh
npm run typecheck -w @genoffice/ai-electron
npm test -w @genoffice/ai-electron
npm run test:upstream-sync
```
