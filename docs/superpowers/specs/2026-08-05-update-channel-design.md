# Update Channel Setting (Stable / Beta) — Design

Date: 2026-08-05
Status: Approved (design), pending implementation plan

## Problem

Every merge to `main` publishes a new release to the single update channel
(`latest.yml` / `latest-mac.yml`). All installed clients poll that feed
(15s after launch, then every 4h) and show the strong-guidance update
modal, so users see an update prompt for effectively every code change.

## Goal

Two update channels selectable in the app:

- **Stable (default)** — releases are published only by a manual promote
  action. Existing installed clients land here with zero migration.
- **Beta** — current cadence: every merge to `main` publishes, clients on
  this channel update continuously.

## Design

### Release pipeline (CI)

- `mac-build.yml` / `windows-build.yml`: release builds (pushes to `main`
  / `release_*`) keep building and uploading installers to the same CDN
  prefix, but publish the feed as `beta.yml` / `beta-mac.yml` instead of
  `latest.yml` / `latest-mac.yml`. `latest*.yml` is no longer touched by
  the per-merge pipeline.
- New **promote workflow** (`workflow_dispatch`, input: a version already
  published to beta). Promote mode — no rebuild: it rewrites the stable
  feed (`latest.yml` / `latest-mac.yml`) to point at that version's
  already-uploaded artifacts. The binaries shipped to stable users are
  byte-identical to the beta-validated ones.
- Guard: refuse to promote a version that is not strictly newer than the
  version currently in the stable feed (same monotonicity rule
  `mac-release-upload.cjs` already applies to uploads).
- Windows and mac version sequences are independent (each platform's
  patch comes from its own workflow run number, and each has its own
  feed), so the promote workflow takes a per-platform version input and
  promotes each platform's feed separately (either or both per run).

### Client (apps/shell)

- `userData/app-settings.json` gains `updateChannel: 'stable' | 'beta'`.
  Missing or invalid value → `stable`.
- `updater.ts` reads the setting at startup: stable → electron-updater
  channel `latest` (today's behavior), beta → channel `beta`.
- New IPC pair on the home API: `getUpdateChannel()` /
  `setUpdateChannel(channel)`. Setting the channel persists it, updates
  `autoUpdater.channel`, and triggers an immediate `checkForUpdates()` so
  the switch gives instant feedback.
- **No downgrade**: a beta user switching to stable while beta is ahead
  keeps the installed version until stable catches up (electron-updater
  default; no extra code). A stable user switching to beta upgrades
  immediately and then continuously.
- Existing dismissal logic (`dismissedVersion`) is unchanged.

### UI

- Account popover in `Home.tsx` (language / version / sign-out menu):
  add an "Update channel" row below the language row, reusing the
  language row's flyout pattern — two options, Stable (default, checked)
  and Beta, checkmark on the active one.
- Strings added to `strings.ts` for all supported languages.

### Testing

- `apps/shell/tests/updater.test.ts`: defaults to stable; reads beta
  from settings; channel switch persists and triggers a recheck; invalid
  stored value falls back to stable.
- Promote script unit tests: monotonicity guard rejects promoting a
  version not strictly newer than the current stable feed.

## Out of scope

- No per-user gradual rollout / percentage rollout.
- No third channel (alpha/nightly).
- The docs app updater (`apps/docs/src/main/updater.ts`) is not changed
  in this iteration.
