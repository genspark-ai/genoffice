# Fork release pipeline

Operational manual for shipping installers from this fork
(`shuairongzeng/genoffice`). Everything below is about
`.github/workflows/release.yml` and the `scripts/release-*.cjs` helpers; for
the packaging configuration itself see `apps/shell/electron-builder.cjs`.

## What the pipeline does

```
push tag v0.13.0 (or manual dispatch with a version)
        │
        ▼
   guard ─── version parses as x.y.z, and is newer than the latest published
        │     GitHub Release (same version allowed for re-runs, older rejected —
        │     the forward-only rule shared with scripts/update-feed-utils.cjs)
        ▼
   build (macos-14 arm64 / windows-latest x64 / ubuntu-22.04 x64)
        │     npm ci → npm run typecheck → npm test → npm run build:all
        │     mac: universal xlsx-sidecar (both rustup Apple targets)
        │     → node scripts/release-package.cjs (electron-builder, version
        │       injected, --publish never, no signing secrets = unsigned)
        │     → node scripts/release-artifacts.cjs (assert + file list)
        ▼
   publish ─ one GitHub Release: all three legs' artifacts, release notes
             from git log (scripts/release-notes.cjs), uploaded with gh
```

A red quality gate never produces a package: typecheck and the full unit
suite run in every build leg **before** packaging, and `publish` needs all
three legs. `release-artifacts.cjs` additionally asserts the expected file
names — electron-builder can exit 0 while an artifact silently never got
built, so the workflow checks rather than trusts (same philosophy as the
`beforePack` assertions in `apps/shell/electron-builder.cjs`).

## Cutting a release

1. (Recommended) bump `version` in `apps/shell/package.json` to match — CI
   injects the tag version into the packaged app anyway, but keeping them in
   sync means local `dist:*` builds report the same version.
2. Tag and push:

   ```bash
   git tag v0.13.0
   git push origin v0.13.0
   ```

3. Watch the **Release** workflow run. When it is green, the release exists
   at `https://github.com/shuairongzeng/genoffice/releases/tag/v0.13.0`.

Manual dispatch (builds whatever `main` points at): **Actions → Release →
Run workflow**, enter the bare version (`0.13.0`). The workflow creates the
`v<version>` tag at the current commit during publishing.

The version must be plain `x.y.z` semver and must move forward: the guard
rejects anything older than the latest published release (matching the
update-feed guard, so releases and feed can never go backwards). Re-running
a failed workflow re-publishes the same version fine; rolling back means
`gh release delete v<newer>` first, then re-run.

## Artifacts

| platform               | files                                                                           | notes                                                                                                  |
| ---------------------- | ------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------ |
| macOS (Apple Silicon)  | `GenOffice-<v>-arm64.dmg`, `GenOffice-<v>-arm64-mac.zip`                        | arm64 only; Intel builds are opt-in upstream via `GENOFFICE_MAC_X64` and are not enabled for this fork |
| Windows (x64)          | `GenOffice Setup <v>.exe`                                                       | NSIS, per-user selectable dir                                                                          |
| Linux (x64)            | `GenOffice-<v>.AppImage`, `genoffice_<v>_amd64.deb`, `genoffice-<v>.x86_64.rpm` | rpm is a download-only target, never part of an update feed                                            |
| update feed (optional) | `latest.yml`, `latest-mac.yml`, `latest-linux.yml`                              | only produced when `GENOFFICE_UPDATE_URL` is set                                                       |

On the runners they land in `apps/shell/release/` during the build; the
workflow then attaches exactly the asserted set to the GitHub Release. For
local builds the same directory is where `npm run dist:mac` / `dist:win` /
`dist:linux` write.

## Unsigned builds — what your users see

A fork has no code-signing or notarization secrets by default, and the
pipeline is built to degrade silently instead of failing:

- **macOS**: electron-builder finds no signing identity in the runner keychain
  and skips signing with a warning; `apps/shell/build/notarize-dmg.js` sees no
  Apple credentials and skips notarization. Users get
  _“GenOffice.app can't be opened because it is from an unidentified
  developer”_ — right-click → **Open** (twice on older macOS), or
  `xattr -dr com.apple.quarantine /Applications/GenOffice.app`. The release
  notes spell this out.
- **Windows**: `GENOFFICE_WIN_SIGN_MODE` stays unset, so electron-builder
  produces an unsigned NSIS installer and unsigned `xlsx-sidecar.exe`.
  SmartScreen shows _“Windows protected your PC”_ → **More info → Run
  anyway**. Note: security features that judge child processes individually
  (Smart App Control, WDAC) can kill the unsigned sidecar at runtime — that
  path needs real signing, see below.
- **Linux**: no signing concept; AppImage/deb/rpm install as-is.

To enable signing later: set `APPLE_ID`, `APPLE_APP_SPECIFIC_PASSWORD`,
`APPLE_TEAM_ID` (or `APPLE_KEYCHAIN_PROFILE`) as repository secrets — the mac
leg and `notarize-dmg.js` pick them up automatically and the release notes
switch to signed wording. Windows signing would additionally require
`GENOFFICE_WIN_SIGN_MODE` plus `scripts/win-sign.cjs`, which this fork does
not carry (upstream-only script).

## In-app updates: off by default

Without `GENOFFICE_UPDATE_URL` the electron-builder publish config is
omitted, so **no `app-update.yml` is baked into the app** (verified against
a local build: `apps/shell/release/mac-arm64/GenOffice.app/Contents/Resources`
contains none). At runtime that means
(`apps/shell/src/main/updater.ts:595`):

- the background check after 15 s rejects and is only logged — no dialog,
  no crash (`updater.ts:770`);
- a manual check shows the "couldn't check for updates" warning;
- Linux deb/rpm installs never self-update anyway; only the AppImage does.

So fork users get updates by watching this repository's
**Releases** (Watch → Custom → Releases) and downloading the new installer.

### Wiring your own update feed

1. Host the release files at any HTTPS base URL that serves
   `latest.yml` / `latest-mac.yml` / `latest-linux.yml` next to the
   installers (GitHub Pages, S3 + CloudFront, any static file host).
2. Set the `GENOFFICE_UPDATE_URL` repository secret to that base URL. The
   next release build bakes it into the app and emits the `latest*.yml`
   files, which the workflow asserts and attaches to the GitHub Release —
   mirror them (plus the installers the yml files reference) to your host.
3. Keep the host forward-only: electron-updater never downgrades
   (`updater.ts:623`), so an old version at the feed URL just means users
   stay put — but never let the feed advertise a version whose artifacts are
   not actually there.
4. macOS caveat: Squirrel.Mac applies updates only for signed, notarized
   builds (`updater.ts:608`). An unsigned fork can host a feed, but mac
   users still update manually — only Windows (NSIS) and Linux (AppImage)
   self-update unsigned.

## Workflow inputs, guards and re-runs

| mechanism           | behavior                                                                                                    |
| ------------------- | ----------------------------------------------------------------------------------------------------------- |
| `push` tag `v*`     | version from the tag                                                                                        |
| `workflow_dispatch` | version from the required `version` input; tag created at the dispatch commit                               |
| guard failure       | no runner time burned: bad semver or non-forward version exits before `build`                               |
| leg failure         | other legs keep running (`fail-fast: false`); `publish` is skipped                                          |
| `publish` re-run    | same-version rule lets it through; existing release is uploaded into with `--clobber` (idempotent)          |
| rollback            | delete the newer release first (`gh release delete v<x> --yes`), then re-run the workflow for the older tag |

## Local release builds

The repo's own scripts produce the same (unsigned) packages locally:

```bash
npm run dist:mac    # dmg + zip, universal sidecar
npm run dist:win    # NSIS exe (on a Windows host)
npm run dist:linux  # AppImage + deb + rpm (needs the `rpm` apt package)
```

Differences to CI: the packaged version comes from
`apps/shell/package.json` (CI injects the tag version), and a local
`apps/shell/electron-builder.env` (gitignored) can supply
`GENOFFICE_UPDATE_URL` to bake a feed URL for testing.

macOS note: `native:build:universal` needs the `x86_64-apple-darwin` target
of the _rustup_ toolchain. If a Homebrew rust shadows rustup in `PATH`
(`cargo --version` reports "(Homebrew)"), the cross build fails with
`can't find crate for core`; prefix the command with
`PATH="$HOME/.cargo/bin:$PATH"` (CI's dtolnay rust-toolchain action is
unaffected).
