# Docs SDK versioning and delivery

The Docs SDK is a private development package, currently version `0.1.0`, built
from the Docs workspace version. It can be copied, packed and installed locally;
this project does not promise a published registry package or signed desktop
installers merely because the corresponding build commands exist.

## Compatibility policy

Unreleased logical-text change: external `apply_ops/findReplace` now requires
exactly one literal match by default, including text split across formatting
runs. Existing bulk callers must pass `match: "all"` and an explicit target
scope. Zero/multiple precise matches refuse atomically; identical replacements
do not advance revision or split undo history. Native UI defaults are unchanged.
The replacement inherits match-start formatting and stays inside one paragraph
without crossing cells, protected/deleted spans, fields, images or formulas.
Occurrence counts are separate from the existing block-count summaries.

Unreleased safety change: public live-document HTML writes now require the
document/revision-bound image tags returned by `read_blocks`. Bare `doc:B.I`,
invalid/stale references and unsupported HTTP/data `<img>` sources refuse the
whole write rather than silently omitting requested content. Coordinate-based
media reads remain available; new pictures use `insert_picture`. This is an
intentional tightening of unsafe behavior, not a claim of compatible legacy
write semantics. See the authoritative SDK guide for reference diagnostics and
supported image versus protected-object boundaries.

The public contract is the [SDK guide](../apps/docs/sdk/README.md), its exported
types and observed public behavior. During `0.x`, a minor version may contain a
breaking public contract change. Patch versions are intended for compatible
fixes; document any unavoidable behavior change explicitly. Do not infer stable
API guarantees from the current development version.

The development `0.1.0` header/footer tools now support optional `textboxIndex`
coordinates and report each textbox's current-part lifetime and editability.
Supported replacements retain their original OOXML containers and compatibility
twins. Previously accepted writes to unsupported structured textboxes now fail
before mutation; consumers must refresh context and select an editable container
or leave the part unchanged. Existing section/view/scope defaults are unchanged.

Before a versioned delivery, update the Docs workspace version and any native
package identity that represents that delivery, then record consumer-facing
release notes in GitHub Releases or the delivery issue. Notes must identify:

- Added, changed or removed editor, host-service, runtime, tool and UI contracts.
- Compatibility and migration requirements, including new host services or peers.
- Document format, persistence, permission or cancellation behavior changes.
- Current integration targets, actual verification results and known limitations.

Do not rename retained `@genoffice/*` workspace identifiers, the `genoffice` CLI,
compatibility environment variables or private resource schemes solely to change
branding. These names do not imply affiliation with upstream or an SDK dependency
on its agent.

The imported downstream integration adds Qt `configureResources`, optional
`readOnly` MCP server creation and `readOnlyConnection` in trusted-host start/state
results. Existing `connection` consumers keep their editing access; review clients
must receive only the separate read-only configuration. Both connections end with
their shared document binding. Resource selection still needs read grants and does
not grant writes. Include these additive contracts and their lifetime rules in
the next delivery's compatibility notes. Successful saving of the active document
also publishes its saved baseline before the next call; no caller-side delay is
required. Save As with `adopt: false` exports a copy and retains the active
baseline, identity and dirty state.

## Build and consume

From the repository root with Node 24+ and npm 10+:

```sh
npm ci
npm run build:sdk
npm run check:sdk-package
```

The SDK guide documents local installation and examples. Package acceptance must
install the actual built package into a consumer outside the monorepo, resolve
public entry points/types and exercise authorized file processing. Workspace
symlinks are not evidence that the independent consumer works.

The package contains editor assets, Node bundles and their required dependency
closure, public declarations, examples, Qt host sources, PDFium WASM and applicable
license/attribution files. Electron, Qt and Chromium runtimes are supplied by the
host. `build-info.json` records the source revision, upstream baseline, dirty
status and build Node version. A dirty build can aid development, but cannot be
presented as an exact reproduction of its recorded commit.

## Delivery acceptance

For each delivery, record evidence in its GitHub issue or release, not in a
repository-local process ledger:

1. Identify source commit, SDK version, upstream baseline and build provenance.
2. Verify independent installation, runtime/type imports, assets, services,
   examples, host sources, notices and the documented quickstart.
3. Run applicable existing file-fidelity, file-grant, document-lifetime,
   hosted-save, selection and instruction-hook coverage.
4. Exercise the actual supported browser and Electron SDK modes, and the Qt
   native targets on an appropriate platform. Record OS and engine versions,
   commands and results; mark skipped or unavailable targets explicitly.
5. For claimed MCP integrations, record the real target client/version and
   verify discovery and live-document use. A reference client alone does not
   establish interoperability with other clients.
6. Check applicable licenses, upstream notices and changed-file modification
   notices, and review public documentation against the delivered artifact.
7. Review and commit source, tests and documentation; account for all task-owned
   files and report the final working-tree status.

Use the root SDK commands and the existing native Qt test instructions in the
SDK guide. CI builds SDK prerequisites before host integration tests; inspect
the job results rather than inferring acceptance from desktop compilation.
Successful packaging alone does not complete the broader Docs milestone.

## Desktop and downstream packaging

Desktop builds and SDK builds are separate deliverables. Desktop packaging uses
its own native dependencies, installer identities and signing configuration.
Source-built artifacts may be unsigned. Do not link upstream downloads as
ForaOffice builds or claim platforms for which no fork artifact was delivered.
The inherited Linux wrapper recipes identify when they fetch upstream GenOffice
binaries and therefore omit this fork's changes.

Build-injected desktop settings remain explicit:

| Setting                                                       | When absent                                                                        |
| ------------------------------------------------------------- | ---------------------------------------------------------------------------------- |
| `GENOFFICE_UPDATE_URL`                                        | No update feed is baked into the package; in-app automatic updates remain disabled |
| `GENOFFICE_GA4_MEASUREMENT_ID` and `GENOFFICE_GA4_API_SECRET` | No analytics credentials are injected; the tracker is inactive                     |
| `GENOFFICE_FONT_CDN_URL`                                      | Downloadable-font catalog entries/prompts are hidden; local fonts remain usable    |

These desktop options are not requirements of the embedded editor contract.
Changing a distribution's data collection or network services requires matching
[privacy disclosure](../PRIVACY.md), not just a new configuration value.
