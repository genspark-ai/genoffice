#!/usr/bin/env node
/**
 * scripts/release-notes.cjs — generate the GitHub Release body from git
 * history: every commit since the previous tag (capped), the download table,
 * per-platform install hints and the fork's no-secrets caveats.
 *
 * Signing flags: --mac-signed / --win-signed switch the install sections to
 * the signed wording (the release workflow passes them only when the
 * matching secrets exist; a fork without certificates ships unsigned and
 * says so). --feed-url marks the release as having an in-app update channel.
 *
 * Usage:
 *   node scripts/release-notes.cjs --version 0.13.0 [--repo owner/name] \
 *     [--max-commits 200] [--mac-signed] [--win-signed] [--feed-url https://…]
 *
 * The body is printed to stdout; run from a checkout with full history
 * (fetch-depth: 0) so the previous tag can be resolved.
 */

const { execFileSync } = require('node:child_process')

function git(args) {
  return execFileSync('git', args, { encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 })
}

function parseArgs(argv) {
  const out = {
    version: null,
    repo: process.env.GITHUB_REPOSITORY ?? null,
    maxCommits: 200,
    macSigned: false,
    winSigned: false,
    feedUrl: null,
  }
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i]
    if (arg === '--mac-signed') out.macSigned = true
    else if (arg === '--win-signed') out.winSigned = true
    else if (arg === '--version' || arg === '--repo' || arg === '--feed-url') {
      const value = argv[++i]
      if (!value) {
        console.error(`missing value after ${arg}`)
        process.exit(2)
      }
      out[arg === '--version' ? 'version' : arg === '--repo' ? 'repo' : 'feedUrl'] = value
    } else if (arg === '--max-commits') {
      out.maxCommits = Number(argv[++i]) || 200
    } else {
      console.error(`unknown argument: ${arg}`)
      process.exit(2)
    }
  }
  return out
}

// HEAD^ so a run on the new tag itself describes everything since the
// previous tag; HEAD as the fallback covers manual dispatches where the
// workflow commit is not tagged yet. --exclude keeps the fallback from
// matching the tag being released (empty log on the very first release).
function previousTag() {
  const exclude = `v${version}`
  for (const rev of ['HEAD^', 'HEAD']) {
    try {
      const tag = git(['describe', '--tags', '--abbrev=0', '--exclude', exclude, rev]).trim()
      if (tag) return tag
    } catch {
      // no reachable tag yet
    }
  }
  return null
}

function commitsSince(prev, max) {
  const range = prev ? `${prev}..HEAD` : 'HEAD'
  const all = git(['log', '--oneline', '--no-merges', range]).trim().split('\n')
  const list = all.filter(Boolean)
  return { list: list.slice(0, max), omitted: Math.max(0, list.length - max) }
}

const { version, repo, maxCommits, macSigned, winSigned, feedUrl } = parseArgs(
  process.argv.slice(2),
)
if (!/^\d+\.\d+\.\d+(?:[-+][0-9A-Za-z.-]+)?$/.test(version ?? '')) {
  console.error(`--version must be a plain x.y.z semver, got "${version}"`)
  process.exit(2)
}

const prev = previousTag()
const { list, omitted } = commitsSince(prev, maxCommits)
const tag = `v${version}`

const changes =
  list.length > 0
    ? list.map((line) => `- ${line}`).join('\n')
    : '- (no commits found since the previous tag)'
const omittedNote = omitted > 0 ? `\n\n…and ${omitted} more (see the full history).` : ''
const compareLink = repo && prev ? `\n\nFull diff: ${compareUrl(repo, prev, tag)}` : ''

function compareUrl(ownerRepo, from, to) {
  return `https://github.com/${ownerRepo}/compare/${from}...${to}`
}

const macSection = macSigned
  ? 'macOS: open the `.dmg` and drag GenOffice to Applications. The build is\nsigned and notarized.'
  : `macOS: open the \`.dmg\` and drag GenOffice to Applications. This is an
**unsigned, un-notarized** fork build: Gatekeeper blocks the first launch —
right-click the app and choose **Open** (or run
\`xattr -dr com.apple.quarantine /Applications/GenOffice.app\`), then confirm.`

const winSection = winSigned
  ? `Windows: run \`GenOffice Setup ${version}.exe\`. The installer is code-signed.`
  : `Windows: run \`GenOffice Setup ${version}.exe\`. The installer is
**unsigned**: SmartScreen shows "Windows protected your PC" — choose
**More info → Run anyway**. Import-only on Windows Defender/Smart App
Control machines is not guaranteed for unsigned child processes.`

const feedSection = feedUrl
  ? `In-app updates are enabled: this release is published to the update channel
at ${feedUrl} (\`latest*.yml\` + installers). Keep that host in sync with the
GitHub Releases — the updater never downgrades.`
  : `This fork ships **without an in-app update feed**: there is no bundled
\`app-update.yml\`, so the app never checks for updates on its own. Watch this
repository's **Releases** (Watch → Custom → Releases) and install new
versions manually. See docs/release.md for how to wire up an own feed.`

const body = `# GenOffice ${tag}

${prev ? `Changes since ${prev} (${list.length} commits):` : 'Initial fork release:'}

${changes}${omittedNote}${compareLink}

## Downloads

| Platform | File |
| --- | --- |
| macOS (Apple Silicon) | \`GenOffice-${version}-arm64.dmg\` |
| Windows (x64) | \`GenOffice Setup ${version}.exe\` |
| Linux (AppImage / deb / rpm) | \`GenOffice-${version}.AppImage\`, \`genoffice_${version}_amd64.deb\`, \`genoffice-${version}.x86_64.rpm\` |

## Installing

${macSection}

${winSection}

Linux: \`chmod +x\` the \`.AppImage\` and run it, or install the \`deb\`
(\`sudo apt install ./genoffice_${version}_amd64.deb\`) / \`rpm\`
(\`sudo dnf install ./genoffice-${version}.x86_64.rpm\`).

## Updates

${feedSection}

---

Release notes generated by \`scripts/release-notes.cjs\` from commit
\`${git(['rev-parse', '--short', 'HEAD']).trim()}\`.
`

process.stdout.write(body)
