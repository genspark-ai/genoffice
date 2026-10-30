#!/usr/bin/env node
/**
 * scripts/release-artifacts.cjs — assert the expected installer set exists and
 * list the release files.
 *
 * Mirrors the beforePack philosophy in apps/shell/electron-builder.cjs:
 * electron-builder can exit 0 while an artifact never materialized, so the
 * workflow asserts the expected files by name before anything is uploaded.
 *
 * With --feed (the fork owner set GENOFFICE_UPDATE_URL) the electron-updater
 * metadata written next to the installers is asserted and listed too; without
 * it no feed files are produced and none are expected.
 *
 * Usage:
 *   node scripts/release-artifacts.cjs --platform mac|win|linux|all \
 *     --version 0.13.0 [--dist apps/shell/release] [--feed]
 *
 * Missing or empty files exit 1 with the offending names. On success the
 * publishable file paths are printed one per line on stdout (logs go to
 * stderr), so callers can feed the list straight to `gh release upload`.
 */

const { statSync } = require('node:fs')
const { join, resolve } = require('node:path')

// Names follow electron-builder's defaults plus the pinned deb/rpm
// artifactName patterns in apps/shell/electron-builder.cjs:
// mac is arm64-only (GENOFFICE_MAC_X64 unset), win is the x64 NSIS build
// (no arch suffix), linux keeps the published names (GenOffice-<v>.AppImage /
// genoffice_<v>_amd64.deb / genoffice-<v>.x86_64.rpm).
const EXPECTED = {
  mac: (v) => [`GenOffice-${v}-arm64.dmg`, `GenOffice-${v}-arm64-mac.zip`],
  win: (v) => [`GenOffice Setup ${v}.exe`],
  linux: (v) => [
    `GenOffice-${v}.AppImage`,
    `genoffice_${v}_amd64.deb`,
    `genoffice-${v}.x86_64.rpm`,
  ],
}
const FEED_FILES = {
  mac: ['latest-mac.yml'],
  win: ['latest.yml'],
  linux: ['latest-linux.yml'],
}

function parseArgs(argv) {
  const out = { platform: null, version: null, dist: null, feed: false }
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i]
    if (arg === '--feed') out.feed = true
    else if (arg === '--platform' || arg === '--version' || arg === '--dist') {
      const value = argv[++i]
      if (!value) {
        console.error(`missing value after ${arg}`)
        process.exit(2)
      }
      out[arg === '--platform' ? 'platform' : arg === '--version' ? 'version' : 'dist'] = value
    } else {
      console.error(`unknown argument: ${arg}`)
      process.exit(2)
    }
  }
  return out
}

const root = join(__dirname, '..')
const { platform, version, dist, feed } = parseArgs(process.argv.slice(2))
const distDir = resolve(process.cwd(), dist ?? join(root, 'apps', 'shell', 'release'))

const platforms = platform === 'all' ? Object.keys(EXPECTED) : [platform]
if (platforms.some((p) => !EXPECTED[p])) {
  console.error(`--platform must be one of ${Object.keys(EXPECTED).join('|')}|all`)
  process.exit(2)
}
if (!/^\d+\.\d+\.\d+(?:[-+][0-9A-Za-z.-]+)?$/.test(version ?? '')) {
  console.error(`--version must be a plain x.y.z semver, got "${version}"`)
  process.exit(2)
}

const wanted = []
for (const p of platforms) {
  wanted.push(...EXPECTED[p](version).map((name) => ({ name, platform: p })))
  if (feed) wanted.push(...FEED_FILES[p].map((name) => ({ name, platform: p })))
}

const files = []
const missing = []
for (const { name, platform: p } of wanted) {
  const path = join(distDir, name)
  let size = -1
  try {
    size = statSync(path).size
  } catch {
    // reported below
  }
  if (size > 0) {
    files.push(path)
  } else {
    missing.push(`${p}: ${name}${size === 0 ? ' (empty)' : ''}`)
  }
}

if (missing.length > 0) {
  console.error(`[release-artifacts] missing expected files in ${distDir}:`)
  for (const line of missing) console.error(`  ${line}`)
  process.exit(1)
}

console.error(`[release-artifacts] ${files.length} expected file(s) present in ${distDir}`)
for (const path of files) console.log(path)
