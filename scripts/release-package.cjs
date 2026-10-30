#!/usr/bin/env node
/**
 * scripts/release-package.cjs — run electron-builder for one release platform
 * with the release version injected.
 *
 * The release version comes from the git tag (or the manual dispatch input),
 * not from apps/shell/package.json: electron-builder receives it via
 * -c.extraMetadata.version, the same mechanism the config comment in
 * apps/shell/electron-builder.cjs describes, so the packaged app, its
 * artifacts and the bundled `genoffice --version` all report the tag version
 * (the beforePack hook rebuilds the CLI bundle when the two disagree).
 *
 * --publish never keeps electron-builder from uploading anything: publishing
 * to the GitHub Release is done by the release workflow's publish job, and a
 * generic update-feed URL (GENOFFICE_UPDATE_URL) is never uploaded to anyway.
 *
 * Expects `npm run build:all` (and for mac, the universal sidecar build) to
 * have run already — the config's beforePack hook asserts every
 * extraResources source and fails the build when one is missing.
 *
 * Usage: node scripts/release-package.cjs --platform mac|win|linux --version 0.13.0
 */

const { execFileSync } = require('node:child_process')
const { join } = require('node:path')

const PLATFORM_FLAG = { mac: '--mac', win: '--win', linux: '--linux' }

function parseArgs(argv) {
  const out = { platform: null, version: null }
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i]
    if (arg === '--platform' || arg === '--version') {
      const value = argv[++i]
      if (!value) {
        console.error(`missing value after ${arg}`)
        process.exit(2)
      }
      out[arg === '--platform' ? 'platform' : 'version'] = value
    } else {
      console.error(`unknown argument: ${arg}`)
      process.exit(2)
    }
  }
  return out
}

const root = join(__dirname, '..')
const { platform, version } = parseArgs(process.argv.slice(2))

if (!PLATFORM_FLAG[platform]) {
  console.error(`--platform must be one of ${Object.keys(PLATFORM_FLAG).join('|')}`)
  process.exit(2)
}
if (!/^\d+\.\d+\.\d+(?:[-+][0-9A-Za-z.-]+)?$/.test(version ?? '')) {
  console.error(`--version must be a plain x.y.z semver, got "${version}"`)
  process.exit(2)
}

const builder = join(root, 'node_modules', 'electron-builder', 'cli.js')
const cliArgs = [
  builder,
  PLATFORM_FLAG[platform],
  '--publish',
  'never',
  `-c.extraMetadata.version=${version}`,
]

console.error(`[release-package] electron-builder ${PLATFORM_FLAG[platform]} ${version}`)
try {
  execFileSync(process.execPath, cliArgs, { cwd: join(root, 'apps', 'shell'), stdio: 'inherit' })
} catch (err) {
  process.exitCode = err.status ?? 1
}
