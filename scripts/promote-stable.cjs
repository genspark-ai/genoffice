#!/usr/bin/env node
/**
 * scripts/promote-stable.cjs — promote an already-published beta build to the
 * stable update channel. No rebuild: the versioned feed archive uploaded at
 * beta-publish time (GenOffice-mac-arm64-<v>.yml / GenOffice-win-<v>.yml) is
 * re-uploaded as the stable feed (latest-mac.yml / latest.yml). The binaries
 * it points to are already on the CDN under the same prefix.
 *
 * The marketing download aliases (GenOffice.dmg / GenOfficeSetup.exe) also
 * track the stable channel: per-merge beta publishes skip them, and this
 * script re-points each alias to the promoted version via a server-side
 * blob copy (no download/re-upload).
 *
 * Usage:
 *   node scripts/promote-stable.cjs [--mac <version>|latest] [--win <version>|latest] [--force] [--dry-run]
 *
 * At least one of --mac/--win is required (version sequences are independent
 * per platform). "latest" resolves to the version currently served by the
 * platform's beta feed. Guard: refuses to publish a version that is not
 * strictly newer than the current stable feed unless --force is passed.
 *
 * Requires: az CLI + AZURE_STORAGE_CONNECTION_STRING + GENOFFICE_UPDATE_URL
 * (https://<cdn-host>/<container>/<prefix>, same convention as the upload
 * scripts — container/prefix are derived from its path).
 */

const { execFileSync } = require('child_process')
const fs = require('fs')
const os = require('os')
const path = require('path')
const https = require('https')
const { ymlVersion, assertPromotable } = require('./update-feed-utils.cjs')

const dryRun = process.argv.includes('--dry-run')
const force = process.argv.includes('--force')

function fatal(msg) {
  console.error(`[promote-stable] ERROR: ${msg}`)
  process.exit(1)
}

function argValue(flag) {
  const idx = process.argv.indexOf(flag)
  if (idx === -1) return null
  const v = process.argv[idx + 1]
  // 'latest' resolves to the newest version on the platform's beta feed
  if (v === 'latest') return v
  if (!v || !/^\d+\.\d+\.\d+$/.test(v)) fatal(`${flag} requires an x.y.z version or "latest"`)
  return v
}

// installer names follow the build pipelines: electron-builder's default
// mac artifact name (arm64 CI builds) and windows-build.yml's staged
// versioned copy. The alias is the stable marketing download link.
const PLATFORMS = [
  {
    flag: '--mac',
    archive: (v) => `GenOffice-mac-arm64-${v}.yml`,
    feed: 'latest-mac.yml',
    betaFeed: 'beta-mac.yml',
    installer: (v) => `GenOffice-${v}-arm64.dmg`,
    alias: 'GenOffice.dmg',
  },
  {
    flag: '--win',
    archive: (v) => `GenOffice-win-${v}.yml`,
    feed: 'latest.yml',
    betaFeed: 'beta.yml',
    installer: (v) => `GenOfficeSetup-v${v}.exe`,
    alias: 'GenOfficeSetup.exe',
  },
]

function channelTarget() {
  const raw = process.env.GENOFFICE_UPDATE_URL
  if (!raw) fatal('GENOFFICE_UPDATE_URL env not set')
  const base = raw.replace(/\/+$/, '')
  let segments
  try {
    segments = new URL(base).pathname.replace(/^\/+/, '').split('/')
  } catch {
    fatal('GENOFFICE_UPDATE_URL is not a valid URL')
  }
  const container = segments[0]
  const prefix = segments.slice(1).join('/')
  if (!container || !prefix) {
    fatal('GENOFFICE_UPDATE_URL must look like https://<cdn-host>/<container>/<prefix>')
  }
  return { container, prefix, base }
}

function fetchText(url) {
  return new Promise((resolve) => {
    https
      .get(url, { headers: { 'cache-control': 'no-cache' } }, (res) => {
        if (res.statusCode !== 200) {
          res.resume()
          resolve(null)
          return
        }
        let body = ''
        res.on('data', (c) => (body += c))
        res.on('end', () => resolve(body))
      })
      .on('error', () => resolve(null))
  })
}

/** HEAD probe so a missing source blob fails the promote before any write */
function blobExists(url) {
  return new Promise((resolve) => {
    const req = https.request(url, { method: 'HEAD' }, (res) => {
      res.resume()
      resolve(res.statusCode === 200)
    })
    req.on('error', () => resolve(false))
    req.end()
  })
}

/** server-side copy within the same container (installer alias re-point);
 * same-account copies complete synchronously enough for CI use */
function azCopyBlob(target, sourceBlob, destBlob) {
  execFileSync(
    'az',
    [
      'storage',
      'blob',
      'copy',
      'start',
      '--destination-container',
      target.container,
      '--destination-blob',
      `${target.prefix}/${destBlob}`,
      '--source-container',
      target.container,
      '--source-blob',
      `${target.prefix}/${sourceBlob}`,
      '--output',
      'none',
    ],
    { stdio: 'inherit' },
  )
}

function azUpload(target, localPath, blobName) {
  execFileSync(
    'az',
    [
      'storage',
      'blob',
      'upload',
      '--container-name',
      target.container,
      '--file',
      localPath,
      '--name',
      `${target.prefix}/${blobName}`,
      '--overwrite',
      '--content-type',
      'text/yaml',
      '--content-cache-control',
      'no-cache',
      '--output',
      'none',
    ],
    { stdio: 'inherit' },
  )
}

async function promoteOne(target, platform, requestedVersion) {
  let version = requestedVersion
  if (version === 'latest') {
    // resolve to whatever the beta channel currently serves — the archive
    // and installer for that version exist by construction (same publish)
    const betaYml = await fetchText(`${target.base}/${platform.betaFeed}`)
    if (!betaYml) fatal(`${platform.betaFeed} not found on the CDN — nothing published to beta yet`)
    version = ymlVersion(betaYml)
    if (!version) fatal(`could not parse version from ${platform.betaFeed}`)
    console.log(`[promote-stable] ${platform.flag} latest resolves to ${version}`)
  }
  const archiveName = platform.archive(version)
  const archived = await fetchText(`${target.base}/${archiveName}`)
  if (!archived) fatal(`${archiveName} not found on the CDN — was ${version} published to beta?`)
  const archivedVersion = ymlVersion(archived)
  if (archivedVersion !== version) {
    fatal(`${archiveName} declares version ${archivedVersion}, expected ${version}`)
  }

  const installerName = platform.installer(version)
  if (!(await blobExists(`${target.base}/${installerName}`))) {
    fatal(`${installerName} not found on the CDN — cannot re-point ${platform.alias}`)
  }

  const stable = await fetchText(`${target.base}/${platform.feed}`)
  const stableVersion = stable ? ymlVersion(stable) : null
  const verdict = assertPromotable(version, stableVersion, force)
  if (!verdict.ok) fatal(`${platform.feed}: ${verdict.reason}`)

  console.log(`[promote-stable] ${platform.feed}: ${stableVersion ?? '(none)'} -> ${version}`)
  console.log(`[promote-stable] ${platform.alias} -> ${installerName}`)
  if (dryRun) return

  // alias first, feed last: a client reading the new feed mid-promote must
  // already find every artifact it references
  azCopyBlob(target, installerName, platform.alias)
  const tmp = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'promote-')), platform.feed)
  fs.writeFileSync(tmp, archived)
  azUpload(target, tmp, platform.feed)
  console.log(`[promote-stable] published ${target.base}/${platform.feed}`)
}

async function main() {
  const requested = PLATFORMS.map((p) => ({ p, version: argValue(p.flag) })).filter(
    (r) => r.version,
  )
  if (requested.length === 0) fatal('nothing to do — pass --mac <version> and/or --win <version>')
  const target = channelTarget()
  if (!dryRun && !process.env.AZURE_STORAGE_CONNECTION_STRING) {
    fatal('AZURE_STORAGE_CONNECTION_STRING env not set')
  }
  for (const { p, version } of requested) await promoteOne(target, p, version)
}

main().catch((err) => fatal(err.message))
