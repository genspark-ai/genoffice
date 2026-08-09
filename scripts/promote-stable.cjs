#!/usr/bin/env node
/**
 * scripts/promote-stable.cjs — promote an already-published beta build to the
 * stable update channel. No rebuild: the versioned feed archive uploaded at
 * beta-publish time (GenOffice-mac-arm64-<v>.yml / GenOffice-win-<v>.yml /
 * GenOffice-linux-<v>.yml) is re-uploaded as the stable feed
 * (latest-mac.yml / latest.yml / latest-linux.yml). The binaries it points
 * to are already on the CDN under the same prefix.
 *
 * The marketing download aliases (GenOffice.dmg / GenOfficeSetup.exe /
 * GenOffice.AppImage + GenOffice.deb) also track the stable channel:
 * per-merge beta publishes skip them, and this script re-points each alias
 * by downloading the promoted installer to the runner and re-uploading it
 * under the alias name — the same auth path the build pipelines use. (A
 * server-side `az storage blob copy start` cannot read the source blob: the
 * container is private and the connection string does not sign the copy
 * source, so Azure rejects it with CannotVerifyCopySource.)
 *
 * Usage:
 *   node scripts/promote-stable.cjs [--mac <version>|latest] [--win <version>|latest] [--linux <version>|latest] [--force] [--dry-run]
 *
 * At least one of --mac/--win/--linux is required (version sequences are
 * independent per platform). "latest" resolves to the version currently
 * served by the platform's beta feed. Guard: refuses to publish a version
 * that is not strictly newer than the current stable feed unless --force is
 * passed.
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
// mac/linux artifact names (arm64 mac CI builds, x64 linux) and
// windows-build.yml's staged versioned copy. Each alias is a stable
// marketing download link; linux has two (AppImage + deb).
const PLATFORMS = [
  {
    flag: '--mac',
    archive: (v) => `GenOffice-mac-arm64-${v}.yml`,
    feed: 'latest-mac.yml',
    betaFeed: 'beta-mac.yml',
    aliases: [{ installer: (v) => `GenOffice-${v}-arm64.dmg`, alias: 'GenOffice.dmg' }],
  },
  {
    flag: '--win',
    archive: (v) => `GenOffice-win-${v}.yml`,
    feed: 'latest.yml',
    betaFeed: 'beta.yml',
    aliases: [{ installer: (v) => `GenOfficeSetup-v${v}.exe`, alias: 'GenOfficeSetup.exe' }],
  },
  {
    flag: '--linux',
    archive: (v) => `GenOffice-linux-${v}.yml`,
    feed: 'latest-linux.yml',
    betaFeed: 'beta-linux.yml',
    aliases: [
      { installer: (v) => `GenOffice-${v}.AppImage`, alias: 'GenOffice.AppImage' },
      { installer: (v) => `genoffice_${v}_amd64.deb`, alias: 'GenOffice.deb' },
    ],
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

// installer content types match the build pipelines' upload steps
// (mac-release-upload.cjs / windows-build.yml)
function contentTypeFor(name) {
  if (name.endsWith('.dmg')) return 'application/x-apple-diskimage'
  if (name.endsWith('.exe')) return 'application/x-msdownload'
  if (name.endsWith('.AppImage')) return 'application/vnd.appimage'
  if (name.endsWith('.deb')) return 'application/vnd.debian.binary-package'
  return 'application/octet-stream'
}

function azDownload(target, blobName, localPath) {
  execFileSync(
    'az',
    [
      'storage',
      'blob',
      'download',
      '--container-name',
      target.container,
      '--name',
      `${target.prefix}/${blobName}`,
      '--file',
      localPath,
      '--output',
      'none',
    ],
    { stdio: 'inherit' },
  )
}

function azUpload(target, localPath, blobName, { contentType, noCache = false } = {}) {
  const args = [
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
    contentType,
    '--output',
    'none',
  ]
  if (noCache) args.push('--content-cache-control', 'no-cache')
  execFileSync('az', args, { stdio: 'inherit' })
}

/** re-point the installer alias: pull the promoted installer to the runner
 * and re-upload it under the alias name (same upload auth as the build
 * pipelines — a server-side copy cannot authenticate its private source) */
function repointAlias(target, tmpDir, sourceBlob, aliasBlob) {
  const localPath = path.join(tmpDir, sourceBlob)
  azDownload(target, sourceBlob, localPath)
  azUpload(target, localPath, aliasBlob, { contentType: contentTypeFor(aliasBlob) })
  fs.rmSync(localPath, { force: true })
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

  const aliases = platform.aliases.map(({ installer, alias }) => ({
    installerName: installer(version),
    alias,
  }))
  for (const { installerName, alias } of aliases) {
    if (!(await blobExists(`${target.base}/${installerName}`))) {
      fatal(`${installerName} not found on the CDN — cannot re-point ${alias}`)
    }
  }

  const stable = await fetchText(`${target.base}/${platform.feed}`)
  const stableVersion = stable ? ymlVersion(stable) : null
  const verdict = assertPromotable(version, stableVersion, force)
  if (!verdict.ok) fatal(`${platform.feed}: ${verdict.reason}`)

  console.log(`[promote-stable] ${platform.feed}: ${stableVersion ?? '(none)'} -> ${version}`)
  for (const { installerName, alias } of aliases) {
    console.log(`[promote-stable] ${alias} -> ${installerName}`)
  }
  if (dryRun) return

  // aliases first, feed last: a client reading the new feed mid-promote must
  // already find every artifact it references
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'promote-'))
  for (const { installerName, alias } of aliases) {
    repointAlias(target, tmpDir, installerName, alias)
  }
  const tmp = path.join(tmpDir, platform.feed)
  fs.writeFileSync(tmp, archived)
  azUpload(target, tmp, platform.feed, { contentType: 'text/yaml', noCache: true })
  console.log(`[promote-stable] published ${target.base}/${platform.feed}`)
}

async function main() {
  const requested = PLATFORMS.map((p) => ({ p, version: argValue(p.flag) })).filter(
    (r) => r.version,
  )
  if (requested.length === 0)
    fatal('nothing to do — pass --mac <version>, --win <version>, and/or --linux <version>')
  const target = channelTarget()
  if (!dryRun && !process.env.AZURE_STORAGE_CONNECTION_STRING) {
    fatal('AZURE_STORAGE_CONNECTION_STRING env not set')
  }
  for (const { p, version } of requested) await promoteOne(target, p, version)
}

main().catch((err) => fatal(err.message))
