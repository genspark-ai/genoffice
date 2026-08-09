#!/usr/bin/env node
/**
 * scripts/linux-release-upload.cjs — publish a built Linux release to the
 * CDN update channel (counterpart of mac-release-upload.cjs; normally run
 * by linux-build.yml, but works for local publishes the same way).
 *
 * Flow:
 *   1. npm run dist:linux -- -c.extraMetadata.version=0.5.X  (with
 *      GENOFFICE_UPDATE_URL in the environment / electron-builder.env so the
 *      publish config is baked and latest-linux.yml is generated)
 *   2. node scripts/linux-release-upload.cjs [--channel <name>] [--dry-run] [--force]
 *
 * --channel selects the update feed name (default 'latest'); produces
 * <channel>-linux.yml on the CDN — the name electron-updater derives on
 * Linux x64 for that channel.
 *
 * Uploads from apps/shell/release:
 *   - GenOffice-<v>.AppImage      auto-update payload (electron-updater's
 *                                 AppImageUpdater replaces the file in place)
 *   - genoffice_<v>_amd64.deb     apt install package; listed in the feed but
 *                                 deb installs run no updater — uploaded so
 *                                 every feed entry resolves
 *   - latest-linux.yml            auto-update feed (uploaded LAST, no-cache);
 *                                 url: entries rewritten to absolute CDN URLs,
 *                                 with a versioned archive copy
 *                                 GenOffice-linux-<v>.yml kept alongside for
 *                                 the promote workflow / rollback
 *
 * On the latest channel the artifacts are also aliased to GenOffice.AppImage
 * / GenOffice.deb (stable marketing links); beta publishes leave the aliases
 * to the promote workflow — same split as the dmg/exe aliases.
 *
 * Safety: refuses to upload unless the local latest-linux.yml version is
 * strictly newer than the version currently served by the CDN (--force
 * overrides, e.g. for rollback).
 *
 * Requires: az CLI + AZURE_STORAGE_CONNECTION_STRING env + GENOFFICE_UPDATE_URL
 * (https://<cdn-host>/<container>/<prefix>; container/prefix for az uploads
 * are derived from its path — same convention as the other release scripts).
 */

const { execFileSync } = require('child_process')
const fs = require('fs')
const path = require('path')
const https = require('https')
const { ymlVersion, semverNewer } = require('./update-feed-utils.cjs')

const PRODUCT = 'GenOffice'
const RELEASE_DIR = path.resolve(__dirname, '..', 'apps/shell/release')

const dryRun = process.argv.includes('--dry-run')
const force = process.argv.includes('--force')

function fatal(msg) {
  console.error(`[linux-release] ERROR: ${msg}`)
  process.exit(1)
}

function channelFlag() {
  const idx = process.argv.indexOf('--channel')
  const name = idx === -1 ? 'latest' : process.argv[idx + 1]
  if (!name || !/^[a-z]+$/.test(name)) fatal(`invalid --channel "${name}"`)
  return name
}

const CHANNEL = channelFlag()
const FEED_NAME = `${CHANNEL}-linux.yml`

// container + blob prefix are derived from the channel URL's path
// (https://<cdn-host>/<container>/<prefix>) so a single env var drives both
// the public download base and the az upload target
function channelTarget() {
  const raw = process.env.GENOFFICE_UPDATE_URL
  if (!raw) fatal('GENOFFICE_UPDATE_URL env not set (public base URL of the update channel)')
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

const { container: CONTAINER, prefix: PREFIX, base: CDN_BASE } = channelTarget()

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

/** rewrite relative url: entries to absolute CDN URLs so the feed stays
 * valid no matter where it is fetched from (same as mac-release-upload) */
function rewriteYmlUrls(text) {
  return text.replace(/^(\s*(?:-\s+)?url:\s*)(\S+)$/gm, (line, head, url) =>
    /^https?:\/\//.test(url) ? line : `${head}${CDN_BASE}/${url}`,
  )
}

function contentTypeFor(name) {
  if (name.endsWith('.AppImage')) return 'application/vnd.appimage'
  if (name.endsWith('.deb')) return 'application/vnd.debian.binary-package'
  if (name.endsWith('.yml')) return 'text/yaml'
  return 'application/octet-stream'
}

function azUpload(localPath, blobName) {
  const args = [
    'storage',
    'blob',
    'upload',
    '--container-name',
    CONTAINER,
    '--file',
    localPath,
    '--name',
    `${PREFIX}/${blobName}`,
    '--overwrite',
    '--content-type',
    contentTypeFor(blobName),
    '--output',
    'none',
  ]
  if (blobName.endsWith('.yml')) args.push('--content-cache-control', 'no-cache')
  console.log(`[linux-release] uploading ${blobName} ...`)
  execFileSync('az', args, { stdio: 'inherit' })
}

async function main() {
  const ymlPath = path.join(RELEASE_DIR, 'latest-linux.yml')
  if (!fs.existsSync(ymlPath))
    fatal(
      `latest-linux.yml not found in ${RELEASE_DIR} — package with GENOFFICE_UPDATE_URL set first`,
    )
  const yml = fs.readFileSync(ymlPath, 'utf8')
  const version = ymlVersion(yml)
  if (!version) fatal('could not parse version from latest-linux.yml')

  const files = fs
    .readdirSync(RELEASE_DIR)
    .filter((f) => (f.endsWith('.AppImage') || f.endsWith('.deb')) && !f.endsWith('.blockmap'))
    .filter((f) => f.includes(version))
  const appImages = files.filter((f) => f.endsWith('.AppImage'))
  const debs = files.filter((f) => f.endsWith('.deb'))
  if (appImages.length === 0)
    fatal(`no ${version} AppImage in ${RELEASE_DIR}; it is required for auto-update`)
  // every file referenced by the feed must actually be uploaded
  for (const m of yml.matchAll(/url:\s*(\S+)/g)) {
    if (!files.includes(m[1]))
      fatal(`latest-linux.yml references ${m[1]} which is not in ${RELEASE_DIR}`)
  }

  console.log(`[linux-release] local version:  ${version}`)
  const remoteYml = await fetchText(`${CDN_BASE}/${FEED_NAME}`)
  const remoteVersion = remoteYml ? ymlVersion(remoteYml) : null
  console.log(`[linux-release] CDN version:    ${remoteVersion ?? '(none published yet)'}`)
  if (remoteVersion && !semverNewer(version, remoteVersion) && !force) {
    fatal(
      `local ${version} is not newer than published ${remoteVersion}; ` +
        `pass --force to roll back`,
    )
  }

  const versionedYml = `${PRODUCT}-linux-${version}.yml`
  console.log(
    `[linux-release] will upload: ${[...appImages, ...debs, versionedYml, FEED_NAME].join(', ')}`,
  )
  if (dryRun) {
    console.log('[linux-release] dry run — nothing uploaded')
    return
  }
  if (!process.env.AZURE_STORAGE_CONNECTION_STRING) {
    fatal('AZURE_STORAGE_CONNECTION_STRING env not set')
  }

  // the served feed carries absolute CDN download URLs; keep the local
  // electron-builder original untouched and upload from a rewritten copy
  const cdnYmlPath = path.join(RELEASE_DIR, 'latest-linux.cdn.yml')
  fs.writeFileSync(cdnYmlPath, rewriteYmlUrls(yml))

  // payloads first, feed last — an updater that reads the new feed mid-upload
  // must already find the binaries it points to
  for (const f of [...appImages, ...debs]) azUpload(path.join(RELEASE_DIR, f), f)
  // the GenOffice.AppImage / GenOffice.deb marketing aliases must track the
  // STABLE channel only: beta publishes skip them, and the promote workflow
  // re-points them to the promoted version (scripts/promote-stable.cjs)
  if (CHANNEL === 'latest') {
    for (const f of appImages) azUpload(path.join(RELEASE_DIR, f), `${PRODUCT}.AppImage`)
    for (const f of debs) azUpload(path.join(RELEASE_DIR, f), `${PRODUCT}.deb`)
  }
  azUpload(cdnYmlPath, versionedYml) // immutable archive copy (promote/rollback)
  azUpload(cdnYmlPath, FEED_NAME)

  console.log('')
  console.log(`[linux-release] published v${version}`)
  if (CHANNEL === 'latest') {
    console.log(`[linux-release] stable AppImage: ${CDN_BASE}/${PRODUCT}.AppImage`)
    console.log(`[linux-release] stable deb:      ${CDN_BASE}/${PRODUCT}.deb`)
  }
  console.log(`[linux-release] update feed: ${CDN_BASE}/${FEED_NAME}`)
  console.log(`[linux-release] feed archive: ${CDN_BASE}/${versionedYml}`)
}

main().catch((err) => fatal(err.message))
