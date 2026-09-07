#!/usr/bin/env node
/**
 * Snapshot the published rhwp-studio (0.8.6 pages build) for offline embed.
 * Runtime never talks to github.io — this script is the only network step.
 *
 * `--ensure` skips the download when a complete snapshot is already present.
 */
import { createWriteStream, existsSync, readdirSync } from 'node:fs'
import { mkdir, readFile, rm, unlink, writeFile } from 'node:fs/promises'
import { dirname, extname, join } from 'node:path'
import { Readable } from 'node:stream'
import { pipeline } from 'node:stream/promises'
import { fileURLToPath } from 'node:url'
import {
  PWA_FILES,
  REQUIRED_ASSET_EXTS,
  REQUIRED_RELATIVE,
  eagerPagePrefetch,
  isPwaPath,
  stripPwaHtml,
} from './studio-snapshot.mjs'

const ORIGIN = 'https://edwardkim.github.io'
const PREFIX = '/rhwp/'
const OUT = join(dirname(fileURLToPath(import.meta.url)), '..', 'vendor', 'rhwp-studio')
const ENSURE = process.argv.includes('--ensure')

const TEXT_EXT = new Set([
  '.html',
  '.js',
  '.css',
  '.json',
  '.webmanifest',
  '.svg',
  '.txt',
  '.map',
])

const BUNDLED_FONTS = [
  'fonts/Cafe24Ssurround-v2.0.woff2',
  'fonts/Cafe24Supermagic-Regular-v1.0.woff2',
  'fonts/D2Coding-Regular.woff2',
  'fonts/GowunBatang-Regular.woff2',
  'fonts/GowunDodum-Regular.woff2',
  'fonts/Happiness-Sans-Bold.woff2',
  'fonts/Happiness-Sans-Regular.woff2',
  'fonts/Happiness-Sans-Title.woff2',
  'fonts/HappinessSansVF.woff2',
  'fonts/LatinModernMath-Regular.woff2',
  'fonts/NanumGothic-Regular.woff2',
  'fonts/NanumGothicCoding-Regular.woff2',
  'fonts/NanumMyeongjo-Regular.woff2',
  'fonts/NotoSansKR-Bold.woff2',
  'fonts/NotoSansKR-ExtraLight.woff2',
  'fonts/NotoSansKR-Regular.woff2',
  'fonts/NotoSerifKR-Bold.woff2',
  'fonts/NotoSerifKR-Regular.woff2',
  'fonts/Pretendard-Black.woff2',
  'fonts/Pretendard-Bold.woff2',
  'fonts/Pretendard-ExtraBold.woff2',
  'fonts/Pretendard-ExtraLight.woff2',
  'fonts/Pretendard-Light.woff2',
  'fonts/Pretendard-Medium.woff2',
  'fonts/Pretendard-Regular.woff2',
  'fonts/Pretendard-SemiBold.woff2',
  'fonts/Pretendard-Thin.woff2',
  'fonts/SourceHanSerifK-OldHangul-subset.woff2',
  'fonts/SpoqaHanSans-Regular.woff2',
]

function extOf(path) {
  const q = path.split('?')[0]
  const i = q.lastIndexOf('.')
  return i >= 0 ? q.slice(i).toLowerCase() : ''
}

function isTextPath(path) {
  return TEXT_EXT.has(extOf(path)) || path.endsWith('/') || !extOf(path)
}

function destFor(urlPath) {
  const clean = urlPath.split('?')[0]
  const rel = clean.slice(PREFIX.length)
  return join(OUT, rel === '' || rel.endsWith('/') ? `${rel}index.html` : rel)
}

function discover(text) {
  const found = new Set()
  const re = /(?:\/rhwp\/)[A-Za-z0-9._~:/?#[\]@!$&'()*+,;=%-]+/g
  for (const match of text.matchAll(re)) {
    let path = match[0]
    path = path.replace(/["')\s>].*$/, '')
    const cut = path.search(/[#?]/)
    if (cut >= 0) path = path.slice(0, cut)
    if (!path.startsWith(PREFIX) || isPwaPath(path)) continue
    found.add(path)
  }
  return found
}

async function download(urlPath) {
  const res = await fetch(`${ORIGIN}${urlPath}`)
  if (!res.ok || !res.body) throw new Error(`${res.status} ${urlPath}`)
  const dest = destFor(urlPath)
  await mkdir(dirname(dest), { recursive: true })
  if (isTextPath(urlPath)) {
    let text = await res.text()
    if (dest.endsWith('index.html')) text = stripPwaHtml(text)
    else if (dest.endsWith('.js')) text = tryEagerPrefetch(text, dest)
    await writeFile(dest, text)
    return text
  }
  await pipeline(Readable.fromWeb(res.body), createWriteStream(dest))
  return ''
}

async function stripPwaFiles() {
  for (const name of PWA_FILES) {
    const path = join(OUT, name)
    if (!existsSync(path)) continue
    await unlink(path)
  }
  const index = join(OUT, 'index.html')
  if (!existsSync(index)) return
  const next = stripPwaHtml(await readFile(index, 'utf8'))
  await writeFile(index, next)
}

function tryEagerPrefetch(js, label) {
  try {
    return eagerPagePrefetch(js)
  } catch (err) {
    process.stderr.write(
      `prefetch patch skipped (${label}): ${err instanceof Error ? err.message : err}\n`,
    )
    return js
  }
}

async function patchStudioJs() {
  const dir = join(OUT, 'assets')
  if (!existsSync(dir)) return
  for (const name of readdirSync(dir)) {
    if (extname(name) !== '.js') continue
    const path = join(dir, name)
    const next = tryEagerPrefetch(await readFile(path, 'utf8'), name)
    await writeFile(path, next)
  }
}

function missingRequired() {
  const missing = REQUIRED_RELATIVE.filter((rel) => !existsSync(join(OUT, rel)))
  const assets = existsSync(join(OUT, 'assets')) ? readdirSync(join(OUT, 'assets')) : []
  for (const ext of REQUIRED_ASSET_EXTS) {
    if (!assets.some((name) => extname(name) === ext)) missing.push(`assets/*${ext}`)
  }
  return missing
}

function isComplete() {
  if (missingRequired().length > 0) return false
  if (PWA_FILES.some((name) => existsSync(join(OUT, name)))) return false
  return true
}

async function vendor() {
  await rm(OUT, { recursive: true, force: true })
  await mkdir(OUT, { recursive: true })
  const queue = [PREFIX, `${PREFIX}index.html`, ...BUNDLED_FONTS.map((path) => `${PREFIX}${path}`)]
  const seen = new Set()
  const failed = []
  while (queue.length) {
    const path = queue.pop()
    if (!path || seen.has(path) || isPwaPath(path)) continue
    seen.add(path)
    try {
      const text = await download(path)
      process.stdout.write(`  ${path}\n`)
      if (text) for (const next of discover(text)) queue.push(next)
    } catch (err) {
      failed.push(`${path}: ${err instanceof Error ? err.message : err}`)
    }
  }
  await stripPwaFiles()
  await patchStudioJs()
  const missing = missingRequired()
  if (missing.length || failed.length) {
    const details = [...missing.map((rel) => `missing ${rel}`), ...failed]
    throw new Error(`rhwp-studio vendor failed:\n${details.join('\n')}`)
  }
  process.stdout.write(`vendored ${seen.size} files → ${OUT}\n`)
}

async function main() {
  if (ENSURE) {
    await stripPwaFiles()
    await patchStudioJs()
    if (isComplete()) {
      process.stdout.write(`rhwp-studio snapshot ready → ${OUT}\n`)
      return
    }
  }
  await vendor()
}

await main()
