#!/usr/bin/env node
/**
 * Snapshot the published rhwp-studio (0.8.6 pages build) for offline embed.
 * Runtime never talks to github.io — this script is the only network step.
 */
import { createWriteStream } from 'node:fs'
import { mkdir, rm, writeFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { Readable } from 'node:stream'
import { pipeline } from 'node:stream/promises'
import { fileURLToPath } from 'node:url'

const ORIGIN = 'https://edwardkim.github.io'
const PREFIX = '/rhwp/'
const OUT = join(dirname(fileURLToPath(import.meta.url)), '..', 'vendor', 'rhwp-studio')

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
    if (!path.startsWith(PREFIX)) continue
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
    const text = await res.text()
    await writeFile(dest, text)
    return text
  }
  await pipeline(Readable.fromWeb(res.body), createWriteStream(dest))
  return ''
}

async function main() {
  await rm(OUT, { recursive: true, force: true })
  await mkdir(OUT, { recursive: true })
  const extras = [
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
    'icons/icon-128.png',
    'icons/icon-192.png',
    'icons/icon-512.png',
  ]
  const queue = [
    PREFIX,
    `${PREFIX}index.html`,
    `${PREFIX}manifest.webmanifest`,
    ...extras.map((path) => `${PREFIX}${path}`),
  ]
  const seen = new Set()
  while (queue.length) {
    const path = queue.pop()
    if (!path || seen.has(path)) continue
    seen.add(path)
    try {
      const text = await download(path)
      process.stdout.write(`  ${path}\n`)
      if (text) for (const next of discover(text)) queue.push(next)
    } catch (err) {
      process.stderr.write(`skip ${path}: ${err instanceof Error ? err.message : err}\n`)
    }
  }
  process.stdout.write(`vendored ${seen.size} files → ${OUT}\n`)
}

await main()
