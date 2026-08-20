#!/usr/bin/env node
/**
 * Generates the KĀRYA icon set for all six apps.
 *
 * Design: rounded-square tile, indigo→violet diagonal gradient, white K
 * monogram built from three rounded bars (stem + two arms). The same geometry
 * is mirrored in the in-app mark (apps/shell/src/renderer/src/assets/
 * karya-logo.svg), so the window/installer icon and the in-app logo match.
 *
 * Emits (per app): build/icon.png (1024), build/icon.ico, build/icon.icns
 * Shell additionally: build/icon-mac.png and build/icons/{16..1024}x1024.png
 * (the linux hicolor set electron-builder's linux target consumes).
 *
 * Pure Node — pngjs for encoding; ICO/ICNS containers are written by hand
 * (PNG-compressed entries, the modern standard both OSes accept).
 */
import pngjs from 'pngjs'
import { writeFileSync, mkdirSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const { PNG } = pngjs
const root = join(dirname(fileURLToPath(import.meta.url)), '..')

// ── Design ───────────────────────────────────────────────
const INDIGO = [79, 70, 229] // #4F46E5
const VIOLET = [139, 92, 246] // #8B5CF6
const WHITE = [255, 255, 255]
const RADIUS = 0.225 // tile corner radius as a fraction of size

// K monogram: three capsules in unit space (centers of the rounded bars).
const BARS = [
  // stem: vertical, slightly left of center
  { x1: 0.3, y1: 0.245, x2: 0.3, y2: 0.755, r: 0.066 },
  // upper arm: rises from the stem's right edge to the top-right
  { x1: 0.372, y1: 0.345, x2: 0.735, y2: 0.222, r: 0.058 },
  // lower arm: falls from the stem's right edge to the bottom-right
  { x1: 0.372, y1: 0.515, x2: 0.735, y2: 0.778, r: 0.058 },
]

function sdSegment(px, py, x1, y1, x2, y2) {
  const dx = x2 - x1
  const dy = y2 - y1
  const len2 = dx * dx + dy * dy
  let t = len2 === 0 ? 0 : ((px - x1) * dx + (py - y1) * dy) / len2
  t = Math.max(0, Math.min(1, t))
  const ex = x1 + t * dx - px
  const ey = y1 + t * dy - py
  return Math.sqrt(ex * ex + ey * ey)
}

function insideTile(px, py) {
  const h = 0.5 - RADIUS
  const qx = Math.abs(px - 0.5) - h
  const qy = Math.abs(py - 0.5) - h
  const ax = Math.max(qx, 0)
  const ay = Math.max(qy, 0)
  return ax * ax + ay * ay <= RADIUS * RADIUS
}

/** [r,g,b,255] for an opaque sample, or null outside the tile */
function sampleColor(px, py) {
  if (!insideTile(px, py)) return null
  const t = (px + py) / 2 // diagonal gradient parameter
  const bg = [
    Math.round(INDIGO[0] + (VIOLET[0] - INDIGO[0]) * t),
    Math.round(INDIGO[1] + (VIOLET[1] - INDIGO[1]) * t),
    Math.round(INDIGO[2] + (VIOLET[2] - INDIGO[2]) * t),
  ]
  for (const bar of BARS) {
    if (sdSegment(px, py, bar.x1, bar.y1, bar.x2, bar.y2) <= bar.r) {
      return [...WHITE, 255]
    }
  }
  return [...bg, 255]
}

/** Render one size with SS×SS supersampling per pixel. */
function renderPNG(size, ss = 4) {
  const img = new PNG({ width: size, height: size })
  const step = 1 / (size * ss)
  let idx = 0
  for (let y = 0; y < size; y++) {
    for (let x = 0; x < size; x++) {
      const x0 = x / size
      const y0 = y / size
      let r = 0
      let g = 0
      let b = 0
      let n = 0
      for (let sy = 0; sy < ss; sy++) {
        for (let sx = 0; sx < ss; sx++) {
          const c = sampleColor(x0 + (sx + 0.5) * step, y0 + (sy + 0.5) * step)
          if (c) {
            r += c[0]
            g += c[1]
            b += c[2]
            n++
          }
        }
      }
      const total = ss * ss
      img.data[idx] = n > 0 ? Math.round(r / n) : 0
      img.data[idx + 1] = n > 0 ? Math.round(g / n) : 0
      img.data[idx + 2] = n > 0 ? Math.round(b / n) : 0
      img.data[idx + 3] = Math.round((n / total) * 255)
      idx += 4
    }
  }
  return PNG.sync.write(img)
}

// ── Containers ───────────────────────────────────────────

/** Windows .ico with PNG-compressed entries (16..256). */
function buildICO(pngFor) {
  const sizes = [16, 24, 32, 48, 64, 128, 256]
  const header = Buffer.alloc(6)
  header.writeUInt16LE(0, 0) // reserved
  header.writeUInt16LE(1, 2) // type: icon
  header.writeUInt16LE(sizes.length, 4)
  let offset = 6 + 16 * sizes.length
  const dir = []
  const chunks = []
  for (const s of sizes) {
    const buf = pngFor(s)
    const entry = Buffer.alloc(16)
    entry.writeUInt8(s >= 256 ? 0 : s, 0) // width (0 = 256)
    entry.writeUInt8(s >= 256 ? 0 : s, 1) // height
    entry.writeUInt8(0, 2) // palette
    entry.writeUInt8(0, 3) // reserved
    entry.writeUInt16LE(1, 4) // planes
    entry.writeUInt16LE(32, 6) // bpp
    entry.writeUInt32LE(buf.length, 8)
    entry.writeUInt32LE(offset, 12)
    dir.push(entry)
    chunks.push(buf)
    offset += buf.length
  }
  return Buffer.concat([header, ...dir, ...chunks])
}

/** macOS .icns with PNG-compressed entries covering 16..1024 (+@2x). */
function buildICNS(pngFor) {
  const entries = [
    ['icp4', 16],
    ['icp5', 32],
    ['icp6', 64],
    ['ic07', 128],
    ['ic08', 256],
    ['ic09', 512],
    ['ic10', 1024],
    ['ic11', 64], // 32@2x
    ['ic12', 512], // 256@2x
    ['ic13', 1024], // 512@2x
    ['ic14', 32], // 16@2x
  ]
  const chunks = []
  for (const [type, s] of entries) {
    const buf = pngFor(s)
    const head = Buffer.alloc(8)
    head.write(type, 0, 'ascii')
    head.writeUInt32BE(buf.length + 8, 4)
    chunks.push(head, buf)
  }
  const body = Buffer.concat(chunks)
  const header = Buffer.alloc(8)
  header.write('icns', 0, 'ascii')
  header.writeUInt32BE(body.length + 8, 4)
  return Buffer.concat([header, body])
}

// ── Emit ─────────────────────────────────────────────────

const cache = new Map()
const pngFor = (size) => {
  if (!cache.has(size)) cache.set(size, renderPNG(size))
  return cache.get(size)
}

const apps = ['docs', 'slides', 'sheets', 'pdf', 'markdown', 'shell']
for (const app of apps) {
  const dir = join(root, 'apps', app, 'build')
  mkdirSync(dir, { recursive: true })
  writeFileSync(join(dir, 'icon.png'), pngFor(1024))
  writeFileSync(join(dir, 'icon.ico'), buildICO(pngFor))
  writeFileSync(join(dir, 'icon.icns'), buildICNS(pngFor))
  console.log(`wrote ${dir}/icon.{png,ico,icns}`)
}

// shell extras: mac png + linux hicolor set (names match the existing layout)
writeFileSync(join(root, 'apps/shell/build/icon-mac.png'), pngFor(1024))
const iconsDir = join(root, 'apps/shell/build/icons')
mkdirSync(iconsDir, { recursive: true })
for (const s of [16, 32, 48, 128, 256, 512, 1024]) {
  writeFileSync(join(iconsDir, `${s}x${s}.png`), pngFor(s))
}
console.log('wrote shell build/icons/*.png and build/icon-mac.png')
