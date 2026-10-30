#!/usr/bin/env node
/**
 * Wallpaper generator.
 *
 * The Home desktop ships procedurally drawn wallpapers from
 * apps/shell/build/wallpapers (electron-builder maps it to resources/wallpapers).
 * Each one is authored as a draw function below — brand hues, gradients, arcs,
 * star fields — rendered at 1920x1080 through the preinstalled napi canvas
 * bindings (@napi-rs/canvas, the same 2D-canvas API as the `canvas` package,
 * already present in node_modules as pdfjs's optional dependency). Generated
 * PNGs and the wallpapers.json manifest the desktop IPC reads are committed,
 * like the repo's other generated fixtures.
 *
 *   node tools/gen-wallpapers.mjs
 */
import { mkdirSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { createCanvas } from '@napi-rs/canvas'

const HERE = dirname(fileURLToPath(import.meta.url))
const ROOT = join(HERE, '..')
const OUT_DIR = join(ROOT, 'apps', 'shell', 'build', 'wallpapers')
const MANIFEST = 'wallpapers.json'

/** every wallpaper renders at the same 16:9 canvas */
const WIDTH = 1920
const HEIGHT = 1080

/** the manifest's public shape — the id the renderer addresses wallpapers by */
export const WALLPAPERS = [
  { id: 'brand-aurora', nameKey: 'wpNameAurora', filename: 'brand-aurora.png' },
  { id: 'starfield-night', nameKey: 'wpNameStarfield', filename: 'starfield-night.png' },
  { id: 'warm-dusk', nameKey: 'wpNameDusk', filename: 'warm-dusk.png' },
  { id: 'dot-grid-dark', nameKey: 'wpNameDots', filename: 'dot-grid-dark.png' },
  { id: 'minimal-light', nameKey: 'wpNameMinimal', filename: 'minimal-light.png' },
  { id: 'teal-deep', nameKey: 'wpNameTeal', filename: 'teal-deep.png' },
]

export class GenWallpapersError extends Error {}

/** deterministic PRNG (mulberry32): a rerun must byte-match the committed PNGs */
function seededRandom(seed) {
  let state = seed >>> 0
  return () => {
    state = (state + 0x6d2b79f5) >>> 0
    let t = state
    t = Math.imul(t ^ (t >>> 15), t | 1)
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61)
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296
  }
}

/** the four brand hues from apps/shell/build/logo.svg (see desktop.css) */
const BRAND_BLUE = '#0f7fff'
const BRAND_GREEN = '#16a34a'
const BRAND_ORANGE = '#f97316'
const BRAND_RED = '#ef4444'

/** a soft radial glow of `color` centered at (cx, cy) with the given radius */
function glow(ctx, cx, cy, radius, color, alpha) {
  const gradient = ctx.createRadialGradient(cx, cy, 0, cx, cy, radius)
  gradient.addColorStop(0, hexA(color, alpha))
  gradient.addColorStop(0.55, hexA(color, alpha * 0.45))
  gradient.addColorStop(1, hexA(color, 0))
  ctx.fillStyle = gradient
  ctx.fillRect(cx - radius, cy - radius, radius * 2, radius * 2)
}

/** hex + alpha into an rgba() fill color */
function hexA(hex, alpha) {
  const value = hex.slice(1)
  const r = Number.parseInt(value.slice(0, 2), 16)
  const g = Number.parseInt(value.slice(2, 4), 16)
  const b = Number.parseInt(value.slice(4, 6), 16)
  return `rgba(${r}, ${g}, ${b}, ${Math.max(0, Math.min(1, alpha)).toFixed(4)})`
}

// ── the six wallpapers ──────────────────────────────────

/** four-hue soft glows drifting over a deep blue base, like the live aurora */
function drawBrandAurora(ctx, w, h, random) {
  ctx.fillStyle = '#0b1526'
  ctx.fillRect(0, 0, w, h)
  ctx.fillStyle = '#101d33'
  ctx.fillRect(0, h * 0.62, w, h * 0.38)
  glow(ctx, w * 0.16, h * 0.14, w * 0.34, BRAND_BLUE, 0.5)
  glow(ctx, w * 0.84, h * 0.2, w * 0.28, BRAND_GREEN, 0.34)
  glow(ctx, w * 0.5, h * 0.86, w * 0.36, BRAND_RED, 0.24)
  glow(ctx, w * 0.9, h * 0.78, w * 0.3, BRAND_ORANGE, 0.3)
  // a faint second echo of each glow, lower and dimmer, for depth
  for (let i = 0; i < 3; i += 1) {
    const x = random() * w
    const y = h * (0.55 + random() * 0.35)
    glow(ctx, x, y, w * (0.1 + random() * 0.12), BRAND_BLUE, 0.08)
  }
}

/** near-black night sky: a milky band, then stars in three brightness tiers */
function drawStarfieldNight(ctx, w, h, random) {
  ctx.fillStyle = '#05070d'
  ctx.fillRect(0, 0, w, h)
  // the milky way: a diagonal band of pale haze
  ctx.save()
  ctx.translate(w * 0.5, h * 0.5)
  ctx.rotate(-0.32)
  const band = ctx.createLinearGradient(0, -h * 0.5, 0, h * 0.5)
  band.addColorStop(0, 'rgba(38, 50, 76, 0)')
  band.addColorStop(0.5, 'rgba(70, 90, 130, 0.32)')
  band.addColorStop(1, 'rgba(38, 50, 76, 0)')
  ctx.fillStyle = band
  ctx.fillRect(-w, -h * 0.34, w * 2, h * 0.68)
  ctx.restore()
  glow(ctx, w * 0.72, h * 0.24, w * 0.2, '#3b5bff', 0.12)
  glow(ctx, w * 0.2, h * 0.8, w * 0.24, '#ff5b8a', 0.07)
  // stars: mostly pinpoints, some mid, a few bright with a cross flare
  const star = (x, y, r, alpha) => {
    ctx.fillStyle = `rgba(226, 236, 255, ${alpha})`
    ctx.beginPath()
    ctx.arc(x, y, r, 0, Math.PI * 2)
    ctx.fill()
  }
  for (let i = 0; i < 620; i += 1) {
    star(random() * w, random() * h, 0.5 + random() * 0.9, 0.25 + random() * 0.5)
  }
  for (let i = 0; i < 90; i += 1) {
    star(random() * w, random() * h, 1.4 + random() * 1.1, 0.55 + random() * 0.35)
  }
  for (let i = 0; i < 12; i += 1) {
    const x = random() * w
    const y = random() * h
    const r = 2.2 + random() * 1.6
    glow(ctx, x, y, r * 7, '#9db4ff', 0.28)
    star(x, y, r, 0.95)
    // cross flare on the brightest
    ctx.strokeStyle = 'rgba(214, 228, 255, 0.5)'
    ctx.lineWidth = 0.8
    ctx.beginPath()
    ctx.moveTo(x - r * 4, y)
    ctx.lineTo(x + r * 4, y)
    ctx.moveTo(x, y - r * 4)
    ctx.lineTo(x, y + r * 4)
    ctx.stroke()
  }
}

/** warm dusk gradient with geometric arcs rising like a sun setting */
function drawWarmDusk(ctx, w, h) {
  const sky = ctx.createLinearGradient(0, 0, 0, h)
  sky.addColorStop(0, '#2b1a3d')
  sky.addColorStop(0.42, '#6e2f4f')
  sky.addColorStop(0.68, '#c65a3c')
  sky.addColorStop(0.85, '#e88a4a')
  sky.addColorStop(1, '#f4b066')
  ctx.fillStyle = sky
  ctx.fillRect(0, 0, w, h)
  // the setting sun's glow at the horizon
  glow(ctx, w * 0.5, h * 0.88, w * 0.34, '#ffd9a0', 0.5)
  // concentric arcs over the sun, alternating stroke tints
  const cx = w * 0.5
  const cy = h * 0.88
  for (let i = 0; i < 7; i += 1) {
    const radius = w * (0.16 + i * 0.075)
    ctx.strokeStyle = hexA(i % 2 === 0 ? '#ffd9a0' : '#ffffff', 0.22 - i * 0.02)
    ctx.lineWidth = 3.5
    ctx.beginPath()
    ctx.arc(cx, cy, radius, Math.PI, Math.PI * 2)
    ctx.stroke()
  }
  // a dark hill silhouette across the bottom
  ctx.fillStyle = '#241526'
  ctx.beginPath()
  ctx.moveTo(0, h)
  ctx.lineTo(0, h * 0.9)
  ctx.quadraticCurveTo(w * 0.3, h * 0.78, w * 0.55, h * 0.92)
  ctx.quadraticCurveTo(w * 0.8, h * 1.02, w, h * 0.86)
  ctx.lineTo(w, h)
  ctx.closePath()
  ctx.fill()
}

/** dark slate with a fading dot grid and a few accent dots, workmanlike */
function drawDotGridDark(ctx, w, h, random) {
  ctx.fillStyle = '#14181f'
  ctx.fillRect(0, 0, w, h)
  const sheen = ctx.createRadialGradient(w * 0.7, h * 0.25, 0, w * 0.7, h * 0.25, w * 0.7)
  sheen.addColorStop(0, 'rgba(52, 70, 96, 0.55)')
  sheen.addColorStop(1, 'rgba(20, 24, 31, 0)')
  ctx.fillStyle = sheen
  ctx.fillRect(0, 0, w, h)
  // the grid, dimming toward the edges (the mirror of desktop.css's mask)
  const step = 34
  for (let y = step / 2; y < h; y += step) {
    for (let x = step / 2; x < w; x += step) {
      const dx = (x - w * 0.5) / (w * 0.62)
      const dy = (y - h * 0.45) / (h * 0.62)
      const fade = Math.max(0, 1 - (dx * dx + dy * dy) * 0.8)
      ctx.fillStyle = `rgba(148, 166, 190, ${(0.1 + fade * 0.22).toFixed(3)})`
      ctx.beginPath()
      ctx.arc(x, y, 1.6, 0, Math.PI * 2)
      ctx.fill()
    }
  }
  // a scatter of brand-hue accent dots
  const accents = [BRAND_BLUE, BRAND_GREEN, BRAND_ORANGE, BRAND_RED]
  for (let i = 0; i < 26; i += 1) {
    const x = random() * w
    const y = random() * h
    const color = accents[Math.floor(random() * accents.length)]
    glow(ctx, x, y, 8 + random() * 14, color, 0.4)
    ctx.fillStyle = hexA(color, 0.8)
    ctx.beginPath()
    ctx.arc(x, y, 1.4 + random() * 1.6, 0, Math.PI * 2)
    ctx.fill()
  }
}

/** the light-theme one: near-white paper, one soft band, thin quiet rules */
function drawMinimalLight(ctx, w, h) {
  ctx.fillStyle = '#f4f5f7'
  ctx.fillRect(0, 0, w, h)
  const tint = ctx.createLinearGradient(0, 0, w, h)
  tint.addColorStop(0, '#eef1f5')
  tint.addColorStop(0.5, '#f7f8fa')
  tint.addColorStop(1, '#eef0f4')
  ctx.fillStyle = tint
  ctx.fillRect(0, 0, w, h)
  // one soft blue band across the upper third
  const band = ctx.createLinearGradient(0, 0, w, 0)
  band.addColorStop(0, 'rgba(15, 127, 255, 0)')
  band.addColorStop(0.5, 'rgba(15, 127, 255, 0.16)')
  band.addColorStop(1, 'rgba(22, 163, 74, 0.08)')
  ctx.fillStyle = band
  ctx.fillRect(0, h * 0.16, w, h * 0.2)
  // three thin rules, like a letterhead
  ctx.strokeStyle = 'rgba(100, 116, 139, 0.35)'
  ctx.lineWidth = 2
  for (const at of [0.52, 0.62, 0.72]) {
    ctx.beginPath()
    ctx.moveTo(w * 0.12, h * at)
    ctx.lineTo(w * (0.12 + 0.5 * (1 - (at - 0.52) * 3)), h * at)
    ctx.stroke()
  }
  // a single accent dot where the first rule ends
  ctx.fillStyle = 'rgba(15, 127, 255, 0.8)'
  ctx.beginPath()
  ctx.arc(w * 0.62, h * 0.52, 7, 0, Math.PI * 2)
  ctx.fill()
}

/** deep teal sea: stacked wave curves lightening toward the horizon */
function drawTealDeep(ctx, w, h) {
  const sea = ctx.createLinearGradient(0, 0, 0, h)
  sea.addColorStop(0, '#03211f')
  sea.addColorStop(0.45, '#05403c')
  sea.addColorStop(1, '#0a6b60')
  ctx.fillStyle = sea
  ctx.fillRect(0, 0, w, h)
  glow(ctx, w * 0.3, h * 0.2, w * 0.3, '#0f7fff', 0.14)
  // wave curves: sine strokes with growing amplitude downward
  for (let i = 0; i < 9; i += 1) {
    const baseY = h * (0.3 + i * 0.08)
    const amp = 14 + i * 9
    const phase = i * 1.1
    const alpha = 0.1 + i * 0.035
    ctx.strokeStyle = `rgba(140, 235, 220, ${alpha.toFixed(3)})`
    ctx.lineWidth = 2.5 + i * 0.6
    ctx.beginPath()
    for (let x = 0; x <= w; x += 8) {
      const y = baseY + Math.sin(x / 260 + phase) * amp
      if (x === 0) ctx.moveTo(x, y)
      else ctx.lineTo(x, y)
    }
    ctx.stroke()
  }
}

/** draw functions by wallpaper id; each gets (ctx, width, height, rng) */
const DRAWERS = {
  'brand-aurora': drawBrandAurora,
  'starfield-night': drawStarfieldNight,
  'warm-dusk': drawWarmDusk,
  'dot-grid-dark': drawDotGridDark,
  'minimal-light': drawMinimalLight,
  'teal-deep': drawTealDeep,
}

/** a stable seed per id, so every rerun draws the identical picture */
const seedFor = (id) => {
  let hash = 2166136261
  for (const ch of id) hash = Math.imul(hash ^ ch.charCodeAt(0), 16777619)
  return hash >>> 0
}

async function renderWallpaper(entry) {
  const canvas = createCanvas(WIDTH, HEIGHT)
  const ctx = canvas.getContext('2d')
  const draw = DRAWERS[entry.id]
  if (!draw) throw new GenWallpapersError(`${entry.id}: no draw function`)
  draw(ctx, WIDTH, HEIGHT, seededRandom(seedFor(entry.id)))
  return canvas.encode('png')
}

/** every generated byte must be a real PNG of the promised size */
function validateOutput(entry, bytes) {
  if (bytes.length < 50_000) {
    throw new GenWallpapersError(`${entry.filename}: suspiciously small (${bytes.length} bytes)`)
  }
  const signature = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]
  if (!signature.every((byte, i) => bytes[i] === byte)) {
    throw new GenWallpapersError(`${entry.filename}: not a PNG (bad signature)`)
  }
  // IHDR's width/height are big-endian at offsets 16..23
  const width = (bytes[16] << 24) | (bytes[17] << 16) | (bytes[18] << 8) | bytes[19]
  const height = (bytes[20] << 24) | (bytes[21] << 16) | (bytes[22] << 8) | bytes[23]
  if (width !== WIDTH || height !== HEIGHT) {
    throw new GenWallpapersError(
      `${entry.filename}: ${width}x${height}, expected ${WIDTH}x${HEIGHT}`,
    )
  }
}

async function main() {
  mkdirSync(OUT_DIR, { recursive: true })
  for (const entry of WALLPAPERS) {
    process.stdout.write(`building ${entry.filename}… `)
    const bytes = await renderWallpaper(entry)
    validateOutput(entry, bytes)
    writeFileSync(join(OUT_DIR, entry.filename), bytes)
    console.log(`ok (${(bytes.length / 1024).toFixed(0)} KiB)`)
  }
  const manifest = WALLPAPERS.map(({ id, nameKey, filename }) => ({ id, nameKey, filename }))
  writeFileSync(join(OUT_DIR, MANIFEST), `${JSON.stringify(manifest, null, 2)}\n`, 'utf8')
  console.log(`wrote ${MANIFEST}: ${manifest.length} wallpapers in apps/shell/build/wallpapers`)
}

const invokedDirectly = process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href
if (invokedDirectly) {
  main().catch((err) => {
    console.error(`gen-wallpapers: ${err instanceof Error ? err.message : String(err)}`)
    process.exit(1)
  })
}
