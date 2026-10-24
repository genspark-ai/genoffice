import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { inflateSync } from 'node:zlib'
import { describe, expect, it } from 'vitest'

// The native new-tab menu pairs a 1x PNG with an @2x PNG and lets Electron pick
// the representation by scale factor. A fully transparent 1x file therefore
// renders as an invisible icon at 100% display scaling while 200% looks fine
// (#2133). These checks decode the shipped assets directly so an empty or
// miscopied icon cannot come back unnoticed.
const ASSETS = fileURLToPath(new URL('../src/main/assets/', import.meta.url))

const MENU_ICONS = ['docx', 'xlsx', 'pptx', 'md', 'html', 'pdf', 'home'] as const

type DecodedPng = { width: number; height: number; rgba: Buffer }

function decodePng(buf: Buffer): DecodedPng {
  expect(buf.subarray(0, 8).equals(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]))).toBe(true)
  let width = 0
  let height = 0
  let colorType = 0
  let palette: Buffer | null = null
  let transparency: Buffer | null = null
  const idat: Buffer[] = []
  let pos = 8
  while (pos < buf.length) {
    const length = buf.readUInt32BE(pos)
    const type = buf.toString('ascii', pos + 4, pos + 8)
    const data = buf.subarray(pos + 8, pos + 8 + length)
    if (type === 'IHDR') {
      width = data.readUInt32BE(0)
      height = data.readUInt32BE(4)
      colorType = data[9]
      expect(data[8]).toBe(8) // 8-bit depth only
      expect(data[12]).toBe(0) // no interlacing
    } else if (type === 'PLTE') palette = data
    else if (type === 'tRNS') transparency = data
    else if (type === 'IDAT') idat.push(data)
    else if (type === 'IEND') break
    pos += 12 + length
  }
  const channels = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 }[colorType] ?? 0
  expect(channels).toBeGreaterThan(0)
  const raw = inflateSync(Buffer.concat(idat))
  const stride = width * channels
  const rgba = Buffer.alloc(width * height * 4)
  let previous = Buffer.alloc(stride)
  for (let y = 0; y < height; y++) {
    const filter = raw[y * (stride + 1)]
    const current = Buffer.from(raw.subarray(y * (stride + 1) + 1, (y + 1) * (stride + 1)))
    for (let i = 0; i < stride; i++) {
      const left = i >= channels ? current[i - channels] : 0
      const up = previous[i]
      const upperLeft = i >= channels ? previous[i - channels] : 0
      if (filter === 1) current[i] = (current[i] + left) & 0xff
      else if (filter === 2) current[i] = (current[i] + up) & 0xff
      else if (filter === 3) current[i] = (current[i] + ((left + up) >> 1)) & 0xff
      else if (filter === 4) {
        const predictor = left + up - upperLeft
        const pa = Math.abs(predictor - left)
        const pb = Math.abs(predictor - up)
        const pc = Math.abs(predictor - upperLeft)
        current[i] = (current[i] + (pa <= pb && pa <= pc ? left : pb <= pc ? up : upperLeft)) & 0xff
      }
    }
    for (let x = 0; x < width; x++) {
      const target = (y * width + x) * 4
      const source = x * channels
      if (colorType === 6) rgba.set(current.subarray(source, source + 4), target)
      else if (colorType === 2) {
        rgba[target] = current[source]
        rgba[target + 1] = current[source + 1]
        rgba[target + 2] = current[source + 2]
        rgba[target + 3] = 255
      } else if (colorType === 4) {
        rgba[target] = rgba[target + 1] = rgba[target + 2] = current[source]
        rgba[target + 3] = current[source + 1]
      } else if (colorType === 0) {
        rgba[target] = rgba[target + 1] = rgba[target + 2] = current[source]
        rgba[target + 3] = 255
      } else if (colorType === 3 && palette) {
        const index = current[source] * 3
        rgba[target] = palette[index]
        rgba[target + 1] = palette[index + 1]
        rgba[target + 2] = palette[index + 2]
        rgba[target + 3] = transparency ? (transparency[current[source]] ?? 255) : 255
      }
    }
    previous = current
  }
  return { width, height, rgba }
}

function visiblePixels(png: DecodedPng): number {
  let count = 0
  for (let i = 3; i < png.rgba.length; i += 4) if (png.rgba[i] > 8) count++
  return count
}

// The 1x icon is the 2x2 box average of its @2x sibling — the relationship the
// intact icons already follow, and what the regenerated pdf/home files restore.
function boxDownsample(png: DecodedPng): DecodedPng {
  const width = png.width / 2
  const height = png.height / 2
  const rgba = Buffer.alloc(width * height * 4)
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      let red = 0
      let green = 0
      let blue = 0
      let alphaSum = 0
      for (let dy = 0; dy < 2; dy++) {
        for (let dx = 0; dx < 2; dx++) {
          const source = ((y * 2 + dy) * png.width + x * 2 + dx) * 4
          const alpha = png.rgba[source + 3] / 255
          red += png.rgba[source] * alpha
          green += png.rgba[source + 1] * alpha
          blue += png.rgba[source + 2] * alpha
          alphaSum += png.rgba[source + 3]
        }
      }
      const alpha = alphaSum / 4
      const target = (y * width + x) * 4
      rgba[target] = alpha > 0 ? Math.round(red / (alpha / 255) / 4) : 0
      rgba[target + 1] = alpha > 0 ? Math.round(green / (alpha / 255) / 4) : 0
      rgba[target + 2] = alpha > 0 ? Math.round(blue / (alpha / 255) / 4) : 0
      rgba[target + 3] = Math.round(alpha)
    }
  }
  return { width, height, rgba }
}

function averageDifference(a: DecodedPng, b: DecodedPng): number {
  let total = 0
  for (let i = 0; i < a.rgba.length; i++) total += Math.abs(a.rgba[i] - b.rgba[i])
  return total / a.rgba.length
}

describe('new-tab menu icons', () => {
  it('ships a visible icon at both scale factors', () => {
    for (const name of MENU_ICONS) {
      for (const [suffix, size] of [
        ['', 16],
        ['@2x', 32],
      ] as const) {
        const png = decodePng(readFileSync(`${ASSETS}menu-${name}${suffix}.png`))
        expect(png.width, `menu-${name}${suffix}.png width`).toBe(size)
        expect(png.height, `menu-${name}${suffix}.png height`).toBe(size)
        // The thinnest glyph in the set (the home outline) keeps 83 visible
        // pixels at 1x; the broken assets shipped with 0.
        expect(
          visiblePixels(png),
          `menu-${name}${suffix}.png visible pixels`,
        ).toBeGreaterThanOrEqual(40)
      }
    }
  })

  it('keeps every 1x icon distinct', () => {
    const seen = new Map<string, string>()
    for (const name of MENU_ICONS) {
      const bytes = readFileSync(`${ASSETS}menu-${name}.png`).toString('base64')
      expect(seen.has(bytes), `menu-${name}.png duplicates menu-${seen.get(bytes)}`).toBe(false)
      seen.set(bytes, name)
    }
  })

  it('derives each 1x icon from its @2x sibling', () => {
    for (const name of MENU_ICONS) {
      const oneX = decodePng(readFileSync(`${ASSETS}menu-${name}.png`))
      const downsampled = boxDownsample(decodePng(readFileSync(`${ASSETS}menu-${name}@2x.png`)))
      // The intact icons sit at ~1-2 per channel against a box downsample;
      // anything shipping a different drawing at 1x would land far higher.
      expect(
        averageDifference(oneX, downsampled),
        `menu-${name}.png vs @2x downsample`,
      ).toBeLessThan(16)
    }
  })
})
