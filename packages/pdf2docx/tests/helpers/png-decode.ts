/**
 * Decode our own encoder's output back to RGBA8: every colour type it may
 * pick (grey, RGB, indexed 1/2/4/8 bpp with optional tRNS, grey+alpha, RGBA),
 * filter 0 rows, one zlib stream.
 */
import { inflateSync } from 'node:zlib'

export function decodePngRgba(png: Uint8Array): {
  rgba: Uint8Array
  width: number
  height: number
} {
  const view = new DataView(png.buffer, png.byteOffset, png.byteLength)
  const width = view.getUint32(16)
  const height = view.getUint32(20)
  const bitDepth = png[24]!
  const colorType = png[25]!
  const idat: Uint8Array[] = []
  let plte: Uint8Array = new Uint8Array(0)
  let trns: Uint8Array = new Uint8Array(0)
  let pos = 8
  while (pos < png.length) {
    const len = view.getUint32(pos)
    const type = String.fromCharCode(png[pos + 4]!, png[pos + 5]!, png[pos + 6]!, png[pos + 7]!)
    const data = png.subarray(pos + 8, pos + 8 + len)
    if (type === 'IDAT') idat.push(data)
    else if (type === 'PLTE') plte = data
    else if (type === 'tRNS') trns = data
    pos += 12 + len
  }
  const zlib = new Uint8Array(idat.reduce((n, c) => n + c.length, 0))
  let off = 0
  for (const c of idat) {
    zlib.set(c, off)
    off += c.length
  }
  const raw: Uint8Array = inflateSync(zlib)
  const channels =
    colorType === 0 ? 1 : colorType === 2 ? 3 : colorType === 3 ? 1 : colorType === 4 ? 2 : 4
  const bytesPerRow = Math.ceil((width * channels * bitDepth) / 8)
  const rgba = new Uint8Array(width * height * 4)
  for (let y = 0; y < height; y++) {
    const row = raw.subarray(y * (1 + bytesPerRow) + 1, (y + 1) * (1 + bytesPerRow))
    for (let x = 0; x < width; x++) {
      const o = (y * width + x) * 4
      if (colorType === 3) {
        const perByte = 8 / bitDepth
        const idx =
          bitDepth === 8
            ? row[x]!
            : (row[Math.floor(x / perByte)]! >> (8 - bitDepth - (x % perByte) * bitDepth)) &
              ((1 << bitDepth) - 1)
        rgba[o] = plte[idx * 3]!
        rgba[o + 1] = plte[idx * 3 + 1]!
        rgba[o + 2] = plte[idx * 3 + 2]!
        rgba[o + 3] = idx < trns.length ? trns[idx]! : 255
      } else if (colorType === 0) {
        rgba[o] = rgba[o + 1] = rgba[o + 2] = row[x]!
        rgba[o + 3] = 255
      } else if (colorType === 4) {
        rgba[o] = rgba[o + 1] = rgba[o + 2] = row[x * 2]!
        rgba[o + 3] = row[x * 2 + 1]!
      } else if (colorType === 2) {
        rgba[o] = row[x * 3]!
        rgba[o + 1] = row[x * 3 + 1]!
        rgba[o + 2] = row[x * 3 + 2]!
        rgba[o + 3] = 255
      } else {
        rgba.set(row.subarray(x * 4, x * 4 + 4), o)
      }
    }
  }
  return { rgba, width, height }
}
