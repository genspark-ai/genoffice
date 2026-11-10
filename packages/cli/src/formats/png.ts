import { PNG } from 'pngjs'

export interface PxRect {
  x: number
  y: number
  w: number
  h: number
}

/** Crop a PNG to `rect` grown by `pad` on each side, clamped to the image. */
export function cropPng(png: Buffer, rect: PxRect, pad = 0): { png: Buffer; rect: PxRect } {
  const src = PNG.sync.read(png)
  const x0 = clamp(Math.floor(rect.x - pad), 0, src.width - 1)
  const y0 = clamp(Math.floor(rect.y - pad), 0, src.height - 1)
  const x1 = clamp(Math.ceil(rect.x + rect.w + pad), x0 + 1, src.width)
  const y1 = clamp(Math.ceil(rect.y + rect.h + pad), y0 + 1, src.height)
  const out = new PNG({ width: x1 - x0, height: y1 - y0 })
  for (let y = y0; y < y1; y++) {
    src.data.copy(
      out.data,
      (y - y0) * out.width * 4,
      (y * src.width + x0) * 4,
      (y * src.width + x1) * 4,
    )
  }
  return { png: PNG.sync.write(out), rect: { x: x0, y: y0, w: x1 - x0, h: y1 - y0 } }
}

export interface GridTile {
  page: number
  x: number
  y: number
  w: number
  h: number
}

export interface GridSheet {
  png: Buffer
  width: number
  height: number
  cols: number
  rows: number
  tiles: GridTile[]
}

const GUTTER = 8

/** Pages downscaled to `tileWidth` and laid out row-major on white, `cols` per row. */
export function contactSheet(
  pages: { page: number; png: Buffer }[],
  cols: number,
  tileWidth: number,
): GridSheet {
  const decoded = pages.map((p) => ({ page: p.page, img: PNG.sync.read(p.png) }))
  const tileHeight = Math.max(
    1,
    ...decoded.map((d) => Math.round((d.img.height * tileWidth) / d.img.width)),
  )
  const rows = Math.ceil(decoded.length / cols)
  const width = cols * tileWidth + (cols + 1) * GUTTER
  const height = rows * tileHeight + (rows + 1) * GUTTER
  const sheet = new PNG({ width, height })
  sheet.data.fill(255)
  const tiles: GridTile[] = []
  decoded.forEach((d, i) => {
    const col = i % cols
    const row = Math.floor(i / cols)
    const w = tileWidth
    const h = Math.round((d.img.height * tileWidth) / d.img.width)
    const x = GUTTER + col * (tileWidth + GUTTER)
    const y = GUTTER + row * (tileHeight + GUTTER)
    downscaleInto(d.img, sheet, x, y, w, h)
    tiles.push({ page: d.page, x, y, w, h })
  })
  return { png: PNG.sync.write(sheet), width, height, cols, rows, tiles }
}

/** Box-filter downscale of `src` into a `w`×`h` area of `dst` at (dx, dy). */
function downscaleInto(src: PNG, dst: PNG, dx: number, dy: number, w: number, h: number): void {
  const sx = src.width / w
  const sy = src.height / h
  for (let y = 0; y < h; y++) {
    const y0 = Math.floor(y * sy)
    const y1 = Math.max(y0 + 1, Math.floor((y + 1) * sy))
    for (let x = 0; x < w; x++) {
      const x0 = Math.floor(x * sx)
      const x1 = Math.max(x0 + 1, Math.floor((x + 1) * sx))
      let r = 0
      let g = 0
      let b = 0
      let n = 0
      for (let yy = y0; yy < y1; yy++) {
        for (let xx = x0; xx < x1; xx++) {
          const i = (yy * src.width + xx) * 4
          const a = src.data[i + 3]! / 255
          r += src.data[i]! * a + 255 * (1 - a)
          g += src.data[i + 1]! * a + 255 * (1 - a)
          b += src.data[i + 2]! * a + 255 * (1 - a)
          n++
        }
      }
      const o = ((dy + y) * dst.width + dx + x) * 4
      dst.data[o] = Math.round(r / n)
      dst.data[o + 1] = Math.round(g / n)
      dst.data[o + 2] = Math.round(b / n)
      dst.data[o + 3] = 255
    }
  }
}

function clamp(v: number, lo: number, hi: number): number {
  return Math.max(lo, Math.min(hi, v))
}
