import { describe, it, expect } from 'vitest'
import { imageDpiFromDataUrl, imageSizeFromBytes } from '../src/image-dpi'
import { resolveFill } from '../src/fill'
import { makeViewport } from '../src/coords'
import type { Fill } from '@genoffice/pptx-engine'

function pngWithPhys(ppmX: number, ppmY: number, unit = 1, withPhys = true): string {
  const chunk = (type: string, data: number[]) => {
    const len = data.length
    return [
      (len >>> 24) & 255,
      (len >>> 16) & 255,
      (len >>> 8) & 255,
      len & 255,
      ...type.split('').map((c) => c.charCodeAt(0)),
      ...data,
      0,
      0,
      0,
      0,
    ]
  }
  const be32 = (v: number) => [(v >>> 24) & 255, (v >>> 16) & 255, (v >>> 8) & 255, v & 255]
  const bytes = [
    0x89,
    0x50,
    0x4e,
    0x47,
    0x0d,
    0x0a,
    0x1a,
    0x0a,
    ...chunk('IHDR', [...be32(2), ...be32(2), 8, 6, 0, 0, 0]),
    ...(withPhys ? chunk('pHYs', [...be32(ppmX), ...be32(ppmY), unit]) : []),
    ...chunk('IDAT', [0, 0]),
    ...chunk('IEND', []),
  ]
  return `data:image/png;base64,${Buffer.from(bytes).toString('base64')}`
}

function jfif(units: number, xd: number, yd: number): string {
  const bytes = [
    0xff,
    0xd8,
    0xff,
    0xe0,
    0x00,
    0x10,
    0x4a,
    0x46,
    0x49,
    0x46,
    0x00,
    0x01,
    0x01,
    units,
    (xd >> 8) & 255,
    xd & 255,
    (yd >> 8) & 255,
    yd & 255,
    0,
    0,
    0xff,
    0xda,
    0x00,
    0x02,
  ]
  return `data:image/jpeg;base64,${Buffer.from(bytes).toString('base64')}`
}

describe('image dpi tag', () => {
  it('reads PNG pHYs in pixels per metre', () => {
    const d = imageDpiFromDataUrl(pngWithPhys(5906, 5906))
    expect(d?.x).toBeCloseTo(150, 0)
    expect(d?.y).toBeCloseTo(150, 0)
  })

  it('ignores PNG pHYs that only states an aspect ratio, and untagged PNGs', () => {
    expect(imageDpiFromDataUrl(pngWithPhys(1, 1, 0))).toBeUndefined()
    expect(imageDpiFromDataUrl(pngWithPhys(1, 1, 1, false))).toBeUndefined()
    // implausible densities (1 px/m) are treated as untagged
    expect(imageDpiFromDataUrl(pngWithPhys(1, 1))).toBeUndefined()
  })

  it('reads JFIF density in dpi or dots per cm, none when units=0', () => {
    expect(imageDpiFromDataUrl(jfif(1, 75, 75))).toEqual({ x: 75, y: 75 })
    expect(imageDpiFromDataUrl(jfif(2, 100, 100))?.x).toBeCloseTo(254)
    expect(imageDpiFromDataUrl(jfif(0, 1, 1))).toBeUndefined()
  })

  it('tolerates non-image and non-base64 urls', () => {
    expect(imageDpiFromDataUrl(undefined)).toBeUndefined()
    expect(imageDpiFromDataUrl('data:image/svg+xml;utf8,<svg/>')).toBeUndefined()
    expect(imageDpiFromDataUrl('data:image/png;base64,AAAA')).toBeUndefined()
  })

  it('tile scale follows the image dpi tag, 144dpi when untagged', () => {
    const vp = makeViewport({ cx: 9525 * 1000, cy: 9525 * 1000 }, 1000)
    const f: Fill = {
      type: 'image',
      mediaRef: 'ppt/media/t.png',
      mode: 'tile',
      tile: { tx: 0, ty: 0, sx: 1, sy: 1, algn: 'tl' },
    }
    const tagged = resolveFill(f, vp, () => pngWithPhys(5906, 5906))
    if (tagged.kind === 'image') expect(tagged.tile?.scaleX).toBeCloseTo(96 / 150, 3)
    const texture = resolveFill(f, vp, () => jfif(1, 75, 75))
    if (texture.kind === 'image') expect(texture.tile?.scaleX).toBeCloseTo(96 / 75, 3)
    const untagged = resolveFill(f, vp, () => jfif(0, 1, 1))
    if (untagged.kind === 'image') expect(untagged.tile?.scaleX).toBeCloseTo(96 / 144, 3)
  })
})

describe('imageSizeFromBytes', () => {
  it('reads the PNG IHDR size', () => {
    const b64 = pngWithPhys(2835, 2835).split(',')[1]!
    expect(imageSizeFromBytes(Buffer.from(b64, 'base64'))).toEqual({ w: 2, h: 2 })
  })

  it('reads the first JPEG SOF frame after APP segments', () => {
    const sof0 = [
      0xff, 0xc0, 0x00, 0x11, 8, 0x01, 0x2c, 0x02, 0x80, 3, 1, 0x22, 0, 2, 0x11, 1, 3, 0x11, 1,
    ]
    const bytes = Buffer.concat([
      Buffer.from(jfif(1, 72, 72).split(',')[1]!, 'base64').subarray(0, 20),
      Buffer.from(sof0),
      Buffer.from([0xff, 0xda, 0, 2]),
    ])
    expect(imageSizeFromBytes(bytes)).toEqual({ w: 640, h: 300 })
  })

  it('returns undefined for other formats', () => {
    expect(imageSizeFromBytes(Buffer.from('GIF89a'.padEnd(30, '\0')))).toBeUndefined()
  })
})
