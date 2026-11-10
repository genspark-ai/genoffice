import { describe, expect, it } from 'vitest'
import {
  exportImageMime,
  exportImageMimeFromPath,
  remoteExportImageMime,
  sniffExportImageMime,
} from '../src/shared/export-image-mime'
import {
  FALLBACK_IMAGE_SIZE,
  MAX_RASTER_DIM_PX,
  decodeImageDataUrl,
  fitDocxWidth,
  rasterSize,
  toDocxImage,
} from '../src/renderer/export/exportImage'
import type { ImageDecoder } from '../src/renderer/export/exportImage'

const PNG_HEADER = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
const SVG =
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 800 200"><rect width="800" height="200"/></svg>'
const svgBase64 = Buffer.from(SVG).toString('base64')

describe('export image mime', () => {
  it('covers every displayable format and normalizes aliases', () => {
    expect(exportImageMimeFromPath('demo.svg')).toBe('image/svg+xml')
    expect(exportImageMimeFromPath('https://x.test/a/pic.WEBP?w=1#f')).toBe('image/webp')
    expect(exportImageMimeFromPath('pic.tiff')).toBeNull()
    expect(exportImageMime('image/jpg')).toBe('image/jpeg')
    expect(exportImageMime('Image/SVG+XML; charset=utf-8')).toBe('image/svg+xml')
    expect(exportImageMime('text/html')).toBeNull()
  })

  it('remote bytes win over the header, the header over the URL', () => {
    expect(remoteExportImageMime(PNG_HEADER, 'text/plain', 'https://x.test/a.svg')).toBe(
      'image/png',
    )
    const text = new TextEncoder().encode(SVG)
    expect(remoteExportImageMime(text, 'image/svg+xml', 'https://x.test/a')).toBe('image/svg+xml')
    expect(remoteExportImageMime(text, 'application/octet-stream', 'https://x.test/a.svg')).toBe(
      'image/svg+xml',
    )
    expect(remoteExportImageMime(text, 'text/html', 'https://x.test/a')).toBeNull()
    expect(sniffExportImageMime(new Uint8Array([0x42, 0x4d, 0, 0]))).toBe('image/bmp')
  })
})

describe('data URL images', () => {
  it('decodes base64 and percent-encoded payloads', () => {
    expect(decodeImageDataUrl(`data:image/svg+xml;base64,${svgBase64}`)).toEqual({
      base64: svgBase64,
      mime: 'image/svg+xml',
    })
    const utf8 = decodeImageDataUrl(`data:image/svg+xml;charset=utf-8,${encodeURIComponent(SVG)}`)
    expect(utf8?.mime).toBe('image/svg+xml')
    expect(Buffer.from(utf8!.base64, 'base64').toString()).toBe(SVG)
  })

  it('keeps a bare % literal in an unencoded payload the way browsers do', () => {
    const raw = '<svg xmlns="http://www.w3.org/2000/svg" width="100%" viewBox="0 0 1 1"/>'
    const decoded = decodeImageDataUrl(`data:image/svg+xml,${raw}`)
    expect(Buffer.from(decoded!.base64, 'base64').toString()).toBe(raw)
  })

  it('rejects other schemes, unknown mimes and broken base64', () => {
    expect(decodeImageDataUrl('https://x.test/a.png')).toBeNull()
    expect(decodeImageDataUrl('data:text/plain;base64,aGk=')).toBeNull()
    expect(decodeImageDataUrl('data:image/png;base64,@@@')).toBeNull()
  })
})

describe('docx image sizing', () => {
  it('shrinks to the text column and keeps the aspect ratio', () => {
    expect(fitDocxWidth({ width: 1240, height: 600 }, 620)).toEqual({ widthPx: 620, heightPx: 300 })
    expect(fitDocxWidth({ width: 300, height: 100 }, 620)).toEqual({ widthPx: 300, heightPx: 100 })
  })

  it('vectors raster at 2x doc size, bitmaps never upscale, both stay under the cap', () => {
    const doc = { widthPx: 620, heightPx: 155 }
    expect(rasterSize(doc, { width: 8000, height: 2000 }, true)).toEqual({
      width: 1240,
      height: 310,
    })
    expect(rasterSize(doc, { width: 700, height: 175 }, false)).toEqual({ width: 700, height: 175 })
    const tall = rasterSize({ widthPx: 100, heightPx: 9000 }, { width: 100, height: 9000 }, true)
    expect(Math.max(tall.width, tall.height)).toBeLessThanOrEqual(MAX_RASTER_DIM_PX)
  })
})

describe('toDocxImage', () => {
  const measured =
    (width: number, height: number): ImageDecoder =>
    async () => ({ width, height, source: {} as CanvasImageSource })
  const undecodable: ImageDecoder = async () => null

  it('passes native bytes through, scaled to the column', async () => {
    const image = await toDocxImage(
      { base64: 'iVBORw0KGgo=', mime: 'image/png' },
      620,
      measured(1240, 600),
    )
    expect(image).toEqual({
      base64: 'iVBORw0KGgo=',
      mime: 'image/png',
      widthPx: 620,
      heightPx: 300,
    })
  })

  it('an undecodable native picture still lands at the fallback size', async () => {
    const image = await toDocxImage({ base64: 'iVBORw0KGgo=', mime: 'image/png' }, 620, undecodable)
    expect(image).toMatchObject({ mime: 'image/png', widthPx: FALLBACK_IMAGE_SIZE.width })
  })

  it('an undecodable SVG yields null so the caller keeps the alt text', async () => {
    expect(
      await toDocxImage({ base64: svgBase64, mime: 'image/svg+xml' }, 620, undecodable),
    ).toBeNull()
    expect(
      await toDocxImage({ base64: '!!!', mime: 'image/svg+xml' }, 620, measured(1, 1)),
    ).toBeNull()
  })

  it('hands the decoder a size-pinned SVG so it does not decode at 300x150', async () => {
    let seen = ''
    const spy: ImageDecoder = async (dataUrl) => {
      seen = Buffer.from(dataUrl.slice(dataUrl.indexOf(',') + 1), 'base64').toString()
      return null
    }
    await toDocxImage({ base64: svgBase64, mime: 'image/svg+xml' }, 620, spy)
    expect(seen).toMatch(/^<svg width="800" height="200"/)
  })
})
