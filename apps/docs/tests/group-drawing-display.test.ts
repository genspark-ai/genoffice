import { describe, expect, it } from 'vitest'
import type { TextboxDisplay } from '@genoffice/docx-engine'
import { pictureLayerDecls } from '../src/renderer/editor/image-rotation'
import { textboxFloatPlacement } from '../src/renderer/editor/protected-render'

const PNG = 'data:image/png;base64,AAAA'

describe('cropped picture layers', () => {
  const urls = (decls: string[]): string[] =>
    Array.from(decls[0].matchAll(/url\("([^"]*)"\)/g), (m) => m[1])

  it('an uncropped layer paints its bitmap as is', () => {
    const decls = pictureLayerDecls([{ dataUrl: PNG, x: 0, y: 0, w: 1, h: 1 }])
    expect(urls(decls)).toEqual([PNG])
  })

  it('a cropped layer shows only the kept window, at the placed proportions', () => {
    const decls = pictureLayerDecls([
      {
        dataUrl: PNG,
        x: 0,
        y: 0,
        w: 1,
        h: 0.5,
        crop: { l: 0.5, t: 0, r: 0, b: 0.25 },
        aspect: 2,
      },
    ])
    const [url] = urls(decls)
    expect(url.startsWith('data:image/svg+xml,')).toBe(true)
    const svg = decodeURIComponent(url.slice('data:image/svg+xml,'.length))
    // kept window 2000 x 1000 of a bitmap drawn 4000 x 1333.33, starting halfway across
    expect(svg).toContain('width="2000" height="1000" viewBox="2000 0 2000 1000"')
    expect(svg).toContain(`<image href="${PNG}" width="4000" height="1333.33"`)
  })

  it('a bitmap served by URL stays uncropped: an SVG image cannot load it', () => {
    const served = 'genoffice-docx-media://hash/word/media/image1.png'
    const decls = pictureLayerDecls([
      { dataUrl: served, x: 0, y: 0, w: 1, h: 1, crop: { l: 0.5, t: 0, r: 0, b: 0 }, aspect: 1 },
    ])
    expect(urls(decls)).toEqual([served])
  })
})

describe('vertical line placement', () => {
  const line = (extra: Partial<TextboxDisplay>): TextboxDisplay => ({
    paras: [],
    readOnly: true,
    floating: true,
    prst: 'line',
    widthPx: 12,
    heightPx: 100,
    offsetXEmu: 952500,
    offsetYEmu: 0,
    ...extra,
  })
  const leftOf = (css: string): number => Number(/, (-?[\d.]+)px,/.exec(css)?.[1])

  it('centres the band on the authored x', () => {
    const plain = leftOf(textboxFloatPlacement(line({})))
    const vert = leftOf(textboxFloatPlacement(line({ lineVert: true })))
    expect(plain).toBe(100)
    expect(vert).toBe(94)
  })
})
