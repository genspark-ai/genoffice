import { describe, expect, it, vi } from 'vitest'

// The harfbuzz wasm only resolves in electron builds; ground-truth measuring never shapes
vi.mock('../../../node_modules/harfbuzzjs/dist/harfbuzz.js', () => ({ default: vi.fn() }))
vi.mock('../../../node_modules/harfbuzzjs/dist/harfbuzz.wasm?asset', () => ({ default: '' }))

import { gtMeasure, refineComplexWidths } from '../src/main/shaped-metrics'

/** Run the renderer-side script against a stub page and record the canvas font of each measure. */
async function refineOn(page: { displayFontFamily?: (name: string) => string }) {
  const fonts: string[] = []
  const ctx = {
    font: '',
    measureText(_text: string) {
      fonts.push(this.font)
      return { width: 50 }
    },
  }
  const window = {
    slidesApi: { privateFontFaces: async () => [] },
    ...(page.displayFontFamily ? { __genofficeDisplayFontFamily: page.displayFontFamily } : {}),
  }
  const document = {
    fonts: { forEach: () => {} },
    createElement: () => ({ getContext: () => ctx }),
  }
  const wc = {
    executeJavaScript: (code: string) =>
      new Function('window', 'document', `return ${code}`)(window, document) as Promise<unknown>,
  }
  const added = await refineComplexWidths(wc)
  return { added, fonts }
}

describe('ground-truth measureText uses the drawing font stack', () => {
  it("measures with the renderer's displayFontFamily stack", async () => {
    expect(gtMeasure('Latin', 'NanumGothic', 14)).toBeNull()
    const { added, fonts } = await refineOn({
      displayFontFamily: (n) => `'${n}', 'Malgun Gothic', 'Apple SD Gothic Neo', sans-serif`,
    })
    expect(added).toBe(true)
    expect(fonts).toEqual([
      "100px 'NanumGothic', 'Malgun Gothic', 'Apple SD Gothic Neo', sans-serif",
    ])
    // cached: the second layout pass reads the measured width (50 per 100px)
    expect(gtMeasure('Latin', 'NanumGothic', 14)).toBe(7)
  })

  it('keeps the generic stack when the page exposes no drawing stack', async () => {
    expect(gtMeasure('Cell', 'Some Face', 10, true)).toBeNull()
    const { fonts } = await refineOn({})
    expect(fonts).toEqual([
      "bold 100px \"Some Face\", 'PingFang SC', 'Microsoft YaHei', 'Yu Gothic', 'Malgun Gothic', sans-serif",
    ])
  })
})
