import { afterEach, describe, expect, it, vi } from 'vitest'
import { parseDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { checkMissingFonts, collectDocFonts } from '../src/renderer/font-check'

/** fake 2d context: known families get distinct widths, unknown ones inherit the generic fallback */
function stubCanvas(availableFamilies: string[]) {
  const known = new Map(availableFamilies.map((f, i) => [f, 100 + (i + 1) * 10]))
  let current = 50
  const fake = {
    set font(spec: string) {
      const family = /"([^"]+)"/.exec(spec)?.[1]
      current =
        family !== undefined && known.has(family)
          ? known.get(family)!
          : spec.includes('monospace')
            ? 50
            : 60
    },
    measureText: () => ({ width: current }),
  }
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(
    fake as unknown as CanvasRenderingContext2D,
  )
}

afterEach(() => {
  vi.restoreAllMocks()
})

describe('collectDocFonts', () => {
  it('gathers distinct run fonts from the parsed model', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:r><w:rPr><w:rFonts w:eastAsia="仿宋_GB2312"/></w:rPr><w:t>正文</w:t></w:r>' +
          '<w:r><w:rPr><w:rFonts w:eastAsia="楷体_GB2312"/></w:rPr><w:t>（承办）</w:t></w:r></w:p>' +
          '<w:p><w:r><w:rPr><w:rFonts w:eastAsia="仿宋_GB2312"/></w:rPr><w:t>再来一段</w:t></w:r></w:p>',
      }),
    )
    const fonts = collectDocFonts(doc)
    expect(fonts).toContain('仿宋_GB2312')
    expect(fonts).toContain('楷体_GB2312')
    expect(fonts.filter((f) => f === '仿宋_GB2312')).toHaveLength(1)
  })
})

describe('checkMissingFonts', () => {
  it('reports a missing font with its first available substitute', () => {
    stubCanvas(['STFangsong', 'STKaiti'])
    // fresh module per stub: the 2d context is cached module-wide
    expect(checkMissingFonts(['仿宋_GB2312'])).toEqual([
      { name: '仿宋_GB2312', substitute: 'STFangsong' },
    ])
  })

  it('reports null substitute when only the bundled fallback is left', () => {
    stubCanvas([])
    expect(checkMissingFonts(['楷体_GB2312'])).toEqual([{ name: '楷体_GB2312', substitute: null }])
  })

  it('says nothing when the declared font resolves', () => {
    stubCanvas(['黑体'])
    expect(checkMissingFonts(['黑体'])).toEqual([])
  })
})
