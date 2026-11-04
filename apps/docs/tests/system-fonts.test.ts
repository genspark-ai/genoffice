/**
 * Unit tests for the system-font table bridge (the Local Font Access API).
 *
 * Covers:
 *   - a weight-qualified declared name keeps its own resolution
 *   - a declared name the browser resolves itself is not aliased to its family
 *   - a rejected queryLocalFonts() does not poison the cache
 */

import { afterEach, describe, expect, it, vi } from 'vitest'

interface LocalFontFace {
  readonly family: string
  readonly fullName: string
  readonly postscriptName: string
  readonly style?: string
}

/** the family at the head of a css font chain */
const head = (chain: string): string => chain.split(',')[0].replace(/['"]/g, '').trim()

/** A canvas 2d stub: a family in `resolvable` measures wider than the generic
    fallbacks, so isFontAvailable() reports it as present; every other family
    measures like the fallback, so it reads as absent. */
function stubCanvas(resolvable: readonly string[]): void {
  const present = new Set(resolvable)
  const ctx = {
    font: '',
    measureText(_text: string) {
      const quoted = /"([^"]+)"/.exec(ctx.font)
      const name = quoted
        ? quoted[1]!
        : ctx.font
            .replace(/^32px\s*/, '')
            .split(',')[0]!
            .trim()
      return { width: present.has(name) ? 200 : 100 } as TextMetrics
    },
  }
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(((contextId: string) =>
    contextId === '2d'
      ? (ctx as unknown as CanvasRenderingContext2D)
      : null) as unknown as HTMLCanvasElement['getContext'])
}

function stubQueryLocalFonts(query: () => Promise<readonly LocalFontFace[]>): void {
  ;(
    window as unknown as { queryLocalFonts?: () => Promise<readonly LocalFontFace[]> }
  ).queryLocalFonts = query
}

/** a fresh module graph: the font table cache and the alias lookup are
    module state, so each test starts from an empty one */
async function freshModules() {
  vi.resetModules()
  const systemFonts = await import('../src/renderer/system-fonts')
  const lineMetrics = await import('../src/renderer/line-metrics')
  return { systemFonts, lineMetrics }
}

afterEach(() => {
  vi.restoreAllMocks()
  delete (window as unknown as { queryLocalFonts?: unknown }).queryLocalFonts
})

describe('system font table', () => {
  it('a weight-qualified declared name keeps its own weight', async () => {
    // nothing resolves here, so the alias map is the only way a declared name
    // could change: the weight suffix must survive it
    stubCanvas([])
    const { systemFonts, lineMetrics } = await freshModules()
    stubQueryLocalFonts(async () => [
      {
        family: 'PingFang SC',
        fullName: 'PingFang SC',
        postscriptName: 'PingFangSC-Regular',
        style: 'Regular',
      },
      {
        family: 'PingFang SC',
        fullName: 'PingFang SC Semibold',
        postscriptName: 'PingFangSC-Semibold',
        style: 'Semibold',
      },
    ])
    await systemFonts.loadSystemFontAliases()
    expect(head(lineMetrics.cssFontFamily('PingFang SC Semibold'))).toBe('PingFang SC Semibold')
  })

  it('a declared name the browser resolves itself is not aliased to its family', async () => {
    // the face reports no style, so its full name reaches the alias map; the
    // browser resolves the name itself, so the alias must not apply
    stubCanvas(['Helvetica Neue Light'])
    const { systemFonts, lineMetrics } = await freshModules()
    stubQueryLocalFonts(async () => [
      {
        family: 'Helvetica Neue',
        fullName: 'Helvetica Neue Light',
        postscriptName: 'HelveticaNeue-Light',
      },
    ])
    await systemFonts.loadSystemFontAliases()
    expect(head(lineMetrics.cssFontFamily('Helvetica Neue Light'))).toBe('Helvetica Neue Light')
  })

  it('a localized name resolves through the English family the table maps it to', async () => {
    stubCanvas([])
    const { systemFonts, lineMetrics } = await freshModules()
    stubQueryLocalFonts(async () => [
      {
        family: 'FZFangSong-Z02',
        fullName: '\u65b9\u6b63\u4eff\u5b8b_GBK',
        postscriptName: 'FZFangSong-Z02',
        style: 'Regular',
      },
    ])
    await systemFonts.loadSystemFontAliases()
    expect(head(lineMetrics.cssFontFamily('\u65b9\u6b63\u4eff\u5b8b_GBK'))).toBe('FZFangSong-Z02')
  })

  it('a rejected query does not poison the cache: a later query fills it', async () => {
    stubCanvas([])
    const { systemFonts, lineMetrics } = await freshModules()
    let calls = 0
    // the query rejects without transient user activation, which is what a
    // document open runs into
    stubQueryLocalFonts(async () => {
      calls += 1
      throw new Error('NotAllowedError: must be handling a user gesture')
    })
    await systemFonts.loadSystemFontAliases()
    expect(calls).toBe(1)
    // the rejected query is not cached: the next query (the font picker's own
    // click) runs and fills the cache
    stubQueryLocalFonts(async () => {
      calls += 1
      return [
        {
          family: 'FZFangSong-Z02',
          fullName: '\u65b9\u6b63\u4eff\u5b8b_GBK',
          postscriptName: 'FZFangSong-Z02',
          style: 'Regular',
        },
      ]
    })
    await systemFonts.loadSystemFontAliases()
    expect(calls).toBe(2)
    expect(head(lineMetrics.cssFontFamily('\u65b9\u6b63\u4eff\u5b8b_GBK'))).toBe('FZFangSong-Z02')
  })
})
