/**
 * Unit tests for the font alias map (the installed faces' name tables).
 *
 * Covers:
 *   - the alias map comes from the main process over IPC
 *   - the candidate the browser resolves is the one that is used
 *   - a declared name the browser resolves itself keeps its own weight
 *   - an alias that arrives later notifies the subscribers
 *   - the font picker's family list still comes from queryLocalFonts
 */

import type { ParsedDoc } from '@genoffice/docx-engine'
import { afterEach, describe, expect, it, vi } from 'vitest'

interface LocalFontFace {
  readonly family: string
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

/** the main process answers with the other spellings of the installed faces */
function stubFontAliases(
  answer: (names: readonly string[]) => Promise<Record<string, readonly string[]>>,
): void {
  ;(window as unknown as { desktop: { fontAliases: unknown } }).desktop = {
    fontAliases: answer,
  }
}

/** a fresh module graph: the font table cache and the alias lookup are
    module state, so each test starts from an empty one */
async function freshModules() {
  vi.resetModules()
  const systemFonts = await import('../src/renderer/system-fonts')
  const lineMetrics = await import('../src/renderer/line-metrics')
  const fontCheck = await import('../src/renderer/font-check')
  return { systemFonts, lineMetrics, fontCheck }
}

afterEach(() => {
  vi.restoreAllMocks()
  delete (window as unknown as { queryLocalFonts?: unknown }).queryLocalFonts
  delete (window as unknown as { desktop?: unknown }).desktop
})

// the font pipeline's module graph is heavy: import it once at file scope so
// the transform is paid while the file is collected, not inside a test's
// timeout (the dynamic imports below re-evaluate it, but the transform is
// already cached)
await import('../src/renderer/line-metrics')

describe('font aliases', () => {
  it('the alias map comes from the main process', async () => {
    // the browser resolves the English spelling but not the localized one
    stubCanvas(['FangSong_GB2312'])
    const { systemFonts, lineMetrics } = await freshModules()
    stubFontAliases(async () => ({ 仿宋_GB2312: ['FangSong_GB2312'] }))
    await systemFonts.resolveFontAliases(['仿宋_GB2312'])
    expect(head(lineMetrics.cssFontFamily('仿宋_GB2312'))).toBe('FangSong_GB2312')
  })

  it('the first candidate the browser resolves is the one that is used', async () => {
    stubCanvas(['FangSong_GB2312'])
    const { systemFonts, lineMetrics } = await freshModules()
    // the canonical spelling comes first, but only the second one resolves
    stubFontAliases(async () => ({ 仿宋_GB2312: ['FangSong_GB2312', 'No Such Face'] }))
    await systemFonts.resolveFontAliases(['仿宋_GB2312'])
    expect(head(lineMetrics.cssFontFamily('仿宋_GB2312'))).toBe('FangSong_GB2312')
  })

  it('a declared name stays when no candidate resolves', async () => {
    stubCanvas([])
    const { systemFonts, lineMetrics } = await freshModules()
    stubFontAliases(async () => ({ 仿宋_GB2312: ['No Such Face'] }))
    await systemFonts.resolveFontAliases(['仿宋_GB2312'])
    expect(head(lineMetrics.cssFontFamily('仿宋_GB2312'))).toBe('仿宋_GB2312')
  })

  it('a declared name the browser resolves itself keeps its own weight', async () => {
    // the alias map offers the family, and the browser resolves both, but the
    // declared name must not collapse to the family
    stubCanvas(['PingFang SC Semibold', 'PingFang SC'])
    const { systemFonts, lineMetrics } = await freshModules()
    stubFontAliases(async () => ({ 'PingFang SC Semibold': ['PingFang SC'] }))
    await systemFonts.resolveFontAliases(['PingFang SC Semibold'])
    expect(head(lineMetrics.cssFontFamily('PingFang SC Semibold'))).toBe('PingFang SC Semibold')
  })

  it('an alias that arrives later notifies the subscribers', async () => {
    stubCanvas(['FangSong_GB2312'])
    const { systemFonts } = await freshModules()
    const seen: string[][] = []
    const off = systemFonts.onFontAliasesChanged((names) => seen.push([...names]))
    stubFontAliases(async () => ({ 仿宋_GB2312: ['FangSong_GB2312'] }))
    await systemFonts.resolveFontAliases(['仿宋_GB2312'])
    expect(seen).toEqual([['仿宋_GB2312']])
    // the same answer announces again: a document laid out before the map was
    // filled must still be re-styled
    await systemFonts.resolveFontAliases(['仿宋_GB2312'])
    expect(seen).toEqual([['仿宋_GB2312'], ['仿宋_GB2312']])
    off()
    // an unsubscribed listener hears nothing more
    await systemFonts.resolveFontAliases(['仿宋_GB2312'])
    expect(seen).toEqual([['仿宋_GB2312'], ['仿宋_GB2312']])
  })

  it('an alias the document does not declare covers nothing', async () => {
    stubCanvas(['FangSong_GB2312'])
    const { systemFonts, fontCheck } = await freshModules()
    const seen: string[][] = []
    systemFonts.onFontAliasesChanged((names) => seen.push([...names]))
    stubFontAliases(async () => ({ 仿宋_GB2312: ['FangSong_GB2312'] }))
    await systemFonts.resolveFontAliases(['仿宋_GB2312'])
    expect(seen).toEqual([['仿宋_GB2312']])
    const parsed = {
      blocks: [{ runs: [{ font: 'SimSun' }] }],
      styles: new Map(),
      docDefaults: {},
    } as unknown as ParsedDoc
    expect(fontCheck.docFontsCoveredBy(parsed, seen[0]!)).toEqual([])
  })

  it('the headless export can wait for the alias request to settle', async () => {
    stubCanvas(['FangSong_GB2312'])
    const { systemFonts } = await freshModules()
    // the answer lands after the open, so the export must be able to wait
    let release: (() => void) | undefined
    const gate = new Promise<void>((resolve) => {
      release = resolve
    })
    stubFontAliases(async () => {
      await gate
      return { 仿宋_GB2312: ['FangSong_GB2312'] }
    })
    const request = systemFonts.resolveFontAliases(['仿宋_GB2312'])
    let settled = false
    const waited = systemFonts.fontAliasesSettled().then(() => {
      settled = true
    })
    await new Promise((resolve) => setTimeout(resolve, 0))
    // the answer has not landed yet, so the export still waits
    expect(settled).toBe(false)
    release?.()
    await request
    await waited
    expect(settled).toBe(true)
  })

  it('only a declared font counts as covered', async () => {
    const { fontCheck } = await freshModules()
    const parsed = {
      blocks: [{ runs: [{ font: '仿宋_GB2312' }] }],
      styles: new Map(),
      docDefaults: {},
    } as unknown as ParsedDoc
    expect(fontCheck.docFontsCoveredBy(parsed, ['仿宋_GB2312'])).toEqual(['仿宋_GB2312'])
    // an alias for a font the document does not declare must not reflow it
    expect(fontCheck.docFontsCoveredBy(parsed, ['PingFang SC'])).toEqual([])
  })
})

describe('system font table', () => {
  it("the picker's family list still comes from queryLocalFonts", async () => {
    const { systemFonts } = await freshModules()
    stubQueryLocalFonts(async () => [{ family: 'PingFang SC' }, { family: 'Helvetica Neue' }])
    expect(await systemFonts.loadSystemFontFamilies()).toEqual(['Helvetica Neue', 'PingFang SC'])
  })

  it('a rejected query does not poison the cache: a later query fills it', async () => {
    const { systemFonts } = await freshModules()
    let calls = 0
    // the query rejects without transient user activation, which is what a
    // document open runs into
    stubQueryLocalFonts(async () => {
      calls += 1
      throw new Error('NotAllowedError: must be handling a user gesture')
    })
    expect(await systemFonts.loadSystemFontFamilies()).toEqual([])
    expect(calls).toBe(1)
    // the rejected query is not cached: the next query (the font picker's own
    // click) runs and fills the cache
    stubQueryLocalFonts(async () => {
      calls += 1
      return [{ family: 'PingFang SC' }]
    })
    expect(await systemFonts.loadSystemFontFamilies()).toEqual(['PingFang SC'])
    expect(calls).toBe(2)
  })
})
