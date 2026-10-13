/**
 * The strut alias takes the EA chain's canvas metrics. A bundled web face
 * that leads the chain but has not loaded yet makes canvas measure the next
 * installed family, so the result must not be cached and the face must be
 * requested; once it is in, callers re-derive the CSS.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  fontChainMetricsPct,
  takeSettledMetricFaces,
  unloadedBundledPrimary,
} from '../src/renderer/line-metrics'

const KR_CHAIN =
  '"맑은 고딕", "Malgun Gothic", "GenOffice Sans KR", "Apple SD Gothic Neo", sans-serif'

function stubFonts(loaded: Set<string>) {
  const load = vi.fn((spec: string) => {
    loaded.add(/"([^"]+)"/.exec(spec)?.[1] ?? '')
    return Promise.resolve([])
  })
  const fonts = { check: (spec: string) => loaded.has(/"([^"]+)"/.exec(spec)?.[1] ?? ''), load }
  Object.defineProperty(document, 'fonts', { value: fonts, configurable: true })
  return load
}

function stubCanvas(metrics: { ascent: number; descent: number }) {
  const fake = {
    font: '',
    measureText: () => ({
      width: 50,
      fontBoundingBoxAscent: metrics.ascent,
      fontBoundingBoxDescent: metrics.descent,
    }),
  }
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(
    fake as unknown as CanvasRenderingContext2D,
  )
}

afterEach(() => {
  vi.restoreAllMocks()
  Reflect.deleteProperty(document, 'fonts')
})

describe('unloadedBundledPrimary', () => {
  it('names the bundled face that leads a chain of missing families until it loads', () => {
    stubFonts(new Set())
    stubCanvas({ ascent: 900, descent: 300 })
    expect(unloadedBundledPrimary(KR_CHAIN)).toBe('GenOffice Sans KR')
  })

  it('null once the face is loaded, for generic-only chains and without FontFaceSet', () => {
    stubFonts(new Set(['GenOffice Sans KR']))
    expect(unloadedBundledPrimary(KR_CHAIN)).toBeNull()
    expect(unloadedBundledPrimary('serif')).toBeNull()
    Reflect.deleteProperty(document, 'fonts')
    expect(unloadedBundledPrimary(KR_CHAIN)).toBeNull()
  })
})

describe('fontChainMetricsPct with a pending bundled face', () => {
  it('requests the face, returns the fallback geometry uncached, then re-measures once loaded', () => {
    const loaded = new Set<string>()
    const load = stubFonts(loaded)
    const metrics = { ascent: 900, descent: 300 }
    stubCanvas(metrics)
    const chain = `${KR_CHAIN}, "probe-a"`
    expect(fontChainMetricsPct(chain)).toEqual({ ascentPct: 90, descentPct: 30 })
    expect(load).toHaveBeenCalledWith('16px "GenOffice Sans KR"')
    // the stub resolves the load synchronously: the face is in now
    expect(takeSettledMetricFaces()).toBe(true)
    expect(takeSettledMetricFaces()).toBe(false)
    metrics.ascent = 1160
    metrics.descent = 288
    expect(fontChainMetricsPct(chain)).toEqual({ ascentPct: 116, descentPct: 28.8 })
    // settled faces cache like any other
    metrics.ascent = 1
    expect(fontChainMetricsPct(chain)).toEqual({ ascentPct: 116, descentPct: 28.8 })
  })
})
