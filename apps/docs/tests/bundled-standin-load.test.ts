/**
 * Opening a document waits for the bundled stand-in cuts its families chain
 * to: canvas measures the system fallback until a web face is in, and the
 * autofit table pass sizes columns from those advances once.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { loadBundledStandIns } from '../src/renderer/line-metrics'

function stubFonts(loaded: Set<string>, settle: () => Promise<unknown> = () => Promise.resolve()) {
  const load = vi.fn(async (spec: string) => {
    await settle()
    loaded.add(spec)
    return []
  })
  const fonts = { check: (spec: string) => loaded.has(spec), load }
  Object.defineProperty(document, 'fonts', { value: fonts, configurable: true })
  return load
}

function stubCanvas() {
  const fake = { font: '', measureText: () => ({ width: 50 }) }
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(
    fake as unknown as CanvasRenderingContext2D,
  )
}

afterEach(() => {
  vi.restoreAllMocks()
  vi.useRealTimers()
  Reflect.deleteProperty(document, 'fonts')
})

describe('loadBundledStandIns', () => {
  it('loads every cut of the stand-in a missing family resolves to, once per face', async () => {
    const loaded = new Set<string>()
    const load = stubFonts(loaded)
    stubCanvas()
    await loadBundledStandIns(['Calibri', 'Calibri Light', 'Cambria'])
    const specs = load.mock.calls.map(([spec]) => spec)
    expect(specs).toEqual([
      '16px "Carlito GO"',
      'bold 16px "Carlito GO"',
      'italic 16px "Carlito GO"',
      'bold italic 16px "Carlito GO"',
      '16px "Caladea"',
      'bold 16px "Caladea"',
      'italic 16px "Caladea"',
      'bold italic 16px "Caladea"',
    ])
    expect(loaded.has('bold 16px "Carlito GO"')).toBe(true)
  })

  it('skips cuts already in and chains that lead with a generic family', async () => {
    const loaded = new Set(['16px "Carlito GO"', 'bold 16px "Carlito GO"'])
    const load = stubFonts(loaded)
    stubCanvas()
    await loadBundledStandIns(['Calibri', 'sans-serif'])
    expect(load.mock.calls.map(([spec]) => spec)).toEqual([
      'italic 16px "Carlito GO"',
      'bold italic 16px "Carlito GO"',
    ])
  })

  it('gives up on a face that never loads instead of holding the open', async () => {
    vi.useFakeTimers()
    stubFonts(new Set(), () => new Promise(() => {}))
    stubCanvas()
    let done = false
    const p = loadBundledStandIns(['Calibri'], 1000).then(() => {
      done = true
    })
    await vi.advanceTimersByTimeAsync(999)
    expect(done).toBe(false)
    await vi.advanceTimersByTimeAsync(1)
    await p
    expect(done).toBe(true)
  })

  it('is a no-op without a FontFaceSet loader', async () => {
    Reflect.deleteProperty(document, 'fonts')
    await expect(loadBundledStandIns(['Calibri'])).resolves.toBeUndefined()
  })
})
