import { describe, expect, it } from 'vitest'

import { probePageGeometry } from '../src/renderer/page-probe'

/** A document whose pages report A4 geometry, with optional per-page rejections */
function fakeDoc(numPages: number, failing: Record<number, Error> = {}) {
  return {
    numPages,
    getPage: async (n: number) => {
      if (failing[n]) throw failing[n]
      return {
        getViewport: ({ scale }: { scale: number }) => ({
          width: 595 * scale,
          height: 842 * scale,
        }),
        rotate: 90,
        view: [0, 0],
        userUnit: 1,
      }
    },
  }
}

describe('probePageGeometry', () => {
  it('probes every page in order on a healthy document', async () => {
    const r = await probePageGeometry(fakeDoc(3))
    expect(r.pages).toHaveLength(3)
    expect(r.failures).toEqual([])
    expect(r.pages[0]).toEqual({ width: 595, height: 842, rotate: 90, origin: [0, 0], userUnit: 1 })
  })

  it('a broken page degrades to the previous geometry instead of failing the document', async () => {
    const r = await probePageGeometry(fakeDoc(450, { 217: new Error('bad xref stream') }))
    expect(r.pages).toHaveLength(450)
    // the broken page inherits its neighbor's geometry so the row layout keeps
    // its shape; the failure is reported for the banner
    expect(r.pages[216]).toEqual(r.pages[215])
    expect(r.failures).toEqual([{ page: 217, message: 'bad xref stream' }])
    // the pages after the broken one still probe normally
    expect(r.pages[217]).toEqual({
      width: 595,
      height: 842,
      rotate: 90,
      origin: [0, 0],
      userUnit: 1,
    })
  })

  it('a broken first page falls back to A4', async () => {
    const r = await probePageGeometry(fakeDoc(2, { 1: new Error('nope') }))
    expect(r.pages[0]).toEqual({ width: 595, height: 842, rotate: 0, origin: [0, 0], userUnit: 1 })
    expect(r.pages[1]).toEqual({ width: 595, height: 842, rotate: 90, origin: [0, 0], userUnit: 1 })
    expect(r.failures).toEqual([{ page: 1, message: 'nope' }])
  })

  it('every page failing still returns a full-length geometry list', async () => {
    const r = await probePageGeometry(
      fakeDoc(3, { 1: new Error('a'), 2: new Error('b'), 3: new Error('c') }),
    )
    expect(r.pages).toHaveLength(3)
    expect(r.failures).toHaveLength(3)
  })
})
