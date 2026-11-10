import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { EmbeddedFont } from '@genoffice/docx-engine'
import { adoptEmbeddedFonts } from '../src/renderer/embedded-fonts'

const setEmbeddedLineMetrics = vi.fn()
vi.mock('../src/renderer/line-metrics', () => ({
  noteEmbeddedFontsChanged: () => {},
  setEmbeddedLineMetrics: (faces: unknown) => setEmbeddedLineMetrics(faces),
}))

class FakeFontFace {
  family: string
  static gate: Record<string, () => void> = {}
  constructor(family: string) {
    this.family = family
  }
  load(): Promise<this> {
    return new Promise((resolve) => {
      const earlier = FakeFontFace.gate[this.family]
      FakeFontFace.gate[this.family] = () => {
        earlier?.()
        resolve(this)
      }
    })
  }
}

const font = (family: string, extra: Partial<EmbeddedFont> = {}) => ({
  family,
  data: new Uint8Array([0, 1, 0, 0]),
  bold: false,
  italic: false,
  ...extra,
})

describe('adoptEmbeddedFonts', () => {
  const faces = new Set<FakeFontFace>()
  beforeEach(() => {
    vi.stubGlobal('FontFace', FakeFontFace)
    Object.defineProperty(document, 'fonts', {
      configurable: true,
      value: {
        add: (f: FakeFontFace) => faces.add(f),
        delete: (f: FakeFontFace) => faces.delete(f),
        dispatchEvent: () => true,
      },
    })
  })
  afterEach(() => {
    vi.unstubAllGlobals()
    faces.clear()
    FakeFontFace.gate = {}
    setEmbeddedLineMetrics.mockClear()
  })

  it("hands the loaded faces' line boxes over, regular cut first", async () => {
    const box = { ascent: 0.984, descent: 0.273, lineGap: 0 }
    const p = adoptEmbeddedFonts([
      font('Serif', { bold: true, lineMetrics: { ascent: 1, descent: 0.3, lineGap: 0 } }),
      font('Serif', { lineMetrics: box }),
      font('Plain'),
    ])
    FakeFontFace.gate['Serif']()
    FakeFontFace.gate['Plain']()
    await p
    expect(setEmbeddedLineMetrics).toHaveBeenLastCalledWith([
      { styled: false, family: 'Serif', ...box },
      { styled: true, family: 'Serif', ascent: 1, descent: 0.3, lineGap: 0 },
    ])
  })

  it('drops a superseded adoption instead of leaking its faces', async () => {
    const first = adoptEmbeddedFonts([font('Doc One')])
    const second = adoptEmbeddedFonts([font('Doc Two')])
    FakeFontFace.gate['Doc Two']()
    await expect(second).resolves.toBe(true)
    FakeFontFace.gate['Doc One']()
    await expect(first).resolves.toBe(false)
    expect([...faces].map((f) => f.family)).toEqual(['Doc Two'])
  })

  it('revokes the previous document faces', async () => {
    const a = adoptEmbeddedFonts([font('A')])
    FakeFontFace.gate['A']()
    await a
    const b = adoptEmbeddedFonts([font('B')])
    FakeFontFace.gate['B']()
    await b
    expect([...faces].map((f) => f.family)).toEqual(['B'])
  })

  it('lets the winner revoke faces a superseded adoption never saw', async () => {
    const a = adoptEmbeddedFonts([font('A')])
    FakeFontFace.gate['A']()
    await a
    const stale = adoptEmbeddedFonts([font('Stale')])
    const winner = adoptEmbeddedFonts([])
    await winner
    expect(faces.size).toBe(0)
    FakeFontFace.gate['Stale']()
    await expect(stale).resolves.toBe(false)
    expect(faces.size).toBe(0)
  })
})
