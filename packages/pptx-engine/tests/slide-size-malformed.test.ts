/** Malformed p:sldSz must fall back to defaults instead of yielding NaN. */
import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { openPptx, createBlankPptx } from '../src/index'

async function withSlideSize(xmlPatch: (xml: string) => string): Promise<Uint8Array> {
  const blank = await createBlankPptx()
  const zip = await JSZip.loadAsync(blank)
  const path = 'ppt/presentation.xml'
  const xml = await zip.file(path)!.async('text')
  zip.file(path, xmlPatch(xml))
  return zip.generateAsync({ type: 'uint8array' })
}

describe('readPresentation malformed sldSz', () => {
  it('falls back to defaults when cx/cy are non-numeric', async () => {
    const bytes = await withSlideSize((xml) =>
      xml.replace(/cx="[^"]*"/, 'cx="not-a-number"').replace(/cy="[^"]*"/, 'cy=""'),
    )
    const opened = await openPptx(bytes)
    expect(Number.isFinite(opened.deck.size.cx)).toBe(true)
    expect(Number.isFinite(opened.deck.size.cy)).toBe(true)
    expect(opened.deck.size).toEqual({ cx: 9144000, cy: 6858000 })
  })

  it('falls back per axis when only one dimension is corrupt', async () => {
    const bytes = await withSlideSize((xml) => xml.replace(/cx="[^"]*"/, 'cx="oops"'))
    const opened = await openPptx(bytes)
    expect(opened.deck.size.cx).toBe(9144000)
    expect(opened.deck.size.cy).toBe(6858000)
  })

  it('rejects non-positive dimensions instead of propagating them', async () => {
    const bytes = await withSlideSize((xml) =>
      xml.replace(/cx="[^"]*"/, 'cx="0"').replace(/cy="[^"]*"/, 'cy="-10"'),
    )
    const opened = await openPptx(bytes)
    expect(opened.deck.size).toEqual({ cx: 9144000, cy: 6858000 })
  })
})
