/** New slides keep p:sldId/@id inside ST_SlideId (256..2147483647) when the deck already sits at the ceiling. */
import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import {
  openPptx,
  savePptx,
  createBlankPptx,
  duplicateSlide,
  insertSlideWithLayout,
} from '../src/index'
import { nextSlideId } from '../src/slide-ids'

async function deckAtCeiling() {
  const zip = await JSZip.loadAsync(await createBlankPptx())
  const pres = await zip.file('ppt/presentation.xml')!.async('string')
  zip.file('ppt/presentation.xml', pres.replace('<p:sldId id="256"', '<p:sldId id="2147483647"'))
  return openPptx(await zip.generateAsync({ type: 'uint8array' }))
}

describe('slide id allocation', () => {
  it('reads single-quoted ids when allocating the next id', () => {
    const pres =
      '<p:sldIdLst><p:sldId id=\'256\' r:id=\'rId2\'/><p:sldId id="257" r:id="rId3"/></p:sldIdLst>'
    expect(nextSlideId(pres)).toBe(258)
  })

  it.each([
    ['duplicateSlide', (o: Awaited<ReturnType<typeof openPptx>>) => duplicateSlide(o, 0)],
    [
      'insertSlideWithLayout',
      (o: Awaited<ReturnType<typeof openPptx>>) =>
        insertSlideWithLayout(o, 0, o.deck.slides[0]!.layoutPath!),
    ],
  ])(
    '%s falls back to the lowest free id when max+1 would exceed ST_SlideId',
    async (_n, insert) => {
      const opened = await deckAtCeiling()
      expect(insert(opened)).toBeTruthy()
      const saved = await JSZip.loadAsync(await savePptx(opened))
      const ids = [
        ...(await saved.file('ppt/presentation.xml')!.async('string')).matchAll(
          /<p:sldId\s[^>]*\bid="(\d+)"/g,
        ),
      ].map((m) => Number(m[1]))
      expect(ids).toHaveLength(2)
      expect(new Set(ids).size).toBe(2)
      for (const id of ids) expect(id).toBeGreaterThanOrEqual(256)
      for (const id of ids) expect(id).toBeLessThanOrEqual(2147483647)
    },
  )
})
