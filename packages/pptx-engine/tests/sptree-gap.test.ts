/** Non-shape spTree children between two shapes (mc:AlternateContent etc.) are preserved verbatim on rebuild. */
import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { openPptx, savePptx, addElement, createBlankPptx, patchSlideXml } from '../src/index'

const ALT =
  '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
  '<mc:Choice Requires="p14"><p:contentPart r:id="rId99"/></mc:Choice>' +
  '<mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="90" name="alt"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp></mc:Fallback>' +
  '</mc:AlternateContent>'

/** Minimal pptx with mc:AlternateContent injected between two rectangles */
async function deckWithGap(): Promise<Buffer> {
  const opened = await openPptx(await createBlankPptx())
  const slide = opened.deck.slides[0]!
  addElement(slide, {
    kind: 'rect',
    offset: { x: 0, y: 0, cx: 100, cy: 100 },
    fillColor: '#111111',
  })
  addElement(slide, {
    kind: 'rect',
    offset: { x: 200, y: 200, cx: 100, cy: 100 },
    fillColor: '#222222',
  })
  const zip = await JSZip.loadAsync(await savePptx(opened))
  const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
  const at = xml.indexOf('</p:sp>') + '</p:sp>'.length
  zip.file('ppt/slides/slide1.xml', xml.slice(0, at) + ALT + xml.slice(at))
  return zip.generateAsync({ type: 'nodebuffer' })
}

describe('non-shape spTree children preserved', () => {
  it('scan records the gap; rebuild without edits == original bytes', async () => {
    const opened = await openPptx(await deckWithGap())
    const slide = opened.deck.slides[0]!
    expect(slide.elements.some((e) => e.anchor.gapAfter?.includes('mc:AlternateContent'))).toBe(
      true,
    )
    expect(patchSlideXml(slide)).toBe(slide.originalXml)
  })

  it('after editing a shape and saving, the non-shape fragment between the two shapes keeps its bytes', async () => {
    const opened = await openPptx(await deckWithGap())
    const slide = opened.deck.slides[0]!
    const last = slide.elements[slide.elements.length - 1]!
    last.transform.offset.x = 999999
    last.dirtyTransform = true
    const zip = await JSZip.loadAsync(await savePptx(opened))
    const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
    expect(xml).toContain(ALT)
    expect(xml).toContain('x="999999"')
    // The fragment still sits before the edited shape (original position)
    expect(xml.indexOf(ALT)).toBeLessThan(xml.indexOf('x="999999"'))
  })
})
