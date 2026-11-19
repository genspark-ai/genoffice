import { describe, expect, it } from 'vitest'

import { inlinePictureDrawingXml, resizeInlinePictureXml } from '../src/index'

/**
 * A run picture resized in the editor (header / table cell) saves its new size:
 * wp:extent and the picture's a:ext follow, nothing else in the drawing moves.
 */
describe('resizeInlinePictureXml', () => {
  const xml = inlinePictureDrawingXml({ rId: 'rId1', docPrId: 1, widthPx: 158.74, heightPx: 39.44 })

  it('rewrites both extents to the new size', () => {
    const out = resizeInlinePictureXml(xml, 240, 60)
    expect(out).toContain('<wp:extent cx="2286000" cy="571500"/>')
    expect(out).toContain('<a:ext cx="2286000" cy="571500"/>')
    expect(out.replace(/c[xy]="\d+"/g, '')).toBe(xml.replace(/c[xy]="\d+"/g, ''))
  })

  it('leaves the drawing byte-identical when the size is the parsed one', () => {
    expect(resizeInlinePictureXml(xml, 158.74, 39.44)).toBe(xml)
  })

  it('follows the picture transform when pic:spPr carries attributes', () => {
    const word = xml.replace('<pic:spPr>', '<pic:spPr bwMode="auto">')
    const out = resizeInlinePictureXml(word, 240, 60)
    expect(out).toContain(
      '<pic:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="2286000" cy="571500"/>',
    )
  })

  it('does not touch floating pictures or invalid sizes', () => {
    const anchored = xml
      .replace('<wp:inline ', '<wp:anchor ')
      .replace('</wp:inline>', '</wp:anchor>')
    expect(resizeInlinePictureXml(anchored, 240, 60)).toBe(anchored)
    expect(resizeInlinePictureXml(xml, 0, 60)).toBe(xml)
  })
})
