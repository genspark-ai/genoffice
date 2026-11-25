import { describe, expect, it } from 'vitest'
import {
  applyShapeWrapAt,
  buildLineParagraphXml,
  buildShapeParagraphXml,
  patchShapeStyles,
  patchTextboxSizes,
} from '../src/generate'
import { parseDocx } from '../src/parse'
import { buildDocx } from './helpers/build-docx'

const SHAPE = buildShapeParagraphXml({
  prst: 'heart',
  fillHex: '4472C4',
  borderHex: '2F5496',
  withTextbox: true,
})

const firstBox = async (paragraphXml: string) => {
  const doc = await parseDocx(await buildDocx({ bodyXml: paragraphXml }))
  return doc.blocks.find((b) => b.textboxes?.length)?.textboxes?.[0]
}

describe('patchShapeStyles rotation', () => {
  it('writes a:xfrm@rot in 60000ths of a degree', () => {
    const out = patchShapeStyles(SHAPE, [{ rotDeg: 90 }])
    expect(out).toContain('<a:xfrm rot="5400000">')
    // the generated xfrm carried no rotation before the patch
    expect(SHAPE).not.toContain('rot="')
  })

  it('keeps fractional degrees and normalizes into one turn', () => {
    expect(patchShapeStyles(SHAPE, [{ rotDeg: 45.5 }])).toContain('rot="2730000"')
    expect(patchShapeStyles(SHAPE, [{ rotDeg: 270 }])).toContain('rot="16200000"')
    expect(patchShapeStyles(SHAPE, [{ rotDeg: -90 }])).toContain('rot="16200000"')
    expect(patchShapeStyles(SHAPE, [{ rotDeg: 765 }])).toContain('rot="2700000"')
  })

  it('replaces an existing rotation instead of appending a second one', () => {
    const once = patchShapeStyles(SHAPE, [{ rotDeg: 90 }])
    const twice = patchShapeStyles(once, [{ rotDeg: 180 }])
    expect(twice).toContain('rot="10800000"')
    expect(twice).not.toContain('5400000')
    expect(twice.match(/rot="/g)).toHaveLength(1)
  })

  it('clears the attribute with null and with 0', () => {
    const rotated = patchShapeStyles(SHAPE, [{ rotDeg: 90 }])
    for (const clear of [null, 0, 360]) {
      const out = patchShapeStyles(rotated, [{ rotDeg: clear }])
      expect(out).not.toContain('rot="')
      expect(out).toContain('<a:xfrm>')
      // everything else survives the clear
      expect(out).toContain('4472C4')
      expect(out).toContain('2F5496')
    }
  })

  it('leaves fill and outline untouched when only the rotation is patched', () => {
    const out = patchShapeStyles(SHAPE, [{ rotDeg: 90 }])
    expect(out).toContain('<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>')
    expect(out).toContain('<a:ln><a:solidFill><a:srgbClr val="2F5496"/></a:solidFill></a:ln>')
  })
})

describe('patchShapeStyles outline width', () => {
  it('writes a:ln@w in EMU from CSS pixels', () => {
    const out = patchShapeStyles(SHAPE, [{ borderWidthPx: 2 }])
    expect(out).toContain('<a:ln w="19050">')
    // the colour slot the generator wrote is still there
    expect(out).toContain('<a:srgbClr val="2F5496"/>')
  })

  it('replaces an existing width', () => {
    const twice = patchShapeStyles(patchShapeStyles(SHAPE, [{ borderWidthPx: 2 }]), [
      { borderWidthPx: 4 },
    ])
    expect(twice).toContain('<a:ln w="38100">')
    expect(twice).not.toContain('19050')
  })

  it('clears the attribute with null', () => {
    const sized = patchShapeStyles(SHAPE, [{ borderWidthPx: 2 }])
    const out = patchShapeStyles(sized, [{ borderWidthPx: null }])
    expect(out).not.toContain('w="19050"')
    expect(out).toContain('<a:ln><a:solidFill><a:srgbClr val="2F5496"/></a:solidFill></a:ln>')
  })

  it('leaves the rotation and colours untouched when only the width is patched', () => {
    const rotated = patchShapeStyles(SHAPE, [{ rotDeg: 90 }])
    const out = patchShapeStyles(rotated, [{ borderWidthPx: 2 }])
    expect(out).toContain('rot="5400000"')
    expect(out).toContain('4472C4')
  })

  it('sets the width on a stroke-only connector without losing its arrow end', () => {
    const line = buildLineParagraphXml({ kind: 'lineArrow' })
    const out = patchShapeStyles(line, [{ borderWidthPx: 3 }])
    expect(out).toContain('<a:ln w="28575">')
    expect(out).toContain('<a:tailEnd type="triangle"/>')
  })
})

describe('patchShapeStyles no-op', () => {
  it('returns the paragraph byte-identically when nothing is requested', () => {
    expect(patchShapeStyles(SHAPE, [{}])).toBe(SHAPE)
    expect(patchShapeStyles(SHAPE, [null, undefined])).toBe(SHAPE)
  })
})

describe('rotation and outline width round-trip through the parser', () => {
  it('restores rotDeg and borderWidthPx on the display box', async () => {
    const out = patchShapeStyles(SHAPE, [{ rotDeg: 90, borderWidthPx: 2 }])
    const box = await firstBox(out)
    expect(box?.rotDeg).toBe(90)
    expect(box?.borderWidthPx).toBe(2)
  })

  it('round-trips a rotated connector', async () => {
    const xml = patchShapeStyles(buildLineParagraphXml({ kind: 'line' }), [{ rotDeg: 270 }])
    const box = await firstBox(xml)
    expect(box?.rotDeg).toBe(270)
  })

  it('round-trips a connector stroke width', async () => {
    const xml = patchShapeStyles(buildLineParagraphXml({ kind: 'lineArrow' }), [
      { borderWidthPx: 3 },
    ])
    const box = await firstBox(xml)
    expect(box?.borderWidthPx).toBe(3)
  })
})

describe('a shape without text is patched through its own drawing', () => {
  const BARE = buildShapeParagraphXml({ prst: 'rect', fillHex: '4472C4' })
  const LABELLED = buildShapeParagraphXml({ prst: 'heart', fillHex: 'FF0000', withTextbox: true })
  const run = (paragraphXml: string) => /<w:r>[\s\S]*<\/w:r>/.exec(paragraphXml)![0]
  const BOTH = `<w:p>${run(BARE)}${run(LABELLED)}</w:p>`
  const boxes = async (paragraphXml: string) => {
    const doc = await parseDocx(await buildDocx({ bodyXml: paragraphXml }))
    return doc.blocks.find((b) => b.textboxes?.length)?.textboxes ?? []
  }

  it('the parse records which drawing each ungrouped shape came from', async () => {
    expect((await boxes(BOTH)).map((b) => b.drawingIndex)).toEqual([0, 1])
  })

  it('is left alone by an ordinal patch, as before', () => {
    expect(patchShapeStyles(BARE, [{ rotDeg: 90 }])).toBe(BARE)
  })

  it('takes a rotation and a size addressed to its drawing', async () => {
    const turned = patchShapeStyles(BARE, [{ rotDeg: 90, drawingIndex: 0 }])
    expect(turned).toContain('<a:xfrm rot="5400000">')
    const sized = patchTextboxSizes(turned, [{ wPx: 200, hPx: 100, drawingIndex: 0 }])
    expect((await boxes(sized))[0]).toMatchObject({ rotDeg: 90, widthPx: 200, heightPx: 100 })
  })

  it('takes a wrap addressed to its drawing', () => {
    const out = applyShapeWrapAt(BARE, { boxIndex: 0, drawingIndex: 0 }, 'behind')
    expect(out).toContain('behindDoc="1"')
    expect(applyShapeWrapAt(BARE, { boxIndex: 0 }, 'behind')).toBe(BARE)
  })

  it('patches each of two shapes in one paragraph, and only the one addressed', async () => {
    const parsed = await boxes(BOTH)
    const out = patchShapeStyles(BOTH, [
      { rotDeg: 90, drawingIndex: parsed[0].drawingIndex },
      { rotDeg: 180, drawingIndex: parsed[1].drawingIndex },
    ])
    expect((await boxes(out)).map((b) => b.rotDeg)).toEqual([90, 180])
    const second = patchShapeStyles(BOTH, [null, { rotDeg: 180, drawingIndex: 1 }])
    expect((await boxes(second)).map((b) => b.rotDeg)).toEqual([undefined, 180])
    const firstDrawing = (xml: string) => xml.slice(0, xml.indexOf('</w:drawing>'))
    expect(firstDrawing(second)).toBe(firstDrawing(BOTH))
  })

  it('a patch without a drawing still goes to the nth drawing that holds text', async () => {
    const out = patchShapeStyles(BOTH, [{ rotDeg: 45 }])
    expect((await boxes(out)).map((b) => b.rotDeg)).toEqual([undefined, 45])
  })
})
