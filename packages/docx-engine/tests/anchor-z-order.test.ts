import { describe, expect, it } from 'vitest'
import { buildAnchoredTextboxParagraphXml } from '../src/generate'
import { parseDocx } from '../src/parse'
import { buildDocx } from './helpers/build-docx'

const boxPara = (zOrder: number, text: string, behindDoc = false): string =>
  buildAnchoredTextboxParagraphXml({
    anchor: 'paragraph',
    xEmu: 91440,
    yEmu: 12700,
    widthEmu: 914400,
    heightEmu: 914400,
    fillHex: '1A1A2E',
    zOrder,
    behindDoc,
    id: 10 + zOrder,
    paragraphs: [{ runs: [{ text }] }],
  })

const parseBoxes = async (bodyXml: string) => {
  const parsed = await parseDocx(await buildDocx({ bodyXml }))
  return parsed.blocks.flatMap((b) => b.textboxes ?? [])
}

describe('floating shape z-order (wp:anchor relativeHeight)', () => {
  it('keeps Word-style compact ranks raw', async () => {
    const boxes = await parseBoxes(boxPara(3, 'front') + boxPara(1, 'back'))
    expect(boxes.map((b) => b.z)).toEqual([3, 1])
  })

  it('rank 0 (the base value) stays unset', async () => {
    const boxes = await parseBoxes(boxPara(0, 'base'))
    expect(boxes[0].z).toBeUndefined()
  })

  it('re-ranks wild producer values by document z, stable by order', async () => {
    // LibreOffice-style raw relativeHeight (1, 3, 2) decodes to huge negatives
    const xml = (boxPara(0, 'a') + boxPara(0, 'b') + boxPara(0, 'c'))
      .replace('relativeHeight="251658240"', 'relativeHeight="3"')
      .replace('relativeHeight="251658240"', 'relativeHeight="1"')
      .replace('relativeHeight="251658240"', 'relativeHeight="2"')
    const boxes = await parseBoxes(xml)
    const [a, b, c] = boxes.map((x) => x.z!)
    // compact ranks (the raw deltas are ~-251658237), ordered by decoded value
    expect(boxes.every((x) => Math.abs(x.z!) < 10000)).toBe(true)
    expect(b).toBeLessThan(c)
    expect(c).toBeLessThan(a)
  })

  it('behind flag and z coexist', async () => {
    const boxes = await parseBoxes(boxPara(5, 'bg', true))
    expect(boxes[0].behind).toBe(true)
    expect(boxes[0].z).toBe(5)
  })
})
