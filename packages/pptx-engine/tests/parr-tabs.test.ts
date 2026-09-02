/** pPr marR / tabLst / defTabSz: parse + rebuild round-trip. */
import { describe, it, expect } from 'vitest'
import { parseSlide } from '../src/parse'
import { generateParagraphXml } from '../src/generate'
import type { TextElement } from '../src/types'

const slideWith = (sp: string) =>
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a"><p:cSld>' +
  `<p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${sp}</p:spTree></p:cSld></p:sld>`
const spWith = (txBodyInner: string) =>
  '<p:sp><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>' +
  `<p:txBody>${txBodyInner}</p:txBody></p:sp>`

const parseOne = (txBodyInner: string) => {
  const slide = parseSlide({
    path: 'ppt/slides/slide1.xml',
    slideXml: slideWith(spWith(txBodyInner)),
    ctx: {},
  })
  return slide.elements[0] as TextElement
}

describe('pPr marR / tabLst / defTabSz', () => {
  it('parses marR, defTabSz and sorted tab stops', () => {
    const el = parseOne(
      '<a:bodyPr/><a:p><a:pPr marR="4128770" defTabSz="457200">' +
        '<a:tabLst><a:tab pos="3384550" algn="l"/><a:tab pos="914400"/></a:tabLst>' +
        '</a:pPr><a:r><a:t>x</a:t></a:r></a:p>',
    )
    const p = el.text!.paragraphs[0]!
    expect(p.marR).toBe(4128770)
    expect(p.defTabSz).toBe(457200)
    expect(p.tabStops).toEqual([{ pos: 914400 }, { pos: 3384550, algn: 'l' }])
    expect(p.pPrExplicit).toMatchObject({ marR: true, tabLst: true, defTabSz: true })
  })

  it('rebuild keeps marR/defTabSz attributes and the tabLst child', () => {
    const el = parseOne(
      '<a:bodyPr/><a:p><a:pPr marL="12700" marR="5080" defTabSz="457200">' +
        '<a:tabLst><a:tab pos="3384550" algn="l"/></a:tabLst>' +
        '</a:pPr><a:r><a:t>x</a:t></a:r></a:p>',
    )
    const xml = generateParagraphXml(el.text!.paragraphs[0]!)
    expect(xml).toContain('marR="5080"')
    expect(xml).toContain('defTabSz="457200"')
    expect(xml).toContain('<a:tabLst><a:tab pos="3384550" algn="l"/></a:tabLst>')
  })

  it('paragraphs without them write none (no baked-in defaults)', () => {
    const el = parseOne('<a:bodyPr/><a:p><a:pPr marL="12700"/><a:r><a:t>x</a:t></a:r></a:p>')
    const xml = generateParagraphXml(el.text!.paragraphs[0]!)
    expect(xml).not.toContain('marR=')
    expect(xml).not.toContain('defTabSz=')
    expect(xml).not.toContain('tabLst')
  })
})
