/** Picture opacity: alphaModFix parsing + byte-surgery write-back. */
import { describe, it, expect } from 'vitest'
import { parseSlide } from '../src/parse'
import { setPictureOpacity } from '../src/index'
import type { PictureElement } from '../src/types'

const slideWith = (inner: string) =>
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld>' +
  `<p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${inner}</p:spTree></p:cSld></p:sld>`
const pic = (blip: string) =>
  '<p:pic><p:nvPicPr><p:cNvPr id="7" name="P"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
  `<p:blipFill>${blip}<a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
  '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr></p:pic>'

const parseOne = (blip: string) =>
  parseSlide({ path: 'ppt/slides/slide1.xml', slideXml: slideWith(pic(blip)), ctx: {} })

describe('picture opacity', () => {
  it('parses alphaModFix → opacity', () => {
    const slide = parseOne('<a:blip r:embed="rId2"><a:alphaModFix amt="40000"/></a:blip>')
    const el = slide.elements[0] as PictureElement
    expect(el.opacity).toBeCloseTo(0.4)
  })

  it('setting opacity: self-closing blip expanded for injection; restoring 1 removes the marker', () => {
    const slide = parseOne('<a:blip r:embed="rId2"/>')
    const el = slide.elements[0] as PictureElement
    expect(setPictureOpacity(slide, el.id, 0.35)).toBe(true)
    expect(el.anchor.originalXml).toContain(
      '<a:blip r:embed="rId2"><a:alphaModFix amt="35000"/></a:blip>',
    )
    expect(el.opacity).toBeCloseTo(0.35)
    expect(setPictureOpacity(slide, el.id, 1)).toBe(true)
    expect(el.anchor.originalXml).not.toContain('alphaModFix')
    expect(el.opacity).toBeUndefined()
  })

  it('existing alphaModFix is replaced, not stacked', () => {
    const slide = parseOne('<a:blip r:embed="rId2"><a:alphaModFix amt="40000"/></a:blip>')
    const el = slide.elements[0] as PictureElement
    setPictureOpacity(slide, el.id, 0.8)
    const hits = el.anchor.originalXml.match(/alphaModFix/g) ?? []
    expect(hits.length).toBe(1)
    expect(el.anchor.originalXml).toContain('amt="80000"')
  })
})

describe('picture softEdge', () => {
  it('parses a:softEdge rad → PictureElement.softEdge', () => {
    const slideXml =
      '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>' +
      '<p:pic><p:nvPicPr><p:cNvPr id="7" name="P"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
      '<p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
      '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>' +
      '<a:effectLst><a:softEdge rad="63500"/></a:effectLst></p:spPr></p:pic>' +
      '</p:spTree></p:cSld></p:sld>'
    const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml, ctx: {} })
    expect((slide.elements[0] as PictureElement).softEdge).toBe(63500)
  })
})
