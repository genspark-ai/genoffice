import { describe, it, expect } from 'vitest'
import { parseSlide } from '../src/parse'

const slideWith = (bodyShapes: string) =>
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld>' +
  `<p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${bodyShapes}</p:spTree></p:cSld></p:sld>`

const sp = (rPr: string, text: string) =>
  '<p:sp><p:nvSpPr><p:cNvPr id="5" name="TextBox 4"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>' +
  '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>' +
  `<p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US">${rPr}</a:rPr><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`

const FONTS =
  '<a:latin typeface="Sakkal Majalla"/><a:cs typeface="Sakkal Majalla"/><a:sym typeface="Wingdings 2"/>'

describe('a:sym font slot (prod_068 Wingdings status dots)', () => {
  it('a PUA-only run draws with the sym typeface', () => {
    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slideWith(sp(FONTS, '')),
      ctx: {},
    })
    const run = (slide.elements[0] as any).text.paragraphs[0].runs[0]
    expect(run.fontFamily).toBe('Wingdings 2')
  })

  it('a plain-text run keeps the latin typeface even with a:sym present', () => {
    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slideWith(sp(FONTS, 'hello')),
      ctx: {},
    })
    const run = (slide.elements[0] as any).text.paragraphs[0].runs[0]
    expect(run.fontFamily).toBe('Sakkal Majalla')
  })
})
