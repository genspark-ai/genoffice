import { describe, expect, it } from 'vitest'
import { parseSlide } from '../src/parse'
import { parseDefaultTextStyle } from '../src/placeholder'

const slideWith = (bodyShapes: string) =>
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld>' +
  `<p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${bodyShapes}</p:spTree></p:cSld></p:sld>`

const PRES =
  '<?xml version="1.0"?><p:presentation xmlns:p="p" xmlns:a="a">' +
  '<p:defaultTextStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="1200" b="1"><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr>' +
  '</p:defaultTextStyle></p:presentation>'

const JA = String.fromCodePoint(0x65e5, 0x672c, 0x8a9e)

const THEME = { minorFont: 'Calibri', majorFont: 'Calibri Light', colors: {} } as any

const box = (paragraphs: string) =>
  '<p:sp><p:nvSpPr><p:cNvPr id="5" name="TextBox 4"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>' +
  '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>' +
  `<p:txBody><a:bodyPr/>${paragraphs}</p:txBody></p:sp>`

function firstParagraph(sp: string) {
  const ctx = { defaultTextStyle: parseDefaultTextStyle(PRES, THEME), theme: THEME }
  const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml: slideWith(sp), ctx })
  return (slide.elements[0] as any).text.paragraphs[0]
}

describe('text style provenance', () => {
  it('names the inheritance layer for inherited values and the run for explicit ones', () => {
    const p = firstParagraph(
      box('<a:p><a:r><a:rPr lang="en-US" i="1"/><a:t>plain</a:t></a:r></a:p>'),
    )
    expect(p.runs[0].styleSrc).toEqual({
      fontSize: 'presentation defaultTextStyle',
      bold: 'presentation defaultTextStyle',
      italic: 'run',
      color: 'default',
      fontFamily: 'theme minor via presentation defaultTextStyle',
    })
    expect(p.alignSrc).toBe('presentation defaultTextStyle')
    const explicit = firstParagraph(
      box(
        '<a:p><a:pPr algn="r"/><a:r><a:rPr sz="2000"><a:latin typeface="Georgia"/></a:rPr><a:t>x</a:t></a:r></a:p>',
      ),
    )
    expect(explicit.runs[0].styleSrc).toMatchObject({ fontSize: 'run', fontFamily: 'run' })
    expect(explicit.alignSrc).toBe('paragraph')
  })

  it('credits a concrete paragraph defRPr face without the theme ref, and the picked script slot', () => {
    const para = firstParagraph(
      box(
        '<a:p><a:pPr><a:defRPr sz="1800"><a:latin typeface="Georgia"/></a:defRPr></a:pPr><a:r><a:rPr lang="en-US"/><a:t>x</a:t></a:r></a:p>',
      ),
    )
    expect(para.runs[0].fontFamily).toBe('Georgia')
    expect(para.runs[0].styleSrc).toMatchObject({
      fontSize: 'paragraph defRPr',
      fontFamily: 'paragraph defRPr',
    })
    const ea = firstParagraph(
      box(
        `<a:p><a:r><a:rPr lang="ja-JP"><a:ea typeface="Meiryo"/></a:rPr><a:t>${JA}</a:t></a:r></a:p>`,
      ),
    )
    expect(ea.runs[0].fontFamily).toBe('Meiryo')
    expect(ea.runs[0].styleSrc.fontFamily).toBe('run')
    const eaFromParagraph = firstParagraph(
      box(
        `<a:p><a:pPr><a:defRPr><a:ea typeface="Meiryo"/></a:defRPr></a:pPr><a:r><a:rPr lang="ja-JP"/><a:t>${JA}</a:t></a:r></a:p>`,
      ),
    )
    expect(eaFromParagraph.runs[0].fontFamily).toBe('Meiryo')
    expect(eaFromParagraph.runs[0].styleSrc.fontFamily).toBe('paragraph defRPr')
  })
})
