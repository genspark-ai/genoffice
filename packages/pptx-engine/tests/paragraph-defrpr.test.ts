/**
 * Paragraph-level <a:pPr><a:defRPr>: PowerPoint resolves run properties as
 * run rPr → paragraph defRPr → lstStyle/placeholder/master chain. Runs that omit
 * sz/b/fill must pick them up from the paragraph defaults (python-pptx paragraph.font,
 * WPS exports), and the rebuild path must write the node back so those runs keep
 * their look after a structural edit.
 */
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

const YAHEI = '<a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/>'

describe('paragraph <a:pPr><a:defRPr> defaults', () => {
  it('runs without sz/b/fill inherit them from the paragraph defRPr', () => {
    // Shape lstStyle says 18pt regular black; the paragraph defaults say 10pt bold white
    const el = parseOne(
      '<a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr sz="1800" b="0">' +
        '<a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle>' +
        '<a:p><a:pPr algn="ctr"><a:defRPr sz="1000" b="1">' +
        '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>' +
        '<a:latin typeface="Microsoft YaHei"/></a:defRPr></a:pPr>' +
        `<a:r><a:rPr>${YAHEI}</a:rPr><a:t>Core</a:t></a:r>` +
        `<a:r><a:rPr sz="1200" b="0"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill>${YAHEI}</a:rPr><a:t>Layer</a:t></a:r>` +
        '</a:p>',
    )
    const [plain, explicit] = el.text!.paragraphs[0]!.runs
    expect(plain).toMatchObject({ fontSize: 10, bold: true, color: '#FFFFFF' })
    // The values are still inherited from the run's point of view: surgical patches and
    // the rebuild must not bake them into rPr (the defRPr keeps providing them)
    expect(plain).toMatchObject({
      fontSizeImplicit: true,
      boldImplicit: true,
      colorInherited: true,
    })
    // Explicit run attributes win over the paragraph defaults
    expect(explicit).toMatchObject({ fontSize: 12, bold: false, color: '#C00000' })
    expect(explicit.fontSizeImplicit).toBeUndefined()
  })

  it('a partial defRPr only overrides the attributes it carries', () => {
    const el = parseOne(
      '<a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr sz="1800" i="1">' +
        '<a:solidFill><a:srgbClr val="336699"/></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle>' +
        '<a:p><a:pPr><a:defRPr b="1"/></a:pPr><a:r><a:rPr/><a:t>x</a:t></a:r></a:p>',
    )
    const run = el.text!.paragraphs[0]!.runs[0]!
    expect(run).toMatchObject({ fontSize: 18, bold: true, italic: true, color: '#336699' })
  })

  it('the empty-paragraph mark follows the paragraph defRPr size', () => {
    const el = parseOne(
      '<a:bodyPr/><a:p><a:pPr><a:defRPr sz="900"/></a:pPr><a:endParaRPr lang="en-US"/></a:p>',
    )
    const mark = el.text!.paragraphs[0]!.runs[0]!
    expect(mark.paraMark).toBe(true)
    expect(mark.fontSize).toBe(9)
  })

  it('paragraphs without a defRPr are unchanged (no model field, nothing written)', () => {
    const el = parseOne('<a:bodyPr/><a:p><a:pPr algn="ctr"/><a:r><a:rPr/><a:t>x</a:t></a:r></a:p>')
    const p = el.text!.paragraphs[0]!
    expect(p.defRPr).toBeUndefined()
    expect(generateParagraphXml(p)).not.toContain('defRPr')
  })

  it('rebuild writes the defRPr back after the bullet/tab children, keeping theme font refs', () => {
    const el = parseOne(
      '<a:bodyPr/><a:p><a:pPr algn="ctr"><a:buChar char="-"/>' +
        '<a:defRPr sz="1000" b="1" cap="all"><a:solidFill><a:srgbClr val="333333"/></a:solidFill>' +
        '<a:latin typeface="+mn-lt"/><a:ea typeface="Microsoft YaHei"/></a:defRPr></a:pPr>' +
        '<a:r><a:rPr/><a:t>x</a:t></a:r></a:p>',
    )
    const p = el.text!.paragraphs[0]!
    expect(p.defRPr).toEqual({
      fontSize: 10,
      bold: true,
      cap: 'all',
      color: '#333333',
      latinFont: '+mn-lt',
      eaFont: 'Microsoft YaHei',
    })
    const xml = generateParagraphXml(p)
    expect(xml).toContain(
      '<a:pPr algn="ctr"><a:buChar char="-"/>' +
        '<a:defRPr sz="1000" b="1" cap="all"><a:solidFill><a:srgbClr val="333333"/></a:solidFill>' +
        '<a:latin typeface="+mn-lt"/><a:ea typeface="Microsoft YaHei"/></a:defRPr></a:pPr>',
    )
    // The run itself stays attribute-free: it keeps inheriting from the rewritten defRPr
    expect(xml).toContain('<a:r><a:rPr/><a:t>x</a:t></a:r>')
  })
})

describe('theme colors survive rebuild (no srgbClr materialization)', () => {
  const theme = { colors: { accent1: '4472C4', dk1: '000000', lt1: 'FFFFFF' } } as any
  const rebuilt = (txBodyInner: string) => {
    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slideWith(spWith(txBodyInner)),
      ctx: { theme },
    })
    const el = slide.elements[0] as TextElement
    return generateParagraphXml(el.text!.paragraphs[0]!)
  }

  it('keeps a paragraph defRPr schemeClr + lumMod instead of baking srgbClr', () => {
    const out = rebuilt(
      '<a:bodyPr/><a:p><a:pPr algn="ctr"><a:defRPr sz="1000" b="1"><a:solidFill>' +
        '<a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:pPr>' +
        '<a:r><a:rPr lang="en"/><a:t>x</a:t></a:r></a:p>',
    )
    expect(out).toContain('<a:defRPr')
    expect(out).toContain('<a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr>')
    expect(out).not.toContain('srgbClr')
  })

  it('still bakes a plain srgb defRPr color', () => {
    const out = rebuilt(
      '<a:bodyPr/><a:p><a:pPr><a:defRPr sz="1000"><a:solidFill>' +
        '<a:srgbClr val="FF0000"/></a:solidFill></a:defRPr></a:pPr>' +
        '<a:r><a:rPr lang="en"/><a:t>x</a:t></a:r></a:p>',
    )
    expect(out).toContain('<a:srgbClr val="FF0000"/>')
    expect(out).not.toContain('schemeClr')
  })

  it('keeps a bullet buClr schemeClr + tint instead of baking srgbClr', () => {
    const out = rebuilt(
      '<a:bodyPr/><a:p><a:pPr><a:buClr><a:schemeClr val="accent1"><a:tint val="60000"/></a:schemeClr></a:buClr>' +
        '<a:buChar char="•"/></a:pPr><a:r><a:rPr lang="en"/><a:t>x</a:t></a:r></a:p>',
    )
    expect(out).toContain(
      '<a:buClr><a:schemeClr val="accent1"><a:tint val="60000"/></a:schemeClr></a:buClr>',
    )
    expect(out).not.toContain('srgbClr')
  })

  it('still bakes a plain srgb buClr color', () => {
    const out = rebuilt(
      '<a:bodyPr/><a:p><a:pPr><a:buClr><a:srgbClr val="C00000"/></a:buClr>' +
        '<a:buChar char="•"/></a:pPr><a:r><a:rPr lang="en"/><a:t>x</a:t></a:r></a:p>',
    )
    expect(out).toContain('<a:buClr><a:srgbClr val="C00000"/></a:buClr>')
    expect(out).not.toContain('schemeClr')
  })
})
