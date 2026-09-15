/**
 * <a:fld> in-paragraph ordering: a field must stay at its document
 * position relative to plain runs — "fld(2) + r(h)" renders as "2h", not "h2".
 */
import { describe, it, expect } from 'vitest'
import { parseSlide } from '../src/parse'
import { generateParagraphXml, patchTextElementXml } from '../src/generate'
import type { TextElement } from '../src/types'

const wrap = (spTree: string) =>
  `<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld><p:spTree>` +
  `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
  spTree +
  `</p:spTree></p:cSld></p:sld>`

const textBox = (paragraphs: string) => `<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="TextBox 1"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm></p:spPr>
  <p:txBody><a:bodyPr/>${paragraphs}</p:txBody>
</p:sp>`

const runsOf = (spTree: string, paraIndex = 0) => {
  const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml: wrap(spTree), ctx: {} })
  const el = slide.elements[0] as TextElement
  return el.text!.paragraphs[paraIndex]!.runs
}

const parseEl = (spTree: string) => {
  const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml: wrap(spTree), ctx: {} })
  return slide.elements[0] as TextElement
}

describe('a:fld document-order preservation', () => {
  it('fld before a plain run stays first (fld("2") + r("h") → "2h")', () => {
    const runs = runsOf(
      textBox(
        `<a:p>` +
          `<a:fld id="{B6D32A31-0000-0000-0000-000000000000}" type="datetime2"><a:rPr lang="en-US"/><a:t>2</a:t></a:fld>` +
          `<a:r><a:rPr lang="en-US"/><a:t>h</a:t></a:r>` +
          `</a:p>`,
      ),
    )
    expect(runs.map((r) => r.text)).toEqual(['2', 'h'])
    expect(runs[0]!.field).toBe('datetime2')
    expect(runs[1]!.field).toBeUndefined()
  })

  it('run + fld + run keeps interleaved order (r("8~") + fld("10") + r("h") → "8~10h")', () => {
    const runs = runsOf(
      textBox(
        `<a:p>` +
          `<a:r><a:rPr lang="en-US"/><a:t>8~</a:t></a:r>` +
          `<a:fld id="{B6D32A31-0000-0000-0000-000000000001}" type="datetime10"><a:rPr lang="en-US"/><a:t>10</a:t></a:fld>` +
          `<a:r><a:rPr lang="en-US"/><a:t>h</a:t></a:r>` +
          `</a:p>`,
      ),
    )
    expect(runs.map((r) => r.text)).toEqual(['8~', '10', 'h'])
    expect(runs.map((r) => r.field)).toEqual([undefined, 'datetime10', undefined])
  })

  it('slidenum fld alone in its paragraph still parses with field set', () => {
    const runs = runsOf(
      textBox(
        `<a:p>` +
          `<a:fld id="{B6D32A31-0000-0000-0000-000000000002}" type="slidenum"><a:rPr lang="en-US"/><a:t>7</a:t></a:fld>` +
          `</a:p>`,
      ),
    )
    expect(runs.map((r) => r.text)).toEqual(['7'])
    expect(runs[0]!.field).toBe('slidenum')
  })

  it('a slidenum fld with no cached text is kept as a field run, not an empty paragraph mark', () => {
    // Layout-authored fields are saved without <a:t>; the value is substituted at render
    // time, so the endParaRPr mark must not replace the run (that dropped the number)
    const runs = runsOf(
      textBox(
        `<a:p>` +
          `<a:fld id="{B6D32A31-0000-0000-0000-000000000003}" type="slidenum"><a:rPr lang="zh-CN" sz="2000" b="1"/></a:fld>` +
          `<a:endParaRPr lang="zh-CN" sz="2000" b="1"/>` +
          `</a:p>`,
      ),
    )
    expect(runs).toHaveLength(1)
    expect(runs[0]!.field).toBe('slidenum')
    expect(runs[0]!.text).toBe('')
    expect(runs[0]!.paraMark).toBeUndefined()
  })

  it('a:tab between runs stays in order as a tab stop (r("Col1") + tab + r("Col2"))', () => {
    // Defensive: ECMA-376 CT_TextParagraph allows pPr/r/br/fld/endParaRPr only —
    // a:tab lives in a:tabLst — and PowerPoint encodes tabs as a literal tab in
    // <a:t>. The bare form below is only seen from third-party producers.
    const runs = runsOf(
      textBox(
        `<a:p>` +
          `<a:r><a:rPr lang="en-US"/><a:t>Col1</a:t></a:r>` +
          `<a:tab/>` +
          `<a:r><a:rPr lang="en-US"/><a:t>Col2</a:t></a:r>` +
          `</a:p>`,
      ),
    )
    expect(runs.map((r) => r.text).join('')).toBe('Col1\tCol2')
  })

  it('literal tab in <a:t> (PowerPoint encoding) stays in order', () => {
    const runs = runsOf(textBox(`<a:p><a:r><a:rPr lang="en-US"/><a:t>Col1\tCol2</a:t></a:r></a:p>`))
    expect(runs.map((r) => r.text).join('')).toBe('Col1\tCol2')
  })

  it('bare <a:tab/> survives an in-place text patch (no rebuild fallback)', () => {
    const el = parseEl(
      textBox(
        `<a:p>` +
          `<a:r><a:rPr lang="en-US"/><a:t>Col1</a:t></a:r>` +
          `<a:tab/>` +
          `<a:r><a:rPr lang="en-US"/><a:t>Col2</a:t></a:r>` +
          `</a:p>`,
      ),
    )
    el.text!.paragraphs[0]!.runs[0]!.text = 'Col1!'
    const out = patchTextElementXml(el, el.anchor.originalXml)
    expect(out).toContain('<a:tab/>')
    expect(out).toContain('Col1!')
    expect(out).toContain('Col2')
  })

  it('literal single-tab run patches in place and tabLst stops are ignored as spans', () => {
    const el = parseEl(
      textBox(
        `<a:p><a:pPr><a:tabLst><a:tab pos="914400"/></a:tabLst></a:pPr>` +
          `<a:r><a:rPr lang="en-US"/><a:t>Col1</a:t></a:r>` +
          `<a:r><a:rPr lang="en-US"/><a:t>\t</a:t></a:r>` +
          `<a:r><a:rPr lang="en-US"/><a:t>Col2</a:t></a:r>` +
          `</a:p>`,
      ),
    )
    expect(el.text!.paragraphs[0]!.tabStops).toEqual([{ pos: 914400 }])
    el.text!.paragraphs[0]!.runs[0]!.text = 'Col1!'
    const out = patchTextElementXml(el, el.anchor.originalXml)
    expect(out).toContain('<a:tabLst><a:tab pos="914400"/></a:tabLst>')
    expect(out).toContain('Col1!')
  })

  it('rebuild normalizes a bare tab to a valid literal (schema: tabLst → defRPr order kept)', () => {
    const el = parseEl(
      textBox(
        `<a:p>` +
          `<a:r><a:rPr lang="en-US"/><a:t>Col1</a:t></a:r>` +
          `<a:tab/>` +
          `<a:r><a:rPr lang="en-US"/><a:t>Col2</a:t></a:r>` +
          `</a:p>`,
      ),
    )
    const xml = generateParagraphXml(el.text!.paragraphs[0]!)
    expect(xml).toContain('\t')
    expect(xml).not.toContain('<a:tab/>')
  })
})
