/**
 * Full-rebuild rPr fidelity: explicit "off" attributes (b="0", u="none",
 * strike="noStrike", spc="0", baseline="0", cap="none", kern) are overrides of
 * inherited styling — a rebuild that drops them re-inherits from the
 * placeholder/master chain. Bodies containing <a:fld> always take the rebuild
 * path, so slide-number placeholders were losing these (edit-eval identity
 * sweep, prod_016 / Lecture_Week_5).
 */
import { describe, it, expect } from 'vitest'
import { parseSlide } from '../src/parse'
import { rebuildTxBody } from '../src/generate'
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

const parseEl = (spTree: string): TextElement => {
  const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml: wrap(spTree), ctx: {} })
  return slide.elements[0] as TextElement
}

const EXPLICIT_OFF_RPR =
  'kumimoji="0" lang="en-US" sz="900" b="0" i="0" u="none" strike="noStrike" ' +
  'kern="1200" cap="none" spc="0" baseline="0"'

describe('rebuildTxBody keeps explicit-off rPr attributes', () => {
  it('a fld run re-emits b/i/u/strike/kern/cap/spc/baseline overrides', () => {
    const el = parseEl(
      textBox(
        `<a:p>` +
          `<a:fld id="{B6D32A31-0000-0000-0000-000000000002}" type="slidenum">` +
          `<a:rPr ${EXPLICIT_OFF_RPR}/><a:t>7</a:t></a:fld>` +
          `</a:p>`,
      ),
    )
    const xml = rebuildTxBody(el, el.anchor.originalXml)
    const rpr = /<a:fld[^>]*>(<a:rPr[^>]*\/?>)/.exec(xml)?.[1] ?? ''
    expect(rpr).toContain('b="0"')
    expect(rpr).toContain('i="0"')
    expect(rpr).toContain('u="none"')
    expect(rpr).toContain('strike="noStrike"')
    expect(rpr).toContain('kern="1200"')
    expect(rpr).toContain('cap="none"')
    expect(rpr).toContain('spc="0"')
    expect(rpr).toContain('baseline="0"')
  })

  it('a plain run in the same rebuilt body keeps its overrides too', () => {
    const el = parseEl(
      textBox(
        `<a:p>` +
          `<a:r><a:rPr ${EXPLICIT_OFF_RPR}/><a:t>page </a:t></a:r>` +
          `<a:fld id="{B6D32A31-0000-0000-0000-000000000003}" type="slidenum">` +
          `<a:rPr lang="en-US"/><a:t>7</a:t></a:fld>` +
          `</a:p>`,
      ),
    )
    const xml = rebuildTxBody(el, el.anchor.originalXml)
    const rpr = /<a:r><a:rPr[^>]*/.exec(xml)?.[0] ?? ''
    expect(rpr).toContain('b="0"')
    expect(rpr).toContain('u="none"')
    expect(rpr).toContain('strike="noStrike"')
    expect(rpr).toContain('cap="none"')
  })

  it('inherited (attr-less) runs still write no override attributes', () => {
    const el = parseEl(
      textBox(
        `<a:p>` +
          `<a:r><a:rPr lang="en-US"/><a:t>plain</a:t></a:r>` +
          `<a:fld id="{B6D32A31-0000-0000-0000-000000000004}" type="slidenum">` +
          `<a:rPr lang="en-US"/><a:t>7</a:t></a:fld>` +
          `</a:p>`,
      ),
    )
    const xml = rebuildTxBody(el, el.anchor.originalXml)
    for (const frag of ['b="', 'i="', 'u="', 'strike="', 'cap="', 'spc="', 'baseline="']) {
      expect(xml).not.toContain(frag)
    }
  })

  it('non-plain color nodes restore verbatim instead of baking the resolved value', () => {
    const el = parseEl(
      textBox(
        `<a:p>` +
          `<a:fld id="{B6D32A31-0000-0000-0000-000000000006}" type="slidenum">` +
          `<a:rPr lang="en-US"><a:solidFill><a:prstClr val="black"><a:tint val="75000"/></a:prstClr></a:solidFill></a:rPr>` +
          `<a:t>7</a:t></a:fld>` +
          `<a:r><a:rPr lang="en-US"><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill></a:rPr>` +
          `<a:t>x</a:t></a:r>` +
          `</a:p>`,
      ),
    )
    const xml = rebuildTxBody(el, el.anchor.originalXml)
    expect(xml).toContain('<a:prstClr val="black"><a:tint val="75000"/></a:prstClr>')
    expect(xml).toContain('<a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr>')
    expect(xml).not.toMatch(/<a:srgbClr/)
  })

  it('positive styles keep winning over the off markers (u/strike style preserved)', () => {
    const el = parseEl(
      textBox(
        `<a:p>` +
          `<a:r><a:rPr lang="en-US" b="1" u="dbl" strike="dblStrike"/><a:t>styled</a:t></a:r>` +
          `<a:fld id="{B6D32A31-0000-0000-0000-000000000005}" type="slidenum">` +
          `<a:rPr lang="en-US"/><a:t>7</a:t></a:fld>` +
          `</a:p>`,
      ),
    )
    const xml = rebuildTxBody(el, el.anchor.originalXml)
    expect(xml).toContain('b="1"')
    expect(xml).toContain('u="dbl"')
    expect(xml).toContain('strike="dblStrike"')
  })
})
