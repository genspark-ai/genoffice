/**
 * Run-level format fidelity: latin/ea dual fonts kept + underline style not collapsed.
 * Modeled via parseSlide (with latinFont/eaFont/fontImplicit/underlineStyle markers),
 * then asserted at the byte level through patchTextElementXml.
 */
import { describe, it, expect } from 'vitest'
import { parseSlide } from '../src/parse'
import { patchTextElementXml, setElementFont } from '../src/index'
import type { TextElement } from '../src/types'

const slideWith = (sp: string) =>
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a"><p:cSld>' +
  `<p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${sp}</p:spTree></p:cSld></p:sld>`
const spWith = (runs: string) =>
  '<p:sp><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>' +
  `<p:txBody><a:bodyPr/><a:p>${runs}</a:p></p:txBody></p:sp>`

const parseEl = (runs: string) => {
  const slide = parseSlide({
    path: 'ppt/slides/slide1.xml',
    slideXml: slideWith(spWith(runs)),
    ctx: {},
  })
  return { slide, el: slide.elements[0] as TextElement }
}

describe('latin/ea dual fonts preserved', () => {
  const DUAL =
    '<a:r><a:rPr lang="en" sz="1800"><a:latin typeface="Calibri"/><a:ea typeface="微软雅黑"/></a:rPr><a:t>混排 text</a:t></a:r>'

  it('parses latinFont/eaFont verbatim', () => {
    const { el } = parseEl(DUAL)
    const r = el.text!.paragraphs[0]!.runs[0]!
    expect(r.latinFont).toBe('Calibri')
    expect(r.eaFont).toBe('微软雅黑')
    expect(r.fontImplicit).toBeUndefined()
  })

  it('text-only edit: dual font declarations keep their original bytes', () => {
    const { el } = parseEl(DUAL)
    el.text!.paragraphs[0]!.runs[0]!.text = 'edited 编辑'
    const out = patchTextElementXml(el, el.anchor.originalXml)
    expect(out).toContain('<a:latin typeface="Calibri"/>')
    expect(out).toContain('<a:ea typeface="微软雅黑"/>')
    expect(out).toContain('edited 编辑')
  })

  it('theme font reference (+mn-lt) not materialized when the font is unchanged', () => {
    const { el } = parseEl('<a:r><a:rPr><a:latin typeface="+mn-lt"/></a:rPr><a:t>x</a:t></a:r>')
    el.text!.paragraphs[0]!.runs[0]!.text = 'y'
    const out = patchTextElementXml(el, el.anchor.originalXml)
    expect(out).toContain('typeface="+mn-lt"')
  })

  it('run without font declaration (inherited): text-only edit does not inject latin/ea', () => {
    const { el } = parseEl('<a:r><a:rPr sz="1800"/><a:t>x</a:t></a:r>')
    expect(el.text!.paragraphs[0]!.runs[0]!.fontImplicit).toBe(true)
    el.text!.paragraphs[0]!.runs[0]!.text = 'y'
    const out = patchTextElementXml(el, el.anchor.originalXml)
    expect(out).not.toContain('<a:latin')
  })

  it('setElementFont explicit font change: both latin/ea rewritten to the new font', () => {
    const { slide, el } = parseEl(DUAL)
    setElementFont(slide, el.id, { fontFamily: 'Arial' })
    const out = patchTextElementXml(el, el.anchor.originalXml)
    expect(out).toContain('<a:latin typeface="Arial"/>')
    expect(out).toContain('<a:ea typeface="Arial"/>')
    expect(out).not.toContain('微软雅黑')
  })
})

describe('underline style not collapsed', () => {
  it('u="dbl" stays dbl after a text-only edit', () => {
    const { el } = parseEl('<a:r><a:rPr u="dbl"/><a:t>x</a:t></a:r>')
    const r = el.text!.paragraphs[0]!.runs[0]!
    expect(r.underline).toBe(true)
    expect(r.underlineStyle).toBe('dbl')
    r.text = 'y'
    const out = patchTextElementXml(el, el.anchor.originalXml)
    expect(out).toContain('u="dbl"')
    expect(out).not.toContain('u="sng"')
  })

  it('removing underline: existing u becomes none; runs without u get nothing injected', () => {
    const { el } = parseEl(
      '<a:r><a:rPr u="wavy"/><a:t>a</a:t></a:r><a:r><a:rPr/><a:t>b</a:t></a:r>',
    )
    const [r1, r2] = el.text!.paragraphs[0]!.runs
    r1!.underline = false
    delete r1!.underlineStyle
    r2!.text = 'b2'
    const out = patchTextElementXml(el, el.anchor.originalXml)
    expect(out).toContain('u="none"')
    expect(out).not.toContain('u="wavy"')
    // The second run had no u originally; do not inject one
    expect(out.match(/\su="/g)!.length).toBe(1)
  })

  it('newly enabled underline writes sng', () => {
    const { el } = parseEl('<a:r><a:rPr/><a:t>x</a:t></a:r>')
    el.text!.paragraphs[0]!.runs[0]!.underline = true
    const out = patchTextElementXml(el, el.anchor.originalXml)
    expect(out).toContain('u="sng"')
  })
})

describe('complex-script font follows the font change', () => {
  // Arabic, Hebrew and Indic text renders from a:cs, not a:latin, so a font change that
  // leaves a:cs alone looks like it did nothing at all.
  const ARABIC =
    '<a:r><a:rPr lang="ar-SA"><a:latin typeface="Calibri"/><a:cs typeface="Traditional Arabic"/></a:rPr>' +
    '<a:t>مرحبا</a:t></a:r>'

  it('parses csFont verbatim', () => {
    const { el } = parseEl(ARABIC)
    expect(el.text!.paragraphs[0]!.runs[0]!.csFont).toBe('Traditional Arabic')
  })

  it('changing the font rewrites a:cs as well as a:latin', () => {
    const { slide, el } = parseEl(ARABIC)
    expect(setElementFont(slide, el.id, { fontFamily: 'Amiri' })).toBe(true)
    const out = patchTextElementXml(el, el.anchor.originalXml)
    expect(/<a:latin[^>]*typeface="([^"]*)"/.exec(out)?.[1]).toBe('Amiri')
    expect(/<a:cs[^>]*typeface="([^"]*)"/.exec(out)?.[1]).toBe('Amiri')
  })

  /** The three slots a font change writes, in the order CT_TextCharacterProperties requires. */
  const slotOrder = (xml: string) => [...xml.matchAll(/<a:(latin|ea|cs)\b/g)].map((m) => m[1])

  it('injects a:cs for a run that declared no fonts at all', () => {
    // the shape of the repo's own Arabic fixture: a bare run inheriting everything
    const { slide, el } = parseEl('<a:r><a:t>مرحبا</a:t></a:r>')
    expect(setElementFont(slide, el.id, { fontFamily: 'Amiri' })).toBe(true)
    const out = patchTextElementXml(el, el.anchor.originalXml)
    expect(slotOrder(out)).toEqual(['latin', 'ea', 'cs'])
    expect(/<a:cs[^>]*typeface="([^"]*)"/.exec(out)?.[1]).toBe('Amiri')
  })

  it('keeps the schema order when the run declares only a:cs', () => {
    const { slide, el } = parseEl(
      '<a:r><a:rPr lang="ar-SA"><a:cs typeface="Traditional Arabic"/></a:rPr><a:t>مرحبا</a:t></a:r>',
    )
    expect(setElementFont(slide, el.id, { fontFamily: 'Amiri' })).toBe(true)
    const out = patchTextElementXml(el, el.anchor.originalXml)
    expect(slotOrder(out)).toEqual(['latin', 'ea', 'cs'])
  })

  it('leaves a typeface inside an extension payload alone', () => {
    const { slide, el } = parseEl(
      '<a:r><a:rPr lang="ar-SA"><a:latin typeface="Calibri"/>' +
        '<a:extLst><a:ext uri="{FF2B5EF4}"><a:cs typeface="Opaque Extension Font"/></a:ext></a:extLst>' +
        '</a:rPr><a:t>مرحبا</a:t></a:r>',
    )
    expect(setElementFont(slide, el.id, { fontFamily: 'Amiri' })).toBe(true)
    const out = patchTextElementXml(el, el.anchor.originalXml)
    expect(out).toContain('typeface="Opaque Extension Font"')
    const ownProps = out.slice(0, out.indexOf('<a:extLst'))
    expect(/<a:cs[^>]*typeface="([^"]*)"/.exec(ownProps)?.[1]).toBe('Amiri')
  })

  it('leaves a typeface nested in a subtree alone and sets the run own slots', () => {
    // a:ln may carry its own extension payload with its own font reference; only the run's
    // direct children are ours. Note the nested one comes first in document order.
    const { slide, el } = parseEl(
      '<a:r><a:rPr lang="ar-SA">' +
        '<a:ln><a:extLst><a:ext uri="{FF2B5EF4}"><a:latin typeface="Nested Outline Font"/></a:ext>' +
        '</a:extLst></a:ln><a:latin typeface="Calibri"/>' +
        '</a:rPr><a:t>مرحبا</a:t></a:r>',
    )
    expect(setElementFont(slide, el.id, { fontFamily: 'Amiri' })).toBe(true)
    const out = patchTextElementXml(el, el.anchor.originalXml)
    expect(out).toContain('typeface="Nested Outline Font"')
    const own = out.slice(out.indexOf('</a:ln>'))
    expect(slotOrder(own)).toEqual(['latin', 'ea', 'cs'])
    expect(/<a:latin[^>]*typeface="([^"]*)"/.exec(own)?.[1]).toBe('Amiri')
  })

  it('a structural rebuild keeps the font the user picked in a:cs', () => {
    const { slide, el } = parseEl(ARABIC)
    expect(setElementFont(slide, el.id, { fontFamily: 'Amiri' })).toBe(true)
    el.text!.paragraphs[0]!.runs.push({ text: ' added' }) // forces the rebuild path
    const out = patchTextElementXml(el, el.anchor.originalXml)
    expect(/<a:cs[^>]*typeface="([^"]*)"/.exec(out)?.[1]).toBe('Amiri')
  })

  const refont = (runXml: string) => {
    const { slide, el } = parseEl(runXml)
    expect(setElementFont(slide, el.id, { fontFamily: 'Amiri' })).toBe(true)
    return patchTextElementXml(el, el.anchor.originalXml)
  }

  it('rewrites a single-quoted typeface rather than adding a second one', () => {
    const out = refont(
      "<a:r><a:rPr lang='ar-SA'><a:latin pitchFamily='34' typeface='Old'/></a:rPr>" +
        '<a:t>مرحبا</a:t></a:r>',
    )
    expect(out.match(/typeface\s*=/g)).toHaveLength(3) // latin, ea, cs: one each
    expect(out).toContain("pitchFamily='34'")
    expect(out).not.toContain("typeface='Old'")
  })

  it('expands a self-closing rPr to the full slot group, attributes kept', () => {
    // the open-tag pattern's [^"'>] accepts '/', so an unguarded match would swallow a
    // self-closing tag whole and route it to the latin/ea-only fallback
    const out = refont('<a:r><a:rPr lang="ar-SA" b="1"/><a:t>مرحبا</a:t></a:r>')
    // the run attribute writer may add its own attrs (e.g. i="0"); the declared ones survive
    expect(out).toMatch(/<a:rPr[^>]*\blang="ar-SA"[^>]*\bb="1"[^>]*>/)
    expect(slotOrder(out)).toEqual(['latin', 'ea', 'cs'])
    expect(/<a:cs[^>]*typeface="([^"]*)"/.exec(out)?.[1]).toBe('Amiri')
  })

  it('expands a bare self-closing rPr as well', () => {
    const out = refont('<a:r><a:rPr/><a:t>مرحبا</a:t></a:r>')
    expect(slotOrder(out)).toEqual(['latin', 'ea', 'cs'])
    expect(/<a:cs[^>]*typeface="([^"]*)"/.exec(out)?.[1]).toBe('Amiri')
  })

  it('does not add a second slot group to a run using mc:AlternateContent', () => {
    // the branch children only exist once a consumer picks a branch, so injecting a direct
    // group alongside them would leave two a:latin after preprocessing
    const out = refont(
      '<a:r><a:rPr lang="ar-SA"><mc:AlternateContent>' +
        '<mc:Choice Requires="a14"><a:latin typeface="Old"/></mc:Choice>' +
        '<mc:Fallback><a:latin typeface="Old"/></mc:Fallback>' +
        '</mc:AlternateContent></a:rPr><a:t>مرحبا</a:t></a:r>',
    )
    expect(out.match(/<a:latin\b/g)).toHaveLength(2)
  })

  it('changes the real slot when a comment holds element-looking text', () => {
    // note the run attribute writer already reaches inside a comment on main too, so the comment
    // bytes are not asserted here; what matters is that the font lands on the actual slot
    const out = refont(
      '<a:r><!-- marker <a:rPr/> --><a:rPr lang="ar-SA"><a:latin typeface="Old"/></a:rPr>' +
        '<a:t>مرحبا</a:t></a:r>',
    )
    expect(/<a:latin[^>]*typeface="([^"]*)"/.exec(out)?.[1]).toBe('Amiri')
    expect(out.match(/<a:latin\b/g)).toHaveLength(1)
  })

  it('still injects a font into an unscannable run that declared none', () => {
    // the unscannable path has to keep doing everything the engine did before, injection included
    const out = refont('<a:r><a:rPr/><!-- guard --><a:t>مرحبا</a:t></a:r>')
    expect(/<a:latin[^>]*typeface="([^"]*)"/.exec(out)?.[1]).toBe('Amiri')
    expect(/<a:ea[^>]*typeface="([^"]*)"/.exec(out)?.[1]).toBe('Amiri')
  })

  it('changes only the attribute actually named typeface', () => {
    const out = refont(
      '<a:r><a:rPr lang="ar-SA"><a:latin z:data="typeface=\'Trap\'" typeface="Own"/></a:rPr>' +
        '<a:t>مرحبا</a:t></a:r>',
    )
    expect(out).toContain('z:data="typeface=\'Trap\'"')
    expect(out).toContain('<a:latin z:data="typeface=\'Trap\'" typeface="Amiri"/>')
  })

  it('fills in a slot that carries no attributes at all', () => {
    // an attribute pattern permissive enough for unquoted values also swallows the slash of a
    // self-closing tag, which then reappears in front of whatever is added
    const out = refont(
      '<a:r><a:rPr lang="ar-SA"><a:latin typeface="Old"/><a:cs/></a:rPr>' +
        '<a:t>\u0645\u0631\u062d\u0628\u0627</a:t></a:r>',
    )
    expect(out).not.toContain('/ typeface')
    expect(out).toContain('<a:cs typeface="Amiri"/>')
    expect(slotOrder(out)).toEqual(['latin', 'ea', 'cs'])
  })

  it('does not repeat the attribute when the slot already carries the target font', () => {
    // setting a slot to the typeface it already has leaves the element untouched, which is not
    // the same as the element having no typeface
    const out = refont(
      '<a:r><a:rPr lang="ar-SA"><a:latin typeface="Amiri"/><a:cs typeface="Traditional Arabic"/>' +
        '</a:rPr><a:t>\u0645\u0631\u062d\u0628\u0627</a:t></a:r>',
    )
    expect(out).not.toMatch(/typeface="[^"]*"\s+typeface=/)
    expect(/<a:cs[^>]*typeface="([^"]*)"/.exec(out)?.[1]).toBe('Amiri')
  })

  it('still applies the font when the properties cannot be located', () => {
    // a payload carrying its own a:r closes the slice early; falling back beats doing nothing
    const out = refont(
      '<a:r><a:rPr lang="en"><a:latin typeface="Old"/><a:extLst><a:ext uri="{x}">' +
        '<x:payload xmlns:x="urn:x" xmlns:a="urn:payload"><a:r><a:t>payload</a:t></a:r></x:payload>' +
        '</a:ext></a:extLst></a:rPr><a:t>visible</a:t></a:r>',
    )
    expect(out).toContain('typeface="Amiri"')
    expect(out).not.toContain('typeface="Old"')
  })

  it('uses the run own properties, not an a:rPr that appears earlier in the text', () => {
    const out = refont(
      '<a:r><a:ext uri="{FF2B5EF4}"><a:rPr/></a:ext>' +
        '<a:rPr lang="ar-SA"><a:latin typeface="OldLatin"/><a:cs typeface="OldCS"/></a:rPr>' +
        '<a:t>\u0645\u0631\u062d\u0628\u0627</a:t></a:r>',
    )
    // no font slot goes into the payload; the run attribute writer touches its rPr on main too
    const payload = out.slice(out.indexOf('<a:ext uri='), out.indexOf('</a:ext>'))
    expect(payload).not.toContain('<a:latin')
    expect(payload).not.toContain('<a:cs')
    const own = out.slice(out.indexOf('</a:ext>'))
    expect(/<a:latin[^>]*typeface="([^"]*)"/.exec(own)?.[1]).toBe('Amiri')
    expect(/<a:cs[^>]*typeface="([^"]*)"/.exec(own)?.[1]).toBe('Amiri')
  })

  it('handles a closing tag carrying whitespace before its bracket', () => {
    const out = refont(
      '<a:r><a:rPr lang="ar-SA"><a:latin typeface="Old"/></a:rPr ><a:t>\u0645\u0631\u062d\u0628\u0627</a:t></a:r>',
    )
    // the closing tag survives intact rather than being spliced through
    expect(out).toContain('</a:rPr >')
    expect(out).not.toContain('<<')
    expect(slotOrder(out)).toEqual(['latin', 'ea', 'cs'])
    expect(/<a:latin[^>]*typeface="([^"]*)"/.exec(out)?.[1]).toBe('Amiri')
  })

  it('is not cut short by an a:rPr nested in an extension payload', () => {
    // the run's own closing tag has to be found by depth: the payload's closing tag comes first
    // in the text, and stopping there would leave the real slots untouched
    const out = refont(
      '<a:r><a:rPr lang="ar-SA"><a:latin typeface="Old"/>' +
        '<a:extLst><a:ext uri="{FF2B5EF4}">' +
        '<a:rPr><a:latin typeface="Payload"/></a:rPr>' +
        '</a:ext></a:extLst></a:rPr><a:t>\u0645\u0631\u062d\u0628\u0627</a:t></a:r>',
    )
    expect(out).toContain('typeface="Payload"')
    const own = out.slice(0, out.indexOf('<a:extLst'))
    expect(slotOrder(own)).toEqual(['latin', 'ea', 'cs'])
    expect(/<a:latin[^>]*typeface="([^"]*)"/.exec(own)?.[1]).toBe('Amiri')
  })

  it('reads a name containing a middle dot whole', () => {
    // U+00B7 continues an XML name, so a scan that stopped short would read a:latin\u00B7ext as
    // a:latin and write the typeface into the middle of the tag
    const out = refont(
      '<a:r><a:rPr lang="ar-SA"><a:latin\u00B7ext><z:payload keep="yes"/></a:latin\u00B7ext>' +
        '</a:rPr><a:t>\u0645\u0631\u062d\u0628\u0627</a:t></a:r>',
    )
    expect(out).toContain('<z:payload keep="yes"/>')
    expect(out).not.toMatch(/<a:latin typeface="Amiri"\u00B7/)
  })

  it('does not drop a repeated slot and its attributes', () => {
    const out = refont(
      '<a:r><a:rPr lang="ar-SA"><a:latin typeface="L1" charset="00"/>' +
        '<a:latin typeface="L2" pitchFamily="34"/></a:rPr><a:t>\u0645\u0631\u062d\u0628\u0627</a:t></a:r>',
    )
    expect(out).toContain('charset="00"')
    expect(out).toContain('pitchFamily="34"')
    expect(out.match(/<a:latin\b/g)).toHaveLength(2)
  })

  it('sees a wrapper whose prefix starts with an underscore', () => {
    // XML names may begin with _, so a scan anchored on ASCII letters would miss the wrapper
    // entirely and then empty its branch
    const out = refont(
      '<a:r><a:rPr lang="ar-SA"><_mc:AlternateContent>' +
        '<_mc:Fallback><a:latin typeface="Fallback"/></_mc:Fallback>' +
        '</_mc:AlternateContent></a:rPr><a:t>مرحبا</a:t></a:r>',
    )
    // the branch keeps its content and the font lands inside it, byte for byte as before
    expect(out).toContain('<_mc:Fallback><a:latin typeface="Amiri"/>')
    expect(out).toContain('</_mc:Fallback>')
  })

  it('leaves a payload nested inside a paired slot alone', () => {
    const out = refont(
      '<a:r><a:rPr lang="ar-SA"><a:latin typeface="Own">' +
        '<z:payload typeface="Nested"/></a:latin></a:rPr><a:t>مرحبا</a:t></a:r>',
    )
    expect(out).toContain('<z:payload typeface="Nested"/>')
    expect(out).toContain('<a:latin typeface="Amiri">')
  })

  it('treats a markup-compatibility wrapper as unscannable whatever prefix it uses', () => {
    const out = refont(
      '<a:r><a:rPr lang="ar-SA"><x:AlternateContent>' +
        '<x:Choice Requires="a14"><a:latin typeface="Old"/></x:Choice>' +
        '</x:AlternateContent></a:rPr><a:t>مرحبا</a:t></a:r>',
    )
    // no second group appended alongside the branch-supplied one
    expect(out.match(/<a:latin\b/g)).toHaveLength(1)
    expect(out).not.toContain('<a:cs')
  })
})
