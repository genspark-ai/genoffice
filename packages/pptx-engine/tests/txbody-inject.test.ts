/**
 * Injecting a <p:txBody> into an element that never had one (a shape gains text for
 * the first time). CT_Shape's sequence is nvSpPr/spPr/style?/txBody?/extLst?, so the
 * insertion point — not just the presence of the element — decides whether PowerPoint
 * can still open the deck.
 */
import { describe, it, expect } from 'vitest'
import { parseSlide } from '../src/parse'
import { rebuildTxBody, isConnectorXml } from '../src/generate'
import { patchSlideXml, patchGroupChildText, findGroupChild } from '../src/index'
import type { GroupElement, TextElement } from '../src/types'

const wrap = (spTree: string) =>
  `<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld><p:spTree>` +
  `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
  spTree +
  `</p:spTree></p:cSld></p:sld>`

const SPPR = `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>`
const STYLE = `<p:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>`
const nvSpPr = (id: number, name: string, inner = '') =>
  `<p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr>${inner}</p:nvPr></p:nvSpPr>`

/** Autoshape without a <p:txBody> (AI/converter output, or text deleted in another tool) */
const SHAPE = `<p:sp>${nvSpPr(2, 'Rect 1')}${SPPR}</p:sp>`
/** Same, but carrying the theme style reference PowerPoint writes for every autoshape */
const SHAPE_STYLED = `<p:sp>${nvSpPr(3, 'Rect 2')}${SPPR}${STYLE}</p:sp>`
/** Placeholder: geometry is inherited, so spPr is routinely self-closing */
const PLACEHOLDER = `<p:sp>${nvSpPr(4, 'PH', '<p:ph type="body" idx="1"/>')}<p:spPr/></p:sp>`
const CONNECTOR = `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="5" name="Straight Connector 1"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>${SPPR}${STYLE}</p:cxnSp>`

/** Type text into the first element the way the setText op does. */
function typeInto(elXml: string) {
  const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml: wrap(elXml), ctx: {} })
  const el = slide.elements[0] as TextElement
  el.text = {
    paragraphs: [{ runs: [{ text: 'HELLO' }] }],
    insets: { l: 91440, t: 45720, r: 91440, b: 45720 },
    anchor: 'middle',
  }
  el.dirty = true
  return { el, out: patchSlideXml(slide) }
}

/** Order of a <p:sp>'s modelled children, as written back. */
const childOrder = (xml: string) =>
  [...xml.matchAll(/<(p:spPr|p:style|p:txBody)\b/g)].map((m) => m[1]).join(' ')

describe('the flags that decide whether a fresh body gets centered', () => {
  const parseOne = (elXml: string) =>
    parseSlide({ path: 'ppt/slides/slide1.xml', slideXml: wrap(elXml), ctx: {} })
      .elements[0] as TextElement

  it('an autoshape is neither a placeholder nor a text box', () => {
    const el = parseOne(SHAPE)
    expect(el.placeholder).toBeUndefined()
    expect(el.txBox).toBeUndefined()
  })

  it('<p:ph> surfaces as the placeholder type', () => {
    expect(parseOne(PLACEHOLDER).placeholder).toBe('body')
  })

  it('<p:cNvSpPr txBox="1"> surfaces as txBox', () => {
    const xml = `<p:sp><p:nvSpPr><p:cNvPr id="6" name="TextBox 1"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>${SPPR}</p:sp>`
    expect(parseOne(xml).txBox).toBe(true)
  })
})

describe('txBody injection into an element that had none', () => {
  it('plain shape: txBody lands after spPr and the text round-trips', () => {
    const { out } = typeInto(SHAPE)
    expect(out).toContain('<p:txBody>')
    expect(childOrder(out)).toBe('p:spPr p:txBody')
    const el = parseSlide({ path: 'p', slideXml: out, ctx: {} }).elements[0] as TextElement
    expect(el.text?.paragraphs[0]?.runs[0]?.text).toBe('HELLO')
  })

  it('styled shape: txBody follows p:style (CT_Shape sequence), not spPr', () => {
    const { out } = typeInto(SHAPE_STYLED)
    expect(childOrder(out)).toBe('p:spPr p:style p:txBody')
    const el = parseSlide({ path: 'p', slideXml: out, ctx: {} }).elements[0] as TextElement
    expect(el.text?.paragraphs[0]?.runs[0]?.text).toBe('HELLO')
  })

  it('self-closing <p:spPr/>: the text is not dropped', () => {
    const { out } = typeInto(PLACEHOLDER)
    expect(out).toContain('<p:spPr/><p:txBody>')
    const el = parseSlide({ path: 'p', slideXml: out, ctx: {} }).elements[0] as TextElement
    expect(el.text?.paragraphs[0]?.runs[0]?.text).toBe('HELLO')
  })

  it('connector: refused — CT_Connector has no txBody child', () => {
    const { out } = typeInto(CONNECTOR)
    expect(out).not.toContain('<p:txBody>')
    expect(isConnectorXml(CONNECTOR)).toBe(true)
    expect(isConnectorXml(SHAPE)).toBe(false)
  })

  it('anchor is written into the synthesized bodyPr and survives a reload', () => {
    const { out } = typeInto(SHAPE)
    expect(out).toContain('<a:bodyPr anchor="ctr"/>')
    const el = parseSlide({ path: 'p', slideXml: out, ctx: {} }).elements[0] as TextElement
    expect(el.text?.anchor).toBe('middle')
  })

  it('insets match the parser defaults, so the live model and the reload agree', () => {
    const { el, out } = typeInto(SHAPE)
    const reloaded = parseSlide({ path: 'p', slideXml: out, ctx: {} }).elements[0] as TextElement
    expect(reloaded.text?.insets).toEqual(el.text?.insets)
  })

  it('an empty body still writes one <a:p> (CT_TextBody requires it)', () => {
    const slide = parseSlide({ path: 'p', slideXml: wrap(SHAPE), ctx: {} })
    const el = slide.elements[0] as TextElement
    el.text = { paragraphs: [] }
    expect(rebuildTxBody(el, el.anchor.originalXml)).toContain('<a:lstStyle/><a:p/></p:txBody>')
  })

  it('group child gains a txBody inside the group bytes', () => {
    const grpSpPr = `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/><a:chOff x="0" y="0"/><a:chExt cx="914400" cy="457200"/></a:xfrm></p:grpSpPr>`
    const group = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="9" name="Group 1"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>${grpSpPr}${SHAPE_STYLED}</p:grpSp>`
    const slide = parseSlide({ path: 'p', slideXml: wrap(group), ctx: {} })
    const grp = slide.elements[0] as GroupElement
    const child = findGroupChild(slide, grp.id, grp.children[0]!.id)!.child as TextElement
    child.text = { paragraphs: [{ runs: [{ text: 'IN GROUP' }] }], anchor: 'middle' }
    expect(patchGroupChildText(slide, grp.id, child)).toBe(true)
    expect(childOrder(grp.anchor.originalXml)).toBe('p:spPr p:style p:txBody')
    const reloaded = parseSlide({ path: 'p', slideXml: patchSlideXml(slide), ctx: {} })
    const reChild = (reloaded.elements[0] as GroupElement).children[0] as TextElement
    expect(reChild.text?.paragraphs[0]?.runs[0]?.text).toBe('IN GROUP')
  })
})
