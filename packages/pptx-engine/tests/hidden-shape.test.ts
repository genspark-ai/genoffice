/**
 * <p:cNvPr hidden="1"> shapes: PowerPoint never paints them (slideshow, PDF export,
 * or the editing canvas). They must not surface as renderable elements, while their
 * bytes stay in the file (silent passthrough) so saves replay them verbatim.
 */
import { describe, it, expect } from 'vitest'
import { parseSlide, parseDecorations } from '../src/parse'

const slideWith = (body: string) =>
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld>' +
  `<p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${body}</p:spTree></p:cSld></p:sld>`

const sp = (id: number, extra = '') =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s${id}"${extra}/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr>` +
  `<p:txBody><a:bodyPr/><a:p><a:r><a:t>t${id}</a:t></a:r></a:p></p:txBody></p:sp>`

const pic = (id: number, extra = '') =>
  `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="p${id}"${extra}/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
  `<p:blipFill><a:blip r:embed="rId2"/><a:stretch/></p:blipFill>` +
  `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></p:spPr></p:pic>`

describe('cNvPr hidden="1"', () => {
  it('hidden sp becomes a silent passthrough, visible sibling still parses', () => {
    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slideWith(sp(2, ' hidden="1"') + sp(3)),
      ctx: {},
    })
    expect(slide.elements).toHaveLength(2)
    expect(slide.elements[0]!.type).toBe('passthrough')
    expect((slide.elements[0] as any).noChip).toBe(true)
    expect(slide.elements[1]!.type).not.toBe('passthrough')
  })

  it('hidden pic is skipped too', () => {
    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slideWith(pic(2, ' hidden="1"')),
      ctx: {},
    })
    expect(slide.elements[0]!.type).toBe('passthrough')
  })

  it('hidden group child is dropped while visible children stay', () => {
    const grp =
      `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
      `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/>` +
      `<a:chOff x="0" y="0"/><a:chExt cx="914400" cy="914400"/></a:xfrm></p:grpSpPr>` +
      sp(11, ' hidden="1"') +
      sp(12) +
      `</p:grpSp>`
    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slideWith(grp),
      ctx: {},
    })
    const g = slide.elements[0] as any
    expect(g.type).toBe('group')
    expect(g.children).toHaveLength(1)
    expect(g.children[0].name).toBe('s12')
  })

  it('hidden master/layout decorations do not render', () => {
    const decorations = parseDecorations(slideWith(sp(2, ' hidden="1"') + sp(3)), {})
    expect(decorations).toHaveLength(1)
    expect((decorations[0] as any).name).toBe('s3')
  })
})
