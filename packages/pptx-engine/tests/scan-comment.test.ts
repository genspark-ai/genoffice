/**
 * A slide may legally carry an XML comment inside <p:spTree> (PowerPoint and other
 * producers do it to annotate shapes). The byte-level scanners match tags with a
 * class that cannot cross a '<', so a comment whose body contains markup
 * (<!-- <p:sp> -->) used to be scanned as real tags: the comment's <p:sp> counted
 * as an unclosed open tag, the spTree close was eaten as its close, and the next
 * </p:cSld> threw 'unexpected closing tag'. These tests pin that comments are
 * skipped whole, in scanSlide and in the group-child slicer.
 */
import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { openPptx, savePptx, addElement, createBlankPptx, scanSlide } from '../src/index'
import { sliceGroupChildXmls } from '../src/parse'

const SP = (id: number) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>`

const COMMENT = '<!-- <p:sp> annotated, not a real shape --> -->'

const SLIDE_XML =
  '<p:sld xmlns:p="p"><p:cSld><p:spTree>' +
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr/>' +
  SP(2) +
  COMMENT +
  SP(3) +
  '</p:spTree></p:cSld></p:sld>'

describe('XML comments inside spTree', () => {
  it('scanSlide skips a comment body containing markup and keeps both shapes', () => {
    const scan = scanSlide(SLIDE_XML)
    expect(scan.elements.map((e) => e.name)).toEqual(['p:sp', 'p:sp'])
    expect(scan.elements.map((e) => SLIDE_XML.slice(e.start, e.end))).toEqual([SP(2), SP(3)])
  })

  it('the comment survives a rebuild: the bytes between the shapes are replayed as a gap', () => {
    const scan = scanSlide(SLIDE_XML)
    expect(scan.elements[0]!.gapAfter).toBe(COMMENT)
    expect(scan.bodyPrefix).toContain('<p:spTree>')
    expect(scan.bodySuffix).toBe('</p:spTree></p:cSld></p:sld>')
  })

  it('a slide whose spTree holds a comment still opens', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    addElement(slide, { kind: 'rect', offset: { x: 0, y: 0, cx: 100, cy: 100 } })
    addElement(slide, { kind: 'rect', offset: { x: 200, y: 200, cx: 100, cy: 100 } })
    const zip = await JSZip.loadAsync(await savePptx(opened))
    const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
    const at = xml.indexOf('</p:sp>') + '</p:sp>'.length
    zip.file('ppt/slides/slide1.xml', xml.slice(0, at) + COMMENT + xml.slice(at))

    const reopened = await openPptx(await zip.generateAsync({ type: 'nodebuffer' }))
    expect(reopened.deck.slides[0]!.elements).toHaveLength(2)
  })

  it('sliceGroupChildXmls slices group children past a comment', () => {
    const grpXml =
      '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
      '<p:grpSpPr/>' +
      SP(2) +
      COMMENT +
      SP(3) +
      '</p:grpSp>'
    expect(sliceGroupChildXmls(grpXml)).toEqual([SP(2), SP(3)])
  })
})
