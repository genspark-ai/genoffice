/**
 * positionH page anchors: the renderer measures the offset from the paper
 * edge (one left margin before the column) and never clamps the picture to
 * the column width; a page-relative Y above the body top lifts a side-wrapped
 * run picture there instead of starting it at the anchor line.
 */
import { Editor } from '@tiptap/core'
import { parseDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { DocInlineImage, editorExtensions } from '../src/renderer/editor/extensions'

const PIC =
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="6700000" cy="2540000"/></a:xfrm></pic:spPr></pic:pic></a:graphicData></a:graphic>'

const anchor = (positionH: string, positionV: string, wrap: string) =>
  '<w:r><w:drawing>' +
  '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251658240" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/>' +
  positionH +
  positionV +
  '<wp:extent cx="6700000" cy="2540000"/>' +
  wrap +
  '<wp:docPr id="1" name="Image 1"/>' +
  PIC +
  '</wp:anchor></w:drawing></w:r>'

const PAGE_H =
  '<wp:positionH relativeFrom="page"><wp:posOffset>444500</wp:posOffset></wp:positionH>'
const COLUMN_H =
  '<wp:positionH relativeFrom="column"><wp:posOffset>-311785</wp:posOffset></wp:positionH>'
const PARA_V =
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>232695</wp:posOffset></wp:positionV>'
const PAGE_V = '<wp:positionV relativeFrom="page"><wp:posOffset>23495</wp:posOffset></wp:positionV>'

async function render(bodyXml: string) {
  const parsed = await parseDocx(await buildDocx({ bodyXml, withImage: true }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return editor.view.dom as HTMLElement
}

/** jsdom drops calc()/min() declarations it cannot parse: read the run image's spec instead */
async function runImageStyle(bodyXml: string): Promise<string> {
  const parsed = await parseDocx(await buildDocx({ bodyXml, withImage: true }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  let style = ''
  editor.state.doc.descendants((node) => {
    if (node.type.name !== 'docInlineImage') return true
    const spec = (
      DocInlineImage.config.renderHTML as unknown as (p: {
        node: typeof node
        HTMLAttributes: object
      }) => unknown[]
    ).call(DocInlineImage, { node, HTMLAttributes: {} })
    style = String((spec[1] as Record<string, string>).style ?? '')
    return false
  })
  return style
}

describe('page-relative anchor x', () => {
  it('a block wrapTopAndBottom picture keeps the page x and its full width', async () => {
    const dom = await render(`<w:p>${anchor(PAGE_H, PARA_V, '<wp:wrapTopAndBottom/>')}</w:p>`)
    const wrap = dom.querySelector<HTMLElement>(
      '.doc-protected.img-wrap-topBottom > .doc-img-wrap',
    )!
    expect(wrap.style.marginLeft).toBe('calc(46.7px - var(--doc-margin-left,0px))')
    const img = wrap.querySelector<HTMLElement>('img.doc-protected-img')!
    expect(img.style.maxWidth).toBe('none')
  })

  it('a run-level side-wrapped picture positions from the page edge', async () => {
    const style = await runImageStyle(
      `<w:p>${anchor(PAGE_H, PARA_V, '<wp:wrapSquare wrapText="right"/>')}<w:r><w:t>T</w:t></w:r></w:p>`,
    )
    expect(style).toContain('margin-left:calc(46.7px - var(--doc-margin-left,0px))')
  })

  it('a page-relative Y above the body top lifts a side-wrapped run picture there', async () => {
    const style = await runImageStyle(
      `<w:p>${anchor(COLUMN_H, PAGE_V, '<wp:wrapSquare wrapText="bothSides"/>')}<w:r><w:t>T</w:t></w:r></w:p>`,
    )
    expect(style).toContain('margin-top:min(0px, calc(2.5px - var(--doc-margin-top,0px)))')
  })
})
