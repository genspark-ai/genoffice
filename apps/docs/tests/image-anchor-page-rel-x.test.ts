import { Editor } from '@tiptap/core'
import { parseDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

const PIC =
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'

const anchor = (relH: string, posOffsetEmu: number) =>
  '<w:p><w:r><w:t>before</w:t></w:r><w:r><w:drawing>' +
  '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/>' +
  `<wp:positionH relativeFrom="${relH}"><wp:posOffset>${posOffsetEmu}</wp:posOffset></wp:positionH>` +
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="1905000" cy="952500"/><wp:wrapNone/>' +
  '<wp:docPr id="1" name="pic 1"/>' +
  PIC +
  '</wp:anchor></w:drawing></w:r><w:r><w:t>after</w:t></w:r></w:p>'

async function render(bodyXml: string) {
  const parsed = await parseDocx(await buildDocx({ bodyXml, withImage: true }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  const img = editor.view.dom.querySelector<HTMLImageElement>('.doc-inline-img-anchor > img')!
  const out = { transform: img.style.transform, data: { ...img.dataset } }
  editor.destroy()
  return out
}

// a two-column journal page anchored its right-column figure 11.8 cm from the
// paper edge; measured from the column it landed one column further right
describe('run-level wrapNone pictures with a page/margin-relative X', () => {
  it('measures a paper-edge X past the left margin and marks the column undo', async () => {
    const r = await render(anchor('page', 4261041))
    expect(r.transform).toBe(
      'translate(calc(447.4px - var(--doc-margin-left,0px) - var(--page-rel-x-undo,0px)), calc(0.0px - var(--anchor-before,0px)))',
    )
    expect(r.data.pageRelX).toBe('1')
    expect(r.data.pageRelXFrom).toBe('page')
    expect(r.data.pageRelXPx).toBe('447.4')
  })

  it('keeps a margin-relative X at the body left but still undoes the column shift', async () => {
    const r = await render(anchor('margin', 914400))
    expect(r.transform).toBe(
      'translate(calc(96.0px - var(--page-rel-x-undo,0px)), calc(0.0px - var(--anchor-before,0px)))',
    )
    expect(r.data.pageRelX).toBe('1')
    expect(r.data.pageRelXFrom).toBeUndefined()
  })

  it('leaves a column-relative X on the plain pixel translate', async () => {
    const r = await render(anchor('column', 914400))
    expect(r.transform).toBe(
      'translate(calc(96.0px - var(--anchor-ind,0px)), calc(0.0px - var(--anchor-before,0px)))',
    )
    expect(r.data.pageRelX).toBeUndefined()
  })
})
