import { Editor } from '@tiptap/core'
import { parseDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { pinParagraphRelVPictures } from '../src/renderer/editor/pagination-gaps'
import type { BlockPlan } from '../src/renderer/pagination-plan'

const PIC =
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'

const anchor = (relV: string, posOffsetEmu: number) =>
  '<w:p><w:pPr><w:spacing w:before="150"/></w:pPr><w:r><w:t>before</w:t></w:r><w:r><w:drawing>' +
  '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/>' +
  '<wp:positionH relativeFrom="column"><wp:posOffset>914400</wp:posOffset></wp:positionH>' +
  `<wp:positionV relativeFrom="${relV}"><wp:posOffset>${posOffsetEmu}</wp:posOffset></wp:positionV>` +
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
  return { editor, relV: parsed.blocks[0].runs?.[1]?.image?.relV }
}

// a portrait page anchored a small wrapNone picture 199pt below a paragraph
// carrying spacing-before 7.5pt: Word measured the offset from above the
// spacing, so the picture sat 7.5pt too low
describe('run-level wrapNone pictures with a paragraph-relative Y', () => {
  it('subtracts the host paragraph spacing-before from a paragraph-relative offset', async () => {
    const { editor, relV } = await render(anchor('paragraph', 2532380))
    expect(relV).toBeUndefined()
    const img = editor.view.dom.querySelector<HTMLImageElement>('.doc-inline-img-anchor > img')!
    expect(img.style.transform).toBe(
      'translate(calc(96.0px - var(--anchor-ind,0px)), calc(265.9px - var(--anchor-before,0px)))',
    )
    const host = img.closest<HTMLElement>('.doc-anchor-origin')!
    expect(host.style.marginTop).toBe('7.5pt')
    host.style.marginTop = '10px'
    pinParagraphRelVPictures(editor.view.dom)
    expect(img.style.getPropertyValue('--anchor-before')).toBe('10.0px')
    editor.destroy()
  })

  it('takes the planned space before over the computed margin, 0 where the page lead dropped it', async () => {
    const { editor } = await render(anchor('paragraph', 0))
    const img = editor.view.dom.querySelector<HTMLImageElement>('.doc-inline-img-anchor > img')!
    const host = img.closest<HTMLElement>('.doc-anchor-origin')!
    host.style.marginTop = '10px'
    const plan = { kind: 'para', spaceBeforePt: 12 } as unknown as BlockPlan
    pinParagraphRelVPictures(editor.view.dom, (el) => (el === host ? plan : undefined))
    expect(img.style.getPropertyValue('--anchor-before')).toBe('16.0px')
    pinParagraphRelVPictures(editor.view.dom, () => plan, new Set([host]))
    expect(img.style.getPropertyValue('--anchor-before')).toBe('0.0px')
    editor.destroy()
  })

  it('keeps a line-relative offset on the line top', async () => {
    const { editor, relV } = await render(anchor('line', 95250))
    expect(relV).toBe('line')
    const img = editor.view.dom.querySelector<HTMLImageElement>('.doc-inline-img-anchor > img')!
    expect(img.style.transform).toBe('translate(calc(96.0px - var(--anchor-ind,0px)), 10.0px)')
    editor.destroy()
  })
})
