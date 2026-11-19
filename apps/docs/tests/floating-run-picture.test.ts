import { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import { parseDocx, saveDocx, type Run } from '@genoffice/docx-engine'
import { beforeAll, describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { setFloatingWrap } from '../src/renderer/editor/floating-z-order'

const EMU = 9525

/** a Word picture floating beside the text of its own paragraph (wrapSquare) */
const anchoredPicture = (wrap = '<wp:wrapSquare wrapText="bothSides"/>') =>
  '<w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" ' +
  'relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/>' +
  '<wp:positionH relativeFrom="column"><wp:posOffset>476250</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>95250</wp:posOffset></wp:positionV>' +
  `<wp:extent cx="1905000" cy="1270000"/>${wrap}<wp:docPr id="1" name="Picture 1"/>` +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="Picture 1"/><pic:cNvPicPr/></pic:nvPicPr>' +
  '<pic:blipFill><a:blip r:embed="rId10"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="1270000"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
  '</a:graphicData></a:graphic></wp:anchor></w:drawing>'

async function openRunPictureDoc() {
  const source = await buildDocx({
    bodyXml:
      `<w:p><w:r><w:t>Text that wraps around the picture.</w:t></w:r><w:r>${anchoredPicture()}</w:r></w:p>` +
      '<w:p><w:r><w:t>Next paragraph.</w:t></w:r></w:p>',
    withImage: true,
  })
  const parsed = await parseDocx(source)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  let picturePos = -1
  editor.state.doc.descendants((node, pos) => {
    if (node.type.name === 'docInlineImage') picturePos = pos
  })
  return { editor, parsed, source, picturePos }
}

async function savedRunPicture(editor: Editor, parsed: Awaited<ReturnType<typeof parseDocx>>) {
  const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
  const reparsed = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
  const runs = reparsed.blocks.flatMap((block) => ('runs' in block ? (block.runs ?? []) : []))
  return { plan, image: runs.find((run: Run) => run.image)?.image }
}

const select = (editor: Editor, pos: number) =>
  editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, pos)))

const patchPicture = (editor: Editor, pos: number, attrs: Record<string, unknown>) => {
  const node = editor.state.doc.nodeAt(pos)!
  editor.view.dispatch(editor.state.tr.setNodeMarkup(pos, undefined, { ...node.attrs, ...attrs }))
}

describe('floating run pictures (anchored drawings sharing a text paragraph)', () => {
  // jsdom has no hit testing: a press reaches the picture by its target alone
  beforeAll(() => {
    const doc = document as unknown as Record<string, unknown>
    doc.caretRangeFromPoint ??= () => null
    doc.elementsFromPoint ??= () => []
  })

  it('opens as a floating run picture and saves byte-identical when untouched', async () => {
    const { editor, parsed, source, picturePos } = await openRunPictureDoc()
    expect(editor.state.doc.nodeAt(picturePos)?.attrs.wrap).toBe('square-left')
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(0)
    expect(await saveDocx(parsed, plan.saveBlocks)).toEqual(source)
    editor.destroy()
  })

  it('saves a moved and resized picture', async () => {
    const { editor, parsed, picturePos } = await openRunPictureDoc()
    patchPicture(editor, picturePos, {
      offsetXEmu: 150 * EMU,
      offsetYEmu: 40 * EMU,
      widthPx: 260,
      heightPx: 173,
    })
    const { plan, image } = await savedRunPicture(editor, parsed)
    expect(plan.changedCount).toBe(1)
    expect([image?.offsetXEmu, image?.offsetYEmu]).toEqual([150 * EMU, 40 * EMU])
    expect([image?.widthPx, image?.heightPx]).toEqual([260, 173])
    expect(image?.wrap).toBe('square-left')
    editor.destroy()
  })

  it('changes the wrap through Wrap Text and back in line', async () => {
    const { editor, parsed, picturePos } = await openRunPictureDoc()
    select(editor, picturePos)
    setFloatingWrap(editor, 'behind')
    expect(editor.state.doc.nodeAt(picturePos)?.attrs.wrap).toBe('behind')
    expect((await savedRunPicture(editor, parsed)).image?.wrap).toBe('behind')

    setFloatingWrap(editor, null)
    const inline = editor.state.doc.nodeAt(picturePos)!
    expect([inline.attrs.wrap, inline.attrs.offsetXEmu, inline.attrs.offsetYEmu]).toEqual([
      null,
      null,
      null,
    ])
    const { image } = await savedRunPicture(editor, parsed)
    expect(image?.wrap).toBeUndefined()
    expect(image?.xml).toContain('<wp:inline')
    // back in the flow it gets the in-line picture's selection handles
    expect(editor.view.dom.querySelector('.doc-inline-pic .img-resize-handle')).toBeTruthy()
    editor.destroy()
  })

  it('drags the picture by its body: the drop becomes its anchor offsets', async () => {
    const { editor, picturePos } = await openRunPictureDoc()
    const img = editor.view.nodeDOM(picturePos) as HTMLImageElement
    expect(img.tagName).toBe('IMG')
    img.dispatchEvent(
      new MouseEvent('mousedown', { bubbles: true, button: 0, clientX: 100, clientY: 100 }),
    )
    expect(editor.state.selection instanceof NodeSelection && editor.state.selection.from).toBe(
      picturePos,
    )
    window.dispatchEvent(new MouseEvent('mousemove', { clientX: 140, clientY: 130 }))
    expect(img.style.transform).toBe('translate(40px, 30px)')
    window.dispatchEvent(new MouseEvent('mouseup', { clientX: 140, clientY: 130 }))
    const moved = editor.state.doc.nodeAt(picturePos)!
    expect(moved.attrs.offsetXEmu).toBe(476250 + 40 * EMU)
    expect(moved.attrs.offsetYEmu).toBe(95250 + 30 * EMU)
    expect(moved.attrs.wrap).toBe('square-left')
    editor.destroy()
  })

  it('a plain click selects the picture without moving it', async () => {
    const { editor, picturePos } = await openRunPictureDoc()
    const before = editor.state.doc.nodeAt(picturePos)!.attrs
    const img = editor.view.nodeDOM(picturePos) as HTMLImageElement
    img.dispatchEvent(
      new MouseEvent('mousedown', { bubbles: true, button: 0, clientX: 50, clientY: 50 }),
    )
    window.dispatchEvent(new MouseEvent('mouseup', { clientX: 51, clientY: 50 }))
    expect(editor.state.selection instanceof NodeSelection).toBe(true)
    expect(editor.state.doc.nodeAt(picturePos)!.attrs).toEqual(before)
    editor.destroy()
  })
})
