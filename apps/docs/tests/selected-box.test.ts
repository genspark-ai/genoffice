/**
 * A paragraph anchoring several drawings packs them into one docProtected
 * node. Pressing one of them records its index, and the ribbon reads and
 * edits that box: a behind-text cover photo next to text boxes is a picture,
 * not the first text box.
 */
import { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import { afterEach, describe, expect, it } from 'vitest'
import { parseDocx, type TextboxDisplay } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { computeFormatState } from '../src/renderer/components/ribbon-format-state'
import { setActiveSubEditor } from '../src/renderer/editor/active-editor'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions, textboxSubExtensions } from '../src/renderer/editor/extensions'
import { sendBackward, setFloatingWrap } from '../src/renderer/editor/floating-z-order'
import { selectBoxAt, selectedBoxIndex } from '../src/renderer/editor/selected-box'

const WPS_NS = 'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
const SHAPE_URI = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
const PIC_URI = 'http://schemas.openxmlformats.org/drawingml/2006/picture'

/** a front text box and a behind-text photo anchored in the same paragraph (a cover page) */
const PARAGRAPH_XML =
  `<w:p><w:r><w:drawing><wp:anchor behindDoc="0" relativeHeight="251660288" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">` +
  `<wp:simplePos x="0" y="0"/>` +
  `<wp:positionH relativeFrom="column"><wp:posOffset>100000</wp:posOffset></wp:positionH>` +
  `<wp:positionV relativeFrom="paragraph"><wp:posOffset>200000</wp:posOffset></wp:positionV>` +
  `<wp:extent cx="952500" cy="476250"/><wp:wrapNone/><wp:docPr id="4" name="Text Box 4"/>` +
  `<a:graphic><a:graphicData uri="${SHAPE_URI}">` +
  `<wps:wsp ${WPS_NS}><wps:cNvPr id="4" name="Text Box 4"/><wps:cNvSpPr txBox="1"/><wps:spPr>` +
  `<a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
  `</wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>title</w:t></w:r></w:p>` +
  `</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>` +
  `</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>` +
  `<w:r><w:drawing><wp:anchor behindDoc="1" relativeHeight="251659264" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">` +
  `<wp:simplePos x="0" y="0"/>` +
  `<wp:positionH relativeFrom="column"><wp:posOffset>-100000</wp:posOffset></wp:positionH>` +
  `<wp:positionV relativeFrom="paragraph"><wp:posOffset>-200000</wp:posOffset></wp:positionV>` +
  `<wp:extent cx="1905000" cy="952500"/><wp:wrapNone/><wp:docPr id="1" name="Picture 2"/>` +
  `<a:graphic><a:graphicData uri="${PIC_URI}">` +
  `<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="Picture 2"/><pic:cNvPicPr/></pic:nvPicPr>` +
  `<pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill>` +
  `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></a:xfrm></pic:spPr>` +
  `</pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>`

let editor: Editor | null = null

async function open(): Promise<{ editor: Editor; pos: number; boxes: TextboxDisplay[] }> {
  const parsed = await parseDocx(await buildDocx({ bodyXml: PARAGRAPH_XML, withImage: true }))
  const element = document.createElement('div')
  document.body.appendChild(element)
  editor = new Editor({
    element,
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  let pos = -1
  editor.state.doc.descendants((node, at) => {
    if (node.type.name === 'docProtected' && node.attrs.textboxes) pos = at
    return pos === -1
  })
  expect(pos).toBeGreaterThanOrEqual(0)
  const boxes = editor.state.doc.nodeAt(pos)!.attrs.textboxes as TextboxDisplay[]
  expect(boxes.map((b) => b.picture ?? false)).toEqual([false, true])
  return { editor, pos, boxes }
}

afterEach(() => {
  const container = (editor?.view.dom as HTMLElement | undefined)?.parentElement
  editor?.destroy()
  container?.remove()
  editor = null
})

describe('pressed box of a multi-drawing node', () => {
  it('a plain node selection addresses the first box', async () => {
    const { editor, pos } = await open()
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, pos)))
    expect(selectedBoxIndex(editor.state)).toBe(0)
    const fs = computeFormatState(editor)
    expect(fs.textboxSelected).toBe(true)
    expect(fs.imageSelected).toBe(false)
    expect(fs.imageWrap).toBe('front')
  })

  it('pressing the photo shows Picture Format state for the photo', async () => {
    const { editor, pos } = await open()
    editor.view.dispatch(selectBoxAt(editor.state.tr, pos, 1))
    expect(selectedBoxIndex(editor.state)).toBe(1)
    const fs = computeFormatState(editor)
    expect(fs.pictureBox).toBe(true)
    expect(fs.imageSelected).toBe(true)
    expect(fs.textboxSelected).toBe(false)
    expect(fs.imageWrap).toBe('behind')
    expect(fs.imageDataUrl).toContain('data:image/png')
    expect(fs.imageWidthPx).toBe(200)
  })

  it('wrap and order edits land on the pressed box and leave the first box and its mirror alone', async () => {
    const { editor, pos, boxes: initial } = await open()
    const initialMirror = editor.state.doc.nodeAt(pos)!.attrs.imageWrap
    editor.view.dispatch(selectBoxAt(editor.state.tr, pos, 1))
    setFloatingWrap(editor, 'front')
    let boxes = editor.state.doc.nodeAt(pos)!.attrs.textboxes as TextboxDisplay[]
    expect(boxes[1].behind).toBeFalsy()
    expect(boxes[1].floating).toBe(true)
    expect(boxes[0]).toEqual(initial[0])
    expect(editor.state.doc.nodeAt(pos)!.attrs.imageWrap).toBe(initialMirror)
    expect(selectedBoxIndex(editor.state)).toBe(1)

    sendBackward(editor)
    boxes = editor.state.doc.nodeAt(pos)!.attrs.textboxes as TextboxDisplay[]
    expect(boxes[1].z).toBe((initial[1].z ?? 0) - 1)
    expect(boxes[0].z).toBe(initial[0].z)
  })

  it('typing into a sibling text box addresses that box, not the pressed photo', async () => {
    const { editor, pos } = await open()
    editor.view.dispatch(selectBoxAt(editor.state.tr, pos, 1))
    const textboxEl = editor.view.dom.querySelector('.doc-textbox:not(.doc-textbox-picture)')!
    expect(textboxEl).toBeTruthy()
    // the node view's own sub-editor is private: mount a stand-in inside the same box
    const host = document.createElement('div')
    textboxEl.appendChild(host)
    const sub = new Editor({ element: host, extensions: textboxSubExtensions })
    setActiveSubEditor(sub)
    try {
      expect(selectedBoxIndex(editor.state)).toBe(0)
      const fs = computeFormatState(editor)
      expect(fs.textboxSelected).toBe(true)
      expect(fs.pictureBox).toBe(false)
    } finally {
      setActiveSubEditor(null)
      sub.destroy()
    }
  })

  it('moving the selection elsewhere forgets the pressed box', async () => {
    const { editor, pos } = await open()
    editor.view.dispatch(selectBoxAt(editor.state.tr, pos, 1))
    editor.commands.setTextSelection(1)
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, pos)))
    expect(selectedBoxIndex(editor.state)).toBe(0)
  })
})
