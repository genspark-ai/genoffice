/**
 * Shape Format Arrange: the stacking commands the ribbon and the object
 * context menu share. Ranks are document-global among floating anchors, and
 * an inline anchor is promoted to a front float so the reorder is visible —
 * Word's "Bring to Front" on an inline picture does the same.
 */
import { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import { describe, expect, it } from 'vitest'
import { parseDocx } from '@genoffice/docx-engine'
import {
  buildDocx,
  IMAGE_PARAGRAPH_XML,
} from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  bringForward,
  bringToFront,
  sendBackward,
  sendToBack,
} from '../src/renderer/editor/floating-z-order'

async function openAnchors(): Promise<Editor> {
  const source = await buildDocx({
    bodyXml: IMAGE_PARAGRAPH_XML + IMAGE_PARAGRAPH_XML,
    withImage: true,
  })
  const parsed = await parseDocx(source)
  const element = document.createElement('div')
  document.body.appendChild(element)
  return new Editor({
    element,
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
}

function anchorPositions(editor: Editor): number[] {
  const out: number[] = []
  editor.state.doc.forEach((node, pos) => {
    if (node.type.name === 'docProtected') out.push(pos)
  })
  return out
}

function select(editor: Editor, index: number): void {
  editor.view.dispatch(
    editor.state.tr.setSelection(
      NodeSelection.create(editor.state.doc, anchorPositions(editor)[index]),
    ),
  )
}

function setAttrs(editor: Editor, index: number, attrs: Record<string, unknown>): void {
  select(editor, index)
  editor.commands.updateAttributes('docProtected', attrs)
}

function attrsOf(
  editor: Editor,
  index: number,
): { imageWrap: string | null; imageZOrder: number | null } {
  const node = editor.state.doc.nodeAt(anchorPositions(editor)[index])!
  return { imageWrap: node.attrs.imageWrap ?? null, imageZOrder: node.attrs.imageZOrder ?? null }
}

function close(editor: Editor): void {
  const element = editor.view.dom as HTMLElement
  editor.destroy()
  element.remove()
}

describe('floating z-order commands', () => {
  it('bringForward promotes an inline anchor to a front float one step up', async () => {
    const editor = await openAnchors()
    select(editor, 0)
    bringForward(editor)
    expect(attrsOf(editor, 0)).toEqual({ imageWrap: 'front', imageZOrder: 1 })
    expect(attrsOf(editor, 1)).toEqual({ imageWrap: null, imageZOrder: null }) // sibling untouched
    close(editor)
  })

  it('sendBackward on an inline anchor floats it in front with rank -1', async () => {
    const editor = await openAnchors()
    select(editor, 0)
    sendBackward(editor)
    expect(attrsOf(editor, 0)).toEqual({ imageWrap: 'front', imageZOrder: -1 })
    close(editor)
  })

  it('bringToFront outranks the highest float in the document, behind ranks included', async () => {
    const editor = await openAnchors()
    setAttrs(editor, 0, { imageWrap: 'front', imageZOrder: 3 })
    setAttrs(editor, 1, { imageWrap: 'behind', imageZOrder: 5 })
    select(editor, 0)
    bringToFront(editor)
    expect(attrsOf(editor, 0)).toEqual({ imageWrap: 'front', imageZOrder: 6 })
    expect(attrsOf(editor, 1)).toEqual({ imageWrap: 'behind', imageZOrder: 5 })
    close(editor)
  })

  it('sendToBack drops below the lowest float in the document', async () => {
    const editor = await openAnchors()
    setAttrs(editor, 0, { imageWrap: 'front', imageZOrder: 2 })
    setAttrs(editor, 1, { imageWrap: 'behind', imageZOrder: -1 })
    select(editor, 0)
    sendToBack(editor)
    expect(attrsOf(editor, 0)).toEqual({ imageWrap: 'front', imageZOrder: -2 })
    close(editor)
  })

  it('bringForward/sendBackward step one rank and leave the float wrap alone', async () => {
    const editor = await openAnchors()
    setAttrs(editor, 0, { imageWrap: 'front', imageZOrder: 3 })
    select(editor, 0)
    bringForward(editor)
    expect(attrsOf(editor, 0)).toEqual({ imageWrap: 'front', imageZOrder: 4 })
    sendBackward(editor)
    expect(attrsOf(editor, 0)).toEqual({ imageWrap: 'front', imageZOrder: 3 })
    close(editor)
  })
})
