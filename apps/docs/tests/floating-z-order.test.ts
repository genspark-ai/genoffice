/**
 * Shape Format Arrange: the stacking commands the ribbon and the object
 * context menu share. Ranks are document-global among floating anchors, and
 * an inline anchor is promoted to a front float so the reorder is visible —
 * Word's "Bring to Front" on an inline picture does the same.
 */
import { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import { describe, expect, it } from 'vitest'
import { parseDocx, type TextboxDisplay } from '@genoffice/docx-engine'
import {
  buildDocx,
  IMAGE_PARAGRAPH_XML,
} from '../../../packages/docx-engine/tests/helpers/build-docx'
import { insertShapeAt } from '../src/renderer/components/ribbon-tabs'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  boxWithZ,
  bringForward,
  bringToFront,
  sendBackward,
  sendToBack,
  setFloatingWrap,
  shapeWrapOf,
} from '../src/renderer/editor/floating-z-order'

async function openDoc(bodyXml: string, withImage = false): Promise<Editor> {
  const source = await buildDocx({ bodyXml, withImage })
  const parsed = await parseDocx(source)
  const element = document.createElement('div')
  document.body.appendChild(element)
  return new Editor({
    element,
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
}

async function openAnchors(): Promise<Editor> {
  return openDoc(IMAGE_PARAGRAPH_XML + IMAGE_PARAGRAPH_XML, true)
}

function insertShape(editor: Editor): number {
  insertShapeAt(editor, 'rect')
  let pos = -1
  editor.state.doc.descendants((node, at) => {
    if (node.type.name === 'docProtected' && node.attrs.textboxes) pos = at
    return true
  })
  editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, pos)))
  return pos
}

function boxesOf(editor: Editor, pos: number): TextboxDisplay[] {
  return editor.state.doc.nodeAt(pos)!.attrs.textboxes as TextboxDisplay[]
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
  const container = (editor.view.dom as HTMLElement).parentElement
  editor.destroy()
  container?.remove()
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

  it('steps a behind anchor one rank and keeps its wrap', async () => {
    const editor = await openAnchors()
    setAttrs(editor, 0, { imageWrap: 'behind', imageZOrder: -1 })
    select(editor, 0)
    bringForward(editor)
    expect(attrsOf(editor, 0)).toEqual({ imageWrap: 'behind', imageZOrder: 0 })
    close(editor)
    expect(document.body.children).toHaveLength(0)
  })
})

describe('shape box-model mapping', () => {
  const box = (extra: Partial<TextboxDisplay> = {}): TextboxDisplay => ({ paras: [], ...extra })

  it('shapeWrapOf reads the mode out of the box fields', () => {
    expect(shapeWrapOf(box({ behind: true }))).toBe('behind')
    expect(shapeWrapOf(box({ wrapSide: 'left' }))).toBe('square-left')
    expect(shapeWrapOf(box({ wrapSide: 'right' }))).toBe('square-right')
    expect(shapeWrapOf(box({ bandBottomPx: 40 }))).toBe('topBottom')
    expect(shapeWrapOf(box({ floating: true }))).toBe('front')
    expect(shapeWrapOf(box())).toBeNull()
  })

  it('boxWithZ floats an in-flow box, preserves a side wrap and a float mode', () => {
    expect(boxWithZ(box(), 2)).toEqual({
      paras: [],
      z: 2,
      floating: true,
      behind: false,
      noWrap: true,
    })
    const side = box({ wrapSides: true, wrapSide: 'left', wrapEdgePx: 0, wrapGapPx: 12 })
    expect(boxWithZ(side, 3)).toEqual({ ...side, z: 3 })
    const floater = box({ floating: true, behind: true, noWrap: true, z: 1 })
    expect(boxWithZ(floater, 4)).toEqual({ ...floater, z: 4 })
  })
})

describe('shape wrap and stacking commands', () => {
  it('setFloatingWrap(behind) writes the shape box and mirrors the node attr', async () => {
    const editor = await openDoc('<w:p><w:r><w:t>Body text</w:t></w:r></w:p>')
    const pos = insertShape(editor)
    setFloatingWrap(editor, 'behind')
    const boxes = boxesOf(editor, pos)
    expect(boxes[0].behind).toBe(true)
    expect(boxes[0].floating).toBe(true)
    expect(editor.state.doc.nodeAt(pos)!.attrs.imageWrap).toBe('behind')
    close(editor)
  })

  it('setFloatingWrap(square-left) sets the box side wrap', async () => {
    const editor = await openDoc('<w:p><w:r><w:t>Body text</w:t></w:r></w:p>')
    const pos = insertShape(editor)
    setFloatingWrap(editor, 'square-left')
    const boxes = boxesOf(editor, pos)
    expect(boxes[0].wrapSides).toBe(true)
    expect(boxes[0].wrapSide).toBe('left')
    expect(editor.state.doc.nodeAt(pos)!.attrs.imageWrap).toBe('square-left')
    close(editor)
  })

  it('setFloatingWrap(null) clears every wrap field from the box', async () => {
    const editor = await openDoc('<w:p><w:r><w:t>Body text</w:t></w:r></w:p>')
    const pos = insertShape(editor)
    setFloatingWrap(editor, 'behind')
    setFloatingWrap(editor, null)
    const box = boxesOf(editor, pos)[0]
    expect('floating' in box).toBe(false)
    expect('behind' in box).toBe(false)
    expect('noWrap' in box).toBe(false)
    expect('wrapSides' in box).toBe(false)
    expect('wrapSide' in box).toBe(false)
    expect('bandTopPx' in box).toBe(false)
    expect('bandBottomPx' in box).toBe(false)
    expect(shapeWrapOf(box)).toBeNull()
    expect(editor.state.doc.nodeAt(pos)!.attrs.imageWrap).toBeNull()
    close(editor)
  })

  it('bringToFront ranks a shape on the box model', async () => {
    const editor = await openDoc('<w:p><w:r><w:t>Body text</w:t></w:r></w:p>')
    const pos = insertShape(editor)
    bringToFront(editor)
    const boxes = boxesOf(editor, pos)
    expect(boxes[0].z).toBe(1)
    expect(boxes[0].floating).toBe(true)
    expect(editor.state.doc.nodeAt(pos)!.attrs.imageZOrder).toBe(1)
    close(editor)
  })

  it('bringToFront shares the document rank scale with floating images', async () => {
    const editor = await openDoc(
      '<w:p><w:r><w:t>Body text</w:t></w:r></w:p>' + IMAGE_PARAGRAPH_XML,
      true,
    )
    let imagePos = -1
    editor.state.doc.descendants((node, at) => {
      if (node.type.name === 'docProtected' && node.attrs.blockType === 'image') imagePos = at
      return true
    })
    editor.view.dispatch(
      editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, imagePos)),
    )
    editor.commands.updateAttributes('docProtected', { imageWrap: 'front', imageZOrder: 3 })

    const pos = insertShape(editor)
    bringToFront(editor)
    expect(boxesOf(editor, pos)[0].z).toBe(4)
    close(editor)
  })
})
