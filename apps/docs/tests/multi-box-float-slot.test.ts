/**
 * A paragraph anchoring several shapes renders one node with every box
 * positioned by its own float slot: the editor's chrome rotator exists only
 * for a single shape, so a multi-box node must keep the slot on the boxes
 * themselves after the sub-editors mount (a group of nine shapes otherwise
 * stacks in flow at the column edge).
 */
import { Editor } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import { buildShapeParagraphXml, parseDocx, type TextboxDisplay } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { textboxBoxStyle } from '../src/renderer/editor/protected-render'

const shapeRun = (id: number, offsetEmu: number): string =>
  buildShapeParagraphXml({ prst: 'roundRect', withTextbox: true, id })
    .replace(
      '<wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH>',
      `<wp:positionH relativeFrom="column"><wp:posOffset>${offsetEmu}</wp:posOffset></wp:positionH>`,
    )
    .replace(/^<w:p[^>]*>/, '')
    .replace(/<\/w:p>$/, '')

describe('multi-box anchor node', () => {
  it('keeps every floating box on its own slot after the sub-editors mount', async () => {
    const source = await buildDocx({
      bodyXml: `<w:p>${shapeRun(1, 500000)}${shapeRun(2, 2600000)}${shapeRun(3, 4700000)}</w:p>`,
    })
    const parsed = await parseDocx(source)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    let pos = -1
    editor.state.doc.descendants((node, at) => {
      if (pos === -1 && node.type.name === 'docProtected' && node.attrs.textboxes) pos = at
      return true
    })
    const boxes = editor.state.doc.nodeAt(pos)!.attrs.textboxes as TextboxDisplay[]
    expect(boxes).toHaveLength(3)
    const wrapper = editor.view.nodeDOM(pos) as HTMLElement
    expect(wrapper.querySelector('.box-rotator')).toBeNull()
    const els = Array.from(wrapper.querySelectorAll<HTMLElement>(':scope > .doc-textbox'))
    expect(els).toHaveLength(3)
    els.forEach((el, i) => {
      expect(el.querySelector('.doc-textbox-editor')).not.toBeNull()
      const reference = document.createElement('div')
      reference.setAttribute('style', textboxBoxStyle(boxes[i]))
      expect(reference.style.position).toBe('absolute')
      expect(el.style.position).toBe('absolute')
      expect(el.style.left).toBe(reference.style.left)
      expect(el.style.top).toBe(reference.style.top)
    })
    // an external model change (undo, AI edit) re-applies the same slot
    editor.commands.updateAttributes('docProtected', {
      textboxes: boxes.map((b) => ({ ...b, rotDeg: 15 })),
    })
    els.forEach((el) => {
      expect(el.style.position).toBe('absolute')
      expect(el.style.transform).toBe('rotate(15deg)')
    })
    editor.destroy()
  })
})
