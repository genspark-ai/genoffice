/**
 * Tab in a body paragraph inserts a tab character (Word tab-stop), instead of
 * leaving the editor to cycle ribbon buttons (issue 101). List indent and
 * table cell motion keep their existing Tab bindings.
 */
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

const paragraphEditor = (text = 'Hello') =>
  new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [{ type: 'docParagraph', content: [{ type: 'text', text }] }],
    } as never,
  })

/** Drive the real keymap path. `commands.keyboardShortcut` only replays
 *  document steps, so table Tab (a selection-only transaction) would look like a no-op. */
const pressKey = (editor: Editor, key: string, shift = false) => {
  const event = new KeyboardEvent('keydown', {
    key,
    shiftKey: shift,
    bubbles: true,
    cancelable: true,
  })
  editor.view.someProp('handleKeyDown', (fn) => fn(editor.view, event))
}

describe('Tab key inserts a tab stop', () => {
  it('inserts a tab character in a normal paragraph', () => {
    const editor = paragraphEditor()
    editor.commands.setTextSelection(6)
    pressKey(editor, 'Tab')
    expect(editor.state.doc.textContent).toBe('Hello\t')
    editor.destroy()
  })

  it('inserts a tab character in a heading', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docHeading',
            attrs: { level: 1 },
            content: [{ type: 'text', text: 'Title' }],
          },
        ],
      } as never,
    })
    editor.commands.setTextSelection(6)
    pressKey(editor, 'Tab')
    expect(editor.state.doc.textContent).toBe('Title\t')
    editor.destroy()
  })

  it('Shift-Tab in a paragraph does not cycle away or insert a character', () => {
    const editor = paragraphEditor()
    editor.commands.setTextSelection(6)
    pressKey(editor, 'Tab', true)
    expect(editor.state.doc.textContent).toBe('Hello')
    editor.destroy()
  })

  it('still indents a list item', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docListItem',
            attrs: { kind: 'bullet', numId: '3', ilvl: 0 },
            content: [{ type: 'text', text: 'item' }],
          },
        ],
      } as never,
    })
    editor.commands.setTextSelection(2)
    pressKey(editor, 'Tab')
    expect(editor.state.doc.child(0).attrs.ilvl).toBe(1)
    expect(editor.state.doc.textContent).toBe('item')
    editor.destroy()
  })

  it('still moves to the next table cell', async () => {
    const source = await buildDocx({
      bodyXml:
        '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>' +
        '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
        '<w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>' +
        '<w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
    })
    const parsed = await parseDocx(source)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const cells: number[] = []
    editor.state.doc.descendants((node, pos) => {
      if (node.type.name === 'docTableCell') cells.push(pos)
    })
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, cells[0] + 2)),
    )
    const from = editor.state.selection.from
    pressKey(editor, 'Tab')
    expect(editor.state.selection.from).toBeGreaterThan(from)
    expect(editor.state.doc.textContent).toBe('AB')
    editor.destroy()
  })

  it('saves the inserted tab as w:tab', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>Hi</w:t></w:r></w:p>' }),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    editor.commands.setTextSelection(3)
    pressKey(editor, 'Tab')
    editor.commands.insertContent('there')

    const plan = pmDocToSavePlan(editor.getJSON() as never, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const JSZip = (await import('jszip')).default
    const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
    expect(xml).toContain('<w:tab/>')
    expect(xml).toContain('Hi')
    expect(xml).toContain('there')
    editor.destroy()
  })
})
