/**
 * Enter at the end of a list item must continue the list (Word
 * behavior); Enter in an empty item leaves it. ProseMirror's default splitBlock
 * produced the schema's default block, so every new line dropped out of the list.
 */
import { Editor } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import { buildBlankDocx, parseDocx, saveDocx } from '@genoffice/docx-engine'
import { blocksToPmDoc, pmDocToSavePlan } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

/** The Enter shortcut delegates to this command; commands.keyboardShortcut() builds its own
 * transaction and would discard the command's dispatch, so tests drive the command itself. */
const pressEnter = (editor: Editor) =>
  (editor.commands as unknown as { continueDocList: () => boolean }).continueDocList()

const listEditor = (kind: 'bullet' | 'ordered', text = 'first') =>
  new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docListItem',
          attrs: { kind, numId: '3', ilvl: 0 },
          content: [{ type: 'text', text }],
        },
      ],
    } as never,
  })

const blockKinds = (editor: Editor): string[] => {
  const out: string[] = []
  editor.state.doc.forEach((node) => out.push(node.type.name))
  return out
}

describe('Enter continues a list', () => {
  it.each(['bullet', 'ordered'] as const)(
    '%s: a non-empty item splits into a sibling item',
    (kind) => {
      const editor = listEditor(kind)
      editor.commands.setTextSelection(editor.state.doc.content.size - 1)
      pressEnter(editor)

      expect(blockKinds(editor)).toEqual(['docListItem', 'docListItem'])
      const second = editor.state.doc.child(1)
      expect(second.attrs.kind).toBe(kind)
      expect(second.attrs.numId).toBe('3')
      expect(second.attrs.ilvl).toBe(0)
      // The new item must not claim the original's docx anchor
      expect(second.attrs.docxIndex).toBeNull()

      editor.destroy()
    },
  )

  it('the continued item keeps its level and can still be indented with Tab', () => {
    const editor = listEditor('bullet')
    editor.commands.updateAttributes('docListItem', { ilvl: 1 })
    editor.commands.setTextSelection(editor.state.doc.content.size - 1)
    pressEnter(editor)
    expect(editor.state.doc.child(1).attrs.ilvl).toBe(1)

    editor.commands.keyboardShortcut('Tab')
    expect(editor.state.doc.child(1).attrs.ilvl).toBe(2)
    editor.destroy()
  })

  it('Enter in an empty item leaves the list', () => {
    const editor = listEditor('bullet', '')
    editor.commands.setTextSelection(1)
    pressEnter(editor)
    expect(blockKinds(editor)).toEqual(['docParagraph'])
    editor.destroy()
  })

  it('a continued item saves with its own w:numPr', async () => {
    const parsed = await parseDocx(await buildBlankDocx())
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    editor.commands.setTextSelection(1)
    editor.chain().setNode('docListItem', { kind: 'bullet', numId: '1', ilvl: 0 }).run()
    editor.commands.insertContent('one')
    pressEnter(editor)
    editor.commands.insertContent('two')

    const plan = pmDocToSavePlan(editor.getJSON() as never, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const JSZip = (await import('jszip')).default
    const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
    expect((xml.match(/<w:numPr>/g) ?? []).length).toBe(2)
    expect(xml).toContain('one')
    expect(xml).toContain('two')

    editor.destroy()
  })
})
