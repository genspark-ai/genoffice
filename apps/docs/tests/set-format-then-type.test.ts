/**
 * With a collapsed cursor (no selection), setting two character
 * formats in a row must keep both - font then size used to drop the font,
 * because the ribbon rebuilt the whole docTextStyle mark from a stale
 * getAttributes read-back. The ribbon now passes only the patch and lets
 * setMark merge with the caret's stored mark.
 */
import { Editor } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import { buildBlankDocx, parseDocx, saveDocx } from '@genoffice/docx-engine'
import { blocksToPmDoc, pmDocToSavePlan } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

const strOf = async (bytes: Uint8Array, name: string): Promise<string> => {
  const JSZip = (await import('jszip')).default
  const zip = await JSZip.loadAsync(bytes)
  return zip.file(name)!.async('string')
}

describe('set format then type', () => {
  it('font → size → type keeps both on the stored mark and in the saved file', async () => {
    const parsed = await parseDocx(await buildBlankDocx())
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    editor.commands.setTextSelection(1)

    // two consecutive ribbon-style calls, no selection in between
    editor.chain().setMark('docTextStyle', { font: 'Times New Roman' }).run()
    editor.chain().setMark('docTextStyle', { sizeHalfPoints: 48 }).run()

    const stored = editor.state.storedMarks?.find((m) => m.type.name === 'docTextStyle')
    expect(stored?.attrs.font).toBe('Times New Roman')
    expect(stored?.attrs.sizeHalfPoints).toBe(48)

    // typing carries the stored mark; the saved run has both w:rFonts and w:sz
    editor.view.dispatch(editor.state.tr.insertText('hello'))
    const plan = pmDocToSavePlan(editor.getJSON() as never, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const docXml = await strOf(saved, 'word/document.xml')
    expect(docXml).toContain('Times New Roman')
    expect(docXml).toMatch(/<w:sz w:val="48"/)

    editor.destroy()
  })

  it('reverse order (size → font) keeps both too', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: { type: 'doc', content: [{ type: 'docParagraph' }] } as never,
    })
    editor.commands.setTextSelection(1)
    editor.chain().setMark('docTextStyle', { sizeHalfPoints: 48 }).run()
    editor.chain().setMark('docTextStyle', { font: 'Times New Roman' }).run()
    const stored = editor.state.storedMarks?.find((m) => m.type.name === 'docTextStyle')
    expect(stored?.attrs.sizeHalfPoints).toBe(48)
    expect(stored?.attrs.font).toBe('Times New Roman')
    editor.destroy()
  })

  it('with a selection, partial patches still merge per run', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'abc' }] }],
      } as never,
    })
    editor.commands.setTextSelection({ from: 1, to: 4 })
    editor.chain().setMark('docTextStyle', { font: 'Georgia' }).run()
    editor.chain().setMark('docTextStyle', { sizeHalfPoints: 24 }).run()
    const marks: Record<string, unknown>[] = []
    editor.state.doc.descendants((node) => {
      node.marks.forEach((m) => m.type.name === 'docTextStyle' && marks.push(m.attrs))
      return true
    })
    expect(marks[0]?.font).toBe('Georgia')
    expect(marks[0]?.sizeHalfPoints).toBe(24)
    editor.destroy()
  })
})
