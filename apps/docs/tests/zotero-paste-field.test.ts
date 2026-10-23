/**
 * A pasted field must keep its instruction.
 *
 * The clipboard carries the instruction in `data-instr-field` (the mark's
 * renderHTML writes it), so the mark's parseHTML has to read it back. Without
 * that, a pasted citation saves as an empty `w:instrText`, the next parse
 * classifies the whole paragraph as a protected passthrough, and the
 * paragraph locks with a blue frame and no caret (genoffice#2099).
 */
import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { buildBlankDocx, parseDocx, saveDocx } from '@genoffice/docx-engine'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

const ZOTERO_INSTR =
  'ADDIN ZOTERO_ITEM CSL_CITATION {"citationItems":[{"id":"ITEM-1"}],"properties":{"formattedCitation":"(Smith, 2020)"}}'
const DATE_INSTR = 'DATE \\@ "yyyy-MM-dd"'

function makeEditor(): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content: [{ type: 'docParagraph' }] },
  })
}

/** Every instrField mark in the document, in document order. */
function fieldInstrs(editor: Editor): string[] {
  const out: string[] = []
  editor.state.doc.descendants((node) => {
    for (const mark of node.marks) {
      if (mark.type.name === 'instrField') out.push(String(mark.attrs.instr))
    }
  })
  return out
}

/** The instruction as it travels on the clipboard: the mark's renderHTML. */
function clipboardHtml(instr: string, text: string): string {
  return `<span data-instr-field='${instr}' class="doc-ref-field">${text}</span>`
}

describe('pasting a field', () => {
  it('keeps a pasted Zotero citation instruction', () => {
    const editor = makeEditor()
    editor.commands.insertContent(clipboardHtml(ZOTERO_INSTR, '(Smith, 2020)'))
    expect(fieldInstrs(editor)).toContain(ZOTERO_INSTR)
  })

  it('keeps a pasted plain field instruction', () => {
    const editor = makeEditor()
    editor.commands.insertContent(clipboardHtml(DATE_INSTR, '2026-01-01'))
    expect(fieldInstrs(editor)).toContain(DATE_INSTR)
  })

  it('does not invent an instruction for a span without one', () => {
    const editor = makeEditor()
    editor.commands.insertContent('<span class="doc-ref-field">plain</span>')
    expect(fieldInstrs(editor)).toEqual([])
  })
})

/** The whole reported cycle: paste a citation, save, reopen. A lost
 *  instruction saves as an empty w:instrText, and the next parse classifies
 *  the paragraph as a protected passthrough — the blue, uneditable frame the
 *  reporter saw after the autosave. */
describe('a pasted citation survives a save and a reopen', () => {
  it('reopens as an editable paragraph, not a protected passthrough', async () => {
    const parsed = await parseDocx(await buildBlankDocx())
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    try {
      editor.commands.insertContent(clipboardHtml(ZOTERO_INSTR, '(Smith, 2020)'))

      const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
      const bytes = await saveDocx(parsed, plan.saveBlocks)

      // the saved field carries the instruction, not an empty one
      const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')
      expect(xml).toContain('ZOTERO_ITEM')
      expect(xml).not.toMatch(/<w:instrText[^>]*>\s*<\/w:instrText>/)

      // and the paragraph it lands in is still editable
      const reopened = await parseDocx(bytes)
      const blocks = reopened.blocks.filter((block) => !block.hidden)
      expect(blocks.map((block) => block.type)).not.toContain('passthrough')
    } finally {
      editor.destroy()
    }
  })
})
