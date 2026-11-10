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
import { DOMSerializer } from '@tiptap/pm/model'
import JSZip from 'jszip'
import { buildBlankDocx, parseDocx, saveDocx, type Block } from '@genoffice/docx-engine'
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

/** The clipboard as the editor writes it: the schema's DOMSerializer over the
 *  copied slice. Runtime attrs (fieldId/fieldPart) are not rendered, so the
 *  paste has to rebuild the field structure from the instruction alone. */
function copyHtml(editor: Editor, from: number, to: number): string {
  const fragment = editor.state.doc.slice(from, to).content
  const holder = document.createElement('div')
  holder.appendChild(DOMSerializer.fromSchema(editor.schema).serializeFragment(fragment))
  return holder.innerHTML
}

const citationRun = (text: string, fieldId: number, fieldPart: string) => ({
  type: 'text',
  text,
  marks: [{ type: 'instrField', attrs: { instr: ZOTERO_INSTR, fieldId, fieldPart } }],
})

async function saveAndReopen(editor: Editor, parsed: Awaited<ReturnType<typeof parseDocx>>) {
  const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
  const bytes = await saveDocx(parsed, plan.saveBlocks)
  const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')
  const blocks = (await parseDocx(bytes)).blocks.filter((block) => !block.hidden)
  return { xml, blocks }
}

const fieldTokens = (xml: string) =>
  [...xml.matchAll(/w:fldCharType="(begin|end)"/g)].map((m) => m[1])
const partsOf = (blocks: Block[]) =>
  blocks.flatMap((block) => (block.runs ?? []).flatMap((run) => run.zoteroFieldPart ?? []))
const idsOf = (blocks: Block[]) =>
  new Set(blocks.flatMap((block) => (block.runs ?? []).flatMap((run) => run.zoteroFieldId ?? [])))

describe('a copied citation pastes back as a balanced field', () => {
  async function bibliographyEditor() {
    const parsed = await parseDocx(await buildBlankDocx())
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    editor.commands.setContent({
      type: 'doc',
      content: [
        { type: 'docParagraph', content: [citationRun('Alpha (2024)', 1, 'begin')] },
        { type: 'docParagraph', content: [citationRun('Beta (2023)', 1, 'inside')] },
        { type: 'docParagraph', content: [citationRun('Gamma (2022)', 1, 'end')] },
      ],
    })
    return { parsed, editor }
  }

  it('a whole multi-paragraph citation pastes as one begin/inside/end field', async () => {
    const source = await bibliographyEditor()
    const target = await bibliographyEditor()
    try {
      const html = copyHtml(source.editor, 0, source.editor.state.doc.content.size)
      expect(html).toContain('data-instr-field')
      target.editor.commands.setContent({ type: 'doc', content: [{ type: 'docParagraph' }] })
      target.editor.commands.insertContent(html)
      const { xml, blocks } = await saveAndReopen(target.editor, target.parsed)
      expect(fieldTokens(xml)).toEqual(['begin', 'end'])
      expect(partsOf(blocks)).toEqual(['begin', 'inside', 'end'])
      expect(idsOf(blocks).size).toBe(1)
    } finally {
      source.editor.destroy()
      target.editor.destroy()
    }
  })

  it('a paragraph copied out of the middle of a field pastes as a complete field', async () => {
    const source = await bibliographyEditor()
    try {
      const middle = source.editor.state.doc.child(0).nodeSize
      const html = copyHtml(source.editor, middle + 1, middle + 1 + 'Beta (2023)'.length)
      source.editor.commands.setContent({
        type: 'doc',
        content: [
          { type: 'docParagraph', content: [{ type: 'text', text: 'before ' }] },
          { type: 'docParagraph', content: [{ type: 'text', text: 'after' }] },
        ],
      })
      source.editor.commands.setTextSelection(source.editor.state.doc.child(0).nodeSize - 1)
      source.editor.commands.insertContent(html)
      const { xml, blocks } = await saveAndReopen(source.editor, source.parsed)
      expect(fieldTokens(xml)).toEqual(['begin', 'end'])
      expect(blocks.map((block) => block.runs?.map((run) => run.text).join(''))).toEqual([
        'before Beta (2023)',
        'after',
      ])
      expect(partsOf(blocks)).toEqual(['single'])
    } finally {
      source.editor.destroy()
    }
  })

  it('pasting a citation inside another citation ends the first instead of nesting', async () => {
    const parsed = await parseDocx(await buildBlankDocx())
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [{ type: 'docParagraph', content: [citationRun('(Smith, 2020)', 1, 'single')] }],
      },
    })
    try {
      editor.commands.setTextSelection(1 + '(Smi'.length)
      editor.commands.insertContent(
        clipboardHtml(ZOTERO_INSTR.replace('ITEM-1', 'ITEM-2'), '(Doe, 2021)'),
      )
      const { xml, blocks } = await saveAndReopen(editor, parsed)
      expect(fieldTokens(xml)).toEqual(['begin', 'end', 'begin', 'end', 'begin', 'end'])
      expect(blocks[0].runs?.map((run) => run.text)).toEqual(['(Smi', '(Doe, 2021)', 'th, 2020)'])
      expect(partsOf(blocks)).toEqual(['single', 'single', 'single'])
      expect(idsOf(blocks).size).toBe(3)
    } finally {
      editor.destroy()
    }
  })
})
