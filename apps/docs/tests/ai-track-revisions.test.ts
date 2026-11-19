import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import {
  BLANK_BULLET_NUM_ID,
  BLANK_ORDERED_NUM_ID,
  buildBlankDocx,
  parseDocx,
  saveDocx,
} from '@genoffice/docx-engine'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { applyRevisionsBy, collectRevisions } from '../src/renderer/editor/revisions'
import { executeTool } from '../src/renderer/ai/tools'

/**
 * Guard tests for AI edits recorded as tracked revisions (review accept/reject
 * instead of the yellow-highlight flow): insert_content marks new text ins,
 * replace_blocks keeps the old text struck through (del) before the new.
 */

const NUM_IDS = { bullet: BLANK_BULLET_NUM_ID, ordered: BLANK_ORDERED_NUM_ID }
const TRACK = { author: 'AI Assistant' }

async function createEditor(withText = false) {
  const { editorExtensions } = await import('../src/renderer/editor/extensions')
  const parsed = await parseDocx(await buildBlankDocx())
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
  })
  editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
  if (withText) {
    await executeTool(
      editor,
      {
        id: 'seed',
        name: 'insert_content',
        input: { html: '<p>Original first paragraph.</p><p>Original second paragraph.</p>' },
      },
      NUM_IDS,
    )
  }
  return { editor, parsed }
}

/** [text, hasIns, hasDel] triples of every text node */
function inlineMarks(doc: PmNode): Array<[string, boolean, boolean]> {
  const out: Array<[string, boolean, boolean]> = []
  const walk = (node: PmNode) => {
    if (node.type === 'text') {
      const marks = (node.marks ?? []).map((m) => m.type)
      out.push([node.text ?? '', marks.includes('ins'), marks.includes('del')])
    }
    for (const child of node.content ?? []) walk(child)
  }
  walk(doc)
  return out
}

describe('AI edits merged into the revision model', () => {
  it('insert_content in tracking mode marks new text as ins (author "AI Assistant")', async () => {
    const { editor } = await createEditor(true)
    await executeTool(
      editor,
      {
        id: 't',
        name: 'insert_content',
        input: { html: '<p>Paragraph added by AI.</p>', afterBlockIndex: 1 },
      },
      NUM_IDS,
      TRACK,
    )
    const marks = inlineMarks(editor.getJSON() as PmNode)
    const inserted = marks.find(([text]) => text.includes('added by AI'))
    expect(inserted?.[1]).toBe(true)
    // Pre-existing text carries no marks
    expect(marks.find(([text]) => text.includes('Original first paragraph'))?.[1]).toBe(false)
    editor.destroy()
  })

  it('replace_blocks in tracking mode keeps old text struck through + new text underlined', async () => {
    const { editor } = await createEditor(true)
    const exec = await executeTool(
      editor,
      {
        id: 't',
        name: 'replace_blocks',
        input: { startBlockIndex: 0, endBlockIndex: 0, html: '<p>Rewritten first paragraph.</p>' },
      },
      NUM_IDS,
      TRACK,
    )
    expect(exec.isError).toBeFalsy()
    const marks = inlineMarks(editor.getJSON() as PmNode)
    expect(marks.find(([text]) => text.includes('Original first paragraph'))?.[2]).toBe(true) // del
    expect(marks.find(([text]) => text.includes('Rewritten'))?.[1]).toBe(true) // ins
    expect(marks.find(([text]) => text.includes('Original second paragraph'))).toEqual([
      'Original second paragraph.',
      false,
      false,
    ])
    editor.destroy()
  })

  it('accept all: struck-through paragraphs removed entirely; new text kept with marks stripped', async () => {
    const { editor } = await createEditor(true)
    await executeTool(
      editor,
      {
        id: 't',
        name: 'replace_blocks',
        input: { startBlockIndex: 0, endBlockIndex: 0, html: '<p>Rewritten first paragraph.</p>' },
      },
      NUM_IDS,
      TRACK,
    )
    applyRevisionsBy(editor, 'AI Assistant', 'accept')
    const text = editor.state.doc.textContent
    expect(text).toContain('Rewritten first paragraph')
    expect(text).not.toContain('Original first paragraph')
    expect(inlineMarks(editor.getJSON() as PmNode).every(([, ins, del]) => !ins && !del)).toBe(true)
    // Whole-paragraph deletion leaves no empty shell behind
    expect(editor.state.doc.childCount).toBe(2)
    editor.destroy()
  })

  it('reject all: inserted paragraphs removed entirely; original text kept with marks stripped', async () => {
    const { editor } = await createEditor(true)
    await executeTool(
      editor,
      {
        id: 't',
        name: 'replace_blocks',
        input: { startBlockIndex: 0, endBlockIndex: 0, html: '<p>Rewritten first paragraph.</p>' },
      },
      NUM_IDS,
      TRACK,
    )
    applyRevisionsBy(editor, 'AI Assistant', 'reject')
    const text = editor.state.doc.textContent
    expect(text).toContain('Original first paragraph')
    expect(text).not.toContain('Rewritten')
    expect(editor.state.doc.childCount).toBe(2)
    editor.destroy()
  })

  it('revision marks are saved as w:ins / w:del so Word can accept/reject them', async () => {
    const { editor, parsed } = await createEditor(true)
    await executeTool(
      editor,
      {
        id: 't',
        name: 'replace_blocks',
        input: { startBlockIndex: 0, endBlockIndex: 0, html: '<p>Rewritten first paragraph.</p>' },
      },
      NUM_IDS,
      TRACK,
    )
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    editor.destroy()
    const reparsed = await parseDocx(saved)
    const xml = reparsed.internal.documentXml
    expect(xml).toMatch(/<w:ins w:id="\d+" w:author="AI Assistant"/)
    expect(xml).toMatch(/<w:del w:id="\d+" w:author="AI Assistant"/)
    expect(xml).toContain('<w:delText xml:space="preserve">Original first paragraph.</w:delText>')
  })

  it('first generation into a blank document: replaces the empty paragraph directly, content carries ins marks', async () => {
    const { editor } = await createEditor(false)
    await executeTool(
      editor,
      { id: 't', name: 'insert_content', input: { html: '<h1>Title</h1><p>Body text.</p>' } },
      NUM_IDS,
      TRACK,
    )
    // No orphan empty paragraph left behind
    expect(editor.state.doc.childCount).toBe(2)
    const marks = inlineMarks(editor.getJSON() as PmNode)
    expect(marks.every(([, ins]) => ins)).toBe(true)
    editor.destroy()
  })

  it('tracks structural AI replacement containing a table as block revisions', async () => {
    const { editor } = await createEditor(false)
    await executeTool(
      editor,
      {
        id: 't1',
        name: 'insert_content',
        input: { html: '<p>Some text.</p><table><tr><th>A</th></tr><tr><td>1</td></tr></table>' },
      },
      NUM_IDS,
    )
    const exec = await executeTool(
      editor,
      {
        id: 't2',
        name: 'replace_blocks',
        input: { startBlockIndex: 0, endBlockIndex: 1, html: '<p>Replacement text.</p>' },
      },
      NUM_IDS,
      TRACK,
    )
    expect(exec.isError).toBeFalsy()
    expect(editor.state.doc.textContent).toContain('Some text.')
    expect(collectRevisions(editor.state.doc).map((revision) => revision.kind)).toEqual([
      'blockDel',
      'blockDel',
      'blockIns',
    ])
    expect(inlineMarks(editor.getJSON() as PmNode).every(([, ins, del]) => !ins && !del)).toBe(true)
    applyRevisionsBy(editor, TRACK.author, 'accept')
    expect(editor.state.doc.textContent).not.toContain('Some text.')
    expect(editor.state.doc.textContent).toContain('Replacement text.')
    editor.destroy()
  })

  it('apply_ops format changes in tracking mode record rPrChange; reject restores the old format', async () => {
    const { editor } = await createEditor(true)
    const exec = await executeTool(
      editor,
      {
        id: 't',
        name: 'apply_ops',
        input: {
          ops: [{ op: 'setFont', target: { blockIndexes: [0] }, bold: true, color: 'FF0000' }],
        },
      },
      NUM_IDS,
      TRACK,
    )
    expect(exec.isError).toBeFalsy()
    const revs = collectRevisions(editor.state.doc)
    expect(revs).toMatchObject([{ kind: 'rPrChange', author: 'AI Assistant' }])
    const boldAt1 = () => !!editor.state.doc.nodeAt(1)?.marks.some((m) => m.type.name === 'bold')
    expect(boldAt1()).toBe(true)
    applyRevisionsBy(editor, TRACK.author, 'reject')
    expect(collectRevisions(editor.state.doc)).toHaveLength(0)
    expect(boldAt1()).toBe(false)
    const ts = editor.state.doc.nodeAt(1)?.marks.find((m) => m.type.name === 'docTextStyle')
    expect(ts?.attrs.color ?? null).toBeNull()
    // recorder stays off for untracked edits afterwards
    editor.commands.insertContentAt(1, 'x')
    expect(collectRevisions(editor.state.doc)).toHaveLength(0)
    editor.destroy()
  })

  it('apply_ops paragraph changes in tracking mode record pPrChange; reject restores alignment', async () => {
    const { editor } = await createEditor(true)
    await executeTool(
      editor,
      {
        id: 't',
        name: 'apply_ops',
        input: {
          ops: [{ op: 'setParagraphFormat', target: { blockIndexes: [0] }, align: 'center' }],
        },
      },
      NUM_IDS,
      TRACK,
    )
    expect(collectRevisions(editor.state.doc)).toMatchObject([
      { kind: 'pPrChange', author: 'AI Assistant' },
    ])
    expect(editor.state.doc.firstChild?.attrs.align).toBe('center')
    applyRevisionsBy(editor, TRACK.author, 'reject')
    expect(collectRevisions(editor.state.doc)).toHaveLength(0)
    expect(editor.state.doc.firstChild?.attrs.align ?? null).toBeNull()
    editor.destroy()
  })
})
