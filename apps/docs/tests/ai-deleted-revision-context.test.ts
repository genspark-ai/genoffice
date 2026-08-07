import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { buildDocumentContext, countWords } from '../src/renderer/ai/protocol'
import { executeTool } from '../src/renderer/ai/tools'

/**
 * Tracked deletions (w:del revisions shown struck through in All Markup) are
 * not current content. The bug: the model saw them as live text and kept
 * "deleting" them in a loop while the tracker re-materialized them. These
 * tests pin the fix: snapshots exclude/tag deleted text, and delete/replace
 * tools report a discriminable skip instead of a silent no-op.
 */

const NUM_IDS = { bullet: null, ordered: null }
const REV = { author: 'User', date: '2026-01-01T00:00:00Z' }

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
  marks?: Array<{ type: string; attrs?: Record<string, unknown> }>
}

const text = (t: string): JsonNode => ({ type: 'text', text: t })
const delText = (t: string): JsonNode => ({
  type: 'text',
  text: t,
  marks: [{ type: 'del', attrs: REV }],
})
const para = (...content: JsonNode[]): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content,
})

const liveEditors: Editor[] = []

afterEach(() => {
  for (const editor of liveEditors.splice(0)) editor.destroy()
})

/** 0 live | 1 fully del-marked | 2 live + del tail | 3 blockRevision del */
function createEditor(): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        para(text('Keep this.')),
        para(delText('Old reference 74.')),
        para(text('New text '), delText('old tail')),
        {
          ...para(text('Old block reference 75.')),
          attrs: { docxIndex: null, blockRevision: { kind: 'del', ...REV } },
        },
      ],
    },
  })
  liveEditors.push(editor)
  return editor
}

describe('document snapshot vs tracked deletions', () => {
  it('block list tags deleted blocks, hides deleted text and excludes it from stats', () => {
    const context = buildDocumentContext(createEditor())
    expect(context).toContain('1|p|[tracked deletion] Old reference 74.')
    expect(context).toContain('3|p|[tracked deletion] Old block reference 75.')
    expect(context).toContain('2|p|New text')
    expect(context).not.toContain('old tail')
    expect(context).toContain(`Full-text stats: words ${countWords('Keep this.New text ')}`)
    expect(context).toContain('Tracked changes:')
  })

  it('read_blocks omits deleted revision text', async () => {
    const r = await executeTool(
      createEditor(),
      { id: 't', name: 'read_blocks', input: { startBlockIndex: 0, endBlockIndex: 3 } },
      NUM_IDS,
    )
    expect(r.output).toContain('Keep this.')
    expect(r.output).toContain('New text')
    expect(r.output).not.toContain('Old reference')
    expect(r.output).not.toContain('old tail')
    expect(r.output).not.toContain('Old block reference')
  })

  it('read_blocks on a deleted-only range says so instead of "(range is empty)"', async () => {
    const r = await executeTool(
      createEditor(),
      { id: 't', name: 'read_blocks', input: { startBlockIndex: 1, endBlockIndex: 1 } },
      NUM_IDS,
    )
    expect(r.output).toContain('pending tracked deletions')
    expect(r.isError).not.toBe(true)
  })

  it('atom-only blocks (formula/ruby) are not mis-tagged as deleted', () => {
    const math = (marks?: JsonNode['marks']): JsonNode => ({
      type: 'docInlineMath',
      attrs: { omml: '', mathml: '', latex: 'E=mc^2', text: 'E = mc 2' },
      ...(marks ? { marks } : {}),
    })
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          para(math()),
          para(delText('old caption '), math()),
          para(math([{ type: 'del', attrs: REV }])),
        ],
      },
    })
    liveEditors.push(editor)
    const context = buildDocumentContext(editor)
    // live formula (alone or next to a deleted run) stays live content
    expect(context).toContain('0|p|')
    expect(context).not.toContain('0|p|[tracked deletion]')
    expect(context).not.toContain('1|p|[tracked deletion]')
    // del-marked formula is a pending deletion like any other content
    expect(context).toContain('2|p|[tracked deletion]')
  })

  it('protected blocks with a block-level deletion are tagged and hidden too', async () => {
    const protectedBlock = (label: string, deleted: boolean): JsonNode => ({
      type: 'docProtected',
      attrs: {
        docxIndex: null,
        blockType: 'image',
        label,
        previewText: `${label} preview`,
        ...(deleted ? { blockRevision: { kind: 'del', ...REV } } : {}),
      },
    })
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [protectedBlock('Old figure', true), protectedBlock('New figure', false)],
      },
    })
    liveEditors.push(editor)
    const context = buildDocumentContext(editor)
    expect(context).toContain('0|Old figure|[tracked deletion] Old figure preview')
    expect(context).toContain('1|New figure|New figure preview')
    expect(context).not.toContain('1|New figure|[tracked deletion]')
    expect(context).toContain('Tracked changes:')
    const r = await executeTool(
      editor,
      { id: 't', name: 'read_blocks', input: { startBlockIndex: 0, endBlockIndex: 0 } },
      NUM_IDS,
    )
    expect(r.output).toContain('pending tracked deletions')
  })

  it('a deleted block between two lists keeps them separate in read_blocks', async () => {
    const li = (t: string): JsonNode => ({
      type: 'docListItem',
      attrs: { docxIndex: null, kind: 'bullet', numId: null, ilvl: 0 },
      content: [text(t)],
    })
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: { type: 'doc', content: [li('alpha'), para(delText('gone')), li('beta')] },
    })
    liveEditors.push(editor)
    const r = await executeTool(
      editor,
      { id: 't', name: 'read_blocks', input: { startBlockIndex: 0, endBlockIndex: 2 } },
      NUM_IDS,
    )
    expect((r.output.match(/<ul>/g) ?? []).length).toBe(2)
    expect(r.output).not.toContain('gone')
  })

  it('a clean document gets no tracked-changes notice', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: { type: 'doc', content: [para(text('Plain paragraph.'))] },
    })
    liveEditors.push(editor)
    const context = buildDocumentContext(editor)
    expect(context).not.toContain('[tracked deletion]')
    expect(context).not.toContain('Tracked changes:')
  })
})

describe('edit tools vs tracked deletions', () => {
  it('deleteBlocks on deleted blocks skips them and tells the model why', async () => {
    const editor = createEditor()
    const r = await executeTool(
      editor,
      {
        id: 't',
        name: 'apply_commands',
        input: { commands: [{ deleteBlocks: { target: { blockIndexes: [1, 3] } } }] },
      },
      NUM_IDS,
    )
    expect(r.mutated).toBe(false)
    expect(r.output).toContain('pending tracked deletion')
    expect(editor.state.doc.childCount).toBe(4)
  })

  it('containsText only matches live text, so deleted-only text matches nothing', async () => {
    const editor = createEditor()
    const r = await executeTool(
      editor,
      {
        id: 't',
        name: 'apply_commands',
        input: { commands: [{ deleteBlocks: { target: { containsText: 'Old reference' } } }] },
      },
      NUM_IDS,
    )
    expect(r.mutated).toBe(false)
    expect(editor.state.doc.childCount).toBe(4)
  })

  it('replaceAllText leaves deleted revision text alone and reports the skip', async () => {
    const editor = createEditor()
    const r = await executeTool(
      editor,
      {
        id: 't',
        name: 'apply_commands',
        input: {
          commands: [{ replaceAllText: { containsText: 'Old reference 74.', replaceText: '[9]' } }],
        },
      },
      NUM_IDS,
    )
    expect(r.mutated).toBe(false)
    expect(r.output).toContain('pending tracked deletion')
    expect(editor.state.doc.textContent).toContain('Old reference 74.')
  })

  it('live targets still work: deleteBlocks by containsText removes the live block', async () => {
    const editor = createEditor()
    const r = await executeTool(
      editor,
      {
        id: 't',
        name: 'apply_commands',
        input: { commands: [{ deleteBlocks: { target: { containsText: 'Keep this' } } }] },
      },
      NUM_IDS,
    )
    expect(r.mutated).toBe(true)
    expect(editor.state.doc.childCount).toBe(3)
  })

  it('tracked replace_blocks leaves a struck copy the next snapshot tags as deleted', async () => {
    const editor = createEditor()
    await executeTool(
      editor,
      {
        id: 't',
        name: 'replace_blocks',
        input: { startBlockIndex: 0, endBlockIndex: 0, html: '<p>Rewritten opening.</p>' },
      },
      NUM_IDS,
      { author: 'AI Assistant' },
    )
    const context = buildDocumentContext(editor)
    expect(context).toContain('[tracked deletion] Keep this.')
    expect(context).toContain('Rewritten opening.')
  })
})
