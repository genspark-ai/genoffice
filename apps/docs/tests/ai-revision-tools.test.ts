import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { collectRevisions } from '../src/renderer/editor/revisions'
import { listRevisionEntries, selectRevisions } from '../src/renderer/ai/revision-ops'
import { buildRevisionsContext } from '../src/renderer/ai/protocol'
import { executeTool } from '../src/renderer/ai/tools'

const NUM_IDS = { bullet: null, ordered: null }
const CAROL = { author: 'Carol', date: '2026-03-01T10:00:00Z' }
const DAVE = { author: 'Dave', date: '2026-05-01T10:00:00Z' }

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
  marks?: Array<{ type: string; attrs?: Record<string, unknown> }>
}

const text = (t: string, ...marks: JsonNode['marks'] & unknown[]): JsonNode => ({
  type: 'text',
  text: t,
  ...(marks.length ? { marks } : {}),
})
const para = (attrs: Record<string, unknown>, ...content: JsonNode[]): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null, ...attrs },
  content,
})

const liveEditors: Editor[] = []
afterEach(() => {
  for (const editor of liveEditors.splice(0)) editor.destroy()
})

/**
 * 0 plain | 1 Carol insertion + Carol deletion | 2 Dave bold formatting change |
 * 3 paragraph alignment change by Dave | 4 block inserted by Carol
 */
function createEditor(): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        para({}, text('Untouched.')),
        para(
          {},
          text('Base '),
          text('added words', { type: 'ins', attrs: CAROL }),
          text(' and '),
          text('gone words', { type: 'del', attrs: CAROL }),
          text('.'),
        ),
        para(
          {},
          text(
            'Now bold',
            { type: 'bold' },
            { type: 'rprChange', attrs: { ...DAVE, old: { bold: false } } },
          ),
        ),
        para(
          {
            align: 'center',
            pPrChange: JSON.stringify({ ...DAVE, old: { align: 'left' } }),
          },
          text('Centered now.'),
        ),
        para({ blockRevision: { kind: 'ins', ...CAROL } }, text('Whole new paragraph.')),
      ],
    },
  })
  liveEditors.push(editor)
  return editor
}

const bodyText = (editor: Editor) =>
  editor.state.doc.textBetween(0, editor.state.doc.content.size, '|')

describe('revision entries', () => {
  it('lists positional ids, types, block indexes and formatting diffs', () => {
    const entries = listRevisionEntries(createEditor().state.doc)
    expect(entries.map((e) => [e.id, e.type, e.blockIndex, e.author])).toEqual([
      ['r1', 'insertion', 1, 'Carol'],
      ['r2', 'deletion', 1, 'Carol'],
      ['r3', 'formatting', 2, 'Dave'],
      ['r4', 'formatting', 3, 'Dave'],
      ['r5', 'insertion', 4, 'Carol'],
    ])
    expect(entries[2]!.change).toBe('bold: none → yes')
    expect(entries[3]!.change).toBe('align: left → center')
    expect(entries[1]!.text).toBe('gone words')
  })

  it('read_revisions shows ids so the model can target them', async () => {
    const editor = createEditor()
    const r = await executeTool(editor, { id: 't', name: 'read_revisions', input: {} }, NUM_IDS)
    expect(r.output).toContain('- r2 | block 1 | deleted by Carol on 2026-03-01: "gone words"')
    expect(r.output).toContain('(align: left → center)')
    expect(buildRevisionsContext(editor)).toContain('ids are positional')
  })

  it('an empty selection names what is pending', () => {
    const r = selectRevisions(createEditor().state.doc, { author: 'Nobody' })
    expect(r).toEqual({
      error: expect.stringContaining(
        'authors Carol (3), Dave (2); types insertion (2), deletion (1), formatting (2)',
      ),
    })
    const unknown = selectRevisions(createEditor().state.doc, { ids: ['r9'] })
    expect(unknown).toEqual({
      error: expect.stringContaining('unknown revision id(s) r9; valid ids are r1–r5'),
    })
  })
})

describe('accept_changes / reject_changes tools', () => {
  it('accepts everything', async () => {
    const editor = createEditor()
    const r = await executeTool(
      editor,
      { id: 't', name: 'accept_changes', input: { all: true } },
      NUM_IDS,
    )
    expect(r.isError).not.toBe(true)
    expect(r.mutated).toBe(true)
    expect(r.output).toContain('Accepted 5 tracked change(s)')
    expect(r.output).toContain('No tracked changes remain')
    expect(collectRevisions(editor.state.doc)).toEqual([])
    expect(bodyText(editor)).toBe(
      'Untouched.|Base added words and .|Now bold|Centered now.|Whole new paragraph.',
    )
    expect(editor.state.doc.child(3).attrs.align).toBe('center')
  })

  it('rejects by author, leaving the other author pending', async () => {
    const editor = createEditor()
    const r = await executeTool(
      editor,
      { id: 't', name: 'reject_changes', input: { author: 'carol' } },
      NUM_IDS,
    )
    expect(r.isError).not.toBe(true)
    expect(r.output).toContain('Rejected 3 tracked change(s) (r1, r2, r5)')
    expect(r.output).toContain('Still 2 pending: authors Dave (2)')
    expect(bodyText(editor)).toBe('Untouched.|Base  and gone words.|Now bold|Centered now.')
    const left = listRevisionEntries(editor.state.doc)
    expect(left.map((e) => [e.id, e.author])).toEqual([
      ['r1', 'Dave'],
      ['r2', 'Dave'],
    ])
  })

  it('rejects formatting by type and restores the old format', async () => {
    const editor = createEditor()
    const r = await executeTool(
      editor,
      { id: 't', name: 'reject_changes', input: { type: 'formatting' } },
      NUM_IDS,
    )
    expect(r.isError).not.toBe(true)
    expect(editor.state.doc.child(3).attrs.align).toBe('left')
    expect(editor.state.doc.child(3).attrs.pPrChange).toBeNull()
    const boldLeft = editor.state.doc.child(2).firstChild!.marks.some((m) => m.type.name === 'bold')
    expect(boldLeft).toBe(false)
    expect(listRevisionEntries(editor.state.doc).map((e) => e.type)).toEqual([
      'insertion',
      'deletion',
      'insertion',
    ])
  })

  it('accepts by id and by block, in one transaction each', async () => {
    const editor = createEditor()
    const byId = await executeTool(
      editor,
      { id: 't', name: 'accept_changes', input: { ids: ['r2'] } },
      NUM_IDS,
    )
    expect(byId.output).toContain('Accepted 1 tracked change(s) (r2)')
    expect(bodyText(editor)).toContain('Base added words and .')
    const byBlock = await executeTool(
      editor,
      { id: 't', name: 'accept_changes', input: { blockIndex: 1 } },
      NUM_IDS,
    )
    expect(byBlock.output).toContain('Accepted 1 tracked change(s)')
    expect(editor.state.doc.child(1).firstChild!.marks.some((m) => m.type.name === 'ins')).toBe(
      false,
    )
    const before = await executeTool(
      editor,
      { id: 't', name: 'accept_changes', input: { before: '2026-04-01' } },
      NUM_IDS,
    )
    // r1 is now the block insertion (the two text changes above are gone)
    expect(before.output).toContain('Accepted 1 tracked change(s) (r3)')
    expect(collectRevisions(editor.state.doc).map((r) => r.author)).toEqual(['Dave', 'Dave'])
  })

  it('selects author-less changes as "unknown"', async () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          para({}, text('Stays.')),
          para({}, text('no author', { type: 'ins', attrs: { author: '' } })),
        ],
      },
    })
    liveEditors.push(editor)
    expect(listRevisionEntries(editor.state.doc)[0]!.author).toBe('unknown')
    const r = await executeTool(
      editor,
      { id: 't', name: 'reject_changes', input: { author: 'unknown' } },
      NUM_IDS,
    )
    expect(r.isError).not.toBe(true)
    expect(bodyText(editor)).toBe('Stays.')
  })

  it('refuses a missing or malformed selector without touching the document', async () => {
    const editor = createEditor()
    const snapshot = editor.state.doc.toJSON()
    const empty = await executeTool(editor, { id: 't', name: 'accept_changes', input: {} }, NUM_IDS)
    expect(empty.isError).toBe(true)
    expect(empty.output).toContain('give a selector')
    const badType = await executeTool(
      editor,
      { id: 't', name: 'reject_changes', input: { type: 'typo' } },
      NUM_IDS,
    )
    expect(badType.isError).toBe(true)
    expect(badType.output).toContain('type must be one of insertion, deletion, formatting, move')
    const nobody = await executeTool(
      editor,
      { id: 't', name: 'reject_changes', input: { author: 'Nobody', type: 'deletion' } },
      NUM_IDS,
    )
    expect(nobody.isError).toBe(true)
    expect(nobody.output).toContain('no tracked change matches author="Nobody", type="deletion"')
    expect(editor.state.doc.toJSON()).toEqual(snapshot)
  })
})
