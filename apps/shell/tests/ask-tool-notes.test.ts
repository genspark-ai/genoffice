import { describe, expect, it } from 'vitest'

import type { AgentToolCall } from '@genoffice/agent-core'

import { appendToolNote, toolNoteDetail, type ToolNote } from '../src/renderer/src/ask-tool-notes'

function call(name: string, input: Record<string, unknown>): AgentToolCall {
  return { id: `c-${name}`, name, input }
}

const note = (over: Partial<ToolNote> = {}): ToolNote => ({
  label: 'Building a file',
  kind: 'cli',
  count: 1,
  ...over,
})

describe('toolNoteDetail', () => {
  it('names the CLI command and its --out file', () => {
    const detail = toolNoteDetail(
      call('run_genoffice', {
        args: ['slides', 'create', 'spec.json', '--out', '/out/brief.pptx'],
      }),
    )
    expect(detail).toBe('slides create · brief.pptx')
  })

  it('falls back to the file argument of an in-place edit', () => {
    const detail = toolNoteDetail(
      call('run_genoffice', { args: ['docs', 'apply', '/out/report.docx', 'ops.json'] }),
    )
    expect(detail).toBe('docs apply · report.docx')
  })

  it('marks a --help probe and a guide load', () => {
    expect(toolNoteDetail(call('run_genoffice', { args: ['sheet', 'read', '--help'] }))).toBe(
      'sheet read · --help',
    )
    expect(toolNoteDetail(call('load_guide', { domain: 'slides' }))).toBe('guide slides')
  })

  it('resolves a knowledge-base read to the file name', () => {
    const files = ['/kb/a/lease.docx', '/kb/b/notes.md']
    expect(toolNoteDetail(call('read_knowledge_file', { index: 1 }), (i) => files[i])).toBe(
      'notes.md',
    )
    expect(toolNoteDetail(call('read_knowledge_file', { index: '0' }), (i) => files[i])).toBe(
      'lease.docx',
    )
    expect(
      toolNoteDetail(call('read_knowledge_file', { index: 7 }), (i) => files[i]),
    ).toBeUndefined()
  })

  it('leaves a search without detail', () => {
    expect(toolNoteDetail(call('search_knowledge_base', { query: 'rent' }))).toBeUndefined()
  })
})

describe('appendToolNote', () => {
  it('folds consecutive calls on the same target into a count', () => {
    let notes: ToolNote[] = []
    notes = appendToolNote(notes, note({ detail: 'slides apply · deck.pptx' }))
    notes = appendToolNote(notes, note({ detail: 'slides apply · deck.pptx' }))
    notes = appendToolNote(notes, note({ detail: 'slides apply · deck.pptx' }))
    expect(notes).toHaveLength(1)
    expect(notes[0]!.count).toBe(3)
  })

  it('keeps a different target or kind as its own chip', () => {
    let notes: ToolNote[] = []
    notes = appendToolNote(notes, note({ detail: 'slides create · deck.pptx' }))
    notes = appendToolNote(notes, note({ detail: 'slides apply · deck.pptx' }))
    notes = appendToolNote(
      notes,
      note({ kind: 'read', label: 'Reading a file', detail: 'deck.pptx' }),
    )
    notes = appendToolNote(notes, note({ detail: 'slides apply · deck.pptx' }))
    expect(notes.map((n) => [n.detail, n.count])).toEqual([
      ['slides create · deck.pptx', 1],
      ['slides apply · deck.pptx', 1],
      ['deck.pptx', 1],
      ['slides apply · deck.pptx', 1],
    ])
  })
})
