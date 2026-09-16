import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import type { Node as ProseNode } from '@tiptap/pm/model'
import type { Run } from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { inlineToRuns, runsToInline, type PmNode } from '../src/renderer/editor/convert'
import {
  buildNotesContext,
  noteInsertPos,
  protectedNoteMarkBlock,
} from '../src/renderer/ai/note-ops'

function createEditor(content: PmNode[]): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
}

describe('footnote / endnote reference conversion', () => {
  it('maps Run.noteRef to a docNoteRef atom and back', () => {
    const runs: Run[] = [
      { text: 'body' },
      { text: '2', noteRef: { kind: 'footnote', id: '3' } },
      { text: 'continued' },
    ]
    const inline = runsToInline(runs)
    expect(inline).toEqual([
      { type: 'text', text: 'body' },
      { type: 'docNoteRef', attrs: { kind: 'footnote', id: '3', num: 2 } },
      { type: 'text', text: 'continued' },
    ])
    expect(inlineToRuns(inline)).toEqual(runs)
  })

  it('maps Run.xeTerm to a trailing docXeMark atom and back', () => {
    const runs: Run[] = [{ text: 'artificial intelligence', xeTerm: 'artificial intelligence' }]
    const inline = runsToInline(runs)
    expect(inline).toEqual([
      { type: 'text', text: 'artificial intelligence' },
      { type: 'docXeMark', attrs: { term: 'artificial intelligence' } },
    ])
    // splits back into a plain text run + an empty XE-carrier run
    expect(inlineToRuns(inline)).toEqual([
      { text: 'artificial intelligence' },
      { text: '', xeTerm: 'artificial intelligence' },
    ])
  })

  it('renders both atoms in the editor schema without dropping them', () => {
    const editor = createEditor([
      {
        type: 'docParagraph',
        attrs: { docxIndex: 0 },
        content: [
          { type: 'text', text: 'a' },
          { type: 'docNoteRef', attrs: { kind: 'endnote', id: '5', num: 1 } },
          { type: 'docXeMark', attrs: { term: 'glossary term' } },
        ],
      },
    ])
    const para = (editor.getJSON() as PmNode).content![0]
    const types = (para.content ?? []).map((n) => n.type)
    expect(types).toEqual(['text', 'docNoteRef', 'docXeMark'])
    const html = editor.getHTML()
    expect(html).toContain('data-note-ref="5"')
    expect(html).toContain('data-xe-term="glossary term"')
    editor.destroy()
  })
})

describe('noteInsertPos', () => {
  it('skips afterText matches that end inside a field result', () => {
    const editor = createEditor([
      {
        type: 'docParagraph',
        attrs: { docxIndex: 0 },
        content: [
          { type: 'text', text: 'See ' },
          { type: 'text', text: 'Table 12', marks: [{ type: 'refField', attrs: { name: 'tbl' } }] },
          { type: 'text', text: ' below' },
        ],
      },
    ])
    const doc = editor.state.doc
    expect(noteInsertPos(doc, 0, 'Table 1')).toEqual({
      error: expect.stringContaining('ends inside the field { REF tbl }'),
    })
    expect(noteInsertPos(doc, 0, 'Table 12')).toEqual({ pos: 1 + 'See Table 12'.length })
    expect(noteInsertPos(doc, 0, 'See')).toEqual({ pos: 1 + 'See'.length })
    editor.destroy()
  })
})

describe('buildNotesContext', () => {
  it('locates marks that live only in protected block XML', () => {
    const editor = createEditor([
      {
        type: 'docParagraph',
        attrs: { docxIndex: 0 },
        content: [
          { type: 'text', text: 'a' },
          { type: 'docNoteRef', attrs: { kind: 'footnote', id: '1', num: 1 } },
        ],
      },
      {
        type: 'docProtected',
        attrs: {
          docxIndex: 1,
          blockType: 'passthrough',
          label: 'p',
          previewText: '',
          genXml: '<w:p><w:r><w:footnoteReference w:id="2"/></w:r></w:p>',
        },
      },
    ])
    const doc = editor.state.doc
    const footnotes = [
      { id: '1', text: 'first' },
      { id: '2', text: 'second' },
      { id: '3', text: 'orphan' },
    ]
    const xmlOf = (block: ProseNode) => String(block.attrs.genXml ?? '')
    const out = buildNotesContext(doc, footnotes, [], (kind, id) =>
      protectedNoteMarkBlock(doc, xmlOf, kind, id),
    )
    expect(out).toContain('id 1 (mark 1 in block 0)')
    expect(out).toContain('id 2 (mark in protected block 1, cannot be deleted here)')
    expect(out).toContain('id 3 (no reference mark in the text)')
  })
})
