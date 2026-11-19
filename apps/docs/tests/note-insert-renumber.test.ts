import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { nextNoteId, parseNotesXml, type NoteInfo } from '@genoffice/docx-engine'
import { buildNotesXml } from '../../../packages/docx-engine/src/notes'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  convertNotes,
  deleteNote,
  submitNote,
  type ReviewContext,
} from '../src/renderer/review-actions'
import { noteBlockSections, renumber } from '../src/renderer/ai/note-ops'
import { setNoteProps } from '../src/renderer/note-format'

/** the in-text reference marks of one kind, in document order */
function marks(editor: Editor, kind: 'footnote' | 'endnote'): Array<{ id: string; num: number }> {
  const out: Array<{ id: string; num: number }> = []
  editor.state.doc.descendants((node) => {
    if (node.type.name !== 'docNoteRef' || node.attrs.kind !== kind) return true
    out.push({ id: String(node.attrs.id), num: Number(node.attrs.num) })
    return false
  })
  return out
}

function makeEditor(): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { docxIndex: 0 },
          content: [
            { type: 'text', text: 'alpha' },
            { type: 'docNoteRef', attrs: { kind: 'footnote', id: '3', num: 1 } },
          ],
        },
        {
          type: 'docParagraph',
          attrs: { docxIndex: 1 },
          content: [
            { type: 'text', text: 'beta' },
            { type: 'docNoteRef', attrs: { kind: 'footnote', id: '5', num: 2 } },
          ],
        },
      ],
    },
  })
}

function makeCtx(editor: Editor, footnotes: NoteInfo[]): { ctx: ReviewContext; state: NoteInfo[] } {
  const state = [...footnotes]
  const ctx = {
    editor,
    notePrompt: { kind: 'footnote' },
    setFootnotes: (value: unknown) => {
      const next =
        typeof value === 'function'
          ? (value as (prev: NoteInfo[]) => NoteInfo[])(state)
          : (value as NoteInfo[])
      state.splice(0, state.length, ...next)
    },
    setNotesDirty: () => undefined,
    footnotes,
  } as unknown as ReviewContext
  return { ctx, state }
}

describe('inserting a note above the existing ones', () => {
  it('renumbers the marks already in the body into document order', () => {
    const editor = makeEditor()
    // the caret sits before both existing marks, so the new one comes first
    editor.commands.setTextSelection(1)
    const { ctx, state } = makeCtx(editor, [
      { id: '3', text: 'first note' },
      { id: '5', text: 'second note' },
    ])

    submitNote(ctx, 'inserted note')

    const newId = nextNoteId([
      { id: '3', text: '' },
      { id: '5', text: '' },
    ])
    // document order is the new mark, then the two that were already there
    expect(marks(editor, 'footnote')).toEqual([
      { id: newId, num: 1 },
      { id: '3', num: 2 },
      { id: '5', num: 3 },
    ])
    // the note list itself still appends
    expect(state.map((n) => n.id)).toEqual(['3', '5', newId])
    editor.destroy()
  })

  it('leaves the marks of the other kind alone', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            attrs: { docxIndex: 0 },
            content: [
              { type: 'text', text: 'alpha' },
              { type: 'docNoteRef', attrs: { kind: 'endnote', id: '9', num: 1 } },
              { type: 'docNoteRef', attrs: { kind: 'footnote', id: '3', num: 1 } },
            ],
          },
        ],
      },
    })
    editor.commands.setTextSelection(1)
    const { ctx } = makeCtx(editor, [{ id: '3', text: 'a footnote' }])

    submitNote(ctx, 'inserted note')

    expect(marks(editor, 'endnote')).toEqual([{ id: '9', num: 1 }])
    expect(marks(editor, 'footnote').map((m) => m.num)).toEqual([1, 2])
    editor.destroy()
  })
})

describe('deleting a note whose list order differs from document order', () => {
  it('renumbers the surviving marks by document order, not by list index', () => {
    // the marks read 3, 5, 7 down the body, but the list was created 7, 3, 5
    // (the caret moved between insertions), so the two orders disagree
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            attrs: { docxIndex: 0 },
            content: [
              { type: 'text', text: 'alpha' },
              { type: 'docNoteRef', attrs: { kind: 'footnote', id: '3', num: 1 } },
            ],
          },
          {
            type: 'docParagraph',
            attrs: { docxIndex: 1 },
            content: [
              { type: 'text', text: 'beta' },
              { type: 'docNoteRef', attrs: { kind: 'footnote', id: '5', num: 2 } },
            ],
          },
          {
            type: 'docParagraph',
            attrs: { docxIndex: 2 },
            content: [
              { type: 'text', text: 'gamma' },
              { type: 'docNoteRef', attrs: { kind: 'footnote', id: '7', num: 3 } },
            ],
          },
        ],
      },
    })
    const { ctx, state } = makeCtx(editor, [
      { id: '7', text: 'third note' },
      { id: '3', text: 'first note' },
      { id: '5', text: 'second note' },
    ])

    deleteNote(ctx, 'footnote', '5')

    // '3' still comes before '7' in the body, so it keeps number 1; numbering
    // off the surviving list (7, 3) would hand it 2 and read 1, 2 descending
    expect(marks(editor, 'footnote')).toEqual([
      { id: '3', num: 1 },
      { id: '7', num: 2 },
    ])
    // the list drops the note, and its own order is left alone
    expect(state.map((n) => n.id)).toEqual(['7', '3'])
    editor.destroy()
  })
})

describe('numbering under the document note properties', () => {
  afterEach(() => setNoteProps({}, []))
  const sections = (...ends: number[]) =>
    ends.map((lastBlockIndex) => ({ lastBlockIndex, sectPrXml: '' }))

  const renumbered = (editor: Editor) => {
    const tr = editor.state.tr
    renumber(tr, 'footnote')
    editor.view.dispatch(tr)
    return marks(editor, 'footnote').map((m) => m.num)
  }

  it('counts from the start number and starts over after each section', () => {
    const editor = makeEditor()
    setNoteProps({ footnote: { numStart: 4 } }, sections(0, 1))
    expect(renumbered(editor)).toEqual([4, 5])
    // block 0 closes the first section
    setNoteProps({ footnote: { numStart: 4, numRestart: 'eachSect' } }, sections(0, 1))
    expect(renumbered(editor)).toEqual([4, 4])
    editor.destroy()
  })

  it('a number format change redraws the marks without renumbering them', () => {
    const editor = makeEditor()
    setNoteProps({ footnote: { numFmt: 'lowerRoman' } }, sections(1))
    expect(renumbered(editor)).toEqual([1, 2])
    expect(editor.view.dom.querySelectorAll('.doc-note-ref')[1]?.textContent).toBe('ii')
    editor.destroy()
  })

  it('each section counts and draws by its own properties', () => {
    const editor = makeEditor()
    const texts = () =>
      [...editor.view.dom.querySelectorAll('.doc-note-ref')].map((el) => el.textContent)
    // block 0 closes the first section; the second starts over at 3 in capitals
    setNoteProps(
      [
        { footnote: { numFmt: 'lowerLetter' }, endnote: {} },
        { footnote: { numFmt: 'upperRoman', numStart: 3, numRestart: 'eachSect' }, endnote: {} },
      ],
      sections(0, 1),
    )
    expect(renumbered(editor)).toEqual([1, 3])
    expect(texts()).toEqual(['a', 'III'])
    // a section that counts on adds the running count to its own start number
    setNoteProps(
      [
        { footnote: {}, endnote: {} },
        { footnote: { numStart: 3 }, endnote: {} },
      ],
      sections(0, 1),
    )
    expect(renumbered(editor)).toEqual([1, 4])
    expect(texts()).toEqual(['1', '4'])
    editor.destroy()
  })

  it('a note in a paragraph typed after a saved break is numbered in the section the break opens', () => {
    const editor = makeEditor()
    // block 0 closes the first section; the new paragraph has no place in the file yet
    editor.commands.insertContentAt(editor.state.doc.child(0).nodeSize, {
      type: 'docParagraph',
      attrs: { docxIndex: null },
      content: [
        { type: 'text', text: 'typed' },
        { type: 'docNoteRef', attrs: { kind: 'footnote', id: '9', num: 1 } },
      ],
    })
    const list = sections(0, 1)
    setNoteProps(
      [
        { footnote: {}, endnote: {} },
        { footnote: { numFmt: 'upperRoman', numStart: 7, numRestart: 'eachSect' }, endnote: {} },
      ],
      list,
    )
    expect(noteBlockSections(editor.state.doc, list)).toEqual([0, 1, 1])
    expect(renumbered(editor)).toEqual([1, 7, 8])
    expect(
      [...editor.view.dom.querySelectorAll('.doc-note-ref')].map((el) => el.textContent),
    ).toEqual(['1', 'VII', 'VIII'])
    // behind a break made in this session the break paragraph itself moves the section on
    expect(
      noteBlockSections(editor.state.doc, [{ ...list[0]!, pendingBreak: true }, list[1]!]),
    ).toEqual([0, 0, 1])
    editor.destroy()
  })

  it('a custom mark shows its text and takes no number', () => {
    const editor = makeEditor()
    editor.commands.setTextSelection(1)
    const { ctx, state } = makeCtx(editor, [
      { id: '3', text: 'first note' },
      { id: '5', text: 'second note' },
    ])
    ctx.notePrompt = { kind: 'footnote', customMark: '\u2020' }
    submitNote(ctx, 'marked')
    expect(marks(editor, 'footnote').map((m) => m.num)).toEqual([3, 1, 2])
    expect(editor.view.dom.querySelector('.doc-note-ref')?.textContent).toBe('\u2020')
    expect(state[2]).toMatchObject({ text: '\u2020 marked', noRefMark: true, customMark: '\u2020' })
    editor.destroy()
  })
})

describe('converting notes', () => {
  it('moves every footnote to the endnotes, marks included, and back again', () => {
    const editor = makeEditor()
    const lists = {
      footnote: [
        { id: '3', text: 'first note' },
        { id: '5', text: 'second note' },
      ] as NoteInfo[],
      endnote: [] as NoteInfo[],
    }
    const ctx = {
      editor,
      get footnotes() {
        return lists.footnote
      },
      get endnotes() {
        return lists.endnote
      },
      setFootnotes: (value: NoteInfo[]) => (lists.footnote = value),
      setEndnotes: (value: NoteInfo[]) => (lists.endnote = value),
      setNotesDirty: () => undefined,
    } as unknown as ReviewContext

    expect(convertNotes(ctx, 'toEndnotes')).toBe(2)
    expect(marks(editor, 'footnote')).toEqual([])
    expect(marks(editor, 'endnote')).toEqual([
      { id: '1', num: 1 },
      { id: '2', num: 2 },
    ])
    expect(lists.footnote).toEqual([])
    expect(lists.endnote).toEqual([
      { id: '1', text: 'first note', origin: { kind: 'footnote', id: '3' } },
      { id: '2', text: 'second note', origin: { kind: 'footnote', id: '5' } },
    ])

    expect(convertNotes(ctx, 'toFootnotes')).toBe(2)
    expect(marks(editor, 'footnote').map((m) => m.num)).toEqual([1, 2])
    // the bytes still come from where the file has them
    expect(lists.footnote.map((n) => n.origin)).toEqual([
      { kind: 'footnote', id: '3' },
      { kind: 'footnote', id: '5' },
    ])
    expect(convertNotes(ctx, 'toFootnotes')).toBe(0)
    editor.destroy()
  })

  it('a note added in this session never takes the bytes of a deleted note whose id it reuses', () => {
    const editor = makeEditor()
    const lists = {
      footnote: [
        { id: '3', text: 'kept note' },
        { id: '5', text: 'deleted note' },
      ] as NoteInfo[],
      endnote: [] as NoteInfo[],
    }
    const ctx = {
      editor,
      notePrompt: { kind: 'footnote' },
      get footnotes() {
        return lists.footnote
      },
      get endnotes() {
        return lists.endnote
      },
      setFootnotes: (value: NoteInfo[]) => (lists.footnote = value),
      setEndnotes: (value: NoteInfo[]) => (lists.endnote = value),
      setNotesDirty: () => undefined,
    } as unknown as ReviewContext
    // with the last two notes gone, the next id handed out is one the file still has
    deleteNote(ctx, 'footnote', '5')
    deleteNote(ctx, 'footnote', '3')
    submitNote(ctx, 'new note')
    expect(lists.footnote).toEqual([{ id: '1', text: 'new note', created: true }])

    expect(convertNotes(ctx, 'toEndnotes')).toBe(1)
    expect(lists.endnote).toEqual([{ id: '1', text: 'new note', created: true }])

    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    const original =
      `<w:footnotes ${W}><w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r>` +
      '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve"> long gone</w:t></w:r></w:p></w:footnote></w:footnotes>'
    const other = { xml: original, styleIds: new Set<string>() }
    const saved = buildNotesXml('endnote', lists.endnote, null, other)
    expect(parseNotesXml(saved, 'endnote').map((n) => n.text)).toEqual(['new note'])
    expect(saved).not.toContain('<w:b/>')
    // the same holds where it was made: a fresh entry, not the old one's bytes with new text
    const inPlace = buildNotesXml(
      'footnote',
      [{ id: '1', text: 'new note', created: true }],
      original,
    )
    expect(inPlace).not.toContain('<w:b/>')
    editor.destroy()
  })
})
