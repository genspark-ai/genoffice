import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { renumber } from '../src/renderer/ai/note-ops'
import {
  assignNoteNumbers,
  footnotesRestartPerPage,
  setFootnotePages,
  setNoteProps,
  type NoteCountRule,
  type NoteRefSite,
} from '../src/renderer/note-format'
import { NOTE_SETTLE_CAP, NotePageSettle } from '../src/renderer/note-page-settle'

const rule = (restart: NoteCountRule['restart'], start = 1, fmt = 'decimal'): NoteCountRule => ({
  start,
  restart,
  fmt,
})
const nums = (refs: NoteRefSite[], rules: NoteCountRule[]) =>
  [...assignNoteNumbers(refs, (s) => rules[s]!).values()].map((m) => m.num)
const site = (id: string, page?: number, section = 0): NoteRefSite => ({
  id,
  section,
  ...(page === undefined ? {} : { page }),
})

describe('note numbers by page', () => {
  it('starts over at the first mark of each page', () => {
    const refs = [site('a', 0), site('b', 0), site('c', 1), site('d', 3), site('e', 3)]
    expect(nums(refs, [rule('eachPage')])).toEqual([1, 2, 1, 1, 2])
  })

  it('opens every page at the start value of the section', () => {
    const refs = [site('a', 0), site('b', 0), site('c', 1)]
    expect(nums(refs, [rule('eachPage', 4)])).toEqual([4, 5, 4])
  })

  it('counts on without pages, and puts a mark of unknown page on the page before it', () => {
    expect(nums([site('a'), site('b'), site('c')], [rule('eachPage')])).toEqual([1, 2, 3])
    expect(nums([site('a', 0), site('b', 1), site('new')], [rule('eachPage')])).toEqual([1, 1, 2])
  })

  it('a note referenced twice keeps the number of its first mark', () => {
    const refs = [site('a', 0), site('b', 0), site('a', 1), site('c', 1)]
    const shown = assignNoteNumbers(refs, () => rule('eachPage'))
    expect([shown.get('a')!.num, shown.get('b')!.num, shown.get('c')!.num]).toEqual([1, 2, 1])
  })

  it('pages do not restart a continuous or per-section count', () => {
    const refs = [site('a', 0), site('b', 1), site('c', 1, 1), site('d', 2, 1)]
    expect(nums(refs, [rule('continuous'), rule('continuous')])).toEqual([1, 2, 3, 4])
    expect(nums(refs, [rule('continuous'), rule('eachSect')])).toEqual([1, 2, 1, 2])
  })

  it('a section without marks still starts a per-section count over', () => {
    const refs = [site('a', 0, 0), site('b', 0, 2)]
    expect(nums(refs, [rule('continuous'), rule('eachSect'), rule('continuous')])).toEqual([1, 1])
  })

  it('each section counts its pages by its own rule', () => {
    // a per-page section, then a continuous one that begins on the same page
    const refs = [site('a', 0, 0), site('b', 1, 0), site('c', 1, 1), site('d', 2, 1)]
    expect(nums(refs, [rule('eachPage'), rule('continuous', 1, 'lowerRoman')])).toEqual([
      1, 1, 2, 3,
    ])
    const fmts = [
      ...assignNoteNumbers(
        refs,
        (s) => [rule('eachPage'), rule('eachPage', 1, 'lowerRoman')][s]!,
      ).values(),
    ]
    expect(fmts.map((m) => m.fmt)).toEqual(['decimal', 'decimal', 'lowerRoman', 'lowerRoman'])
  })

  it('a per-page section without notes does not start a later continuous section over', () => {
    const refs = [site('a', 0, 0), site('b', 0, 2)]
    expect(nums(refs, [rule('continuous'), rule('eachPage'), rule('continuous')])).toEqual([1, 2])
  })

  it('is asked for only when a section restarts its footnotes per page', () => {
    const none = { footnote: {}, endnote: {} }
    expect(footnotesRestartPerPage([none])).toBe(false)
    expect(
      footnotesRestartPerPage([none, { footnote: { numRestart: 'eachSect' }, endnote: {} }]),
    ).toBe(false)
    // endnotes have no per-page restart
    expect(footnotesRestartPerPage([{ footnote: {}, endnote: { numRestart: 'eachPage' } }])).toBe(
      false,
    )
    expect(
      footnotesRestartPerPage([none, { footnote: { numRestart: 'eachPage' }, endnote: {} }]),
    ).toBe(true)
  })
})

describe('renumbering the marks of a document by page', () => {
  afterEach(() => {
    setNoteProps({})
    setFootnotePages(new Map())
  })

  const refNode = (kind: string, id: string, extra: Record<string, unknown> = {}) => ({
    type: 'docNoteRef',
    attrs: { kind, id, num: 9, ...extra },
  })
  const makeEditor = () =>
    new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: ['1', '2', '3', '4'].map((id, i) => ({
          type: 'docParagraph',
          attrs: { docxIndex: i },
          content: [
            { type: 'text', text: `para ${id}` },
            refNode('footnote', id),
            refNode('endnote', id),
            ...(id === '2' ? [refNode('footnote', 'star', { mark: '*' })] : []),
          ],
        })),
      },
    })
  const shown = (editor: Editor, kind: string) => {
    const out: number[] = []
    editor.state.doc.descendants((node) => {
      if (node.type.name === 'docNoteRef' && node.attrs.kind === kind && node.attrs.mark == null)
        out.push(Number(node.attrs.num))
      return true
    })
    return out
  }
  const run = (editor: Editor, kind: 'footnote' | 'endnote') => {
    const tr = editor.state.tr
    renumber(tr, kind)
    editor.view.dispatch(tr)
  }

  it('footnotes follow the pages of the last layout; a custom mark takes no number', () => {
    const editor = makeEditor()
    setNoteProps({ footnote: { numRestart: 'eachPage', numStart: 3 } })
    setFootnotePages(
      new Map([
        ['1', 0],
        ['2', 0],
        ['star', 0],
        ['3', 1],
        ['4', 1],
      ]),
    )
    run(editor, 'footnote')
    expect(shown(editor, 'footnote')).toEqual([3, 4, 3, 4])
    editor.destroy()
  })

  it('before any layout the marks count on, as the file opens', () => {
    const editor = makeEditor()
    setNoteProps({ footnote: { numRestart: 'eachPage' } })
    run(editor, 'footnote')
    expect(shown(editor, 'footnote')).toEqual([1, 2, 3, 4])
    editor.destroy()
  })

  it('endnotes ignore a per-page restart and the footnote pages', () => {
    const editor = makeEditor()
    setNoteProps({ endnote: { numRestart: 'eachPage' } })
    setFootnotePages(
      new Map([
        ['1', 0],
        ['2', 1],
        ['3', 2],
        ['4', 3],
      ]),
    )
    run(editor, 'endnote')
    expect(shown(editor, 'endnote')).toEqual([1, 2, 3, 4])
    editor.destroy()
  })
})

describe('the stopping rule of the renumber loop', () => {
  const doc = {}

  it('does nothing while the layout asks for the numbers it was made with', () => {
    const settle = new NotePageSettle()
    expect(settle.step(doc, 'A', 'A')).toBe(false)
    expect(settle.stats).toEqual({ maxRounds: 0, stop: null })
  })

  it('renumbers until the numbers hold and reports the rounds it took', () => {
    const settle = new NotePageSettle()
    expect(settle.step(doc, 'A', 'B')).toBe(true)
    const own = {}
    settle.carry(own)
    expect(settle.step(own, 'B', 'C')).toBe(true)
    const own2 = {}
    settle.carry(own2)
    expect(settle.step(own2, 'C', 'C')).toBe(false)
    expect(settle.stats).toEqual({ maxRounds: 2, stop: null })
  })

  it('an oscillation stops on the numbers the page is drawn with and stays there', () => {
    const settle = new NotePageSettle()
    let input: object = doc
    expect(settle.step(input, 'A', 'B')).toBe(true)
    settle.carry((input = {}))
    // B asks for A again: the page keeps B
    expect(settle.step(input, 'B', 'A')).toBe(false)
    expect(settle.stats.stop).toBe('cycle')
    for (let i = 0; i < 20; i++) expect(settle.step(input, 'B', 'A')).toBe(false)
    expect(settle.stats.maxRounds).toBe(1)
  })

  it('a longer cycle is caught the same way', () => {
    const settle = new NotePageSettle()
    let input: object = doc
    const asks: Record<string, string> = { A: 'B', B: 'C', C: 'B' }
    let current = 'A'
    let rounds = 0
    while (settle.step(input, current, asks[current]!)) {
      current = asks[current]!
      settle.carry((input = {}))
      rounds++
    }
    expect([current, rounds, settle.stats.stop]).toEqual(['C', 2, 'cycle'])
  })

  it('numbers that never repeat stop at the cap', () => {
    const settle = new NotePageSettle()
    let input: object = doc
    let n = 0
    while (settle.step(input, `s${n}`, `s${n + 1}`)) {
      n++
      settle.carry((input = {}))
      expect(n).toBeLessThanOrEqual(NOTE_SETTLE_CAP)
    }
    expect(n).toBe(NOTE_SETTLE_CAP)
    expect(settle.stats).toEqual({ maxRounds: NOTE_SETTLE_CAP, stop: 'cap' })
    expect(settle.step(input, `s${n}`, `s${n + 1}`)).toBe(false)
  })

  it('an edit opens a new run after a stop', () => {
    const settle = new NotePageSettle()
    let input: object = doc
    settle.step(input, 'A', 'B')
    settle.carry((input = {}))
    expect(settle.step(input, 'B', 'A')).toBe(false)
    // the user typed: another document, the loop may try again, A included
    expect(settle.step({}, 'B', 'A')).toBe(true)
  })

  it('the same start always ends on the same numbers', () => {
    const asks: Record<string, string> = { A: 'B', B: 'C', C: 'D', D: 'C' }
    const end = () => {
      const settle = new NotePageSettle()
      let input: object = {}
      let current = 'A'
      while (settle.step(input, current, asks[current]!)) {
        current = asks[current]!
        settle.carry((input = {}))
      }
      return current
    }
    expect([end(), end(), end()]).toEqual(['D', 'D', 'D'])
  })
})
