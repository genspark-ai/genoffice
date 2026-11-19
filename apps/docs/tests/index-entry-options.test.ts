/** Mark Index Entry: the XE field's switches, Mark All, and the index that reads them. */
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  indexEntryOptions,
  indexEntrySwitches,
  indexEntryTerm,
  markIndexEntries,
  occurrenceEnds,
  type IndexEntry,
} from '../src/renderer/editor/index-entry'
import { indexLines, insertIndex } from '../src/renderer/commands/reference-commands'

const entry = (over: Partial<IndexEntry>): IndexEntry => ({
  main: 'Fruit',
  sub: '',
  target: 'page',
  cross: '',
  bookmark: '',
  bold: false,
  italic: false,
  ...over,
})

describe('index entry switches', () => {
  it('writes a subentry after a colon and the options as XE switches', () => {
    expect(indexEntryTerm(entry({ sub: 'Apples' }))).toBe('Fruit:Apples')
    expect(indexEntrySwitches(entry({}))).toBeNull()
    expect(indexEntrySwitches(entry({ target: 'cross', cross: 'See "Pomes"', bold: true }))).toBe(
      '\\t "See Pomes" \\b',
    )
    expect(indexEntrySwitches(entry({ target: 'range', bookmark: 'Orchard', italic: true }))).toBe(
      '\\r Orchard \\i',
    )
  })

  it('reads them back, and a switch letter inside a quoted text is not a switch', () => {
    expect(indexEntryOptions('\\t "See \\b notes" \\i')).toEqual({
      cross: 'See \\b notes',
      bold: false,
      italic: true,
    })
    expect(indexEntryOptions('\\r Orchard \\b')).toEqual({
      bookmark: 'Orchard',
      bold: true,
      italic: false,
    })
    expect(indexEntryOptions(null)).toEqual({ bold: false, italic: false })
  })
})

describe('index lines', () => {
  it('a cross-reference stands where the page would, after the pages when there are some', () => {
    const lines = indexLines(
      [
        { term: 'Apples', cross: 'See Fruit' },
        { term: 'Fruit', page: 2 },
        { term: 'Fruit', cross: 'See also Trees' },
      ],
      false,
    )
    expect(lines.map((l) => [l.text, l.pageNo ?? '', !!l.noPage])).toEqual([
      ['Apples. See Fruit', '', true],
      ['Fruit', '2. See also Trees', false],
    ])
  })

  it('a bold or italic page is a run of its own; plain pages stay one text', () => {
    const [line] = indexLines(
      [
        { term: 'Fruit', page: 5, bold: true },
        { term: 'Fruit', page: 2 },
        { term: 'Fruit', page: 9, italic: true },
      ],
      false,
    )
    expect(line!.pageNo).toBe('2, 5, 9')
    expect(line!.pageRuns).toEqual([
      { text: '2' },
      { text: ', ' },
      { text: '5', bold: true },
      { text: ', ' },
      { text: '9', italic: true },
    ])
    expect(indexLines([{ term: 'Fruit', page: 2 }], false)[0]!.pageRuns).toBeUndefined()
  })
})

describe('marking in the document', () => {
  let editor: Editor
  beforeEach(() => {
    editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            attrs: { bookmarks: ['Orchard'] },
            content: [{ type: 'text', text: 'Apples grow; apples fall. Apples keep.' }],
          },
          { type: 'docParagraph', content: [{ type: 'text', text: 'No fruit here.' }] },
          { type: 'docParagraph', content: [{ type: 'text', text: 'More Apples.' }] },
          { type: 'docParagraph', content: [{ type: 'text', text: 'end' }] },
        ],
      },
    })
  })
  afterEach(() => editor.destroy())

  const marks = () => {
    const out: Array<[string, string | null]> = []
    editor.state.doc.descendants((node) => {
      if (node.type.name === 'docXeMark') out.push([node.attrs.term, node.attrs.switches])
    })
    return out
  }

  it('Mark All marks the first occurrence in each paragraph, capitals and all', () => {
    const ends = occurrenceEnds(editor, 'Apples')
    expect(ends).toHaveLength(2)
    expect(occurrenceEnds(editor, 'Apples', true)).toHaveLength(3)
    markIndexEntries(editor, ends, 'Apples', '\\b')
    expect(marks()).toEqual([
      ['Apples', '\\b'],
      ['Apples', '\\b'],
    ])
    // each mark sits right after the word it marks
    expect(editor.state.doc.child(0).child(0).text).toBe('Apples')
    expect(editor.state.doc.child(0).child(1).type.name).toBe('docXeMark')
  })

  it('the index lists a bold page in bold and a page range entry on the page of its bookmark', () => {
    markIndexEntries(editor, [occurrenceEnds(editor, 'More')[0]!], 'Apples', '\\b')
    markIndexEntries(editor, [occurrenceEnds(editor, 'fruit')[0]!], 'Orchards', '\\r Orchard')
    markIndexEntries(editor, [occurrenceEnds(editor, 'No')[0]!], 'Pears', '\\t "See Apples"')
    editor.commands.setTextSelection(editor.state.doc.content.size - 1)
    // each top-level block is on a page of its own
    const pagesOf = (positions: number[]) =>
      positions.map((pos) => {
        let page = 0
        editor.state.doc.forEach((_, offset, index) => {
          if (offset === pos) page = index + 1
        })
        return page
      })
    insertIndex(editor, [], pagesOf)
    const lines: Array<[string, string, unknown]> = []
    editor.state.doc.forEach((node) => {
      const shown = node.attrs.fieldDisplay as
        { left: string; right: string; rightRuns?: unknown } | undefined
      if (shown && /Index\d/.test(String(node.attrs.genXml)))
        lines.push([shown.left, shown.right, shown.rightRuns])
    })
    expect(lines).toEqual([
      ['Apples', '3', [{ text: '3', bold: true }]],
      ['Orchards', '1', undefined],
      ['Pears. See Apples', '', undefined],
    ])
    const xml: string[] = []
    editor.state.doc.forEach((node) => xml.push(String(node.attrs.genXml ?? '')))
    expect(xml.join('')).toMatch(/<w:b\/><w:noProof\/><\/w:rPr><w:t xml:space="preserve">3<\/w:t>/)
  })
})
