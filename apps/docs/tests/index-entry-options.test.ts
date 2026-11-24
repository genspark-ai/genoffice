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
import { pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'

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

describe('a page range entry', () => {
  it('reads first page, en dash, last page; a range on one page is that page', () => {
    const lines = indexLines(
      [
        { term: 'Orchards', page: 3, pageTo: 5 },
        { term: 'Orchards', page: 9 },
        { term: 'Pears', page: 4 },
        { term: 'Orchards', page: 3 },
      ],
      false,
    )
    expect(lines.map((l) => [l.text, l.pageNo])).toEqual([
      ['Orchards', '3, 3\u20135, 9'],
      ['Pears', '4'],
    ])
  })

  it('runs from the page its bookmark opens on to the page the bookmark closes on', () => {
    const para = (docxIndex: number, text: string, bookmarks?: string[]) => ({
      type: 'docParagraph',
      attrs: { docxIndex, ...(bookmarks ? { bookmarks } : {}) },
      content: [{ type: 'text', text }],
    })
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          para(0, 'Apples grow.', ['Orchard']),
          para(1, 'Apples fall.'),
          para(2, 'Apples keep.'),
          para(3, 'Pears too.', ['Row']),
          para(4, 'end'),
        ],
      },
    })
    const p = (inner: string) => `<w:p>${inner}<w:r><w:t>x</w:t></w:r></w:p>`
    const blocks = [
      p('<w:bookmarkStart w:id="7" w:name="Orchard"/>'),
      p(''),
      `<w:p><w:r><w:t>x</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p>`,
      // the bookmark closes ahead of the next paragraph's text: it ends with this one
      p('<w:bookmarkStart w:id="8" w:name="Row"/>'),
      p('<w:bookmarkEnd w:id="8"/>'),
    ].map((originalXml, docxIndex) => ({ type: 'paragraph', docxIndex, originalXml }))
    markIndexEntries(editor, [occurrenceEnds(editor, 'fall')[0]!], 'Orchards', '\\r Orchard')
    markIndexEntries(editor, [occurrenceEnds(editor, 'too')[0]!], 'Rows', '\\r Row')
    editor.commands.setTextSelection(editor.state.doc.content.size - 1)
    const indexOf = (pos: number) => {
      let at = 0
      editor.state.doc.forEach((_, offset, index) => {
        if (offset === pos) at = index
      })
      return at
    }
    // a block starts on page 2n+1 and ends on the page after it
    const pagesOf = (positions: number[], edge?: 'top' | 'bottom') =>
      positions.map((pos) => indexOf(pos) * 2 + (edge === 'bottom' ? 2 : 1))
    insertIndex(editor, blocks as never, pagesOf)
    const lines: Array<[string, string]> = []
    editor.state.doc.forEach((node) => {
      const shown = node.attrs.fieldDisplay as { left: string; right: string } | undefined
      if (shown && /Index\d/.test(String(node.attrs.genXml))) lines.push([shown.left, shown.right])
    })
    expect(lines).toEqual([
      ['Orchards', '1\u20136'],
      ['Rows', '7\u20138'],
    ])
    editor.destroy()
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
    // saved, the bold page stays a run of its own behind one separator
    const saved = pmDocToSavePlan(editor.getJSON() as PmNode, [])
      .saveBlocks.map((block) => (block.kind === 'xml' ? block.xml : ''))
      .join('')
    expect(saved).toMatch(
      />Apples<\/w:t><\/w:r><w:r><w:rPr><w:noProof\/><\/w:rPr><w:t xml:space="preserve">, <\/w:t><\/w:r><w:r><w:rPr><w:b\/><w:noProof\/><\/w:rPr><w:t xml:space="preserve">3<\/w:t>/,
    )
    expect(saved).toContain('>, 1</w:t>')
  })
})
