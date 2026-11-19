/** The Index and Tables options are the switches of the table's field: written, followed on the page, kept by Update, read back. */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { generateCaptionXml, type FieldDisplay } from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  authoritiesOptionsOf,
  authoritiesSpec,
  captionBody,
  DEFAULT_SPECS,
  figuresOptionsOf,
  figuresSpec,
  indexOptionsOf,
  indexSpec,
  tocOptionsOf,
  tocSpec,
  type TableSpec,
} from '../src/renderer/editor/index-tables'
import { indexSectionEdges } from '../src/renderer/commands/reference-commands'
import { collectHeadings } from '../src/renderer/editor/headings'
import { applyTocPageDisplays } from '../src/renderer/editor/toc-refresh'
import { updateTocField } from '../src/renderer/editor/toc-update'
import {
  indexLines,
  insertIndex,
  insertTableOfAuthorities,
  insertTableOfFigures,
  insertToc,
  markCitation,
  tableAtSelection,
  tableSpecOfPayload,
  tableSpecPayload,
  updateIndex,
  updateTableOfFigures,
} from '../src/renderer/commands/reference-commands'

describe('Index and Tables options as field switches', () => {
  it('reads the default table of contents Word writes and writes each option as its switch', () => {
    const word: TableSpec = { instr: ' TOC \\o "1-3" \\h \\z \\u ', leader: 'dot' }
    const options = tocOptionsOf(word)
    expect(options).toEqual({
      levels: 3,
      styles: [],
      outline: true,
      entryFields: false,
      pageNumbers: true,
      rightAlign: true,
      leader: 'dot',
      hyperlinks: true,
    })
    expect(tocSpec(options, word)).toEqual(word)
    expect(tocSpec({ ...options, levels: 2 }, word).instr).toBe(' TOC \\o "1-2" \\h \\z \\u ')
    expect(tocSpec({ ...options, pageNumbers: false }, word).instr).toBe(
      ' TOC \\o "1-3" \\h \\z \\u \\n ',
    )
    expect(tocSpec({ ...options, rightAlign: false }, word).instr).toBe(
      ' TOC \\o "1-3" \\h \\z \\u \\p " " ',
    )
    expect(tocSpec({ ...options, hyperlinks: false }, word).instr).toBe(' TOC \\o "1-3" \\z \\u ')
    expect(tocSpec({ ...options, leader: 'underscore' }, word).leader).toBe('underscore')
  })

  it('keeps the switches the dialog does not show', () => {
    const base: TableSpec = {
      instr: ' TOC \\o "2-4" \\t "Appendix,1" \\p "-" \\w ',
      leader: 'none',
    }
    const options = tocOptionsOf(base)
    expect(options).toMatchObject({ levels: 4, rightAlign: false, hyperlinks: false })
    expect(tocSpec({ ...options, levels: 5 }, base).instr).toBe(
      ' TOC \\o "2-5" \\t "Appendix,1" \\p "-" \\w ',
    )
  })

  it('a table of figures lists a label with \\c, or without label and number with \\a', () => {
    const base: TableSpec = { instr: ' TOC \\h \\z \\c "Figure" ', leader: 'dot' }
    const options = figuresOptionsOf(base)
    expect(options).toMatchObject({ label: 'Figure', includeLabel: true, pageNumbers: true })
    const bare = figuresSpec({ ...options, includeLabel: false, label: 'Table' }, base)
    expect(bare.instr).toBe(' TOC \\h \\z \\a "Table" ')
    expect(figuresOptionsOf(bare)).toMatchObject({ label: 'Table', includeLabel: false })
    expect(figuresSpec({ ...options, pageNumbers: false }, base).instr).toBe(
      ' TOC \\h \\z \\c "Figure" \\n ',
    )
  })

  it('an index is run-in with \\r and right aligns its pages with a tab as \\e', () => {
    const base: TableSpec = { instr: ' INDEX \\c "2" \\z "1033" ', leader: 'dot' }
    expect(indexOptionsOf(base)).toEqual({
      runIn: false,
      accented: false,
      rightAlign: false,
      leader: 'dot',
      columns: 2,
    })
    const next = indexSpec(
      { runIn: true, accented: false, rightAlign: true, leader: 'hyphen', columns: 2 },
      base,
    )
    expect(next.instr).toBe(' INDEX \\c "2" \\z "1033" \\r \\e "\t" ')
    expect(indexOptionsOf(next)).toMatchObject({ runIn: true, rightAlign: true, leader: 'hyphen' })
    expect(
      indexSpec({ ...indexOptionsOf(next), rightAlign: false, runIn: false }, next).instr,
    ).toBe(' INDEX \\c "2" \\z "1033" ')
  })

  it('an index carries its column count as \\c, and none for one column', () => {
    const base: TableSpec = { instr: ' INDEX \\c "2" \\z "1033" ', leader: 'dot' }
    const options = indexOptionsOf(base)
    expect(indexSpec({ ...options, columns: 3 }, base).instr).toBe(' INDEX \\c "3" \\z "1033" ')
    expect(indexSpec({ ...options, columns: 1 }, base).instr).toBe(' INDEX \\z "1033" ')
    expect(indexSpec({ ...indexOptionsOf(DEFAULT_SPECS.index), columns: 2 }).instr).toBe(
      ' INDEX \\c "2" ',
    )
  })

  it('a table of authorities lists every category without \\c and uses passim with \\p', () => {
    expect(authoritiesOptionsOf(DEFAULT_SPECS.authorities)).toEqual({
      category: 0,
      passim: true,
      keepFormatting: true,
      leader: 'dot',
    })
    const one = authoritiesSpec({
      category: 2,
      passim: false,
      keepFormatting: true,
      leader: 'dot',
    })
    expect(one.instr).toBe(' TOA \\h \\c "2" ')
    // Word lists a citation as it was marked unless \\f takes its formatting off
    const plain = authoritiesSpec({ ...authoritiesOptionsOf(one), keepFormatting: false }, one)
    expect(plain.instr).toBe(' TOA \\h \\c "2" \\f ')
    expect(authoritiesOptionsOf(plain).keepFormatting).toBe(false)
    expect(authoritiesOptionsOf(one)).toMatchObject({ category: 2, passim: false })
  })

  it('strips a caption down to its own text', () => {
    expect(captionBody('Figure 1 Overview', 'Figure')).toBe('Overview')
    expect(captionBody('Figure 2-3: Sales by region', 'Figure')).toBe('Sales by region')
    expect(captionBody('Table IV. Costs', 'Table')).toBe('Costs')
    expect(captionBody('3 Detail', 'Figure')).toBe('Detail')
  })

  it('carries a spec through a command payload', () => {
    const spec: TableSpec = { instr: ' TOC \\o "1-2" \\p " " ', leader: 'hyphen' }
    expect(tableSpecOfPayload(tableSpecPayload(spec))).toEqual(spec)
    expect(tableSpecOfPayload('Figure')).toBeUndefined()
    expect(tableSpecOfPayload(undefined)).toBeUndefined()
  })

  it('lays subentries out indented or run-in', () => {
    const marks = [
      { term: 'Zebra', page: 4 },
      { term: 'Apple:Red', page: 2 },
      { term: 'Apple:Green', page: 1 },
      { term: 'Apple:Red', page: 5 },
      { term: 'Mango', page: 3 },
    ]
    expect(indexLines(marks, false).map((l) => [l.level, l.text, l.pageNo ?? ''])).toEqual([
      [1, 'Apple', ''],
      [2, 'Green', '1'],
      [2, 'Red', '2, 5'],
      [1, 'Mango', '3'],
      [1, 'Zebra', '4'],
    ])
    expect(indexLines(marks, true).map((l) => [l.text, l.pageNo ?? ''])).toEqual([
      ['Apple: Green, 1; Red, 2, 5', ''],
      ['Mango', '3'],
      ['Zebra', '4'],
    ])
  })
})

const heading = (level: number, text: string) => ({
  type: 'docHeading',
  attrs: { level },
  content: [{ type: 'text', text }],
})
const paragraph = (text: string, marks: string[] = []) => ({
  type: 'docParagraph',
  content: [
    { type: 'text', text },
    ...marks.map((term) => ({ type: 'docXeMark', attrs: { term } })),
  ],
})
const caption = (label: string, n: number, text: string) => ({
  type: 'docProtected',
  attrs: { blockType: 'passthrough', genXml: generateCaptionXml(label, n, text) },
})

describe('tables generated from a spec', () => {
  let editor: Editor
  beforeEach(() => {
    editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          paragraph('start'),
          heading(1, 'Alpha'),
          paragraph('about alpha', ['Fruit:Apple', 'Zebra']),
          heading(2, 'Beta'),
          caption('Figure', 1, 'Overview'),
          heading(3, 'Gamma'),
          paragraph('about gamma', ['Fruit:Pear']),
        ],
      },
    })
    editor.commands.setTextSelection(2)
    vi.spyOn(window, 'alert').mockImplementation(() => {})
  })
  afterEach(() => {
    editor.destroy()
    vi.restoreAllMocks()
  })

  const tableLines = (style: RegExp) => {
    const out: Array<{ xml: string; shown: FieldDisplay; pos: number }> = []
    editor.state.doc.forEach((node, pos) => {
      const xml = String(node.attrs.genXml ?? '')
      if (style.test(xml)) out.push({ xml, shown: node.attrs.fieldDisplay as FieldDisplay, pos })
    })
    return out
  }
  const tocLines = () => tableLines(/w:pStyle w:val="TOC\d"/)
  const pages = () => [3, 4, 5]

  it('lists the levels of \\o, writes the leader, and Update keeps both', () => {
    insertToc(editor, [], pages, { instr: ' TOC \\o "1-2" \\h \\z \\u ', leader: 'hyphen' })
    const check = () => {
      const lines = tocLines()
      expect(lines.map((l) => [l.shown.left, l.shown.right, l.shown.leader])).toEqual([
        ['Alpha', '3', 'hyphen'],
        ['Beta', '4', 'hyphen'],
      ])
      expect(lines[0].xml).toContain(' TOC \\o "1-2" \\h \\z \\u ')
      expect(lines[0].xml).toContain('w:leader="hyphen"')
      expect(lines[0].shown.noLink).toBeUndefined()
    }
    check()
    expect(updateTocField(editor, [], pages)).toBe(true)
    check()
  })

  it('shows no page numbers with \\n and the page after the separator with \\p', () => {
    insertToc(editor, [], pages, { instr: ' TOC \\o "1-3" \\n ', leader: 'dot' })
    let lines = tocLines()
    expect(lines.map((l) => [l.shown.left, l.shown.right, l.shown.noPage])).toEqual([
      ['Alpha', '', true],
      ['Beta', '', true],
      ['Gamma', '', true],
    ])
    expect(lines.map((l) => l.xml).join('')).not.toContain('w:tab')
    expect(lines[0].shown.noLink).toBe(true)

    editor.commands.setNodeSelection(lines[0].pos)
    expect(tableAtSelection(editor, [], 'toc')?.spec.instr).toBe(' TOC \\o "1-3" \\n ')
    // with the selection in the table, the dialog's spec rebuilds it in place
    insertToc(editor, [], pages, { instr: ' TOC \\o "1-1" \\p " " \\h ', leader: 'dot' })
    lines = tocLines()
    expect(lines.map((l) => [l.shown.left, l.shown.right, l.shown.pageSep])).toEqual([
      ['Alpha', '3', ' '],
    ])
    expect(lines[0].xml).toContain('<w:t xml:space="preserve"> 3</w:t>')
    expect(lines[0].xml).not.toContain('<w:tabs>')
  })

  it('a table of figures drops label and number with \\a, and Update keeps the switch', () => {
    insertTableOfFigures(editor, [], () => [7], undefined, {
      instr: ' TOC \\h \\z \\a "Figure" ',
      leader: 'underscore',
    })
    const check = () => {
      const lines = tableLines(/TableofFigures/)
      expect(lines.map((l) => [l.shown.left, l.shown.right, l.shown.leader])).toEqual([
        ['Overview', '7', 'underscore'],
      ])
      expect(lines[0].xml).toContain(' TOC \\h \\z \\a "Figure" ')
    }
    check()
    expect(updateTableOfFigures(editor, [], () => [7])).toBe(true)
    check()
    // the table of figures is not taken for the table of contents
    expect(updateTocField(editor, [], pages)).toBe(false)
  })

  it('an index lists pages after its entries, or right aligned behind a tab, run-in or indented', () => {
    const pagesOf = (positions: number[]) => positions.map((_, i) => i + 1)
    insertIndex(editor, [], pagesOf, { instr: ' INDEX ', leader: 'dot' })
    let lines = tableLines(/w:pStyle w:val="Index\d"/)
    expect(lines.map((l) => [l.shown.level, l.shown.left, l.shown.right, l.shown.pageSep])).toEqual(
      [
        [1, 'Fruit', '', undefined],
        [2, 'Apple', '1', ', '],
        [2, 'Pear', '3', ', '],
        [1, 'Zebra', '2', ', '],
      ],
    )
    expect(lines[1].xml).toContain('<w:t xml:space="preserve">, 1</w:t>')
    // a subentry indents by what its paragraph is written with, not by the TOC level ladder
    expect(lines.map((l) => l.shown.indentLeftTwips)).toEqual([0, 240, 240, 0])
    expect(lines[1].xml).toContain('<w:ind w:left="240"/>')

    editor.commands.setNodeSelection(lines[0].pos)
    insertIndex(editor, [], pagesOf, { instr: ' INDEX \\r \\e "\t" ', leader: 'hyphen' })
    const check = () => {
      lines = tableLines(/w:pStyle w:val="Index\d"/)
      expect(lines.map((l) => [l.shown.left, l.shown.right, l.shown.leader])).toEqual([
        ['Fruit: Apple, 1; Pear, 3', '', 'hyphen'],
        ['Zebra', '2', 'hyphen'],
      ])
      expect(lines[0].xml).toContain(' INDEX \\r \\e "\t" ')
      expect(lines[1].xml).toContain('w:leader="hyphen"')
    }
    check()
    expect(updateIndex(editor, [], pagesOf)).toBe(true)
    check()
  })

  it('a heading page refresh leaves the lines of an index alone', () => {
    editor.commands.insertContentAt(editor.state.doc.content.size, paragraph('more', ['Alpha']))
    editor.commands.setTextSelection(editor.state.doc.content.size - 1)
    insertIndex(editor, [], (positions) => positions.map(() => 9), {
      instr: ' INDEX ',
      leader: 'dot',
    })
    insertToc(editor, [], pages, { instr: ' TOC \\o "1-1" ', leader: 'dot' })
    const tr = editor.state.tr
    const refreshed = applyTocPageDisplays(
      editor.state.doc,
      tr,
      collectHeadings(editor.state.doc),
      ['21', '22', '23'],
    )
    expect(refreshed).toBe(true)
    editor.view.dispatch(tr)
    // "Alpha" is a heading and an index entry: only the table of contents follows the heading
    expect(tocLines().map((l) => [l.shown.left, l.shown.right])).toEqual([['Alpha', '21']])
    const index = tableLines(/w:pStyle w:val="Index\d"/)
    expect(index.find((l) => l.shown.left === 'Alpha')?.shown.right).toBe('9')
  })

  it('a table of authorities lists every category, and five pages or more as passim', () => {
    for (let i = 0; i < 5; i++) markCitation(editor, 3, { long: 'Baldwin v. Alberti', category: 1 })
    markCitation(editor, 3, { long: 'RCW 4.16', category: 2 })
    const pagesOf = (positions: number[]) => positions.map((_, i) => i + 1)
    const headingOf = (category: number) => `Heading ${category}`
    editor.commands.setTextSelection(editor.state.doc.content.size - 1)
    insertTableOfAuthorities(editor, [], pagesOf, headingOf, undefined, {
      instr: ' TOA \\h \\p ',
      leader: 'dot',
    })
    const lines = tableLines(/TOAHeading|TableofAuthorities/)
    expect(lines.map((l) => [l.shown.left, l.shown.right])).toEqual([
      ['Heading 1', ''],
      ['Baldwin v. Alberti', 'passim'],
      ['Heading 2', ''],
      ['RCW 4.16', '1'],
    ])
    expect(lines[0].xml).toContain(' TOA \\h \\p ')
  })
})

describe('the section an index has to itself', () => {
  const sections = [
    { index: 0, firstBlock: 0, lastBlock: 2 },
    { index: 1, firstBlock: 3, lastBlock: 6 },
    { index: 2, firstBlock: 7, lastBlock: 9 },
  ]
  const never = () => false

  it('ends at the index, or at a paragraph after it that only carries the break', () => {
    expect(indexSectionEdges([3, 6], sections, never)).toMatchObject({
      startsIt: true,
      endsIt: true,
    })
    expect(indexSectionEdges([3, 5], sections, (block) => block === 6)).toMatchObject({
      startsIt: true,
      endsIt: true,
    })
  })

  it('is shared when the block after the index is text, though that text closes the section', () => {
    expect(indexSectionEdges([3, 5], sections, never)).toMatchObject({
      startsIt: true,
      endsIt: false,
    })
    expect(indexSectionEdges([4, 6], sections, never)).toMatchObject({
      startsIt: false,
      endsIt: true,
    })
  })

  it('in the last section a block after the index is never its closer', () => {
    expect(indexSectionEdges([7, 8], sections, () => true)).toMatchObject({ endsIt: false })
    expect(indexSectionEdges([7, 9], sections, never)).toMatchObject({ endsIt: true })
    expect(indexSectionEdges([12, 13], sections, never)).toBeNull()
  })
})
