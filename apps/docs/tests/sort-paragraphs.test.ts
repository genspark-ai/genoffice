// @vitest-environment jsdom
import { Editor, type JSONContent } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import { CellSelection } from '@tiptap/pm/tables'
import { describe, expect, it } from 'vitest'
import { sortKeyChoices } from '../src/renderer/components/MenuDialogs'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { acceptAllRevisions, rejectAllRevisions } from '../src/renderer/editor/revisions'
import {
  compareByKeys,
  compareParagraphs,
  sortFields,
  sortScope,
  sortSelection,
  type SortBy,
  type SortKey,
  type SortOptions,
  type SortScope,
} from '../src/renderer/editor/sort-paragraphs'

const para = (text: string) => ({
  type: 'docParagraph',
  content: text ? [{ type: 'text', text }] : [],
})

const cell = (text: string, attrs: Record<string, unknown> = {}) => ({
  type: 'docTableCell',
  attrs,
  content: [para(text)],
})

const tableOf = (rows: string[][]) => ({
  type: 'docTable',
  content: rows.map((row) => ({ type: 'docTableRow', content: row.map((text) => cell(text)) })),
})

function makeEditor(content: JSONContent[]): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
}

const linesOf = (editor: Editor) => {
  const lines: string[] = []
  editor.state.doc.forEach((node) => lines.push(node.textContent))
  return lines
}

const key = (field: number, by: SortBy = 'text', descending = false): SortKey => ({
  field,
  by,
  descending,
})

const options = (keys: SortKey[], extra: Partial<SortOptions> = {}): SortOptions => ({
  keys,
  header: false,
  separator: '\t',
  caseSensitive: false,
  ...extra,
})

const one = (by: SortBy, descending = false) => options([key(0, by, descending)])

/** the table's rows, each as its cells' texts; a tracked row is prefixed with its record */
const rowsOf = (editor: Editor) => {
  const rows: string[] = []
  editor.state.doc.descendants((node) => {
    if (node.type.name !== 'docTableRow') return true
    const texts: string[] = []
    node.forEach((child) => texts.push(child.textContent))
    const kind = (node.attrs.rowRevision as { kind?: string } | null)?.kind
    rows.push((kind ? kind + ':' : '') + texts.join('|'))
    return false
  })
  return rows
}

const PEOPLE = [
  ['Name', 'City', 'Age'],
  ['Grace', 'Paris', '41'],
  ['Ada', 'London', '36'],
  ['Alan', 'London', '9'],
  ['Linus', 'Paris', '30'],
]

/** a caret inside the table's first cell */
const intoTable = (editor: Editor) => editor.commands.setTextSelection(4)

describe('Sort', () => {
  it('compares as text, number or date', () => {
    expect(compareParagraphs('text', 'apple', 'Banana')).toBeLessThan(0)
    expect(compareParagraphs('number', 'item 10', 'item 9')).toBeGreaterThan(0)
    expect(compareParagraphs('text', 'item 10', 'item 9')).toBeLessThan(0)
    expect(compareParagraphs('number', 'none', '3')).toBeLessThan(0)
    expect(compareParagraphs('date', '2026-03-01', '2025-12-31')).toBeGreaterThan(0)
    // two paragraphs without a number or a date tie, so they keep their order
    expect(compareParagraphs('number', 'none', 'nothing')).toBe(0)
    expect(compareParagraphs('date', 'soon', 'later')).toBe(0)
    // case tells two texts apart only when asked to
    expect(compareParagraphs('text', 'apple', 'Apple')).toBe(0)
    expect(compareParagraphs('text', 'apple', 'Apple', true)).not.toBe(0)
  })

  it('the first key that tells two records apart decides, each in its own direction', () => {
    const keys = [key(2), key(3, 'number', true)]
    const record = (...cells: string[]) => ['', ...cells]
    expect(compareByKeys(keys, record('Ada', 'London', '36'), record('Zed', 'Paris', '1'))).toBe(-1)
    // the same city: the larger number comes first
    expect(
      compareByKeys(keys, record('Ada', 'London', '36'), record('Alan', 'London', '9')),
    ).toBeLessThan(0)
    expect(
      compareByKeys(keys, record('Alan', 'London', '9'), record('Ada', 'London', '36')),
    ).toBeGreaterThan(0)
    // a tie on every key, and a field the record does not have
    expect(compareByKeys(keys, record('Ada', 'London', '36'), record('Bo', 'london', '36'))).toBe(0)
    expect(compareByKeys([key(5)], record('a'), record('b'))).toBe(0)
    expect(compareByKeys([], record('a'), record('b'))).toBe(0)
  })

  it('a separator parts a paragraph into fields', () => {
    expect(sortFields('a\tb\tc', '\t')).toEqual(['a\tb\tc', 'a', 'b', 'c'])
    expect(sortFields('a,b', ',')).toEqual(['a,b', 'a', 'b'])
    expect(sortFields('a,b', '')).toEqual(['a,b', 'a,b'])
  })

  it('with nothing selected, sorts every paragraph of the body', () => {
    const editor = makeEditor(['pear', 'Apple', 'fig'].map(para))
    expect(sortSelection(editor, one('text'))).toBeNull()
    expect(linesOf(editor)).toEqual(['Apple', 'fig', 'pear'])
    expect(sortSelection(editor, one('text', true))).toBeNull()
    expect(linesOf(editor)).toEqual(['pear', 'fig', 'Apple'])
    editor.destroy()
  })

  it('paragraphs without a number stay in their order ahead of the numbered ones', () => {
    const editor = makeEditor(['zeta', '7', 'alpha', '2', 'mid'].map(para))
    expect(sortSelection(editor, one('number'))).toBeNull()
    expect(linesOf(editor)).toEqual(['zeta', 'alpha', 'mid', '2', '7'])
    editor.destroy()
  })

  it('sorts only the paragraphs the selection touches', () => {
    const editor = makeEditor(['zebra', '30', '4', '200', 'aardvark'].map(para))
    // from inside "30" to inside "200"
    editor.commands.setTextSelection({ from: 9, to: 17 })
    expect(sortSelection(editor, one('number'))).toBeNull()
    expect(linesOf(editor)).toEqual(['zebra', '4', '30', '200', 'aardvark'])
    editor.destroy()
  })

  it('paragraphs sort by two fields, ties keeping their order, under a header', () => {
    const editor = makeEditor(
      ['City,Name', 'Paris,Zoe', 'London,Bo', 'Paris,Al', 'London,Ada', 'London,Bo'].map(para),
    )
    const byCityThenName = [key(1), key(2, 'text', true)]
    expect(
      sortSelection(editor, options(byCityThenName, { header: true, separator: ',' })),
    ).toBeNull()
    expect(linesOf(editor)).toEqual([
      'City,Name',
      'London,Bo',
      'London,Bo',
      'London,Ada',
      'Paris,Zoe',
      'Paris,Al',
    ])
    // one undo puts every paragraph back
    editor.commands.undo()
    expect(linesOf(editor)[1]).toBe('Paris,Zoe')
    // with a header, two paragraphs leave one to sort
    const short = makeEditor(['b', 'a'].map(para))
    expect(sortSelection(short, options([key(0)], { header: true }))).toBe('needsParagraphs')
    expect(linesOf(short)).toEqual(['b', 'a'])
    editor.destroy()
    short.destroy()
  })

  it('case decides between two texts only when the sort is case sensitive', () => {
    const editor = makeEditor(['b', 'A', 'a', 'B'].map(para))
    expect(sortSelection(editor, one('text'))).toBeNull()
    expect(linesOf(editor)).toEqual(['A', 'a', 'b', 'B'])
    expect(sortSelection(editor, options([key(0)], { caseSensitive: true }))).toBeNull()
    expect(linesOf(editor)).toEqual(['a', 'A', 'b', 'B'])
    editor.destroy()
  })

  it('leaves a range with a table in it alone', () => {
    const editor = makeEditor([para('b'), para('a'), tableOf([['cell']])])
    editor.commands.setTextSelection(1)
    const before = linesOf(editor)
    expect(sortScope(editor.state)).toBe('needsParagraphs')
    expect(sortSelection(editor, one('text'))).toBe('needsParagraphs')
    expect(linesOf(editor)).toEqual(before)
    // a single paragraph has nothing to sort against
    editor.commands.setTextSelection({ from: 1, to: 2 })
    expect(sortSelection(editor, one('text'))).toBe('needsParagraphs')
    editor.destroy()
  })

  it('with the caret in a table, its rows sort by two columns under the header row', () => {
    const editor = makeEditor([tableOf(PEOPLE), para('after')])
    intoTable(editor)
    expect(sortScope(editor.state)).toEqual({ kind: 'table', rows: PEOPLE })
    const table = editor.state.doc.firstChild!
    expect(sortSelection(editor, options([key(2), key(3, 'number')], { header: true }))).toBeNull()
    expect(rowsOf(editor)).toEqual([
      'Name|City|Age',
      'Alan|London|9',
      'Ada|London|36',
      'Linus|Paris|30',
      'Grace|Paris|41',
    ])
    expect(linesOf(editor)[1]).toBe('after')
    // the save writes the moved rows out instead of patching texts into the old ones
    expect(editor.state.doc.firstChild!.attrs.originalStructure).toBeNull()
    editor.commands.undo()
    expect(rowsOf(editor)).toEqual(PEOPLE.map((row) => row.join('|')))
    expect(editor.state.doc.firstChild!.eq(table)).toBe(true)
    // without a header the first row sorts with the rest
    expect(sortSelection(editor, options([key(1, 'text', true)]))).toBeNull()
    expect(rowsOf(editor)[0]).toBe('Name|City|Age')
    expect(rowsOf(editor)[1]).toBe('Linus|Paris|30')
    editor.destroy()
  })

  it('a table selected whole sorts all its rows; a selection of cells sorts the rows it spans', () => {
    const editor = makeEditor([tableOf(PEOPLE)])
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, 0)))
    expect(sortSelection(editor, options([key(1)], { header: true }))).toBeNull()
    expect(rowsOf(editor).map((row) => row.split('|')[0])).toEqual([
      'Name',
      'Ada',
      'Alan',
      'Grace',
      'Linus',
    ])
    // rows 2 and 3 (Alan, Grace), descending
    const starts: number[] = []
    editor.state.doc.descendants((node, pos) => {
      if (node.type.name === 'docTableCell') starts.push(pos)
      return node.type.name !== 'docTableCell'
    })
    editor.view.dispatch(
      editor.state.tr.setSelection(CellSelection.create(editor.state.doc, starts[6], starts[9])),
    )
    expect(sortScope(editor.state)).toMatchObject({ rows: [PEOPLE[3], PEOPLE[1]] })
    expect(sortSelection(editor, options([key(1, 'text', true)]))).toBeNull()
    expect(rowsOf(editor).map((row) => row.split('|')[0])).toEqual([
      'Name',
      'Ada',
      'Grace',
      'Alan',
      'Linus',
    ])
    editor.destroy()
  })

  it('refuses a table whose rows hold merged cells', () => {
    const merged = tableOf(PEOPLE)
    merged.content[1].content = [cell('Grace', { colspan: 2 }), cell('41')]
    const editor = makeEditor([merged])
    intoTable(editor)
    const before = rowsOf(editor)
    expect(sortScope(editor.state)).toBe('mergedCells')
    expect(sortSelection(editor, options([key(1)]))).toBe('mergedCells')
    expect(rowsOf(editor)).toEqual(before)
    editor.destroy()
  })

  it('tracked, a row that changed place is struck where it was and inserted where it went', () => {
    const editor = makeEditor([tableOf(PEOPLE)])
    editor.storage.trackChanges.enabled = true
    editor.storage.trackChanges.author = 'Reviewer'
    intoTable(editor)
    expect(sortSelection(editor, options([key(1)], { header: true }))).toBeNull()
    // Linus was last and stays last: no record on a row that did not move
    expect(rowsOf(editor)).toEqual([
      'Name|City|Age',
      'del:Grace|Paris|41',
      'del:Ada|London|36',
      'del:Alan|London|9',
      'ins:Ada|London|36',
      'ins:Alan|London|9',
      'ins:Grace|Paris|41',
      'Linus|Paris|30',
    ])
    const row = editor.state.doc.firstChild!.child(1)
    expect(row.attrs.rowRevision).toMatchObject({ kind: 'del', author: 'Reviewer' })
    expect(row.attrs.rawTrPr).toContain('<w:del ')
    const tracked = editor.state.doc
    acceptAllRevisions(editor)
    expect(rowsOf(editor).map((text) => text.split('|')[0])).toEqual([
      'Name',
      'Ada',
      'Alan',
      'Grace',
      'Linus',
    ])
    editor.view.dispatch(
      editor.state.tr.replaceWith(0, editor.state.doc.content.size, tracked.content),
    )
    rejectAllRevisions(editor)
    expect(rowsOf(editor)).toEqual(PEOPLE.map((cells) => cells.join('|')))
    editor.destroy()
  })

  it('the dialog lists a table by its columns and paragraphs by their fields', () => {
    const table: SortScope = { kind: 'table', rows: PEOPLE }
    expect(sortKeyChoices(table, '\t', false)).toEqual([
      { field: 1, name: null },
      { field: 2, name: null },
      { field: 3, name: null },
    ])
    expect(sortKeyChoices(table, '\t', true).map((choice) => choice.name)).toEqual([
      'Name',
      'City',
      'Age',
    ])
    const paragraphs: SortScope = { kind: 'paragraphs', texts: ['City,Name', 'Paris,Zoe,x'] }
    // no tab in them: the paragraph is the only key
    expect(sortKeyChoices(paragraphs, '\t', false)).toEqual([{ field: 0, name: null }])
    expect(sortKeyChoices(paragraphs, ',', true)).toEqual([
      { field: 0, name: null },
      { field: 1, name: 'City' },
      { field: 2, name: 'Name' },
      { field: 3, name: null },
    ])
    expect(sortKeyChoices(null, '\t', false)).toEqual([{ field: 0, name: null }])
  })
})
