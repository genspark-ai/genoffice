// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import type { Block, FieldDisplay } from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  figuresOptionsOf,
  figuresSpec,
  indexOptionsOf,
  indexSpec,
  tocEntryField,
  tocOptionsOf,
  tocPlan,
  tocSpec,
  tocStyleLevels,
  withIndexHeadings,
  type TableSpec,
} from '../src/renderer/editor/index-tables'
import {
  entryStyleUpsert,
  ENTRY_STYLES,
  formatIndexHeading,
  formatStyleUpserts,
  TABLE_FORMATS,
} from '../src/renderer/editor/table-formats'
import { autoMarkIndex, concordanceRows } from '../src/renderer/editor/index-automark'
import { collectPlanEntries } from '../src/renderer/editor/toc-update'
import {
  indexLetter,
  indexLines,
  insertIndex,
  insertTableOfAuthorities,
  insertTableOfFigures,
  insertToc,
  markCitation,
} from '../src/renderer/commands/reference-commands'
import { citationRunsOf } from '../src/renderer/editor/citations'

describe('Options as field switches', () => {
  it('reads the styles of \\t with their levels, whatever the list separator', () => {
    expect(tocStyleLevels('Appendix,1,Note Title,2')).toEqual([
      { name: 'Appendix', level: 1 },
      { name: 'Note Title', level: 2 },
    ])
    expect(tocStyleLevels('Appendix;1;Note Title;2')).toEqual([
      { name: 'Appendix', level: 1 },
      { name: 'Note Title', level: 2 },
    ])
    expect(tocStyleLevels('Caption')).toEqual([{ name: 'Caption', level: 1 }])
  })

  it('a table of contents names its sources: \\o headings, \\t styles, \\u outline levels, \\f TC fields', () => {
    const word: TableSpec = { instr: ' TOC \\o "1-3" \\h \\z \\u ', leader: 'dot' }
    const options = tocOptionsOf(word)
    const styled = tocSpec(
      { ...options, styles: [{ name: 'Appendix', level: 2 }], outline: false, entryFields: true },
      word,
    )
    expect(styled.instr).toBe(' TOC \\o "1-3" \\h \\z \\t "Appendix,2" \\f ')
    expect(tocOptionsOf(styled)).toMatchObject({
      levels: 3,
      styles: [{ name: 'Appendix', level: 2 }],
      outline: false,
      entryFields: true,
    })
    const only = tocSpec({ ...tocOptionsOf(styled), levels: 0, entryFields: false }, styled)
    expect(only.instr).toBe(' TOC \\h \\z \\t "Appendix,2" ')
    expect(tocPlan(only.instr)).toMatchObject({ headings: false, entryFields: false })
    expect(tocOptionsOf(only).levels).toBe(0)
    // an identifier the field already names is kept
    const named: TableSpec = { instr: ' TOC \\f "B" \\l "2-3" ', leader: 'dot' }
    expect(tocPlan(named.instr)).toMatchObject({
      headings: false,
      entryFields: true,
      entryId: 'B',
      entryLevels: [2, 3],
    })
    expect(tocSpec(tocOptionsOf(named), named).instr).toBe(' TOC \\f "B" \\l "2-3" ')
  })

  it('a table of figures takes a style with \\t and TC fields of one identifier with \\f', () => {
    const base: TableSpec = { instr: ' TOC \\h \\z \\c "Figure" ', leader: 'dot' }
    const next = figuresSpec({ ...figuresOptionsOf(base), style: 'Caption', entryId: 'F' }, base)
    expect(next.instr).toBe(' TOC \\h \\z \\c "Figure" \\t "Caption,1" \\f "F" ')
    expect(figuresOptionsOf(next)).toMatchObject({
      label: 'Figure',
      style: 'Caption',
      entryId: 'F',
    })
    expect(figuresSpec({ ...figuresOptionsOf(next), style: '', entryId: '' }, next).instr).toBe(
      base.instr,
    )
  })

  it('reads a TC field: text, identifier and level', () => {
    expect(tocEntryField('"Appendix A" \\f F \\l 2')).toEqual({
      text: 'Appendix A',
      id: 'F',
      level: 2,
    })
    expect(tocEntryField('"Say \\"hi\\"" \\n')).toEqual({ text: 'Say "hi"', level: 1 })
    expect(tocEntryField('Entry')).toEqual({ text: 'Entry', level: 1 })
  })

  it('an index files accented letters apart with \\a and heads each letter with \\h', () => {
    const base: TableSpec = { instr: ' INDEX \\c "2" ', leader: 'dot' }
    const accented = indexSpec({ ...indexOptionsOf(base), accented: true }, base)
    expect(accented.instr).toBe(' INDEX \\c "2" \\a ')
    expect(indexOptionsOf(accented).accented).toBe(true)
    expect(withIndexHeadings(accented.instr, 'A')).toBe(' INDEX \\c "2" \\a \\h "A" ')
    expect(withIndexHeadings(' INDEX \\h "A" \\c "2" ', null)).toBe(' INDEX \\c "2" ')
  })
})

describe('index letter headings', () => {
  const marks = ['apple', '\u00c9clair', 'egg', '\u00e9t\u00e9', 'Zebra', '42'].map((term, i) => ({
    term,
    page: i + 1,
  }))
  const shape = (lines: ReturnType<typeof indexLines>) =>
    lines.map((l) => (l.style === 'IndexHeading' ? `[${l.text}]` : l.text))

  it('files an entry under its first letter, without the accent unless asked', () => {
    expect(indexLetter('\u00e9t\u00e9', false)).toBe('E')
    expect(indexLetter('\u00e9t\u00e9', true)).toBe('\u00c9')
    expect(indexLetter('42', false)).toBe('')
    expect(indexLetter('\u4e2d', false)).toBe('')
  })

  it('puts a heading above each letter; entries before the first letter have none', () => {
    expect(shape(indexLines(marks, false, ', ', { pattern: 'A' }))).toEqual([
      '42',
      '[A]',
      'apple',
      '[E]',
      '\u00c9clair',
      'egg',
      '\u00e9t\u00e9',
      '[Z]',
      'Zebra',
    ])
  })

  it('gives accented letters a heading of their own with \\a, and follows the pattern', () => {
    expect(shape(indexLines(marks, false, ', ', { pattern: '-A-', accented: true }))).toEqual([
      '42',
      '[-A-]',
      'apple',
      '[-E-]',
      'egg',
      '[-\u00c9-]',
      '\u00c9clair',
      '\u00e9t\u00e9',
      '[-Z-]',
      'Zebra',
    ])
  })

  it('only the A of the pattern stands for the letter', () => {
    const one = [{ term: 'Zebra', page: 1 }]
    expect(shape(indexLines(one, false, ', ', { pattern: 'Letter A' }))[0]).toBe('[Letter Z]')
    expect(shape(indexLines(one, false, ', ', { pattern: '(a)' }))[0]).toBe('[(z)]')
  })

  it('has no headings without \\h, and the order it always had', () => {
    expect(shape(indexLines(marks, false))).toEqual(
      shape(indexLines(marks, false, ', ', { accented: true })),
    )
    expect(indexLines(marks, false).some((l) => l.style === 'IndexHeading')).toBe(false)
  })
})

describe('formats and entry styles', () => {
  it('lists the formats of each tab, From template first', () => {
    for (const formats of Object.values(TABLE_FORMATS)) expect(formats[0]).toBe('template')
    expect(TABLE_FORMATS.index).toContain('bulleted')
    expect(TABLE_FORMATS.figures).toContain('centered')
  })

  it('From template defines nothing; another format redefines the entry styles', () => {
    expect(formatStyleUpserts('toc', 'template', new Map(), 'Normal')).toEqual([])
    const classic = formatStyleUpserts('toc', 'classic', new Map(), 'Normal')
    expect(classic.map((u) => u.styleId)).toEqual(['TOC1', 'TOC2', 'TOC3'])
    expect(classic[0]).toMatchObject({
      name: 'toc 1',
      builtIn: true,
      basedOn: 'Normal',
      rPr: { bold: true, caps: true },
    })
    // a style the document has is patched, never recreated
    const has = new Map([['TOC1', { name: 'toc 1' }]])
    const patched = formatStyleUpserts('toc', 'simple', has, 'Normal')[0]
    expect(patched.name).toBeUndefined()
    // a size an earlier format set is cleared as well
    expect(patched.rPr).toMatchObject({ bold: false, caps: false, sizeHalfPoints: null })
  })

  it('finds a built-in entry style under the id the document gives it', () => {
    const styles = new Map([['Verzeichnis1', { name: 'TOC 1' }]])
    expect(
      entryStyleUpsert(ENTRY_STYLES.toc[0], { rPr: { bold: true } }, styles, 'Normal'),
    ).toEqual({
      styleId: 'Verzeichnis1',
      rPr: { bold: true },
    })
  })

  it('an index format brings letter headings, Simple and From template none', () => {
    expect(formatIndexHeading('template')).toBeUndefined()
    expect(formatIndexHeading('classic')).toBe('A')
    expect(formatIndexHeading('simple')).toBeNull()
  })
})

const heading = (level: number, text: string, attrs: Record<string, unknown> = {}) => ({
  type: 'docHeading',
  attrs: { level, ...attrs },
  content: [{ type: 'text', text }],
})
type Json = Record<string, unknown>
const paragraph = (text: string, attrs: Json = {}, marks: Json[] = []) => ({
  type: 'docParagraph',
  attrs,
  content: [{ type: 'text', text }, ...marks],
})
const tc = (instr: string) => ({ type: 'docXeMark', attrs: { term: '', tc: instr } })

describe('tables built from Options sources', () => {
  let editor: Editor
  beforeEach(() => {
    editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          heading(1, 'Alpha'),
          paragraph('Appendix one', { styleId: 'Appendix' }),
          heading(2, 'Outline only', { outlineOnly: true }),
          paragraph('body text', {}, [tc('"Marked entry" \\l 2'), tc('"Figure entry" \\f F')]),
          heading(2, 'Beta'),
          paragraph('Plate one', { styleId: 'Plate' }),
        ],
      },
    })
    editor.storage.listNumbering.styles = new Map([
      ['Appendix', { styleId: 'Appendix', name: 'Appendix Title', type: 'paragraph' }],
      ['Plate', { styleId: 'Plate', name: 'Plate', type: 'paragraph' }],
    ])
    editor.commands.setTextSelection(editor.state.doc.content.size - 1)
    vi.spyOn(window, 'alert').mockImplementation(() => {})
  })
  afterEach(() => {
    editor.destroy()
    vi.restoreAllMocks()
  })

  const entries = (instr: string) =>
    collectPlanEntries(editor, tocPlan(instr)).map((e) => `${e.level}:${e.text}`)
  const lines = (style: RegExp) => {
    const out: Array<{ shown: FieldDisplay; styleId: unknown; xml: string }> = []
    editor.state.doc.forEach((node) => {
      const xml = String(node.attrs.genXml ?? '')
      if (style.test(xml))
        out.push({
          shown: node.attrs.fieldDisplay as FieldDisplay,
          styleId: node.attrs.styleId,
          xml,
        })
    })
    return out
  }

  it('lists headings alone by default, outline-only paragraphs only with \\u', () => {
    expect(entries(' TOC \\o "1-3" \\h \\z ')).toEqual(['1:Alpha', '2:Beta'])
    expect(entries(' TOC \\o "1-3" \\h \\z \\u ')).toEqual(['1:Alpha', '2:Outline only', '2:Beta'])
    expect(entries(' TOC \\o "1-1" ')).toEqual(['1:Alpha'])
    // outline levels alone: a heading has one through its style, and Options reads Styles as off
    expect(entries(' TOC \\u ')).toEqual(['1:Alpha', '2:Outline only', '2:Beta'])
    expect(tocOptionsOf({ instr: ' TOC \\u ', leader: 'dot' })).toMatchObject({
      levels: 0,
      outline: true,
    })
    const outlineOnly = { instr: ' TOC \\u ', leader: 'dot' } as const
    expect(tocSpec(tocOptionsOf(outlineOnly), outlineOnly).instr).toBe(' TOC \\u ')
  })

  it('lists the paragraphs of a style named in \\t at its level, by style name', () => {
    expect(entries(' TOC \\o "1-3" \\t "Appendix Title,3" ')).toEqual([
      '1:Alpha',
      '3:Appendix one',
      '2:Beta',
    ])
    expect(entries(' TOC \\t "appendix title,1,Heading 2,4" ')).toEqual([
      '1:Appendix one',
      '4:Beta',
    ])
  })

  it('lists TC fields with \\f: those without an identifier, or those of the one named', () => {
    expect(entries(' TOC \\f ')).toEqual(['2:Marked entry'])
    expect(entries(' TOC \\f \\l "1-1" ')).toEqual([])
    expect(entries(' TOC \\f F ')).toEqual(['1:Figure entry'])
    expect(entries(' TOC \\o "1-1" \\f ')).toEqual(['1:Alpha', '2:Marked entry'])
  })

  it('inserts a table of contents over those sources, each line in its TOC style', () => {
    insertToc(
      editor,
      [],
      () => null,
      { instr: ' TOC \\o "1-1" \\t "Appendix Title,2" \\f ', leader: 'dot' },
      (positions) => positions.map((_, i) => i + 1),
    )
    const toc = lines(/w:pStyle w:val="TOC\d"/)
    expect(toc.map((l) => [l.shown.left, l.shown.right, l.styleId])).toEqual([
      ['Alpha', '1', 'TOC1'],
      ['Appendix one', '2', 'TOC2'],
      ['Marked entry', '3', 'TOC2'],
    ])
    expect(toc[0].xml).toContain(' TOC \\o "1-1" \\t "Appendix Title,2" \\f ')
  })

  it('a table of figures lists a style, or the TC fields of an identifier, in place of captions', () => {
    const pages = (positions: number[]) => positions.map(() => 7)
    insertTableOfFigures(editor, [], pages, undefined, {
      instr: ' TOC \\h \\z \\c "Figure" \\t "Plate,1" ',
      leader: 'dot',
    })
    expect(lines(/TableofFigures/).map((l) => [l.shown.left, l.shown.right, l.styleId])).toEqual([
      ['Plate one', '7', 'TableofFigures'],
    ])
    editor.commands.setContent({
      type: 'doc',
      content: [paragraph('x', {}, [tc('"Figure entry" \\f F')])],
    })
    insertTableOfFigures(editor, [], pages, undefined, {
      instr: ' TOC \\h \\z \\c "Figure" \\f "F" ',
      leader: 'dot',
    })
    expect(lines(/TableofFigures/).map((l) => l.shown.left)).toEqual(['Figure entry'])
  })

  it('writes a line in the id the document stores the built-in style under', () => {
    editor.storage.listNumbering.styles = new Map([
      ['Verzeichnis1', { styleId: 'Verzeichnis1', name: 'toc 1', type: 'paragraph' }],
    ])
    const spec = { instr: ' TOC \\o "1-2" ', leader: 'dot' } as const
    insertToc(editor, [], () => null, spec)
    const withSpec = lines(/w:pStyle w:val="(?:Verzeichnis1|TOC2)"/)
    expect(withSpec.map((l) => l.styleId)).toEqual(['Verzeichnis1', 'TOC2'])
    editor.commands.setTextSelection(1)
    insertToc(editor, [], () => null)
    expect(lines(/w:pStyle w:val="Verzeichnis1"/).map((l) => l.styleId)).toEqual([
      'Verzeichnis1',
      'Verzeichnis1',
    ])
    expect(lines(/w:pStyle w:val="TOC1"/)).toEqual([])
  })

  it('a TC field is not an index entry', () => {
    insertIndex(editor, [], () => null)
    expect(lines(/w:pStyle w:val="Index\d"/)).toEqual([])
  })
})

describe('AutoMark', () => {
  const cell = (text: string) => ({ paras: [text] })
  const table = (rows: string[][]): Block =>
    ({
      id: 't',
      type: 'table',
      docxIndex: 0,
      originalXml: '',
      table: { rows: rows.map((r) => r.map(cell)) },
    }) as unknown as Block

  it('reads a concordance table: the text to find, then its entry', () => {
    expect(
      concordanceRows([
        table([
          ['apple', 'Fruit:Apple'],
          ['Zebra', ''],
          ['', 'ignored'],
        ]),
      ]),
    ).toEqual([
      { find: 'apple', entry: 'Fruit:Apple' },
      { find: 'Zebra', entry: 'Zebra' },
    ])
  })

  it('marks the first occurrence in each paragraph once, and nothing the second time', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [paragraph('an apple and an apple'), paragraph('Apple, apple, Zebra')],
      },
    })
    const rows = [
      { find: 'apple', entry: 'Fruit:Apple' },
      { find: 'Zebra', entry: 'Zebra' },
    ]
    expect(autoMarkIndex(editor, rows)).toBe(3)
    const marks: string[] = []
    editor.state.doc.descendants((node, pos) => {
      if (node.type.name === 'docXeMark')
        marks.push(`${node.attrs.term}@${editor.state.doc.textBetween(0, pos, '|')}`)
    })
    expect(marks).toEqual([
      'Fruit:Apple@an apple',
      'Fruit:Apple@an apple and an apple|Apple, apple',
      'Zebra@an apple and an apple|Apple, apple, Zebra',
    ])
    expect(autoMarkIndex(editor, rows)).toBe(0)
    editor.destroy()
  })
})

describe('a table of authorities keeps the formatting a citation was marked with', () => {
  let editor: Editor
  const long = 'Baldwin v. Alberti (1961)'
  beforeEach(() => {
    editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              { type: 'text', text: 'See ' },
              { type: 'text', text: 'Baldwin v. Alberti', marks: [{ type: 'italic' }] },
              { type: 'text', text: ' (1961)' },
              { type: 'text', text: ' above.' },
            ],
          },
        ],
      },
    })
    vi.spyOn(window, 'alert').mockImplementation(() => {})
  })
  afterEach(() => {
    editor.destroy()
    vi.restoreAllMocks()
  })
  const from = 5
  const to = from + long.length
  const table = () => {
    const out: Array<{ shown: FieldDisplay; xml: string }> = []
    editor.state.doc.forEach((node) => {
      const xml = String(node.attrs.genXml ?? '')
      if (/TableofAuthorities/.test(xml))
        out.push({ shown: node.attrs.fieldDisplay as FieldDisplay, xml })
    })
    return out
  }

  it('reads the selection in pieces; not when it is plain or the text was changed', () => {
    expect(citationRunsOf(editor.state.doc, from, to, long)).toEqual([
      { text: 'Baldwin v. Alberti', italic: true },
      { text: ' (1961)' },
    ])
    expect(citationRunsOf(editor.state.doc, from, to, 'Baldwin')).toBeUndefined()
    // a quote cannot stand in the field's text: the pieces lose it as the long citation does
    editor.commands.insertContentAt(editor.state.doc.content.size - 1, [
      { type: 'text', text: ' The ' },
      { type: 'text', text: '"Lemon"', marks: [{ type: 'bold' }] },
      { type: 'text', text: ' test' },
    ])
    const end = editor.state.doc.content.size - 1
    expect(citationRunsOf(editor.state.doc, end - 16, end, 'The Lemon test')).toEqual([
      { text: 'The ' },
      { text: 'Lemon', bold: true },
      { text: ' test' },
    ])
    expect(citationRunsOf(editor.state.doc, to - 6, to, '(1961)')).toBeUndefined()
  })

  it('lists the citation formatted, and plain with \\f', () => {
    const runs = citationRunsOf(editor.state.doc, from, to, long)
    markCitation(editor, to, { long, category: 1, runs })
    const heading = () => 'Cases'
    editor.commands.setTextSelection(editor.state.doc.content.size - 1)
    insertTableOfAuthorities(editor, [], () => null, heading, undefined, {
      instr: ' TOA \\h \\p ',
      leader: 'dot',
    })
    expect(table()[0].shown.leftRuns).toEqual(runs)
    expect(table()[0].xml).toContain('<w:i/>')
    editor.commands.setTextSelection(editor.state.doc.content.size - 1)
    insertTableOfAuthorities(editor, [], () => null, heading, undefined, {
      instr: ' TOA \\h \\p \\f ',
      leader: 'dot',
    })
    const plain = table().at(-1)!
    expect(plain.shown.left).toBe(long)
    expect(plain.shown.leftRuns).toBeUndefined()
    expect(plain.xml).not.toContain('<w:i/>')
  })
})
