// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx, TABLE_XML } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { bookmarkRange } from '../src/renderer/editor/bookmark-range'
import { addBookmark } from '../src/renderer/editor/cross-reference'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { convertTableToText, convertTextToTable } from '../src/renderer/editor/table-convert'
import {
  acceptAllRevisions,
  applyRevisions,
  collectRevisions,
  rejectAllRevisions,
  type TrackChangesStorage,
} from '../src/renderer/editor/revisions'

const editors = new Set<Editor>()
afterEach(() => {
  for (const editor of editors) editor.destroy()
  editors.clear()
})

async function open(bodyXml: string, tracked = false) {
  const parsed = await parseDocx(await buildDocx({ bodyXml }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks),
  })
  editors.add(editor)
  if (tracked) {
    const storage = editor.storage.trackChanges as TrackChangesStorage
    storage.enabled = true
    storage.author = 'Reviewer'
  }
  return { editor, parsed }
}

async function saved(editor: Editor, parsed: Awaited<ReturnType<typeof parseDocx>>) {
  const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
  const bytes = await saveDocx(parsed, plan.saveBlocks)
  const zip = await JSZip.loadAsync(bytes)
  const xml = await zip.file('word/document.xml')!.async('string')
  return { xml: xml.slice(xml.indexOf('<w:body>')), bytes }
}

const P = (text: string, pPr = '') =>
  `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r><w:t>${text}</w:t></w:r></w:p>`

/** the body as lines: a paragraph's text, or a table's rows as `a|b` */
const linesOf = (editor: Editor) => {
  const lines: string[] = []
  editor.state.doc.forEach((node) => {
    if (node.type.name !== 'docTable') {
      lines.push(node.textContent)
      return
    }
    node.forEach((row) => {
      const cells: string[] = []
      row.forEach((cell) => cells.push(cell.textContent))
      lines.push(cells.join('|'))
    })
  })
  return lines
}

const selectBlocks = (editor: Editor, count: number) => {
  let end = 0
  for (let i = 0; i < count; i++) end += editor.state.doc.child(i).nodeSize
  editor.view.dispatch(
    editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 1, end - 1)),
  )
}

const comma = { cols: 2, separator: { kind: 'text', text: ',' }, autoFit: 'window' } as const
const cellParagraphs = (xml: string) =>
  [...xml.matchAll(/<w:tc>[\s\S]*?<\/w:tc>/g)].map((m) => /<w:p>[\s\S]*<\/w:p>/.exec(m[0])![0])
const kinds = (editor: Editor) => collectRevisions(editor.state.doc).map((r) => r.kind)

describe('Convert Text to Table keeps paragraph formatting', () => {
  it('writes each cell paragraph formatted as the paragraph it was cut from', async () => {
    const { editor, parsed } = await open(
      P('a,b', '<w:jc w:val="center"/><w:ind w:left="720"/>') +
        P('c,d', '<w:spacing w:before="240" w:after="120"/>') +
        P('e,f') +
        P('end'),
    )
    selectBlocks(editor, 3)
    expect(convertTextToTable(editor, comma, 600)).toBe(true)
    expect(linesOf(editor)).toEqual(['a|b', 'c|d', 'e|f', 'end'])
    const cells = cellParagraphs((await saved(editor, parsed)).xml)
    expect(cells).toHaveLength(6)
    for (const cell of cells.slice(0, 2)) {
      expect(cell).toContain('<w:jc w:val="center"/>')
      expect(cell).toMatch(/<w:ind [^>]*w:left="720"/)
    }
    for (const cell of cells.slice(2, 4)) {
      expect(cell).toMatch(/<w:spacing [^>]*w:before="240"/)
      expect(cell).toMatch(/<w:spacing [^>]*w:after="120"/)
    }
    // a paragraph that sets nothing takes the new table's own cell paragraph
    for (const cell of cells.slice(4)) {
      expect(cell).toMatch(/<w:spacing [^>]*w:after="0"/)
      expect(cell).not.toContain('<w:jc')
    }
  })

  it('keeps the style of a heading and what a paragraph is named by', async () => {
    const { editor, parsed } = await open(
      P('Title,one', '<w:pStyle w:val="Heading1"/>') +
        '<w:p><w:bookmarkStart w:id="1" w:name="Total"/><w:r><w:t>x,y</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>' +
        P('end'),
    )
    selectBlocks(editor, 2)
    expect(convertTextToTable(editor, comma, 600)).toBe(true)
    const { xml } = await saved(editor, parsed)
    const cells = cellParagraphs(xml)
    expect(cells[0]).toContain('<w:pStyle w:val="Heading1"/>')
    expect(cells[1]).toContain('<w:pStyle w:val="Heading1"/>')
    expect(xml.match(/w:name="Total"/g)).toHaveLength(1)
    expect(cells[2]).toContain('w:name="Total"')
  })
})

describe('Convert under Track Changes', () => {
  const kindsOf = (editor: Editor) => [...new Set(kinds(editor))].sort()
  const reopen = async (bytes: Uint8Array) => {
    const reopened = await parseDocx(bytes)
    const again = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(reopened.blocks),
    })
    editors.add(again)
    return again
  }

  it('records text to table as the text deleted and the table rows inserted', async () => {
    const { editor, parsed } = await open(P('a,b') + P('c,d') + P('end'), true)
    selectBlocks(editor, 2)
    expect(convertTextToTable(editor, comma, 600)).toBe(true)
    // the text of an inserted row is part of the row's change
    expect(kindsOf(editor)).toEqual(['del', 'paraMarkDel', 'rowIns'])
    // the struck text stays in its order, in front of the table
    expect(linesOf(editor)).toEqual(['a,b', 'c,d', 'a|b', 'c|d', 'end'])
    // the caret is in the new table
    expect(editor.state.selection.$from.parent.textContent).toBe('a')
    const { xml, bytes } = await saved(editor, parsed)
    // run-level and row-level records only: no block wrapped in w:ins / w:del
    expect(xml).not.toMatch(/<w:(?:ins|del) [^>]*><w:(?:p|tbl)[ >]/)
    expect(xml.match(/<w:trPr><w:ins /g)).toHaveLength(2)
    expect(xml.match(/<w:delText[^>]*>[^<]+<\/w:delText>/g)).toHaveLength(2)
    expect(xml.match(/<w:pPr><w:rPr><w:del /g)).toHaveLength(2)

    const rejected = await reopen(bytes)
    expect(kindsOf(rejected)).toEqual(['del', 'paraMarkDel', 'rowIns'])
    rejectAllRevisions(rejected)
    expect(linesOf(rejected)).toEqual(['a,b', 'c,d', 'end'])
    expect(kinds(rejected)).toEqual([])
    acceptAllRevisions(editor)
    expect(linesOf(editor)).toEqual(['a|b', 'c|d', 'end'])
    expect(kinds(editor)).toEqual([])
  })

  it('drops text the same reviser had just inserted instead of striking it', async () => {
    const { editor } = await open(P('a,b') + P('end'), true)
    editor.commands.insertContentAt(4, 'X')
    expect(kinds(editor)).toEqual(['ins'])
    selectBlocks(editor, 1)
    expect(convertTextToTable(editor, comma, 600)).toBe(true)
    expect(linesOf(editor)).toEqual(['a,b', 'a|bX', 'end'])
    rejectAllRevisions(editor)
    expect(linesOf(editor)).toEqual(['a,b', 'end'])
  })

  const BOOKMARKED =
    '<w:p><w:bookmarkStart w:id="1" w:name="Total"/><w:r><w:t>x,y</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>'
  const namedTotal = (editor: Editor) => {
    const named: string[] = []
    editor.state.doc.descendants((node) => {
      if ((node.attrs.bookmarks as string[] | null)?.includes('Total'))
        named.push(`${node.type.name}:${node.textContent}`)
    })
    return named
  }
  it('a tracked text to table carries the bookmark to the cell: once in the file, kept by accept and by reject', async () => {
    const { editor, parsed } = await open(BOOKMARKED + P('end'), true)
    selectBlocks(editor, 1)
    expect(convertTextToTable(editor, comma, 600)).toBe(true)
    // the struck paragraph is first, the table's first cell second
    expect(linesOf(editor)).toEqual(['x,y', 'x|y', 'end'])
    expect(namedTotal(editor)).toEqual(['docParagraph:x'])
    const { xml, bytes } = await saved(editor, parsed)
    expect(xml.match(/w:name="Total"/g)).toHaveLength(1)
    expect(xml.indexOf('w:name="Total"')).toBeGreaterThan(xml.indexOf('<w:tbl>'))

    const accepted = await reopen(bytes)
    acceptAllRevisions(accepted)
    expect(linesOf(accepted)).toEqual(['x|y', 'end'])
    expect(namedTotal(accepted)).toEqual(['docParagraph:x'])

    rejectAllRevisions(editor)
    expect(linesOf(editor)).toEqual(['x,y', 'end'])
    expect(namedTotal(editor)).toEqual(['docParagraph:x,y'])
    expect((await saved(editor, parsed)).xml.match(/w:name="Total"/g)).toHaveLength(1)
  })

  it('the deletion rejected before the insertion: the name waits on the paragraph until the table goes', async () => {
    const { editor } = await open(BOOKMARKED + P('end'), true)
    selectBlocks(editor, 1)
    convertTextToTable(editor, comma, 600)
    const revisions = () => collectRevisions(editor.state.doc)
    applyRevisions(
      editor,
      revisions().filter((r) => r.kind === 'del' || r.kind === 'paraMarkDel'),
      'reject',
    )
    expect(namedTotal(editor)).toEqual(['docParagraph:x'])
    rejectAllRevisions(editor)
    expect(linesOf(editor)).toEqual(['x,y', 'end'])
    expect(namedTotal(editor)).toEqual(['docParagraph:x,y'])
  })

  it('a bookmark that covers a selection goes to the table whole: no end is left in the struck text', async () => {
    const { editor, parsed } = await open(P('x,y') + P('end'), true)
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 1, 4)))
    addBookmark(editor, 'Span')
    selectBlocks(editor, 1)
    convertTextToTable(editor, comma, 600)
    const { xml } = await saved(editor, parsed)
    expect(xml.match(/<w:bookmarkStart [^>]*w:name="Span"/g)).toHaveLength(1)
    expect(xml.match(/<w:bookmarkEnd /g)).toHaveLength(1)
    expect(xml.indexOf('w:name="Span"')).toBeGreaterThan(xml.indexOf('<w:tbl>'))
    acceptAllRevisions(editor)
    expect(bookmarkRange(editor.state.doc, 'Span')).not.toBeNull()
  })

  it('a tracked table to text carries a cell bookmark to the text, and a rejection gives it back', async () => {
    const table = TABLE_XML.replace(
      '<w:p><w:r><w:t>A1</w:t></w:r></w:p>',
      '<w:p><w:bookmarkStart w:id="1" w:name="Total"/><w:r><w:t>A1</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>',
    )
    expect(table).toContain('w:name="Total"')
    const { editor, parsed } = await open(table + P('end'), true)
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 4)))
    expect(
      convertTableToText(editor, { separator: { kind: 'text', text: '\t' }, nested: true }),
    ).toBe(true)
    expect(namedTotal(editor)).toEqual(['docParagraph:A1\tB1'])
    const { xml, bytes } = await saved(editor, parsed)
    expect(xml.match(/w:name="Total"/g)).toHaveLength(1)

    const accepted = await reopen(bytes)
    acceptAllRevisions(accepted)
    expect(linesOf(accepted)).toEqual(['A1\tB1', 'A2\tB2', 'end'])
    expect(namedTotal(accepted)).toEqual(['docParagraph:A1\tB1'])

    rejectAllRevisions(editor)
    expect(linesOf(editor)).toEqual(['A1|B1', 'A2|B2', 'end'])
    expect(namedTotal(editor)).toEqual(['docParagraph:A1'])
  })

  it('records table to text as the table rows deleted and the text inserted', async () => {
    const { editor, parsed } = await open(TABLE_XML + P('end'), true)
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 4)))
    expect(
      convertTableToText(editor, { separator: { kind: 'text', text: '\t' }, nested: true }),
    ).toBe(true)
    expect(kindsOf(editor)).toEqual(['ins', 'paraMarkIns', 'rowDel'])
    expect(linesOf(editor)).toEqual(['A1|B1', 'A2|B2', 'A1\tB1', 'A2\tB2', 'end'])
    const { xml, bytes } = await saved(editor, parsed)
    expect(xml).not.toMatch(/<w:(?:ins|del) [^>]*><w:(?:p|tbl)[ >]/)
    expect(xml.match(/<w:trPr><w:del /g)).toHaveLength(2)
    const accepted = await reopen(bytes)
    acceptAllRevisions(accepted)
    expect(linesOf(accepted)).toEqual(['A1\tB1', 'A2\tB2', 'end'])
    expect(kinds(accepted)).toEqual([])
    rejectAllRevisions(editor)
    expect(linesOf(editor)).toEqual(['A1|B1', 'A2|B2', 'end'])
    expect(kinds(editor)).toEqual([])
  })
})

describe('whole blocks deleted while tracking', () => {
  it('come back as deletions in the order they stood in', async () => {
    const { editor } = await open(P('one') + P('two') + P('end'), true)
    const size = editor.state.doc.child(0).nodeSize + editor.state.doc.child(1).nodeSize
    editor.view.dispatch(
      editor.state.tr.replaceWith(0, size, editor.schema.nodes.docParagraph.create()),
    )
    expect(linesOf(editor).slice(0, 2)).toEqual(['one', 'two'])
    expect(kinds(editor).filter((kind) => kind === 'blockDel')).toHaveLength(2)
  })
})
