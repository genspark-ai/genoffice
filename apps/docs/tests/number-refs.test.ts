/**
 * Cross-references to a numbered paragraph (REF \r, \n, \w): what the dialog inserts,
 * what a Word file's own fields show, and how both follow the numbering. Saved
 * results are read back from document.xml.
 */
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import JSZip from 'jszip'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { afterEach, describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { dirtyFieldParts } from '../../../packages/docx-engine/tests/helpers/dirty-fields'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import {
  insertCrossReference,
  listRefTargets,
  type RefContent,
} from '../src/renderer/editor/cross-reference'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { numberRefSwitches, refreshNumberRefs } from '../src/renderer/editor/number-refs'

const WORDS = { above: 'above', below: 'below', onPage: (n: number) => `on page ${n}` }

const NUMBERING =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:abstractNum w:abstractNumId="0">' +
  '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>' +
  '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="(%2)"/></w:lvl>' +
  '</w:abstractNum>' +
  '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
  '</w:numbering>'

const item = (text: string, ilvl = 0, bookmark?: string) =>
  `<w:p><w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="1"/></w:numPr></w:pPr>` +
  (bookmark ? `<w:bookmarkStart w:id="${bookmark.length}" w:name="${bookmark}"/>` : '') +
  `<w:r><w:t>${text}</w:t></w:r>` +
  (bookmark ? `<w:bookmarkEnd w:id="${bookmark.length}"/>` : '') +
  '</w:p>'

const ref = (instr: string, result: string) =>
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  `<w:r><w:t>${result}</w:t></w:r>` +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>'

const para = (inner: string) => `<w:p>${inner}</w:p>`

const editors: Editor[] = []
afterEach(() => editors.splice(0).forEach((e) => e.destroy()))

async function open(bodyXml: string) {
  const parsed = await parseDocx(await buildDocx({ bodyXml, numberingXml: NUMBERING }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
  })
  editors.push(editor)
  editor.storage.listNumbering.defs = parsed.numbering
  editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
  return { editor, parsed }
}

async function saved(editor: Editor, parsed: Awaited<ReturnType<typeof parseDocx>>) {
  const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
  const bytes = await saveDocx(parsed, plan.saveBlocks)
  const zip = await JSZip.loadAsync(bytes)
  return { bytes, xml: await zip.file('word/document.xml')!.async('string') }
}

/** each REF field of the saved body as [instruction, result] */
const fieldsOf = (xml: string): Array<[string, string]> =>
  [
    ...xml.matchAll(
      /<w:instrText[^>]*>([^<]*)<\/w:instrText>[\s\S]*?w:fldCharType="separate"[\s\S]*?w:fldCharType="end"/g,
    ),
  ].map((m) => [
    m[1].trim().replace(/\s+/g, ' '),
    [...m[0].matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((t) => t[1]).join(''),
  ])

const refTexts = (editor: Editor): string[] => {
  const out: string[] = []
  editor.state.doc.descendants((node) => {
    if (node.isText && node.marks.some((m) => m.type.name === 'refField')) out.push(node.text!)
  })
  return out
}

function caretAtEnd(editor: Editor) {
  const end = editor.state.doc.content.size - 1
  editor.view.dispatch(
    editor.state.tr.setSelection(TextSelection.near(editor.state.doc.resolve(end))),
  )
}

const insertItemBefore = (editor: Editor, pos: number, ilvl = 0) =>
  editor.view.dispatch(
    editor.state.tr.insert(
      pos,
      editor.schema.nodes.docListItem.create(
        { kind: 'ordered', numId: '1', ilvl },
        editor.schema.text('new'),
      ),
    ),
  )

describe('REF switches of a paragraph number', () => {
  it('reads the context, the separator and above/below', () => {
    expect(numberRefSwitches(' REF _Ref1 \\r \\h ')).toEqual({
      context: 'relative',
      aboveBelow: false,
    })
    expect(numberRefSwitches('REF _Ref1 \\h \\n \\* MERGEFORMAT')).toEqual({
      context: 'none',
      aboveBelow: false,
    })
    expect(numberRefSwitches('REF _Ref1 \\w \\p \\d "-" \\h')).toEqual({
      context: 'full',
      separator: '-',
      aboveBelow: true,
    })
    expect(numberRefSwitches('REF _Ref1 \\h')).toBeNull()
    expect(numberRefSwitches('PAGEREF _Ref1 \\r')).toBeNull()
    expect(numberRefSwitches('REF _Ref1 \\r \\* Upper')).toEqual({
      context: 'relative',
      aboveBelow: false,
      cased: 'REF _Ref1 \\r \\* Upper',
    })
    // a format this editor does not apply leaves the result to Word
    expect(numberRefSwitches('REF _Ref1 \\r \\* roman')).toBeNull()
    expect(numberRefSwitches('REF _Ref1 \\r \\# 0')).toBeNull()
    expect(numberRefSwitches('REF _Ref1 \\r \\* DBCHAR')).toBeNull()
  })
})

describe('Cross-reference to a numbered item', () => {
  const BODY =
    item('Scope') +
    item('Terms', 1) +
    item('Fees', 1) +
    item('Notices') +
    para('<w:r><w:t>See </w:t></w:r>')

  it('lists every numbered paragraph with its number', async () => {
    const { editor, parsed } = await open(BODY)
    const targets = listRefTargets(editor, parsed.blocks, { kind: 'numbered' })
    expect(targets.map((t) => [t.text, t.level])).toEqual([
      ['1. Scope', 1],
      ['(a) Terms', 2],
      ['(b) Fees', 2],
      ['2. Notices', 1],
    ])
  })

  it('inserts REF with the switch of each choice, the number as its result, saved without w:dirty', async () => {
    const { editor, parsed } = await open(BODY)
    const insert = (content: RefContent, extra: Record<string, unknown> = {}) => {
      caretAtEnd(editor)
      const target = listRefTargets(editor, parsed.blocks, { kind: 'numbered' })[2]
      expect(
        insertCrossReference(editor, parsed.blocks, {
          target,
          content,
          hyperlink: true,
          includeAboveBelow: false,
          pages: null,
          words: WORDS,
          ...extra,
        }),
      ).toBe(true)
    }
    insert('paraNumber')
    insert('paraNumberNoContext')
    insert('paraNumberFullContext')
    insert('paraNumberFullContext', { separator: '-' })
    insert('paraNumber', { includeAboveBelow: true })
    insert('text')
    const out = await saved(editor, parsed)
    const name = /w:name="(_Ref\d+)"/.exec(out.xml)![1]
    expect(out.xml.match(/w:name="_Ref/g)).toHaveLength(1)
    expect(fieldsOf(out.xml)).toEqual([
      [`REF ${name} \\r \\h`, '1.(b)'],
      [`REF ${name} \\n \\h`, '(b)'],
      [`REF ${name} \\w \\h`, '1.(b)'],
      [`REF ${name} \\w \\d "-" \\h`, '1-(b)'],
      [`REF ${name} \\r \\p \\h`, '1.(b) above'],
      [`REF ${name} \\h`, 'Fees'],
    ])
    expect(await dirtyFieldParts(out.bytes)).toEqual([])

    const again = await parseDocx(out.bytes)
    const reopened = again.blocks.flatMap((b) => ('runs' in b ? (b.runs ?? []) : []))
    expect(reopened.filter((r) => r.refField === name).map((r) => r.text)).toEqual([
      '1.(b)',
      '(b)',
      '1.(b)',
      '1-(b)',
      '1.(b) above',
      'Fees',
    ])
  })

  it('a reference made inside the list shows the number relative to where it stands', async () => {
    const { editor, parsed } = await open(BODY)
    // the caret in "Terms", pointing at its sibling "Fees"
    let terms = 0
    editor.state.doc.forEach((node, pos) => {
      if (node.textContent === 'Terms') terms = pos
    })
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, terms + 6)),
    )
    const target = listRefTargets(editor, parsed.blocks, { kind: 'numbered' })[2]
    insertCrossReference(editor, parsed.blocks, {
      target,
      content: 'paraNumber',
      hyperlink: false,
      includeAboveBelow: true,
      pages: null,
      words: WORDS,
    })
    expect(refTexts(editor)).toEqual(['(b) below'])
  })

  it('follows the numbering: a paragraph added in front renumbers the reference, an undo takes both back', async () => {
    const { editor, parsed } = await open(BODY)
    caretAtEnd(editor)
    insertCrossReference(editor, parsed.blocks, {
      target: listRefTargets(editor, parsed.blocks, { kind: 'numbered' })[3],
      content: 'paraNumber',
      hyperlink: true,
      includeAboveBelow: false,
      pages: null,
      words: WORDS,
    })
    expect(refTexts(editor)).toEqual(['2'])
    insertItemBefore(editor, 0)
    expect(refTexts(editor)).toEqual(['3'])
    const out = await saved(editor, parsed)
    expect(fieldsOf(out.xml).map((f) => f[1])).toEqual(['3'])
    expect(await dirtyFieldParts(out.bytes)).toEqual([])
    editor.commands.undo()
    expect(refTexts(editor)).toEqual(['2'])
  })
})

describe('paragraph-number references a Word file already holds', () => {
  const FILE =
    item('Scope') +
    item('Terms', 1, '_Ref100') +
    item('Notices', 0, '_Ref200') +
    para(
      '<w:r><w:t>See </w:t></w:r>' +
        ref('REF _Ref100 \\w \\h', '1.(a)') +
        ref('REF _Ref200 \\r \\h', '2') +
        // a number the file never brought up to date
        ref('REF _Ref200 \\n \\h', '7'),
    )

  it('shows the numbers of the file and writes the file back unchanged', async () => {
    const { editor, parsed } = await open(FILE)
    expect(refTexts(editor)).toEqual(['1.(a)', '2', '7'])
    const out = await saved(editor, parsed)
    expect(fieldsOf(out.xml)).toEqual([
      ['REF _Ref100 \\w \\h', '1.(a)'],
      ['REF _Ref200 \\r \\h', '2'],
      ['REF _Ref200 \\n \\h', '7'],
    ])
  })

  it('an edit renumbers the references that agreed with the numbering and leaves the stale one', async () => {
    const { editor, parsed } = await open(FILE)
    insertItemBefore(editor, 0)
    expect(refTexts(editor)).toEqual(['2.(a)', '3', '7'])
    const out = await saved(editor, parsed)
    expect(fieldsOf(out.xml).map((f) => f[1])).toEqual(['2.(a)', '3', '7'])
    expect(await dirtyFieldParts(out.bytes)).toEqual([])
  })

  it('a new reference beside a stale one to the same paragraph follows the numbering on its own', async () => {
    const { editor, parsed } = await open(FILE)
    const end = editor.state.doc.content.size
    editor.view.dispatch(editor.state.tr.insert(end, editor.schema.nodes.docParagraph.create()))
    caretAtEnd(editor)
    insertCrossReference(editor, parsed.blocks, {
      target: listRefTargets(editor, parsed.blocks, { kind: 'numbered' })[2],
      content: 'paraNumberNoContext',
      hyperlink: true,
      includeAboveBelow: false,
      pages: null,
      words: WORDS,
    })
    expect(refTexts(editor)).toEqual(['1.(a)', '2', '7', '2'])
    insertItemBefore(editor, 0)
    expect(refTexts(editor)).toEqual(['2.(a)', '3', '7', '3'])
  })

  it('a change of the list definition reaches the references', async () => {
    const { editor } = await open(FILE)
    const defs = editor.storage.listNumbering.defs
    const def = defs.get('1')!
    editor.storage.listNumbering.defs = new Map(defs).set('1', {
      ...def,
      levels: { ...def.levels, 0: { ...def.levels[0], numFmt: 'upperRoman' } },
    })
    editor.view.dispatch(editor.state.tr)
    expect(refTexts(editor)).toEqual(['I.(a)', 'II', '7'])
  })

  it('Update Field brings the stale one up to date as well', async () => {
    const { editor } = await open(FILE)
    const tr = editor.state.tr
    expect(refreshNumberRefs(tr, editor.storage.listNumbering.defs, { words: WORDS })).toBe(1)
    editor.view.dispatch(tr)
    expect(refTexts(editor)).toEqual(['1.(a)', '2', '2'])
  })
})
