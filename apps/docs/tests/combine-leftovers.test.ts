import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { compareDocuments, parseDocx, readSections, saveDocx } from '@genoffice/docx-engine'
import {
  buildDocx,
  type BuildDocxOptions,
} from '../../../packages/docx-engine/tests/helpers/build-docx'
import { applyTrackedChangesXml } from '../../../packages/docx-engine/src/revision-xml'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  blocksToPmDoc,
  pmDocOptions,
  pmDocToSavePlan,
  type PmNode,
} from '../src/renderer/editor/convert'
import { combineDocuments, type CombineOptions } from '../src/renderer/editor/combine'
import {
  acceptAllRevisions,
  collectRevisions,
  rejectAllRevisions,
} from '../src/renderer/editor/revisions'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const AUTHOR = 'Rae Visor'
const DATE = '2026-11-22T10:00:00Z'
const BEA = 'w:author="Bea Revise" w:date="2026-11-02T09:00:00Z"'
const OTHER_PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPj/HwADBwIAMCbHYQAAAABJRU5ErkJggg=='

const r = (text: string, rPr = '') =>
  `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`
const p = (content: string, pPr = '') =>
  `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${content.startsWith('<') ? content : r(content)}</w:p>`
const tc = (content: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${content.startsWith('<') ? content : p(content)}</w:tc>`
const tr = (cells: string[], trPr = '') =>
  `<w:tr>${trPr ? `<w:trPr>${trPr}</w:trPr>` : ''}${cells.map((c) => (c.startsWith('<w:tc>') ? c : tc(c))).join('')}</w:tr>`
const tbl = (rows: string[]) =>
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>${rows.join('')}</w:tbl>`
const picture =
  '<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Picture 1"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="Picture 1"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId10"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
  '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
const math = (text: string) => `<m:oMath><m:r><m:t>${text}</m:t></m:r></m:oMath>`

const INTRO = p('This agreement is made between the supplier and the customer.')
const SIGNED = p('Signed by both parties.')

const editors: Editor[] = []
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy()
})

type Side = string | (BuildDocxOptions & { png?: string })

async function load(side: Side) {
  const options = typeof side === 'string' ? { bodyXml: side } : side
  let bytes = await buildDocx(options)
  if (typeof side !== 'string' && side.png) {
    const zip = await JSZip.loadAsync(bytes)
    zip.file('word/media/image1.png', side.png, { base64: true })
    bytes = await zip.generateAsync({ type: 'uint8array' })
  }
  const parsed = await parseDocx(bytes)
  const json = blocksToPmDoc(parsed.blocks, readSections(parsed), pmDocOptions(parsed))
  return { parsed, json }
}

async function combine(
  originalSide: Side,
  revisedSide: Side,
  options: Partial<CombineOptions> = {},
) {
  const original = await load(originalSide)
  const revised = await load(revisedSide)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    editorProps: { handleScrollToSelection: () => true },
    content: original.json,
  })
  editors.push(editor)
  const summary = combineDocuments(editor, revised.json.content ?? [], {
    author: AUTHOR,
    date: DATE,
    formatting: true,
    ...options,
  })
  const save = async () => {
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, original.parsed.blocks)
    const bytes = await saveDocx(original.parsed, plan.saveBlocks, {
      ...(Object.keys(plan.newPictures).length ? { newPictures: plan.newPictures } : {}),
    })
    const xml = (await parseDocx(bytes)).internal.documentXml
    return { bytes, xml, body: /<w:body>([\s\S]*)<w:sectPr/.exec(xml)?.[1] ?? '' }
  }
  return { editor, summary, save }
}

const revisions = (editor: Editor) =>
  collectRevisions(editor.state.doc).map((rev) => `${rev.kind}|${rev.author}`)
const paragraphs = (editor: Editor): PmNode[] => {
  const out: PmNode[] = []
  editor.state.doc.descendants((node) => {
    if (!node.isTextblock) return true
    out.push(node.toJSON() as PmNode)
    return false
  })
  return out
}
const rowCount = (xml: string) => (xml.match(/<w:tr>/g) ?? []).length
function resolved(xml: string, mode: 'accept' | 'reject'): string {
  const out = applyTrackedChangesXml(xml, mode)
  expect(out.remaining).toBe(0)
  return out.xml
}

describe('Formatting compared into the open document', () => {
  it('a paragraph whose formatting alone differs takes the revised formatting as a tracked change', async () => {
    const { editor, summary, save } = await combine(
      INTRO + p('Same words here.', '<w:keepNext/><w:jc w:val="left"/>') + SIGNED,
      INTRO + p('Same words here.', '<w:jc w:val="center"/>') + SIGNED,
    )
    expect(summary).toMatchObject({ formatChanges: 1, insertions: 0, deletions: 0 })
    expect(revisions(editor)).toEqual([`pPrChange|${AUTHOR}`])
    expect(editor.state.doc.child(1).attrs).toMatchObject({ align: 'center', keepNext: null })
    const { body } = await save()
    expect(body).toMatch(
      new RegExp(
        `<w:pPr><w:jc w:val="center"/><w:pPrChange w:id="\\d+" w:author="${AUTHOR}" w:date="${DATE}"><w:pPr><w:keepNext/><w:jc w:val="left"/></w:pPr></w:pPrChange></w:pPr>`,
      ),
    )
    rejectAllRevisions(editor)
    expect(editor.state.doc.child(1).attrs).toMatchObject({ align: 'left', keepNext: true })
    expect(revisions(editor)).toEqual([])
  })

  it('a table cell paragraph is compared too, and accepting keeps the revised formatting', async () => {
    const { editor, summary, save } = await combine(
      INTRO + tbl([tr(['Item', 'Price'])]) + SIGNED,
      INTRO + tbl([tr(['Item', tc(p('Price', '<w:jc w:val="right"/>'))])]) + SIGNED,
    )
    expect(summary.formatChanges).toBe(1)
    expect(revisions(editor)).toEqual([`pPrChange|${AUTHOR}`])
    expect((await save()).body).toMatch(
      /<w:tc>.*?<w:pPr><w:jc w:val="right"\/><w:pPrChange [^>]*><w:pPr><\/w:pPr><\/w:pPrChange><\/w:pPr><w:r><w:t[^>]*>Price</,
    )
    acceptAllRevisions(editor)
    expect(paragraphs(editor)[2]!.attrs!.align).toBe('right')
    expect((await save()).body).not.toContain('w:pPrChange')
  })

  it('shown in the revised document, the paragraph keeps its formatting and records the original', async () => {
    const { editor, summary } = await combine(
      INTRO + p('Same words here.', '<w:jc w:val="center"/>') + SIGNED,
      INTRO + p('Same words here.', '<w:jc w:val="left"/>') + SIGNED,
      { openIsRevised: true },
    )
    expect(summary.formatChanges).toBe(1)
    expect(editor.state.doc.child(1).attrs.align).toBe('center')
    rejectAllRevisions(editor)
    expect(editor.state.doc.child(1).attrs.align).toBe('left')
  })

  it('text whose formatting alone differs is a tracked formatting change that rejects to the original', async () => {
    const { editor, summary, save } = await combine(
      INTRO + p('This line turns bold.') + SIGNED,
      INTRO + p(r('This line ') + r('turns bold', '<w:b/>') + r('.')) + SIGNED,
    )
    expect(summary).toMatchObject({ formatChanges: 1, insertions: 0, deletions: 0 })
    expect(revisions(editor)).toEqual([`rPrChange|${AUTHOR}`])
    expect((await save()).body).toMatch(
      new RegExp(
        `<w:rPr><w:b/><w:bCs/><w:rPrChange w:id="\\d+" w:author="${AUTHOR}" w:date="${DATE}">.*?</w:rPr><w:t[^>]*>turns bold<`,
      ),
    )
    rejectAllRevisions(editor)
    expect((await save()).body).not.toContain('<w:b/>')
  })

  it('with Formatting off, or between identical documents, nothing is recorded', async () => {
    const off = await combine(
      INTRO + p('Same words here.', '<w:jc w:val="left"/>') + tbl([tr(['Item', 'Price'])]),
      INTRO +
        p(r('Same words here.', '<w:b/>'), '<w:jc w:val="center"/>') +
        tbl([tr(['Item', 'Price'])]),
      { formatting: false },
    )
    expect(off.summary).toMatchObject({ formatChanges: 0 })
    expect(revisions(off.editor)).toEqual([])
    const xml =
      INTRO +
      p(
        r('Styled ', '<w:b/><w:color w:val="FF0000"/>') + r('words.'),
        '<w:pStyle w:val="Heading1"/><w:spacing w:before="120"/>',
      ) +
      tbl([tr(['Item', tc(p('Price', '<w:jc w:val="right"/>'))])]) +
      SIGNED
    const same = await combine(xml, xml)
    expect(same.summary).toMatchObject({ formatChanges: 0, insertions: 0, deletions: 0 })
    expect(revisions(same.editor)).toEqual([])
  })

  it("Combine brings the other document's tracked formatting changes in under their author", async () => {
    const { editor, summary, save } = await combine(
      INTRO + p('This line turns bold.') + p('Centred later.') + SIGNED,
      INTRO +
        p(
          r('This line ') +
            `<w:r><w:rPr><w:b/><w:rPrChange w:id="3" ${BEA}><w:rPr/></w:rPrChange></w:rPr><w:t xml:space="preserve">turns bold</w:t></w:r>` +
            r('.'),
        ) +
        p(
          'Centred later.',
          `<w:jc w:val="center"/><w:pPrChange w:id="4" ${BEA}><w:pPr/></w:pPrChange>`,
        ) +
        SIGNED,
      { keepRevisions: true },
    )
    expect(summary).toMatchObject({ formatChanges: 0, carried: 2 })
    expect(revisions(editor)).toEqual(['rPrChange|Bea Revise', 'pPrChange|Bea Revise'])
    const { body } = await save()
    expect(body).toMatch(/<w:rPrChange w:id="\d+" w:author="Bea Revise"/)
    expect(body).toMatch(/<w:pPrChange w:id="\d+" w:author="Bea Revise"/)
    rejectAllRevisions(editor)
    expect(editor.state.doc.child(2).attrs.align).toBeNull()
  })
})

describe('Table rows compared into the open document', () => {
  const HEAD = tr(['Item', 'Price'])
  const BOLT = tr(['Bolt', 'ten dollars'])
  const NUT = tr(['Nut', 'two dollars'])

  it('a row only the revised table has comes in as a tracked row insertion', async () => {
    const { editor, summary, save } = await combine(
      INTRO + tbl([HEAD, BOLT]) + SIGNED,
      INTRO + tbl([HEAD, BOLT, NUT]) + SIGNED,
    )
    expect(summary).toMatchObject({ insertions: 1, deletions: 0, skipped: 0 })
    expect(revisions(editor)).toEqual([`rowIns|${AUTHOR}`])
    const { body } = await save()
    expect(body).toMatch(
      new RegExp(
        `<w:tr><w:trPr><w:ins w:id="\\d+" w:author="${AUTHOR}" w:date="${DATE}"/></w:trPr>`,
      ),
    )
    expect(rowCount(resolved(body, 'accept'))).toBe(3)
    expect(rowCount(resolved(body, 'reject'))).toBe(2)
    rejectAllRevisions(editor)
    expect(editor.state.doc.child(1).childCount).toBe(2)
  })

  it('a row only the open table has is a tracked row deletion, and a cell beside it is diffed', async () => {
    const { editor, summary, save } = await combine(
      INTRO + tbl([HEAD, NUT, BOLT]) + SIGNED,
      INTRO + tbl([HEAD, tr(['Bolt', 'twelve dollars'])]) + SIGNED,
    )
    expect(summary).toMatchObject({ insertions: 1, deletions: 2, skipped: 0 })
    expect(revisions(editor)).toEqual([`rowDel|${AUTHOR}`, `del|${AUTHOR}`, `ins|${AUTHOR}`])
    const { body } = await save()
    expect(body).toMatch(new RegExp(`<w:trPr><w:del w:id="\\d+" w:author="${AUTHOR}"`))
    expect(rowCount(resolved(body, 'accept'))).toBe(2)
    expect(rowCount(resolved(body, 'reject'))).toBe(3)
    acceptAllRevisions(editor)
    expect(editor.state.doc.child(1).childCount).toBe(2)
    expect(editor.state.doc.child(1).textContent).toBe('ItemPriceBolttwelve dollars')
  })

  it("Combine brings in a row the other document tracks, under that row's author", async () => {
    const tracked = tr(
      [
        tc(p(`<w:ins w:id="5" ${BEA}>${r('Nut')}</w:ins>`)),
        tc(p(`<w:ins w:id="6" ${BEA}>${r('two dollars')}</w:ins>`)),
      ],
      `<w:ins w:id="4" ${BEA}/>`,
    )
    const { editor, summary, save } = await combine(
      INTRO + tbl([HEAD, BOLT]) + SIGNED,
      INTRO + tbl([HEAD, tracked, BOLT]) + SIGNED,
      { keepRevisions: true },
    )
    expect(summary.insertions).toBe(0)
    expect(revisions(editor)).toEqual(['rowIns|Bea Revise'])
    const { body } = await save()
    expect(body).toMatch(/<w:trPr><w:ins w:id="\d+" w:author="Bea Revise"/)
    expect(body).toMatch(/<w:ins w:id="\d+" w:author="Bea Revise"[^>]*><w:r><w:t[^>]*>Nut</)
    expect(body).not.toContain(AUTHOR)
  })

  it('Combine brings in a cell the other document tracks as inserted, on a row both tables have', async () => {
    const cell = `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:cellIns w:id="7" ${BEA}/></w:tcPr>${p('ten dollars')}</w:tc>`
    const { editor, summary, save } = await combine(
      INTRO + tbl([HEAD, BOLT]) + SIGNED,
      INTRO + tbl([HEAD, tr(['Bolt', cell])]) + SIGNED,
      { keepRevisions: true },
    )
    expect(summary).toMatchObject({ insertions: 0, deletions: 0, carried: 1 })
    expect(revisions(editor)).toEqual(['cellIns|Bea Revise'])
    const { body } = await save()
    expect(body).toMatch(
      /<w:cellIns w:id="\d+" w:author="Bea Revise" w:date="2026-11-02T09:00:00Z"\/>/,
    )
    expect(body).not.toContain(AUTHOR)
    // without Combine the other document reads as accepted: nothing to bring in
    const compared = await combine(
      INTRO + tbl([HEAD, BOLT]) + SIGNED,
      INTRO + tbl([HEAD, tr(['Bolt', cell])]) + SIGNED,
    )
    expect(revisions(compared.editor)).toEqual([])
  })

  const WIDE = [tr(['Item', 'Stock', 'Price']), tr(['Bolt', 'many', 'ten dollars'])]
  const cellsPerRow = (xml: string) =>
    [...xml.matchAll(/<w:tr>[\s\S]*?<\/w:tr>/g)].map((m) => (m[0].match(/<w:tc>/g) ?? []).length)
  const tableRows = (editor: Editor) => {
    const rows: string[][] = []
    editor.state.doc.child(1).forEach((row) => {
      const cells: string[] = []
      row.forEach((cell) => cells.push(cell.textContent))
      rows.push(cells)
    })
    return rows
  }

  it('a column only the revised table has comes in as tracked inserted cells', async () => {
    const { editor, summary, save } = await combine(
      INTRO + tbl([HEAD, BOLT]) + SIGNED,
      INTRO + tbl(WIDE) + SIGNED,
    )
    expect(summary).toMatchObject({ insertions: 1, deletions: 0, skipped: 0 })
    expect(revisions(editor)).toEqual([`cellIns|${AUTHOR}`, `cellIns|${AUTHOR}`])
    expect(tableRows(editor)).toEqual([
      ['Item', 'Stock', 'Price'],
      ['Bolt', 'many', 'ten dollars'],
    ])
    const { body } = await save()
    expect(body.match(/<w:cellIns w:id="\d+" w:author="Rae Visor"/g)).toHaveLength(2)
    expect(body.match(/<w:gridCol /g)).toHaveLength(3)
    expect(cellsPerRow(resolved(body, 'accept'))).toEqual([3, 3])
    expect(cellsPerRow(resolved(body, 'reject'))).toEqual([2, 2])
    rejectAllRevisions(editor)
    expect(tableRows(editor)).toEqual([
      ['Item', 'Price'],
      ['Bolt', 'ten dollars'],
    ])
    expect(revisions(editor)).toEqual([])
  })

  it('a column only the open table has is tracked deleted cells, and a cell beside it is diffed', async () => {
    const { editor, summary, save } = await combine(
      INTRO + tbl(WIDE) + SIGNED,
      INTRO + tbl([HEAD, tr(['Bolt', 'twelve dollars'])]) + SIGNED,
    )
    expect(summary).toMatchObject({ insertions: 1, deletions: 2, skipped: 0 })
    expect(revisions(editor)).toEqual([
      `cellDel|${AUTHOR}`,
      `cellDel|${AUTHOR}`,
      `del|${AUTHOR}`,
      `ins|${AUTHOR}`,
    ])
    const { body } = await save()
    expect(body.match(/<w:cellDel w:id="\d+" w:author="Rae Visor"/g)).toHaveLength(2)
    expect(cellsPerRow(resolved(body, 'accept'))).toEqual([2, 2])
    expect(cellsPerRow(resolved(body, 'reject'))).toEqual([3, 3])
    acceptAllRevisions(editor)
    expect(tableRows(editor)).toEqual([
      ['Item', 'Price'],
      ['Bolt', 'twelve dollars'],
    ])
  })

  it('a column and rows changed together: every row gets a cell for every column', async () => {
    const SCREW = tr(['Screw', 'three dollars'])
    const revised = [
      tr(['Item', 'Price', 'Stock']),
      tr(['Bolt', 'ten dollars', 'many']),
      tr(['Screw', 'three dollars', 'some']),
      tr(['Washer', 'one dollar', 'few']),
    ]
    const both = () =>
      combine(INTRO + tbl([HEAD, NUT, BOLT, SCREW]) + SIGNED, INTRO + tbl(revised) + SIGNED)
    const accepted = await both()
    expect(accepted.summary.skipped).toBe(0)
    expect(tableRows(accepted.editor).map((row) => row.length)).toEqual([3, 3, 3, 3, 3])
    acceptAllRevisions(accepted.editor)
    expect(tableRows(accepted.editor)).toEqual([
      ['Item', 'Price', 'Stock'],
      ['Bolt', 'ten dollars', 'many'],
      ['Screw', 'three dollars', 'some'],
      ['Washer', 'one dollar', 'few'],
    ])
    expect(revisions(accepted.editor)).toEqual([])
    const rejected = await both()
    rejectAllRevisions(rejected.editor)
    expect(tableRows(rejected.editor)).toEqual([
      ['Item', 'Price'],
      ['Nut', 'two dollars'],
      ['Bolt', 'ten dollars'],
      ['Screw', 'three dollars'],
    ])
    expect(revisions(rejected.editor)).toEqual([])
  })

  it('a row that goes is struck whole when a column comes into the middle of it', async () => {
    const SCREW = tr(['Screw', 'three dollars'])
    const { editor, summary } = await combine(
      INTRO + tbl([HEAD, NUT, BOLT, SCREW]) + SIGNED,
      INTRO +
        tbl([
          tr(['Item', 'Stock', 'Price']),
          tr(['Bolt', 'many', 'ten dollars']),
          tr(['Screw', 'some', 'three dollars']),
        ]) +
        SIGNED,
    )
    expect(summary.skipped).toBe(0)
    const struck: Array<[string, boolean]> = []
    editor.state.doc
      .child(1)
      .child(1)
      .descendants((node) => {
        if (node.isText) struck.push([node.text!, node.marks.some((m) => m.type.name === 'del')])
      })
    expect(struck).toEqual([
      ['Nut', true],
      ['two dollars', true],
    ])
    expect(tableRows(editor)[1]).toEqual(['Nut', '', 'two dollars'])
    rejectAllRevisions(editor)
    expect(tableRows(editor)).toEqual([
      ['Item', 'Price'],
      ['Nut', 'two dollars'],
      ['Bolt', 'ten dollars'],
      ['Screw', 'three dollars'],
    ])
  })

  it('a table with merged cells whose columns changed is still left alone and counted', async () => {
    const merged =
      '<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:gridSpan w:val="2"/></w:tcPr>' +
      `${p('Parts')}</w:tc></w:tr>`
    const { editor, summary } = await combine(
      INTRO + tbl([merged, HEAD, BOLT]) + SIGNED,
      INTRO + tbl([tr(['Parts', '', '']), ...WIDE]) + SIGNED,
    )
    expect(summary.skipped).toBe(1)
    expect(revisions(editor)).toEqual([])
  })

  it('a table only one document has comes in, or goes, as a whole tracked table', async () => {
    const added = await combine(INTRO + SIGNED, INTRO + tbl([HEAD, BOLT]) + SIGNED)
    expect(added.summary).toMatchObject({ insertions: 1, skipped: 0 })
    expect(added.editor.state.doc.child(1).attrs.blockRevision).toMatchObject({
      kind: 'ins',
      author: AUTHOR,
    })
    rejectAllRevisions(added.editor)
    expect(added.editor.state.doc.childCount).toBe(2)

    const removed = await combine(INTRO + tbl([HEAD, BOLT]) + SIGNED, INTRO + SIGNED)
    expect(removed.summary).toMatchObject({ deletions: 1, skipped: 0 })
    expect(removed.editor.state.doc.child(1).attrs.blockRevision).toMatchObject({
      kind: 'del',
      author: AUTHOR,
    })
    acceptAllRevisions(removed.editor)
    expect(removed.editor.state.doc.childCount).toBe(2)
  })
})

describe('Moves, pictures and equations combined into the open document', () => {
  const CLAUSE = 'A clause that moves to the end.'
  const from = `<w:p><w:pPr><w:rPr><w:moveFrom w:id="7" ${BEA}/></w:rPr></w:pPr><w:moveFromRangeStart w:id="8" ${BEA} w:name="move1"/><w:moveFrom w:id="9" ${BEA}>${r(CLAUSE)}</w:moveFrom><w:moveFromRangeEnd w:id="8"/></w:p>`
  const to = `<w:p><w:pPr><w:rPr><w:moveTo w:id="10" ${BEA}/></w:rPr></w:pPr><w:moveToRangeStart w:id="11" ${BEA} w:name="move1"/><w:moveTo w:id="12" ${BEA}>${r(CLAUSE)}</w:moveTo><w:moveToRangeEnd w:id="11"/></w:p>`

  it('a move the other document tracks stays a move when both its ends come in', async () => {
    const { editor, save } = await combine(INTRO + p(CLAUSE) + SIGNED, INTRO + from + SIGNED + to, {
      keepRevisions: true,
    })
    expect(revisions(editor).filter((kind) => kind.startsWith('move'))).toEqual([
      'moveFrom|Bea Revise',
      'moveTo|Bea Revise',
    ])
    const { body } = await save()
    expect(body).toMatch(/<w:moveFrom w:id="\d+" w:author="Bea Revise"/)
    expect(body).toMatch(/<w:moveTo w:id="\d+" w:author="Bea Revise"/)
    const names = [...body.matchAll(/<w:move(?:From|To)RangeStart [^>]*w:name="([^"]+)"/g)].map(
      (m) => m[1],
    )
    expect(names).toHaveLength(2)
    expect(names[0]).toBe(names[1])
  })

  it('a move with one end only comes in as a plain insertion', async () => {
    const { editor, save } = await combine(INTRO + SIGNED, INTRO + SIGNED + to, {
      keepRevisions: true,
    })
    expect(revisions(editor).some((kind) => kind.startsWith('move'))).toBe(false)
    expect((await save()).body).not.toContain('w:moveTo')
  })

  it('a picture replaced by another is the old one struck and the new one inserted, saved with both bitmaps', async () => {
    const figure = p(r('Logo: ') + picture + r(' as shown.'))
    const { editor, summary, save } = await combine(
      { bodyXml: INTRO + figure + SIGNED, withImage: true },
      { bodyXml: INTRO + figure + SIGNED, withImage: true, png: OTHER_PNG },
    )
    expect(summary).toMatchObject({ insertions: 1, deletions: 1, skipped: 0 })
    expect(revisions(editor)).toEqual([`del|${AUTHOR}`, `ins|${AUTHOR}`])
    const { bytes, body } = await save()
    expect(body).toMatch(/<w:del [^>]*><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:drawing>/)
    expect(body).toMatch(/<w:ins [^>]*><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:drawing>/)
    const zip = await JSZip.loadAsync(bytes)
    const media = Object.keys(zip.files).filter((name) => name.startsWith('word/media/'))
    const bitmaps = await Promise.all(media.map((name) => zip.file(name)!.async('base64')))
    expect(bitmaps).toContain(OTHER_PNG)
    expect(bitmaps).toHaveLength(2)
    rejectAllRevisions(editor)
    expect(revisions(editor)).toEqual([])
    expect((await save()).body.match(/<w:drawing>/g)).toHaveLength(1)
  })

  it('an equation that changed is the old one struck and the new one inserted', async () => {
    const { editor, summary, save } = await combine(
      INTRO + p(r('Area is ') + math('a^2') + r(' here.')) + SIGNED,
      INTRO + p(r('Area is ') + math('a^3') + r(' here.')) + SIGNED,
    )
    expect(summary).toMatchObject({ insertions: 1, deletions: 1, skipped: 0 })
    const { body } = await save()
    expect(body).toMatch(/<w:del [^>]*><m:oMath><m:r><m:t>a\^2<\/m:t><\/m:r><\/m:oMath><\/w:del>/)
    expect(body).toMatch(/<w:ins [^>]*><m:oMath><m:r><m:t>a\^3<\/m:t><\/m:r><\/m:oMath><\/w:ins>/)
    acceptAllRevisions(editor)
    expect((await save()).body.match(/<m:oMath>/g)).toHaveLength(1)
    expect((await save()).body).toContain('a^3')
  })
})

describe('a comparison in a new document: text that moved between table cells', () => {
  const A = 'The supplier shall deliver the goods to the address of the customer.'
  const B = 'Delivery takes place on working days only.'
  const C = 'The customer inspects the goods on arrival.'
  const original = INTRO + tbl([tr([`${A} ${B}`, C])]) + SIGNED
  const revised = INTRO + tbl([tr([A, `${C} ${B}`])]) + SIGNED

  async function opened() {
    const { bytes } = await compareDocuments(
      await parseDocx(await buildDocx({ bodyXml: original })),
      await parseDocx(await buildDocx({ bodyXml: revised })),
      { author: AUTHOR, date: DATE },
    )
    const parsed = await parseDocx(bytes)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      editorProps: { handleScrollToSelection: () => true },
      content: blocksToPmDoc(parsed.blocks, readSections(parsed), pmDocOptions(parsed)),
    })
    editors.push(editor)
    const cells = async () => {
      const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
      const xml = (await parseDocx(await saveDocx(parsed, plan.saveBlocks))).internal.documentXml
      expect(xml).not.toMatch(/<w:move(?:From|To)/)
      return [...xml.matchAll(/<w:tc>[\s\S]*?<\/w:tc>/g)].map((cell) =>
        [...cell[0].matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((t) => t[1]).join(''),
      )
    }
    return { editor, cells }
  }

  it('is listed as a move, and accepting or rejecting it in the editor gives the revised or the original cells', async () => {
    const accepted = await opened()
    expect(revisions(accepted.editor).sort()).toEqual([`moveFrom|${AUTHOR}`, `moveTo|${AUTHOR}`])
    acceptAllRevisions(accepted.editor)
    expect(revisions(accepted.editor)).toEqual([])
    expect(await accepted.cells()).toEqual([A, `${C} ${B}`])

    const rejected = await opened()
    rejectAllRevisions(rejected.editor)
    expect(revisions(rejected.editor)).toEqual([])
    expect(await rejected.cells()).toEqual([`${A} ${B}`, C])
  })
})
