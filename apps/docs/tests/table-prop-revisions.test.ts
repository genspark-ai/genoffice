import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { parseDocx, saveDocx, type Block } from '@genoffice/docx-engine'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import {
  acceptAllRevisions,
  collectRevisions,
  rejectAllRevisions,
  type RevisionRange,
} from '../src/renderer/editor/revisions'
import { revisionGroup } from '../src/renderer/editor/revision-filters'
import { revGroupsOf } from '../src/renderer/editor/margin-annotations'
import {
  buildDocx,
  PROP_REVISION_AT as AT,
  PROP_REVISION_TABLE_XML,
} from '../../../packages/docx-engine/tests/helpers/build-docx'

async function open(
  bodyXml: string,
  options: { withNumbering?: boolean; sectPrExtra?: string } = {},
) {
  const bytes = await buildDocx({ bodyXml, ...options })
  const parsed = await parseDocx(bytes)
  const editor = new Editor({ extensions: editorExtensions, content: blocksToPmDoc(parsed.blocks) })
  return { editor, parsed }
}

async function saveAndReload(editor: Editor, parsed: Awaited<ReturnType<typeof parseDocx>>) {
  const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
  const bytes = await saveDocx(parsed, plan.saveBlocks, {
    trailingSectPrChange: plan.trailingSectPrChange,
  })
  const documentXml = await (
    await JSZip.loadAsync(bytes)
  )
    .file('word/document.xml')!
    .async('string')
  return { reparsed: await parseDocx(bytes), documentXml, plan }
}

const kinds = (editor: Editor) => collectRevisions(editor.state.doc).map((r) => r.kind)

describe('table property revisions in the editor', () => {
  it('enumerates every record as a Formatting revision with author and date', async () => {
    const { editor } = await open(PROP_REVISION_TABLE_XML)
    const revs = collectRevisions(editor.state.doc)
    expect(revs.map((r) => r.kind).sort()).toEqual(
      [
        'cellMerge',
        'tblGridChange',
        'tblPrChange',
        'tblPrExChange',
        'tcPrChange',
        'trPrChange',
      ].sort(),
    )
    for (const r of revs) {
      expect(r.author).toBe('Alice')
      expect(r.date).toBe('2026-01-01T00:00:00Z')
      expect(revisionGroup(r.kind)).toBe('format')
    }
    // the merged-away cell has no node: its record sits on the restart cell (row 2)
    const merge = revs.find((r) => r.kind === 'cellMerge')!
    expect(editor.state.doc.resolve(merge.from + 1).parent.textContent).toBe('A2')
    // balloons: records sharing a node and a label fold into one bubble (table, row, cell, merge)
    expect(revGroupsOf(editor.state.doc).filter((g) => g.kind === 'fmt')).toHaveLength(4)
    editor.destroy()
  })

  it('a structural edit keeps the pending merge record on the merged-away cell', async () => {
    const { editor, parsed } = await open(PROP_REVISION_TABLE_XML)
    const table = editor.state.doc.firstChild!
    editor.commands.command(({ tr }) => {
      tr.setNodeMarkup(0, undefined, { ...table.attrs, tblAlign: 'right' })
      return true
    })
    const { documentXml } = await saveAndReload(editor, parsed)
    expect(documentXml).toContain('<w:jc w:val="right"/>')
    expect(documentXml.match(/<w:cellMerge\b/g)).toHaveLength(1)
    expect(documentXml).toContain('<w:vMerge/><w:cellMerge w:id="26"')
    expect(kinds(editor)).toContain('cellMerge')
    editor.destroy()
  })

  it('a three-row merge rebuilt by a structural edit writes the merge record once', async () => {
    const xml = PROP_REVISION_TABLE_XML.replace(
      '</w:tbl>',
      '<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:vMerge/></w:tcPr><w:p/></w:tc>' +
        '<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>B4</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
    )
    const { editor, parsed } = await open(xml)
    const table = editor.state.doc.firstChild!
    expect(table.child(1).child(0).attrs.rowspan).toBe(3)
    editor.commands.command(({ tr }) => {
      tr.setNodeMarkup(0, undefined, { ...table.attrs, tblAlign: 'right' })
      return true
    })
    const { documentXml } = await saveAndReload(editor, parsed)
    expect(documentXml.match(/<w:cellMerge\b/g)).toHaveLength(1)
    expect(documentXml.match(/<w:vMerge\/>/g)).toHaveLength(2)
    editor.destroy()
  })

  it('accept all keeps the current properties and the saved table carries no records', async () => {
    const { editor, parsed } = await open(PROP_REVISION_TABLE_XML)
    acceptAllRevisions(editor)
    expect(kinds(editor)).toEqual([])
    const { reparsed, documentXml } = await saveAndReload(editor, parsed)
    expect(documentXml).not.toMatch(/PrChange|PrExChange|tblGridChange|cellMerge/)
    expect(documentXml).toContain('<w:jc w:val="center"/>')
    expect(documentXml).toContain('<w:trHeight w:val="600"')
    expect(documentXml).toContain('w:fill="FF0000"')
    expect(documentXml).toContain('<w:tblPrEx><w:tblCellMar>')
    const table = reparsed.blocks[0].table!
    expect(table.rows[2][0].vMerge).toBe('continue')
    expect(table.tblPrChange).toBeUndefined()
    editor.destroy()
  })

  it('reject all restores the previous properties on the page and in the file', async () => {
    const { editor, parsed } = await open(PROP_REVISION_TABLE_XML)
    rejectAllRevisions(editor)
    expect(kinds(editor)).toEqual([])
    const table = editor.state.doc.firstChild!
    expect(table.attrs.tblAlign).toBe('left')
    expect(table.firstChild!.attrs.heightTwips).toBe(300)
    expect(table.firstChild!.firstChild!.attrs.fill).toBeNull()
    // the grid went back to 3000/5000 twips
    expect(table.attrs.colWidthsPct.map((p: number) => Math.round(p))).toEqual([38, 62])
    // the vertical merge was undone: row 3 has its own A cell again
    expect(table.child(1).firstChild!.attrs.rowspan).toBe(1)
    expect(table.child(2).childCount).toBe(2)
    const { reparsed, documentXml } = await saveAndReload(editor, parsed)
    expect(documentXml).not.toMatch(/PrChange|PrExChange|tblGridChange|cellMerge/)
    expect(documentXml).not.toContain('<w:jc')
    expect(documentXml).toContain('<w:trHeight w:val="300"')
    expect(documentXml).not.toContain('w:fill="FF0000"')
    expect(documentXml).not.toContain('<w:tblPrEx')
    const model = reparsed.blocks[0].table!
    // the old grid came back through the px grid (15 twips per px)
    expect(model.colWidthsTwips).toEqual([3000, 4995])
    expect(model.rows[2][0].vMerge).toBeUndefined()
    expect(model.rows.map((r) => r.map((c) => c.paras[0]))).toEqual([
      ['A1', 'B1'],
      ['A2', 'B2'],
      ['', 'B3'],
    ])
    editor.destroy()
  })
})

const NUMBERING_AND_SECTION_XML =
  `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/><w:numberingChange w:id="31" ${AT} w:original="1."/></w:numPr></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>` +
  `<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:sectPrChange w:id="32" ${AT}><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:sectPrChange></w:sectPr></w:pPr></w:p>` +
  '<w:p><w:r><w:t>Tail</w:t></w:r></w:p>'

describe('numbering and section property revisions in the editor', () => {
  const sectPrExtra = `<w:sectPrChange w:id="40" ${AT}><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:sectPrChange>`

  it('lists the numbering change, the section-break change and the trailing section change', async () => {
    const { editor } = await open(NUMBERING_AND_SECTION_XML, { withNumbering: true, sectPrExtra })
    const revs = collectRevisions(editor.state.doc)
    expect(revs.map((r) => r.kind)).toEqual(['numberingChange', 'sectPrChange', 'sectPrChange'])
    // the trailing record rides on the last paragraph
    const last = editor.state.doc.lastChild!
    expect(String(last.attrs.sectPrChange)).toContain('"trailing":true')
    expect(revs[2].from).toBe(editor.state.doc.content.size - last.nodeSize)
    editor.destroy()
  })

  it('accept all strips every record from the saved document', async () => {
    const { editor, parsed } = await open(NUMBERING_AND_SECTION_XML, {
      withNumbering: true,
      sectPrExtra,
    })
    acceptAllRevisions(editor)
    expect(kinds(editor)).toEqual([])
    const { documentXml, plan } = await saveAndReload(editor, parsed)
    expect(plan.trailingSectPrChange).toBe('accept')
    expect(documentXml).not.toMatch(/numberingChange|sectPrChange/)
    expect(documentXml).toContain('<w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr>')
    expect(documentXml).toContain('<w:numId w:val="1"/></w:numPr>')
    editor.destroy()
  })

  it('reject all restores both previous sectPrs; the numbering record just drops', async () => {
    const { editor, parsed } = await open(NUMBERING_AND_SECTION_XML, {
      withNumbering: true,
      sectPrExtra,
    })
    rejectAllRevisions(editor)
    expect(kinds(editor)).toEqual([])
    const { documentXml, plan, reparsed } = await saveAndReload(editor, parsed)
    expect(plan.trailingSectPrChange).toBe('reject')
    expect(documentXml).not.toMatch(/numberingChange|sectPrChange/)
    expect(documentXml).toContain('<w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr>')
    const trailing = reparsed.blocks.find((b: Block) => b.hidden)!.originalXml!
    expect(trailing).toContain('w:w="12240"')
    expect(trailing).not.toContain('w:pgMar')
    editor.destroy()
  })

  it('a last paragraph with its own section record does not host the trailing one', async () => {
    const body =
      '<w:p><w:r><w:t>Body</w:t></w:r></w:p>' +
      `<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:sectPrChange w:id="32" ${AT}><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:sectPrChange></w:sectPr></w:pPr><w:r><w:t>Break</w:t></w:r></w:p>`
    const { editor, parsed } = await open(body, { sectPrExtra })
    const revs = collectRevisions(editor.state.doc)
    expect(revs.map((r) => r.kind)).toEqual(['sectPrChange', 'sectPrChange'])
    expect(String(editor.state.doc.firstChild!.attrs.sectPrChange)).toContain('"trailing":true')
    expect(String(editor.state.doc.lastChild!.attrs.sectPrChange)).not.toContain('trailing')
    acceptAllRevisions(editor)
    const { documentXml, plan } = await saveAndReload(editor, parsed)
    expect(plan.trailingSectPrChange).toBe('accept')
    expect(documentXml).not.toContain('sectPrChange')
    editor.destroy()
  })

  it('a table-only body leaves the trailing record pending instead of accepting it', async () => {
    const { editor, parsed } = await open(PROP_REVISION_TABLE_XML, { sectPrExtra })
    expect(kinds(editor)).not.toContain('sectPrChange')
    acceptAllRevisions(editor)
    const { documentXml, plan } = await saveAndReload(editor, parsed)
    expect(plan.trailingSectPrChange).toBeUndefined()
    expect(documentXml.match(/<w:sectPrChange\b/g)).toHaveLength(1)
    editor.destroy()
  })

  it('an untouched document keeps its records', async () => {
    const { editor, parsed } = await open(NUMBERING_AND_SECTION_XML, {
      withNumbering: true,
      sectPrExtra,
    })
    const { documentXml, plan } = await saveAndReload(editor, parsed)
    expect(plan.trailingSectPrChange).toBeUndefined()
    expect(documentXml.match(/<w:sectPrChange\b/g)).toHaveLength(2)
    expect(documentXml).toContain('<w:numberingChange')
    const again: RevisionRange['kind'][] = kinds(editor)
    expect(again).toHaveLength(3)
    editor.destroy()
  })
})
