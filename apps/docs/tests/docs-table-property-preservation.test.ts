import { afterEach, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { CellSelection, TableMap } from '@tiptap/pm/tables'
import { deleteRowWithSource, splitCellWithSource } from '../src/renderer/ai/table-ops'
import { computeFormatState } from '../src/renderer/components/ribbon-format-state'
import {
  deleteCells,
  splitCellsInto,
  splitGridCell,
  splitTableAtSelection,
} from '../src/renderer/editor/table-ops'
import { readFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import JSZip from 'jszip'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { createDocsRuntime } from '../src/renderer/runtime/docs-runtime'
import type { DocsQueryResult } from '../src/shared/document-query'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'

const cleanup: Array<() => void> = []
afterEach(() =>
  cleanup
    .splice(0)
    .reverse()
    .forEach((dispose) => dispose()),
)

async function openSynthetic(bodyXml: string, extraStylesXml?: string) {
  const parsed = await parseDocx(await buildDocx({ bodyXml, extraStylesXml }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  const runtime = createDocsRuntime({
    current: () => ({
      editor,
      identity: editor,
      numIds: { bullet: null, ordered: null },
      autoAcceptUntracked: true,
    }),
  })
  cleanup.push(
    () => editor.destroy(),
    () => runtime.dispose(),
  )
  const caller = runtime.connect({ id: 'property-source' })
  caller.context()
  return {
    editor,
    parsed,
    caller,
    async saved() {
      const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
      const bytes = await saveDocx(parsed, plan.saveBlocks)
      const zip = await JSZip.loadAsync(bytes)
      return parseXml(await zip.file('word/document.xml')!.async('string'))
    },
  }
}

const hiddenMergeTable =
  '<w:tbl><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
  '<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>START</w:t></w:r></w:p></w:tc><w:tc><w:p/></w:tc></w:tr>' +
  '<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p><w:r><w:t>RETAINED_HIDDEN</w:t></w:r></w:p></w:tc><w:tc><w:p/></w:tc></w:tr></w:tbl>'

it('refuses hidden continuation migration at native and dialog command boundaries', async () => {
  const { editor } = await openSynthetic(hiddenMergeTable + '<w:p/>')
  const table = editor.state.doc.firstChild!
  const map = TableMap.get(table)
  editor.view.dispatch(
    editor.state.tr.setSelection(CellSelection.create(editor.state.doc, 1 + map.map[0])),
  )
  const before = editor.getJSON()
  expect(computeFormatState(editor)).toMatchObject({
    canSplitCell: false,
    canSplitCellsDialog: false,
  })
  for (const command of [splitCellWithSource, splitCellsInto(2, 1, false)]) {
    expect(command(editor.state)).toBe(false)
    expect(command(editor.state, editor.view.dispatch)).toBe(false)
    expect(editor.getJSON()).toEqual(before)
  }
  expect(splitGridCell(table, { row: 0, col: 0 }, 2, 1)).toBeNull()
  editor.view.dispatch(
    editor.state.tr.setSelection(CellSelection.create(editor.state.doc, 1 + map.map[1])),
  )
  expect(deleteRowWithSource(editor.state)).toBe(false)
  expect(deleteRowWithSource(editor.state, editor.view.dispatch)).toBe(false)
  expect(computeFormatState(editor).canDeleteRow).toBe(false)
  expect(deleteCells('row')(editor.state)).toBe(false)
  expect(deleteCells('row')(editor.state, editor.view.dispatch)).toBe(false)
  editor.view.dispatch(
    editor.state.tr.setSelection(CellSelection.create(editor.state.doc, 1 + map.map[3])),
  )
  expect(splitTableAtSelection()(editor.state)).toBe(false)
  expect(splitTableAtSelection()(editor.state, editor.view.dispatch)).toBe(false)
  expect(computeFormatState(editor).canSplitTable).toBe(false)
  expect(editor.getJSON()).toEqual(before)
  expect(editor.commands.undo()).toBe(false)
})

it('retains every pre-existing duplicate source binding when splitting another table', async () => {
  const wideTable =
    '<w:tbl><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>OTHER</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
  const { editor } = await openSynthetic(hiddenMergeTable + wideTable + '<w:p/>')
  const original = editor.state.doc.firstChild!
  // Tracked replacement legitimately retains the original and a replacement with its source bindings.
  editor.view.dispatch(editor.state.tr.insert(original.nodeSize, original))
  const before = editor.getJSON().content!.slice(0, 2)
  const thirdPos = original.nodeSize * 2
  const third = editor.state.doc.nodeAt(thirdPos)!
  const pos = thirdPos + 1 + TableMap.get(third).map[0]
  editor.view.dispatch(editor.state.tr.setSelection(CellSelection.create(editor.state.doc, pos)))
  expect(splitCellWithSource(editor.state, editor.view.dispatch)).toBe(true)
  expect(editor.getJSON().content!.slice(0, 2)).toEqual(before)
  const pieces = editor.state.doc.nodeAt(thirdPos)!.firstChild!
  expect(pieces.child(0).attrs.sourceCell).toBeTruthy()
  expect(pieces.child(1).attrs.sourceCell).toBeNull()
})

it.each(['native', 'dialog'] as const)(
  'keeps inherited properties on new %s split pieces without duplicating source identity',
  async (kind) => {
    const { editor, saved } = await openSynthetic(
      '<w:tbl><w:tblPr><w:tblStyle w:val="Inherited"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
        '<w:tr><w:tc><w:tcPr><w:tcW w:w="4800" w:type="dxa"/><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>START</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>',
      '<w:style w:type="table" w:styleId="Inherited"><w:name w:val="Inherited"/><w:tcPr><w:shd w:fill="FF0000"/></w:tcPr></w:style>',
    )
    const map = TableMap.get(editor.state.doc.firstChild!)
    editor.view.dispatch(
      editor.state.tr.setSelection(CellSelection.create(editor.state.doc, 1 + map.map[0])),
    )
    expect(
      (kind === 'native' ? splitCellWithSource : splitCellsInto(1, 2, false))(
        editor.state,
        editor.view.dispatch,
      ),
    ).toBe(true)
    const row = editor.state.doc.firstChild!.firstChild!
    expect(row.child(0).attrs.sourceCell).toBeTruthy()
    expect(row.child(1).attrs.sourceCell).toBeNull()
    expect((await saved()).getElementsByTagNameNS(ns, 'shd')).toHaveLength(0)
  },
)

it('retains the surviving orphan continuation by source when identical rows are deleted', async () => {
  const row = (hidden: string) =>
    '<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p><w:r><w:t>' +
    hidden +
    '</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>SAME</w:t></w:r></w:p></w:tc></w:tr>'
  const { caller, saved } = await openSynthetic(
    '<w:tbl><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
      row('HIDDEN_A') +
      row('HIDDEN_B') +
      '</w:tbl><w:p/>',
  )
  const result = await caller.execute({
    id: 'delete-first-identical-row',
    name: 'apply_ops',
    input: { ops: [{ op: 'deleteTableRow', target: { blockIndexes: [0] }, row: 0 }] },
  })
  expect(result.isError, result.output).not.toBe(true)
  const xml = await saved()
  expect(xml.documentElement.textContent).toContain('HIDDEN_B')
  expect(xml.documentElement.textContent).not.toContain('HIDDEN_A')
})

it.each(['delete', 'split'] as const)(
  'refuses ambiguous hidden continuation migration before a %s operation',
  async (kind) => {
    const { caller, editor } = await openSynthetic(
      '<w:tbl><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>' +
        '<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>START</w:t></w:r></w:p></w:tc></w:tr>' +
        '<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p><w:r><w:t>RETAINED_HIDDEN</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>',
    )
    const before = caller.context()
    const content = editor.getJSON()
    const result = await caller.execute({
      id: 'ambiguous-hidden-migration',
      name: 'apply_ops',
      input: {
        ops: [
          kind === 'delete'
            ? { op: 'deleteTableRow', target: { blockIndexes: [0] }, row: 0 }
            : { op: 'splitTableCell', target: { blockIndexes: [0] }, cell: { row: 0, col: 0 } },
        ],
      },
    })
    expect(result).toMatchObject({ isError: true, committed: false, revision: before.revision })
    expect(result.output).toContain('hidden continuation')
    expect(editor.getJSON()).toEqual(content)
    expect(editor.commands.undo()).toBe(false)
  },
)

it('keeps source bindings private and does not copy one cell source onto new split pieces', async () => {
  const source = await readFile(
    resolve(__dirname, '../hosts/qt/tests/fixtures/mixed-run-table.docx'),
  )
  const originalZip = await JSZip.loadAsync(source)
  const { editor, caller, saved } = await openSynthetic(
    (await originalZip.file('word/document.xml')!.async('string')).match(
      /<w:body>([\s\S]*?)<w:sectPr/,
    )![1],
  )
  const before = await caller.execute({
    id: 'source-private',
    name: 'query_document',
    input: { type: 'docTable' },
  })
  const queried = (JSON.parse(before.output) as DocsQueryResult).items[0]
  const table = (editor.getJSON() as PmNode).content!.find((node) => node.type === 'docTable')!
  expect(before.output).not.toContain(String(table.attrs?.sourceTable))
  expect(editor.view.dom.outerHTML).not.toContain(String(table.attrs?.sourceTable))
  const result = await caller.execute({
    id: 'split-source',
    name: 'apply_ops',
    input: {
      ops: [
        {
          op: 'splitTableCell',
          target: { blockIndexes: [queried.blockIndex] },
          cell: { row: 0, col: 0 },
        },
      ],
    },
  })
  expect(result.isError, result.output).not.toBe(true)
  const sources: string[] = []
  editor.state.doc.descendants((node) => {
    if (node.attrs.sourceCell) sources.push(JSON.stringify(node.attrs.sourceCell))
  })
  expect(new Set(sources).size).toBe(sources.length)
  const xml = await saved()
  expect(xml.documentElement.outerHTML).not.toContain('sourceCell')
  expect(xml.documentElement.outerHTML).not.toContain(String(table.attrs?.sourceTable))
})

it('saves actual source-row relocation even when visible geometry and text are identical', async () => {
  const row = (hidden: string) =>
    '<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p><w:r><w:t>' +
    hidden +
    '</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>SAME</w:t></w:r></w:p></w:tc></w:tr>'
  const { editor, saved } = await openSynthetic(
    '<w:tbl><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
      row('HIDDEN_A') +
      row('HIDDEN_B') +
      '</w:tbl><w:p/>',
  )
  const table = editor.state.doc.firstChild!
  editor.view.dispatch(
    editor.state.tr.replaceWith(
      0,
      table.nodeSize,
      table.type.create(table.attrs, [table.child(1), table.child(0)]),
    ),
  )
  const xml = await saved()
  expect(xml.documentElement.textContent).toBe('HIDDEN_BSAMEHIDDEN_ASAME')
})

it.each(['row', 'fill', 'rewrite'] as const)(
  'preserves the actual conditional-style source during a %s edit',
  async (kind) => {
    const styles =
      '<w:style w:type="table" w:styleId="Conditional"><w:name w:val="Conditional"/>' +
      '<w:tblStylePr w:type="firstRow"><w:tcPr><w:shd w:fill="FF0000"/></w:tcPr></w:tblStylePr>' +
      '<w:tblStylePr w:type="lastRow"><w:tcPr><w:shd w:fill="0000FF"/></w:tcPr></w:tblStylePr></w:style>'
    const row = (text: string) =>
      '<w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>' +
      '<w:p><w:r><w:t>' +
      text +
      '</w:t></w:r></w:p></w:tc></w:tr>'
    const { caller, parsed, saved } = await openSynthetic(
      '<w:tbl><w:tblPr><w:tblStyle w:val="Conditional"/>' +
        '<w:tblLook w:firstRow="1" w:lastRow="1" w:noHBand="1" w:noVBand="1"/></w:tblPr>' +
        '<w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>' +
        row('RED') +
        row('BLUE') +
        '</w:tbl><w:p/>',
      styles,
    )
    expect(parsed.blocks[0].table?.rows.map((cells) => cells[0].fill)).toEqual(['FF0000', '0000FF'])
    expect(parsed.blocks[0].table?.rows[0][0].rawTcPr).toBe(
      parsed.blocks[0].table?.rows[1][0].rawTcPr,
    )
    const result =
      kind === 'rewrite'
        ? await caller.execute({
            id: 'rewrite-ordinary-cell',
            name: 'replace_blocks',
            input: {
              startBlockIndex: 0,
              endBlockIndex: 0,
              html: '<table><tr><td>EDITED</td></tr><tr><td>BLUE</td></tr></table>',
            },
          })
        : await caller.execute({
            id: 'conditional-style-edit',
            name: 'apply_ops',
            input: {
              ops: [
                kind === 'row'
                  ? {
                      op: 'insertTableRow',
                      target: { blockIndexes: [0] },
                      row: 1,
                      position: 'after',
                    }
                  : {
                      op: 'setTableCellFormat',
                      target: { blockIndexes: [0] },
                      range: { row: 0, col: 0 },
                      fill: '#0000FF',
                    },
              ],
            },
          })
    expect(result.isError, result.output).not.toBe(true)
    const xml = await saved()
    const cells = [...xml.getElementsByTagNameNS(ns, 'tc')]
    expect(cells[1].getElementsByTagNameNS(ns, 'shd')).toHaveLength(0)
    if (kind !== 'fill') expect(cells[0].getElementsByTagNameNS(ns, 'shd')).toHaveLength(0)
    else
      expect(cells[0].getElementsByTagNameNS(ns, 'shd')[0].getAttributeNS(ns, 'fill')).toBe(
        '0000FF',
      )
    if (kind === 'rewrite') expect(xml.documentElement.textContent).toContain('EDITED')
  },
)
const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const parseXml = (xml: string) => new DOMParser().parseFromString(xml, 'application/xml')

it.each(['fill', 'row'] as const)(
  'preserves unedited grid, widths and themed cell properties after a %s edit',
  async (kind) => {
    const source = await readFile(
      resolve(__dirname, '../hosts/qt/tests/fixtures/mixed-run-table.docx'),
    )
    const originalZip = await JSZip.loadAsync(source)
    const original = parseXml(await originalZip.file('word/document.xml')!.async('string'))
    const parsed = await parseDocx(source)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const runtime = createDocsRuntime({
      current: () => ({
        editor,
        identity: editor,
        numIds: { bullet: null, ordered: null },
        autoAcceptUntracked: true,
      }),
    })
    cleanup.push(
      () => editor.destroy(),
      () => runtime.dispose(),
    )
    const caller = runtime.connect({ id: `table-properties-${kind}` })
    const before = caller.context()
    const query = await caller.execute({
      id: 'query-table',
      name: 'query_document',
      input: { type: 'docTable' },
    })
    expect(query.isError, query.output).not.toBe(true)
    const table = (JSON.parse(query.output) as DocsQueryResult).items[0]
    const result = await caller.execute({
      id: 'edit-table',
      name: 'apply_ops',
      input: {
        ops: [
          kind === 'fill'
            ? {
                op: 'setTableCellFormat',
                target: { blockIndexes: [table.blockIndex] },
                range: { row: 0, col: 0 },
                fill: '#112233',
              }
            : {
                op: 'insertTableRow',
                target: { blockIndexes: [table.blockIndex] },
                row: 1,
                position: 'after',
              },
        ],
      },
    })
    expect(result.isError, result.output).not.toBe(true)
    expect(result.committed).toBe(true)
    expect(result.revision).toBeGreaterThan(before.revision)
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const zip = await JSZip.loadAsync(saved)
    const xml = parseXml(await zip.file('word/document.xml')!.async('string'))
    const reopened = await parseDocx(saved)
    const grid = (doc: Document) => doc.getElementsByTagNameNS(ns, 'tblGrid')[0].outerHTML
    expect(grid(xml)).toBe(grid(original))
    const cells = (doc: Document) => [...doc.getElementsByTagNameNS(ns, 'tc')]
    const neighbor = (doc: Document) => cells(doc).find((cell) => cell.textContent === 'NEIGHBOR')!
    expect(neighbor(xml).getElementsByTagNameNS(ns, 'tcPr')[0].outerHTML).toBe(
      neighbor(original).getElementsByTagNameNS(ns, 'tcPr')[0].outerHTML,
    )
    const continuation = (doc: Document) =>
      cells(doc).find((cell) => {
        const merge = cell.getElementsByTagNameNS(ns, 'vMerge')[0]
        return merge && merge.getAttributeNS(ns, 'val') !== 'restart'
      })!
    expect(continuation(xml).getElementsByTagNameNS(ns, 'tcW')[0].outerHTML).toBe(
      continuation(original).getElementsByTagNameNS(ns, 'tcW')[0].outerHTML,
    )
    expect(reopened.blocks.find((block) => block.type === 'table')?.table?.rows).toHaveLength(
      kind === 'row' ? 3 : 2,
    )
    if (kind === 'row') {
      const rows = xml.getElementsByTagNameNS(ns, 'tr')
      expect(
        [...rows[rows.length - 1].getElementsByTagNameNS(ns, 'tcW')].map((width) =>
          width.getAttributeNS(ns, 'w'),
        ),
      ).toEqual(['1255', '974'])
    }
    if (kind === 'fill')
      expect(cells(xml)[0].getElementsByTagNameNS(ns, 'shd')[0].getAttributeNS(ns, 'fill')).toBe(
        '112233',
      )
    expect(xml.documentElement.textContent).toContain('AFTER_TABLE')
  },
)
