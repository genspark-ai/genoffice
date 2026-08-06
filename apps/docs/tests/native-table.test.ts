import { Editor } from '@tiptap/core'
import {
  CellSelection,
  addColumnAfter,
  addRowAfter,
  columnResizingPluginKey,
  mergeCells,
  splitCell,
} from '@tiptap/pm/tables'
import { NodeSelection, TextSelection } from '@tiptap/pm/state'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import {
  blocksToPmDoc,
  pmDocToSavePlan,
  pmTableToModel,
  type PmNode,
} from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  constrainSelectedTableWidth,
  constrainTableWidthAtCell,
  fitColumnWidths,
  setSelectedColumnWidth,
} from '../src/renderer/editor/table-sizing'

const TABLE =
  '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
  '<w:tr><w:tc><w:tcPr><w:shd w:fill="D9EAF7"/></w:tcPr><w:p><w:r><w:rPr>' +
  '<w:rFonts w:ascii="Calibri" w:eastAsia="Calibri"/><w:b/><w:color w:val="1F4E78"/>' +
  '</w:rPr><w:t>A</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>' +
  '<w:tr><w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'

function cellPositions(editor: Editor): number[] {
  const positions: number[] = []
  editor.state.doc.descendants((node, pos) => {
    if (node.type.name === 'docTableCell' || node.type.name === 'docTableHeader')
      positions.push(pos)
  })
  return positions
}

async function openTable(): Promise<{
  editor: Editor
  parsed: Awaited<ReturnType<typeof parseDocx>>
  source: Uint8Array
}> {
  const source = await buildDocx({ bodyXml: TABLE })
  const parsed = await parseDocx(source)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, parsed, source }
}

describe('native editable tables', () => {
  it('redistributes requested column widths within the section content box', () => {
    expect(fitColumnWidths([200, 200, 200], new Map([[0, 500]]), 600)).toEqual([500, 50, 50])
    const many = fitColumnWidths(new Array(20).fill(100), new Map([[0, 1000]]), 600)
    expect(many.reduce((sum, width) => sum + width, 0)).toBeCloseTo(600, 1)
    expect(many.every((width) => width > 0)).toBe(true)
  })

  it('converts imported tables to schema-native editable nodes', async () => {
    const { editor } = await openTable()
    const json = editor.getJSON() as unknown as PmNode

    expect(json.content?.[0].type).toBe('docTable')
    expect(json.content?.[0].content?.[0].content?.[0]).toMatchObject({
      type: 'docTableCell',
      attrs: { fill: 'D9EAF7', bold: true, color: '1F4E78', colspan: 1, rowspan: 1 },
    })
    expect(json.content?.[0].content?.[0].content?.[0].content?.[0].content?.[0].marks).toEqual([
      { type: 'bold' },
      {
        type: 'docTextStyle',
        attrs: {
          color: '1F4E78',
          sizeHalfPoints: null,
          font: 'Calibri',
          fontAscii: 'Calibri',
          charSpacingTwips: null,
          charScaleEm: null,
          highlight: null,
          vertAlign: null,
          em: null,
          styleId: null,
          rawRPr:
            '<w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="Calibri"/><w:b/><w:color w:val="1F4E78"/></w:rPr>',
        },
      },
    ])
    expect(editor.view.dom.querySelector('table.doc-table td')?.textContent).toBe('A')
    expect(editor.view.dom.querySelector('[data-doc-protected="table"]')).toBeNull()
    editor.destroy()
  })

  it('keeps an untouched imported table byte-identical', async () => {
    const { editor, parsed, source } = await openTable()
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)

    expect(plan.changedCount).toBe(0)
    expect(await saveDocx(parsed, plan.saveBlocks)).toEqual(source)
    editor.destroy()
  })

  it('uses the targeted cell-text patch when structure is unchanged', async () => {
    const { editor, parsed } = await openTable()
    const firstCell = cellPositions(editor)[0]
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, firstCell + 2)),
    )
    editor.view.dispatch(editor.state.tr.insertText('Edited '))

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const reparsed = await parseDocx(saved)
    expect(plan.changedCount).toBe(1)
    expect(reparsed.blocks[0].table?.rows[0][0].paras).toEqual(['Edited A'])
    expect(reparsed.blocks[0].originalXml).toContain('<w:tblStyle w:val="TableGrid"/>')
    expect(reparsed.blocks[0].originalXml).toContain('<w:t>B</w:t>')
    editor.destroy()
  })

  it('persists Ribbon-style rich marks by safely regenerating the table', async () => {
    const { editor, parsed } = await openTable()
    const secondCell = cellPositions(editor)[1]
    editor.view.dispatch(
      editor.state.tr.setSelection(CellSelection.create(editor.state.doc, secondCell)),
    )
    editor
      .chain()
      .setMark('bold')
      .setMark('docTextStyle', {
        color: 'FF0000',
        font: 'Arial',
        sizeHalfPoints: 28,
        highlight: null,
        vertAlign: null,
      })
      .run()

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const reparsed = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    const run = reparsed.blocks[0].table?.rows[0][1].richParas?.[0].runs[0]
    expect(plan.changedCount).toBe(1)
    expect(run).toMatchObject({
      text: 'B',
      bold: true,
      color: 'FF0000',
      font: 'Arial',
      sizeHalfPoints: 28,
    })
    expect(reparsed.blocks[0].originalXml).toContain('<w:tblStyle w:val="TableGrid"/>')
    editor.destroy()
  })

  it('round-trips row and column additions through generated OOXML', async () => {
    const { editor, parsed } = await openTable()
    const firstCell = cellPositions(editor)[0]
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, firstCell + 2)),
    )
    expect(addRowAfter(editor.state, editor.view.dispatch)).toBe(true)
    expect(addColumnAfter(editor.state, editor.view.dispatch)).toBe(true)

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const reparsed = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    expect(plan.changedCount).toBe(1)
    expect(reparsed.blocks[0].table?.rows).toHaveLength(3)
    expect(reparsed.blocks[0].table?.rows.every((row) => row.length === 3)).toBe(true)
    expect(reparsed.blocks[0].table?.rows[0][0]).toMatchObject({ fill: 'D9EAF7', bold: true })
    expect(reparsed.blocks[0].table?.rows[0][0].richParas?.[0].runs[0]).toMatchObject({
      text: 'A',
      bold: true,
      color: '1F4E78',
      font: 'Calibri',
    })
    expect(reparsed.blocks[0].originalXml).toContain('<w:tblStyle w:val="TableGrid"/>')
    editor.destroy()
  })

  it('round-trips clamped Ribbon and drag-resized column grids', async () => {
    const { editor, parsed } = await openTable()
    const firstCell = cellPositions(editor)[0]
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, firstCell + 2)),
    )

    expect(setSelectedColumnWidth(900, 624)(editor.state, editor.view.dispatch)).toBe(true)
    expect(constrainSelectedTableWidth(624)(editor.state, editor.view.dispatch)).toBe(true)

    const model = pmTableToModel(editor.getJSON().content![0] as unknown as PmNode)
    expect(model.colWidthsTwips?.reduce((sum, width) => sum + width, 0)).toBeLessThanOrEqual(9360)
    expect(model.colWidthsTwips?.every((width) => width > 0)).toBe(true)

    const plan = pmDocToSavePlan(editor.getJSON() as unknown as PmNode, parsed.blocks)
    const reparsed = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    const savedWidths = reparsed.blocks[0].table?.colWidthsTwips ?? []
    expect(savedWidths.reduce((sum, width) => sum + width, 0)).toBeLessThanOrEqual(9360)
    expect(savedWidths).toHaveLength(2)
    editor.destroy()
  })

  it('renders legacy over-wide grids clamped to the content box', async () => {
    const { editor } = await openTable()
    const table = editor.state.doc.firstChild!
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(0, undefined, {
        ...table.attrs,
        widthPx: 1200,
        colWidthsPct: [80, 80],
      }),
    )
    // jsdom's style parser drops min(); assert on the toDOM output spec instead
    const spec = editor.schema.nodes.docTable.spec.toDOM!(editor.state.doc.firstChild!) as [
      string,
      Record<string, string>,
      [string, Record<string, string>, ...Array<[string, Record<string, string>]>],
    ]
    expect(spec[1].style).toContain('width:min(1200px,100%)')
    const cols = spec[2].slice(2) as Array<[string, Record<string, string>]>
    expect(cols.map((col) => col[1].style)).toEqual(['width:50.00%', 'width:50.00%'])
    editor.destroy()
  })

  it('clamps a drag-committed grid without any selection inside the table', async () => {
    const { editor } = await openTable()
    // the resize handle sets no selection; a table NodeSelection is not "in table" either
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, 0)))
    const cells = cellPositions(editor)
    for (const pos of [cells[0], cells[2]]) {
      const cell = editor.state.doc.nodeAt(pos)!
      editor.view.dispatch(
        editor.state.tr.setNodeMarkup(pos, undefined, { ...cell.attrs, colwidth: [900] }),
      )
    }

    expect(constrainSelectedTableWidth(624)(editor.state, editor.view.dispatch)).toBe(false)
    expect(constrainTableWidthAtCell(cells[0], 624)(editor.state, editor.view.dispatch)).toBe(true)
    const model = pmTableToModel(editor.getJSON().content![0] as unknown as PmNode)
    expect(model.colWidthsTwips!.reduce((sum, width) => sum + width, 0)).toBeLessThanOrEqual(9360)
    expect(model.colWidthsTwips!.every((width) => width > 0)).toBe(true)
    editor.destroy()
  })

  it('constrains after the resize plugin commits, via the activeHandle mouseup path', async () => {
    const { editor } = await openTable()
    editor.view.dom.style.setProperty('--section-content-w', '624px')
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, 0)))
    const cells = cellPositions(editor)
    editor.view.dispatch(editor.state.tr.setMeta(columnResizingPluginKey, { setHandle: cells[0] }))
    editor.view.dispatch(
      editor.state.tr.setMeta(columnResizingPluginKey, {
        setDragging: { startX: 0, startWidth: 300 },
      }),
    )
    editor.view.dom.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }))

    // the plugin's window-level finish runs after our handler's microtask checkpoint:
    // replay that ordering, committing an oversized dragged column afterwards
    await Promise.resolve()
    for (const pos of [cells[0], cells[2]]) {
      const cell = editor.state.doc.nodeAt(pos)!
      editor.view.dispatch(
        editor.state.tr.setNodeMarkup(pos, undefined, { ...cell.attrs, colwidth: [900] }),
      )
    }
    editor.view.dispatch(editor.state.tr.setMeta(columnResizingPluginKey, { setDragging: null }))

    await new Promise((resolve) => setTimeout(resolve, 5))
    const model = pmTableToModel(editor.getJSON().content![0] as unknown as PmNode)
    expect(model.colWidthsTwips!.reduce((sum, width) => sum + width, 0)).toBeLessThanOrEqual(9360)
    editor.destroy()
  })

  it('maps merged and split cells between ProseMirror and TableModel', async () => {
    const { editor } = await openTable()
    let positions = cellPositions(editor)
    editor.view.dispatch(
      editor.state.tr.setSelection(
        CellSelection.create(editor.state.doc, positions[0], positions[1]),
      ),
    )
    expect(mergeCells(editor.state, editor.view.dispatch)).toBe(true)

    let table = editor.getJSON().content![0] as PmNode
    let model = pmTableToModel(table)
    expect(model.rows[0][0]).toMatchObject({ colSpan: 2, paras: ['A', 'B'] })

    positions = cellPositions(editor)
    editor.view.dispatch(
      editor.state.tr.setSelection(CellSelection.create(editor.state.doc, positions[0])),
    )
    expect(splitCell(editor.state, editor.view.dispatch)).toBe(true)
    table = editor.getJSON().content![0] as PmNode
    model = pmTableToModel(table)
    expect(model.rows[0]).toHaveLength(2)
    expect(model.rows[0][0].colSpan).toBeUndefined()
    expect(model.rows[0][1].colSpan).toBeUndefined()
    editor.destroy()
  })

  it('deletes a table NodeSelection with Backspace', async () => {
    const { editor } = await openTable()
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, 0)))

    expect(editor.commands.keyboardShortcut('Backspace')).toBe(true)
    expect(editor.state.doc.firstChild?.type.name).not.toBe('docTable')
    editor.destroy()
  })

  it('deletes a full CellSelection with forward Delete but preserves partial selections', async () => {
    const partial = await openTable()
    const partialCells = cellPositions(partial.editor)
    partial.editor.view.dispatch(
      partial.editor.state.tr.setSelection(
        CellSelection.create(partial.editor.state.doc, partialCells[0], partialCells[1]),
      ),
    )
    partial.editor.commands.keyboardShortcut('Delete')
    expect(partial.editor.state.doc.firstChild?.type.name).toBe('docTable')
    partial.editor.destroy()

    const whole = await openTable()
    const allCells = cellPositions(whole.editor)
    whole.editor.view.dispatch(
      whole.editor.state.tr.setSelection(
        CellSelection.create(whole.editor.state.doc, allCells[0], allCells.at(-1)!),
      ),
    )
    expect(whole.editor.commands.keyboardShortcut('Delete')).toBe(true)
    expect(whole.editor.state.doc.firstChild?.type.name).not.toBe('docTable')
    whole.editor.destroy()
  })
})
