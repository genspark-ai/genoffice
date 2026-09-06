import { Editor } from '@tiptap/core'
import {
  CellSelection,
  addColumnAfter,
  addRowAfter,
  columnResizingPluginKey,
  mergeCells,
  splitCell,
} from '@tiptap/pm/tables'
import { DOMSerializer } from '@tiptap/pm/model'
import { NodeSelection, TextSelection } from '@tiptap/pm/state'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import {
  blocksToPmDoc,
  pmDocToSavePlan,
  pmTableToModel,
  type PmNode,
} from '../src/renderer/editor/convert'
import { editorExtensions, rowHeightCss, tableRowEatCss } from '../src/renderer/editor/extensions'
import { renderTableSpec } from '../src/renderer/editor/protected-render'
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
          eaSlotEmpty: null,
          sizeHalfPoints: null,
          font: 'Calibri',
          fontAscii: 'Calibri',
          csFont: null,
          charSpacingTwips: null,
          charScaleEm: null,
          highlight: null,
          shading: null,
          vertAlign: null,
          em: null,
          boldOff: null,
          italicOff: null,
          kern: null,
          caps: null,
          vanish: null,
          eaLang: null,
          cs: null,
          rtl: null,
          styleId: null,
          rawRPr:
            '<w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="Calibri"/><w:b/><w:color w:val="1F4E78"/></w:rPr>',
          themeRFonts: null,
          themeColor: null,
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

  it('persists a table alignment change (tblAlign → w:jc) through save and reload', async () => {
    const { editor, parsed } = await openTable()
    const table = editor.state.doc.firstChild!
    // the Table Layout ribbon sets tblAlign via updateAttributes('docTable', …)
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(0, undefined, { ...table.attrs, tblAlign: 'center' }),
    )

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(1)
    const reparsed = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    expect(reparsed.blocks[0].table?.align).toBe('center')
    expect(reparsed.blocks[0].originalXml).toContain('<w:jc w:val="center"/>')
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

  it('renders legacy over-wide grids clamped to the paper edge (content box + right margin)', async () => {
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
    expect(spec[1].style).toContain(
      'width:min(1200px,calc(var(--doc-content-w,100%) + var(--doc-margin-right,0px)))',
    )
    const cols = spec[2].slice(2) as Array<[string, Record<string, string>]>
    expect(cols.map((col) => col[1].style)).toEqual(['width:50.00%', 'width:50.00%'])
    editor.destroy()
  })

  it('fixed-layout tables hold the declared width instead of narrowing to the paper', async () => {
    const { editor } = await openTable()
    const table = editor.state.doc.firstChild!
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(0, undefined, {
        ...table.attrs,
        widthPx: 1200,
        tblFixedLayout: true,
      }),
    )
    const spec = editor.schema.nodes.docTable.spec.toDOM!(editor.state.doc.firstChild!) as [
      string,
      Record<string, string>,
    ]
    expect(spec[1].style).toContain('width:1200px')
    expect(spec[1].style).toContain('max-width:none')
    expect(spec[1].style).not.toContain('min(')

    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(0, undefined, {
        ...editor.state.doc.firstChild!.attrs,
        tblAlign: 'center',
      }),
    )
    const centered = editor.schema.nodes.docTable.spec.toDOM!(editor.state.doc.firstChild!) as [
      string,
      Record<string, string>,
    ]
    expect(centered[1].style).toContain('margin-left:calc((var(--doc-content-w,100%) - 1200px)/2)')
    editor.destroy()
  })

  it('resolves a pct table width against its own section column', async () => {
    // w:tblW type="pct" is a share of the section's TEXT COLUMN. The canvas pads by
    // the first section's margins, so a bare 100% made every table of a document
    // with a full-bleed cover section (w:pgMar w:left="0") span the whole paper and
    // hang off its right edge once the section's own left inset is applied.
    const { editor } = await openTable()
    const table = editor.state.doc.firstChild!
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(0, undefined, {
        ...table.attrs,
        tblAutoFit: 'fixed',
        widthPx: null,
        widthPct: 100,
      }),
    )
    const spec = editor.schema.nodes.docTable.spec.toDOM!(editor.state.doc.firstChild!) as [
      string,
      Record<string, string>,
    ]
    expect(spec[1].style).toContain('width:calc(var(--doc-content-w,100%) * 1)')
    expect(spec[1].style).not.toContain('width:100%')
    editor.destroy()
  })

  it('resolves an AutoFit-to-Window table against its own section column', async () => {
    // the imported shape the bug came in on: <w:tblLayout w:type="autofit"/> +
    // <w:tblW w:w="5000" w:type="pct"/> parses as AutoFit to Window
    const pctTable =
      '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/>' +
      '<w:tblW w:w="5000" w:type="pct"/><w:tblLayout w:type="autofit"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    const source = await buildDocx({ bodyXml: pctTable })
    const parsed = await parseDocx(source)
    expect(parsed.blocks[0].table?.autoFit).toBe('window')
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const spec = editor.schema.nodes.docTable.spec.toDOM!(editor.state.doc.firstChild!) as [
      string,
      Record<string, string>,
    ]
    expect(spec[1].style).toContain('width:var(--doc-content-w,100%)')
    expect(spec[1].style).not.toContain('width:100%')
    editor.destroy()
  })

  it('takes a positive table indent out of the right-margin spill allowance', async () => {
    const { editor } = await openTable()
    const table = editor.state.doc.firstChild!
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(0, undefined, {
        ...table.attrs,
        widthPx: 1200,
        indentTwips: 1450,
      }),
    )
    const spec = editor.schema.nodes.docTable.spec.toDOM!(editor.state.doc.firstChild!) as [
      string,
      Record<string, string>,
    ]
    expect(spec[1].style).toContain(
      'width:min(1200px,calc(var(--doc-content-w,100%) + var(--doc-margin-right,0px) - 96.7px))',
    )
    expect(spec[1].style).toContain('margin-left:96.7px')
    editor.destroy()
  })

  it('hangs a start-aligned bidiVisual table from the right margin, indent mirrored', async () => {
    const { editor } = await openTable()
    const table = editor.state.doc.firstChild!
    const specOf = () =>
      editor.schema.nodes.docTable.spec.toDOM!(editor.state.doc.firstChild!) as [
        string,
        Record<string, string>,
      ]
    // prod100 sas 073: fixed layout, 10253 twips wide, w:tblInd -893 — Word puts
    // the right edge 60px past the right margin and lets the table spill left
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(0, undefined, {
        ...table.attrs,
        widthPx: 683,
        tblFixedLayout: true,
        bidiVisual: true,
        indentTwips: -893,
      }),
    )
    let spec = specOf()
    expect(spec[1].dir).toBe('rtl')
    expect(spec[1].style).toContain('width:683px')
    expect(spec[1].style).toContain('margin-left:calc(var(--doc-content-w,100%) - 683px + 59.5px)')

    // autofit grid with a positive indent: the spill allowance is the LEFT margin
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(0, undefined, {
        ...editor.state.doc.firstChild!.attrs,
        tblFixedLayout: false,
        indentTwips: 1450,
      }),
    )
    spec = specOf()
    const width = 'min(683px,calc(var(--doc-content-w,100%) + var(--doc-margin-left,0px) - 96.7px))'
    expect(spec[1].style).toContain(`width:${width}`)
    expect(spec[1].style).toContain(
      `margin-left:calc(var(--doc-content-w,100%) - ${width} - 96.7px)`,
    )

    // explicit right/center alignment keeps the LTR placement rules
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(0, undefined, {
        ...editor.state.doc.firstChild!.attrs,
        tblAlign: 'right',
      }),
    )
    spec = specOf()
    expect(spec[1].style).toContain('margin-left:auto')
    expect(spec[1].style).not.toContain('96.7px')
    editor.destroy()
  })

  it('a floating table (w:tblpPr) drops alignment/indent margins for the float gaps', async () => {
    const { editor } = await openTable()
    const table = editor.state.doc.firstChild!
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(0, undefined, {
        ...table.attrs,
        widthPx: 400,
        tblFloat: 'right',
        tblAlign: 'center',
        indentTwips: 1450,
      }),
    )
    const spec = editor.schema.nodes.docTable.spec.toDOM!(editor.state.doc.firstChild!) as [
      string,
      Record<string, string>,
    ]
    expect(spec[1].class).toContain('doc-table-float-right')
    expect(spec[1].style ?? '').not.toContain('margin-left:')
    expect(spec[1].style ?? '').not.toContain('margin-right:')
    // the indent must not shrink the float width either
    expect(spec[1].style ?? '').not.toContain('96.7px')
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

  it('wraps hRule="exact" row cells in a fixed-height clip box; atLeast rows stay unwrapped', async () => {
    const xml =
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
      '<w:tr><w:trPr><w:trHeight w:val="907" w:hRule="exact"/></w:trPr>' +
      '<w:tc><w:p><w:r><w:t>X</w:t></w:r></w:p></w:tc></w:tr>' +
      '<w:tr><w:trPr><w:trHeight w:val="907" w:hRule="atLeast"/></w:trPr>' +
      '<w:tc><w:p><w:r><w:t>Y</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    const parsed = await parseDocx(await buildDocx({ bodyXml: xml }))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const rows = editor.view.dom.querySelectorAll('tr')
    const clip = rows[0].querySelector(':scope > td > div.cell-clip') as HTMLElement
    expect(clip).toBeTruthy()
    expect(clip.style.height).toBe('60.5px')
    expect(rows[1].querySelector('.cell-clip')).toBeNull()
    editor.destroy()
  })

  it('advances declared-height rows by the horizontal gridline width like Word', async () => {
    const bordered =
      '<w:tbl><w:tblPr><w:tblBorders>' +
      '<w:top w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/>' +
      '<w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/>' +
      '</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
      '<w:tr><w:trPr><w:trHeight w:val="814"/></w:trPr>' +
      '<w:tc><w:p><w:r><w:t>X</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    const parsed = await parseDocx(await buildDocx({ bodyXml: bordered }))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const table = editor.view.dom.querySelector('table.doc-table') as HTMLElement
    // sz=4 eighths = 0.5pt gridline = 0.67px at 96dpi
    expect(table.style.getPropertyValue('--doc-row-eat')).toBe('0.67px')
    const tr = table.querySelector('tr') as HTMLElement
    expect(tr.getAttribute('style')).toContain('calc(54.3px + var(--doc-row-eat,0px))')
    editor.destroy()

    const borderless =
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
      '<w:tr><w:trPr><w:trHeight w:val="814"/></w:trPr>' +
      '<w:tc><w:p><w:r><w:t>X</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    const parsed2 = await parseDocx(await buildDocx({ bodyXml: borderless }))
    const editor2 = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed2.blocks) as never,
    })
    const table2 = editor2.view.dom.querySelector('table.doc-table') as HTMLElement
    expect(table2.style.getPropertyValue('--doc-row-eat')).toBe('')
    editor2.destroy()
  })

  it('advances declared-height rows by explicit tcBorders when the table has no tblBorders', async () => {
    const cell = (borders: string, text: string): string =>
      `<w:tc><w:tcPr><w:tcBorders>${borders}</w:tcBorders></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`
    const lined =
      '<w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/>' +
      '<w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/>'
    const bare =
      '<w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/>'
    const xml =
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
      `<w:tr><w:trPr><w:trHeight w:val="225"/></w:trPr>${cell(bare, 'Title')}</w:tr>` +
      `<w:tr><w:trPr><w:trHeight w:val="225"/></w:trPr>${cell(lined, 'A')}</w:tr>` +
      `<w:tr><w:trPr><w:trHeight w:val="450"/></w:trPr>${cell(lined, 'B')}</w:tr>` +
      '<w:tr><w:trPr><w:trHeight w:val="225"/></w:trPr><w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc></w:tr>' +
      '</w:tbl>'
    const parsed = await parseDocx(await buildDocx({ bodyXml: xml }))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const table = editor.view.dom.querySelector('table.doc-table') as HTMLElement
    expect(table.style.getPropertyValue('--doc-row-eat')).toBe('')
    const rows = Array.from(table.querySelectorAll('tr')).map((tr) => tr.getAttribute('style'))
    // explicit nil borders: no gridline, no advance; sz=4 = 0.5pt = 0.67px per lined row
    expect(rows[0]).toMatch(/^height:\s*15(\.0)?px;?$/)
    expect(rows[1]).toContain('calc(15.0px + 0.67px * var(--doc-row-grid,1))')
    expect(rows[2]).toContain('calc(30.0px + 0.67px * var(--doc-row-grid,1))')
    expect(rows[3]).toContain('calc(15.0px + var(--doc-row-eat,0px))')
    editor.destroy()

    const spec = renderTableSpec({
      rows: [
        [
          {
            paras: ['A'],
            borders: {
              top: { style: 'single', szEighths: 4 },
              bottom: { style: 'single', szEighths: 4 },
            },
          },
        ],
        [{ paras: ['B'], borders: { top: { style: 'nil' }, bottom: { style: 'nil' } } }],
      ],
      rowHeightsTwips: [225, 225],
    })
    const dom = document.createElement('div')
    dom.appendChild(DOMSerializer.renderSpec(document, spec as never).dom)
    const trs = Array.from(dom.querySelectorAll('tr')).map((tr) => tr.getAttribute('style'))
    expect(trs[0]).toContain('calc(15.0px + 0.67px * var(--doc-row-grid,1))')
    expect(trs[1]).toMatch(/^height:\s*15(\.0)?px;?$/)
  })

  it('row gridline: mixed cells take the larger of cell line and table gridline, spaced tables none, gap cells ignored', () => {
    const thin = { style: 'single', szEighths: 4 }
    const thick = { style: 'single', szEighths: 12 }
    // one cell explicit thin, one inheriting the table insideH: whichever is wider wins at render time
    expect(rowHeightCss(225, [{ top: thin, bottom: thin }, null])).toBe(
      'height:calc(15.0px + max(0.67px, var(--doc-row-eat,0px)) * var(--doc-row-grid,1))',
    )
    expect(rowHeightCss(225, [{ top: thick, bottom: thick }])).toBe(
      'height:calc(15.0px + 2.00px * var(--doc-row-grid,1))',
    )
    expect(tableRowEatCss({ insideH: thin } as never, true)).toEqual(['--doc-row-grid:0'])
    expect(tableRowEatCss({ insideH: thin } as never, false)).toEqual(['--doc-row-eat:0.67px'])
    // an all-nil row keeps its bare height even though a gap placeholder (borders null) sits in it
    const nil = { style: 'nil' }
    expect(rowHeightCss(225, [{ top: nil, bottom: nil }])).toBe('height:15.0px')
    // nested tables must not inherit a spaced parent's grid switch or its gridline
    const css = readFileSync(resolve(__dirname, '../src/renderer/styles.css'), 'utf8')
    expect(css).toMatch(/\.doc-table \{[^}]*--doc-row-eat: 0px;[^}]*--doc-row-grid: 1;/s)
  })

  it('insets cell text by max(margin, border/2) and spans the outer half-borders (Word probe 2026-09-03)', async () => {
    const xml =
      '<w:tbl><w:tblPr><w:tblBorders>' +
      '<w:left w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/>' +
      '<w:insideV w:val="single" w:sz="4"/></w:tblBorders>' +
      '<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>' +
      '</w:tblPr><w:tblGrid><w:gridCol w:w="1700"/><w:gridCol w:w="1700"/></w:tblGrid>' +
      '<w:tr><w:tc><w:tcPr><w:tcW w:w="1700" w:type="dxa"/>' +
      '<w:tcBorders><w:left w:val="single" w:sz="16"/></w:tcBorders>' +
      '<w:tcMar><w:left w:w="200" w:type="dxa"/></w:tcMar></w:tcPr>' +
      '<w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:tcPr><w:tcW w:w="1700" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>' +
      '</w:tr></w:tbl>'
    const parsed = await parseDocx(await buildDocx({ bodyXml: xml }))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const table = editor.view.dom.querySelector('table.doc-table') as HTMLElement
    // 2 x 113px grid + half of the 1px outer borders on each side
    expect(table.getAttribute('style')).toContain('width: min(227px,')
    expect(table.style.getPropertyValue('--doc-bw-v')).toBe('1px')
    expect(table.style.getPropertyValue('--doc-bw-l')).toBe('1px')
    expect(table.style.getPropertyValue('--doc-cell-pad-l')).toBe('7.2px')
    expect(table.style.getPropertyValue('--doc-cell-pad-t')).toBe('0.0px')
    const [a, b] = [...table.querySelectorAll('td')] as HTMLElement[]
    // the cell's own 2pt left border and tcMar feed the inset rule as variables, not as padding
    expect(a.style.getPropertyValue('--cell-bw-l')).toBe('3px')
    expect(a.style.getPropertyValue('--doc-cell-pad-l')).toBe('13.3px')
    expect(a.style.paddingLeft).toBe('')
    expect(b.style.getPropertyValue('--cell-bw-l')).toBe('')
    editor.destroy()
  })

  it('hands the border snap delta back to the row height (Word probe 2026-09-04: rows grow by the true sz/8 pt width)', async () => {
    const xml =
      '<w:tbl><w:tblPr><w:tblBorders>' +
      '<w:top w:val="single" w:sz="2"/><w:bottom w:val="single" w:sz="6"/>' +
      '<w:insideH w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="1700"/><w:gridCol w:w="1700"/></w:tblGrid>' +
      '<w:tr><w:tc><w:tcPr><w:tcW w:w="1700" w:type="dxa"/>' +
      '<w:tcBorders><w:bottom w:val="single" w:sz="8"/></w:tcBorders></w:tcPr>' +
      '<w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:tcPr><w:tcW w:w="1700" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>' +
      '</w:tr></w:tbl>'
    const parsed = await parseDocx(await buildDocx({ bodyXml: xml }))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const table = editor.view.dom.querySelector('table.doc-table') as HTMLElement
    // drawn 1px each; true widths 0.333 / 1.0 / 0.667px
    expect(table.style.getPropertyValue('--doc-bd-t')).toBe('-0.667px')
    expect(table.style.getPropertyValue('--doc-bd-b')).toBe('0.000px')
    expect(table.style.getPropertyValue('--doc-bd-h')).toBe('-0.333px')
    const [a, b] = [...table.querySelectorAll('td')] as HTMLElement[]
    // a 1pt cell border is drawn 1px but measures 1.333px in Word
    expect(a.style.getPropertyValue('--cell-bd-b')).toBe('0.333px')
    expect(a.style.getPropertyValue('--cell-bd-t')).toBe('')
    expect(b.style.getPropertyValue('--cell-bd-b')).toBe('')
    editor.destroy()
  })

  it('keeps whole-border padding and the bare grid width for spaced (separate-border) tables', async () => {
    const xml =
      '<w:tbl><w:tblPr><w:tblCellSpacing w:w="30" w:type="dxa"/><w:tblBorders>' +
      '<w:left w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/>' +
      '<w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="1700"/><w:gridCol w:w="1700"/></w:tblGrid>' +
      '<w:tr><w:tc><w:tcPr><w:tcW w:w="1700" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:tcPr><w:tcW w:w="1700" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>' +
      '</w:tr></w:tbl>'
    const parsed = await parseDocx(await buildDocx({ bodyXml: xml }))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const table = editor.view.dom.querySelector('table.doc-table') as HTMLElement
    expect(table.getAttribute('style')).toContain('width: min(226px,')
    expect(table.style.getPropertyValue('--doc-bw-share')).toBe('0')
    expect(table.style.getPropertyValue('--doc-bd-share')).toBe('1')
    editor.destroy()
  })

  it('still wraps an exact row whose padding consumes the whole height (clip height 0)', async () => {
    const xml =
      '<w:tbl><w:tblPr><w:tblCellMar><w:top w:w="500" w:type="dxa"/>' +
      '<w:bottom w:w="500" w:type="dxa"/></w:tblCellMar></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
      '<w:tr><w:trPr><w:trHeight w:val="907" w:hRule="exact"/></w:trPr>' +
      '<w:tc><w:p><w:r><w:t>X</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    const parsed = await parseDocx(await buildDocx({ bodyXml: xml }))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const clip = editor.view.dom.querySelector('td > div.cell-clip') as HTMLElement
    expect(clip).toBeTruthy()
    expect(clip.style.height).toBe('0px')
    editor.destroy()
  })
})
