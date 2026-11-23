import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { generateTableModelXml, parseDocx, upsertStyleXml } from '@genoffice/docx-engine'
import { describe, expect, it, vi } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmTableToModel, type PmNode } from '../src/renderer/editor/convert'
import { measureWidthIn } from '../src/renderer/components/TablePropertiesDialog'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  applyTablePreset,
  repeatHeaderState,
  setTableAutoFit,
  setTableLookOption,
  tablePresetStyleXml,
  toggleRepeatHeaderRows,
  updateSelectedTableAttrs,
} from '../src/renderer/editor/table-properties'
import {
  setSelectedColumnPreferredWidth,
  shareColumnWidths,
} from '../src/renderer/editor/table-sizing'

const TABLE =
  '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/>' +
  '<w:tblW w:w="5000" w:type="pct"/><w:tblLayout w:type="autofit"/>' +
  '<w:tblCellMar><w:left w:w="120" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tblCellMar>' +
  '<w:tblLook w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>' +
  '</w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
  '<w:tr><w:trPr><w:tblHeader/></w:trPr>' +
  '<w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>' +
  '<w:tr><w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'

async function openTable(): Promise<Editor> {
  const parsed = await parseDocx(await buildDocx({ bodyXml: TABLE }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  let firstCell = -1
  editor.state.doc.descendants((node, pos) => {
    if (
      firstCell < 0 &&
      (node.type.name === 'docTableCell' || node.type.name === 'docTableHeader')
    ) {
      firstCell = pos
    }
  })
  editor.view.dispatch(
    editor.state.tr.setSelection(TextSelection.near(editor.state.doc.resolve(firstCell + 1))),
  )
  return editor
}

describe('editable table properties', () => {
  it('loads AutoFit, margins, style options, and repeat-header state', async () => {
    const editor = await openTable()
    const table = editor.getJSON().content![0] as unknown as PmNode
    expect(table.attrs).toMatchObject({
      tblAutoFit: 'window',
      cellMar: { left: 120, right: 120 },
      tblLook: {
        firstRow: true,
        lastRow: false,
        firstColumn: true,
        lastColumn: false,
        bandedRows: true,
        bandedColumns: false,
      },
    })
    expect(table.content![0].attrs).toMatchObject({ repeatHeader: true })
    expect(repeatHeaderState(editor.state)).toEqual({ enabled: true, active: true })
    editor.destroy()
  })

  it('toggles leading header rows and persists the edited model flag', async () => {
    const editor = await openTable()
    expect(toggleRepeatHeaderRows()(editor.state, editor.view.dispatch)).toBe(true)
    const table = editor.getJSON().content![0] as unknown as PmNode
    expect(table.content![0].attrs).toMatchObject({
      repeatHeader: false,
      repeatHeaderEdited: true,
    })
    expect(pmTableToModel(table as PmNode).repeatHeaderRows).toEqual([false, null])
    editor.destroy()
  })

  it('switches all three AutoFit modes and materializes a fixed grid', async () => {
    const editor = await openTable()
    expect(setTableAutoFit('contents', 600)(editor.state, editor.view.dispatch)).toBe(true)
    let table = editor.getJSON().content![0] as unknown as PmNode
    expect(table.attrs).toMatchObject({
      tblAutoFit: 'contents',
      tblAutoFitEdited: true,
      widthPx: null,
      widthPct: null,
      colWidthsPct: null,
    })

    expect(setTableAutoFit('window', 600)(editor.state, editor.view.dispatch)).toBe(true)
    expect((editor.getJSON().content![0] as unknown as PmNode).attrs).toMatchObject({
      tblAutoFit: 'window',
      widthPct: 100,
    })

    expect(setTableAutoFit('fixed', 600)(editor.state, editor.view.dispatch)).toBe(true)
    table = editor.getJSON().content![0] as unknown as PmNode
    expect(table.attrs).toMatchObject({ tblAutoFit: 'fixed', widthPct: null })
    expect(table.content![0].content![0].attrs?.colwidth).toEqual([300])
    expect(pmTableToModel(table as PmNode).autoFit).toBe('fixed')
    editor.destroy()
  })

  it("AutoFit Window drops the table's indent, as Word does", async () => {
    // Word probe 2026-11-23: a fixed 8800-twip table with tblInd -108 on a 9638
    // column becomes pct 5000 with no w:tblInd, its border on both margins
    const indented = TABLE.replace(
      '<w:tblW w:w="5000" w:type="pct"/><w:tblLayout w:type="autofit"/>',
      '<w:tblW w:w="6000" w:type="dxa"/><w:tblInd w:w="-108" w:type="dxa"/>' +
        '<w:tblLayout w:type="fixed"/>',
    )
    const parsed = await parseDocx(await buildDocx({ bodyXml: indented }))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const before = editor.getJSON().content![0] as unknown as PmNode
    expect(before.attrs).toMatchObject({ indentTwips: -108, tblIndentEdited: false })
    expect(pmTableToModel(before).indentTwips).toBeUndefined()
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.near(editor.state.doc.resolve(4))),
    )

    expect(setTableAutoFit('window', 600)(editor.state, editor.view.dispatch)).toBe(true)
    const table = editor.getJSON().content![0] as unknown as PmNode
    expect(table.attrs).toMatchObject({ indentTwips: null, tblIndentEdited: true, widthPct: 100 })
    expect(editor.view.dom.querySelector('table')!.getAttribute('style')).not.toContain(
      'margin-left',
    )
    const model = pmTableToModel(table)
    expect(model).toMatchObject({ autoFit: 'window', indentTwips: 0 })
    const xml = generateTableModelXml(model, indented)
    expect(xml).not.toContain('w:tblInd')
    expect(xml).toContain('<w:tblW w:w="5000" w:type="pct"/>')
    editor.destroy()
  })

  it('updates cell margins, floating position, and table style options', async () => {
    const editor = await openTable()
    expect(
      updateSelectedTableAttrs({
        cellMar: { top: 40, right: 100, bottom: 50, left: 110 },
        cellMarEdited: true,
        tblFloat: 'left',
        tblFloatXTwips: 720,
        tblFloatYTwips: 360,
        tblFloatHorzAnchor: 'margin',
        tblFloatVertAnchor: 'text',
        tblFloatDistance: { top: 60, right: 120, bottom: 70, left: 130 },
        tblFloatEdited: true,
      })(editor.state, editor.view.dispatch),
    ).toBe(true)
    expect(setTableLookOption('lastRow', true)(editor.state, editor.view.dispatch)).toBe(true)

    const table = editor.getJSON().content![0] as PmNode
    const model = pmTableToModel(table)
    expect(model.cellMarTwips).toEqual({ top: 40, right: 100, bottom: 50, left: 110 })
    expect(model.floatSide).toBe('left')
    expect(model.floatPos).toMatchObject({
      xTwips: 720,
      yTwips: 360,
      horzAnchor: 'margin',
      vertAnchor: 'text',
    })
    expect(model.tableLook?.lastRow).toBe(true)
    editor.destroy()
  })

  it('positions a right AutoFit float without restoring fixed widthPx', async () => {
    const editor = await openTable()
    expect(setTableAutoFit('contents', 600)(editor.state, editor.view.dispatch)).toBe(true)
    expect(
      updateSelectedTableAttrs({
        tblFloat: 'right',
        tblFloatSource: 'right',
        tblFloatXTwips: 3000,
        tblFloatWidthPx: 240,
        tblFloatEdited: true,
      })(editor.state, editor.view.dispatch),
    ).toBe(true)

    const table = editor.state.doc.firstChild!
    expect(table.attrs).toMatchObject({
      tblAutoFit: 'contents',
      widthPx: null,
      tblFloatWidthPx: 240,
    })
    const spec = editor.schema.nodes.docTable.spec.toDOM!(table) as [string, Record<string, string>]
    expect(spec[1].class).toContain('doc-table-autofit-contents')
    expect(spec[1].style).toContain('width:auto')
    expect(spec[1].style).toContain(
      'margin-right:max(0px,calc(var(--doc-content-w,100%) - 200.0px - 240.0px))',
    )
    editor.destroy()
  })

  it('keeps a pagination-suppressed float in the saved table model', async () => {
    const editor = await openTable()
    expect(
      updateSelectedTableAttrs({
        tblFloat: null,
        tblFloatSource: 'right',
        tblFloatSuppressed: true,
        tblFloatXTwips: 4800,
        tblFloatYTwips: 240,
        tblFloatEdited: true,
      })(editor.state, editor.view.dispatch),
    ).toBe(true)

    const model = pmTableToModel(editor.getJSON().content![0] as unknown as PmNode)
    expect(model.floatSide).toBe('right')
    expect(model.floatPos).toMatchObject({ xTwips: 4800, yTwips: 240 })
    editor.destroy()
  })

  it("a preset becomes the table's style, with the shading left to the style", async () => {
    const editor = await openTable()
    const preset = {
      styleId: 'BlueBanded',
      name: 'Blue Banded',
      headerFill: 'BDD7EE',
      band1Fill: 'DDEBF7',
      band2Fill: 'FFFFFF',
      borderColor: '9DC3E6',
    }
    expect(applyTablePreset(preset)(editor.state, editor.view.dispatch)).toBe(true)
    const table = editor.getJSON().content![0] as unknown as PmNode
    expect(table.attrs).toMatchObject({
      tblStyleId: 'BlueBanded',
      tblLook: { firstRow: true, bandedRows: true },
      tblLookEdited: true,
    })
    expect(table.content![0].content![0].attrs).toMatchObject({
      fill: null,
      borders: { top: { style: 'single', color: '9DC3E6' } },
    })
    const model = pmTableToModel(table)
    expect(model.tblStyleId).toBe('BlueBanded')
    expect(model.rows[1][0].borders?.left?.color).toBe('9DC3E6')

    const styles = upsertStyleXml(
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:styles>',
      { styleId: preset.styleId, sourceXml: tablePresetStyleXml(preset), basedOn: 'TableNormal' },
    )
    expect(styles).toContain(
      '<w:style w:type="table" w:customStyle="1" w:styleId="BlueBanded"><w:name w:val="Blue Banded"/><w:basedOn w:val="TableNormal"/>',
    )
    expect(styles).toMatch(
      /<w:tblStylePr w:type="firstRow"><w:tcPr><w:shd [^>]*w:fill="BDD7EE"\/>.*<w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd [^>]*w:fill="DDEBF7"\/>/,
    )
    editor.destroy()
  })

  it('drags a floating table handle into persisted x/y offsets', async () => {
    const editor = await openTable()
    updateSelectedTableAttrs({
      tblFloat: 'left',
      tblFloatSource: 'left',
      tblFloatXTwips: 0,
      tblFloatYTwips: 0,
      tblFloatEdited: true,
    })(editor.state, editor.view.dispatch)
    let firstCell = -1
    editor.state.doc.descendants((node, pos) => {
      if (
        firstCell < 0 &&
        (node.type.name === 'docTableCell' || node.type.name === 'docTableHeader')
      ) {
        firstCell = pos
      }
    })
    vi.spyOn(editor.view, 'posAtCoords').mockReturnValue({ pos: firstCell, inside: firstCell })
    editor.view.dom.dispatchEvent(
      new MouseEvent('mousemove', { bubbles: true, clientX: 10, clientY: 10 }),
    )
    const handle =
      editor.view.dom.parentElement!.querySelector<HTMLButtonElement>('.doc-table-handle')!
    handle.dispatchEvent(
      new MouseEvent('mousedown', { bubbles: true, button: 0, clientX: 10, clientY: 10 }),
    )
    window.dispatchEvent(new MouseEvent('mousemove', { clientX: 40, clientY: 30 }))
    window.dispatchEvent(new MouseEvent('mouseup', { clientX: 40, clientY: 30 }))

    expect((editor.getJSON().content![0] as unknown as PmNode).attrs).toMatchObject({
      tblFloatXTwips: 450,
      tblFloatYTwips: 300,
      tblFloatEdited: true,
    })
    editor.destroy()
  })

  it('shares a table between its columns by percent without changing its width', () => {
    expect(shareColumnWidths([200, 200, 200], 0, 1, 50)).toEqual([300, 150, 150])
    expect(shareColumnWidths([100, 300], 1, 2, 25)).toEqual([300, 100])
    // the other columns keep their minimum
    expect(shareColumnWidths([200, 200], 0, 1, 100)).toEqual([360, 40])
    expect(shareColumnWidths([200, 200], 0, 2, 30)).toEqual([200, 200])
  })

  it('a column measured in percent is written as a pct w:tcW and read back as one', async () => {
    const editor = await openTable()
    expect(
      setSelectedColumnPreferredWidth({ pct: 30 }, 600)(editor.state, editor.view.dispatch),
    ).toBe(true)
    let table = editor.getJSON().content![0] as unknown as PmNode
    expect(table.content!.map((row) => row.content![0].attrs?.widthPct)).toEqual([true, true])
    expect(table.content!.map((row) => row.content![1].attrs?.widthPct)).toEqual([false, false])
    // the table stays the share of the page it was
    expect(table.attrs).toMatchObject({ widthPct: 100, tblAutoFit: 'window' })
    expect((table.attrs!.colWidthsPct as number[]).map(Math.round)).toEqual([30, 70])
    const xml = generateTableModelXml(pmTableToModel(table as PmNode))
    expect(xml.match(/<w:tcW [^>]*>/g)).toEqual([
      '<w:tcW w:w="1500" w:type="pct"/>',
      expect.stringContaining('w:type="dxa"'),
      '<w:tcW w:w="1500" w:type="pct"/>',
      expect.stringContaining('w:type="dxa"'),
    ])
    const reread = await parseDocx(await buildDocx({ bodyXml: xml }))
    expect(reread.blocks[0].table!.rows.map((row) => row.map((cell) => !!cell.widthPct))).toEqual([
      [true, false],
      [true, false],
    ])

    // measuring in a length again drops the percent
    expect(
      setSelectedColumnPreferredWidth({ px: 120 }, 600)(editor.state, editor.view.dispatch),
    ).toBe(true)
    table = editor.getJSON().content![0] as unknown as PmNode
    expect(table.content![0].content![0].attrs).toMatchObject({ widthPct: false, colwidth: [120] })
    expect(generateTableModelXml(pmTableToModel(table as PmNode))).not.toContain('w:type="pct"/>')
    editor.destroy()
  })

  it('Measure in restates the typed width in the unit switched to', () => {
    const width = { cm: 4, pct: 50, inPct: false, tableCm: 16 }
    expect(measureWidthIn(width, true)).toMatchObject({ inPct: true, pct: 25, cm: 4 })
    expect(measureWidthIn({ ...width, inPct: true, pct: 75 }, false)).toMatchObject({
      inPct: false,
      cm: 12,
    })
    expect(measureWidthIn(width, false)).toEqual(width)
  })
})
