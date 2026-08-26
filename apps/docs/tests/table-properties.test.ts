import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { parseDocx } from '@genoffice/docx-engine'
import { describe, expect, it, vi } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmTableToModel, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  applyTablePreset,
  repeatHeaderState,
  setTableAutoFit,
  setTableLookOption,
  toggleRepeatHeaderRows,
  updateSelectedTableAttrs,
} from '../src/renderer/editor/table-properties'

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

  it('applies portable built-in shading and border presets', async () => {
    const editor = await openTable()
    expect(
      applyTablePreset({
        headerFill: 'BDD7EE',
        band1Fill: 'DDEBF7',
        band2Fill: 'FFFFFF',
        borderColor: '9DC3E6',
      })(editor.state, editor.view.dispatch),
    ).toBe(true)
    const table = editor.getJSON().content![0] as unknown as PmNode
    expect(table.attrs?.tblStyleId).toBeNull()
    expect(table.content![0].content![0].attrs).toMatchObject({
      fill: 'BDD7EE',
      borders: { top: { style: 'single', color: '9DC3E6' } },
    })
    expect(table.content![1].content![0].attrs?.fill).toBe('DDEBF7')
    const model = pmTableToModel(table)
    expect(model.rows[0][0].fill).toBe('BDD7EE')
    expect(model.rows[1][0].borders?.left?.color).toBe('9DC3E6')
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
})
