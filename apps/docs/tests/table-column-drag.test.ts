import { Editor } from '@tiptap/core'
import { columnResizingPluginKey } from '@tiptap/pm/tables'
import { parseDocx } from '@genoffice/docx-engine'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmTableToModel, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { dragColumnBorder, tableColumnDragKey } from '../src/renderer/editor/table-column-drag'

// three 4000-twip columns: 266.67px each, 800px in all
const TABLE =
  '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
  '<w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc></w:tr>' +
  '<w:tr><w:tc><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:t>E</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:t>F</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'

const sum = (widths: number[]) => widths.reduce((total, width) => total + width, 0)

describe('dragColumnBorder (Word border drag)', () => {
  it('trades width between the two neighbouring columns and keeps the table width', () => {
    const widths = dragColumnBorder([200, 200, 200], 0, 50, 624)
    expect(widths).toEqual([250, 150, 200])
    expect(sum(widths)).toBe(600)
    expect(dragColumnBorder([200, 200, 200], 1, -80, 624)).toEqual([200, 120, 280])
  })

  it('stops at the minimum column width on either side of the border', () => {
    expect(dragColumnBorder([200, 200, 200], 0, 500, 624)).toEqual([360, 40, 200])
    expect(dragColumnBorder([200, 200, 200], 0, -500, 624)).toEqual([40, 360, 200])
  })

  it('moves the table edge with the last border, never past the content width', () => {
    expect(dragColumnBorder([200, 200, 200], 2, -50, 624)).toEqual([200, 200, 150])
    expect(dragColumnBorder([200, 200, 200], 2, 100, 624)).toEqual([200, 200, 224])
    // an already over-wide grid may shrink but never grow
    expect(dragColumnBorder([300, 300, 300], 2, 50, 624)).toEqual([300, 300, 300])
    expect(dragColumnBorder([300, 300, 300], 2, -50, 624)).toEqual([300, 300, 250])
  })

  it('ignores an out-of-range column', () => {
    expect(dragColumnBorder([200, 200], 5, 50, 624)).toEqual([200, 200])
  })
})

async function openTable(): Promise<{ editor: Editor; cells: number[]; table: HTMLTableElement }> {
  const parsed = await parseDocx(await buildDocx({ bodyXml: TABLE }))
  const element = document.createElement('div')
  document.body.appendChild(element)
  const editor = new Editor({
    element,
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  editor.view.dom.style.setProperty('--section-content-w', '900px')
  const cells: number[] = []
  editor.state.doc.descendants((node, pos) => {
    if (node.type.name === 'docTableCell') cells.push(pos)
  })
  return { editor, cells, table: editor.view.dom.querySelector('table')! }
}

function gridOf(editor: Editor): number[] {
  const model = pmTableToModel(editor.getJSON().content![0] as unknown as PmNode)
  return model.colWidthsTwips!.map((twips) => Math.round(twips / 15))
}

function mouse(type: string, clientX: number, buttons = 1): MouseEvent {
  return new MouseEvent(type, { clientX, buttons, button: 0, bubbles: true, cancelable: true })
}

function press(editor: Editor, clientX: number): boolean {
  const plugin = tableColumnDragKey.get(editor.state)!
  const mousedown = plugin.props.handleDOMEvents!.mousedown!
  return Boolean(mousedown.call(plugin, editor.view, mouse('mousedown', clientX)))
}

/** press on the border the resize plugin is hovering (its handle is the cell pos) */
function pressHandle(editor: Editor, cellPos: number, clientX: number): boolean {
  editor.view.dispatch(editor.state.tr.setMeta(columnResizingPluginKey, { setHandle: cellPos }))
  return press(editor, clientX)
}

describe('Word-style column border drag in the editor', () => {
  afterEach(() => {
    vi.restoreAllMocks()
    document.body.innerHTML = ''
  })

  it('previews and commits the same grid, following the pointer at 150% zoom', async () => {
    const { editor, cells, table } = await openTable()
    const before = gridOf(editor)
    expect(before).toEqual([267, 267, 267])
    // 800 model px drawn 1200 screen px wide (canvas zoom 150%)
    vi.spyOn(table, 'getBoundingClientRect').mockReturnValue({
      width: 1200,
      height: 40,
      top: 0,
      left: 0,
      right: 1200,
      bottom: 40,
      x: 0,
      y: 0,
      toJSON: () => ({}),
    })

    expect(pressHandle(editor, cells[0], 400)).toBe(true)
    // 150 screen px to the left = 100 model px
    window.dispatchEvent(mouse('mousemove', 250))
    const cols = Array.from(table.querySelectorAll('col')).map((col) => col.style.width)
    expect(cols.map((width) => Number.parseFloat(width))).toEqual([
      expect.closeTo(20.83, 1),
      expect.closeTo(45.83, 1),
      expect.closeTo(33.33, 1),
    ])

    window.dispatchEvent(mouse('mouseup', 250, 0))
    const after = gridOf(editor)
    expect(after).toEqual([167, 367, 267])
    // the table keeps its width: only the two neighbours traded space
    expect(sum(after)).toBe(sum(before))
    expect(columnResizingPluginKey.getState(editor.state)?.dragging).toBeFalsy()
    editor.destroy()
  })

  it('Escape cancels the drag and restores the table', async () => {
    const { editor, cells, table } = await openTable()
    const styleBefore = table.getAttribute('style')
    const colsBefore = Array.from(table.querySelectorAll('col')).map((col) =>
      col.getAttribute('style'),
    )
    expect(pressHandle(editor, cells[1], 500)).toBe(true)
    window.dispatchEvent(mouse('mousemove', 560))
    window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))
    expect(table.getAttribute('style')).toBe(styleBefore)
    expect(
      Array.from(table.querySelectorAll('col')).map((col) => col.getAttribute('style')),
    ).toEqual(colsBefore)
    window.dispatchEvent(mouse('mouseup', 600, 0))
    expect(gridOf(editor)).toEqual([267, 267, 267])
    editor.destroy()
  })

  it('leaves the table alone when no resize handle is live', async () => {
    const { editor } = await openTable()
    expect(press(editor, 10)).toBe(false)
    editor.destroy()
  })
})
