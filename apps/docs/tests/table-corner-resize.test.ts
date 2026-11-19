import { Editor } from '@tiptap/core'
import { parseDocx } from '@genoffice/docx-engine'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmTableToModel, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { scaleTableGrid } from '../src/renderer/editor/table-column-drag'
import { scaleRowHeights } from '../src/renderer/editor/table-handle'

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

describe('scaleTableGrid (Word corner drag)', () => {
  it('scales every column by the same factor', () => {
    expect(scaleTableGrid([100, 200, 300], -300, 900)).toEqual([50, 100, 150])
    expect(scaleTableGrid([100, 200, 300], 300, 900)).toEqual([150, 300, 450])
  })

  it('never grows past the content width and never shrinks below the column minimum', () => {
    expect(sum(scaleTableGrid([200, 200, 200], 1000, 624))).toBe(624)
    expect(sum(scaleTableGrid([200, 200, 200], -1000, 624))).toBe(120)
    // an uneven grid stops when its narrowest column hits the minimum
    expect(scaleTableGrid([100, 300], -1000, 624)).toEqual([40, 120])
    // an already over-wide grid may shrink but never grow
    expect(scaleTableGrid([300, 300, 300], 50, 624)).toEqual([300, 300, 300])
  })
})

async function openTable(): Promise<{ editor: Editor; table: HTMLTableElement }> {
  const parsed = await parseDocx(await buildDocx({ bodyXml: TABLE }))
  const element = document.createElement('div')
  document.body.appendChild(element)
  const editor = new Editor({
    element,
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  editor.view.dom.style.setProperty('--section-content-w', '900px')
  // jsdom has no hit-testing; prosemirror-tables' hover handler needs it to exist
  document.elementFromPoint ??= () => null
  return { editor, table: editor.view.dom.querySelector('table')! }
}

function gridOf(editor: Editor): number[] {
  const model = pmTableToModel(editor.getJSON().content![0] as unknown as PmNode)
  return model.colWidthsTwips!.map((twips) => Math.round(twips / 15))
}

function mouse(type: string, clientX: number, buttons = 1, clientY = 10): MouseEvent {
  return new MouseEvent(type, { clientX, clientY, buttons, button: 0, bubbles: true })
}

const offsetHeight = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'offsetHeight')!

/** jsdom lays nothing out: rows are 30px tall, the table 60px */
function layoutRows(): void {
  Object.defineProperty(HTMLElement.prototype, 'offsetHeight', {
    configurable: true,
    get(this: HTMLElement) {
      return this.tagName === 'TR' ? 30 : this.tagName === 'TABLE' ? 60 : 0
    },
  })
}

describe('scaleRowHeights (Word corner drag)', () => {
  it('scales every row by the same factor and stops at the row minimum', () => {
    expect(scaleRowHeights([20, 40], 30)).toEqual([30, 60])
    expect(scaleRowHeights([20, 40], -100)).toEqual([4, 8])
  })
})

describe('table corner resize handle', () => {
  afterEach(() => {
    vi.restoreAllMocks()
    Object.defineProperty(HTMLElement.prototype, 'offsetHeight', offsetHeight)
    document.body.innerHTML = ''
  })

  it('a vertical drag scales the row heights into w:trHeight', async () => {
    layoutRows()
    const { editor, table } = await openTable()
    table.querySelector('td')!.dispatchEvent(mouse('mousemove', 10))
    const resizer = document.querySelector<HTMLElement>('.doc-table-resize')!
    resizer.dispatchEvent(mouse('mousedown', 800, 1, 60))
    window.dispatchEvent(mouse('mousemove', 800, 1, 90))
    expect(Array.from(table.rows).map((tr) => tr.style.height)).toEqual(['45px', '45px'])
    window.dispatchEvent(mouse('mouseup', 800, 0, 90))
    const rows: { heightTwips: number; heightRule: string }[] = []
    editor.state.doc.descendants((node) => {
      if (node.type.name === 'docTableRow') rows.push(node.attrs as never)
    })
    expect(rows.map((row) => row.heightTwips)).toEqual([675, 675])
    expect(rows.every((row) => row.heightRule === 'atLeast')).toBe(true)
    // the grid was not touched
    expect(gridOf(editor)).toEqual([267, 267, 267])
    editor.destroy()
  })

  it('appears on hover and commits a proportionally scaled grid on drop', async () => {
    const { editor, table } = await openTable()
    const resizer = document.querySelector<HTMLElement>('.doc-table-resize')!
    expect(resizer.style.display).toBe('none')
    table.querySelector('td')!.dispatchEvent(mouse('mousemove', 10))
    expect(resizer.style.display).toBe('flex')

    vi.spyOn(table, 'getBoundingClientRect').mockReturnValue({
      width: 800,
      height: 40,
      top: 0,
      left: 0,
      right: 800,
      bottom: 40,
      x: 0,
      y: 0,
      toJSON: () => ({}),
    })
    resizer.dispatchEvent(mouse('mousedown', 800))
    window.dispatchEvent(mouse('mousemove', 600))
    expect(table.style.width).toBe('600px')
    window.dispatchEvent(mouse('mouseup', 600, 0))
    expect(gridOf(editor)).toEqual([200, 200, 200])
    editor.destroy()
  })

  it('turns an AutoFit-to-window table into a fixed grid so the new width shows', async () => {
    const { editor } = await openTable()
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(0, undefined, {
        ...editor.state.doc.firstChild!.attrs,
        tblAutoFit: 'window',
        widthPx: null,
        widthPct: 100,
      }),
    )
    const dom = editor.view.dom.querySelector('table')!
    expect(dom.className).toContain('doc-table-autofit-window')
    dom.querySelector('td')!.dispatchEvent(mouse('mousemove', 10))
    const resizer = document.querySelector<HTMLElement>('.doc-table-resize')!
    vi.spyOn(dom, 'getBoundingClientRect').mockReturnValue({
      width: 900,
      height: 40,
      top: 0,
      left: 0,
      right: 900,
      bottom: 40,
      x: 0,
      y: 0,
      toJSON: () => ({}),
    })
    resizer.dispatchEvent(mouse('mousedown', 900))
    window.dispatchEvent(mouse('mouseup', 600, 0))
    const attrs = editor.state.doc.firstChild!.attrs
    expect(attrs.tblAutoFit).toBe('fixed')
    expect(attrs.widthPct).toBeNull()
    expect(Math.round(attrs.widthPx)).toBe(600)
    expect(editor.view.dom.querySelector('table')!.className).toContain('doc-table-autofit-fixed')
    editor.destroy()
  })

  it('Escape cancels the drag and restores the table', async () => {
    const { editor, table } = await openTable()
    const resizer = document.querySelector<HTMLElement>('.doc-table-resize')!
    table.querySelector('td')!.dispatchEvent(mouse('mousemove', 10))
    const styleBefore = table.getAttribute('style')
    resizer.dispatchEvent(mouse('mousedown', 800))
    window.dispatchEvent(mouse('mousemove', 700))
    window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))
    expect(table.getAttribute('style')).toBe(styleBefore)
    expect(resizer.style.display).toBe('none')
    expect(gridOf(editor)).toEqual([267, 267, 267])
    editor.destroy()
  })
})
