import { Editor } from '@tiptap/core'
import { parseDocx } from '@genoffice/docx-engine'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmTableToModel, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  draggedRowHeightTwips,
  rowChromePx,
  tableRowDragKey,
} from '../src/renderer/editor/table-row-drag'

const TABLE =
  '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
  '<w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>' +
  '<w:tr><w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'

const NESTED =
  '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid>' +
  '<w:tr><w:tc>' +
  '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
  '<w:tr><w:tc><w:p><w:r><w:t>inner</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' +
  '<w:p><w:r><w:t>outer</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'

const rect = (top: number, height: number, width = 500) =>
  ({
    top,
    bottom: top + height,
    height,
    left: 0,
    right: width,
    width,
    x: 0,
    y: top,
    toJSON: () => ({}),
  }) as DOMRect

async function openTable(
  bodyXml = TABLE,
): Promise<{ editor: Editor; rows: HTMLTableRowElement[] }> {
  const parsed = await parseDocx(await buildDocx({ bodyXml }))
  const element = document.createElement('div')
  document.body.appendChild(element)
  const editor = new Editor({
    element,
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  const rows = Array.from(editor.view.dom.querySelectorAll<HTMLTableRowElement>('tbody > tr'))
  // jsdom has no layout: two 20px rows, drawn at 150% zoom
  rows.forEach((tr, i) => {
    vi.spyOn(tr, 'getBoundingClientRect').mockReturnValue(rect(i * 30, 30))
    Object.defineProperty(tr, 'offsetHeight', { value: 20, configurable: true })
  })
  return { editor, rows }
}

/** the decoration redraws the <tr>: always look the live element up */
const rowDom = (editor: Editor, i: number) =>
  editor.view.dom.querySelectorAll<HTMLTableRowElement>('tbody > tr')[i]

const rowHeights = (editor: Editor) =>
  pmTableToModel(editor.getJSON().content![0] as unknown as PmNode).rowHeightsTwips

function mouse(type: string, target: Element | null, clientY: number, buttons = 1): MouseEvent {
  const event = new MouseEvent(type, {
    clientY,
    buttons,
    button: 0,
    bubbles: true,
    cancelable: true,
  })
  if (target) Object.defineProperty(event, 'target', { value: target })
  return event
}

function press(editor: Editor, cell: Element, clientY: number): boolean {
  const plugin = tableRowDragKey.get(editor.state)!
  const mousedown = plugin.props.handleDOMEvents!.mousedown!
  return Boolean(mousedown.call(plugin, editor.view, mouse('mousedown', cell, clientY)))
}

describe('rowChromePx', () => {
  it('reads the authored height expression and zeroes the w:trHeight term', () => {
    const tr = document.createElement('tr')
    const td = document.createElement('td')
    tr.appendChild(td)
    // 600 twips = 40.0px; CSSOM would serialize the token as 40px
    tr.setAttribute('style', '--doc-row-pad-t:10px;height:calc(40.0px + 10px + 10px)')
    let probeExpr = ''
    const append = td.appendChild.bind(td)
    vi.spyOn(td, 'appendChild').mockImplementation((node) => {
      probeExpr = (node as HTMLElement).style.height
      vi.spyOn(node as HTMLElement, 'getBoundingClientRect').mockReturnValue(rect(0, 20))
      return append(node)
    })
    expect(rowChromePx(tr, 600)).toBe(20)
    // jsdom folds the probe's calc(); the w:trHeight term is gone either way
    expect(probeExpr).not.toContain('40')
    expect(rowChromePx(tr, 601)).toBe(0)
  })
})

describe('draggedRowHeightTwips', () => {
  it('adds the pointer delta to the drawn height and never goes below the floor', () => {
    expect(draggedRowHeightTwips(20, 10)).toBe(450)
    expect(draggedRowHeightTwips(20, -100)).toBe(60)
  })
})

describe('Word-style row border drag in the editor', () => {
  afterEach(() => {
    vi.restoreAllMocks()
    document.body.innerHTML = ''
  })

  it('previews on the row and commits w:trHeight from the pointer travel at zoom', async () => {
    const { editor, rows } = await openTable()
    expect(rowHeights(editor)).toBeUndefined()
    const cell = rows[0].querySelector('td')!
    // 2px above the first row's bottom border (y=30)
    expect(press(editor, cell, 28)).toBe(true)
    window.dispatchEvent(mouse('mousemove', null, 58))
    // 30 screen px = 20 css px at 150%
    expect(Number.parseFloat(rowDom(editor, 0).style.height)).toBeCloseTo(40, 1)
    window.dispatchEvent(mouse('mousemove', null, 58, 0))
    expect(rowHeights(editor)).toEqual([600, null])
    expect((editor.getJSON().content![0] as unknown as PmNode).content![0].attrs!.heightRule).toBe(
      'atLeast',
    )
    editor.destroy()
  })

  it('leaves a nested table row alone (its border is not the outer row)', async () => {
    const { editor, rows } = await openTable(NESTED)
    const inner = rows.find((tr) => tr.closest('table') !== rows[0].closest('table'))!
    expect(inner).toBeTruthy()
    vi.spyOn(inner, 'getBoundingClientRect').mockReturnValue(rect(0, 10))
    expect(press(editor, inner.querySelector('td')!, 9)).toBe(false)
    expect(rowHeights(editor)).toBeUndefined()
    editor.destroy()
  })

  it('ignores presses away from a border and Escape restores the row', async () => {
    const { editor, rows } = await openTable()
    const cell = rows[1].querySelector('td')!
    expect(press(editor, cell, 45)).toBe(false)
    expect(press(editor, cell, 59)).toBe(true)
    window.dispatchEvent(mouse('mousemove', null, 90))
    expect(rowDom(editor, 1).style.height).not.toBe('')
    window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))
    expect(rowDom(editor, 1).style.height).toBe('')
    window.dispatchEvent(mouse('mouseup', null, 90, 0))
    expect(rowHeights(editor)).toBeUndefined()
    editor.destroy()
  })
})
