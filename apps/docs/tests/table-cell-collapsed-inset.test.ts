import { Editor } from '@tiptap/core'
import { parseDocx } from '@genoffice/docx-engine'
import type { TableModel } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { cellPadPx, collapsedEdgePx } from '../src/renderer/editor/border-metrics'
import { blocksToPmDoc, collapsedCellBw } from '../src/renderer/editor/convert'
import {
  editorExtensions,
  modelEdgeBorders,
  outerBorderPx,
  tableEdgeBorders,
} from '../src/renderer/editor/extensions'
import { renderTableSpec } from '../src/renderer/editor/protected-render'

const single = { style: 'single', szEighths: 4 }
const nil = { style: 'nil' }

describe('collapsed cell edge width', () => {
  it('takes the cell line, else the neighbour line that wins the collapse', () => {
    expect(collapsedEdgePx(single, undefined)).toBe(1)
    expect(collapsedEdgePx(nil, single)).toBe(1)
    expect(collapsedEdgePx(undefined, { style: 'single', szEighths: 16 })).toBe(3)
    expect(collapsedEdgePx(nil, undefined)).toBe(0)
    expect(collapsedEdgePx(nil, nil)).toBe(0)
    expect(collapsedEdgePx(undefined, undefined)).toBeUndefined()
  })

  it('resolves both sides of a row cell from its neighbours', () => {
    const row: TableModel['rows'][number] = [
      { paras: ['a'], borders: { right: single } },
      { paras: ['b'], borders: { left: nil, right: single } },
      { paras: ['c'] },
    ]
    expect(collapsedCellBw(row, 0)).toEqual({ r: 1 })
    expect(collapsedCellBw(row, 1)).toEqual({ l: 1, r: 1 })
    expect(collapsedCellBw(row, 2)).toEqual({ l: 1 })
  })

  it('keeps the 0.667px of a 70-twip cell margin', () => {
    expect(cellPadPx(70)).toBe('4.667px')
    expect(cellPadPx(108)).toBe('7.2px')
    expect(cellPadPx(0)).toBe('0px')
  })

  it('edge cells widen the collapsed table box when tblBorders are silent', () => {
    expect(outerBorderPx(null)).toBe(0)
    expect(outerBorderPx(null, { left: [single, single], right: [single, nil] })).toBe(1)
    expect(
      outerBorderPx({ left: single }, { left: [{ style: 'single', szEighths: 16 }], right: [] }),
    ).toBe(1.5)
  })
})

// fixed-layout questionnaire table: tcW 3681 / 10313, tblCellMar 70, sz 4 tcBorders
// only, the answer cell's own left side nil (the question cell draws that line)
const questionnaire =
  '<w:tbl><w:tblPr><w:tblW w:w="13994" w:type="dxa"/><w:tblLayout w:type="fixed"/>' +
  '<w:tblCellMar><w:left w:w="70" w:type="dxa"/><w:right w:w="70" w:type="dxa"/></w:tblCellMar></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="3681"/><w:gridCol w:w="10313"/></w:tblGrid>' +
  '<w:tr><w:tc><w:tcPr><w:tcW w:w="3681" w:type="dxa"/><w:tcBorders>' +
  '<w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/>' +
  '</w:tcBorders></w:tcPr><w:p><w:r><w:t>Q</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:tcPr><w:tcW w:w="10313" w:type="dxa"/><w:tcBorders>' +
  '<w:top w:val="single" w:sz="4"/><w:left w:val="nil"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/>' +
  '</w:tcBorders></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'

describe('fixed table text area equals tcW minus the cell margins', () => {
  it('editor: the nil-sided cell absorbs the neighbour line and the box spans the outer halves', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: questionnaire }))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const table = editor.view.dom.querySelector('table.doc-table') as HTMLElement
    // 13994 twips = 933px grid + half of the two 1px outer lines
    expect(table.getAttribute('style')).toContain('width: 934px')
    expect(table.style.getPropertyValue('--doc-cell-pad-l')).toBe('4.667px')
    const [q, a] = [...table.querySelectorAll('td')] as HTMLElement[]
    expect(q.style.getPropertyValue('--cell-bw-l')).toBe('1px')
    expect(a.style.getPropertyValue('--cell-bw-l')).toBe('1px')
    expect(a.style.getPropertyValue('--cell-bw-r')).toBe('1px')
    editor.destroy()
  })

  it('read-only render: same inset variables and box width', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: questionnaire }))
    const spec = renderTableSpec(parsed.blocks[0].table as TableModel)
    const html = JSON.stringify(spec)
    expect(html).toContain('width:934px')
    expect(html).toContain('--doc-cell-pad-l:4.667px')
    expect(html.match(/--cell-bw-l:1px/g)).toHaveLength(2)
  })
})

// row 2 opens with a vMerge continuation (dropped from the PM row), so its first PM
// child is the interior cell whose heavy left line must not count as a table edge
const tc = (w: number, borders: string, text: string, tcPr = '') =>
  `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${tcPr}<w:tcBorders>${borders}</w:tcBorders></w:tcPr>` +
  `<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`
const thin = (side: string) => `<w:${side} w:val="single" w:sz="4"/>`
const vMerged =
  '<w:tbl><w:tblPr><w:tblW w:w="5100" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="2100"/><w:gridCol w:w="3000"/></w:tblGrid>' +
  '<w:tr>' +
  tc(
    2100,
    thin('top') + thin('left') + thin('bottom') + thin('right'),
    'A',
    '<w:vMerge w:val="restart"/>',
  ) +
  tc(3000, thin('top') + thin('bottom') + thin('right'), 'B') +
  '</w:tr><w:tr>' +
  tc(2100, thin('left') + thin('bottom'), '', '<w:vMerge/>') +
  tc(3000, '<w:left w:val="single" w:sz="24"/>' + thin('bottom') + thin('right'), 'C') +
  '</w:tr></w:tbl>'

describe('edge borders follow grid columns, not child order', () => {
  it('editor: a heavy interior line after a vertical span does not widen the box', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: vMerged }))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    let tableNode = null as import('@tiptap/pm/model').Node | null
    editor.state.doc.descendants((n) => {
      if (n.type.name === 'docTable') tableNode = n
      return !tableNode
    })
    const edges = tableEdgeBorders(tableNode!)
    expect(edges.left).toHaveLength(1)
    expect(edges.right).toHaveLength(2)
    expect(outerBorderPx(null, edges)).toBe(1)
    // 5100 twips = 340px grid + half of the two 1px outer lines
    const table = editor.view.dom.querySelector('table.doc-table') as HTMLElement
    expect(table.getAttribute('style')).toContain('width: 341px')
    editor.destroy()
  })

  it('read-only render agrees', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: vMerged }))
    const model = parsed.blocks[0].table as TableModel
    expect(modelEdgeBorders(model).left).toHaveLength(1)
    expect(JSON.stringify(renderTableSpec(model))).toContain('width:341px')
  })
})
