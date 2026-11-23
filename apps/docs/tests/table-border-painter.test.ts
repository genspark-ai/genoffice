import { Editor } from '@tiptap/core'
import { parseDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { nearestSide, paintEdge } from '../src/renderer/editor/table-border-painter'

const cell = (text: string) => `<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`
const TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
  `<w:tr>${cell('A')}${cell('B')}</w:tr><w:tr>${cell('C')}${cell('D')}</w:tr></w:tbl>`

async function open(): Promise<Editor> {
  const parsed = await parseDocx(await buildDocx({ bodyXml: TABLE }))
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
}

function cellPositions(editor: Editor): number[] {
  const out: number[] = []
  editor.state.doc.descendants((node, pos) => {
    if (node.type.name === 'docTableCell') out.push(pos)
  })
  return out
}

describe('Border Painter', () => {
  it('takes the cell edge nearest the pointer', () => {
    const box = { left: 0, top: 0, right: 100, bottom: 40 }
    expect(nearestSide(box, 50, 2)).toBe('top')
    expect(nearestSide(box, 50, 39)).toBe('bottom')
    expect(nearestSide(box, 1, 20)).toBe('left')
    expect(nearestSide(box, 97, 20)).toBe('right')
  })

  it('draws the pen on the edge and on the neighbour across it', async () => {
    const editor = await open()
    const pen = { style: 'double', szEighths: 12, color: 'C00000' }
    const [a, , c] = cellPositions(editor)
    paintEdge(editor.view, a, 'right', pen)
    paintEdge(editor.view, cellPositions(editor)[2], 'top', pen)
    const rows = (editor.getJSON().content![0] as unknown as PmNode).content!
    expect(rows[0].content![0].attrs!.borders).toMatchObject({ right: pen, bottom: pen })
    expect(rows[0].content![1].attrs!.borders).toMatchObject({ left: pen })
    expect(rows[1].content![0].attrs!.borders).toMatchObject({ top: pen })
    expect(c).toBeGreaterThan(a)
    editor.destroy()
  })

  it('starts off with a black half-point pen', async () => {
    const editor = await open()
    expect(editor.storage.tableBorderPainter).toMatchObject({
      active: false,
      pen: { style: 'single', szEighths: 4, color: '000000' },
    })
    editor.destroy()
  })
})
