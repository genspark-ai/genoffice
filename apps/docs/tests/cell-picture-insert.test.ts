import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { buildBlankDocx, parseDocx, saveDocx } from '@genoffice/docx-engine'
import { insertTableAt } from '../src/renderer/components/ribbon-tabs'
import { insertPictureAtSelection } from '../src/renderer/editor/insert-picture'
import {
  blocksToPmDoc,
  pmDocToSavePlan,
  tableFormattingSignature,
  type PmNode,
} from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

const TINY_PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='

const picture = {
  dataUrl: `data:image/png;base64,${TINY_PNG_BASE64}`,
  base64: TINY_PNG_BASE64,
  mime: 'image/png',
  widthPx: 120,
  heightPx: 80,
  label: 'Picture',
}

async function openBlank() {
  const source = await buildBlankDocx()
  const parsed = await parseDocx(source)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, parsed }
}

function selectFirstCellParagraph(editor: Editor) {
  let target = -1
  editor.state.doc.descendants((node, pos) => {
    if (target >= 0) return false
    if (node.type.name === 'docTableCell') {
      target = pos + 2
      return false
    }
    return true
  })
  expect(target).toBeGreaterThan(0)
  editor.commands.setTextSelection(target)
}

async function save(editor: Editor, parsed: Awaited<ReturnType<typeof parseDocx>>) {
  const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
  const bytes = await saveDocx(parsed, plan.saveBlocks, { newPictures: plan.newPictures })
  const zip = await JSZip.loadAsync(bytes)
  return { plan, zip, xml: await zip.file('word/document.xml')!.async('string') }
}

describe('insert picture in a table cell', () => {
  it('becomes an inline picture run of the cell paragraph and saves as cell media', async () => {
    const { editor, parsed } = await openBlank()
    insertTableAt(editor, 2, 2)
    selectFirstCellParagraph(editor)
    expect(insertPictureAtSelection(editor, picture)).toBe(true)

    let cellPictures = 0
    let blockPictures = 0
    editor.state.doc.descendants((node, _pos, parent) => {
      if (node.type.name === 'docInlineImage') {
        cellPictures++
        expect(parent?.type.name).toBe('docParagraph')
        expect(node.attrs.widthPx).toBe(120)
        expect(node.attrs.heightPx).toBe(80)
      }
      if (node.type.name === 'docProtected' && node.attrs.blockType === 'image') blockPictures++
      return true
    })
    expect(cellPictures).toBe(1)
    expect(blockPictures).toBe(0)
    expect(editor.state.doc.firstChild?.type.name).toBe('docTable')

    const { plan, zip, xml } = await save(editor, parsed)
    expect(Object.keys(plan.newPictures)).toHaveLength(1)
    const tbl = xml.slice(xml.indexOf('<w:tbl>'), xml.indexOf('</w:tbl>'))
    expect(tbl).toContain('<w:drawing>')
    expect(tbl).toContain('<wp:extent cx="1143000" cy="762000"/>')
    expect(xml).not.toContain('aidocs-new-pic-')
    expect(Object.keys(zip.files).some((f) => /^word\/media\/aidocs\d+\.png$/.test(f))).toBe(true)
  })

  it('saves a picture added to a cell of a table that came from the file', async () => {
    const first = await openBlank()
    insertTableAt(first.editor, 1, 2)
    selectFirstCellParagraph(first.editor)
    first.editor.commands.insertContent('a')
    const saved = await save(first.editor, first.parsed)
    const parsed = await parseDocx(await saved.zip.generateAsync({ type: 'uint8array' }))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    selectFirstCellParagraph(editor)
    expect(insertPictureAtSelection(editor, picture)).toBe(true)
    const { xml, zip } = await save(editor, parsed)
    const tbl = xml.slice(xml.indexOf('<w:tbl>'), xml.indexOf('</w:tbl>'))
    expect(tbl).toContain('<w:drawing>')
    expect(tbl).toContain('<w:t xml:space="preserve">a</w:t>')
    expect(Object.keys(zip.files).some((f) => /^word\/media\/aidocs\d+\.png$/.test(f))).toBe(true)
  })

  it('gives a picture its own token when the plan is rebuilt from the same nodes', async () => {
    const { editor, parsed } = await openBlank()
    insertTableAt(editor, 1, 2)
    selectFirstCellParagraph(editor)
    expect(insertPictureAtSelection(editor, picture)).toBe(true)
    const first = editor.getJSON() as PmNode
    expect(Object.keys(pmDocToSavePlan(first, parsed.blocks).newPictures)).toHaveLength(1)

    const table = first.content![0]
    const row = table.content![0]
    const [cellA, cellB] = row.content!
    const pictureNode = cellA.content![0].content![0]
    const second: PmNode = {
      ...first,
      content: [
        {
          ...table,
          content: [
            {
              ...row,
              content: [
                cellA,
                { ...cellB, content: [{ type: 'docParagraph', content: [{ ...pictureNode }] }] },
              ],
            },
          ],
        },
        ...first.content!.slice(1),
      ],
    }
    expect(Object.keys(pmDocToSavePlan(second, parsed.blocks).newPictures)).toHaveLength(2)
  })

  it('a second picture next to an existing one changes the table signature', () => {
    const img = { type: 'docInlineImage', attrs: { dataUrl: 'data:x', xml: '<w:drawing/>' } }
    const table = (inline: PmNode[]): PmNode =>
      ({
        type: 'docTable',
        content: [
          {
            type: 'docTableRow',
            content: [
              { type: 'docTableCell', content: [{ type: 'docParagraph', content: inline }] },
            ],
          },
        ],
      }) as PmNode
    expect(tableFormattingSignature(table([img as PmNode]))).not.toBe(
      tableFormattingSignature(
        table([img as PmNode, { ...img, attrs: { ...img.attrs, xml: '' } } as PmNode]),
      ),
    )
  })

  it('keeps the body insert as a protected picture block', async () => {
    const { editor } = await openBlank()
    expect(insertPictureAtSelection(editor, picture)).toBe(true)
    let blockPictures = 0
    editor.state.doc.descendants((node) => {
      if (node.type.name === 'docProtected' && node.attrs.blockType === 'image') blockPictures++
      return true
    })
    expect(blockPictures).toBe(1)
  })
})
