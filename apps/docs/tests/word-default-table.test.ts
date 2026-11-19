import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { buildBlankDocx, generateTableModelXml, parseDocx } from '@genoffice/docx-engine'
import { insertTableAt } from '../src/renderer/components/ribbon-tabs'
import { parseHtmlFragment } from '../src/renderer/ai/protocol'
import { blocksToPmDoc, pmTableToModel } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import type { PmNode } from '../src/renderer/editor/convert'

async function openBlank() {
  const parsed = await parseDocx(await buildBlankDocx())
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
}

const cellParagraphs = (table: PmNode): PmNode[] =>
  (table.content ?? []).flatMap((row) =>
    (row.content ?? []).flatMap((cell) =>
      (cell.content ?? []).filter((n) => n.type === 'docParagraph'),
    ),
  )

function expectWordLook(table: PmNode) {
  const borders = table.attrs?.borders as Record<string, { style: string; szEighths: number }>
  for (const side of ['top', 'bottom', 'left', 'right', 'insideH', 'insideV'])
    expect(borders[side]).toMatchObject({ style: 'single', szEighths: 4 })
  for (const row of table.content ?? [])
    for (const cell of row.content ?? []) expect(cell.attrs?.vAlign).toBe('center')
  const paras = cellParagraphs(table)
  expect(paras.length).toBeGreaterThan(0)
  for (const p of paras) expect(p.attrs).toMatchObject({ spaceAfter: 0, lineSpacing: 1 })
  const xml = generateTableModelXml(pmTableToModel(table))
  expect(xml).toContain('<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>')
  expect(xml).toContain('<w:vAlign w:val="center"/>')
  expect(xml).toContain('<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>')
}

describe('new tables look like Word Insert Table', () => {
  it('ribbon insert: single borders, cell paragraphs at single spacing with no space after', async () => {
    const editor = await openBlank()
    insertTableAt(editor, 2, 3)
    expectWordLook(editor.state.doc.firstChild!.toJSON() as PmNode)
  })

  it('AI <table> fragment: same look, header row keeps bold + fill', () => {
    const [table] = parseHtmlFragment(
      '<table><tr><th>Item</th><th>Plan</th></tr><tr><td>Day 1</td><td>Museum</td></tr></table>',
      { bullet: null, ordered: null },
    )
    expectWordLook(table)
    const header = table.content![0].content![0]
    expect(header.type).toBe('docTableHeader')
    expect(header.attrs?.fill).toBe('F2F2F2')
    expect(header.content![0].content![0].marks?.some((m) => m.type === 'bold')).toBe(true)
  })
})
