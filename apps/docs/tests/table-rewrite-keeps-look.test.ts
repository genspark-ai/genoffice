import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { buildBlankDocx, parseDocx } from '@genoffice/docx-engine'
import { inheritTableFormatting } from '../src/renderer/ai/inherit-formatting'
import { parseHtmlFragment } from '../src/renderer/ai/protocol'
import { blocksToPmDoc, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

const NUM_IDS = { bullet: null, ordered: null }

async function editorWith(nodes: PmNode[]) {
  const parsed = await parseDocx(await buildBlankDocx())
  const doc = blocksToPmDoc(parsed.blocks) as PmNode
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { ...doc, content: nodes } as never,
  })
}

const vAligns = (table: PmNode) =>
  (table.content ?? []).flatMap((row) => (row.content ?? []).map((c) => c.attrs?.vAlign ?? null))

describe('rewriting a table keeps the old cell look', () => {
  it('a top-aligned table stays top-aligned when the AI adds a column', async () => {
    const [old] = parseHtmlFragment('<table><tr><td>a</td><td>b</td></tr></table>', NUM_IDS)
    for (const row of old.content!) for (const cell of row.content!) cell.attrs!.vAlign = null
    const editor = await editorWith([old, { type: 'docParagraph' }])
    const [next] = parseHtmlFragment(
      '<table><tr><td>a</td><td>b</td><td>c</td></tr></table>',
      NUM_IDS,
    )
    expect(vAligns(next)).toEqual(['center', 'center', 'center'])
    const merged = inheritTableFormatting(editor.state.doc.firstChild!, next, { anchor: true })
    expect(vAligns(merged)).toEqual([null, null, null])
  })
})
