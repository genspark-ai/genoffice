/** A cell paints only what its table style hands to bare runs. Colour and bold
 * aggregated from the runs stay on the runs, so clearing them really clears. */
import { Editor } from '@tiptap/core'
import { parseDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const RED = (t: string) =>
  `<w:r><w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr><w:t>${t}</w:t></w:r>`
const TABLE = (tblPr: string) =>
  `<w:tbl><w:tblPr>${tblPr}</w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>` +
  `<w:tr><w:tc><w:p>${RED('M/F')}${RED(' choose')}</w:p></w:tc></w:tr></w:tbl>`
const STYLE =
  '<w:style w:type="table" w:styleId="Tinted"><w:name w:val="Tinted"/>' +
  '<w:rPr><w:color w:val="365F91"/></w:rPr></w:style>'

async function open(bodyXml: string, extraStylesXml?: string) {
  const parsed = await parseDocx(await buildDocx({ bodyXml, extraStylesXml }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, td: editor.view.dom.querySelector('td') as HTMLElement }
}

function clearRuns(editor: Editor) {
  editor.commands.selectAll()
  editor.chain().setMark('docTextStyle', { color: null }).unsetMark('bold').run()
}

describe('uniformly formatted cell runs', () => {
  it('leave nothing on the cell, so clearing the runs clears the paint', async () => {
    const { editor, td } = await open(TABLE(''))
    expect(td.style.color).toBe('')
    expect(td.style.fontWeight).toBe('')
    expect((td.querySelector('span[data-doc-style]') as HTMLElement).style.color).toBe(
      'rgb(255, 0, 0)',
    )

    clearRuns(editor)
    for (const span of td.querySelectorAll<HTMLElement>('span[data-doc-style]')) {
      expect(span.style.color).toBe('')
    }
    expect(td.querySelector('strong')).toBeNull()
    expect(td.style.color).toBe('')
    editor.destroy()
  })

  it('still fall back to the table style colour once cleared', async () => {
    const { editor, td } = await open(TABLE('<w:tblStyle w:val="Tinted"/>'), STYLE)
    expect(td.style.color).toBe('rgb(54, 95, 145)')
    clearRuns(editor)
    expect(td.style.color).toBe('rgb(54, 95, 145)')
    editor.destroy()
  })
})
