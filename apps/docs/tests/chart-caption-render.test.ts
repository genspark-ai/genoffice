/** chart + caption in one paragraph: plot on top, caption text under it, caption not field-editable */
import { Editor } from '@tiptap/core'
import { parseDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import {
  buildDocx,
  CHART_PARAGRAPH_XML,
  CHART_PART_XML,
  CHART_RELS,
} from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

const CAPTION =
  '<w:r><w:t xml:space="preserve">Figure </w:t></w:r>' +
  '<w:fldSimple w:instr=" SEQ Figure \\* ARABIC "><w:r><w:t>1</w:t></w:r></w:fldSimple>' +
  '<w:r><w:t xml:space="preserve">: Filler comparison</w:t></w:r>'

describe('chart caption rendering', () => {
  it('renders the SVG chart with the caption below and keeps the caption read-only', async () => {
    const parsed = await parseDocx(
      await buildDocx({
        bodyXml: CHART_PARAGRAPH_XML.replace('</w:r></w:p>', `</w:r>${CAPTION}</w:p>`),
        extraRels: CHART_RELS,
        extraParts: [
          {
            path: 'word/charts/chart1.xml',
            xml: CHART_PART_XML,
            contentType: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
          },
        ],
      }),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const chart = editor.view.dom.querySelector<HTMLElement>('.doc-protected-chart')
    expect(chart).toBeTruthy()
    expect(chart!.querySelector('.doc-chart-svg')).toBeTruthy()
    const caption = chart!.querySelector<HTMLElement>('.doc-chart-caption .doc-field-text')
    expect(caption?.textContent).toBe('Figure 1: Filler comparison')
    expect(chart!.querySelector('.doc-chart')!.compareDocumentPosition(caption!) & 4).toBeTruthy()

    caption!.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }))
    caption!.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, button: 0 }))
    expect(caption!.getAttribute('contenteditable')).toBe('false')
    editor.destroy()
  })
})
