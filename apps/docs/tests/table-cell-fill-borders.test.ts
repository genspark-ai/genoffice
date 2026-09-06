import { Editor } from '@tiptap/core'
import { parseDocx } from '@genoffice/docx-engine'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

const cell = (fill: string | null, text: string) =>
  '<w:tc><w:tcPr><w:tcW w:w="1700" w:type="dxa"/>' +
  '<w:tcBorders><w:top w:val="single" w:sz="4" w:color="9AA6B2"/><w:left w:val="single" w:sz="4" w:color="9AA6B2"/>' +
  '<w:bottom w:val="single" w:sz="4" w:color="9AA6B2"/><w:right w:val="single" w:sz="4" w:color="9AA6B2"/></w:tcBorders>' +
  (fill ? `<w:shd w:val="clear" w:fill="${fill}"/>` : '') +
  `</w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`

const TABLE =
  '<w:tbl><w:tblPr><w:tblBorders><w:insideH w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="1700"/><w:gridCol w:w="1700"/></w:tblGrid>' +
  `<w:tr>${cell('EFEFEF', 'A')}${cell(null, 'B')}</w:tr>` +
  `<w:tr>${cell('EFEFEF', 'C')}${cell(null, 'D')}</w:tr>` +
  '</w:tbl>'

describe('filled cells keep the collapsed seams', () => {
  it('declares every side on filled and unfilled cells alike and positions none of them', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: TABLE }))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const cells = [...editor.view.dom.querySelectorAll('table.doc-table td')] as HTMLElement[]
    expect(cells).toHaveLength(4)
    for (const td of cells) {
      expect(td.style.borderTop).toBe('1px solid rgb(154, 166, 178)')
      expect(td.style.borderBottom).toBe('1px solid rgb(154, 166, 178)')
      expect(td.style.position).toBe('')
    }
    expect(cells[0].style.backgroundColor).toBe('rgb(239, 239, 239)')
    editor.destroy()
  })

  it('never positions table cells in the stylesheet (a positioned fill paints over the borders)', () => {
    const css = readFileSync(resolve(__dirname, '../src/renderer/styles.css'), 'utf8')
    const block = css.match(/\n\.doc-table td,\n\.doc-table th \{([^}]*)\}/)
    expect(block).not.toBeNull()
    expect(block![1]).not.toMatch(/\bposition\s*:/)
    // the column resize handle needs its cell as containing block: positioned only while live
    expect(css).toMatch(
      /\.resize-cursor \.doc-table td,\s*\.resize-cursor \.doc-table th\s*\{\s*position: relative;/,
    )
  })
})
