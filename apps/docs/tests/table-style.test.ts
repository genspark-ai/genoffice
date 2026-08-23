import { Editor } from '@tiptap/core'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { docStyleCss } from '../src/renderer/doc-style-css'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const STYLE =
  '<w:style w:type="table" w:styleId="GridBlue"><w:name w:val="Grid Blue"/>' +
  '<w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr>' +
  '<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="4472C4"/></w:tcPr></w:tblStylePr>' +
  '</w:style>'

const TABLE =
  '<w:tbl><w:tblPr><w:tblStyle w:val="GridBlue"/><w:tblLook w:firstRow="1"/></w:tblPr>' +
  '<w:tr><w:tc><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:tc></w:tr>' +
  '<w:tr><w:tc><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:tc></w:tr>' +
  '</w:tbl>'

describe('table styles in the editor', () => {
  it('renders style-derived header fill and stays byte-identical untouched', async () => {
    const source = await buildDocx({ bodyXml: TABLE, extraStylesXml: STYLE })
    const parsed = await parseDocx(source)
    expect(parsed.blocks[0].table!.rows[0][0].fill).toBe('4472C4')

    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const headerCell = editor.view.dom.querySelector('table.doc-table th, table.doc-table td')
    expect((headerCell as HTMLElement | null)?.style.background).toContain('rgb(68, 114, 196)')

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(0)
    expect(await saveDocx(parsed, plan.saveBlocks)).toEqual(source)
    editor.destroy()
  })

  it('whole-table rPr/pPr (sz, jc, spacing) reach cell text via docStyleCss (LO calendar family)', async () => {
    // Calendar2-shaped style: whole-table 14pt centered text against 32pt docDefaults
    const style =
      '<w:style w:type="table" w:styleId="Calendar2"><w:name w:val="Calendar 2"/>' +
      '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>' +
      '<w:rPr><w:sz w:val="28"/></w:rPr>' +
      '<w:tblStylePr w:type="firstRow"><w:rPr><w:sz w:val="32"/><w:color w:val="5B9BD5"/></w:rPr></w:tblStylePr>' +
      '</w:style>'
    const table =
      '<w:tbl><w:tblPr><w:tblStyle w:val="Calendar2"/><w:tblLook w:firstRow="1"/></w:tblPr>' +
      '<w:tr><w:tc><w:p><w:r><w:t>Mon</w:t></w:r></w:p></w:tc></w:tr>' +
      '<w:tr><w:tc><w:p><w:r><w:t>1</w:t></w:r></w:p></w:tc></w:tr>' +
      '</w:tbl>'
    const parsed = await parseDocx(await buildDocx({ bodyXml: table, extraStylesXml: style }))
    const display = parsed.styles.get('Calendar2')!.tableDisplay!
    expect(display.wholeTable?.sizeHalfPoints).toBe(28)
    expect(display.paraJc).toBe('center')
    expect(display.firstRow?.sizeHalfPoints).toBe(32)

    const css = docStyleCss(parsed as never)
    const sel = '.doc-page table[data-tbl-style="Calendar2"]'
    expect(css).toContain(`${sel} td, ${sel} th { font-size:14pt }`)
    expect(css).toMatch(/tr:first-child td[^{]*\{[^}]*font-size:16pt/)
    expect(css).toMatch(/td p[^{]*\{[^}]*text-align:center/)
    expect(css).toMatch(/td p[^{]*\{[^}]*margin-bottom:0/)
  })
})
