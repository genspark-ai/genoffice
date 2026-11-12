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
    expect(css).toContain(`${sel} td, ${sel} th { font-size:14pt;--doc-base-fs:14pt }`)
    expect(css).toMatch(/tr:first-child td[^{]*\{[^}]*font-size:16pt/)
    expect(css).toMatch(/td p[^{]*\{[^}]*text-align:center/)
    expect(css).toMatch(/td p[^{]*\{[^}]*margin-bottom:0/)
  })

  it('conditional tblStylePr pPr paints inline on the cell paragraph (display only, save untouched)', async () => {
    const style =
      '<w:style w:type="table" w:styleId="Cond"><w:name w:val="Cond"/>' +
      '<w:tblStylePr w:type="firstRow"><w:pPr><w:spacing w:before="0" w:after="0"/><w:jc w:val="center"/></w:pPr></w:tblStylePr>' +
      '</w:style>'
    const table =
      '<w:tbl><w:tblPr><w:tblStyle w:val="Cond"/><w:tblLook w:firstRow="1"/></w:tblPr>' +
      '<w:tr><w:tc><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:tc></w:tr>' +
      '<w:tr><w:tc><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:tc></w:tr>' +
      '</w:tbl>'
    const source = await buildDocx({ bodyXml: table, extraStylesXml: style })
    const parsed = await parseDocx(source)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const [head, body] = [...editor.view.dom.querySelectorAll('table.doc-table p')] as HTMLElement[]
    expect(head.style.textAlign).toBe('center')
    expect(head.style.marginBottom).toBe('0pt')
    expect(body.style.textAlign).toBe('')
    expect(body.style.marginBottom).toBe('')

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(0)
    expect(await saveDocx(parsed, plan.saveBlocks)).toEqual(source)
    editor.destroy()
  })

  it('an rtl run in a conditional cell keeps the table-style weight (resolved, not a style counter)', async () => {
    const styles =
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>' +
      '<w:rPr><w:b/></w:rPr></w:style>' +
      '<w:style w:type="table" w:styleId="HeadBold"><w:name w:val="HeadBold"/>' +
      '<w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:bCs/></w:rPr></w:tblStylePr>' +
      '</w:style>'
    const rtl = (text: string) =>
      `<w:tc><w:p><w:r><w:rPr><w:rtl/></w:rPr><w:t>${text}</w:t></w:r></w:p></w:tc>`
    const table =
      '<w:tbl><w:tblPr><w:tblStyle w:val="HeadBold"/><w:tblLook w:firstRow="1"/></w:tblPr>' +
      `<w:tr>${rtl('Header')}</w:tr><w:tr>${rtl('Body')}</w:tr></w:tbl>`
    const parsed = await parseDocx(await buildDocx({ bodyXml: table, extraStylesXml: styles }))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const [head, body] = [...editor.view.dom.querySelectorAll('table.doc-table span.doc-cs')]
    // first row: table-style bCs on; the Normal w:b without a bCs twin never reaches an rtl run
    expect(head.classList.contains('doc-cs-b')).toBe(true)
    expect(body.classList.contains('doc-cs-nb')).toBe(true)
    expect(head.getAttribute('style') ?? '').not.toContain('font-weight')
    const css = docStyleCss(parsed as never)
    expect(css).not.toMatch(/\.doc-cs[^{]*\{[^}]*font-weight/)
    editor.destroy()
  })
})
