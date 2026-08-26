import { Editor } from '@tiptap/core'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

const TOC =
  '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>' +
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText> TOC \\o "1-3" \\h </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:hyperlink w:anchor="_Toc1"><w:r><w:t>Old title</w:t></w:r>' +
  '<w:r><w:tab/></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText> PAGEREF _Toc1 \\h </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink></w:p>'

const FORMULA =
  '<w:p><m:oMath><m:f><m:num><m:r><m:t>a</m:t></m:r></m:num>' +
  '<m:den><m:r><m:t>b</m:t></m:r></m:den></m:f></m:oMath></w:p>'

describe('protected field and formula editing', () => {
  it('double-click edits visible text and round-trips through targeted XML patches', async () => {
    const source = await buildDocx({ bodyXml: TOC + FORMULA })
    const parsed = await parseDocx(source)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })

    const field = editor.view.dom.querySelector('.doc-protected-field') as HTMLElement
    const title = field.querySelector('.doc-toc-title') as HTMLElement
    const page = field.querySelector('.doc-toc-page') as HTMLElement
    expect(title.getAttribute('contenteditable')).toBe('false')
    title.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }))
    expect(field.draggable).toBe(false)
    title.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, button: 0 }))
    expect(title.getAttribute('contenteditable')).toBe('true')
    title.textContent = 'New & title'
    page.textContent = '12'
    document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }))

    const formula = editor.view.dom.querySelector('.doc-protected-formula') as HTMLElement
    const tokens = Array.from(formula.querySelectorAll<HTMLElement>('.doc-formula-token'))
    expect(formula.draggable).toBe(true)
    tokens[0].dispatchEvent(new MouseEvent('dblclick', { bubbles: true, button: 0 }))
    expect(tokens[0].getAttribute('contenteditable')).toBe('true')
    tokens[0].textContent = 'x'
    tokens[1].textContent = 'y + 1'
    document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }))

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(2)
    const reparsed = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    expect(reparsed.blocks[0].fieldDisplay).toMatchObject({
      left: 'New & title',
      right: '12',
    })
    expect(reparsed.blocks[1].formulaDisplay?.tokens).toEqual(['x', 'y + 1'])
    expect(reparsed.blocks[0].originalXml).toContain(' TOC \\o "1-3" \\h ')
    expect(reparsed.blocks[1].originalXml).toContain('<m:f><m:num>')
    editor.destroy()
  })

  it('keeps untouched field and formula blocks byte-stable', async () => {
    const source = await buildDocx({ bodyXml: TOC + FORMULA })
    const parsed = await parseDocx(source)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(0)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    expect(saved).toEqual(source)
    editor.destroy()
  })

  it('text-field passthrough carries run face/size and preserves spaces (public issue #118)', async () => {
    const rpr =
      '<w:rPr><w:rFonts w:ascii="\u5b8b\u4f53" w:eastAsia="\u5b8b\u4f53"/><w:sz w:val="24"/></w:rPr>'
    const bodyXml =
      `<w:p><w:pPr><w:jc w:val="left"/>${rpr}</w:pPr>` +
      `<w:r>${rpr}<w:t>\u9636\u8d70\u5230\u6cb3</w:t></w:r>` +
      `<w:r>${rpr}<w:t xml:space="preserve">     </w:t></w:r>` +
      `<w:r>${rpr}<w:fldChar w:fldCharType="begin"/></w:r>` +
      `<w:r>${rpr}<w:instrText xml:space="preserve"> INCLUDEPICTURE "/tmp/x.jpeg" \\* MERGEFORMATINET </w:instrText></w:r>` +
      `<w:r>${rpr}<w:fldChar w:fldCharType="end"/></w:r>` +
      `<w:r>${rpr}<w:t>\u5cb8\u8fb9</w:t></w:r></w:p>`
    const parsed = await parseDocx(await buildDocx({ bodyXml }))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const el = editor.view.dom.querySelector<HTMLElement>('.doc-field-text')!
    expect(el).toBeTruthy()
    // mid-paragraph space run must survive into the DOM (pre-wrap keeps it)
    expect(el.textContent).toContain('\u6cb3     \u5cb8')
    expect(el.style.fontSize).toBe('12pt')
    expect(el.style.getPropertyValue('--doc-line-factor')).toBe('1.3029')
    expect(el.style.lineHeight).toContain('--doc-line-grid')
    expect(el.style.textAlign).toBe('left')
    // inline: prosemirror-view injects a higher-specificity white-space:normal
    expect(el.style.whiteSpace).toBe('pre-wrap')
    editor.destroy()
  })
})
