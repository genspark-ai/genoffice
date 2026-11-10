/** w:bdr character borders reach the docTextStyle span; identical adjacent runs share one box */
import { Editor } from '@tiptap/core'
import { parseDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, inlineToRuns, runsToInline } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { runBorderDecls } from '../src/renderer/editor/run-border'

const BDR = '<w:bdr w:val="single" w:sz="6" w:space="1" w:color="000000"/>'
const run = (text: string, rPr = BDR) =>
  `<w:r><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="18"/>${rPr}</w:rPr>` +
  `<w:t xml:space="preserve">${text}</w:t></w:r>`

describe('run border rendering', () => {
  it('draws one bordered span over adjacent runs with the same border', async () => {
    const body =
      '<w:p>' +
      '<w:r><w:t xml:space="preserve">answer </w:t></w:r>' +
      run(' ') +
      run('(') +
      run('28') +
      run(')') +
      run(' ') +
      '<w:r><w:t xml:space="preserve"> here</w:t></w:r>' +
      '</w:p>'
    const parsed = await parseDocx(await buildDocx({ bodyXml: body }))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const boxed = Array.from(
      editor.view.dom.querySelectorAll<HTMLElement>('span[style*="border:"]'),
    )
    expect(boxed.length).toBe(1)
    expect(boxed[0].textContent).toBe(' (28) ')
    expect(boxed[0].style.border).toBe('0.75pt solid rgb(0, 0, 0)')
    expect(boxed[0].style.padding).toBe('1pt')
    expect(boxed[0].style.getPropertyValue('--dk-b-t')).toBe('0.75pt solid #ffffff')
    editor.destroy()
  })

  it('round-trips run.bdr through the mark', () => {
    const bdr = { val: 'single', sz: 6, color: '000000', space: 1 }
    const inline = runsToInline([{ text: 'x', bdr }])
    const mark = inline[0].marks?.find((m) => m.type === 'docTextStyle')
    expect(mark?.attrs?.bdr).toBe(JSON.stringify(bdr))
    expect(inlineToRuns(inline)[0].bdr).toEqual(bdr)
  })

  it('maps Word border styles and the auto color', () => {
    expect(runBorderDecls(JSON.stringify({ val: 'dashed', sz: 12 }))).toEqual([
      'border:1.5pt dashed currentColor',
    ])
    const decls = runBorderDecls(
      JSON.stringify({ val: 'double', sz: 1, color: 'FF0000', space: 2 }),
    )
    expect(decls[0]).toBe('border:0.25pt double #FF0000')
    expect(decls[1]).toBe('padding:2pt')
    expect(decls.filter((d) => d.startsWith('--dk-b-')).length).toBe(4)
    expect(runBorderDecls('not json')).toEqual([])
  })
})
