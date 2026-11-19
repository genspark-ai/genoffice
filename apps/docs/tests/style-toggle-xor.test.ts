/**
 * ECMA-376 17.7.3: a toggle both the paragraph style and the character style
 * turn on is off. The CSS cascade (char-style rule over para-style rule) would
 * keep it on, so the engine's resolved value is painted as a doc-cs-nb class.
 */
import { Editor } from '@tiptap/core'
import { parseDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

const styles =
  '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:rPr><w:b/></w:rPr></w:style>' +
  '<w:style w:type="character" w:styleId="Heading3Char"><w:name w:val="Heading 3 Char"/><w:rPr><w:b/><w:bCs/></w:rPr></w:style>' +
  '<w:style w:type="character" w:styleId="Em"><w:name w:val="Em"/><w:rPr><w:i/></w:rPr></w:style>'

const run = (rPr: string, text: string) =>
  `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t>${text}</w:t></w:r>`

describe('style toggle XOR paint', () => {
  it('paints bold-on-bold regular, leaves single-layer toggles to the cascade', async () => {
    const body =
      '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>' +
      run('<w:rStyle w:val="Heading3Char"/>', 'both') +
      run('<w:rStyle w:val="Heading3Char"/><w:b/>', 'direct') +
      run('<w:rStyle w:val="Em"/>', 'em') +
      run('', 'para') +
      '</w:p>'
    const parsed = await parseDocx(await buildDocx({ bodyXml: body, extraStylesXml: styles }))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const spanOf = (text: string) =>
      [...editor.view.dom.querySelectorAll('span[data-style]')].find(
        (el) => el.textContent === text,
      )!
    expect(spanOf('both').classList.contains('doc-cs-nb')).toBe(true)
    expect(spanOf('both').classList.contains('doc-cs')).toBe(false)
    expect(spanOf('direct').classList.contains('doc-cs-nb')).toBe(false)
    expect(spanOf('direct').closest('strong')).not.toBeNull()
    expect(spanOf('em').className).toBe('')
    expect(editor.view.dom.querySelector('span.doc-cs-nb')?.textContent).toBe('both')
    editor.destroy()
  })
})
