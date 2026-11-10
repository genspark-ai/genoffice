import { Editor } from '@tiptap/core'
import { parseDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

async function htmlOf(bodyXml: string): Promise<string> {
  const parsed = await parseDocx(await buildDocx({ bodyXml }))
  const host = document.createElement('div')
  document.body.appendChild(host)
  const editor = new Editor({
    element: host,
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  const html = editor.getHTML()
  editor.destroy()
  host.remove()
  return html
}

describe('w:sym glyph font', () => {
  it('an undecoded private-use glyph renders in its symbol font, inside the run font span', async () => {
    const html = await htmlOf(
      '<w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr>' +
        '<w:t>a</w:t><w:sym w:font="Wingdings 2" w:char="F0FF"/><w:t>b</w:t></w:r></w:p>',
    )
    const sym = /<span[^>]*data-sym-char="F0FF"[^>]*>/.exec(html)?.[0] ?? ''
    expect(sym).toMatch(/font-family:\s*&quot;Wingdings 2&quot;/)
    expect(html.indexOf('Calibri')).toBeLessThan(html.indexOf('data-sym-char'))
  })

  it('a glyph decoded to Unicode keeps the run font', async () => {
    const html = await htmlOf('<w:p><w:r><w:sym w:font="Wingdings" w:char="F0FC"/></w:r></w:p>')
    const sym = /<span[^>]*data-sym-char="F0FC"[^>]*>/.exec(html)?.[0] ?? ''
    expect(sym).not.toContain('style=')
    expect(html).toContain('>\u2713<')
  })
})
