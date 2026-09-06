/** Explicit w:color auto renders as the default ink, never as a `#auto` colour. */
import { Editor } from '@tiptap/core'
import { parseDocx } from '@genoffice/docx-engine'
import type { ParsedDocFull, StyleDisplay, StyleInfo } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { docStyleCss } from '../src/renderer/doc-style-css'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const STYLES =
  '<w:style w:type="paragraph" w:styleId="RedBody"><w:name w:val="Red Body"/>' +
  '<w:rPr><w:color w:val="FF0000"/></w:rPr></w:style>'

const BODY =
  '<w:p><w:pPr><w:pStyle w:val="RedBody"/></w:pPr>' +
  '<w:r><w:rPr><w:color w:val="auto"/></w:rPr><w:t>automatic</w:t></w:r>' +
  '<w:r><w:t xml:space="preserve"> inherited</w:t></w:r></w:p>'

function parsedWith(styleId: string, display: StyleDisplay): ParsedDocFull {
  const styles = new Map<string, StyleInfo>()
  styles.set(styleId, { styleId, name: styleId, type: 'paragraph', display } as StyleInfo)
  return { styles, docDefaults: {}, blocks: [] } as unknown as ParsedDocFull
}

describe('w:color auto rendering', () => {
  it('paints an auto run in the paper ink over the style colour', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: BODY, extraStylesXml: STYLES }))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const spans = [...editor.view.dom.querySelectorAll('p span')] as HTMLElement[]
    const auto = spans.find((s) => s.textContent === 'automatic')
    expect(auto?.style.color).toBe('var(--docs-paper-ink)')
    expect(auto?.getAttribute('style')).not.toContain('--dk-c')
    editor.destroy()
  })

  it('emits the paper ink for a style whose colour is auto', () => {
    const css = docStyleCss(parsedWith('ResetBody', { color: 'auto' }))
    expect(css).toContain('[data-style="ResetBody"] { color:var(--docs-paper-ink) }')
    expect(css).not.toContain('#auto')
  })
})
