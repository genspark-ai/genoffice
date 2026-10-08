/**
 * `marked` is a Markdown-to-HTML converter, not a sanitiser: it copies link
 * targets and raw inline HTML straight through. A clipboard can therefore hand
 * the docs editor an executable payload that CSP does not cover — a
 * `javascript:` href is a navigation, not an inline handler, so no
 * `script-src`/`unsafe-inline` rule stops it. The rendered HTML is therefore
 * cleaned before it reaches the ProseMirror node view, which is also where it
 * is stored in the document model and re-serialised to .docx.
 */
import { describe, expect, it } from 'vitest'

import { markdownPasteHtml } from '../src/renderer/editor/markdown-paste'

/// Markdown shapes that trigger conversion (see STRONG_SIGNALS).
const md = (...lines: string[]): string => markdownPasteHtml(lines.join('\n')) ?? ''

describe('markdownPasteHtml strips script-carrying markup', () => {
  it('drops a javascript: href, which CSP does not cover', () => {
    const html = md('## T', '', '[click](javascript:alert(1))')
    expect(html).not.toContain('javascript:')
    expect(html).not.toContain('alert(1)')
    // the link text survives as plain text, only the target is gone
    expect(html).toContain('click')
  })

  it('drops a scheme-obfuscated href whatever its case', () => {
    const html = md('## T', '', '[click](JaVaScRiPt:alert(1))')
    expect(html.toLowerCase()).not.toContain('javascript:')
  })

  it('drops a data: URL that could smuggle a document', () => {
    const html = md('## T', '', '[x](data:text/html;base64,PHNjcmlwdD5hbGVydCgxKTwvc2NyaXB0Pg==)')
    expect(html).not.toContain('data:text/html')
  })

  it('keeps an inline base64 picture the image node would persist, but not a data: link', () => {
    const png = 'data:image/png;base64,iVBORw0KGgo='
    const html = md(
      `![p](${png})`,
      '',
      `[x](${png})`,
      '',
      '![s](data:image/svg+xml;base64,PHN2Zy8+)',
    )
    expect(html).toContain(`<img src="${png}"`)
    expect(html).not.toContain(`href="${png}"`)
    expect(html).not.toContain('svg+xml')
  })

  it('removes raw script and iframe elements the converter passes through', () => {
    const html = md('<script>alert(1)</script>', '', '- one', '- two')
    expect(html).not.toContain('<script')
    expect(html).not.toContain('alert(1)')
    const framed = md('<iframe src="https://evil.example"></iframe>', '', '- one', '- two')
    expect(framed).not.toContain('<iframe')
  })

  it('strips inline event handler attributes', () => {
    const html = md('<img src="x" onerror="alert(1)">', '', '- one', '- two')
    expect(html.toLowerCase()).not.toContain('onerror')
  })
})

describe('markdownPasteHtml keeps ordinary formatting', () => {
  it('preserves the document structure it converted', () => {
    const html = md('## Title', '', 'Some **bold** text.', '', '- one', '- two')
    expect(html).toContain('<h2>')
    expect(html).toContain('<strong>bold</strong>')
    expect(html).toContain('<li>one</li>')
  })

  it('keeps http, https and mailto targets', () => {
    const html = md(
      '## T',
      '',
      '[a](https://example.com) [b](http://example.com) [c](mailto:a@b.c)',
    )
    expect(html).toContain('href="https://example.com"')
    expect(html).toContain('href="http://example.com"')
    expect(html).toContain('href="mailto:a@b.c"')
  })

  it('keeps an in-document anchor', () => {
    expect(md('## T', '', '[jump](#section-2)')).toContain('href="#section-2"')
  })
})
