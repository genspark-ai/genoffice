import { afterEach, describe, expect, it } from 'vitest'
import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { Editor } from '@tiptap/core'
import { Markdown } from '@genoffice/ui'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { DOC_NAV_SCHEME, navigateToBlock, parseDocNavHref } from '../src/renderer/ai/doc-nav'
import { AGENT_SYSTEM_PROMPT } from '../src/renderer/ai/protocol'

const editors = new Set<Editor>()
afterEach(() => {
  for (const editor of editors) editor.destroy()
  editors.clear()
})

describe('doc-nav href parsing', () => {
  it('accepts docnav block hrefs and rejects everything else', () => {
    expect(parseDocNavHref('docnav://block/7')).toBe(7)
    expect(parseDocNavHref('docnav://block/0')).toBe(0)
    expect(parseDocNavHref('docnav://block/-1')).toBeNull()
    expect(parseDocNavHref('docnav://para/7')).toBeNull()
    expect(parseDocNavHref('https://example.com')).toBeNull()
  })
})

describe('Markdown nav links', () => {
  const nav = { scheme: DOC_NAV_SCHEME, onNavigate: () => {} }

  it('renders docnav links as anchors and keeps other links literal', () => {
    const html = renderToStaticMarkup(
      createElement(Markdown, {
        text: 'See [Chapter 2](docnav://block/4) and [site](https://example.com).',
        nav,
      }),
    )
    expect(html).toContain('class="ai-md-nav"')
    expect(html).toContain('href="docnav://block/4"')
    expect(html).toContain('Chapter 2')
    // external links keep the pre-existing literal rendering
    expect(html).toContain('[site](https://example.com)')
    expect(html).not.toContain('href="https://example.com"')
  })

  it('without a nav prop every link stays literal (other apps unchanged)', () => {
    const html = renderToStaticMarkup(
      createElement(Markdown, { text: 'See [Chapter 2](docnav://block/4).' }),
    )
    expect(html).toContain('[Chapter 2](docnav://block/4)')
    expect(html).not.toContain('<a')
  })
})

describe('navigateToBlock', () => {
  function createEditor(content: unknown[]): Editor {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: { type: 'doc', content } as never,
    })
    editors.add(editor)
    return editor
  }
  const para = (t: string) => ({
    type: 'docParagraph',
    attrs: { docxIndex: null },
    content: [{ type: 'text', text: t }],
  })
  const image = () => ({
    type: 'docProtected',
    attrs: { docxIndex: null, blockType: 'image', label: 'Image', imageWidthPx: 10 },
  })

  it('selects the cited textblock content', () => {
    const editor = createEditor([para('first'), para('second target'), para('third')])
    navigateToBlock(editor, 1)
    const { from, to } = editor.state.selection
    expect(editor.state.doc.textBetween(from, to)).toBe('second target')
  })

  it('does not throw on atom blocks (images) or out-of-range indexes', () => {
    const editor = createEditor([para('text'), image()])
    expect(() => navigateToBlock(editor, 1)).not.toThrow()
    expect(() => navigateToBlock(editor, 99)).not.toThrow()
    expect(() => navigateToBlock(editor, -3)).not.toThrow()
  })
})

describe('system prompt', () => {
  it('teaches the docnav citation format', () => {
    expect(AGENT_SYSTEM_PROMPT).toContain('docnav://block/N')
  })

  it('carries the template-filling workflow', () => {
    expect(AGENT_SYSTEM_PROMPT).toContain('# Template filling')
    expect(AGENT_SYSTEM_PROMPT).toContain('never invent facts to fill a field')
  })
})
