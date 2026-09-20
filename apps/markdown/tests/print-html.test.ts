import { describe, expect, it, vi } from 'vitest'

import { buildPrintHtml } from '../src/renderer/export/printHtml'

// The KaTeX stylesheet is inlined with Vite's `?inline` query; stub it so the
// test asserts the builder's own behavior rather than the bundler's.
vi.mock('katex/dist/katex.min.css?inline', () => ({ default: '' }))

function editorRoot(innerHtml: string): HTMLElement {
  const root = document.createElement('div')
  root.setAttribute('contenteditable', 'true')
  root.innerHTML = innerHtml
  return root
}

describe('buildPrintHtml', () => {
  it('builds a self-contained document with base, title, and styles', () => {
    const html = buildPrintHtml(editorRoot('<h1>Hello</h1><p>World</p>'), 'Notes')
    expect(html).toContain('<!doctype html>')
    expect(html).toContain('<base href="')
    expect(html).toContain('<title>Notes</title>')
    // KaTeX stylesheet plus the print-theme stylesheet.
    expect(html.match(/<style>/g)?.length).toBeGreaterThanOrEqual(2)
    expect(html).toContain('<h1>Hello</h1>')
  })

  it('escapes the document title', () => {
    const html = buildPrintHtml(editorRoot('<p>x</p>'), 'A&B <C>')
    expect(html).toContain('<title>A&amp;B &lt;C></title>')
    expect(html).not.toContain('<title>A&B <C></title>')
  })

  it('strips editor-only chrome from the clone', () => {
    const html = buildPrintHtml(
      editorRoot(
        '<p contenteditable="true">text<img class="ProseMirror-separator" alt=""></p>' +
          '<div class="md-codeblock-bar"><button>copy</button></div>',
      ),
      'Notes',
    )
    expect(html).not.toContain('contenteditable')
    expect(html).not.toContain('md-codeblock-bar')
    expect(html).not.toContain('ProseMirror-separator')
    expect(html).toContain('<p>text</p>')
  })
})
