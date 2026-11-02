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

  it('prints an aliased wiki link as the alias the canvas shows', () => {
    const html = buildPrintHtml(
      editorRoot(
        '<p><a class="wiki-link"><span class="wiki-alias-hidden">[[Note|</span><span class="wiki-link-resolved">Alias</span><span class="wiki-alias-hidden">]]</span></a>' +
          // the link under the caret shows its syntax and carries no hidden spans
          '<a class="wiki-link">[[Other|Two]]</a>' +
          // no alias: the literal is what the canvas shows
          '<a class="wiki-link">[[Plain]]</a>' +
          // a literal an inline mark splits stays visible on the canvas too
          '<a class="wiki-link">[[Note|bo<strong>ld</strong>]]</a></p>',
      ),
      'x',
    )
    expect(html).toContain('<a class="wiki-link">Alias</a>')
    expect(html).toContain('<a class="wiki-link">Two</a>')
    expect(html).toContain('<a class="wiki-link">[[Plain]]</a>')
    expect(html).toContain('<a class="wiki-link">[[Note|bo<strong>ld</strong>]]</a>')
    expect(html).not.toContain('[[Note|</span>')
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
