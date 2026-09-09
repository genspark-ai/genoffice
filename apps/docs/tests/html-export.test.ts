/**
 * Standalone HTML export: editing chrome is dropped, review wrappers are
 * flattened to their accepted text, authored attributes survive, and the
 * output carries no app class names (jsdom resolves inline styles only, so
 * the cascade itself is exercised on the real app).
 */
import { describe, expect, it } from 'vitest'
import { buildStandaloneHtml, resolveContent } from '../src/renderer/html-export'

function mount(html: string): HTMLElement {
  const root = document.createElement('div')
  root.className = 'doc-page ProseMirror'
  root.innerHTML = html
  document.body.appendChild(root)
  return root
}

describe('buildStandaloneHtml', () => {
  it('drops widgets and tracked deletions, unwraps comments and insertions', () => {
    const root = mount(
      '<p class="doc-p">Hello <span class="doc-comment">noted</span> ' +
        '<span class="doc-ins">added</span><span class="doc-del">gone</span>' +
        '<span class="ProseMirror-widget">⚑</span></p>' +
        '<div class="page-gap">page 2</div>',
    )
    const out = buildStandaloneHtml(root, { title: 'T' })
    expect(out).toContain('<p>Hello noted added</p>')
    expect(out).not.toContain('gone')
    expect(out).not.toContain('⚑')
    expect(out).not.toContain('page 2')
    expect(out).not.toMatch(/class="doc-/)
  })

  it('keeps links, images, table spans and inline widths', () => {
    const root = mount(
      '<p><a href="https://x.test" class="doc-link">x</a>' +
        '<img src="data:image/png;base64,AAAA" alt="pic" style="width:120px;height:40px"></p>' +
        '<table class="doc-table"><tbody><tr><td colspan="2" style="width:200px">c</td></tr></tbody></table>',
    )
    const out = buildStandaloneHtml(root, { title: 'T', textWidthPx: 600 })
    expect(out).toMatch(/<a href="https:\/\/x.test"[^>]*>x<\/a>/)
    expect(out).toMatch(/<img src="data:image\/png;base64,AAAA" alt="pic" width="120" height="40"/)
    expect(out).toMatch(/<td colspan="2" style="[^"]*width:200px/)
    expect(out).toContain('max-width:600px')
    expect(out).toMatch(/^<!DOCTYPE html>/)
  })

  it('escapes text and restores lifted screen-only classes', () => {
    const wrap = document.createElement('div')
    wrap.className = 'workspace page-dark'
    document.body.appendChild(wrap)
    const root = document.createElement('div')
    root.innerHTML = '<p>a &lt; b &amp; "c"</p>'
    wrap.appendChild(root)
    const out = buildStandaloneHtml(root, { title: 'A & B' })
    expect(out).toContain('<p>a &lt; b &amp; "c"</p>')
    expect(out).toContain('<title>A &amp; B</title>')
    expect(wrap.classList.contains('page-dark')).toBe(true)
  })
})

describe('resolveContent', () => {
  it('resolves strings, escapes, attr() and the ordered counter', () => {
    const el = document.createElement('p')
    el.setAttribute('data-marker', '(a)')
    expect(resolveContent('attr(data-marker) "\\00a0"', el)).toBe('(a) ')
    expect(resolveContent('"•"', el)).toBe('•')
    expect(resolveContent('none', el)).toBe('')
    const list = document.createElement('div')
    for (let i = 0; i < 3; i++) {
      const li = document.createElement('p')
      li.className = 'doc-li doc-li-ordered'
      list.appendChild(li)
    }
    expect(resolveContent('counter(doc-ol) "."', list.children[2])).toBe('3.')
  })
})
