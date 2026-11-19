import { afterAll, afterEach, beforeAll, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { buildExtensions } from '../src/renderer/editor/extensions'

// jsdom has no ClipboardEvent; the view still runs its real clipboard parser.
beforeAll(() => vi.stubGlobal('ClipboardEvent', Event))
afterAll(() => vi.unstubAllGlobals())

let editor: Editor
function createEditor(content = '') {
  editor = new Editor({
    extensions: buildExtensions({
      slashController: {
        onOpen() {},
        onUpdate() {},
        onKeyDown: () => false,
        onClose() {},
      },
      slashItems: () => [],
    }),
    content,
  })
  return editor
}
afterEach(() => editor?.destroy())

function pasteText(text: string) {
  const event = new Event('paste', { bubbles: true, cancelable: true })
  Object.defineProperty(event, 'clipboardData', {
    value: {
      getData: (type: string) => (type === 'text/plain' ? text : ''),
      files: [],
      items: [],
      types: ['text/plain'],
    },
  })
  editor.view.dom.dispatchEvent(event)
}

function formulas() {
  const found: { type: string; latex: string }[] = []
  editor.state.doc.descendants((node) => {
    if (node.type.name === 'inlineMath' || node.type.name === 'blockMath')
      found.push({ type: node.type.name, latex: node.attrs.latex })
  })
  return found
}

describe('formula clipboard integration', () => {
  it.each([
    ['$E=mc^2$', 'inlineMath', 'E=mc^2'],
    [String.raw`\(E=mc^2\)`, 'inlineMath', 'E=mc^2'],
    ['$$\nx^2 + y^2\n$$', 'blockMath', 'x^2 + y^2'],
    [String.raw`\[x^2 + y^2\]`, 'blockMath', 'x^2 + y^2'],
  ])('pastes %s as a formula and preserves it on reopen', (text, type, latex) => {
    createEditor()
    pasteText(text)
    expect(formulas()).toEqual([{ type, latex }])
    const saved = editor.getMarkdown()
    editor.commands.setContent(saved, { contentType: 'markdown' })
    expect(formulas()).toEqual([{ type, latex }])
    expect(() => editor.state.doc.check()).not.toThrow()
  })

  it('preserves surrounding HTML marks while converting a formula', () => {
    createEditor().view.pasteHTML(
      String.raw`<p><strong>Energy</strong> \(E=mc^2\) <a href="https://example.com">source</a></p>`,
    )
    expect(formulas()).toEqual([{ type: 'inlineMath', latex: 'E=mc^2' }])
    expect(editor.getHTML()).toContain('<strong>Energy</strong>')
    expect(editor.getHTML()).toContain('href="https://example.com"')
    expect(editor.state.doc.textContent).toBe('Energy  source')
  })

  it('recovers rendered KaTeX from its TeX annotation without duplicate visual text', () => {
    createEditor().view.pasteHTML(
      '<p>Energy <span class="katex"><span class="katex-mathml"><math><semantics><mrow><mi>E</mi></mrow><annotation encoding="application/x-tex">E=mc^2</annotation></semantics></math></span><span class="katex-html">duplicate visual text</span></span>.</p>',
    )
    expect(formulas()).toEqual([{ type: 'inlineMath', latex: 'E=mc^2' }])
    expect(editor.state.doc.textContent).toBe('Energy .')
  })

  it.each([
    'paid $5 and $10',
    String.raw`see \[1\] and \(note\)`,
    '`$x^2$`',
    '```latex\n\\(x^2\\)\n```',
  ])('does not turn literal source into formulas: %s', (text) => {
    createEditor()
    pasteText(text)
    expect(formulas()).toEqual([])
    expect(editor.state.doc.textBetween(0, editor.state.doc.content.size, '\n')).toContain(
      text.startsWith('`$') ? '$x^2$' : text,
    )
  })

  it('preserves source pasted into a code block', () => {
    createEditor('<pre><code></code></pre>')
    pasteText(String.raw`\(x^2\) $y_1$`)
    expect(formulas()).toEqual([])
    expect(editor.state.doc.textContent).toBe(String.raw`\(x^2\) $y_1$`)
  })

  it('preserves formulas inside HTML code', () => {
    createEditor().view.pasteHTML('<p><code>$x^2$</code></p><pre><code>\\(y^2\\)</code></pre>')
    expect(formulas()).toEqual([])
    expect(editor.state.doc.textContent).toBe(String.raw`$x^2$\(y^2\)`)
  })
  it('keeps code source unchanged beside a rendered formula', () => {
    createEditor()
    pasteText('\\(E=mc^2\\)\n\n`\\(x^2\\)`\n\n```latex\n\\[y^2\\]\n```')
    expect(formulas()).toEqual([{ type: 'inlineMath', latex: 'E=mc^2' }])
    const code: string[] = []
    editor.state.doc.descendants((node) => {
      if (node.type.spec.code || node.marks.some((mark) => mark.type.spec.code))
        code.push(node.textContent)
    })
    expect(code).toContain(String.raw`\(x^2\)`)
    expect(editor.state.doc.textBetween(0, editor.state.doc.content.size, '\n')).toContain(
      '```latex\n\\[y^2\\]\n```',
    )
  })

  it('preserves display math from rich HTML', () => {
    createEditor().view.pasteHTML(
      '<div class="katex-display"><span class="katex"><math><semantics><annotation encoding="application/x-tex">x^2 + y^2</annotation></semantics></math><span>duplicate</span></span></div>',
    )
    expect(formulas()).toEqual([{ type: 'blockMath', latex: 'x^2 + y^2' }])
    expect(editor.state.doc.textContent).toBe('')
  })

  it('keeps text and formulas on both sides of a multiline display block', () => {
    createEditor()
    pasteText('Before $a_1$\n\n\\[\nx^2 + y^2\n\\]\n\nAfter $b_2$')
    expect(formulas()).toEqual([
      { type: 'inlineMath', latex: 'a_1' },
      { type: 'blockMath', latex: 'x^2 + y^2' },
      { type: 'inlineMath', latex: 'b_2' },
    ])
    expect(editor.state.doc.textContent).toBe('Before After ')
    editor.commands.undo()
    expect(formulas()).toEqual([])
    expect(editor.state.doc.textContent).toBe('')
  })

  it.each(['inline', 'display'])('removes Wikipedia %s fallback images', (display) => {
    createEditor().view.pasteHTML(
      `<span class="mwe-math-element"><math display="${display === 'display' ? 'block' : 'inline'}"><semantics><annotation encoding="application/x-tex">E=mc^2</annotation></semantics></math><img class="mwe-math-fallback-image-${display}" src="https://example.com/math.svg"></span>`,
    )
    expect(formulas()).toEqual([
      { type: display === 'display' ? 'blockMath' : 'inlineMath', latex: 'E=mc^2' },
    ])
    expect(editor.getHTML()).not.toContain('<img')
  })

  it('keeps surrounding plain text on the default parser path', () => {
    createEditor()
    pasteText('# Title $x^2$\n> quote\n[link](https://example.com)\n---')
    expect(formulas()).toEqual([{ type: 'inlineMath', latex: 'x^2' }])
    expect(editor.state.doc.firstChild?.type.name).toBe('paragraph')
    expect(editor.state.doc.textBetween(0, editor.state.doc.content.size, '\n')).toBe(
      '# Title \n> quote\n[link](https://example.com)\n---',
    )
  })

  it('does not insert empty paragraphs around display math', () => {
    createEditor()
    pasteText('Before\n$$x^2$$\nAfter')
    expect(editor.getJSON().content?.map((node) => [node.type, node.content?.[0]?.text])).toEqual([
      ['paragraph', 'Before'],
      ['blockMath', undefined],
      ['paragraph', 'After'],
    ])
  })

  it('preserves active marks on ordinary plain-text paste', () => {
    createEditor('<p><strong>before after</strong></p>')
    editor.commands.setTextSelection(8)
    pasteText('ordinary text ')
    expect(editor.getHTML()).toContain('<strong>before ordinary text after</strong>')
  })

  it('does not convert escaped dollar delimiters in rich text', () => {
    createEditor().view.pasteHTML(String.raw`<p>\$x^2\$ and $y_1$</p>`)
    expect(formulas()).toEqual([{ type: 'inlineMath', latex: 'y_1' }])
    expect(editor.state.doc.textContent).toBe(String.raw`\$x^2\$ and `)
  })
  it.each([String.raw`<p>\[<br>x^2 + y^2<br>\]</p>`, '<p>$$<br><span>x^2 + y^2</span><br>$$</p>'])(
    'pastes multiline rich display math: %s',
    (html) => {
      createEditor().view.pasteHTML(html)
      expect(formulas()).toEqual([{ type: 'blockMath', latex: 'x^2 + y^2' }])
      expect(editor.state.doc.textContent).toBe('')
    },
  )
})
