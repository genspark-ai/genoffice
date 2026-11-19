import { afterAll, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { buildExtensions } from '../src/renderer/editor/extensions'
import { escapeBrackets } from '../src/renderer/editor/markdownEscape'

// Undestroyed views leave DOMObserver flush timers that fire after jsdom teardown
// ("document is not defined" unhandled error). Editors here are shared per describe,
// so destroy them once at the end of the file.
const editors: Editor[] = []
afterAll(() => {
  for (const e of editors) e.destroy()
})

function createEditor(): Editor {
  const editor = new Editor({
    extensions: buildExtensions({
      slashController: {
        onOpen: () => {},
        onUpdate: () => {},
        onKeyDown: () => false,
        onClose: () => {},
      },
      slashItems: () => [],
    }),
    content: '',
  })
  editors.push(editor)
  return editor
}

/** parse → serialize → parse must be structurally stable */
function roundTrip(editor: Editor, md: string): { out: string; stable: boolean } {
  const manager = editor.markdown!
  const first = manager.parse(md)
  const out = manager.serialize(first)
  const second = manager.parse(out)
  return { out, stable: JSON.stringify(first) === JSON.stringify(second) }
}

describe('markdown round-trip for GFM nodes', () => {
  const editor = createEditor()

  it('task list', () => {
    const md = '- [ ] open item\n- [x] done item'
    const { out, stable } = roundTrip(editor, md)
    expect(out).toContain('- [ ] open item')
    expect(out).toContain('- [x] done item')
    expect(stable).toBe(true)
  })

  it('table', () => {
    const md = '| Name | Value |\n| --- | --- |\n| a | 1 |\n| b | 2 |'
    const { out, stable } = roundTrip(editor, md)
    // the serializer pads cells to align columns
    expect(out).toContain('| Name | Value |')
    expect(out).toMatch(/\| a\s+\| 1\s+\|/)
    expect(stable).toBe(true)
  })

  it('image keeps the authored relative path in markdown', () => {
    const md = '![diagram](assets/diagram.png)'
    const { out, stable } = roundTrip(editor, md)
    expect(out).toContain('![diagram](assets/diagram.png)')
    expect(stable).toBe(true)
  })

  it('plain constructs round-trip', () => {
    const md = '# Title\n\n> quoted\n\n```js\ncode()\n```\n\n---'
    const { out, stable } = roundTrip(editor, md)
    expect(out).toContain('# Title')
    expect(out).toContain('```js')
    expect(stable).toBe(true)
  })

  it('code block language survives the round-trip', () => {
    const { out, stable } = roundTrip(editor, '```python\nprint(1)\n```')
    expect(out).toContain('```python')
    expect(stable).toBe(true)
  })

  it('mermaid blocks stay plain fenced code in the file', () => {
    const md = '```mermaid\nflowchart LR\n    A --> B\n```'
    const { out, stable } = roundTrip(editor, md)
    expect(out).toContain('```mermaid\nflowchart LR\n    A --> B\n```')
    expect(stable).toBe(true)
    expect(editor.markdown!.parse(md).content?.[0]?.attrs?.language).toBe('mermaid')
  })

  it('nested lists sit at the parent item content column', () => {
    // exactly the marker width: "1. " = 3, "- " = 2, so strict CommonMark
    // parsers (GitHub) read the sub-list as nested for any marker
    const ordered = roundTrip(editor, '1. one\n   - sub\n2. two')
    expect(ordered.out).toContain('\n   - sub')
    expect(ordered.stable).toBe(true)

    const bullets = roundTrip(editor, '- a\n  - b\n    - c')
    expect(bullets.out).toContain('\n  - b')
    expect(bullets.out).toContain('\n    - c')
    expect(bullets.stable).toBe(true)

    // files saved by earlier versions used 4-space indents; still parsed as nested
    const legacy = roundTrip(editor, '- a\n    - b\n        - c')
    expect(legacy.out).toContain('\n  - b\n    - c')
  })
})

describe('math nodes — $ / $$ syntax (issue #100)', () => {
  const editor = createEditor()

  function collect(node: Record<string, unknown>, type: string): Record<string, unknown>[] {
    const hits: Record<string, unknown>[] = []
    if (node.type === type) hits.push(node)
    for (const child of (node.content as Record<string, unknown>[] | undefined) ?? []) {
      hits.push(...collect(child, type))
    }
    return hits
  }

  it('$$ array environment parses to block math with row breaks (\\\\) intact', () => {
    const md = [
      '$$',
      '\\begin{array}{ll}',
      '\\max & f = 1.25 x_{1} + 1.5 x_{2}, \\\\',
      '\\text{s.t.} & 0.025 x_{1} + 0.05 x_{2} \\le 400, \\\\',
      '& x_{1}, x_{2} \\ge 0.',
      '\\end{array}',
      '$$',
    ].join('\n')
    const manager = editor.markdown!
    const doc = manager.parse(md) as Record<string, unknown>
    const nodes = collect(doc, 'blockMath')
    expect(nodes.length).toBe(1)
    const latex = String((nodes[0].attrs as { latex: string }).latex)
    expect(latex).toContain('\\begin{array}{ll}')
    // the markdown escape pass must not eat the \\ row separators
    expect(latex).toContain('\\\\')

    const out = manager.serialize(doc)
    expect(out).toContain('\\begin{array}{ll}')
    expect(out).toContain('\\\\')
    expect(JSON.stringify(manager.parse(out))).toBe(JSON.stringify(doc))
  })

  it('$$ aligned environment round-trips', () => {
    const md = '$$\n\\begin{aligned}\na &= b + c \\\\\nd &= e\n\\end{aligned}\n$$'
    const { out, stable } = roundTrip(editor, md)
    expect(out).toContain('\\begin{aligned}')
    expect(out).toContain('\\\\')
    expect(stable).toBe(true)
  })

  it('inline $x_{1}$ becomes an inline math node and round-trips', () => {
    const manager = editor.markdown!
    const doc = manager.parse('the variable $x_{1}$ is free') as Record<string, unknown>
    const nodes = collect(doc, 'inlineMath')
    expect(nodes.length).toBe(1)
    expect((nodes[0].attrs as { latex: string }).latex).toBe('x_{1}')
    expect(manager.serialize(doc)).toBe('the variable $x_{1}$ is free')
  })

  it('currency amounts in plain text never turn into formulas', () => {
    const manager = editor.markdown!
    const md = 'I paid $5 and $10 in total'
    const doc = manager.parse(md) as Record<string, unknown>
    expect(collect(doc, 'inlineMath').length).toBe(0)
    expect(manager.serialize(doc)).toBe(md)
  })
})

describe('only pure markdown syntax is ever produced', () => {
  const editor = createEditor()

  it('underline and highlight commands are gone', () => {
    // both would serialize as non-GFM syntax (`++u++` / `==mark==`)
    expect('toggleUnderline' in editor.commands).toBe(false)
    expect('toggleHighlight' in editor.commands).toBe(false)
  })

  it('alignment and line-height commands are gone', () => {
    expect('setTextAlign' in editor.commands).toBe(false)
    expect('setLineHeight' in editor.commands).toBe(false)
  })

  it('callout and toggle nodes are gone from the schema', () => {
    expect(editor.schema.nodes.callout).toBeUndefined()
    expect(editor.schema.nodes.toggle).toBeUndefined()
  })
})

describe('legacy HTML content degrades to plain markdown, keeping the text', () => {
  const editor = createEditor()

  function parseAndSerialize(md: string): string {
    const manager = editor.markdown!
    return manager.serialize(manager.parse(md))
  }

  it('a styled span drops the styling but keeps the text', () => {
    const out = parseAndSerialize('a <span style="color: #ff0000">red text</span> b')
    expect(out).not.toContain('<span')
    expect(out).toContain('red text')
  })

  it('an aligned paragraph becomes a plain paragraph with marks intact', () => {
    const out = parseAndSerialize(
      '<p style="text-align: center">centered <strong>text</strong></p>',
    )
    expect(out).not.toContain('<p')
    expect(out).toContain('centered **text**')
  })

  it('an aligned heading becomes a plain heading', () => {
    const out = parseAndSerialize('<h2 style="text-align: right">title</h2>')
    expect(out).toBe('## title')
  })

  it('a sized image keeps its size as an img tag, dropping other attributes', () => {
    const out = parseAndSerialize('<img src="assets/d.png" alt="d" width="300" align="center">')
    expect(out).toBe('<img src="assets/d.png" alt="d" width="300" />')
  })

  it('u and mark tags drop the tag but keep the text', () => {
    const out = parseAndSerialize('a <u>underlined</u> and <mark>marked</mark> b')
    expect(out).toBe('a underlined and marked b')
  })
})

describe('bracket escaping on save', () => {
  const editor = createEditor()

  it.each([
    ['wiki-style link', 'See [[Foo]] and [[Bar|alias]].'],
    ['citation marker', 'As shown in [1] and [2, p. 4].'],
    ['bracketed tag', '[TODO] finish the intro'],
  ])('keeps %s verbatim', (_name, md) => {
    const { out, stable } = roundTrip(editor, md)
    expect(out).toBe(md)
    expect(stable).toBe(true)
  })

  it('still escapes text that would parse as a link', () => {
    const { out, stable } = roundTrip(editor, 'literal \\[a](b) here')
    expect(out).toBe('literal \\[a\\](b) here')
    expect(stable).toBe(true)
  })

  it('still escapes a task marker typed as text', () => {
    const md = '- \\[ ] not a task'
    const { out, stable } = roundTrip(editor, md)
    expect(out).toContain('\\[ \\] not a task')
    expect(stable).toBe(true)
  })

  it('still escapes a reference definition typed as text', () => {
    const { out, stable } = roundTrip(editor, '\\[ref]: not a definition')
    expect(out).toBe('\\[ref\\]: not a definition')
    expect(stable).toBe(true)
  })

  it('escapes an indented reference definition too', () => {
    expect(escapeBrackets('   [ref]: /url')).toBe('   \\[ref\\]: /url')
    expect(escapeBrackets('  [x] not a task')).toBe('  \\[x\\] not a task')
  })

  it('keeps escaping the other inline delimiters', () => {
    const { out, stable } = roundTrip(editor, 'a \\* b \\_ c \\~ d')
    expect(out).toBe('a \\* b \\_ c \\~ d')
    expect(stable).toBe(true)
  })
})

describe('formatted inline code', () => {
  const editor = createEditor()
  it.each([
    ['**`bold code`**', 'bold'],
    ['*`italic code`*', 'italic'],
    ['~~`struck code`~~', 'strike'],
    ['[`linked code`](https://example.com)', 'link'],
  ])('opens, edits and reopens %s with valid marks', (source, outerMark) => {
    editor.commands.setContent(source, { contentType: 'markdown' })
    expect(() => editor.state.doc.check()).not.toThrow()
    const text = editor.state.doc.firstChild!.firstChild!
    expect(text.marks.map((mark) => mark.type.name)).toContain(outerMark)
    expect(text.marks.map((mark) => mark.type.name)).toContain('code')
    editor.commands.insertContentAt(2, 'X')
    const saved = editor.getMarkdown()
    editor.commands.setContent(saved, { contentType: 'markdown' })
    expect(() => editor.state.doc.check()).not.toThrow()
    expect(editor.state.doc.textContent).toContain('X')
    const reopened = editor.state.doc.firstChild!.firstChild!
    expect(reopened.marks.map((mark) => mark.type.name)).toContain(outerMark)
    expect(reopened.marks.map((mark) => mark.type.name)).toContain('code')
  })
})

it('keeps emphasis and link syntax inside code literal', () => {
  const editor = createEditor()
  editor.commands.setContent('`**literal** [text](https://example.com)`', {
    contentType: 'markdown',
  })
  expect(() => editor.state.doc.check()).not.toThrow()
  const text = editor.state.doc.firstChild!.firstChild!
  expect(text.text).toBe('**literal** [text](https://example.com)')
  expect(text.marks.map((mark) => mark.type.name)).toEqual(['code'])
})
