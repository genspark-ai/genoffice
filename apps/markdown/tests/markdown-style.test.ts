import { afterAll, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { buildExtensions } from '../src/renderer/editor/extensions'
import {
  type MarkedToken,
  resolveMarkdownStyle,
  styleHintsFromTokens,
} from '../src/renderer/markdown/markdownStyle'
import { buildSourceMap, spliceMarkdown } from '../src/renderer/markdown/sourceSplice'

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

function lex(editor: Editor, source: string): MarkedToken[] {
  const marked = editor.markdown!.instance as unknown as {
    Lexer: new (options: unknown) => { lex(src: string): MarkedToken[] }
    defaults: unknown
  }
  return new marked.Lexer(marked.defaults).lex(source)
}

/** load `source`, replace top-level block `index` with `html`, and save through the splice */
function replaceBlock(editor: Editor, source: string, index: number, html: string): string {
  editor.commands.setContent(source, { contentType: 'markdown' })
  const map = buildSourceMap(editor, editor.state.doc, source)
  expect(map).not.toBeNull()
  let from = 0
  for (let i = 0; i < index; i++) from += editor.state.doc.child(i).nodeSize
  const to = from + editor.state.doc.child(index).nodeSize
  editor.commands.insertContentAt({ from, to }, html)
  return spliceMarkdown(editor, editor.state.doc, map!)
}

/** load `source`, type `text` at the end of the first paragraph of the first block, and save */
function appendToFirstParagraph(editor: Editor, source: string, text: string): string {
  editor.commands.setContent(source, { contentType: 'markdown' })
  const map = buildSourceMap(editor, editor.state.doc, source)
  expect(map).not.toBeNull()
  let pos = -1
  editor.state.doc.descendants((node, nodePos) => {
    if (pos >= 0) return false
    if (node.type.name === 'paragraph') pos = nodePos + node.nodeSize - 1
    return pos < 0
  })
  editor.commands.insertContentAt(pos, text)
  return spliceMarkdown(editor, editor.state.doc, map!)
}

describe('style detection', () => {
  it('reads the conventions a document uses', () => {
    const editor = createEditor()
    const hints = styleHintsFromTokens(
      lex(
        editor,
        [
          'Title',
          '=====',
          '',
          'Some _em_ and __strong__ and _more_ text\\',
          'broken.',
          '',
          '* one',
          '* two',
          '  * nested',
          '',
          '1) x',
          '1) y',
          '',
          '~~~sh',
          'echo',
          '~~~',
          '',
          '* * *',
          '',
          '| a | b |',
          '|---|---|',
          '| _c_ | d |',
          '',
        ].join('\n'),
      ),
    )
    expect(hints).toEqual({
      setext: true,
      em: '_',
      strong: '__',
      hardBreak: 'backslash',
      bullet: '*',
      orderedDelimiter: ')',
      orderedRepeat: true,
      fence: '~',
      rule: '* * *',
      tableAligned: false,
    })
  })

  it('takes the majority when a document mixes conventions', () => {
    const editor = createEditor()
    const hints = styleHintsFromTokens(lex(editor, '- a\n\n* b\n\n- c\n\n*x* _y_ *z*\n'))
    expect(hints.bullet).toBe('-')
    expect(hints.em).toBe('*')
  })

  it('lets only level 1 and 2 headings vote on setext', () => {
    const editor = createEditor()
    const hints = styleHintsFromTokens(
      lex(editor, 'Title\n=====\n\n### a\n\n### b\n\n### c\n\n#### d\n'),
    )
    expect(hints.setext).toBe(true)
    expect(styleHintsFromTokens(lex(editor, '### only\n')).setext).toBeUndefined()
  })

  it('layers hints over the defaults, first answer wins', () => {
    expect(resolveMarkdownStyle({ bullet: '+' }, { bullet: '*', em: '_' })).toMatchObject({
      bullet: '+',
      em: '_',
      strong: '**',
    })
  })
})

describe('edited blocks keep their conventions', () => {
  it('bullet marker and nested indentation', () => {
    const editor = createEditor()
    const out = replaceBlock(
      editor,
      '# T\n\n* a\n* b\n',
      1,
      '<ul><li><p>a</p></li><li><p>b</p><ul><li><p>c</p></li></ul></li></ul>',
    )
    expect(out).toBe('# T\n\n* a\n* b\n  * c\n')
  })

  it('ordered delimiter and repeated numbering', () => {
    const editor = createEditor()
    expect(
      replaceBlock(
        editor,
        '1) x\n2) y\n',
        0,
        '<ol><li><p>x</p></li><li><p>y</p></li><li><p>z</p></li></ol>',
      ),
    ).toBe('1) x\n2) y\n3) z\n')
    expect(
      replaceBlock(
        editor,
        '1. x\n1. y\n',
        0,
        '<ol><li><p>x</p></li><li><p>y</p></li><li><p>z</p></li></ol>',
      ),
    ).toBe('1. x\n1. y\n1. z\n')
  })

  it('emphasis delimiters, per block over the document', () => {
    const editor = createEditor()
    const out = replaceBlock(
      editor,
      'A *star* line.\n\nB _under_ line.\n',
      1,
      '<p>B <em>under</em> and <strong>bold</strong> line.</p>',
    )
    expect(out).toBe('A *star* line.\n\nB _under_ and **bold** line.\n')
  })

  it('setext headings, fences, rules and hard breaks', () => {
    const editor = createEditor()
    expect(replaceBlock(editor, 'Old\n===\n\nbody\n', 0, '<h1>New</h1>')).toBe('New\n===\n\nbody\n')
    expect(replaceBlock(editor, 'T\n---\n\nbody\n', 0, '<h2>Sub</h2>')).toBe('Sub\n---\n\nbody\n')
    expect(
      replaceBlock(
        editor,
        '~~~js\nold\n~~~\n',
        0,
        '<pre><code class="language-js">new\n```</code></pre>',
      ),
    ).toBe('~~~js\nnew\n```\n~~~\n')
    expect(replaceBlock(editor, 'a\n\n* * *\n\nb\n', 0, '<p>a</p><hr>')).toBe(
      'a\n\n* * *\n\n* * *\n\nb\n',
    )
    expect(replaceBlock(editor, 'one\\\ntwo\n', 0, '<p>one<br>three</p>')).toBe('one\\\nthree\n')
  })

  it('falls back to ATX where an underline would not make a heading', () => {
    const editor = createEditor()
    for (const [html, atx] of [
      ['<h1>- item</h1>', '# - item'],
      ['<h1>1. step</h1>', '# 1. step'],
      ['<h1>---</h1>', '# ---'],
      ['<h2>a<br>===</h2>', '## a  \n==='],
    ] as const) {
      expect(replaceBlock(editor, 'Old\n===\n\nbody\n', 0, html)).toBe(`${atx}\n\nbody\n`)
    }
  })

  it('keeps a heading with a line break setext, which ATX could not hold', () => {
    const editor = createEditor()
    const out = replaceBlock(editor, 'Old\n===\n\nbody\n', 0, '<h1>a<br>longer</h1>')
    expect(out).toBe('a  \nlonger\n======\n\nbody\n')
    const reread = editor.markdown!.parse(out).content!
    expect(reread[0]).toMatchObject({ type: 'heading', attrs: { level: 1 } })
    expect(reread[0]!.content!.map((n) => n.type)).toEqual(['text', 'hardBreak', 'text'])
  })

  it('adjacent edited blocks each keep their own conventions', () => {
    const editor = createEditor()
    const source = 'A\n===\n\n# B\n\nx _u_ y\n\nx *s* y\n'
    editor.commands.setContent(source, { contentType: 'markdown' })
    const map = buildSourceMap(editor, editor.state.doc, source)!
    editor.commands.setContent('<h1>A2</h1><h1>B2</h1><p>x <em>u2</em> y</p><p>x <em>s2</em> y</p>')
    expect(spliceMarkdown(editor, editor.state.doc, map)).toBe(
      'A2\n===\n\n# B2\n\nx _u2_ y\n\nx *s2* y\n',
    )
  })

  it('does not pair conventions by count alone when the block shapes differ', () => {
    const editor = createEditor()
    // a `_` paragraph and a `*` list: splitting the paragraph and dropping the
    // list keeps the node count, but the second paragraph is not the list
    const source = 'x _u_ y\n\n* a\n'
    editor.commands.setContent(source, { contentType: 'markdown' })
    const map = buildSourceMap(editor, editor.state.doc, source)!
    editor.commands.setContent('<p>x <em>u</em></p><p><em>y</em></p>')
    expect(spliceMarkdown(editor, editor.state.doc, map)).toBe('x _u_\n\n_y_\n')
  })

  it('new blocks follow the document', () => {
    const editor = createEditor()
    editor.commands.setContent('* a\n\n_x_\n', { contentType: 'markdown' })
    const map = buildSourceMap(editor, editor.state.doc, '* a\n\n_x_\n')!
    editor.commands.insertContentAt(
      editor.state.doc.content.size,
      '<ul><li><p>b <em>y</em></p></li></ul>',
    )
    expect(spliceMarkdown(editor, editor.state.doc, map)).toBe('* a\n\n_x_\n\n* b _y_\n')
  })

  it('a loose list keeps its blank lines and content-column indent after an edit', () => {
    const editor = createEditor()
    const source = '* Item one.\n\n  ```text\n  code\n  ```\n\n  ![alt](/x.png)\n\n* Item two.\n'
    expect(appendToFirstParagraph(editor, source, ' EDITED')).toBe(
      '* Item one. EDITED\n\n  ```text\n  code\n  ```\n\n  ![alt](/x.png)\n\n* Item two.\n',
    )
  })

  it('a tight list stays tight; blocks that need a blank line before them get one', () => {
    const editor = createEditor()
    expect(appendToFirstParagraph(editor, '- a\n  ```\n  x\n  ```\n- b\n', '!')).toBe(
      '- a!\n  ```\n  x\n  ```\n- b\n',
    )
    // a blank line inside an item makes the whole list loose (CommonMark), so
    // the items are separated too; the image is never glued to the paragraph
    expect(appendToFirstParagraph(editor, '- a\n\n  ![i](/i.png)\n- b\n', '!')).toBe(
      '- a!\n\n  ![i](/i.png)\n\n- b\n',
    )
  })

  it('ordered items indent their children by the marker width', () => {
    const editor = createEditor()
    expect(appendToFirstParagraph(editor, '1. a\n   - b\n2. c\n', '!')).toBe(
      '1. a!\n   - b\n2. c\n',
    )
    expect(appendToFirstParagraph(editor, '9. a\n10. b\n    - c\n', '!')).toBe(
      '9. a!\n10. b\n    - c\n',
    )
  })

  it('lazy continuation lines are indented to the content column', () => {
    const editor = createEditor()
    const source =
      '- Use a **x** for deps. Deno reads it,\n  so most projects run.\n- Add a **y**.\n'
    expect(appendToFirstParagraph(editor, source, ' EDITED')).toBe(
      '- Use a **x** for deps. Deno reads it,\n  so most projects run. EDITED\n- Add a **y**.\n',
    )
  })

  it('task items follow the same rules', () => {
    const editor = createEditor()
    expect(appendToFirstParagraph(editor, '* [ ] a\n  * [x] b\n* [ ] c\n', '!')).toBe(
      '* [ ] a!\n  * [x] b\n* [ ] c\n',
    )
  })

  it('the saved list reads back as the document the user saw', () => {
    const editor = createEditor()
    const source = '- a\n\n  b\n- c\n\n1. x\n   - y\n   z\n2. w\n'
    const saved = appendToFirstParagraph(editor, source, '!')
    expect(saved).toBe('- a!\n\n  b\n\n- c\n\n1. x\n   - y\n   z\n2. w\n')
    const reread = editor
      .markdown!.parse(saved)
      .content!.map((json) => editor.schema.nodeFromJSON(json).toJSON())
    const shown = editor.getJSON().content!
    expect(reread).toEqual(shown.slice(0, reread.length))
    expect(shown.slice(reread.length)).toEqual([{ type: 'paragraph' }])
  })

  it('a compact table stays compact, an aligned one stays aligned', () => {
    const editor = createEditor()
    const cell =
      '<table><tr><th>a</th><th>b</th></tr><tr><td>longer cell X</td><td>2</td></tr></table>'
    expect(replaceBlock(editor, 'p\n\n| a | b |\n|---|---|\n| longer cell | 2 |\n', 1, cell)).toBe(
      'p\n\n| a | b |\n| --- | --- |\n| longer cell X | 2 |\n',
    )
    expect(
      replaceBlock(
        editor,
        '| a           | b |\n|-------------|---|\n| longer cell | 2 |\n',
        0,
        cell,
      ),
    ).toBe('| a             | b   |\n| ------------- | --- |\n| longer cell X | 2   |\n')
  })

  it('leaves the plain serializer at its defaults', () => {
    const editor = createEditor()
    editor.commands.setContent('* a\n\n_x_\n\nT\n===\n', { contentType: 'markdown' })
    expect(editor.getMarkdown().trimEnd()).toBe('- a\n\n*x*\n\n# T')
  })
})
