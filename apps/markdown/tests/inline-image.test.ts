import { afterAll, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { buildExtensions } from '../src/renderer/editor/extensions'
import { buildSourceMap, spliceMarkdown } from '../src/renderer/markdown/sourceSplice'

const editors: Editor[] = []
afterAll(() => {
  for (const e of editors) e.destroy()
})

interface Opened {
  editor: Editor
  /** save through the block-level splice against the map built when the file was opened */
  save: () => string
}

function open(source: string): Opened {
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
  editor.commands.setContent(source, { contentType: 'markdown' })
  const map = buildSourceMap(editor, editor.state.doc, source)
  expect(map).not.toBeNull()
  return { editor, save: () => spliceMarkdown(editor, editor.state.doc, map!) }
}

/** position just inside the end of top-level block `index` */
function endOfBlock(editor: Editor, index: number): number {
  let pos = 0
  for (let i = 0; i < index; i++) pos += editor.state.doc.child(i).nodeSize
  return pos + editor.state.doc.child(index).nodeSize - 1
}

/** markless insert: the link mark is inclusive, so insertText after a badge would extend the link */
function typeAtEndOfBlock(editor: Editor, index: number, text: string): void {
  editor.view.dispatch(editor.state.tr.insert(endOfBlock(editor, index), editor.schema.text(text)))
}

function imagesIn(editor: Editor): string[] {
  const out: string[] = []
  editor.state.doc.descendants((node) => {
    if (node.type.name === 'image') out.push(String(node.attrs.src))
  })
  return out
}

describe('images mixed with text', () => {
  it('a paragraph with text around an image is a valid document that stays editable', () => {
    const source = 'Text ![logo](x.png) more text.\n'
    const { editor, save } = open(source)
    expect(() => editor.state.doc.check()).not.toThrow()
    expect(editor.getHTML()).toBe('<p>Text <img src="x.png" alt="logo"> more text.</p>')
    expect(save()).toBe(source)

    expect(() => editor.state.tr.insertText('!', endOfBlock(editor, 0))).not.toThrow()
    typeAtEndOfBlock(editor, 0, ' EDITED')
    expect(imagesIn(editor)).toEqual(['x.png'])
    expect(save()).toBe('Text ![logo](x.png) more text. EDITED\n')
  })

  it('a heading keeps its link-wrapped badge through typing and save', () => {
    const source =
      '# Awesome ML [![Awesome](https://x/badge.svg)](https://github.com/sindresorhus/awesome)\n\nBody.\n'
    const { editor, save } = open(source)
    expect(() => editor.state.doc.check()).not.toThrow()
    const heading = editor.state.doc.child(0)
    expect(heading.type.name).toBe('heading')
    const badge = heading.child(1)
    expect(badge.type.name).toBe('image')
    expect(badge.marks.map((m) => [m.type.name, m.attrs.href])).toEqual([
      ['link', 'https://github.com/sindresorhus/awesome'],
    ])
    expect(editor.getHTML()).toContain(
      '<a target="_blank" rel="noopener noreferrer nofollow" href="https://github.com/sindresorhus/awesome"><img src="https://x/badge.svg" alt="Awesome"></a>',
    )
    expect(save()).toBe(source)

    typeAtEndOfBlock(editor, 0, ' EDITED')
    expect(imagesIn(editor)).toEqual(['https://x/badge.svg'])
    expect(save()).toBe(
      '# Awesome ML [![Awesome](https://x/badge.svg)](https://github.com/sindresorhus/awesome) EDITED\n\nBody.\n',
    )
    expect(editor.getMarkdown()).toBe(
      '# Awesome ML [![Awesome](https://x/badge.svg)](https://github.com/sindresorhus/awesome) EDITED\n\nBody.',
    )
  })

  it('a titled link around a titled image serializes both titles', () => {
    const { editor } = open('[![a](x.png "pic")](https://y "site")\n')
    expect(editor.getMarkdown()).toBe('[![a](x.png "pic")](https://y "site")')
  })

  it('list items, table cells and blockquotes hold images next to text', () => {
    for (const source of [
      '- item ![i](a.png) tail\n',
      '| a | ![i](a.png) |\n| - | - |\n| b | c |\n',
      '> quote ![q](q.png) end\n',
    ]) {
      const { editor, save } = open(source)
      expect(() => editor.state.doc.check()).not.toThrow()
      expect(imagesIn(editor)).toHaveLength(1)
      expect(save()).toBe(source)
    }
  })

  it('typing inside a list item next to an image keeps the image', () => {
    const source = '- item ![i](a.png) tail\n'
    const { editor, save } = open(source)
    const paragraphEnd = 1 + 1 + editor.state.doc.child(0).child(0).child(0).nodeSize - 1
    editor.view.dispatch(editor.state.tr.insertText(' EDITED', paragraphEnd))
    expect(imagesIn(editor)).toEqual(['a.png'])
    expect(save()).toBe('- item ![i](a.png) tail EDITED\n')
  })
})

describe('sized HTML images', () => {
  it('keep their width and height through an edit', () => {
    const source =
      '| [<img src="x/edge.png" alt="Edge" width="24px" height="24px" />](https://y)<br>Edge | b |\n| --- | --- |\n| 1 | 2 |\n'
    const { editor, save } = open(source)
    let at = -1
    editor.state.doc.descendants((node, pos) => {
      if (at < 0 && node.type.name === 'paragraph' && node.textContent === 'b')
        at = pos + node.nodeSize - 1
    })
    editor.view.dispatch(editor.state.tr.insertText(' EDITED', at))
    const saved = save()
    expect(saved).toContain(
      '[<img src="x/edge.png" alt="Edge" width="24px" height="24px" />](https://y)',
    )
    expect(saved).toContain('b EDITED')
    expect(editor.getMarkdown()).toContain(
      '<img src="x/edge.png" alt="Edge" width="24px" height="24px" />',
    )
  })

  it('a sized picture followed by a hard break does not open an HTML block', () => {
    const { editor } = open('<img src="a.png" width="10"><br>Tail text.\n')
    const out = editor.getMarkdown()
    expect(out).toMatch(/^<img src="a\.png" width="10" \/>(?: \\\n|<br>\n?)Tail text\./)
    editor.commands.setContent(out, { contentType: 'markdown' })
    expect(editor.getHTML()).toContain('<br>')
    expect(editor.getHTML()).toContain('Tail text.')
  })

  it('a sized picture with an angle bracket in its alt text is still guarded', () => {
    const { editor } = open('<img src="a.png" alt="a > b" width="10"><br>Tail text.\n')
    const out = editor.getMarkdown()
    expect(out).toMatch(
      /^<img src="a\.png" alt="a &gt; b" width="10" \/>(?: \\\n|<br>\n?)Tail text\./,
    )
  })

  it('a plain picture still serializes as markdown', () => {
    const { editor } = open('Text ![logo](x.png) more.\n')
    expect(editor.getMarkdown()).toContain('![logo](x.png)')
  })
})

describe('standalone image lines', () => {
  it('parse to a paragraph holding only the image and round-trip byte for byte', () => {
    for (const source of [
      'Para.\n\n![a](x.png)\n\nPara2.\n',
      'Para.\n\n![a](x.png "title")\n\nPara2.\n',
    ]) {
      const { editor, save } = open(source)
      expect(() => editor.state.doc.check()).not.toThrow()
      const block = editor.state.doc.child(1)
      expect(block.type.name).toBe('paragraph')
      expect(block.childCount).toBe(1)
      expect(block.firstChild!.type.name).toBe('image')
      expect(save()).toBe(source)
      expect(`${editor.getMarkdown()}\n`).toBe(source)
    }
  })

  it('re-serializes on its own line after a neighbouring edit', () => {
    const source = 'Para.\n\n![a](x.png)\n\nPara2.\n'
    const { editor, save } = open(source)
    typeAtEndOfBlock(editor, 2, ' EDITED')
    expect(save()).toBe('Para.\n\n![a](x.png)\n\nPara2. EDITED\n')
  })
})
