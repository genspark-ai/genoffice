import { execSync } from 'node:child_process'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { afterAll, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { buildExtensions } from '../src/renderer/editor/extensions'
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

function load(editor: Editor, source: string) {
  editor.commands.setContent(source, { contentType: 'markdown' })
  return buildSourceMap(editor, editor.state.doc, source)
}

function save(editor: Editor, source: string): string {
  const map = load(editor, source)
  expect(map).not.toBeNull()
  return spliceMarkdown(editor, editor.state.doc, map!)
}

/** structural keys of the document's top-level nodes, normalized through the schema */
function docKeys(editor: Editor): string[] {
  const out: string[] = []
  editor.state.doc.forEach((node) => out.push(JSON.stringify(node.toJSON())))
  return out
}

function parsedKeys(editor: Editor, markdown: string): string[] {
  const content = editor.markdown!.parse(markdown).content ?? []
  return content.map((json) => JSON.stringify(editor.schema.nodeFromJSON(json).toJSON()))
}

/** the saved text must read back as the document the user saw */
function expectRoundTrip(editor: Editor, saved: string) {
  expect(parsedKeys(editor, saved)).toEqual(docKeys(editor))
}

/** position just inside the end of top-level block `index` */
function endOfBlock(editor: Editor, index: number): number {
  let pos = 0
  for (let i = 0; i < index; i++) pos += editor.state.doc.child(i).nodeSize
  return pos + editor.state.doc.child(index).nodeSize - 1
}

// Everything the plain serializer normalizes or loses: `1)` ordered markers, a
// reference-style link with its definition, a footnote, a raw HTML block, a
// setext heading, wiki brackets, dollar signs, a lazy blockquote line, and
// three blank lines in a row.
const SAMPLE = [
  '# Title',
  '',
  'Para *one* with [[wiki]] and $5 and $10.',
  '',
  '',
  '',
  '- a',
  '- b',
  '',
  '1) x',
  '2) y',
  '',
  '[ref]: https://e.com',
  '',
  'See [docs][ref].',
  '',
  '<div align="center">',
  '  <b>hi</b>',
  '</div>',
  '',
  'Text[^1].',
  '',
  '[^1]: Note.',
  '',
  'Setext',
  '======',
  '',
  '> quote',
  'lazy',
  '',
  '***',
  '',
  'last para  ',
  'hard break',
  '',
].join('\n')

describe('source splice', () => {
  it('saves an unedited document byte for byte', () => {
    const editor = createEditor()
    expect(save(editor, SAMPLE)).toBe(SAMPLE)
    // the plain serializer would have rewritten it
    expect(editor.getMarkdown()).not.toBe(SAMPLE)
  })

  it('re-serializes only the edited block', () => {
    const editor = createEditor()
    const map = load(editor, SAMPLE)!
    editor.commands.insertContentAt(endOfBlock(editor, 1), ' more')
    const out = spliceMarkdown(editor, editor.state.doc, map)
    expectRoundTrip(editor, out)
    expect(out).toContain('Para *one* with [[wiki]] and $5 and $10. more')
    for (const kept of [
      '# Title\n\n',
      '\n\n\n\n- a\n- b\n\n1) x\n2) y\n\n[ref]: https://e.com\n\nSee [docs][ref].',
      '<div align="center">\n  <b>hi</b>\n</div>',
      'Text[^1].\n\n[^1]: Note.\n\nSetext\n======',
      '> quote\nlazy\n\n***\n\nlast para  \nhard break\n',
    ]) {
      expect(out).toContain(kept)
    }
  })

  it('inserts a new block with a single blank line on each side', () => {
    const editor = createEditor()
    const map = load(editor, SAMPLE)!
    // after the bullet list (block 3), before the `1)` list
    editor.commands.insertContentAt(endOfBlock(editor, 3) + 1, '<p>inserted</p>')
    const out = spliceMarkdown(editor, editor.state.doc, map)
    expectRoundTrip(editor, out)
    expect(out).toContain('- a\n- b\n\ninserted\n\n1) x\n2) y\n\n[ref]: https://e.com')
  })

  it('keeps a definition whose block was deleted, and closes the gap', () => {
    const editor = createEditor()
    const map = load(editor, SAMPLE)!
    // delete the `1)` list; the [ref] definition sat in its trailing glue
    const from = endOfBlock(editor, 3) + 1
    editor.commands.deleteRange({ from, to: from + editor.state.doc.child(4).nodeSize })
    const out = spliceMarkdown(editor, editor.state.doc, map)
    expectRoundTrip(editor, out)
    expect(out).not.toContain('1) x')
    expect(out).toContain('- a\n- b\n\n[ref]: https://e.com\n\nSee [docs][ref].')
  })

  it('appends new trailing blocks and keeps the final newline', () => {
    const editor = createEditor()
    const map = load(editor, SAMPLE)!
    editor.commands.insertContentAt(editor.state.doc.content.size, '<p>tail</p>')
    const out = spliceMarkdown(editor, editor.state.doc, map)
    expectRoundTrip(editor, out)
    expect(out.endsWith('last para  \nhard break\n\ntail\n')).toBe(true)
  })

  it('serializes a document whose every block changed like the plain path', () => {
    const editor = createEditor()
    const map = load(editor, '# A\n\nb\n')!
    editor.commands.setContent('<h1>X</h1><p>y</p>')
    const out = spliceMarkdown(editor, editor.state.doc, map)
    expect(out).toBe('# X\n\ny\n')
  })

  it('treats the caret paragraph after a trailing list as empty source', () => {
    const editor = createEditor()
    const source = '# A\n\n- one\n- two\n'
    const map = load(editor, source)!
    expect(editor.state.doc.lastChild?.type.name).toBe('paragraph')
    expect(spliceMarkdown(editor, editor.state.doc, map)).toBe(source)
    editor.commands.insertContentAt(editor.state.doc.content.size - 1, 'tail')
    const out = spliceMarkdown(editor, editor.state.doc, map)
    expectRoundTrip(editor, out)
    expect(out).toBe('# A\n\n- one\n- two\n\ntail\n')
  })

  it('keeps a reference-style image and its definition verbatim', () => {
    for (const source of [
      'Para.\n\n![Alt text][id]\n\n[id]: https://x/a.jpg  "The Dojocat"\n',
      'Para.\n\n![Alt text][id]\n\n[id]: https://x/a.jpg\n',
      'Para.\n\n![Alt text][id] tail\n\n[id]: https://x/a.jpg\n',
    ]) {
      const editor = createEditor()
      expect(save(editor, source)).toBe(source)
      const map = load(editor, source)!
      editor.commands.insertContentAt(endOfBlock(editor, 0), ' EDITED')
      const out = spliceMarkdown(editor, editor.state.doc, map)
      expect(out).toBe(source.replace('Para.', 'Para. EDITED'))
    }
  })

  it('declines a CRLF source rather than guessing', () => {
    const editor = createEditor()
    expect(load(editor, '# A\r\n\r\nb\r\n')).toBeNull()
  })

  it('round-trips every markdown file in the repository', () => {
    const root = execSync('git rev-parse --show-toplevel', { encoding: 'utf8' }).trim()
    const files = execSync('git ls-files -- "*.md"', { cwd: root, encoding: 'utf8' })
      .split('\n')
      .filter(Boolean)
    expect(files.length).toBeGreaterThan(20)
    const editor = createEditor()
    const declined: string[] = []
    const differs: string[] = []
    for (const file of files) {
      const source = readFileSync(join(root, file), 'utf8')
      const map = load(editor, source)
      if (!map) {
        declined.push(file)
        continue
      }
      const out = spliceMarkdown(editor, editor.state.doc, map)
      if (out !== source) differs.push(file)
      // after editing the first block, every other block is still copied verbatim
      if (editor.state.doc.childCount > 1 && editor.state.doc.child(0).isTextblock) {
        editor.commands.insertContentAt(endOfBlock(editor, 0), ' edited')
        const edited = spliceMarkdown(editor, editor.state.doc, map)
        for (const unit of map.units.slice(2)) expect(edited, file).toContain(unit.core)
      }
    }
    expect(differs).toEqual([])
    expect(declined.length, `declined: ${declined.join(', ')}`).toBeLessThanOrEqual(
      Math.ceil(files.length * 0.05),
    )
  })
})
