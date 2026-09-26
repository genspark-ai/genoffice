import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { getDefaults, marked } from 'marked'
import StarterKit from '@tiptap/starter-kit'
import { Markdown } from '@tiptap/markdown'
import { TableKit } from '@tiptap/extension-table'
import { TaskItem, TaskList } from '@tiptap/extension-list'
import { buildExtensions } from '../src/renderer/editor/extensions'

const editors: Editor[] = []
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy()
})

// MarkdownManager registers tokenizers on the global marked singleton at
// construction; reset it so each editor parses with its own set only
function freshEditor(extensions: Editor['options']['extensions']): Editor {
  marked.setOptions(getDefaults())
  const editor = new Editor({ extensions, content: '' })
  editors.push(editor)
  return editor
}

/** `loose` exists only in our schema; drop it before comparing with the stock parse */
function withoutLoose(json: unknown): unknown {
  if (Array.isArray(json)) return json.map(withoutLoose)
  if (!json || typeof json !== 'object') return json
  const out: Record<string, unknown> = {}
  for (const [key, value] of Object.entries(json)) {
    if (key === 'loose') continue
    out[key] = withoutLoose(value)
  }
  if (out.attrs && Object.keys(out.attrs as object).length === 0) delete out.attrs
  return out
}

function parseWith(extensions: Editor['options']['extensions'], md: string): unknown {
  const editor = freshEditor(extensions)
  editor.commands.setContent(md, { contentType: 'markdown' })
  return withoutLoose(editor.getJSON())
}

/** the stock tokenizers, for equivalence checks */
const stock = [
  StarterKit.configure({ underline: false }),
  Markdown.configure({ indentation: { style: 'space', size: 4 } }),
  TableKit.configure({ table: { resizable: false, renderWrapper: true } }),
  TaskList,
  TaskItem.configure({ nested: true }),
]

const bounded = buildExtensions({
  slashController: {
    onOpen: () => {},
    onUpdate: () => {},
    onKeyDown: () => false,
    onClose: () => {},
  },
  slashItems: () => [],
})

const CASES: Record<string, string> = {
  orderedBasic: '1. one\n2. two\n3. three\n\nafter',
  orderedBlankBetweenItems: '1. one\n\n2. two\n\n3. three\n\nafter',
  orderedLazyContinuation: '1. one\ncontinues here\n2. two\nmore\n\nnot part',
  orderedNestedIndent: '1. one\n    - sub a\n    - sub b\n\n    para in item\n2. two\n\nout',
  orderedRomanAlpha: 'i. first\nii. second\n\na) alpha\nb) beta\n\nAA. big\n',
  orderedStart: '5. five\n6. six\n\ntext',
  orderedInterrupted: '1. one\n# heading\n2. two',
  orderedThenBullet: '1. one\n2. two\n\n- bullet\n- bullet2',
  orderedAfterBlankIndentedPara: '1. one\n\n   indented para belongs\n\nflush para does not',
  taskBasic: '- [ ] todo\n- [x] done\n\nafter',
  taskNested: '- [ ] parent\n    - [x] child\n    - [ ] child2\n\n    note\n- [ ] sibling\n\nout',
  taskBlankThenFlush: '- [ ] a\n\n- [ ] b\n\nparagraph\n- [ ] c',
  taskMixedBullet: '- [ ] a\n- plain bullet\n- [ ] b',
  tableBasic: '| a | b |\n|---|---|\n| 1 | 2 |\n\nafter',
  tableNoTrailingBlank: 'intro\n\n| a | b |\n|:--|--:|\n| 1 | 2 |\n| 3 | 4 |',
  tableWithPipesInCells: '| a | b |\n|---|---|\n| x \\| y | `c|d` |\n\nz',
  tableAdjacentParagraph: '| a | b |\n|---|---|\n| 1 | 2 |\ntrailing text\n\nnext',
  paragraphLikeTable: 'a | b\nnot a separator\n\nnext',
  mixed: [
    '# Title',
    '',
    'Paragraph with **bold** and a | pipe.',
    '',
    '1. first',
    '2. second',
    '    - nested bullet',
    '',
    '| h1 | h2 |',
    '|----|----|',
    '| c1 | c2 |',
    '',
    '- [ ] task',
    '- [x] done',
    '',
    '> quote',
    '',
    '```js',
    'const x = 1',
    '```',
    '',
    'Closing paragraph.',
  ].join('\n'),
}

describe('bounded markdown tokenizers', () => {
  it('ends a table at a whitespace-only blank line', () => {
    const md = ['| a | b |', '|---|---|', '| 1 | 2 |', '  ', 'After the table.'].join('\n')
    expect(parseWith(bounded, md)).toEqual(parseWith(stock, md))
    const json = parseWith(bounded, md) as { content: Array<{ type: string }> }
    expect(json.content.map((n) => n.type)).toEqual(['table', 'paragraph'])
  })

  it('stays linear when blank lines carry spaces and cells hold code with pipes', () => {
    const filler = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. '.repeat(20)
    const table = '| h1 | h2 |\n| --- | --- |\n| `a|b` | x |\n| c | d |'
    const md = Array.from(
      { length: 300 },
      (_, i) => `## Section ${i}\n  \n${filler}\n  \n${table}`,
    ).join('\n  \n')
    const editor = freshEditor(bounded)
    const started = performance.now()
    const json = editor.markdown.parse(md)
    const elapsed = performance.now() - started
    expect(json.content?.filter((n) => n.type === 'table')).toHaveLength(300)
    // ~60 ms here; the indexOf('\n\n') bound took ~8 s
    expect(elapsed).toBeLessThan(5000)
  })

  for (const [name, md] of Object.entries(CASES)) {
    it(`parses like the stock tokenizers: ${name}`, () => {
      expect(parseWith(bounded, md)).toEqual(parseWith(stock, md))
    })
  }

  it('parses a large document in linear time', () => {
    const section = [
      '## Section',
      '',
      'Some text, \u4e00\u4e9b\u4e2d\u6587\uff0cwith **marks** and a | pipe character.',
      '',
      '- point one',
      '- point two',
      '',
      '1. step one',
      '2. step two',
      '',
      '| a | b |',
      '|---|---|',
      '| 1 | 2 |',
      '',
      '- [ ] task',
      '',
    ].join('\n')
    const md = section.repeat(2000) // ~400 KB, 12k blocks
    const editor = freshEditor(bounded)
    const started = performance.now()
    const json = editor.markdown.parse(md)
    const elapsed = performance.now() - started
    expect(json.content?.length).toBe(12000)
    // ~0.1 s here; the stock tokenizers take ~10 s. Generous for slow CI.
    expect(elapsed).toBeLessThan(5000)
  })
})
